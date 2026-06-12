# Rapport d'audit & reconstruction — xls-analyzer v2

**Date :** 12/06/2026 · **Périmètre :** source reconstruite par pycdc/pycdas depuis
`PC_XLS_Analyser.exe` (PyInstaller, Python 3.12) · **Livrable :** codebase v2 exécutable,
testée (44 tests), avec provenance tracée fichier par fichier.

**Méthode.** Lecture intégrale des désassemblages pycdas (≈ 18 500 lignes), reconstruction
opcode par opcode quand le disasm existait, ré-implémentation spec-based quand il
n'existait pas (sensitive.py), corrections d'artefacts pycdc prouvées par le bytecode,
puis durcissements ciblés — chaque écart au binaire est numéroté et marqué dans l'en-tête
du fichier concerné.

---

## 1. Verdict d'ensemble

| Question | Réponse |
| --- | --- |
| La source pycdc était-elle exécutable ? | **Non** — 8 modules sur 19 ne passaient pas `ast.parse` ; les modules « propres » contenaient des artefacts sémantiques silencieux (F-01/F-02/F-03). |
| L'architecture d'origine est-elle saine ? | **Oui.** Défense en profondeur réelle : validation amont → préflight fail-closed sur métadonnées ZIP → analyse statique → contenu, écritures atomiques, redaction, loopback. Les choix (jamais décompresser pour décider, fail-closed systématique) sont corrects. |
| La v2 est-elle iso-comportement ? | **Oui pour le cœur vérifié au bytecode** (préflight, qualité ×23 détecteurs, validateurs, scoring, rapport, routes). **Spec-based pour sensitive.py** (disasm absent) avec constantes d'origine + contrat de sévérités prouvé. Quatre durcissements assumés (F-20→F-23). |
| État final | `ast.parse` 19/19, imports 19/19, **44/44 tests**, pipeline bout-en-bout REJECTED/STATIC_DONE/BLOQUÉ conformes, zéro fuite NIR/IBAN dans les artefacts. |

---

## 2. Constats (FAIT, prouvés par bytecode ou test)

### P0 — bloquants à l'exécution (artefacts de décompilation)

| ID | Module | Constat | Correction v2 |
| --- | --- | --- | --- |
| F-01 | reject.py | pycdc a inversé les bases d'enum : `class RejectCode(Enum, str)` → `TypeError` à l'import. Le bytecode module-level charge `str` PUIS `Enum`. | Bases `(str, Enum)` rétablies. |
| F-02 | models.py | Champs Pydantic perdus : `Finding` réduit à 4 champs (réel : 14, prouvé par les `KW_NAMES` de report_builder, CALL 10/15), `VerdictSummary.verdict` absent, validateur `evidence` mutilé. | Modèles reconstitués champ à champ par preuve croisée ; `extra='forbid'` conservé. |
| F-03 | models.py / job_store | `Job` amputé (job_id/status/created_at/updated_at/filename_original) ; `JobStore.__init__`/`update` à corps vide (`pass`). | Reconstitués depuis docstrings d'origine + contrat d'appel pipeline (`AttributeError` sur champ inconnu conservé). |
| F-04 | 8 modules | `atomic_io, main, quality_analyzer, reject, report_builder, risk_scoring, static_analyzer, validators` : SyntaxError pycdc (corps tronqués, `None.get(...)`, double-négatifs). | Reconstruction fidèle intégrale depuis les .disasm (cf. §4 provenance). |
| F-05 | pipeline.py | `_dedup_issues` : artefact `if not key not in seen: continue` (double négatif) qui aurait inversé le dédoublonnage. | Logique rétablie (`if key in seen: continue`). |

### P1 — sécurité / doctrine (présents dans le binaire d'origine)

| ID | Module | Constat | Traitement v2 |
| --- | --- | --- | --- |
| F-20 | sensitive.py | **Faux négatifs marquages** : les motifs `secret\s*defense` etc. ne couvrent ni tirets ni underscores ni points. « SECRET-DÉFENSE », « SECRET_DEFENSE », « Très.Secret » passaient. Pour un outil dont la fonction première est le blocage IGI 1300, c'est l'écart le plus grave du binaire. | **Durci** : normalisation `[-_./\\]+ → espace` avant matching (motifs d'origine conservés). Testé sur 11 variantes + 4 anti-faux-positifs (« secrétaire », « secret professionnel »…). |
| F-21 | sensitive.py | Suites de 13-19 chiffres (`_CARD_RUN`) signalées carte sans validation : tout identifiant long = faux positif CRITICAL. | **Durci** : clé de Luhn exigée. Testé valide/invalide. |
| F-22 | sensitive.py | E-mails détectés par `fullmatch` de cellule uniquement : « contact : jean@x.fr (N3) » invisible. | **Durci** : recherche intra-cellule (`_EMAIL_ANY`) quand la colonne est qualifiée e-mail. |
| F-06 | validators.py | **Asymétrie de redaction assumée par le binaire** (prouvée au bytecode) : NIR/IBAN jamais échoés (`evidence=''`), mais e-mails/téléphones invalides recopiés EN CLAIR dans message+evidence. Écart à la doctrine PII si le rapport circule. | **Conservé tel quel** (fidélité), mais neutralisé en aval : `quality_issues_to_findings` passe tout par `mask_sensitive_text` (comportement du binaire, vérifié au bytecode) — la fuite résiduelle ne concerne que e-mails/téléphones, par conception d'origine. Signalé pour arbitrage : masquer aussi ces deux types coûte une ligne. |
| F-07 | services/preflight_zip.py | Corps `preflight_zip` perdu : impossible de prouver l'ordre exact des contrôles du binaire. | Ré-implémenté selon les RejectCode + seuils config + doctrine documentée ; ordre du moins cher au plus cher, premier échec = rejet. Contrôle `XML_DEPTH` borné en flux (lecture plafonnée), marqué interprétation. |
| F-08 | app_main.py (lanceur) | Déjà constaté en amont : `sys.stderr.isatty()` sous PyInstaller `--noconsole` (stderr=None) → crash au boot ; `uvicorn log_config` par défaut écrit vers des flux absents. | Lanceur corrigé conservé (garde `isatty`, `log_config=None`). |

### P2 — robustesse / qualité (binaire correct mais perfectible)

| ID | Module | Constat | Traitement v2 |
| --- | --- | --- | --- |
| F-23 | sensitive.py | Recouvrement des regex de forme : une colonne NIR espacée matche aussi `_PHONE_SHAPE` → bruit `pii_phone` MEDIUM à côté du `pii_nir` CRITICAL. | **Durci** : colonne déjà qualifiée NIR/IBAN/secret non re-signalée e-mail/téléphone. |
| F-09 | main.py | Rendu HTML résultat : disasm structurellement lisible mais l'assemblage exact au caractère près n'est pas prouvable. | Reconstruit fidèle en structure (classes CSS, colonnes, bannière verdict, actions) ; `html.escape` systématique sur toute donnée utilisateur — y compris titres de feuille hostiles (testé `=HYPERLINK(evil)`). |
| F-10 | report_builder.py | `render_markdown` : sections/libellés prouvés par constantes, mise en forme reconstruite. | Idem F-09 — fidèle en contenu, pas au caractère près. Marqué dans l'en-tête. |

---

## 3. Ce qui a été VÉRIFIÉ conforme au binaire (échantillon des preuves)

- **Préflight** : `MIN_ENTRY_BYTES_FOR_RATIO=4096` conditionne le contrôle de ratio
  (un zip de zéros de 800 Ko passe — comportement du binaire, reproduit et testé) ;
  traversal/doublons/5000 entrées/limites totales identiques.
- **Qualité** : les 23 détecteurs, leurs seuils (`_MIN_TABLE_ROWS=5`, dominance 0.6/0.7/0.8,
  MAD×1.4826×z=6, tolérance totaux `max(1e-6, |total|×1e-9)`, fenêtre 0.95/0.8/0.5 des
  clés, 90 % de tri pour `date_out_of_order`), l'ordre du tuple `_DETECTORS`, les messages
  FR au caractère près, l'exclusion des feuilles de métadonnées et des lignes de total.
- **Validateurs** : NIR sans clé mod-97 (choix documenté du binaire), `_IBAN_LEN` par pays,
  bascule format US `mo>12 & d≤12`, futur strict `> année courante + 1`, cascade
  d'identification NIR→IBAN→e-mail→téléphone au seuil 0.6.
- **Scoring** : poids 0/1/3/7/15, niveau = max d'ordre 0..4.
- **Rapport** : `_SEVERITY_BY_KIND` intégral (45 entrées), `_VERDICT_REASON`, notes
  A/B/C/D, masquage `network_formula`/`unc_path` → « (cible réseau masquée) ».
- **API** : middleware Content-Length→413 (testé), streaming borné 1 Mo/chunk→413,
  `_safe_filename` (contrôle chars, basename, 255), `_csv_safe` anti-injection
  (`=+-@\t\r` → apostrophe, testé sur rapport réel), UTF-8 BOM, routes et codes
  400/404/409/507 identiques.

---

## 4. Provenance par fichier (traçabilité)

| Fichier | Provenance | Notes |
| --- | --- | --- |
| app/reject.py | Fidèle (bytecode) | F-01 corrigé |
| app/atomic_io.py | Fidèle (bytecode) | mode 0o600, fsync dir best-effort |
| app/config.py | pycdc propre | constantes inchangées |
| app/models.py | Reconstruction corrigée | F-02/F-03, preuve croisée KW_NAMES |
| app/risk_scoring.py | Fidèle (bytecode) | |
| app/services/static_analyzer.py | Fidèle (bytecode) | |
| app/services/file_validation.py | Corrigée + spec | signatures ZIP |
| app/services/preflight_zip.py | Corrigée + spec | F-07 |
| app/quality_analyzer.py | Fidèle (bytecode, 8 546 l.) | 23 détecteurs |
| app/sensitive.py | **Spec-based** | constantes d'origine + F-20/21/22/23 |
| app/validators.py | Fidèle (bytecode) | F-06 documenté |
| app/rules_engine.py | Corrigée + spec | une anomalie / (colonne, règle) |
| app/job_store.py | Fidèle + reconstitué | __init__/update |
| app/recalc_client.py | pycdc propre | |
| app/report_builder.py | Fidèle ; render_markdown reconstruit | F-10 |
| app/pipeline.py | Fragments fidèles + run_pipeline spec | transitions SPEC respectées |
| app/main.py | Routes fidèles ; CSS/HTML statiques **au caractère près** ; rendu résultat reconstruit | F-09 |
| app_main.py | Lanceur corrigé (audit amont) | F-08 |
| tests/test_v2.py | Nouveau | 44 tests, T1→T10 |

---

## 5. Écarts planifié / livré

| Prévu | Livré | Écart |
| --- | --- | --- |
| Reconstruction fidèle de tous les modules disasm | Fait (opcode/opcode) | `render_markdown` et `_render_result_html` : fidèles en contenu/structure, pas au caractère près (assemblage f-strings non prouvable) — marqué. |
| Ré-implémentation sensitive.py | Fait | Sans le bytecode, l'agrégation exacte du binaire (par cellule vs par colonne) reste une interprétation ; le contrat kinds/sévérités, lui, est prouvé. |
| Tests | 44 tests, préflight→API | Pas de test de charge ni de fuzzing ZIP systématique (recommandé avant homologation). |
| Recalc runner | Hors périmètre (manifeste seul, comme l'app d'origine) | Conforme à l'invariant « zéro conteneur dans app/ ». |

## 6. Recommandations restantes (par priorité)

1. **Arbitrer F-06** : étendre la redaction aux e-mails/téléphones invalides dans
   `validators.py` (une ligne) — sinon un rapport diffusé contient des PII MEDIUM en clair.
2. Geler `rules.json` par environnement et activer `XLS_RULES_PATH` (le moteur est prêt).
3. Fuzzing du préflight (zips malformés générés) + jeu de fichiers de recette MINARM
   avant toute homologation ; les tests fournis sont un socle, pas une preuve.
4. Si le binaire est reconstruit : PyInstaller `--noconsole` avec le lanceur corrigé,
   et re-dériver les hash de référence du `.exe`.

**Confiance globale : 8/10.** Cœur vérifié au bytecode et testé ; les 2 points retirés
couvrent les zones spec-based (sensitive, preflight, run_pipeline) où la conformité est
au contrat, pas à l'octet.
