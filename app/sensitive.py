# ====================================================================================
# PROVENANCE : RÉIMPLÉMENTATION SPEC-BASED — le corps de ce module n'a PAS pu être
# décompilé (pas de sensitive.disasm.txt ; pycdc n'a restitué que docstring + imports).
# Sources de la spécification :
#   1. Constantes/regex intégrales extraites du .pyc (SURFACE_SECURITE.md) — reprises
#      à l'identique ci-dessous.
#   2. Contrat des kinds prouvé par _SEVERITY_BY_KIND (report_builder.disasm) :
#      classification_marking, pii_nir, pii_iban, pii_card, secret_exposed → CRITICAL ;
#      restricted_marking → HIGH ; pii_email, pii_phone, sensitive_schema → MEDIUM.
#   3. Contrats d'appel : pipeline.analyze_sensitive, job_store.mask_sensitive_filename,
#      report_builder.mask_sensitive_text, validators (_HDR_*, _NIR_SHAPE, _IBAN_SHAPE,
#      _ratio).
#   4. Docstring d'origine (redaction stricte, IGI 1300, lecture seule).
# DURCISSEMENTS au-delà du binaire, marqués [F-20]/[F-21]/[F-22] (cf. RAPPORT_AUDIT.md) :
#   [F-20] normalisation des séparateurs avant détection de marquage (le binaire ratait
#          « SECRET-DÉFENSE », « SECRET_DEFENSE », « T R E S  S E C R E T ») ;
#   [F-21] validation Luhn des suites carte (réduit les faux positifs sur identifiants) ;
#   [F-22] détection d'e-mails/téléphones par RECHERCHE dans la cellule, pas seulement
#          fullmatch (le binaire ratait « contact : jean@x.fr »).
# ====================================================================================
"""Détection de classification (IGI 1300) et de données à caractère personnel (PII).

Fonction de garde-fou MINARM : repérer dans une extraction non protégée les marquages
classifiés et les fuites de données sensibles, et les remonter comme BLOQUANT (CRITICAL).

Contraintes de sécurité (cf. CLAUDE.md) :
- Lecture seule, openpyxl, aucune évaluation de formule, aucun accès réseau.
- **Redaction stricte** : on ne recopie JAMAIS la valeur sensible (NIR, IBAN, carte,
  secret) dans le message, la preuve ou un log. Le constat décrit le TYPE et l'endroit.

Réutilise les helpers de structure de quality_analyzer (vraie ligne d'en-tête, corps,
exclusion des feuilles de métadonnées).
"""
from __future__ import annotations

import re
import unicodedata
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import get_column_letter

from app.quality_analyzer import QualityIssue, _data_rows, _header_info, _is_metadata_sheet, _n

# ----- Constantes extraites du binaire (SURFACE_SECURITE.md), reprises à l'identique --
_LABEL_BOUNDARY_L = '(?<![a-z0-9])'
_LABEL_BOUNDARY_R = '(?![a-z0-9])'

_PHRASE_MARKINGS: tuple[tuple[str, str], ...] = (
    # (motif normalisé, kind) — motifs du binaire ; matching sur texte normalisé.
    (r'tres\s*secret', 'classification_marking'),
    (r'secret\s*defense', 'classification_marking'),
    (r'confidentiel\s*defense', 'classification_marking'),
    (r'(?:cosmic\s*top\s*secret|top\s*secret)', 'classification_marking'),
    (r'(?:(?:nato|otan)\s*secret|secret\s*(?:ue|eu|otan|nato))', 'classification_marking'),
    (r'(?:diffusion\s*restreinte|special\s*france|restreint\s*(?:ue|eu))',
     'restricted_marking'),
)
_BARE_MARKINGS: tuple[tuple[str, str], ...] = (
    ('secret', 'classification_marking'),
    ('restreint', 'restricted_marking'),
)
_BARE_LABELS = re.compile(r'\b(niveau|niv|classification|classifie|mention|marquage)\b')

_CARD_RUN = re.compile(r'(?<!\d)(?:\d[ \-]?){13,19}(?!\d)')
_SECRET_KW = re.compile(
    r'mot de passe|motdepasse|password|passwd|pwd|mdp|secret|token|api[ _]?key|'
    r'cle privee|private key|bearer|client[ _]?secret|secret[ _]?key')
_ASSIGN = re.compile(r'[:=]\s*\S{3,}')
_EMAIL = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')
_EMAIL_ANY = re.compile(r'[^@\s]+@[^@\s]+\.[^@\s]{2,}')  # [F-22] recherche intra-cellule
_NIR_SHAPE = re.compile(r'^\s*\d(?:[ .]?\d){12,14}\s*$')
_IBAN_SHAPE = re.compile(r'^\s*[A-Z]{2}\d{2}[A-Z0-9 ]{10,32}\s*$')
_PHONE_SHAPE = re.compile(r'^\s*\+?(?:[\d][ .\-()]?){8,}\d\s*$')
_HDR_NIR = re.compile(r'\b(nir|nni|s[ée]cu|insee|num[ée]ro de s[ée]curit[ée])\b', re.I)
_HDR_IBAN = re.compile(r'\b(iban|rib|bban)\b', re.I)
_HDR_EMAIL = re.compile(r'\b(e?-?mail|courriel|m[ée]l)\b', re.I)
_HDR_PHONE = re.compile(r'\b(t[ée]l|phone|mobile|portable|gsm|num[ée]ro)\b', re.I)
_HDR_SECRET = re.compile(r'\b(mot de passe|motdepasse|password|passwd|pwd|mdp|secret|'
                         r'api[ _]?key|token|private key)\b', re.I)
_NIR_FULL = re.compile(r'[1-37-8]\d{2}\d{2}(?:\d{2}|2[AB])\d{3}\d{3}\d{2}')
_NUMERIC_CELL = re.compile(r'[0-9][0-9 .\-]*')
_WS_SHEET = re.compile(r'<\w*:?sheet\b[^>]*\bname="([^"]*)"')
_DEFNAME = re.compile(r'<\w*:?definedName\b[^>]*>([^<]+)</')
_DOC_FIELD = re.compile(r'<(?:\w+:)?(title|subject|creator|description|keywords|category|'
                        r'lastModifiedBy|Company|Manager)\b[^>]*>([^<]+)<')
_BULK_PART = re.compile(r'comment.*\.xml$|xl/drawings/.*\.xml$|xl/charts/.*\.xml$|'
                        r'xl/pivotCache/.*\.xml$|customXml/.*\.xml$|.*\.rels$')
_PART_MAX = 20971520


# ------------------------------- Normalisation ---------------------------------------
def _norm(text: object) -> str:
    """Minuscule + accents retirés + séparateurs unifiés en espace simple.

    [F-20] L'unification des séparateurs ([-_./\\]+ -> ' ') ferme les contournements
    « SECRET-DÉFENSE », « SECRET_DEFENSE », « très.secret » que les motifs \\s* du
    binaire laissaient passer après suppression des accents.
    """
    s = unicodedata.normalize('NFKD', str(text))
    s = ''.join(ch for ch in s if not unicodedata.combining(ch)).lower()
    s = re.sub(r'[-_./\\]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _ratio(values: list[str], pattern: re.Pattern) -> float:
    """Part des valeurs d'une colonne conformes au motif (utilisé aussi par validators)."""
    if not values:
        return 0.0
    return sum(1 for v in values if pattern.match(str(v))) / len(values)


def _luhn_ok(digits: str) -> bool:
    """[F-21] Clé de Luhn — filtre les suites de 13-19 chiffres qui ne sont pas des PAN."""
    total, parity = 0, len(digits) % 2
    for i, ch in enumerate(digits):
        d = int(ch)
        if i % 2 == parity:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


def _marking_hits(text: object, *, header: str = '') -> list[tuple[str, str]]:
    """[(kind, libellé du motif)] des marquages détectés dans `text` (normalisé)."""
    norm = _norm(text)
    if not norm:
        return []
    hits: list[tuple[str, str]] = []
    for pattern, kind in _PHRASE_MARKINGS:
        m = re.search(_LABEL_BOUNDARY_L + pattern + _LABEL_BOUNDARY_R, norm)
        if m:
            hits.append((kind, m.group(0)))
    if not hits:
        context = _norm(header) + ' ' + norm
        if _BARE_LABELS.search(context):
            for word, kind in _BARE_MARKINGS:
                if re.search(_LABEL_BOUNDARY_L + word + _LABEL_BOUNDARY_R, norm):
                    hits.append((kind, word))
                    break
    return hits


# ----------------------------- Masquage (redaction) ----------------------------------
def mask_sensitive_text(text: object) -> str:
    """Neutralise NIR/IBAN/carte/secrets dans un texte destiné au rapport ou aux logs.

    Ne supprime pas les marquages de classification (ils sont l'information utile du
    constat) ; supprime les VALEURS nominatives.
    """
    if text is None:
        return ''
    s = str(text)
    s = _CARD_RUN.sub('[masqué]', s)

    def _mask_numeric(m: re.Match) -> str:
        run = re.sub(r'\D', '', m.group(0))
        if len(run) >= 12:
            return '[masqué]'
        return m.group(0)

    s = _NUMERIC_CELL.sub(_mask_numeric, s)
    s = re.sub(r'\b[A-Z]{2}\d{2}[A-Z0-9 ]{10,32}\b', '[masqué]', s)
    s = re.sub(r'(' + _SECRET_KW.pattern + r')\s*[:=]\s*\S+', r'\1=[masqué]', s,
               flags=re.I)
    return s


def mask_sensitive_filename(name: object) -> str:
    """Nom de fichier sûr pour le rapport : valeurs sensibles masquées, extension gardée."""
    if not name:
        return 'fichier'
    p = Path(str(name))
    stem = mask_sensitive_text(p.stem)
    return f'{stem}{p.suffix}'


# ------------------------------- Analyse cellules ------------------------------------
def _issue(kind: str, sheet: str, cell: str, severity: str, message: str,
           *, confidence: str = 'high', evidence: str = '') -> QualityIssue:
    return QualityIssue(kind=kind, sheet=sheet, cell=cell, severity=severity,
                        confidence=confidence, message=message, evidence=evidence)


def _scan_sheet(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)

    # 1) Marquages de classification : toutes cellules + titre de feuille.
    for kind, label in _marking_hits(ws.title):
        issues.append(_issue(kind, ws.title, 'A1',
                             'CRITICAL' if kind == 'classification_marking' else 'HIGH',
                             f'Marquage « {label} » dans le nom de la feuille — fichier '
                             'non diffusable en l\'état.'))
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            header = headers.get(cell.column, '')
            for kind, label in _marking_hits(cell.value, header=header):
                sev = 'CRITICAL' if kind == 'classification_marking' else 'HIGH'
                issues.append(_issue(kind, ws.title, cell.coordinate, sev,
                                     f'Marquage de classification « {label} » détecté '
                                     f'(cellule {cell.coordinate}).'))

    if not headers or not body:
        return issues

    # 2) PII par colonne (en-tête OU forme dominante), redaction stricte.
    for col, name in headers.items():
        cells = [(r, str(ws.cell(r, col).value).strip()) for r in body
                 if ws.cell(r, col).value not in (None, '')]
        raw = [v for _, v in cells]
        letter = get_column_letter(col)

        hdr_nir = bool(_HDR_NIR.search(name))
        hdr_iban = bool(_HDR_IBAN.search(name))
        hdr_email = bool(_HDR_EMAIL.search(name))
        hdr_phone = bool(_HDR_PHONE.search(name))
        hdr_secret = bool(_HDR_SECRET.search(name))
        col_is_idkind = False

        if not cells:
            if hdr_nir or hdr_iban or hdr_secret:
                issues.append(_issue(
                    'sensitive_schema', ws.title, f'{letter}{hdr}', 'MEDIUM',
                    f'Colonne «{name}» : en-tête de donnée sensible sans contenu — '
                    'gabarit destiné à recevoir des données protégées.',
                    confidence='medium'))
            continue

        if hdr_nir or _ratio(raw, _NIR_SHAPE) >= 0.6:
            hits = [(r, v) for r, v in cells
                    if _NIR_FULL.search(re.sub(r'[ .]', '', v))]
            if hits:
                col_is_idkind = True
                issues.append(_issue(
                    'pii_nir', ws.title, f'{letter}{hits[0][0]}', 'CRITICAL',
                    f'Colonne «{name}» : {_n(len(hits), "NIR détecté", "NIR détectés")} '
                    '(numéro de sécurité sociale) — valeur jamais recopiée.'))
        if hdr_iban or _ratio(raw, _IBAN_SHAPE) >= 0.6:
            hits = [(r, v) for r, v in cells if _IBAN_SHAPE.match(v.upper())]
            if hits:
                col_is_idkind = True
                issues.append(_issue(
                    'pii_iban', ws.title, f'{letter}{hits[0][0]}', 'CRITICAL',
                    f'Colonne «{name}» : {_n(len(hits), "IBAN détecté", "IBAN détectés")} '
                    '— valeur jamais recopiée.'))
        if hdr_secret:
            filled = [(r, v) for r, v in cells if v]
            if filled:
                col_is_idkind = True
                issues.append(_issue(
                    'secret_exposed', ws.title, f'{letter}{filled[0][0]}', 'CRITICAL',
                    f'Colonne «{name}» : {_n(len(filled), "secret stocké en clair", "secrets stockés en clair")} '
                    '— valeur jamais recopiée.'))
        # [F-23] Une colonne déjà qualifiée NIR/IBAN/secret n'est pas re-signalée
        # e-mail/téléphone : les regex de forme (héritées du binaire) se recouvrent,
        # le doublon MEDIUM ne ferait que bruiter à côté du CRITICAL.
        if not col_is_idkind and (hdr_email or _ratio(raw, _EMAIL) >= 0.6):
            hits = [r for r, v in cells if _EMAIL_ANY.search(v)]
            if hits:
                issues.append(_issue(
                    'pii_email', ws.title, f'{letter}{hits[0]}', 'MEDIUM',
                    f'Colonne «{name}» : {_n(len(hits), "adresse e-mail", "adresses e-mail")} '
                    '(donnée personnelle).', confidence='medium'))
        if not col_is_idkind and (hdr_phone or _ratio(raw, _PHONE_SHAPE) >= 0.6):
            hits = [r for r, v in cells if _PHONE_SHAPE.match(v)]
            if hits:
                issues.append(_issue(
                    'pii_phone', ws.title, f'{letter}{hits[0]}', 'MEDIUM',
                    f'Colonne «{name}» : {_n(len(hits), "numéro de téléphone", "numéros de téléphone")} '
                    '(donnée personnelle).', confidence='medium'))

    # 3) Cartes bancaires et secrets « clé = valeur » : balayage cellule à cellule.
    for row in ws.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue
            v = cell.value
            m = _CARD_RUN.search(v)
            if m and _luhn_ok(re.sub(r'\D', '', m.group(0))):
                issues.append(_issue(
                    'pii_card', ws.title, cell.coordinate, 'CRITICAL',
                    f'Numéro de carte bancaire plausible (clé de Luhn valide) en '
                    f'{cell.coordinate} — valeur jamais recopiée.'))
            low = _norm(v)
            if _SECRET_KW.search(low) and _ASSIGN.search(v):
                issues.append(_issue(
                    'secret_exposed', ws.title, cell.coordinate, 'CRITICAL',
                    f'Secret en clair (mot de passe/clé/token) en {cell.coordinate} — '
                    'valeur jamais recopiée.'))
    return issues


# --------------------------- Analyse des parts XML (hors cellules) -------------------
def _scan_xml_parts(path: str | Path) -> list[QualityIssue]:
    """Marquages/PII dans les zones que l'analyse cellulaire ne voit pas :
    commentaires, dessins, graphiques, pivotCache, customXml, .rels, docProps,
    noms de feuilles et noms définis du workbook (regex _WS_SHEET/_DEFNAME/_DOC_FIELD)."""
    issues: list[QualityIssue] = []
    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile):
        return issues
    with zf:
        for info in zf.infolist():
            name = info.filename
            interesting = (_BULK_PART.search(name)
                           or name.startswith('docProps/')
                           or name == 'xl/workbook.xml')
            if not interesting or info.file_size > _PART_MAX:
                continue
            try:
                text = zf.read(name).decode('utf-8', errors='replace')
            except (OSError, zipfile.BadZipFile, RuntimeError):
                continue
            candidates: list[str] = []
            if name == 'xl/workbook.xml':
                candidates += _WS_SHEET.findall(text)
                candidates += _DEFNAME.findall(text)
            elif name.startswith('docProps/'):
                candidates += [val for _, val in _DOC_FIELD.findall(text)]
            else:
                candidates.append(text)
            for chunk in candidates:
                for kind, label in _marking_hits(chunk):
                    sev = 'CRITICAL' if kind == 'classification_marking' else 'HIGH'
                    issues.append(_issue(kind, f'part:{name}', '—', sev,
                                         f'Marquage « {label} » hors cellules '
                                         f'(partie OOXML {name}).'))
                m = _CARD_RUN.search(chunk)
                if m and _luhn_ok(re.sub(r'\D', '', m.group(0))):
                    issues.append(_issue('pii_card', f'part:{name}', '—', 'CRITICAL',
                                         f'Numéro de carte plausible hors cellules '
                                         f'({name}) — valeur jamais recopiée.'))
                digits = re.sub(r'[ .]', '', chunk)
                if _NIR_FULL.search(digits) and len(chunk) < 4096:
                    issues.append(_issue('pii_nir', f'part:{name}', '—', 'CRITICAL',
                                         f'NIR plausible hors cellules ({name}) — '
                                         'valeur jamais recopiée.'))
    return issues


def _coerce(path_or_wb):
    if isinstance(path_or_wb, Workbook):
        return (path_or_wb, False, None)
    return (load_workbook(path_or_wb, read_only=False, data_only=False, keep_vba=False),
            True, Path(path_or_wb))


def analyze_sensitive(path_or_wb, *, deadline=None) -> list[QualityIssue]:
    """Détection classification + PII sur les feuilles de données et les parts OOXML."""
    wb, should_close, src = _coerce(path_or_wb)
    try:
        issues: list[QualityIssue] = []
        for ws in wb.worksheets:
            if deadline is not None:
                import time
                if time.monotonic() > deadline:
                    raise TimeoutError('budget temps analyse sensible dépassé')
            if _is_metadata_sheet(ws):
                continue
            issues.extend(_scan_sheet(ws))
        if src is not None:
            issues.extend(_scan_xml_parts(src))
        return issues
    finally:
        if should_close:
            wb.close()
