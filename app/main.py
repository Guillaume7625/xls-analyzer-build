# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE pour les routes, middlewares et helpers texte
# (_limit_body_size, _safe_filename, _stream_to_file, _csv_safe, _cat_label, _CAT_FR
# intégral, upload, get_job, _existing_report, get_report_json/md/csv, post_recalc,
# run — vérifiés sur main.disasm.txt). _BASE_CSS / UPLOAD_HTML / GUIDE_HTML repris
# AU CARACTÈRE PRÈS du binaire. Le rendu HTML du résultat (_render_result_html et ses
# aides) est FIDÈLE EN STRUCTURE (classes CSS, colonnes, bannière verdict, actions)
# mais ré-assemblé — l'HTML exact au caractère près n'est pas garanti.
# ====================================================================================
"""API FastAPI locale de xls-analyzer.

Orchestration PAR FICHIERS uniquement. L'app n'orchestre aucun conteneur et ne
lance aucun recalc : elle valide, fait le preflight, l'analyse statique, écrit le
rapport, et (sur demande) dépose un manifeste de recalc pour le runner.

Bind 127.0.0.1 (config.BIND_HOST). Docs publiques désactivées par défaut
(config.ENABLE_PUBLIC_DOCS). Upload borné en streaming à MAX_UPLOAD_BYTES.
"""
from __future__ import annotations

import csv
import html
import io
import json
import re
from pathlib import Path

from fastapi import Depends, FastAPI, HTTPException, Request, UploadFile
from fastapi.responses import (HTMLResponse, JSONResponse, PlainTextResponse,
                               Response)

from app import recalc_client
from app.config import BIND_HOST, ENABLE_PUBLIC_DOCS, MAX_UPLOAD_BYTES
from app.job_store import REPORT_JSON, REPORT_MD, JobStore
from app.models import Job, JobStatus
from app.pipeline import run_pipeline

app = FastAPI(title='xls-analyzer', version='0.1.0',
              docs_url='/docs' if ENABLE_PUBLIC_DOCS else None,
              redoc_url='/redoc' if ENABLE_PUBLIC_DOCS else None,
              openapi_url='/openapi.json' if ENABLE_PUBLIC_DOCS else None)

_CHUNK = 1048576
_BASE_CSS = '\n*{box-sizing:border-box}\nbody{margin:0;background:#f4f6f8;color:#1f2933;\n font:15px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif}\n.wrap{max-width:64rem;margin:0 auto;padding:0 1.25rem}\na{color:#1e40af}\ncode{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;font-size:.85em;\n background:#eef1f4;padding:.05rem .3rem;border-radius:4px}\n.topbar{background:#1e293b;color:#fff}\n.topbar .wrap{display:flex;align-items:center;justify-content:space-between;height:3.25rem}\n.brand{font-weight:600;letter-spacing:.02em;font-size:1rem}\n.brand .dot{color:#60a5fa}\nh1{font-size:1.35rem;margin:1.5rem 0 .25rem}\nh2{font-size:1rem;color:#1e293b;margin:0 0 .75rem}\n.meta{color:#6b7280;font-size:.85rem;margin:.25rem 0 1.1rem}\n.card{background:#fff;border:1px solid #e3e7eb;border-radius:12px;\n padding:1.1rem 1.25rem;margin:1.1rem 0;box-shadow:0 1px 2px rgba(16,24,40,.04)}\n.status{padding:.2rem .65rem;border-radius:999px;font-size:.78rem;font-weight:600;letter-spacing:.02em}\n.status--ok{background:#dcfce7;color:#166534}\n.status--bad{background:#fee2e2;color:#991b1b}\n.sev{display:inline-block;min-width:4.6rem;text-align:center;padding:.12rem .5rem;\n border-radius:999px;font-size:.72rem;font-weight:700;color:#fff;letter-spacing:.03em}\n.sev--CRITICAL{background:#7f1d1d}\n.sev--HIGH{background:#b91c1c}\n.sev--MEDIUM{background:#c2410c}\n.sev--LOW{background:#475569}\n.sev--INFO{background:#64748b}\n.conf{color:#6b7280;font-size:.8rem;text-transform:capitalize}\n.note-ok{color:#15803d;font-weight:600}\n.note-bad{color:#b91c1c;font-weight:600}\n.empty{color:#15803d;font-size:.9rem;margin:0}\n.actions{display:flex;flex-wrap:wrap;gap:.5rem;margin:.25rem 0 0;padding:0}\n.btn{display:inline-block;padding:.45rem .85rem;border-radius:8px;border:1px solid #d7dce1;\n background:#fff;color:#1f2933;text-decoration:none;font-size:.88rem;font-weight:500}\n.btn:hover{border-color:#b6c0ca;background:#f8fafc}\n.btn--primary{background:#1e293b;border-color:#1e293b;color:#fff}\n.btn--primary:hover{background:#0f172a}\n.tablewrap{overflow-x:auto;border:1px solid #e8ebee;border-radius:10px}\ntable{border-collapse:collapse;width:100%;font-size:.88rem}\nthead th{text-align:left;background:#f1f4f7;color:#475569;font-weight:600;\n padding:.55rem .7rem;border-bottom:1px solid #e3e7eb;white-space:nowrap}\ntbody td{padding:.5rem .7rem;border-top:1px solid #eef1f4;vertical-align:top}\ntbody tr:nth-child(even){background:#fbfcfd}\ntbody tr:hover{background:#f4f7fb}\ntd.mono{font-family:ui-monospace,Menlo,monospace;color:#475569;white-space:nowrap}\n.hero{max-width:36rem;margin-top:3rem}\n.lead{color:#6b7280;margin:.25rem 0 1.5rem}\n.uploader{border:1.5px dashed #c3ccd6;border-radius:14px;background:#fff;\n padding:2rem 1.5rem;text-align:center}\n.uploader input[type=file]{display:block;margin:0 auto 1.1rem;max-width:100%}\n.uploader .btn--primary{padding:.55rem 1.4rem;font-size:.95rem}\n.foot{color:#94a3b8;font-size:.8rem;margin-top:1.25rem}\n.guide-link{text-align:center;margin:1.1rem 0 0}\n.lead-big{font-size:1.05rem;color:#374151;margin:.25rem 0 1.5rem}\n.steps{list-style:none;counter-reset:s;padding:0;margin:.5rem 0}\n.steps li{counter-increment:s;position:relative;padding:.45rem 0 .45rem 2.5rem;color:#374151}\n.steps li::before{content:counter(s);position:absolute;left:0;top:.4rem;width:1.7rem;height:1.7rem;\n background:#1e293b;color:#fff;border-radius:50%;display:grid;place-items:center;font-size:.82rem;font-weight:700}\n.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(15rem,1fr));gap:.9rem}\n.gcard{border:1px solid #e3e7eb;border-radius:10px;padding:.9rem 1rem;background:#fbfcfd}\n.gcard h3{margin:.1rem 0 .55rem;font-size:.95rem;color:#1e293b}\n.chips{display:flex;flex-wrap:wrap;gap:.35rem;margin:.4rem 0 0}\n.chip{font-size:.78rem;background:#eef2f6;color:#334155;border:1px solid #e2e8f0;border-radius:999px;padding:.12rem .55rem}\n.legend{display:flex;flex-wrap:wrap;gap:.5rem 1.1rem;align-items:center;margin:.5rem 0}\n.legend span{display:flex;align-items:center;gap:.45rem;font-size:.85rem;color:#475569}\nul.plain{margin:.4rem 0;padding-left:1.15rem}\nul.plain li{margin:.25rem 0}\n.verdict{display:flex;align-items:center;gap:1.1rem;border-radius:12px;\n padding:1rem 1.25rem;margin:1.1rem 0;border:1px solid}\n.verdict .grade{font-size:2.2rem;font-weight:800;line-height:1;min-width:2.2rem;text-align:center}\n.verdict .vlabel{font-size:1.15rem;font-weight:700;letter-spacing:.02em}\n.verdict .vreason{font-size:.9rem;opacity:.85;margin-top:.15rem}\n.verdict--BLOQUE{background:#fef2f2;border-color:#fecaca;color:#991b1b}\n.verdict--A_VERIFIER{background:#fffbeb;border-color:#fde68a;color:#92400e}\n.verdict--CERTIFIE{background:#f0fdf4;border-color:#bbf7d0;color:#166534}\n.sheets{display:flex;flex-wrap:wrap;gap:.5rem;margin:.1rem 0 1rem}\n.sheetv{font-size:.82rem;border:1px solid #e2e8f0;border-radius:8px;padding:.3rem .6rem;background:#fff;color:#1f2933}\n.sheetv .v{font-weight:700}\n.sheetv .v--BLOQUE{color:#b91c1c}.sheetv .v--A_VERIFIER{color:#b45309}.sheetv .v--CERTIFIE{color:#15803d}\n.counts{font-weight:400;font-size:.78rem;margin-left:.4rem;white-space:nowrap}\n.counts .sev{min-width:0;margin-left:.3rem}\n'
UPLOAD_HTML = f'''<!doctype html>\n<html lang="fr"><head><meta charset="utf-8">\n<title>xls-analyzer</title>\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<style>{_BASE_CSS}</style></head>\n<body>\n<header class="topbar"><div class="wrap">\n <span class="brand">xls-analyzer<span class="dot">.</span></span>\n <span class="status status--ok">analyse locale</span>\n</div></header>\n<main class="wrap hero">\n <h1>Analyse locale d\'un fichier Excel</h1>\n <p class="lead">Contrôle de sécurité et de qualité des données, 100&nbsp;% hors-ligne.</p>\n <form class="uploader" action="/upload" method="post" enctype="multipart/form-data">\n  <input type="file" name="file" accept=".xlsx" required>\n  <button class="btn btn--primary" type="submit">Analyser le fichier</button>\n </form>\n <p class="guide-link"><a href="/guide">📖 Comment ça marche&nbsp;? — guide en 1 minute</a></p>\n <p class="foot">.xlsx uniquement · aucune formule évaluée, aucune macro exécutée, aucun accès réseau.</p>\n</main>\n</body></html>'''
GUIDE_HTML = f'''<!doctype html>\n<html lang="fr"><head><meta charset="utf-8">\n<title>Guide — xls-analyzer</title>\n<meta name="viewport" content="width=device-width, initial-scale=1">\n<style>{_BASE_CSS}</style></head>\n<body>\n<header class="topbar"><div class="wrap">\n <span class="brand">xls-analyzer<span class="dot">.</span></span>\n <a class="btn" href="/">Analyser un fichier</a>\n</div></header>\n<main class="wrap">\n<h1>Guide — comprendre l\'outil en 1 minute</h1>\n<p class="lead-big">Cet outil vérifie vos fichiers Excel : il repère les erreurs de données\net les risques de sécurité, en restant 100&nbsp;% sur votre machine.</p>\n\n<section class="card">\n<h2>À quoi ça sert&nbsp;?</h2>\n<p>Vous déposez un fichier <code>.xlsx</code>. En quelques secondes, l\'outil le passe au crible\net vous liste, en français clair&nbsp;:</p>\n<ul class="plain">\n <li>les <b>erreurs de données</b> : formules cassées, totaux faux, trous, doublons, formats trompeurs…</li>\n <li>les <b>risques de sécurité</b> : macros, fichiers piégés, liens externes…</li>\n</ul>\n<p>Aucune donnée ne quitte votre poste&nbsp;: pas d\'internet, pas de cloud.</p>\n</section>\n\n<section class="card">\n<h2>Comment l\'utiliser&nbsp;?</h2>\n<ol class="steps">\n <li><b>Déposez</b> votre fichier Excel (.xlsx) sur la page d\'accueil.</li>\n <li><b>L\'outil l\'analyse</b> — sans ouvrir vos formules ni lancer la moindre macro.</li>\n <li><b>Lisez le rapport</b>, puis exportez la liste en CSV pour la traiter dans Excel si besoin.</li>\n</ol>\n</section>\n\n<section class="card">\n<h2>Comment lire les résultats&nbsp;?</h2>\n<p>Chaque point relevé porte une <b>pastille de gravité</b>&nbsp;:</p>\n<div class="legend">\n <span><span class="sev sev--CRITICAL">CRITICAL</span> danger — fichier refusé</span>\n <span><span class="sev sev--HIGH">HIGH</span> à corriger en priorité</span>\n <span><span class="sev sev--MEDIUM">MEDIUM</span> à vérifier</span>\n <span><span class="sev sev--LOW">LOW</span> mineur / pour information</span>\n</div>\n<ul class="plain">\n <li><b>Cellule</b> : l\'endroit exact à regarder dans votre fichier (ex. <code>J81</code>).</li>\n <li><b>Détail</b> : ce qui ne va pas, expliqué simplement.</li>\n <li><b>Type</b> : la catégorie du problème (survolez le libellé pour voir le code technique).</li>\n <li><b>Confiance</b> : à quel point l\'outil est sûr de lui (élevée, moyenne, basse).</li>\n</ul>\n</section>\n\n<section class="card">\n<h2>Ce que l\'outil sait repérer dans vos données</h2>\n<div class="grid">\n <div class="gcard"><h3>🧮 Calculs &amp; formules</h3><div class="chips">\n  <span class=chip>Cellule en erreur</span><span class=chip>Formule écrasée</span>\n  <span class=chip>Formule incohérente</span><span class=chip>Total incorrect</span>\n  <span class=chip>Plage incomplète</span><span class=chip>Référence circulaire</span>\n  <span class=chip>Formule sur cellule vide</span></div></div>\n <div class="gcard"><h3>📭 Trous &amp; doublons</h3><div class="chips">\n  <span class=chip>Valeur manquante</span><span class=chip>Ligne en double</span>\n  <span class=chip>Ligne vide</span></div></div>\n <div class="gcard"><h3>🔤 Formats trompeurs</h3><div class="chips">\n  <span class=chip>Nombre en texte</span><span class=chip>Nombre avec unité</span>\n  <span class=chip>Date en texte</span><span class=chip>Date invraisemblable</span>\n  <span class=chip>Dates en désordre</span><span class=chip>Types mélangés</span></div></div>\n <div class="gcard"><h3>👻 Saisie incohérente</h3><div class="chips">\n  <span class=chip>Libellés incohérents</span><span class=chip>Espaces parasites</span></div></div>\n <div class="gcard"><h3>🧱 Structure</h3><div class="chips">\n  <span class=chip>Cellules fusionnées</span><span class=chip>En-tête en double</span>\n  <span class=chip>En-tête manquant</span><span class=chip>Données masquées</span>\n  <span class=chip>Données isolées</span></div></div>\n <div class="gcard"><h3>🎯 Valeurs anormales</h3><div class="chips">\n  <span class=chip>Valeur inhabituelle</span></div></div>\n</div>\n<p class="meta" style="margin-top:1rem">Vous pouvez aussi définir <b>vos propres règles</b> par colonne&nbsp;:\nvaleur obligatoire, plage min/max, liste de valeurs autorisées, format attendu…</p>\n</section>\n\n<section class="card">\n<h2>Et la sécurité&nbsp;?</h2>\n<div class="grid">\n <div class="gcard"><h3>🔒 Fichiers refusés d\'emblée</h3>\n  <p class="meta" style="margin:.1rem 0 .4rem">Un fichier dangereux est rejeté <b>avant</b> toute ouverture&nbsp;:</p>\n  <div class="chips"><span class=chip>Macro VBA</span><span class=chip>ActiveX / OLE</span>\n  <span class=chip>Liens externes</span><span class=chip>Faux .xlsx</span>\n  <span class=chip>Fichier piégé (zip-bomb)</span></div></div>\n <div class="gcard"><h3>👁️ Signalés sans bloquer</h3>\n  <div class="chips"><span class=chip>Feuille masquée</span><span class=chip>Formule volatile</span>\n  <span class=chip>Formule réseau</span><span class=chip>Chemin réseau (UNC)</span></div></div>\n</div>\n</section>\n\n<section class="card">\n<h2>Vos garanties</h2>\n<ul class="plain">\n <li>✅ <b>100&nbsp;% hors-ligne</b> — aucune donnée ne sort de votre machine.</li>\n <li>✅ <b>Aucune formule évaluée</b>, <b>aucune macro exécutée</b>, <b>aucun lien externe suivi</b>.</li>\n <li>✅ Votre fichier <b>n\'est jamais modifié</b>.</li>\n <li>✅ En cas de doute, l\'outil <b>refuse</b> plutôt que de prendre un risque.</li>\n</ul>\n</section>\n\n<p class="actions"><a class="btn btn--primary" href="/">Analyser un fichier</a></p>\n</main>\n</body></html>'''

_CONTROL_CHARS = re.compile(r'[\x00-\x1f\x7f]')


@app.middleware('http')
async def _limit_body_size(request: Request, call_next):
    """Rejette tôt (avant tout parsing) un corps dont Content-Length dépasse la limite.

    Défense anti-DoS complémentaire au contrôle de streaming : un upload honnête
    surdimensionné est refusé sans être bufferisé. Pour les requêtes chunkées sans
    Content-Length, la borne en streaming et un reverse-proxy restent la défense.
    """
    content_length = request.headers.get('content-length')
    if content_length is not None:
        try:
            declared = int(content_length)
        except ValueError:
            return JSONResponse({'detail': 'content-length invalide'}, status_code=400)
        if declared > MAX_UPLOAD_BYTES:
            return JSONResponse({'detail': f'corps > {MAX_UPLOAD_BYTES} octets'},
                                status_code=413)
    return await call_next(request)


def get_store() -> JobStore:
    return JobStore()


def _safe_filename(name: str | None) -> str:
    """Nom de fichier original sûr : basename, sans caractères de contrôle, borné.

    Neutralise les sauts de ligne / caractères de contrôle (injection dans les
    rapports Markdown) et tout composant de chemin. N'est jamais utilisé comme
    chemin réel (l'input est toujours `<job_dir>/input.xlsx`).
    """
    cleaned = _CONTROL_CHARS.sub('', name or '')
    base = Path(cleaned).name or 'upload.bin'
    return base[:255] or 'upload.bin'


@app.get('/', response_class=HTMLResponse)
async def index() -> HTMLResponse:
    return HTMLResponse(UPLOAD_HTML)


@app.get('/guide', response_class=HTMLResponse)
async def guide() -> HTMLResponse:
    return HTMLResponse(GUIDE_HTML)


def _cleanup_job(store: JobStore, job: Job) -> None:
    """Supprime l'input partiel après échec d'upload (jamais le job.json)."""
    try:
        store.input_path(job).unlink(missing_ok=True)
    except OSError:
        pass


async def _stream_to_file(file: UploadFile, dest: Path) -> int:
    """Écrit l'upload vers `dest` par morceaux, borné à MAX_UPLOAD_BYTES.

    Lève HTTPException(413) si la taille dépasse la limite (avant toute analyse).
    """
    total = 0
    with open(dest, 'wb') as out:
        while True:
            chunk = await file.read(_CHUNK)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413,
                                    detail=f'fichier > {MAX_UPLOAD_BYTES} octets')
            out.write(chunk)
    return total


def _load_report(store: JobStore, job: Job) -> dict | None:
    path = store.path(job, REPORT_JSON)
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None


def _wants_html(request: Request) -> bool:
    """Navigateur (formulaire) -> page HTML ; client API -> JSON."""
    accept = request.headers.get('accept', '')
    return 'text/html' in accept and 'application/json' not in accept.split(',')[0]


# ---- Libellés français (mapping intégral récupéré du binaire) ------------------------
_CAT_FR = {
    'classification_marking': 'Marquage classifié',
    'restricted_marking': 'Diffusion restreinte',
    'pii_nir': 'NIR (n° sécu)',
    'pii_iban': 'Coordonnées bancaires (IBAN)',
    'pii_card': 'Carte bancaire en clair',
    'secret_exposed': 'Secret / mot de passe en clair',
    'pii_email': 'Adresses e-mail',
    'pii_phone': 'Numéros de téléphone',
    'invalid_email': 'E-mail invalide',
    'invalid_nir': 'NIR mal formé',
    'invalid_iban': 'IBAN mal formé',
    'invalid_phone': 'Téléphone invalide',
    'placeholder_value': 'Valeur fictive (N/A…)',
    'date_impossible': 'Date impossible',
    'date_us_format': 'Date au format US',
    'date_future': 'Date dans le futur',
    'date_implausible': 'Date trop ancienne',
    'field_incoherence': 'Incohérence entre champs',
    'negative_value': 'Valeur négative anormale',
    'table_offset': 'En-tête décalé',
    'error_cell': 'Cellule en erreur',
    'number_stored_as_text': 'Nombre en texte',
    'number_with_unit': 'Nombre avec unité',
    'overwritten_formula': 'Formule écrasée',
    'inconsistent_formula': 'Formule incohérente',
    'range_omission': 'Plage incomplète',
    'missing_value': 'Valeur manquante',
    'mixed_types': 'Types mélangés',
    'duplicate_row': 'Ligne en double',
    'duplicate_key': "Doublon d'identifiant",
    'total_mismatch': 'Total incorrect',
    'merged_cells': 'Cellules fusionnées',
    'date_as_text': 'Date en texte',
    'date_out_of_range': 'Date invraisemblable',
    'date_out_of_order': 'Dates en désordre',
    'inconsistent_category': 'Libellés incohérents',
    'circular_reference': 'Référence circulaire',
    'missing_header': 'En-tête manquant',
    'duplicate_header': 'En-tête en double',
    'hidden_data': 'Données masquées',
    'outlier': 'Valeur inhabituelle',
    'whitespace_issue': 'Espaces parasites',
    'empty_row': 'Ligne vide',
    'formula_references_empty': 'Formule sur cellule vide',
    'stray_data': 'Données isolées',
    'rule_violation': 'Règle non respectée',
    'analysis_incomplete': 'Analyse incomplète (fail-closed)',
    'sensitive_schema': 'Schéma sensible vide',
    'hidden_sheet': 'Feuille masquée',
    'veryhidden_sheet': 'Feuille très masquée',
    'volatile_formula': 'Formule volatile',
    'network_formula': 'Formule réseau',
    'unc_path': 'Chemin réseau (UNC)',
}


def _cat_label(code: str) -> str:
    """Libellé français lisible ; repli sur le code brut si non traduit."""
    return _CAT_FR.get(code, code)


def _cat_cell(code: str) -> str:
    """Cellule de tableau : libellé FR affiché, code machine en infobulle."""
    return (f'<td title="{html.escape(code, quote=True)}">'
            f'{html.escape(_cat_label(code))}</td>')


_SEV_ORDER = {'CRITICAL': 0, 'HIGH': 1, 'MEDIUM': 2, 'LOW': 3, 'INFO': 4}


def _sev_pill(severity: str | None) -> str:
    sev = severity or 'INFO'
    return f'<span class="sev sev--{html.escape(sev, quote=True)}">{html.escape(sev)}</span>'


def _by_severity(findings: list[dict]) -> list[dict]:
    return sorted(findings, key=lambda f: _SEV_ORDER.get(f.get('severity', 'INFO'), 4))


def _sev_counts(findings: list[dict]) -> str:
    counts: dict[str, int] = {}
    for f in findings:
        sev = f.get('severity', 'INFO')
        counts[sev] = counts.get(sev, 0) + 1
    return ' '.join(f'{_sev_pill(sev)} {counts[sev]}'
                    for sev in _SEV_ORDER if sev in counts)


def _security_table(findings: list[dict]) -> str:
    if not findings:
        return '<p class="empty">Aucun constat de sécurité.</p>'
    rows = []
    for f in _by_severity(findings):
        rows.append(
            '<tr>'
            f'<td>{_sev_pill(f.get("severity"))}</td>'
            f'{_cat_cell(str(f.get("category", "")))}'
            f'<td>{html.escape(str(f.get("sheet") or "—"))}</td>'
            f'<td class="mono">{html.escape(str(f.get("cell") or "—"))}</td>'
            f'<td>{html.escape(str(f.get("title") or f.get("description") or ""))}</td>'
            '</tr>')
    return ('<div class="tablewrap"><table><thead><tr>'
            '<th>Gravité</th><th>Type</th><th>Feuille</th><th>Cellule</th><th>Détail</th>'
            '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div>')


def _quality_table(findings: list[dict]) -> str:
    if not findings:
        return '<p class="empty">Aucune anomalie de qualité détectée. ✔</p>'
    rows = []
    for f in _by_severity(findings):
        rows.append(
            '<tr>'
            f'<td>{_sev_pill(f.get("severity"))}</td>'
            f'<td class="conf">{html.escape(str(f.get("confidence") or ""))}</td>'
            f'{_cat_cell(str(f.get("category", "")))}'
            f'<td>{html.escape(str(f.get("sheet") or "—"))}</td>'
            f'<td class="mono">{html.escape(str(f.get("cell") or "—"))}</td>'
            f'<td>{html.escape(str(f.get("title") or ""))}</td>'
            '</tr>')
    return ('<div class="tablewrap"><table><thead><tr>'
            '<th>Gravité</th><th>Confiance</th><th>Type</th><th>Feuille</th>'
            '<th>Cellule</th><th>Détail</th>'
            '</tr></thead><tbody>' + ''.join(rows) + '</tbody></table></div>')


_VERDICT_FR = {'CERTIFIE': 'CERTIFIÉ', 'A_VERIFIER': 'À VÉRIFIER', 'BLOQUE': 'BLOQUÉ'}


def _verdict_banner(report: dict) -> str:
    verdict = report.get('verdict') or {}
    code = verdict.get('verdict')
    if not code:
        return ''
    code = html.escape(str(code), quote=True)
    grade = html.escape(str(verdict.get('grade', '')))
    reason = html.escape(str(verdict.get('reason', '')))
    label = _VERDICT_FR.get(code, code)
    sheets = ''.join(
        f'<span class="sheetv">{html.escape(str(s.get("sheet", "")))} '
        f'<span class="v v--{html.escape(str(s.get("verdict", "")), quote=True)}">'
        f'{_VERDICT_FR.get(str(s.get("verdict")), str(s.get("verdict")))}</span></span>'
        for s in verdict.get('sheets', []))
    banner = (f'<div class="verdict verdict--{code}">'
              f'<div class="grade">{grade}</div>'
              f'<div><div class="vlabel">{label}</div>'
              f'<div class="vreason">{reason}</div></div></div>')
    if sheets:
        banner += f'<div class="sheets">{sheets}</div>'
    return banner


def _render_result_html(job: Job, report: dict | None) -> str:
    job_id = html.escape(job.job_id, quote=True)
    status = job.status.value
    ok = status not in ('REJECTED', 'FAILED')
    status_pill = (f'<span class="status status--{"ok" if ok else "bad"}">'
                   f'{html.escape(status)}</span>')
    body: list[str] = []
    body.append(f'<h1>Résultat d\'analyse {status_pill}</h1>')
    if report:
        fm = report.get('file_metadata') or {}
        body.append(f'<p class="meta">{html.escape(str(fm.get("filename_original", "")))}'
                    f' · {fm.get("size_bytes", "?")} octets · SHA-256 '
                    f'<code>{html.escape(str(fm.get("sha256", ""))[:16])}…</code>'
                    f' · job <code>{job_id}</code></p>')
        pre = report.get('preflight_summary') or {}
        if not pre.get('passed', False):
            code = html.escape(str(pre.get('reason_code') or ''))
            detail = html.escape(str(pre.get('detail') or ''))
            parts = ''.join(f'<li><code>{html.escape(p)}</code></li>'
                            for p in (pre.get('blocking_parts') or [])[:10])
            body.append('<section class="card"><h2>Fichier refusé au préflight</h2>'
                        f'<p><b>{code}</b> — {detail}</p>'
                        + (f'<ul class="plain">{parts}</ul>' if parts else '')
                        + '</section>')
        else:
            body.append(_verdict_banner(report))
            findings = report.get('findings') or []
            quality = report.get('data_quality') or []
            risk = report.get('risk') or {}
            body.append('<section class="card"><h2>Sécurité '
                        f'({len(findings)} constats · score {risk.get("score", 0)})</h2>'
                        f'<p class="meta">{_sev_counts(findings)}</p>'
                        f'{_security_table(findings)}</section>')
            body.append('<section class="card"><h2>Qualité des données '
                        f'({len(quality)} anomalies)</h2>'
                        f'<p class="meta">{_sev_counts(quality)}</p>'
                        f'{_quality_table(quality)}</section>')
        body.append('<p class="actions">'
                    f'<a class="btn" href="/reports/{job_id}.json">Rapport JSON</a>'
                    f'<a class="btn" href="/reports/{job_id}.md">Rapport Markdown</a>'
                    f'<a class="btn" href="/reports/{job_id}.csv">Anomalies CSV</a>'
                    '<a class="btn btn--primary" href="/">Analyser un autre fichier</a>'
                    '</p>')
    else:
        body.append('<section class="card"><p>Rapport indisponible pour ce job.</p>'
                    '</section>')
    return (f'<!doctype html>\n<html lang="fr"><head><meta charset="utf-8">\n'
            f'<title>Résultat — xls-analyzer</title>\n'
            f'<meta name="viewport" content="width=device-width, initial-scale=1">\n'
            f'<style>{_BASE_CSS}</style></head>\n<body>\n'
            f'<header class="topbar"><div class="wrap">\n'
            f' <span class="brand">xls-analyzer<span class="dot">.</span></span>\n'
            f' <a class="btn" href="/guide">Guide</a>\n'
            f'</div></header>\n<main class="wrap">\n' + '\n'.join(body)
            + '\n</main>\n</body></html>')


@app.post('/upload')
async def upload(request: Request, file: UploadFile,
                 store: JobStore = Depends(get_store)):
    filename = _safe_filename(file.filename)
    job = store.create_job(filename)
    try:
        try:
            await _stream_to_file(file, store.input_path(job))
        except HTTPException:
            _cleanup_job(store, job)
            raise
        except OSError as exc:
            _cleanup_job(store, job)
            raise HTTPException(status_code=507,
                                detail="écriture de l'upload impossible") from exc
    finally:
        await file.close()
    job = run_pipeline(store, job)
    report = _load_report(store, job)
    if _wants_html(request):
        return HTMLResponse(_render_result_html(job, report))
    return JSONResponse({'job_id': job.job_id, 'status': job.status.value,
                         'report': report})


@app.get('/jobs/{job_id}')
async def get_job(job_id: str, store: JobStore = Depends(get_store)) -> JSONResponse:
    try:
        job = store.load_job(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail='job introuvable')
    return JSONResponse(job.model_dump(mode='json'))


def _existing_report(store: JobStore, job_id: str, name: str) -> Path:
    try:
        job_dir = store.job_dir(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    path = job_dir / name
    if not path.exists():
        raise HTTPException(status_code=404, detail='rapport introuvable')
    return path


@app.get('/reports/{job_id}.json')
async def get_report_json(job_id: str,
                          store: JobStore = Depends(get_store)) -> JSONResponse:
    path = _existing_report(store, job_id, REPORT_JSON)
    return JSONResponse(json.loads(path.read_text(encoding='utf-8')))


@app.get('/reports/{job_id}.md', response_class=PlainTextResponse)
async def get_report_md(job_id: str,
                        store: JobStore = Depends(get_store)) -> PlainTextResponse:
    path = _existing_report(store, job_id, REPORT_MD)
    return PlainTextResponse(path.read_text(encoding='utf-8'))


def _csv_safe(value: object) -> str:
    """Anti-injection de formule CSV : préfixe une apostrophe si la valeur commence par
    = + - @ (ou tab/CR) — sinon Excel exécuterait un nom d'onglet « =HYPERLINK(...) »."""
    if value is None:
        return ''
    s = str(value)
    if s and s[0] in '=+-@\t\r':
        return "'" + s
    return s


@app.get('/reports/{job_id}.csv')
async def get_report_csv(job_id: str,
                         store: JobStore = Depends(get_store)) -> Response:
    """Export CSV des anomalies qualité (data_quality). UTF-8 BOM pour Excel."""
    path = _existing_report(store, job_id, REPORT_JSON)
    report = json.loads(path.read_text(encoding='utf-8'))
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(['sévérité', 'confiance', 'type', 'feuille', 'cellule',
                     'titre', 'message'])
    for f in report.get('data_quality', []):
        writer.writerow([
            _csv_safe(f.get('severity', '')),
            _csv_safe(f.get('confidence', '')),
            _csv_safe(f.get('category', '')),
            _csv_safe(f.get('sheet', '')),
            _csv_safe(f.get('cell', '')),
            _csv_safe(f.get('title', '')),
            _csv_safe(f.get('message', f.get('description', ''))),
        ])
    data = buf.getvalue().encode('utf-8-sig')
    return Response(content=data, media_type='text/csv; charset=utf-8',
                    headers={'Content-Disposition':
                             f'attachment; filename="qualite_{job_id}.csv"'})


@app.post('/jobs/{job_id}/recalc')
async def post_recalc(job_id: str,
                      store: JobStore = Depends(get_store)) -> JSONResponse:
    """Dépose un manifeste de recalc (PENDING_RECALC). Ne pilote aucun conteneur."""
    try:
        job = store.load_job(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail='job introuvable')
    try:
        job = recalc_client.request_recalc(store, job)
    except ValueError as exc:
        raise HTTPException(status_code=409, detail=str(exc))
    return JSONResponse({'job_id': job.job_id, 'status': job.status.value,
                         'recalc_request_path': job.recalc_request_path})


def run() -> None:
    import uvicorn
    uvicorn.run(app, host=BIND_HOST, port=8000)


if __name__ == '__main__':
    run()
