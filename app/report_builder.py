# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas report_builder.disasm.txt) pour :
# collect_tool_versions, compute_file_metadata, finding_from_static/_preflight,
# static_findings_to_models, quality_issues_to_findings, build_static_summary,
# _verdict_of, compute_verdict, build_report, apply_recalc_summary, write_report,
# _SEVERITY_BY_KIND (intégral), _VERDICT_REASON, _MASK_STATIC_EVIDENCE, DEFAULT_LIMITS.
# render_markdown : sections et libellés repris des constantes du bytecode, assemblage
# (sauts de ligne, ordre intra-section) reconstruit — fidèle en CONTENU, pas garanti
# au caractère près.
# ====================================================================================
"""Constructeur du rapport final : fusionne préflight, analyse statique, qualité,
sensible, règles et recalc en un `Report` validé, puis le sérialise en JSON et
Markdown (écritures atomiques).

Doctrine : `risk` n'agrège que les findings de SÉCURITÉ ; les anomalies de contenu
(`data_quality`) portent le verdict CERTIFIÉ / À VÉRIFIER / BLOQUÉ.
"""
from __future__ import annotations

import hashlib
from collections import Counter
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from app.atomic_io import write_json_atomic, write_text_atomic
from app.models import (Confidence, FileMetadata, Finding, FindingSource,
                        PreflightSummary, Report, SheetVerdict,
                        StaticAnalysisSummary, Verdict, VerdictSummary)
from app.quality_analyzer import QualityIssue
from app.reject import PreflightResult, Severity
from app.risk_scoring import score_findings
from app.sensitive import mask_sensitive_filename, mask_sensitive_text
from app.services.static_analyzer import Finding as StaticFinding

DEFAULT_LIMITS = [
    "Analyse statique uniquement : aucune formule évaluée, aucune macro exécutée, "
    "aucun lien externe suivi.",
    "Les détections PII/marquages sont heuristiques : elles réduisent le risque "
    "d'oubli mais ne remplacent pas la revue humaine.",
    "Les anomalies de qualité à confiance basse sont des signaux, pas des preuves.",
    "Le verdict porte sur le contenu visible du classeur ; il ne vaut pas décision "
    "d'homologation.",
]

_MASK_STATIC_EVIDENCE = frozenset({'network_formula', 'unc_path'})

_SEVERITY_BY_KIND: dict[str, str] = {
    'classification_marking': 'CRITICAL',
    'pii_nir': 'CRITICAL',
    'pii_iban': 'CRITICAL',
    'pii_card': 'CRITICAL',
    'secret_exposed': 'CRITICAL',
    'restricted_marking': 'HIGH',
    'invalid_nir': 'HIGH',
    'invalid_iban': 'HIGH',
    'field_incoherence': 'HIGH',
    'date_impossible': 'HIGH',
    'total_mismatch': 'HIGH',
    'overwritten_formula': 'HIGH',
    'circular_reference': 'HIGH',
    'pii_email': 'MEDIUM',
    'pii_phone': 'MEDIUM',
    'invalid_email': 'MEDIUM',
    'sensitive_schema': 'MEDIUM',
    'invalid_phone': 'MEDIUM',
    'date_future': 'MEDIUM',
    'date_us_format': 'MEDIUM',
    'date_implausible': 'MEDIUM',
    'negative_value': 'MEDIUM',
    'mixed_types': 'MEDIUM',
    'missing_value': 'MEDIUM',
    'inconsistent_category': 'MEDIUM',
    'duplicate_key': 'MEDIUM',
    'merged_cells': 'MEDIUM',
    'date_as_text': 'MEDIUM',
    'date_out_of_range': 'MEDIUM',
    'range_omission': 'MEDIUM',
    'inconsistent_formula': 'MEDIUM',
    'number_stored_as_text': 'MEDIUM',
    'number_with_unit': 'MEDIUM',
    'whitespace_issue': 'LOW',
    'duplicate_row': 'LOW',
    'empty_row': 'LOW',
    'outlier': 'LOW',
    'date_out_of_order': 'LOW',
    'formula_references_empty': 'LOW',
    'placeholder_value': 'LOW',
    'table_offset': 'LOW',
    'stray_data': 'LOW',
    'missing_header': 'LOW',
    'duplicate_header': 'LOW',
    'hidden_data': 'LOW',
}

_VERDICT_REASON: dict[Verdict, str] = {
    Verdict.BLOQUE: 'Donnée classifiée ou à caractère personnel sensible présente.',
    Verdict.A_VERIFIER: 'Anomalies à corriger avant diffusion.',
    Verdict.CERTIFIE: 'Aucune anomalie bloquante détectée.',
}


def collect_tool_versions() -> dict[str, str]:
    out: dict[str, str] = {}
    for pkg in ('openpyxl', 'fastapi', 'pydantic', 'defusedxml', 'jsonschema'):
        try:
            out[pkg] = version(pkg)
        except PackageNotFoundError:
            continue
    return out


def compute_file_metadata(filename_original: str, path: str | Path) -> FileMetadata:
    """SHA-256 + taille en streaming (jamais tout le fichier en mémoire)."""
    h = hashlib.sha256()
    size = 0
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
            size += len(chunk)
    return FileMetadata(
        filename_original=mask_sensitive_filename(filename_original),
        extension=Path(filename_original).suffix.lower(),
        size_bytes=size,
        sha256=h.hexdigest())


def finding_from_static(sf: StaticFinding, idx: int = 0) -> Finding:
    evidence = sf.evidence
    if sf.category in _MASK_STATIC_EVIDENCE:
        evidence = '(cible réseau masquée)'
    return Finding(
        id=f'static-{idx:04d}',
        severity=Severity(sf.severity),
        category=sf.category,
        title=sf.title,
        source=FindingSource.STATIC,
        sheet=mask_sensitive_text(sf.sheet) if sf.sheet else None,
        cell=sf.cell,
        evidence=mask_sensitive_text(evidence),
        confidence=Confidence.HIGH,
        recommendation='Vérifier la cellule citée ; supprimer ou justifier le contenu.')


def finding_from_preflight(pf: PreflightResult, idx: int = 0) -> Finding:
    return Finding(
        id=f'preflight-{idx:04d}',
        severity=pf.severity or Severity.HIGH,
        category=(pf.reason_code.value.lower() if pf.reason_code else 'preflight'),
        title='Fichier rejeté au préflight',
        source=FindingSource.PREFLIGHT,
        description=pf.detail,
        evidence=mask_sensitive_text(', '.join(pf.blocking_parts[:5])),
        confidence=Confidence.HIGH,
        reason_code=pf.reason_code,
        recommendation='Fournir un .xlsx sain, sans contenu actif ni structure '
                       'pathologique.')


def static_findings_to_models(static_findings: list[StaticFinding]) -> list[Finding]:
    return [finding_from_static(sf, i) for i, sf in enumerate(static_findings)]


def quality_issues_to_findings(issues: list[QualityIssue]) -> list[Finding]:
    """Mappe les anomalies qualité (dataclass) vers des Finding (source=quality).

    La sévérité est recalibrée selon `_SEVERITY_BY_KIND` (hiérarchie métier), sauf pour
    les types non listés (error_cell, rule_violation) qui gardent leur sévérité
    contextuelle.
    """
    return [Finding(
        id=f'quality-{i:04d}',
        severity=Severity(_SEVERITY_BY_KIND.get(q.kind, q.severity)),
        category=mask_sensitive_text(q.kind),
        title=mask_sensitive_text(q.message),
        source=FindingSource.QUALITY,
        sheet=mask_sensitive_text(q.sheet),
        cell=mask_sensitive_text(q.cell),
        evidence=mask_sensitive_text(q.evidence),
        confidence=Confidence(q.confidence),
        recommendation='Vérifier la cellule citée ; corriger la saisie ou la formule.')
        for i, q in enumerate(issues)]


def build_static_summary(findings: list[Finding]) -> StaticAnalysisSummary:
    by_sev = Counter(f.severity.value for f in findings)
    by_cat = Counter(f.category for f in findings)
    return StaticAnalysisSummary(ran=True, total_findings=len(findings),
                                 by_severity=dict(by_sev), by_category=dict(by_cat))


def _verdict_of(findings: list[Finding]) -> tuple[Verdict, str]:
    """(verdict, note) selon la plus haute sévérité présente."""
    sevs = {f.severity.value for f in findings}
    if 'CRITICAL' in sevs:
        return (Verdict.BLOQUE, 'D')
    if 'HIGH' in sevs:
        return (Verdict.A_VERIFIER, 'C')
    if 'MEDIUM' in sevs:
        return (Verdict.A_VERIFIER, 'B')
    return (Verdict.CERTIFIE, 'A')


def compute_verdict(data_quality: list[Finding],
                    sheet_names: list[str] = ()) -> VerdictSummary:
    """Verdict global + par feuille à partir des constats qualité/sensibles.

    `sheet_names` = feuilles de données analysées (hors métadonnées) pour lister aussi
    les feuilles propres en CERTIFIÉ.
    """
    by_sheet: dict[str, list[Finding]] = {mask_sensitive_text(n): []
                                          for n in sheet_names}
    for f in data_quality:
        by_sheet.setdefault(mask_sensitive_text(f.sheet) or '—', []).append(f)
    sheets: list[SheetVerdict] = []
    for name in sorted(by_sheet):
        fs = by_sheet[name]
        v, g = _verdict_of(fs)
        sheets.append(SheetVerdict(
            sheet=name, verdict=v, grade=g,
            counts=dict(Counter(f.severity.value for f in fs))))
    gv, gg = _verdict_of(data_quality)
    return VerdictSummary(verdict=gv, grade=gg, reason=_VERDICT_REASON[gv],
                          sheets=sheets)


def build_report(job_id: str, filename_original: str, status, created_at: str,
                 completed_at: str | None, file_metadata: FileMetadata | None,
                 preflight_summary: PreflightSummary, findings: list[Finding],
                 static_summary: StaticAnalysisSummary | None = None,
                 recalc_summary=None, data_quality: list[Finding] | None = None,
                 verdict: VerdictSummary | None = None,
                 limits: list[str] | None = None,
                 tool_versions: dict[str, str] | None = None) -> Report:
    """Assemble un Report complet et valide (risque agrégé inclus).

    `risk` n'agrège que les findings de SÉCURITÉ ; les anomalies qualité
    (`data_quality`) sont rapportées à part (audience et nature différentes).
    `verdict` = couche contenu (CERTIFIÉ/À VÉRIFIER/BLOQUÉ), n'affecte pas le préflight.
    """
    return Report(
        job_id=job_id,
        filename_original=mask_sensitive_filename(filename_original),
        status=status,
        created_at=created_at,
        completed_at=completed_at,
        file_metadata=file_metadata,
        preflight_summary=preflight_summary,
        static_analysis_summary=static_summary,
        recalc_summary=recalc_summary,
        risk=score_findings(findings),
        verdict=verdict,
        findings=findings,
        data_quality=list(data_quality) if data_quality is not None else [],
        limits=list(limits) if limits is not None else list(DEFAULT_LIMITS),
        tool_versions=(tool_versions if tool_versions is not None
                       else collect_tool_versions()))


def apply_recalc_summary(report: Report, recalc_summary, status, completed_at) -> Report:
    """Fusionne un résumé de recalc dans un rapport existant (nouveau Report)."""
    return report.model_copy(update={'recalc_summary': recalc_summary,
                                     'status': status,
                                     'completed_at': completed_at})


_MD_VERDICT = {Verdict.CERTIFIE: 'CERTIFIÉ', Verdict.A_VERIFIER: 'À VÉRIFIER',
               Verdict.BLOQUE: 'BLOQUÉ'}


def render_markdown(report: Report) -> str:
    """Rapport Markdown lisible hors-ligne. Contenu = celui du JSON, rien de plus."""
    r = report
    lines: list[str] = []
    lines.append(f'# Rapport d\'analyse — {r.filename_original}')
    lines.append('')
    lines.append(f'- **Job** : `{r.job_id}`')
    lines.append(f'- **Statut** : {r.status.value}')
    lines.append(f'- **Créé** : {r.created_at}')
    if r.completed_at:
        lines.append(f'- **Terminé** : {r.completed_at}')
    if r.file_metadata:
        fm = r.file_metadata
        lines.append(f'- **Fichier** : {fm.filename_original} ({fm.size_bytes} octets)')
        lines.append(f'- **SHA-256** : `{fm.sha256}`')
    lines.append('')
    if r.verdict:
        lines.append(f'## Verdict : {_MD_VERDICT[r.verdict.verdict]} '
                     f'(note {r.verdict.grade})')
        lines.append('')
        lines.append(r.verdict.reason)
        if r.verdict.sheets:
            lines.append('')
            lines.append('| Feuille | Verdict | Note |')
            lines.append('| --- | --- | --- |')
            for s in r.verdict.sheets:
                lines.append(f'| {s.sheet} | {_MD_VERDICT[s.verdict]} | {s.grade} |')
        lines.append('')
    lines.append('## Préflight')
    lines.append('')
    if r.preflight_summary.passed:
        lines.append('Conteneur sain : aucun contenu actif, structure conforme.')
    else:
        code = (r.preflight_summary.reason_code.value
                if r.preflight_summary.reason_code else '?')
        lines.append(f'**REJETÉ** — {code} : {r.preflight_summary.detail}')
        for part in r.preflight_summary.blocking_parts[:10]:
            lines.append(f'- `{part}`')
    lines.append('')
    lines.append(f'## Risque (sécurité) : score {r.risk.score}'
                 + (f', niveau {r.risk.level.value}' if r.risk.level else ''))
    lines.append('')
    if r.findings:
        lines.append('| Sévérité | Type | Feuille | Cellule | Titre |')
        lines.append('| --- | --- | --- | --- | --- |')
        for f in r.findings:
            lines.append(f'| {f.severity.value} | {f.category} | {f.sheet or "—"} | '
                         f'{f.cell or "—"} | {f.title or f.description} |')
    else:
        lines.append('Aucun constat de sécurité.')
    lines.append('')
    lines.append(f'## Qualité des données ({len(r.data_quality)} anomalies)')
    lines.append('')
    if r.data_quality:
        lines.append('| Sévérité | Confiance | Type | Feuille | Cellule | Détail |')
        lines.append('| --- | --- | --- | --- | --- | --- |')
        for f in r.data_quality:
            lines.append(f'| {f.severity.value} | {f.confidence.value} | {f.category} '
                         f'| {f.sheet or "—"} | {f.cell or "—"} | {f.title} |')
    else:
        lines.append('Aucune anomalie de qualité détectée.')
    lines.append('')
    if r.recalc_summary:
        lines.append('## Recalc')
        lines.append('')
        lines.append(f'- statut : {r.recalc_summary.status}')
        if r.recalc_summary.detail:
            lines.append(f'- détail : {r.recalc_summary.detail}')
        lines.append('')
    lines.append('## Limites')
    lines.append('')
    for limit in r.limits:
        lines.append(f'- {limit}')
    lines.append('')
    if r.tool_versions:
        lines.append('## Versions')
        lines.append('')
        for k in sorted(r.tool_versions):
            lines.append(f'- {k} {r.tool_versions[k]}')
        lines.append('')
    return '\n'.join(lines)


def write_report(report: Report, json_path: str | Path, md_path: str | Path) -> None:
    """Écrit report.json puis report.md, atomiquement."""
    write_json_atomic(json_path, report.model_dump(mode='json'))
    write_text_atomic(md_path, render_markdown(report))
