# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE des fragments pycdc (_dedup_issues corrigé d'un
# double-négatif artefact, _incomplete_analysis, _cell_count, _deadline, _check_budget,
# _reject, _fail, _write_report restitués) + RECONSTRUCTION SPEC-BASED des corps
# perdus (run_pipeline, _write_incomplete_static_done — fin) selon : la table de
# transitions du docstring d'origine, les constantes env (XLS_MAX_ANALYSIS_CELLS=200000,
# XLS_ANALYSIS_TIMEOUT_SECONDS=10), et les imports du module (qui dictent les phases :
# validate_file → preflight_zip → analyze → quality + validity + sensitive + rules).
# ====================================================================================
"""Orchestrateur de job : applique la table de transitions de docs/SPEC.md.

CREATED -> PREFLIGHT_RUNNING -> (REJECTED | STATIC_RUNNING)
STATIC_RUNNING -> (FAILED | STATIC_DONE)

Fail-closed : toute exception non prévue -> FAILED. Aucune évaluation de formule,
aucun recalc ici (le recalc est piloté ultérieurement par le runner). Zéro conteneur.
"""
from __future__ import annotations

import os
import time
from datetime import datetime, timezone

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from app import rules_engine
from app.atomic_io import write_json_atomic
from app.job_store import REPORT_JSON, REPORT_MD, STATIC_REPORT, JobStore
from app.models import (Confidence, FileMetadata, Finding, FindingSource, Job,
                        JobStatus, PreflightSummary, Report, Severity, Verdict,
                        VerdictSummary)
from app.reject import PreflightResult
from app.quality_analyzer import analyze_quality, data_sheet_names
from app.sensitive import analyze_sensitive
from app.validators import analyze_cached_errors, analyze_validity
from app.report_builder import (build_report, build_static_summary,
                                compute_file_metadata, compute_verdict,
                                finding_from_preflight, quality_issues_to_findings,
                                static_findings_to_models, write_report)
from app.services.file_validation import validate_file
from app.services.preflight_zip import preflight_zip
from app.services.static_analyzer import analyze

MAX_ANALYSIS_CELLS = int(os.environ.get('XLS_MAX_ANALYSIS_CELLS', '200000'))
ANALYSIS_TIMEOUT_SECONDS = float(os.environ.get('XLS_ANALYSIS_TIMEOUT_SECONDS', '10'))
RULES_PATH = os.environ.get('XLS_RULES_PATH', '')


class AnalysisLimitExceeded(Exception):
    """Le classeur dépasse les bornes locales d'analyse complète."""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _dedup_issues(issues):
    """Dédoublonne par (kind, feuille, cellule) — ex. error_cell vu en littéral ET en cache."""
    seen = set()
    out = []
    for i in issues:
        key = (i.kind, i.sheet, i.cell)
        if key in seen:
            continue
        seen.add(key)
        out.append(i)
    return out


def _incomplete_analysis(exc: Exception):
    """Fail-closed : une analyse de contenu incomplète => fichier NON certifiable.

    On n'expose que le TYPE d'exception (jamais son message, qui pourrait contenir une
    valeur de cellule sensible). Un incident d'analyse force « BLOQUÉ » : du contenu
    classifié ou de la PII a pu ne pas être scanné.
    """
    finding = Finding(
        id='quality-0000', severity=Severity.CRITICAL, category='analysis_incomplete',
        title=f'Analyse de contenu incomplète ({type(exc).__name__}) — du contenu '
              'sensible a pu ne pas être scanné. Fichier non certifiable.',
        source=FindingSource.QUALITY, confidence=Confidence.HIGH,
        recommendation="Ne pas diffuser ; rejeter ou réanalyser jusqu'à analyse "
                       'complète.')
    verdict = VerdictSummary(
        verdict=Verdict.BLOQUE, grade='D',
        reason='Analyse de contenu incomplète — non certifiable (fail-closed).',
        sheets=[])
    return ([finding], verdict)


def _cell_count(wb: Workbook) -> int:
    total = 0
    for ws in wb.worksheets:
        total += max(ws.max_row, 0) * max(ws.max_column, 0)
    return total


def _deadline(start: float) -> float:
    return start + ANALYSIS_TIMEOUT_SECONDS


def _check_budget(deadline: float) -> None:
    if time.monotonic() > deadline:
        raise TimeoutError('budget temps du pipeline dépassé')


def _write_report_files(store: JobStore, job: Job, report: Report,
                        status: JobStatus, **extra) -> Job:
    jpath = store.path(job, REPORT_JSON)
    mpath = store.path(job, REPORT_MD)
    write_report(report, jpath, mpath)
    return store.update(job, status=status, report_json_path=str(jpath),
                        report_md_path=str(mpath), **extra)


def _reject(store: JobStore, job: Job, pf: PreflightResult,
            file_metadata: FileMetadata) -> Job:
    """Rejet preflight : report.json/md avec le reason_code, statut REJECTED."""
    preflight_summary = PreflightSummary(
        passed=False, reason_code=pf.reason_code, severity=pf.severity,
        detail=pf.detail, blocking_parts=pf.blocking_parts)
    report = build_report(
        job_id=job.job_id, filename_original=job.filename_original,
        status=JobStatus.REJECTED, created_at=job.created_at, completed_at=_now(),
        file_metadata=file_metadata, preflight_summary=preflight_summary,
        findings=[finding_from_preflight(pf)], static_summary=None)
    return _write_report_files(store, job, report, JobStatus.REJECTED)


def _fail(store: JobStore, job: Job, error: str,
          file_metadata: FileMetadata | None) -> Job:
    """Incident non récupérable après obtention des métadonnées : FAILED + rapport."""
    report = build_report(
        job_id=job.job_id, filename_original=job.filename_original,
        status=JobStatus.FAILED, created_at=job.created_at, completed_at=_now(),
        file_metadata=file_metadata,
        preflight_summary=PreflightSummary(passed=False, detail=error),
        findings=[], static_summary=None)
    return _write_report_files(store, job, report, JobStatus.FAILED, error=error)


def _write_incomplete_static_done(store: JobStore, job: Job,
                                  file_metadata: FileMetadata,
                                  static_findings: list[Finding],
                                  static_summary, exc: Exception) -> Job:
    """Analyse statique OK mais analyse de contenu interrompue : STATIC_DONE BLOQUÉ."""
    quality_findings, verdict = _incomplete_analysis(exc)
    report = build_report(
        job_id=job.job_id, filename_original=job.filename_original,
        status=JobStatus.STATIC_DONE, created_at=job.created_at, completed_at=_now(),
        file_metadata=file_metadata, preflight_summary=PreflightSummary(passed=True),
        findings=static_findings, static_summary=static_summary,
        data_quality=quality_findings, verdict=verdict)
    return _write_report_files(store, job, report, JobStatus.STATIC_DONE)


def run_pipeline(store: JobStore, job: Job) -> Job:
    """Exécute le pipeline complet (hors recalc) et écrit les artefacts."""
    input_path = store.input_path(job)
    job = store.update(job, status=JobStatus.PREFLIGHT_RUNNING)

    try:
        file_metadata = compute_file_metadata(job.filename_original, input_path)
    except OSError as exc:
        return _fail(store, job, f'lecture du fichier impossible '
                                 f'({type(exc).__name__})', None)

    vf = validate_file(job.filename_original, input_path)
    if not vf.passed:
        return _reject(store, job, vf, file_metadata)
    pf = preflight_zip(input_path)
    if not pf.passed:
        return _reject(store, job, pf, file_metadata)

    job = store.update(job, status=JobStatus.STATIC_RUNNING)
    start = time.monotonic()
    deadline = _deadline(start)

    # --- Phase statique (sécurité). Échec ici = FAILED : on ne sait rien dire du
    # fichier, fournir un rapport « propre » serait mensonger.
    try:
        wb = load_workbook(input_path, read_only=False, data_only=False,
                           keep_vba=False)
    except Exception as exc:  # openpyxl lève des types variés sur fichier hostile
        return _fail(store, job, f'ouverture openpyxl impossible '
                                 f'({type(exc).__name__})', file_metadata)
    try:
        try:
            cells = _cell_count(wb)
            if cells > MAX_ANALYSIS_CELLS:
                raise AnalysisLimitExceeded(
                    f'{cells} cellules > {MAX_ANALYSIS_CELLS}')
            static_dc = analyze(wb, deadline=deadline)
        except Exception as exc:
            return _fail(store, job, f'analyse statique interrompue '
                                     f'({type(exc).__name__})', file_metadata)

        static_findings = static_findings_to_models(static_dc)
        static_summary = build_static_summary(static_findings)
        write_json_atomic(store.path(job, STATIC_REPORT),
                          static_summary.model_dump(mode='json'))
        job = store.update(job, static_report_path=str(store.path(job, STATIC_REPORT)))

        # --- Phase contenu (qualité / validité / sensible / règles).
        # Échec ici = STATIC_DONE mais BLOQUÉ (fail-closed : non certifiable).
        try:
            _check_budget(deadline)
            issues = list(analyze_quality(wb, deadline=deadline))
            issues += analyze_validity(wb, deadline=deadline)
            issues += analyze_cached_errors(input_path, deadline=deadline)
            issues += analyze_sensitive(input_path, deadline=deadline)
            if RULES_PATH:
                rules = rules_engine.load_rules(RULES_PATH)
                issues += rules_engine.apply_rules(wb, rules, deadline=deadline)
            issues = _dedup_issues(issues)
            data_quality = quality_issues_to_findings(issues)
            verdict = compute_verdict(data_quality, data_sheet_names(wb))
        except Exception as exc:
            return _write_incomplete_static_done(store, job, file_metadata,
                                                 static_findings, static_summary, exc)
    finally:
        wb.close()

    report = build_report(
        job_id=job.job_id, filename_original=job.filename_original,
        status=JobStatus.STATIC_DONE, created_at=job.created_at, completed_at=_now(),
        file_metadata=file_metadata, preflight_summary=PreflightSummary(passed=True),
        findings=static_findings, static_summary=static_summary,
        data_quality=data_quality, verdict=verdict)
    return _write_report_files(store, job, report, JobStatus.STATIC_DONE)
