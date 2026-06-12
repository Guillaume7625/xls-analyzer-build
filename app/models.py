# ====================================================================================
# PROVENANCE : RECONSTRUCTION CORRIGÉE. pycdc avait perdu des champs et mutilé les
# validateurs. Champs rétablis par preuve croisée sur les KW_NAMES du bytecode :
#   - Finding(id, severity, category, title, …)   ← report_builder.disasm (CALL 10/15)
#   - FileMetadata(filename_original, extension, size_bytes, sha256) ← compute_file_metadata
#   - SheetVerdict(sheet, verdict, grade, counts) / VerdictSummary(verdict, …) ← compute_verdict
#   - Report(job_id, …, recalc_summary, …)        ← build_report (KW_NAMES, 15 champs)
#   - Job(job_id, status, created_at, updated_at, filename_original, …) ← job_store/main
# (Findings F-02, F-03 du rapport.)
# ====================================================================================
"""Modèles Pydantic canoniques : Finding, Report, Job (+ sous-modèles, enums).

Source de vérité des artefacts JSON. Réutilise les enums du noyau vérifié
(`app.reject.Severity`, `app.reject.RejectCode`) — jamais redéfinis ici.
Les schémas JSON de `schemas/` sont générés à partir de ces modèles.
"""
from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, ConfigDict, Field, field_validator

from app.reject import RejectCode, Severity
from app.services.static_analyzer import _redact


class JobStatus(str, Enum):
    """États du job (cf. table de transitions docs/SPEC.md)."""
    CREATED = 'CREATED'
    PREFLIGHT_RUNNING = 'PREFLIGHT_RUNNING'
    STATIC_RUNNING = 'STATIC_RUNNING'
    STATIC_DONE = 'STATIC_DONE'
    PENDING_RECALC = 'PENDING_RECALC'
    RECALC_RUNNING = 'RECALC_RUNNING'
    RECALC_DONE = 'RECALC_DONE'
    DONE = 'DONE'
    REJECTED = 'REJECTED'
    FAILED = 'FAILED'


class FindingSource(str, Enum):
    PREFLIGHT = 'preflight'
    STATIC = 'static'
    RECALC = 'recalc'
    QUALITY = 'quality'


class Confidence(str, Enum):
    LOW = 'low'
    MEDIUM = 'medium'
    HIGH = 'high'


class Verdict(str, Enum):
    """Verdict de contenu (couche au-dessus du préflight fail-closed)."""
    CERTIFIE = 'CERTIFIE'
    A_VERIFIER = 'A_VERIFIER'
    BLOQUE = 'BLOQUE'


_STRICT = ConfigDict(extra='forbid')


class Finding(BaseModel):
    """Constat unitaire. `evidence` est toujours redacted/bornée (fail-closed)."""
    model_config = _STRICT

    id: str
    severity: Severity
    category: str
    title: str = ''
    source: FindingSource
    sheet: str | None = None
    cell: str | None = None
    range: str | None = None
    description: str = ''
    evidence: str = ''
    impact: str = ''
    recommendation: str = ''
    confidence: Confidence = Confidence.MEDIUM
    reason_code: RejectCode | None = None

    @field_validator('evidence')
    @classmethod
    def _cap_evidence(cls, v: str) -> str:
        return _redact(v)


class FileMetadata(BaseModel):
    model_config = _STRICT

    filename_original: str
    extension: str
    size_bytes: int
    sha256: str


class PreflightSummary(BaseModel):
    model_config = _STRICT

    passed: bool
    reason_code: RejectCode | None = None
    severity: Severity | None = None
    detail: str = ''
    blocking_parts: list[str] = Field(default_factory=list)


class StaticAnalysisSummary(BaseModel):
    model_config = _STRICT

    ran: bool
    total_findings: int = 0
    by_severity: dict[str, int] = Field(default_factory=dict)
    by_category: dict[str, int] = Field(default_factory=dict)


class RecalcSummary(BaseModel):
    model_config = _STRICT

    status: str
    detail: str = ''


class RecalcResult(BaseModel):
    """Artefact recalc_result.json écrit par le worker, validé par le runner.

    Schéma strict (extra interdit) : tout artefact non conforme => job FAILED.
    """
    model_config = _STRICT

    status: str
    detail: str = ''
    recalculated_sha256: str | None = None
    duration_seconds: float | None = None


class RiskScore(BaseModel):
    """Agrégat de risque calculé par app.risk_scoring à partir des findings."""
    model_config = _STRICT

    score: int
    level: Severity | None = None
    counts: dict[str, int] = Field(default_factory=dict)


class SheetVerdict(BaseModel):
    """Verdict de contenu pour une feuille."""
    model_config = _STRICT

    sheet: str
    verdict: Verdict
    grade: str
    counts: dict[str, int] = Field(default_factory=dict)


class VerdictSummary(BaseModel):
    """Verdict global + détail par feuille (couche contenu, n'affecte pas le préflight)."""
    model_config = _STRICT

    verdict: Verdict
    grade: str
    reason: str = ''
    sheets: list[SheetVerdict] = Field(default_factory=list)


class Report(BaseModel):
    """Rapport final fusionné, conforme à schemas/report.schema.json."""
    model_config = _STRICT

    job_id: str
    filename_original: str
    status: JobStatus
    created_at: str
    completed_at: str | None = None
    file_metadata: FileMetadata | None = None
    preflight_summary: PreflightSummary
    static_analysis_summary: StaticAnalysisSummary | None = None
    recalc_summary: RecalcSummary | None = None
    risk: RiskScore
    verdict: VerdictSummary | None = None
    findings: list[Finding] = Field(default_factory=list)
    data_quality: list[Finding] = Field(default_factory=list)
    limits: list[str] = Field(default_factory=list)
    tool_versions: dict[str, str] = Field(default_factory=dict)


class Job(BaseModel):
    """État persistant du job dans /run/xls-analyzer/jobs/<job_id>/job.json."""
    model_config = _STRICT

    job_id: str
    status: JobStatus
    created_at: str
    updated_at: str
    filename_original: str
    input_path: str
    static_report_path: str | None = None
    recalc_request_path: str | None = None
    recalc_result_path: str | None = None
    report_json_path: str | None = None
    report_md_path: str | None = None
    error: str | None = None
