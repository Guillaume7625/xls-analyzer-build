# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas risk_scoring.disasm.txt).
# ====================================================================================
"""Scoring de risque : agrège une liste de Finding en un RiskScore.

Score = somme pondérée des sévérités. `level` = plus haute sévérité présente.
Pondérations strictement croissantes ; aucune dépendance aux seuils de config.py.
"""
from __future__ import annotations

from collections import Counter

from app.models import Finding, RiskScore
from app.reject import Severity

_ORDER: dict[Severity, int] = {
    Severity.INFO: 0,
    Severity.LOW: 1,
    Severity.MEDIUM: 2,
    Severity.HIGH: 3,
    Severity.CRITICAL: 4,
}
_WEIGHT: dict[Severity, int] = {
    Severity.INFO: 0,
    Severity.LOW: 1,
    Severity.MEDIUM: 3,
    Severity.HIGH: 7,
    Severity.CRITICAL: 15,
}


def score_findings(findings: list[Finding]) -> RiskScore:
    """Agrège les sévérités des findings en un RiskScore déterministe."""
    counts = Counter(f.severity.value for f in findings)
    score = sum(_WEIGHT[f.severity] for f in findings)
    level = None
    if findings:
        level = max((f.severity for f in findings), key=lambda s: _ORDER[s])
    return RiskScore(score=score, level=level, counts=dict(counts))
