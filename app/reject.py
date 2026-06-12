# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas reject.disasm.txt, vérifiée opcode/opcode)
# Correction d'artefact pycdc : bases d'enum rétablies en (str, Enum) — le bytecode
# module-level charge `str` PUIS `Enum` (offsets 56-58 et 80-82). (Finding F-01)
# ====================================================================================
"""Codes de rejet, severites et resultat structure du preflight."""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class RejectCode(str, Enum):
    EXT_NOT_ALLOWED = 'EXT_NOT_ALLOWED'
    SIGNATURE_MISMATCH = 'SIGNATURE_MISMATCH'
    NOT_A_ZIP = 'NOT_A_ZIP'
    MISSING_OOXML_PART = 'MISSING_OOXML_PART'
    ZIP_PATH_TRAVERSAL = 'ZIP_PATH_TRAVERSAL'
    DUPLICATE_ENTRY = 'DUPLICATE_ENTRY'
    TOO_MANY_ENTRIES = 'TOO_MANY_ENTRIES'
    ENTRY_TOO_LARGE = 'ENTRY_TOO_LARGE'
    TOTAL_TOO_LARGE = 'TOTAL_TOO_LARGE'
    COMPRESSION_RATIO_EXCEEDED = 'COMPRESSION_RATIO_EXCEEDED'
    SHAREDSTRINGS_TOO_LARGE = 'SHAREDSTRINGS_TOO_LARGE'
    XML_DEPTH_EXCEEDED = 'XML_DEPTH_EXCEEDED'
    MACRO_VBAPROJECT = 'MACRO_VBAPROJECT'
    EXTERNAL_LINKS = 'EXTERNAL_LINKS'
    CONNECTIONS = 'CONNECTIONS'
    OLE_OBJECT = 'OLE_OBJECT'
    ACTIVEX = 'ACTIVEX'
    EMBEDDING = 'EMBEDDING'


class Severity(str, Enum):
    INFO = 'INFO'
    LOW = 'LOW'
    MEDIUM = 'MEDIUM'
    HIGH = 'HIGH'
    CRITICAL = 'CRITICAL'


@dataclass
class PreflightResult:
    """Resultat fail-closed du preflight.

    passed=True  -> aucun rejet, openpyxl autorise.
    passed=False -> reason_code renseigne, openpyxl interdit.
    """
    passed: bool
    reason_code: RejectCode | None = None
    severity: Severity | None = None
    detail: str = ''
    blocking_parts: list[str] = field(default_factory=list)

    @classmethod
    def ok(cls) -> 'PreflightResult':
        return cls(passed=True)

    @classmethod
    def reject(cls, code: RejectCode, severity: Severity = Severity.HIGH,
               detail: str = '', parts: list[str] | None = None) -> 'PreflightResult':
        return cls(passed=False, reason_code=code, severity=severity,
                   detail=detail, blocking_parts=parts or [])
