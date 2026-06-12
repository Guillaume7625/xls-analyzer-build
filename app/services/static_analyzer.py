# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas services_static_analyzer.disasm.txt).
# Dataclass Finding et flux à deux passes vérifiés opcode/opcode.
# ====================================================================================
"""Analyse statique openpyxl. Lancee UNIQUEMENT si le preflight passe.

Ne JAMAIS evaluer une formule, executer une macro ni resoudre un lien externe.
Les preuves (evidence) sont tronquees a EVIDENCE_MAX_CHARS par defaut.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook

from app.config import EVIDENCE_MAX_CHARS

_VOLATILE = ('NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT')
_NETWORK = ('WEBSERVICE', 'HYPERLINK', 'RTD')
_UNC_RE = re.compile(r'\\\\[^\\]+\\')


def _redact(text: object) -> str:
    text = str(text)
    if len(text) <= EVIDENCE_MAX_CHARS:
        return text
    return text[:EVIDENCE_MAX_CHARS] + '...[redacted]'


@dataclass
class Finding:
    severity: str
    category: str
    sheet: str | None
    cell: str | None
    title: str
    evidence: str


def _coerce_workbook(path_or_wb):
    if isinstance(path_or_wb, Workbook):
        return (path_or_wb, False)
    return (load_workbook(path_or_wb, read_only=False, data_only=False, keep_vba=False), True)


def _check_deadline(deadline=None) -> None:
    if deadline is None:
        return
    import time
    if time.monotonic() > deadline:
        raise TimeoutError('budget temps analyse statique dépassé')


def analyze(path_or_wb, *, deadline=None) -> list[Finding]:
    findings: list[Finding] = []
    wb, should_close = _coerce_workbook(path_or_wb)
    try:
        for ws in wb.worksheets:
            _check_deadline(deadline)
            if ws.sheet_state == 'hidden':
                findings.append(Finding('MEDIUM', 'hidden_sheet', ws.title, None,
                                        'Feuille cachee', _redact(ws.title)))
            elif ws.sheet_state == 'veryHidden':
                findings.append(Finding('HIGH', 'veryhidden_sheet', ws.title, None,
                                        'Feuille tres cachee (veryHidden)', _redact(ws.title)))
        for ws in wb.worksheets:
            _check_deadline(deadline)
            for row in ws.iter_rows():
                _check_deadline(deadline)
                for cell in row:
                    val = cell.value
                    if not (isinstance(val, str) and val.startswith('=')):
                        continue
                    upper = val.upper()
                    for fn in _VOLATILE:
                        if re.search(rf'\b{fn}\s*\(', upper):
                            findings.append(Finding('MEDIUM', 'volatile_formula', ws.title,
                                                    cell.coordinate,
                                                    f'Fonction volatile {fn}', _redact(val)))
                    for fn in _NETWORK:
                        if re.search(rf'\b{fn}\s*\(', upper):
                            findings.append(Finding('HIGH', 'network_formula', ws.title,
                                                    cell.coordinate,
                                                    f'Fonction reseau/externe {fn}', _redact(val)))
                    if _UNC_RE.search(val):
                        findings.append(Finding('HIGH', 'unc_path', ws.title, cell.coordinate,
                                                'Chemin UNC dans une formule', _redact(val)))
    finally:
        if should_close:
            wb.close()
    return findings
