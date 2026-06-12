# ====================================================================================
# PROVENANCE : RECONSTRUCTION CORRIGÉE. pycdc restituait load_rules/_is_empty avec
# artefacts (`data.get('columns', data)` mutilé, comparaisons perdues) et avait détruit
# le corps de _violation_reason/apply_rules. Reconstitué depuis : constantes du
# bytecode (clés 'required', 'type', 'min', 'max', 'allowed', 'regex', 'max_length',
# 'severity' ; messages FR), rules.example.json livré avec le binaire, et la doctrine
# « une anomalie par (colonne, règle) » du docstring d'origine.
# ====================================================================================
"""Moteur de règles métier par colonne (rules.json, optionnel).

Format (cf. rules.example.json) :
{
  "columns": {
    "Salaire":   {"type": "number", "min": 0, "max": 1000000, "severity": "HIGH"},
    "Matricule": {"required": true, "regex": "^[A-Z]{2}\\\\d{6}$"},
    "Statut":    {"allowed": ["Actif", "Inactif"], "severity": "MEDIUM"}
  }
}
Une anomalie est émise PAR COLONNE en défaut (première cellule citée + décompte),
pour ne pas noyer le rapport. La sévérité vient de la règle, bornée à la liste valide.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import get_column_letter

from app.quality_analyzer import (QualityIssue, _data_rows, _dtype, _header_info,
                                  _is_metadata_sheet, _n, _redact)

_VALID_SEVERITY = frozenset({'INFO', 'LOW', 'MEDIUM', 'HIGH', 'CRITICAL'})


def load_rules(path: str | Path) -> dict[str, dict]:
    """Charge rules.json ; accepte {'columns': {...}} ou directement {...}.

    Fichier absent/illisible/malformé -> {} (les règles sont optionnelles, leur absence
    ne doit jamais bloquer l'analyse ; le pipeline journalise une limite le cas échéant).
    """
    try:
        data = json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    columns = data.get('columns', data)
    if not isinstance(columns, dict):
        return {}
    return {str(k): v for k, v in columns.items() if isinstance(v, dict)}


def _is_empty(value: object) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == '')


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _violation_reason(cell, rule: dict) -> str | None:
    """Motif (FR) si la cellule viole la règle, sinon None. Premier défaut rapporté."""
    value = cell.value
    if _is_empty(value):
        if rule.get('required'):
            return 'valeur requise manquante'
        return None
    rtype = rule.get('type')
    if rtype in ('number', 'integer'):
        if _dtype(cell) != 'n' or not _is_number(value):
            return f'attendu un nombre, trouvé « {value} »'
        if rtype == 'integer' and float(value) != int(value):
            return f'attendu un entier, trouvé « {value} »'
    elif rtype == 'date':
        if not getattr(cell, 'is_date', False):
            return f'attendu une date, trouvé « {value} »'
    elif rtype == 'string':
        if not isinstance(value, str):
            return f'attendu du texte, trouvé « {value} »'
    if 'min' in rule and _is_number(value) and value < rule['min']:
        return f'{value} < minimum autorisé ({rule["min"]})'
    if 'max' in rule and _is_number(value) and value > rule['max']:
        return f'{value} > maximum autorisé ({rule["max"]})'
    if 'allowed' in rule and isinstance(rule['allowed'], list):
        if value not in rule['allowed'] and str(value) not in map(str, rule['allowed']):
            return f'« {value} » hors de la liste autorisée'
    if 'regex' in rule and isinstance(value, str):
        try:
            if not re.fullmatch(str(rule['regex']), value):
                return f'« {value} » ne respecte pas le format attendu'
        except re.error:
            return None
    if 'max_length' in rule and isinstance(value, str) \
            and len(value) > int(rule['max_length']):
        return f'texte de {len(value)} caractères > {rule["max_length"]} autorisés'
    return None


def apply_rules(path_or_wb, rules: dict[str, dict], *, deadline=None) -> list[QualityIssue]:
    """Applique les règles aux colonnes dont l'en-tête correspond exactement à une clé."""
    if not rules:
        return []
    if isinstance(path_or_wb, Workbook):
        wb, should_close = path_or_wb, False
    else:
        wb = load_workbook(path_or_wb, read_only=False, data_only=False, keep_vba=False)
        should_close = True
    try:
        issues: list[QualityIssue] = []
        for ws in wb.worksheets:
            if deadline is not None:
                import time
                if time.monotonic() > deadline:
                    raise TimeoutError('budget temps règles métier dépassé')
            if _is_metadata_sheet(ws):
                continue
            hdr, headers = _header_info(ws)
            if not headers:
                continue
            body = _data_rows(ws)
            for col, name in headers.items():
                rule = rules.get(name)
                if not rule:
                    continue
                severity = str(rule.get('severity', 'MEDIUM')).upper()
                if severity not in _VALID_SEVERITY:
                    severity = 'MEDIUM'
                violations: list[tuple[int, str]] = []
                for r in body:
                    reason = _violation_reason(ws.cell(r, col), rule)
                    if reason:
                        violations.append((r, reason))
                if not violations:
                    continue
                first_row, first_reason = violations[0]
                loc = f'{get_column_letter(col)}{first_row}'
                issues.append(QualityIssue(
                    kind='rule_violation', sheet=ws.title, cell=loc,
                    severity=severity, confidence='high',
                    message=f'Colonne «{name}» : '
                            f'{_n(len(violations), "cellule en défaut", "cellules en défaut")} '
                            f'sur la règle métier (ex. {loc} : {first_reason})',
                    evidence=_redact(name)))
        return issues
    finally:
        if should_close:
            wb.close()
