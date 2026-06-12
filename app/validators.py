# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas validators.disasm.txt, 2 306 lignes,
# 12 fonctions vérifiées opcode/opcode — y compris seuils 0.6, frozenset de masquage
# NIR/IBAN, tolérance « année courante + 1 » des dates futures).
# ====================================================================================
"""Validité & cohérence (Phase 3) — au-delà de la présence de PII, on contrôle que
les données sont *correctes* :

- identifiants : e-mail / NIR (clé) / IBAN (mod 97) / téléphone malformés, valeurs fictives ;
- dates en texte : impossibles (30/02), futures, format US (MM/JJ) ;
- cohérence inter-champs : date de recrutement antérieure à la naissance ;
- bornes : valeur négative dans une colonne qui ne devrait pas l'être ;
- erreurs de formule : lues sur les **valeurs en cache** (data_only=True), car en
  data_only=False openpyxl expose la formule, pas son résultat #VALUE!/#DIV/0!/#REF!.

Lecture seule, aucune évaluation de formule. Redaction stricte (NIR/IBAN jamais recopiés).
"""
from __future__ import annotations

import datetime
import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import get_column_letter

from app.quality_analyzer import (EXCEL_ERRORS, QualityIssue, _data_rows, _header_info,
                                  _is_metadata_sheet, _redact)
from app.sensitive import (_HDR_EMAIL, _HDR_IBAN, _HDR_NIR, _HDR_PHONE, _IBAN_SHAPE,
                           _NIR_SHAPE, _ratio)

_EMAIL_RE = re.compile(r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$')
_PHONE_RE = re.compile(r'^(?:\+33|0)\s*[1-9](?:[ .\-]?\d{2}){4}$')
_FR_DATE = re.compile(r'^\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})\s*$')
_PLACEHOLDERS = {'—', 'néant', 'à compléter', 'n/a', 'a completer', '-', '.', '?', 'x',
                 'na', 'nc', 'ras', 'tbd', 'xxx', 'none', 'null', 'neant'}
_NL = r'(?<![a-zà-ÿ])'
_HDR_DATE = re.compile(_NL + r'(date|naissance|naiss|recrut|embauche|[ée]ch[ée]ance|'
                       r'd[ée]but|fin)', re.I)
_HDR_BIRTH = re.compile(_NL + r'(naissance|naiss|dob)', re.I)
_HDR_HIRE = re.compile(_NL + r'(recrut|embauche|entr[ée]e|prise de poste)', re.I)
_HDR_NONNEG = re.compile(_NL + r'(salaire|montant|prix|co[uû]t|quantit|qte|stock|effectif|'
                         r'nombre|nb_|[âa]ge|taux|valeur|surface|poids|distance|dur[ée]e|'
                         r'heure|capacit)', re.I)
_IBAN_LEN = {'FR': 27, 'MC': 27, 'BE': 16, 'DE': 22, 'ES': 24, 'IT': 27,
             'GB': 22, 'CH': 21, 'LU': 20, 'PT': 25, 'NL': 18}


def _nir_valid(raw: str) -> bool:
    """Structure d'un NIR : 13 ou 15 chiffres, chiffre de sexe plausible (1-4,7,8).

    On ne valide PAS la clé (mod 97) : sur des données synthétiques bien formées mais à
    clé aléatoire, cela noierait le rapport. Le contrôle de clé reste activable au besoin.
    """
    s = re.sub(r'\s', '', raw).upper().replace('2A', '19').replace('2B', '18')
    return len(s) in (13, 15) and s.isdigit() and s[0] in '123478'


def _iban_valid(raw: str) -> bool:
    """Structure d'un IBAN : 2 lettres + 2 chiffres + BBAN, longueur correcte par pays."""
    s = re.sub(r'\s', '', raw).upper()
    if not re.fullmatch(r'[A-Z]{2}\d{2}[A-Z0-9]+', s):
        return False
    expected = _IBAN_LEN.get(s[:2])
    if expected:
        return len(s) == expected
    return 15 <= len(s) <= 34


def _email_valid(v: str) -> bool:
    return bool(_EMAIL_RE.fullmatch(v))


def _phone_valid(v: str) -> bool:
    return bool(_PHONE_RE.fullmatch(v))


def _is_placeholder(v: str) -> bool:
    return v.strip().lower() in _PLACEHOLDERS


def _date_problem(raw: str):
    """(kind, libellé) si la date-texte est problématique, sinon None."""
    m = _FR_DATE.match(raw)
    if not m:
        return None
    d, mo, y = int(m[1]), int(m[2]), int(m[3])
    if y < 100:
        y += 2000 if y < 50 else 1900
    if mo > 12 and d <= 12:
        return ('date_us_format', 'format américain probable (MM/JJ au lieu de JJ/MM)')
    try:
        when = datetime.date(y, mo, d)
        if y < 1900:
            return ('date_implausible', f'année {y} invraisemblable')
        if when.year > datetime.date.today().year + 1:
            return ('date_future', f'date dans le futur (année {y})')
    except ValueError:
        return ('date_impossible', 'date impossible (jour/mois invalides)')
    return None


def _col_cells(ws, col, body) -> list[tuple[int, str]]:
    out: list[tuple[int, str]] = []
    for r in body:
        v = ws.cell(r, col).value
        if v is None or not str(v).strip():
            continue
        out.append((r, str(v).strip()))
    return out


def _validate_identifiers(ws, hdr, headers, body) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    for col, name in headers.items():
        cells = _col_cells(ws, col, body)
        if not cells:
            continue
        raw = [v for _, v in cells]
        if _HDR_NIR.search(name) or _ratio(raw, _NIR_SHAPE) >= 0.6:
            kind, sev, validator, label = 'invalid_nir', 'HIGH', _nir_valid, 'NIR mal formé'
        elif _HDR_IBAN.search(name) or _ratio(raw, _IBAN_SHAPE) >= 0.6:
            kind, sev, validator, label = ('invalid_iban', 'HIGH', _iban_valid,
                                           'IBAN mal formé')
        elif _HDR_EMAIL.search(name) or _ratio(raw, _EMAIL_RE) >= 0.6:
            kind, sev, validator, label = ('invalid_email', 'MEDIUM', _email_valid,
                                           'e-mail invalide')
        elif _HDR_PHONE.search(name) or _ratio(raw, _PHONE_RE) >= 0.6:
            kind, sev, validator, label = ('invalid_phone', 'MEDIUM', _phone_valid,
                                           'téléphone invalide')
        else:
            continue
        for r, v in cells:
            loc = f'{get_column_letter(col)}{r}'
            if _is_placeholder(v):
                issues.append(QualityIssue(
                    kind='placeholder_value', sheet=ws.title, cell=loc,
                    severity='LOW', confidence='high',
                    message=f'Valeur fictive « {v} » dans « {name} » (cellule {loc}) — '
                            'donnée non renseignée.',
                    evidence=_redact(v)))
                continue
            if validator(v):
                continue
            masked = kind in {'invalid_nir', 'invalid_iban'}
            show = '' if masked else f' « {v} »'
            issues.append(QualityIssue(
                kind=kind, sheet=ws.title, cell=loc, severity=sev, confidence='high',
                message=f'{label.capitalize()}{show} dans « {name} » (cellule {loc}).',
                evidence='' if masked else _redact(v)))
    return issues


def _validate_dates(ws, hdr, headers, body) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    for col, name in headers.items():
        if not _HDR_DATE.search(name):
            continue
        for r, v in _col_cells(ws, col, body):
            prob = _date_problem(v)
            if not prob:
                continue
            kind, why = prob
            loc = f'{get_column_letter(col)}{r}'
            sev = 'HIGH' if kind == 'date_impossible' else 'MEDIUM'
            issues.append(QualityIssue(
                kind=kind, sheet=ws.title, cell=loc, severity=sev, confidence='high',
                message=f'Date « {v} » dans « {name} » : {why} (cellule {loc}).',
                evidence=_redact(v)))
    return issues


def _check_coherence(ws, hdr, headers, body) -> list[QualityIssue]:
    birth = next((c for c, n in headers.items() if _HDR_BIRTH.search(n)), None)
    hire = next((c for c, n in headers.items() if _HDR_HIRE.search(n)), None)
    if not birth or not hire:
        return []
    issues: list[QualityIssue] = []
    for r in body:
        b = _as_date(ws.cell(r, birth).value)
        h = _as_date(ws.cell(r, hire).value)
        if not b or not h or not h < b:
            continue
        loc = f'{get_column_letter(hire)}{r}'
        issues.append(QualityIssue(
            kind='field_incoherence', sheet=ws.title, cell=loc,
            severity='HIGH', confidence='high',
            message=f'Incohérence ligne {r} : « {headers[hire]} » ({h.year}) antérieur à '
                    f'« {headers[birth]} » ({b.year}).',
            evidence=''))
    return issues


def _check_ranges(ws, hdr, headers, body) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    for col, name in headers.items():
        if not _HDR_NONNEG.search(name):
            continue
        for r in body:
            v = ws.cell(r, col).value
            if not isinstance(v, (int, float)) or isinstance(v, bool) or not v < 0:
                continue
            loc = f'{get_column_letter(col)}{r}'
            issues.append(QualityIssue(
                kind='negative_value', sheet=ws.title, cell=loc,
                severity='MEDIUM', confidence='high',
                message=f'Valeur négative ({v}) dans « {name} » (cellule {loc}) — '
                        'attendu ≥ 0.',
                evidence=_redact(v)))
    return issues


def _as_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str):
        m = _FR_DATE.match(value)
        if m:
            d, mo, y = int(m[1]), int(m[2]), int(m[3])
            if y < 100:
                y += 2000 if y < 50 else 1900
            try:
                return datetime.date(y, mo, d)
            except ValueError:
                return None
    return None


def _coerce_workbook(path_or_wb, *, data_only: bool):
    if isinstance(path_or_wb, Workbook):
        return (path_or_wb, False)
    return (load_workbook(path_or_wb, read_only=False, data_only=data_only,
                          keep_vba=False), True)


def _check_deadline(deadline=None) -> None:
    if deadline is None:
        return
    import time
    if time.monotonic() > deadline:
        raise TimeoutError('budget temps validateurs dépassé')


def analyze_cached_errors(path_or_wb, *, deadline=None) -> list[QualityIssue]:
    """Erreurs de FORMULE lues sur les valeurs en cache (data_only=True).

    Si le fichier n'a pas de cache (créé sans tableur), data_only renvoie None : aucun
    faux constat (fail-safe). Les cellules-erreur littérales sont déjà vues par ailleurs.
    """
    wb, should_close = _coerce_workbook(path_or_wb, data_only=True)
    try:
        issues: list[QualityIssue] = []
        for ws in wb.worksheets:
            _check_deadline(deadline)
            if _is_metadata_sheet(ws):
                continue
            for row in ws.iter_rows():
                _check_deadline(deadline)
                for cell in row:
                    v = cell.value
                    if not (cell.data_type == 'e'
                            or (isinstance(v, str) and v in EXCEL_ERRORS)):
                        continue
                    issues.append(QualityIssue(
                        kind='error_cell', sheet=ws.title, cell=cell.coordinate,
                        severity='HIGH', confidence='high',
                        message=f'Formule en erreur {v} (cellule {cell.coordinate}) — '
                                'résultat invalide.',
                        evidence=_redact(v)))
        return issues
    finally:
        if should_close:
            wb.close()


def analyze_validity(path_or_wb, *, deadline=None) -> list[QualityIssue]:
    """Tous les contrôles de validité/cohérence sur les feuilles de données."""
    wb, should_close = _coerce_workbook(path_or_wb, data_only=False)
    try:
        issues: list[QualityIssue] = []
        for ws in wb.worksheets:
            _check_deadline(deadline)
            if _is_metadata_sheet(ws):
                continue
            hdr, headers = _header_info(ws)
            if not headers:
                continue
            body = _data_rows(ws)
            issues += _validate_identifiers(ws, hdr, headers, body)
            issues += _validate_dates(ws, hdr, headers, body)
            issues += _check_coherence(ws, hdr, headers, body)
            issues += _check_ranges(ws, hdr, headers, body)
        return issues
    finally:
        if should_close:
            wb.close()
