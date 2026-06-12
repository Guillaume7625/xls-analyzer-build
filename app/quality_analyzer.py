# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas quality_analyzer.disasm.txt, 8 546 lignes
# lues fonction par fonction). Constantes vérifiées contre SURFACE_SECURITE.md et le
# module-level disasm (EXCEL_ERRORS, _ERROR_FR, _KEY_PLACEHOLDERS, _DETECTORS×23).
# Points où le flux exact a été condensé sans perte sémantique : marqués [≈].
# ====================================================================================
"""Contrôle qualité des données (.xlsx) — erreurs fréquentes et impactantes.

Détecteurs (openpyxl pur, lecture seule, aucune formule évaluée) :
  - error_cell            : cellules en erreur Excel (#REF! #DIV/0! #VALUE! #N/A …)
  - number_stored_as_text : nombre saisi en texte dans une colonne numérique (SUM l'ignore)
  - overwritten_formula   : formule effacée, remplacée par une saisie en dur
  - inconsistent_formula  : formule divergente d'un motif voisin cohérent
  - range_omission        : plage d'agrégat (SUM/…) qui semble omettre des données
                            adjacentes (type Reinhart-Rogoff) — confiance basse

Priorités issues de la recherche (Panko/EuSpRIG + cas réels). Haute précision :
on ne signale que sur signal net (les faux positifs érodent la confiance).
"""
from __future__ import annotations

import re
import statistics
from collections import Counter
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.formula.translate import Translator
from openpyxl.workbook.workbook import Workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from app.config import EVIDENCE_MAX_CHARS

EXCEL_ERRORS = frozenset({'#N/A', '#NUM!', '#REF!', '#CALC!', '#NAME?', '#NULL!',
                          '#DIV/0!', '#SPILL!', '#VALUE!', '#GETTING_DATA'})
_ERROR_SEVERITY = {'#N/A': 'MEDIUM'}
_ERROR_FR = {
    '#REF!': 'référence supprimée',
    '#DIV/0!': 'division par zéro',
    '#N/A': 'valeur introuvable',
    '#VALUE!': 'type de valeur incorrect',
    '#NAME?': 'nom de formule non reconnu',
    '#NUM!': 'nombre invalide',
    '#NULL!': 'intersection de plages vide',
    '#SPILL!': 'débordement de plage',
    '#CALC!': 'erreur de calcul',
    '#GETTING_DATA': 'données en cours de chargement',
}
_NUMERIC_TEXT = re.compile(r'^\s*-?\d{1,3}(?:[ \xa0]?\d{3})*(?:[.,]\d+)?\s*$')
_AGG_RANGE = re.compile(
    r'\b(SUM|AVERAGE|COUNT|COUNTA|MIN|MAX|PRODUCT|MEDIAN|STDEV|STDEVP|VAR)'
    r'\s*\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)', re.I)
_MIN_NUMERIC_IN_COL = 3
_NUMERIC_DOMINANCE = 0.6
_MOSTLY_FILLED = 0.7
_MIN_TABLE_ROWS = 5
_TYPE_DOMINANCE = 0.8
_TOTAL_LABEL = re.compile(r'\b(total|totaux|somme|sous-total|sous total|cumul|sum)\b', re.I)
_DATE_TEXT = re.compile(r'^\s*\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}\s*$')
_QUALIFIED_REF = re.compile(
    r"(?:'[^']*'|[A-Za-z0-9_.]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")
_CELL_REF = re.compile(r'(?<![A-Za-z])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?!\()')
_INVISIBLE = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\xa0\u200b\u200c\u200d\ufeff]')
_OUTLIER_ROBUST_Z = 6.0
_SINGLE_REF = re.compile(r'(?<![A-Za-z])\$?([A-Z]{1,3})\$?(\d+)(?!\s*:)(?!\()')


@dataclass
class QualityIssue:
    kind: str
    sheet: str
    cell: str
    severity: str
    confidence: str
    message: str
    evidence: str = ''


def _redact(value: object) -> str:
    text = str(value)
    if len(text) <= EVIDENCE_MAX_CHARS:
        return text
    return text[:EVIDENCE_MAX_CHARS] + '…'


def _n(count: int, singular: str, plural: str) -> str:
    """Accord singulier/pluriel : `_n(1, 'valeur', 'valeurs')` -> '1 valeur'."""
    return f'{count} {singular if count == 1 else plural}'


def _col_label(headers: dict[int, str], col: int) -> str:
    """« colonne « Nom » (lettre) » si l'en-tête est connu, sinon « colonne lettre »."""
    letter = get_column_letter(col)
    name = headers.get(col)
    if name:
        return f'colonne « {name} » ({letter})'
    return f'colonne {letter}'


def _dtype(cell: Cell) -> str | None:
    return getattr(cell, 'data_type', None)


def _is_formula(cell: Cell) -> bool:
    return _dtype(cell) == 'f'


def _is_constant(cell: Cell) -> bool:
    return cell.value is not None and _dtype(cell) != 'f'


def _looks_numeric_text(cell: Cell) -> bool:
    return (_dtype(cell) == 's' and isinstance(cell.value, str)
            and bool(_NUMERIC_TEXT.match(cell.value)))


def _is_number_with_unit(value: object) -> bool:
    """Vrai si `value` est un nombre habillé d'une unité/devise (ex. « 9.3 kL », « 57 356 € »)
    — donc non calculable. Faux pour un nombre pur, une date, du texte sans chiffre."""
    if not isinstance(value, str):
        return False
    s = value.strip()
    if not s or not any(ch.isdigit() for ch in s):
        return False
    core = re.sub(r'^[€$£]\s*', '', s)
    core = re.sub(r'\s*[A-Za-zµ%°/²³€$£]+\.?$', '', core)
    if core == s:
        return False
    return bool(_NUMERIC_TEXT.match(core))


def _expected_at(neighbor: Cell, target_coord: str) -> str | None:
    try:
        return Translator(str(neighbor.value),
                          origin=neighbor.coordinate).translate_formula(target_coord)
    except Exception:
        return None


def _consistent_expected(before: Cell, after: Cell, target_coord: str) -> str | None:
    """Formule attendue à target_coord si before/after partagent le même motif."""
    if not (_is_formula(before) and _is_formula(after)):
        return None
    exp_before = _expected_at(before, target_coord)
    exp_after = _expected_at(after, target_coord)
    if exp_before is not None and exp_before == exp_after:
        return exp_before
    return None


def _detect_error_cells(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (cell.data_type == 'e'
                    or (isinstance(v, str) and v in EXCEL_ERRORS)):
                continue
            issues.append(QualityIssue(
                kind='error_cell', sheet=ws.title, cell=cell.coordinate,
                severity=_ERROR_SEVERITY.get(str(v), 'HIGH'), confidence='high',
                message=f'Cellule en erreur Excel : {v}'
                        + (f' ({_ERROR_FR[str(v)]})' if str(v) in _ERROR_FR else '')
                        + ' — le calcul ne renvoie aucun résultat',
                evidence=_redact(v)))
    return issues


def _detect_number_as_text(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    headers = _header_row(ws)
    by_col: dict[int, list[Cell]] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            by_col.setdefault(cell.column, []).append(cell)
    for col, cells in by_col.items():
        numeric = sum(1 for c in cells if c.data_type == 'n')
        text_nums = [c for c in cells if _looks_numeric_text(c)]
        # [≈] Seuils du binaire : assez de numériques ET dominance numérique.
        if not text_nums or numeric < _MIN_NUMERIC_IN_COL:
            continue
        if numeric / max(1, len(cells)) < _NUMERIC_DOMINANCE:
            continue
        for c in text_nums:
            issues.append(QualityIssue(
                kind='number_stored_as_text', sheet=ws.title, cell=c.coordinate,
                severity='MEDIUM', confidence='high',
                message=f'Nombre stocké comme du texte dans la {_col_label(headers, col)}'
                        f' — ignoré par SUM/AVERAGE',
                evidence=_redact(c.value)))
    return issues


def _detect_formula_anomalies(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    flagged: set[str] = set()
    max_row, max_col = ws.max_row, ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value is None or cell.coordinate in flagged:
                continue
            expected = None
            if 1 < r < max_row:
                expected = _consistent_expected(ws.cell(r - 1, c), ws.cell(r + 1, c),
                                                cell.coordinate)
            if expected is None and 1 < c < max_col:
                expected = _consistent_expected(ws.cell(r, c - 1), ws.cell(r, c + 1),
                                                cell.coordinate)
            if expected is None:
                continue
            if _is_constant(cell):
                flagged.add(cell.coordinate)
                issues.append(QualityIssue(
                    kind='overwritten_formula', sheet=ws.title, cell=cell.coordinate,
                    severity='HIGH', confidence='medium',
                    message='Formule remplacée par une valeur tapée à la main '
                            f'(les cellules voisines calculent : {expected})',
                    evidence=_redact(cell.value)))
            elif _is_formula(cell) and str(cell.value) != expected:
                flagged.add(cell.coordinate)
                issues.append(QualityIssue(
                    kind='inconsistent_formula', sheet=ws.title, cell=cell.coordinate,
                    severity='MEDIUM', confidence='medium',
                    message=f'Formule différente de ses voisines (les autres font : {expected})',
                    evidence=_redact(cell.value)))
    return issues


def _detect_range_omission(ws) -> list[QualityIssue]:
    """Agrégat dont la plage semble s'arrêter juste avant des données adjacentes."""
    issues: list[QualityIssue] = []
    for row in ws.iter_rows():
        for cell in row:
            if not _is_formula(cell):
                continue
            for m in _AGG_RANGE.finditer(str(cell.value)):
                _, c1, r1, c2, r2 = m.groups()
                if c1 == c2:
                    nxt = ws[f'{c1}{int(r2) + 1}']
                elif r1 == r2:
                    nxt = ws.cell(row=int(r1), column=column_index_from_string(c2) + 1)
                else:
                    continue
                if nxt.data_type != 'n' or nxt.value is None:
                    continue
                issues.append(QualityIssue(
                    kind='range_omission', sheet=ws.title, cell=cell.coordinate,
                    severity='MEDIUM', confidence='low',
                    message=f'La plage {m.group(0)} oublie peut-être des données juste à côté '
                            f'({nxt.coordinate} contient une valeur non incluse)',
                    evidence=_redact(cell.value)))
    return issues


_HEADER_SCAN_DEPTH = 8


def _row_is_empty(ws, r: int, max_col: int) -> bool:
    return all(ws.cell(r, c).value in (None, '') for c in range(1, max_col + 1))


def _is_merged_title_row(ws, r: int) -> bool:
    """Ligne couverte par une fusion horizontale large (titre fusionné), pas un en-tête."""
    half = max(2, ws.max_column // 2)
    return any(mr.min_row == r and mr.max_row == r and (mr.max_col - mr.min_col + 1) >= half
               for mr in ws.merged_cells.ranges)


def _looks_like_header(cells) -> dict[int, str] | None:
    """{colonne: nom} si la ligne ressemble franchement à des en-têtes texte, sinon None."""
    nonempty = [c for c in cells if c.value is not None and str(c.value).strip() != '']
    str_cells = [c for c in nonempty if isinstance(c.value, str) and c.value.strip()]
    if not nonempty or len(str_cells) < max(1, int(len(nonempty) * 0.6)):
        return None
    return {c.column: str(c.value).strip() for c in str_cells}


def _header_info(ws) -> tuple[int, dict[int, str]]:
    """(index de la ligne d'en-tête, {colonne: nom}). (0, {}) si aucune table claire.

    Ignore en tête de feuille les titres fusionnés et les lignes vides, puis prend la
    première ligne « en-tête ». Si la première ligne pleine non-titre n'en est pas une,
    on conclut qu'il n'y a pas d'en-tête exploitable.
    """
    max_col = ws.max_column
    for r in range(1, min(ws.max_row, _HEADER_SCAN_DEPTH) + 1):
        if _row_is_empty(ws, r, max_col) or _is_merged_title_row(ws, r):
            continue
        mapping = _looks_like_header([ws.cell(r, c) for c in range(1, max_col + 1)])
        if mapping:
            return (r, mapping)
        return (0, {})
    return (0, {})


def _header_row(ws) -> dict[int, str]:
    """{colonne: nom} de la vraie ligne d'en-tête (compat : nombre de détecteurs l'utilisent)."""
    return _header_info(ws)[1]


def _kind_of(cell: Cell) -> str | None:
    if cell.value is None or (isinstance(cell.value, str) and cell.value == ''):
        return None
    if getattr(cell, 'is_date', False):
        return 'date'
    return {'n': 'nombre', 's': 'texte', 'b': 'booléen',
            'f': 'formule', 'e': 'erreur'}.get(_dtype(cell) or '', 'autre')


def _body_start(ws) -> int:
    """Première ligne de données : juste après la vraie ligne d'en-tête (ou ligne 2)."""
    hdr = _header_info(ws)[0]
    if hdr:
        return hdr + 1
    return 2


def _total_rows(ws) -> set[int]:
    """Lignes contenant un libellé de total/sous-total (à exclure de l'analyse par colonne :
    leurs valeurs sont des sommes, pas des données individuelles)."""
    max_col = ws.max_column
    return {r for r in range(_body_start(ws), ws.max_row + 1)
            if any(isinstance(ws.cell(r, c).value, str)
                   and _TOTAL_LABEL.search(ws.cell(r, c).value)
                   for c in range(1, max_col + 1))}


def _data_rows(ws) -> list[int]:
    """Indices des lignes de DONNÉES (corps hors en-tête et hors lignes de total)."""
    skip = _total_rows(ws)
    return [r for r in range(_body_start(ws), ws.max_row + 1) if r not in skip]


def _detect_missing_values(ws) -> list[QualityIssue]:
    """Colonne à dominante remplie (≥70 %) mais avec des trous = saisies manquantes."""
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)
    n = len(body)
    if not headers or n < _MIN_TABLE_ROWS:
        return issues
    for col, name in headers.items():
        empties = [r for r in body if _kind_of(ws.cell(r, col)) is None]
        filled = n - len(empties)
        if not empties or filled / n < _MOSTLY_FILLED:
            continue
        issues.append(QualityIssue(
            kind='missing_value', sheet=ws.title,
            cell=f'{get_column_letter(col)}{hdr}',
            severity='MEDIUM', confidence='medium',
            message=f'Colonne «{name}» : '
                    f'{_n(len(empties), "valeur manquante", "valeurs manquantes")} sur {n} '
                    f'(ex. {"ligne" if len(empties) == 1 else "lignes"} '
                    f'{", ".join(map(str, empties[:5]))})',
            evidence=_redact(name)))
    return issues


def _detect_mixed_types(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)
    if not headers or len(body) < _MIN_TABLE_ROWS:
        return issues
    for col, name in headers.items():
        kinds: Counter = Counter()
        for r in body:
            k = _kind_of(ws.cell(r, col))
            if k is None or k == 'formule':
                continue
            kinds[k] += 1
        total = sum(kinds.values())
        if total < _MIN_TABLE_ROWS or len(kinds) < 2:
            continue
        dominant, dom_n = kinds.most_common(1)[0]
        if dom_n / total < _TYPE_DOMINANCE:
            continue
        others = ', '.join(f'{cnt} en «{k}»' for k, cnt in kinds.items() if k != dominant)
        exc = _n(total - dom_n, "valeur d'un autre type", "valeurs d'un autre type")
        issues.append(QualityIssue(
            kind='mixed_types', sheet=ws.title,
            cell=f'{get_column_letter(col)}{hdr}',
            severity='MEDIUM', confidence='low',
            message=f'Colonne «{name}» : surtout des «{dominant}», mais {exc} ({others})',
            evidence=_redact(name)))
    return issues


def _detect_duplicate_rows(ws) -> list[QualityIssue]:
    """Lignes de données strictement identiques (doublon complet)."""
    issues: list[QualityIssue] = []
    headers = _header_row(ws)
    body = _data_rows(ws)
    if not headers or len(body) < _MIN_TABLE_ROWS:
        return issues
    cols = sorted(headers)
    seen: dict[tuple, int] = {}
    for r in body:
        key = tuple(ws.cell(r, c).value for c in cols)
        if all(v is None for v in key):
            continue
        if key in seen:
            issues.append(QualityIssue(
                kind='duplicate_row', sheet=ws.title, cell=f'A{r}',
                severity='LOW', confidence='medium',
                message=f'Ligne {r} identique à la ligne {seen[key]} (doublon)',
                evidence=''))
        else:
            seen[key] = r
    return issues


_HDR_KEY = re.compile(r'(?<![a-zà-ÿ])(id|identifiant|matricule|e-?mail|courriel|nir|iban|'
                      r'r[ée]f|code|num[ée]ro|uuid|siret|siren|immatriculation|login|compte)',
                      re.I)
_IDENTIFIER_SHAPE = re.compile(r'^[A-Za-z0-9._@\-+/]{3,}$')
_KEY_PLACEHOLDERS = frozenset({'—', 'néant', 'n/a', '-', '.', '?', 'x',
                               'na', 'nc', 'none', 'null', 'neant'})


def _detect_duplicate_keys(ws) -> list[QualityIssue]:
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)
    issues: list[QualityIssue] = []
    if not headers:
        return issues
    for col, name in headers.items():
        vals = [(r, str(ws.cell(r, col).value).strip())
                for r in body
                if ws.cell(r, col).value not in (None, '')
                and str(ws.cell(r, col).value).strip().casefold() not in _KEY_PLACEHOLDERS]
        if len(vals) < _MIN_TABLE_ROWS:
            continue
        distinct = len({v.casefold() for _, v in vals})
        token_like = sum(1 for _, v in vals if _IDENTIFIER_SHAPE.match(v)) / len(vals)
        dateish = sum(1 for _, v in vals if _DATE_TEXT.match(v)) / len(vals)
        fallback = (distinct / len(vals) >= 0.95 and token_like >= 0.8 and dateish < 0.5)
        if not (_HDR_KEY.search(name) or fallback):
            continue
        seen: dict[str, int] = {}
        pairs: list[tuple[int, int]] = []
        for r, v in vals:
            k = v.casefold()
            if k in seen:
                pairs.append((r, seen[k]))
            else:
                seen[k] = r
        if not pairs:
            continue
        ex = ' ; '.join(f'ligne {a} = ligne {b}' for a, b in pairs[:3])
        issues.append(QualityIssue(
            kind='duplicate_key', sheet=ws.title,
            cell=f'{get_column_letter(col)}{pairs[0][0]}',
            severity='MEDIUM', confidence='medium',
            message=f'Colonne «{name}» : '
                    f'{_n(len(pairs), "identifiant en double", "identifiants en double")}'
                    f" — même valeur sur plusieurs lignes, possible doublon d'enregistrement "
                    f'({ex}).',
            evidence=''))
    return issues


def _detect_total_mismatch(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    max_col = ws.max_column
    total_rows = sorted(_total_rows(ws))
    if len(total_rows) != 1:
        return issues
    for r in total_rows:
        label_cols = {c for c in range(1, max_col + 1)
                      if isinstance(ws.cell(r, c).value, str)
                      and _TOTAL_LABEL.search(ws.cell(r, c).value)}
        for c in range(1, max_col + 1):
            if c in label_cols:
                continue
            cell = ws.cell(r, c)
            if _dtype(cell) != 'n' or cell.value is None:
                continue
            block: list = []
            rr = r - 1
            while rr >= 1:
                up = ws.cell(rr, c)
                if _dtype(up) == 'n' and up.value is not None:
                    block.append(up.value)
                    rr -= 1
                else:
                    break
            if len(block) < 2:
                continue
            total = sum(block)
            if abs(total - cell.value) > max(1e-06, abs(cell.value) * 1e-09):
                issues.append(QualityIssue(
                    kind='total_mismatch', sheet=ws.title, cell=cell.coordinate,
                    severity='HIGH', confidence='medium',
                    message=f'Le total affiché ({cell.value}) ne correspond pas à la somme '
                            f'réelle des {len(block)} lignes au-dessus ({total})',
                    evidence=_redact(cell.value)))
    return issues


def _detect_merged_cells(ws) -> list[QualityIssue]:
    """Cellules fusionnées dans le corps du tableau (cassent tri/SUM/références)."""
    issues: list[QualityIssue] = []
    for rng in sorted(ws.merged_cells.ranges, key=str):
        if not rng.min_row >= 2:
            continue
        issues.append(QualityIssue(
            kind='merged_cells', sheet=ws.title, cell=str(rng),
            severity='MEDIUM', confidence='medium',
            message=f'Cellules fusionnées ({rng}) dans le tableau — gênent tris, filtres '
                    'et formules',
            evidence=''))
    return issues


def _detect_date_issues(ws) -> list[QualityIssue]:
    """Date stockée en texte (colonne date) + date hors plage plausible."""
    issues: list[QualityIssue] = []
    headers = _header_row(ws)
    body = _data_rows(ws)
    if not headers or len(body) < _MIN_TABLE_ROWS:
        return issues
    for col, name in headers.items():
        cells = [ws.cell(r, col) for r in body]
        real_dates = [c for c in cells if getattr(c, 'is_date', False)]
        text_dates = [c for c in cells if _dtype(c) == 's' and isinstance(c.value, str)
                      and _DATE_TEXT.match(c.value)]
        if real_dates and text_dates and len(real_dates) >= _MIN_NUMERIC_IN_COL:
            for c in text_dates:
                issues.append(QualityIssue(
                    kind='date_as_text', sheet=ws.title, cell=c.coordinate,
                    severity='MEDIUM', confidence='medium',
                    message=f'Date saisie comme du texte dans «{name}» — ni triable ni '
                            'calculable comme une vraie date',
                    evidence=_redact(c.value)))
        for c in real_dates:
            year = getattr(c.value, 'year', None)
            if year is None or 1900 <= year <= 2100:
                continue
            issues.append(QualityIssue(
                kind='date_out_of_range', sheet=ws.title, cell=c.coordinate,
                severity='MEDIUM', confidence='low',
                message=f'Date peu plausible (année {year}) dans «{name}» — probable faute '
                        'de saisie',
                evidence=_redact(c.value)))
    return issues


def _norm_label(value: object) -> str:
    return re.sub(r'\s+', ' ', str(value).strip()).casefold()


def _detect_inconsistent_categories(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)
    if not headers or len(body) < _MIN_TABLE_ROWS:
        return issues
    for col, name in headers.items():
        groups: dict[str, set[str]] = {}
        for r in body:
            c = ws.cell(r, col)
            if c.data_type != 's' or not isinstance(c.value, str) or not c.value.strip():
                continue
            groups.setdefault(_norm_label(c.value), set()).add(c.value)
        for variants in groups.values():
            if len(variants) <= 1:
                continue
            sample = sorted(variants)[:4]
            shown = ', '.join(f'« {v} »' for v in sample)
            issues.append(QualityIssue(
                kind='inconsistent_category', sheet=ws.title,
                cell=f'{get_column_letter(col)}{hdr}',
                severity='MEDIUM', confidence='medium',
                message=f'Colonne «{name}» : une même valeur écrite de plusieurs façons '
                        f'({shown}) — à uniformiser',
                evidence=_redact(name)))
    return issues


def _formula_refs(formula: str) -> set[str]:
    """Coordonnées (A1) référencées par une formule, MÊME feuille uniquement."""
    cleaned = _QUALIFIED_REF.sub(' ', formula)
    refs: set[str] = set()
    for m in _CELL_REF.finditer(cleaned):
        c1, r1, c2, r2 = m.groups()
        if c2 is None:
            refs.add(f'{c1}{int(r1)}')
            continue
        ci1, ci2 = sorted((column_index_from_string(c1), column_index_from_string(c2)))
        ri1, ri2 = sorted((int(r1), int(r2)))
        if (ci2 - ci1 + 1) * (ri2 - ri1 + 1) > 4096:
            continue
        for ci in range(ci1, ci2 + 1):
            for ri in range(ri1, ri2 + 1):
                refs.add(f'{get_column_letter(ci)}{ri}')
    return refs


def _detect_circular_references(ws) -> list[QualityIssue]:
    """Cellules formant une référence circulaire (même feuille)."""
    formulas: dict[str, set[str]] = {}
    for row in ws.iter_rows():
        for cell in row:
            if _dtype(cell) == 'f' and isinstance(cell.value, str):
                formulas[cell.coordinate] = _formula_refs(cell.value)
    if not formulas:
        return []
    color: dict[str, int] = {}
    in_cycle: set[str] = set()

    def visit(start: str) -> None:
        stack = [(start, iter(formulas.get(start, ())))]
        color[start] = 1
        while stack:
            node, it = stack[-1]
            advanced = False
            for ref in it:
                if ref not in formulas:
                    continue
                if color.get(ref) == 1:
                    idx = next(i for i, (n, _) in enumerate(stack) if n == ref)
                    in_cycle.update(n for n, _ in stack[idx:])
                elif ref not in color:
                    color[ref] = 1
                    stack.append((ref, iter(formulas.get(ref, ()))))
                    advanced = True
                    break
            if not advanced:
                color[node] = 2
                stack.pop()

    for coord in formulas:
        if coord not in color:
            visit(coord)
    return [QualityIssue(
        kind='circular_reference', sheet=ws.title, cell=coord,
        severity='HIGH', confidence='high',
        message="Référence circulaire (la cellule dépend d'elle-même, directement ou en chaîne)")
        for coord in sorted(in_cycle)]


def _detect_header_issues(ws) -> list[QualityIssue]:
    hdr, headers = _header_info(ws)
    if not headers:
        return []
    issues: list[QualityIssue] = []
    max_row = ws.max_row
    last_hdr_col = max(headers)
    seen: dict[str, str] = {}
    sample_rows = range(hdr + 1, min(max_row, hdr + 50) + 1)
    for cell in next(ws.iter_rows(min_row=hdr, max_row=hdr)):
        col = cell.column
        if col > last_hdr_col:
            continue
        letter = get_column_letter(col)
        has_data = any(ws.cell(r, col).value not in (None, '') for r in sample_rows)
        header_empty = cell.value is None or (isinstance(cell.value, str)
                                              and not cell.value.strip())
        if has_data and header_empty:
            issues.append(QualityIssue(
                kind='missing_header', sheet=ws.title, cell=f'{letter}{hdr}',
                severity='MEDIUM', confidence='medium',
                message=f"Colonne {letter} contient des données mais n'a pas d'en-tête"))
        if not (isinstance(cell.value, str) and cell.value.strip()):
            continue
        key = _norm_label(cell.value)
        if key in seen:
            issues.append(QualityIssue(
                kind='duplicate_header', sheet=ws.title, cell=f'{letter}{hdr}',
                severity='MEDIUM', confidence='medium',
                message=f'En-tête «{cell.value.strip()}» en double (déjà en colonne '
                        f'{seen[key]})',
                evidence=_redact(cell.value)))
        else:
            seen[key] = letter
    return issues


def _detect_hidden_data(ws) -> list[QualityIssue]:
    """Lignes ou colonnes MASQUÉES contenant des données (faussent totaux/revue)."""
    issues: list[QualityIssue] = []
    max_row, max_col = ws.max_row, ws.max_column
    for r, dim in ws.row_dimensions.items():
        if not getattr(dim, 'hidden', False):
            continue
        if not any(ws.cell(r, c).value not in (None, '') for c in range(1, max_col + 1)):
            continue
        issues.append(QualityIssue(
            kind='hidden_data', sheet=ws.title, cell=f'A{r}',
            severity='MEDIUM', confidence='medium',
            message=f'Ligne {r} masquée mais contenant des données (souvent oubliée dans '
                    'les totaux)'))
    for letter, dim in ws.column_dimensions.items():
        if not (getattr(dim, 'hidden', False) and letter.isalpha()):
            continue
        ci = column_index_from_string(letter)
        if not any(ws.cell(r, ci).value not in (None, '') for r in range(1, max_row + 1)):
            continue
        issues.append(QualityIssue(
            kind='hidden_data', sheet=ws.title, cell=f'{letter}1',
            severity='MEDIUM', confidence='medium',
            message=f'Colonne {letter} masquée mais contenant des données (souvent oubliée '
                    'dans les totaux)'))
    return issues


def _detect_outliers(ws) -> list[QualityIssue]:
    """Valeurs aberrantes (écart robuste à la médiane via MAD). Confiance basse."""
    issues: list[QualityIssue] = []
    headers = _header_row(ws)
    body = _data_rows(ws)
    if not headers or len(body) < 8:
        return issues
    for col, name in headers.items():
        pairs = [(r, ws.cell(r, col).value) for r in body
                 if _dtype(ws.cell(r, col)) == 'n' and ws.cell(r, col).value is not None]
        if len(pairs) < 8:
            continue
        nums = [v for _, v in pairs]
        med = statistics.median(nums)
        mad = statistics.median([abs(x - med) for x in nums])
        if mad == 0:
            continue
        candidates = [(r, x) for r, x in pairs
                      if abs(x - med) > _OUTLIER_ROBUST_Z * 1.4826 * mad]
        if not candidates or len(candidates) > max(3, 0.05 * len(pairs)):
            continue
        for r, x in candidates:
            issues.append(QualityIssue(
                kind='outlier', sheet=ws.title, cell=f'{get_column_letter(col)}{r}',
                severity='LOW', confidence='low',
                message=f'Valeur inhabituelle dans «{name}» : {x} — la plupart des valeurs '
                        f'tournent autour de {med}. À vérifier.',
                evidence=_redact(x)))
    return issues


def _detect_whitespace(ws) -> list[QualityIssue]:
    """Espaces superflus en début/fin ou caractères invisibles dans une colonne texte."""
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = _data_rows(ws)
    if not headers or len(body) < _MIN_TABLE_ROWS:
        return issues
    for col, name in headers.items():
        bad = 0
        for r in body:
            c = ws.cell(r, col)
            if _dtype(c) != 's' or not isinstance(c.value, str):
                continue
            if c.value != c.value.strip() or _INVISIBLE.search(c.value):
                bad += 1
        if not bad:
            continue
        issues.append(QualityIssue(
            kind='whitespace_issue', sheet=ws.title,
            cell=f'{get_column_letter(col)}{hdr}',
            severity='LOW', confidence='medium',
            message=f'Colonne «{name}» : {_n(bad, "cellule", "cellules")} avec des espaces '
                    'en trop ou des caractères invisibles — cassent tris, recherches et '
                    'regroupements'))
    return issues


def _detect_units_in_numbers(ws) -> list[QualityIssue]:
    """Nombres avec une unité/devise collée dans une colonne par ailleurs numérique
    (ex. « 9.3 kL » à côté de 10239) — ils ne sont pas pris dans les calculs."""
    issues: list[QualityIssue] = []
    hdr, headers = _header_info(ws)
    body = set(_data_rows(ws))
    by_col: dict[int, list[Cell]] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.row not in body or cell.value is None:
                continue
            by_col.setdefault(cell.column, []).append(cell)
    for col, cells in by_col.items():
        numeric = sum(1 for c in cells if _dtype(c) == 'n')
        unit_cells = [c for c in cells if _is_number_with_unit(c.value)]
        if not unit_cells or numeric < _MIN_NUMERIC_IN_COL:
            continue
        issues.append(QualityIssue(
            kind='number_with_unit', sheet=ws.title,
            cell=f'{get_column_letter(col)}{hdr or 1}',
            severity='HIGH', confidence='medium',
            message=f'{_n(len(unit_cells), "nombre", "nombres")} avec une unité/devise '
                    f'collée dans la {_col_label(headers, col)} '
                    f'(ex. « {unit_cells[0].value} ») — non calculables par Excel',
            evidence=_redact(unit_cells[0].value)))
    return issues


def _detect_empty_rows(ws) -> list[QualityIssue]:
    headers = _header_row(ws)
    max_row = ws.max_row
    if not headers:
        return []
    cols = list(headers)

    def row_empty(r: int) -> bool:
        return all(_kind_of(ws.cell(r, c)) is None for c in cols)

    data_rows = [r for r in range(_body_start(ws), max_row + 1) if not row_empty(r)]
    if len(data_rows) < _MIN_TABLE_ROWS:
        return []
    first, last = data_rows[0], data_rows[-1]
    return [QualityIssue(
        kind='empty_row', sheet=ws.title, cell=f'A{r}',
        severity='LOW', confidence='medium',
        message=f'Ligne {r} entièrement vide au milieu du tableau (entre {first} et {last})')
        for r in range(first + 1, last) if row_empty(r)]


def _detect_date_order(ws) -> list[QualityIssue]:
    issues: list[QualityIssue] = []
    headers = _header_row(ws)
    body = _data_rows(ws)
    if not headers or len(body) < 8:
        return issues
    for col, name in headers.items():
        dated = [(r, ws.cell(r, col).value) for r in body
                 if getattr(ws.cell(r, col), 'is_date', False)]
        if len(dated) < max(8, int(0.8 * len(body))):
            continue
        vals = [d for _, d in dated]
        pairs = list(zip(vals, vals[1:]))
        asc = sum(1 for a, b in pairs if a <= b)
        desc = sum(1 for a, b in pairs if a >= b)
        if not pairs or max(asc, desc) / len(pairs) < 0.9:
            continue
        ascending = asc >= desc
        breaks = [r2 for (r1, a), (r2, b) in zip(dated, dated[1:])
                  if (b < a if ascending else b > a)]
        if not breaks:
            continue
        issues.append(QualityIssue(
            kind='date_out_of_order', sheet=ws.title,
            cell=f'{get_column_letter(col)}{breaks[0]}',
            severity='LOW', confidence='low',
            message=f'Colonne «{name}» : {_n(len(breaks), "date", "dates")} dans le désordre '
                    "(la colonne semble triée mais ne l'est pas partout)"))
    return issues


def _detect_formula_refs_empty(ws) -> list[QualityIssue]:
    """Formule ARITHMÉTIQUE simple référençant une cellule vide (résultat faux silencieux)."""
    issues: list[QualityIssue] = []
    for row in ws.iter_rows():
        for cell in row:
            if _dtype(cell) != 'f' or not isinstance(cell.value, str):
                continue
            cleaned = _QUALIFIED_REF.sub(' ', cell.value)
            if '(' in cleaned:
                continue
            empty_refs: list[str] = []
            for m in _SINGLE_REF.finditer(cleaned):
                coord = f'{m.group(1)}{int(m.group(2))}'
                if coord == cell.coordinate:
                    continue
                if _kind_of(ws[coord]) is not None:
                    continue
                empty_refs.append(coord)
            if not empty_refs:
                continue
            issues.append(QualityIssue(
                kind='formula_references_empty', sheet=ws.title, cell=cell.coordinate,
                severity='LOW', confidence='low',
                message=f'Formule qui utilise une cellule vide ({", ".join(empty_refs[:3])}) '
                        '— le résultat est probablement faux',
                evidence=_redact(cell.value)))
    return issues


def _detect_stray_data(ws) -> list[QualityIssue]:
    """Données présentes dans une colonne AU-DELÀ des en-têtes du tableau."""
    hdr, headers = _header_info(ws)
    if not headers:
        return []
    last_header_col = max(headers)
    stray: set[int] = set()
    for row in ws.iter_rows(min_row=hdr, min_col=last_header_col + 1):
        for cell in row:
            if cell.value not in (None, ''):
                stray.add(cell.column)
    return [QualityIssue(
        kind='stray_data', sheet=ws.title, cell=f'{get_column_letter(c)}1',
        severity='LOW', confidence='low',
        message=f'Données isolées dans la colonne {get_column_letter(c)}, en dehors du '
                'tableau (sans en-tête)')
        for c in sorted(stray)]


_META_SHEET_NAME = re.compile(r'(?:^|[ _\-])(?:cl[ée]|v[ée]rit[ée]|documentation|l[ée]gende|'
                              r'dictionnaire|m[ée]ta|metadata|readme|notice|annexe)', re.I)
_META_SHEET_TITLE = re.compile(r'v[ée]rit[ée]\s+terrain|documentation|grille de v|'
                               r'l[ée]gende|dictionnaire', re.I)


def _is_metadata_sheet(ws) -> bool:
    if _META_SHEET_NAME.search(ws.title or ''):
        return True
    top = ws.cell(1, 1).value
    return isinstance(top, str) and bool(_META_SHEET_TITLE.search(top))


def _detect_table_structure(ws) -> list[QualityIssue]:
    hdr = _header_info(ws)[0]
    if hdr <= 1:
        return []
    return [QualityIssue(
        kind='table_offset', sheet=ws.title, cell=f'A{hdr}',
        severity='LOW', confidence='high',
        message=f'En-tête du tableau en ligne {hdr} (titre fusionné et/ou lignes vides '
                'au-dessus, ignorés)')]


_DETECTORS = (
    _detect_error_cells,
    _detect_number_as_text,
    _detect_formula_anomalies,
    _detect_range_omission,
    _detect_missing_values,
    _detect_mixed_types,
    _detect_duplicate_rows,
    _detect_duplicate_keys,
    _detect_total_mismatch,
    _detect_merged_cells,
    _detect_date_issues,
    _detect_inconsistent_categories,
    _detect_circular_references,
    _detect_header_issues,
    _detect_hidden_data,
    _detect_outliers,
    _detect_whitespace,
    _detect_units_in_numbers,
    _detect_empty_rows,
    _detect_date_order,
    _detect_formula_refs_empty,
    _detect_stray_data,
    _detect_table_structure,
)


def _coerce_workbook(path_or_wb):
    if isinstance(path_or_wb, Workbook):
        return (path_or_wb, False)
    return (load_workbook(path_or_wb, read_only=False, data_only=False, keep_vba=False), True)


def _check_deadline(deadline=None) -> None:
    if deadline is None:
        return
    import time
    if time.monotonic() > deadline:
        raise TimeoutError('budget temps analyse qualité dépassé')


def analyze_quality(path_or_wb, *, deadline=None) -> list[QualityIssue]:
    """Lance tous les détecteurs qualité sur le classeur. Lecture seule."""
    wb, should_close = _coerce_workbook(path_or_wb)
    try:
        issues: list[QualityIssue] = []
        for ws in wb.worksheets:
            _check_deadline(deadline)
            if _is_metadata_sheet(ws):
                continue
            for detector in _DETECTORS:
                _check_deadline(deadline)
                issues.extend(detector(ws))
        return issues
    finally:
        if should_close:
            wb.close()


def data_sheet_names(path_or_wb) -> list[str]:
    """Titres des feuilles de DONNÉES (hors feuilles de documentation/clé)."""
    wb, should_close = _coerce_workbook(path_or_wb)
    try:
        return [ws.title for ws in wb.worksheets if not _is_metadata_sheet(ws)]
    finally:
        if should_close:
            wb.close()


def detect_overwritten_formulas(path: str | Path) -> list[QualityIssue]:
    return [i for i in analyze_quality(path) if i.kind == 'overwritten_formula']
