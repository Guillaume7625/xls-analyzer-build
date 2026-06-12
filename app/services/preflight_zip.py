# ====================================================================================
# PROVENANCE : RECONSTRUCTION CORRIGÉE + SPEC-BASED.
# - _is_traversal : restitué propre par pycdc (repris tel quel).
# - _ooxml_blocker : restitué avec artefacts (comparaisons mutilées) → corrigé selon
#   les constantes du bytecode (préfixes xl/externalLinks, xl/embeddings, activex,
#   oleObject*.bin / media *.bin, xl/connections.xml, xl/vbaProject*).
# - preflight_zip : corps perdu, ré-implémenté depuis la spec (codes RejectCode,
#   seuils config.py, doctrine « jamais décompresser pour décider », ordre des
#   contrôles documenté dans DECOMPILATION.md). Le contrôle de profondeur XML
#   (MAX_XML_DEPTH) est borné par flux et limité aux parts OOXML obligatoires.
# ====================================================================================
"""Préflight ZIP/OOXML fail-closed : tout contrôle se fait sur les MÉTADONNÉES du
répertoire central (infolist). On ne décompresse jamais une entrée pour décider d'un
rejet de taille/ratio — c'est précisément le vecteur zip-bomb.

Ordre des contrôles (du moins cher au plus cher, premier échec = rejet) :
  1. is_zipfile                  -> NOT_A_ZIP
  2. nombre d'entrées            -> TOO_MANY_ENTRIES
  3. doublons d'entrée           -> DUPLICATE_ENTRY
  4. traversée de chemin         -> ZIP_PATH_TRAVERSAL
  5. contenus actifs interdits   -> MACRO_VBAPROJECT / EXTERNAL_LINKS / CONNECTIONS /
                                    ACTIVEX / OLE_OBJECT / EMBEDDING (parts collectées)
  6. taille décompressée/entrée  -> ENTRY_TOO_LARGE
  7. taille décompressée totale  -> TOTAL_TOO_LARGE
  8. ratio de compression        -> COMPRESSION_RATIO_EXCEEDED
  9. sharedStrings.xml           -> SHAREDSTRINGS_TOO_LARGE
 10. parts OOXML obligatoires    -> MISSING_OOXML_PART
 11. profondeur XML (bornée)     -> XML_DEPTH_EXCEEDED
"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from app.config import (MAX_COMPRESSION_RATIO, MAX_ENTRY_UNCOMPRESSED,
                        MAX_SHAREDSTRINGS_BYTES, MAX_TOTAL_UNCOMPRESSED, MAX_XML_DEPTH,
                        MAX_ZIP_ENTRIES, MIN_ENTRY_BYTES_FOR_RATIO, REQUIRED_OOXML_PARTS)
from app.reject import PreflightResult, RejectCode, Severity


def _is_traversal(name: str) -> bool:
    """Vrai si le nom d'entrée sort de l'arborescence (Zip-Slip)."""
    if name.startswith('/') or name.startswith('\\'):
        return True
    if len(name) > 1 and name[1] == ':':
        return True
    parts = name.replace('\\', '/').split('/')
    return '..' in parts


def _ooxml_blocker(name: str) -> tuple[RejectCode, str] | None:
    """(code, détail) si l'entrée désigne un contenu actif interdit, sinon None."""
    low = name.replace('\\', '/').lower()
    if low.startswith('xl/vbaproject') or low.endswith('vbaproject.bin'):
        return (RejectCode.MACRO_VBAPROJECT, 'macro VBA (vbaProject)')
    if low.startswith('xl/externallinks/'):
        return (RejectCode.EXTERNAL_LINKS, 'lien externe (externalLinks)')
    if low == 'xl/connections.xml':
        return (RejectCode.CONNECTIONS, 'connexion de données (connections.xml)')
    if 'activex' in low:
        return (RejectCode.ACTIVEX, 'contrôle ActiveX')
    if low.startswith('xl/embeddings/'):
        if 'oleobject' in low or low.endswith('.bin'):
            return (RejectCode.OLE_OBJECT, 'objet OLE incorporé')
        return (RejectCode.EMBEDDING, 'objet incorporé (embeddings)')
    if low.startswith('xl/media/') and low.endswith('.bin'):
        return (RejectCode.OLE_OBJECT, 'binaire incorporé (media/*.bin)')
    return None


def _xml_depth_exceeded(zf: zipfile.ZipFile, name: str) -> bool:
    """Profondeur XML > MAX_XML_DEPTH ? Lecture en flux, bornée à
    MAX_SHAREDSTRINGS_BYTES — jamais l'arbre complet en mémoire (anti billion-laughs
    structurel : pas d'expansion d'entités avec ElementTree, profondeur plafonnée)."""
    budget = MAX_SHAREDSTRINGS_BYTES
    depth = 0
    try:
        with zf.open(name) as raw:
            stream = io.BufferedReader(raw)  # type: ignore[arg-type]
            for event, elem in ElementTree.iterparse(stream, events=('start', 'end')):
                if event == 'start':
                    depth += 1
                    if depth > MAX_XML_DEPTH:
                        return True
                else:
                    depth -= 1
                    elem.clear()
                budget -= 1
                if budget <= 0:
                    return False
    except (ElementTree.ParseError, OSError, zipfile.BadZipFile, RuntimeError):
        # XML illisible : ce n'est pas le rôle de ce contrôle de trancher —
        # openpyxl échouera proprement plus loin si la part est corrompue.
        return False
    return False


def preflight_zip(path: str | Path) -> PreflightResult:
    """Inspection complète du conteneur. Premier échec = rejet (fail-closed)."""
    path = Path(path)
    if not zipfile.is_zipfile(path):
        return PreflightResult.reject(RejectCode.NOT_A_ZIP, Severity.HIGH,
                                      'conteneur ZIP invalide ou corrompu')
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        return PreflightResult.reject(RejectCode.NOT_A_ZIP, Severity.HIGH,
                                      f'ouverture ZIP impossible ({exc.__class__.__name__})')
    with zf:
        infos = zf.infolist()
        if len(infos) > MAX_ZIP_ENTRIES:
            return PreflightResult.reject(
                RejectCode.TOO_MANY_ENTRIES, Severity.HIGH,
                f'{len(infos)} entrées > {MAX_ZIP_ENTRIES} autorisées')

        seen: set[str] = set()
        for info in infos:
            if info.filename in seen:
                return PreflightResult.reject(
                    RejectCode.DUPLICATE_ENTRY, Severity.HIGH,
                    "entrée en double dans l'archive (masquage de contenu possible)",
                    parts=[info.filename])
            seen.add(info.filename)

        for info in infos:
            if _is_traversal(info.filename):
                return PreflightResult.reject(
                    RejectCode.ZIP_PATH_TRAVERSAL, Severity.CRITICAL,
                    "nom d'entrée sortant de l'arborescence (Zip-Slip)",
                    parts=[info.filename])

        blocking: dict[RejectCode, list[str]] = {}
        details: dict[RejectCode, str] = {}
        for info in infos:
            hit = _ooxml_blocker(info.filename)
            if hit:
                code, detail = hit
                blocking.setdefault(code, []).append(info.filename)
                details.setdefault(code, detail)
        if blocking:
            # Priorité d'affichage : macro > OLE > ActiveX > liens > connexions > embed.
            order = (RejectCode.MACRO_VBAPROJECT, RejectCode.OLE_OBJECT,
                     RejectCode.ACTIVEX, RejectCode.EXTERNAL_LINKS,
                     RejectCode.CONNECTIONS, RejectCode.EMBEDDING)
            code = next(c for c in order if c in blocking)
            all_parts = sorted(p for parts in blocking.values() for p in parts)
            return PreflightResult.reject(
                code, Severity.CRITICAL,
                f'contenu actif interdit : {details[code]}', parts=all_parts)

        total = 0
        for info in infos:
            if info.file_size > MAX_ENTRY_UNCOMPRESSED:
                return PreflightResult.reject(
                    RejectCode.ENTRY_TOO_LARGE, Severity.HIGH,
                    f'{info.filename} : {info.file_size} octets décompressés > '
                    f'{MAX_ENTRY_UNCOMPRESSED}', parts=[info.filename])
            total += info.file_size
        if total > MAX_TOTAL_UNCOMPRESSED:
            return PreflightResult.reject(
                RejectCode.TOTAL_TOO_LARGE, Severity.HIGH,
                f'taille décompressée totale {total} > {MAX_TOTAL_UNCOMPRESSED}')

        for info in infos:
            if info.compress_size < MIN_ENTRY_BYTES_FOR_RATIO:
                continue
            if info.file_size / max(1, info.compress_size) > MAX_COMPRESSION_RATIO:
                return PreflightResult.reject(
                    RejectCode.COMPRESSION_RATIO_EXCEEDED, Severity.CRITICAL,
                    f'{info.filename} : ratio de compression '
                    f'{info.file_size // max(1, info.compress_size)}:1 > '
                    f'{MAX_COMPRESSION_RATIO}:1 (zip-bomb probable)',
                    parts=[info.filename])

        for info in infos:
            if info.filename == 'xl/sharedStrings.xml' \
                    and info.file_size > MAX_SHAREDSTRINGS_BYTES:
                return PreflightResult.reject(
                    RejectCode.SHAREDSTRINGS_TOO_LARGE, Severity.HIGH,
                    f'sharedStrings.xml : {info.file_size} octets > '
                    f'{MAX_SHAREDSTRINGS_BYTES}', parts=[info.filename])

        names = set(seen)
        missing = [p for p in REQUIRED_OOXML_PARTS if p not in names]
        if missing:
            return PreflightResult.reject(
                RejectCode.MISSING_OOXML_PART, Severity.HIGH,
                f'parts OOXML obligatoires absentes : {", ".join(missing)}',
                parts=missing)

        depth_targets = [p for p in ('xl/workbook.xml', 'xl/sharedStrings.xml')
                         if p in names]
        for part in depth_targets:
            if _xml_depth_exceeded(zf, part):
                return PreflightResult.reject(
                    RejectCode.XML_DEPTH_EXCEEDED, Severity.CRITICAL,
                    f'{part} : profondeur XML > {MAX_XML_DEPTH} '
                    '(structure pathologique)', parts=[part])

    return PreflightResult.ok()
