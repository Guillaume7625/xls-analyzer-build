# ====================================================================================
# PROVENANCE : RECONSTRUCTION CORRIGÉE. validate_extension restitué par pycdc avec
# artefact `None.ok()` (→ PreflightResult.ok()). validate_signature et validate_file :
# corps perdus, ré-implémentés depuis la spec (4 octets magiques ZIP, codes
# SIGNATURE_MISMATCH / EXT_NOT_ALLOWED, ordre extension → signature).
# ====================================================================================
"""Validation amont du fichier : extension autorisée puis signature magique.

Première barrière fail-closed, AVANT toute lecture du contenu par zipfile/openpyxl :
- l'extension déclarée doit appartenir à ALLOWED_EXTENSIONS (.xlsx uniquement en V1) ;
- les premiers octets doivent être une signature ZIP (un .xlsx est un conteneur ZIP).
Un fichier renommé (exe→xlsx, xls binaire→xlsx) est rejeté ici, à coût quasi nul.
"""
from __future__ import annotations

from pathlib import Path

from app.config import ALLOWED_EXTENSIONS
from app.reject import PreflightResult, RejectCode, Severity

# Signatures ZIP valides : PK\x03\x04 (local file header), PK\x05\x06 (archive vide),
# PK\x07\x08 (spanned). Tout .xlsx légitime commence par l'une d'elles.
_ZIP_MAGICS = (b'PK\x03\x04', b'PK\x05\x06', b'PK\x07\x08')


def validate_extension(filename_original: str) -> PreflightResult:
    """Rejette toute extension hors liste blanche (comparaison insensible à la casse)."""
    ext = Path(filename_original).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return PreflightResult.reject(
            RejectCode.EXT_NOT_ALLOWED, Severity.HIGH,
            f"extension '{ext or '(aucune)'}' non autorisée (attendu : "
            f"{', '.join(ALLOWED_EXTENSIONS)})")
    return PreflightResult.ok()


def validate_signature(path: str | Path) -> PreflightResult:
    """Rejette tout fichier dont les 4 premiers octets ne sont pas une signature ZIP."""
    try:
        with open(path, 'rb') as fh:
            head = fh.read(4)
    except OSError as exc:
        return PreflightResult.reject(RejectCode.SIGNATURE_MISMATCH, Severity.HIGH,
                                      f'fichier illisible ({exc.__class__.__name__})')
    if head not in _ZIP_MAGICS:
        return PreflightResult.reject(
            RejectCode.SIGNATURE_MISMATCH, Severity.HIGH,
            'signature de fichier invalide : pas un conteneur ZIP/.xlsx '
            '(fichier renommé ?)')
    return PreflightResult.ok()


def validate_file(filename_original: str, path: str | Path) -> PreflightResult:
    """Extension d'abord (gratuit), signature ensuite (4 octets lus)."""
    res = validate_extension(filename_original)
    if not res.passed:
        return res
    return validate_signature(path)
