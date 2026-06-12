# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE (pycdas atomic_io.disasm.txt, vérifiée
# opcode/opcode, table d'exceptions incluse). Mode fichier 0o600 (const 384).
# ====================================================================================
"""Écritures atomiques fail-closed : tmp -> fsync -> os.replace.

Brique partagée par report_builder, job_store, recalc_client (manifeste) et le
runner. Garantit qu'un lecteur ne voit jamais un fichier partiellement écrit :
on écrit dans un fichier temporaire, on fsync, puis on renomme atomiquement.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


def write_bytes_atomic(path: str | Path, data: bytes) -> None:
    """Écrit `data` à `path` de façon atomique et durable.

    En cas d'erreur, le fichier temporaire est supprimé et l'exception remonte
    (fail-closed : pas d'artefact à moitié écrit laissé derrière).
    """
    path = Path(path)
    tmp = path.with_name(f'.{path.name}.tmp.{os.getpid()}')
    try:
        fd = os.open(tmp, os.O_CREAT | os.O_WRONLY | os.O_TRUNC, 0o600)
        try:
            mv = memoryview(data)
            while mv:
                written = os.write(fd, mv)
                mv = mv[written:]
            os.fsync(fd)
        finally:
            os.close(fd)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise
    # fsync du répertoire : durabilité du rename. Best-effort (échoue
    # légitimement sous Windows où open(répertoire) lève OSError).
    try:
        dir_fd = os.open(path.parent, os.O_RDONLY)
        try:
            os.fsync(dir_fd)
        finally:
            os.close(dir_fd)
    except OSError:
        pass


def write_text_atomic(path: str | Path, text: str) -> None:
    write_bytes_atomic(path, text.encode('utf-8'))


def write_json_atomic(path: str | Path, obj: Any) -> None:
    """Sérialise `obj` en JSON (UTF-8, indenté, clés triées) et écrit atomiquement."""
    data = json.dumps(obj, ensure_ascii=False, indent=2, sort_keys=True).encode('utf-8')
    write_bytes_atomic(path, data)
