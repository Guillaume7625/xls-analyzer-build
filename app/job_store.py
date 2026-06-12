# ====================================================================================
# PROVENANCE : RECONSTRUCTION FIDÈLE. Squelette pycdc correct ; __init__ et update
# (corps perdus) reconstitués depuis leur docstring d'origine et le contrat d'appel
# (pipeline : store.update(job, status=..., report_json_path=..., error=...) ;
# env XLS_JOBS_DIR ; AttributeError sur champ inconnu).
# ====================================================================================
"""Stockage des jobs par fichiers, dans <root>/<job_id>/.

Le job_id est tiré de `secrets` (jamais dérivé du nom de fichier). Toutes les
écritures de job.json sont atomiques. Aucun pilotage de conteneur ici (invariant porteur).
"""
from __future__ import annotations

import json
import os
import re
import secrets
from datetime import datetime, timezone
from pathlib import Path

from app.atomic_io import write_json_atomic
from app.models import Job, JobStatus
from app.sensitive import mask_sensitive_filename

_JOB_ID_RE = re.compile('^[0-9a-f]{32}$')
_DEFAULT_JOBS_DIR = '/run/xls-analyzer/jobs'
INPUT_NAME = 'input.xlsx'
JOB_JSON = 'job.json'
STATIC_REPORT = 'static_report.json'
RECALC_REQUEST = 'recalc_request.json'
RECALC_RESULT = 'recalc_result.json'
REPORT_JSON = 'report.json'
REPORT_MD = 'report.md'
RECALC_LOCK = 'recalc.lock'


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def validate_job_id(job_id: str) -> str:
    """Rejette tout id non conforme (fail-closed contre le path traversal)."""
    if not _JOB_ID_RE.match(job_id):
        raise ValueError(f'job_id invalide: {job_id!r}')
    return job_id


class JobStore:
    """Accès au répertoire des jobs. `root` configurable (env XLS_JOBS_DIR)."""

    def __init__(self, root: str | Path | None = None):
        self.root = Path(root or os.environ.get('XLS_JOBS_DIR', _DEFAULT_JOBS_DIR))
        self.root.mkdir(parents=True, exist_ok=True)

    def job_dir(self, job_id: str) -> Path:
        return self.root / validate_job_id(job_id)

    def input_path(self, job: Job) -> Path:
        return self.job_dir(job.job_id) / INPUT_NAME

    def path(self, job: Job, name: str) -> Path:
        return self.job_dir(job.job_id) / name

    def create_job(self, filename_original: str) -> Job:
        """Crée un répertoire de job neuf et son job.json (statut CREATED)."""
        job_id = secrets.token_hex(16)
        d = self.root / job_id
        d.mkdir(parents=True, exist_ok=False)
        now = _now()
        job = Job(job_id=job_id, status=JobStatus.CREATED, created_at=now,
                  updated_at=now,
                  filename_original=mask_sensitive_filename(filename_original),
                  input_path=str(d / INPUT_NAME))
        self._write_job(job)
        return job

    def load_job(self, job_id: str) -> Job:
        data = (self.job_dir(job_id) / JOB_JSON).read_text(encoding='utf-8')
        return Job.model_validate(json.loads(data))

    def update(self, job: Job, status: JobStatus | None = None, **fields) -> Job:
        """Met à jour le job en mémoire et persiste job.json atomiquement.

        `fields` n'accepte que des attributs connus de Job (sinon AttributeError).
        """
        for key in fields:
            if key not in Job.model_fields:
                raise AttributeError(f'champ Job inconnu: {key!r}')
        updates: dict = dict(fields)
        if status is not None:
            updates['status'] = status
        updates['updated_at'] = _now()
        new_job = job.model_copy(update=updates)
        self._write_job(new_job)
        return new_job

    def _write_job(self, job: Job) -> None:
        write_json_atomic(self.job_dir(job.job_id) / JOB_JSON,
                          job.model_dump(mode='json'))
