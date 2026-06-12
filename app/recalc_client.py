# ====================================================================================
# PROVENANCE : DÉCOMPILATION PROPRE (pycdc) — corps intégral restitué sans artefact,
# vérifié. Seule la bannière d'outillage a été ôtée.
# ====================================================================================
"""Demande de recalc côté app : dépose un MANIFESTE, rien d'autre.

Invariant porteur : l'app n'orchestre aucun conteneur et ne lance aucun recalc.
Elle écrit seulement recalc_request.json et passe le job en PENDING_RECALC.
Le runner (process séparé, hors app/) lira ce manifeste et pilotera le recalcul
isolé. Aucun import de moteur de conteneur, aucun port, aucun sous-processus ici.
"""
from __future__ import annotations
from datetime import datetime, timezone
from app.atomic_io import write_json_atomic
from app.job_store import RECALC_REQUEST, RECALC_RESULT, JobStore
from app.models import Job, JobStatus
RECALC_REQUEST_VERSION = 1

def request_recalc(store = None, job = None):
    '''Écrit recalc_request.json et passe le job en PENDING_RECALC.

    Autorisé uniquement depuis STATIC_DONE (transition SPEC). Toute autre source
    lève ValueError (fail-closed : pas de recalc hors séquence).
    '''
    if job.status != JobStatus.STATIC_DONE:
        raise ValueError(f'''recalc non autorisé depuis le statut {job.status.value}''')
    manifest = {
        'version': RECALC_REQUEST_VERSION,
        'job_id': job.job_id,
        'input_path': job.input_path,
        'result_path': str(store.path(job, RECALC_RESULT)),
        'requested_at': datetime.now(timezone.utc).isoformat() }
    request_path = store.path(job, RECALC_REQUEST)
    write_json_atomic(request_path, manifest)
    return store.update(job, status = JobStatus.PENDING_RECALC, recalc_request_path = str(request_path))

