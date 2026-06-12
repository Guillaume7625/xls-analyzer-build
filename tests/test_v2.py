"""Tests xls-analyzer v2 — reconstruction + durcissements.

Couvre :
  T1  imports & enums (F-01/F-02/F-03 corrigés)
  T2  préflight fail-closed (macro, zip-bomb, traversal, doublon, parts manquantes)
  T3  validation amont (extension, signature)
  T4  marquages classification — y compris les contournements F-20 (tirets/underscores)
  T5  PII : NIR, IBAN, carte (Luhn, F-21), e-mail intra-cellule (F-22)
  T6  redaction : aucun NIR/IBAN en clair dans le rapport
  T7  validateurs : NIR/IBAN/date/cohérence/bornes
  T8  qualité : total_mismatch, formule écrasée, doublons, cellules masquées
  T9  pipeline bout-en-bout (STATIC_DONE, verdicts, artefacts atomiques)
  T10 API : upload, anti-injection CSV, 404/400, Content-Length 413
"""
from __future__ import annotations

import io
import json
import os
import zipfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from app.job_store import JobStore
from app.models import JobStatus, Verdict
from app.pipeline import run_pipeline
from app.reject import PreflightResult, RejectCode, Severity
from app.report_builder import quality_issues_to_findings
from app.sensitive import (_luhn_ok, _marking_hits, analyze_sensitive,
                           mask_sensitive_filename, mask_sensitive_text)
from app.services.file_validation import validate_extension, validate_file
from app.services.preflight_zip import preflight_zip
from app.quality_analyzer import analyze_quality
from app.validators import (_date_problem, _iban_valid, _nir_valid,
                            analyze_validity)


# ---------------------------------------------------------------- fixtures utilitaires
def _save(wb: Workbook, tmp_path: Path, name: str = 'f.xlsx') -> Path:
    p = tmp_path / name
    wb.save(p)
    return p


def _wb_simple() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Données'
    ws.append(['Nom', 'Montant', 'Date'])
    for i in range(1, 12):
        ws.append([f'Ligne{i}', 100 + i, f'2024-01-{i:02d}'])
    return wb


# ------------------------------------------------------------------------- T1 : noyau
def test_enums_et_modeles():
    assert RejectCode.MACRO_VBAPROJECT.value == 'MACRO_VBAPROJECT'
    assert Severity.CRITICAL.value == 'CRITICAL'
    r = PreflightResult.reject(RejectCode.NOT_A_ZIP, detail='x')
    assert not r.passed and r.severity is Severity.HIGH and r.blocking_parts == []
    assert PreflightResult.ok().passed


# --------------------------------------------------------------------- T2 : préflight
def test_preflight_fichier_sain(tmp_path):
    p = _save(_wb_simple(), tmp_path)
    assert preflight_zip(p).passed


def test_preflight_macro_vba(tmp_path):
    p = _save(_wb_simple(), tmp_path)
    with zipfile.ZipFile(p, 'a') as zf:
        zf.writestr('xl/vbaProject.bin', b'\x00' * 64)
    res = preflight_zip(p)
    assert not res.passed and res.reason_code is RejectCode.MACRO_VBAPROJECT
    assert 'xl/vbaProject.bin' in res.blocking_parts


def test_preflight_zip_bomb_ratio(tmp_path):
    p = tmp_path / 'bomb.xlsx'
    with zipfile.ZipFile(p, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', '<Types/>')
        zf.writestr('_rels/.rels', '<Relationships/>')
        zf.writestr('xl/workbook.xml', '<workbook/>')
        # 50 Mo de zéros : compressé ≈ 50 Ko (≥ MIN_ENTRY_BYTES_FOR_RATIO=4096),
        # ratio ≈ 1000:1 > 100 — la règle du binaire ne s'applique qu'au-delà de 4096 o.
        zf.writestr('xl/big.xml', b'\x00' * (50 * 1024 * 1024))
    res = preflight_zip(p)
    assert not res.passed
    assert res.reason_code is RejectCode.COMPRESSION_RATIO_EXCEEDED


def test_preflight_traversal_et_doublon(tmp_path):
    p = tmp_path / 't.xlsx'
    with zipfile.ZipFile(p, 'w') as zf:
        zf.writestr('../evil.txt', 'x')
    assert preflight_zip(p).reason_code is RejectCode.ZIP_PATH_TRAVERSAL

    p2 = tmp_path / 'd.xlsx'
    with zipfile.ZipFile(p2, 'w') as zf:
        zf.writestr('a.xml', '1')
        zf.writestr('a.xml', '2')
    assert preflight_zip(p2).reason_code is RejectCode.DUPLICATE_ENTRY


def test_preflight_parts_manquantes(tmp_path):
    p = tmp_path / 'm.xlsx'
    with zipfile.ZipFile(p, 'w') as zf:
        zf.writestr('hello.txt', 'x')
    assert preflight_zip(p).reason_code is RejectCode.MISSING_OOXML_PART


def test_preflight_pas_un_zip(tmp_path):
    p = tmp_path / 'n.xlsx'
    p.write_bytes(b'MZ\x90\x00' + b'\x00' * 64)
    assert preflight_zip(p).reason_code is RejectCode.NOT_A_ZIP


# ------------------------------------------------------------------- T3 : amont
def test_validate_extension_et_signature(tmp_path):
    assert validate_extension('a.xlsx').passed
    assert validate_extension('a.XLSX').passed
    assert validate_extension('a.xlsm').reason_code is RejectCode.EXT_NOT_ALLOWED
    assert validate_extension('a').reason_code is RejectCode.EXT_NOT_ALLOWED
    exe = tmp_path / 'fake.xlsx'
    exe.write_bytes(b'MZ\x90\x00rest')
    assert validate_file('fake.xlsx', exe).reason_code is RejectCode.SIGNATURE_MISMATCH
    good = _save(_wb_simple(), tmp_path, 'g.xlsx')
    assert validate_file('g.xlsx', good).passed


# --------------------------------------------------------- T4 : marquages (dont F-20)
@pytest.mark.parametrize('texte', [
    'SECRET DÉFENSE', 'Secret Defense', 'TRÈS SECRET', 'DIFFUSION RESTREINTE',
    'CONFIDENTIEL DEFENSE', 'NATO SECRET', 'SPECIAL FRANCE',
    # Contournements que le binaire d'origine ratait (F-20) :
    'SECRET-DÉFENSE', 'SECRET_DEFENSE', 'Très.Secret', 'DIFFUSION-RESTREINTE',
])
def test_marquages_detectes(texte):
    assert _marking_hits(texte), f'marquage non détecté : {texte!r}'


@pytest.mark.parametrize('texte', [
    'secrétaire de mairie', 'le secret professionnel est restreint au médecin',
    'restreindre les coûts', 'top des ventes',
])
def test_marquages_faux_positifs_evites(texte):
    hits = _marking_hits(texte)
    assert not hits, f'faux positif marquage : {texte!r} -> {hits}'


def test_marquage_bloque_pipeline(tmp_path):
    wb = _wb_simple()
    wb.active['A5'] = 'Synthèse SECRET-DÉFENSE — ne pas diffuser'
    p = _save(wb, tmp_path)
    issues = analyze_sensitive(p)
    kinds = {i.kind for i in issues}
    assert 'classification_marking' in kinds


# --------------------------------------------------------------- T5 : PII (F-21/F-22)
def test_luhn():
    assert _luhn_ok('4539578763621486')      # VISA de test, Luhn valide
    assert not _luhn_ok('4539578763621487')  # même suite, clé fausse (F-21)


def test_nir_detecte_et_redige(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'RH'
    ws.append(['Agent', 'NIR'])
    nirs = ['185057800608436', '255082B07300417', '178122A12345678',
            '299016311111111', '187049912345678', '156077512345678']
    for i, n in enumerate(nirs):
        ws.append([f'A{i}', n])
    p = _save(wb, tmp_path)
    issues = analyze_sensitive(p)
    nir_issues = [i for i in issues if i.kind == 'pii_nir']
    assert nir_issues, 'NIR non détectés'
    # Redaction : aucune valeur de NIR ne doit transiter dans message/evidence.
    for i in nir_issues:
        for n in nirs:
            assert n not in i.message and n not in i.evidence


def test_email_intra_cellule_F22(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Contact'])
    for i in range(6):
        ws.append([f'joindre : agent{i}@marine.defense.gouv.fr (bureau N3)'])
    p = _save(wb, tmp_path)
    # En-tête « Contact » ne matche pas _HDR_EMAIL ; la forme ne fullmatch pas non plus.
    # F-22 : la détection doit fonctionner via recherche intra-cellule quand l'en-tête
    # le suggère — ici on force l'en-tête e-mail pour valider le chemin de recherche.
    ws['A1'] = 'E-mail'
    p = _save(wb, tmp_path, 'f2.xlsx')
    issues = analyze_sensitive(p)
    assert any(i.kind == 'pii_email' for i in issues)


def test_carte_bancaire_luhn_pipeline(tmp_path):
    wb = _wb_simple()
    wb.active['E2'] = 'CB 4539 5787 6362 1486'
    p = _save(wb, tmp_path)
    issues = analyze_sensitive(p)
    assert any(i.kind == 'pii_card' for i in issues)
    wb2 = _wb_simple()
    wb2.active['E2'] = 'réf dossier 4539578763621487'  # Luhn KO -> pas une carte
    p2 = _save(wb2, tmp_path, 'f2.xlsx')
    assert not any(i.kind == 'pii_card' for i in analyze_sensitive(p2))


def test_mask_sensitive_text_et_filename():
    assert '185057800608436' not in mask_sensitive_text('NIR: 185057800608436')
    assert 'FR7630006000011234567890189' not in mask_sensitive_text(
        'IBAN FR76 3000 6000 0112 3456 7890 189'.replace(' ', ''))
    masked = mask_sensitive_text('password: hunter2secret')
    assert 'hunter2secret' not in masked
    fn = mask_sensitive_filename('export_185057800608436.xlsx')
    assert '185057800608436' not in fn and fn.endswith('.xlsx')


# ----------------------------------------------------------------- T6 : redaction E2E
def test_redaction_dans_findings():
    from app.quality_analyzer import QualityIssue
    q = QualityIssue(kind='pii_nir', sheet='RH', cell='B2', severity='CRITICAL',
                     confidence='high', message='NIR 185057800608436 trouvé',
                     evidence='185057800608436')
    f = quality_issues_to_findings([q])[0]
    assert '185057800608436' not in f.title
    assert '185057800608436' not in f.evidence
    assert f.severity.value == 'CRITICAL'  # recalibrage _SEVERITY_BY_KIND


# ------------------------------------------------------------------ T7 : validateurs
def test_nir_iban_structure():
    assert _nir_valid('185057800608436')
    assert _nir_valid('2 55 08 2B 073 004 17'.replace(' ', ''))
    assert not _nir_valid('585057800608436')   # sexe 5 invalide
    assert not _nir_valid('18505780060843')    # 14 chiffres
    assert _iban_valid('FR7630006000011234567890189')
    assert not _iban_valid('FR763000600001123456789018')   # 26 ≠ 27
    assert _iban_valid('XX12ABCDEFGH12345')                # pays inconnu, 15-34


def test_date_problem():
    assert _date_problem('30/02/2024')[0] == 'date_impossible'
    assert _date_problem('12/25/2023')[0] == 'date_us_format'
    assert _date_problem('01/01/2099')[0] == 'date_future'
    assert _date_problem('01/01/1850')[0] == 'date_implausible'
    assert _date_problem('15/06/2024') is None


def test_coherence_et_bornes(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Agent', 'Date naissance', 'Date recrutement', 'Salaire'])
    import datetime
    for i in range(6):
        ws.append([f'A{i}', datetime.date(1990, 1, 1), datetime.date(2015, 1, 1),
                   2000 + i])
    ws['C3'] = datetime.date(1980, 1, 1)   # recruté avant la naissance
    ws['D4'] = -500                        # salaire négatif
    p = _save(wb, tmp_path)
    kinds = {i.kind for i in analyze_validity(p)}
    assert 'field_incoherence' in kinds
    assert 'negative_value' in kinds


# --------------------------------------------------------------------- T8 : qualité
def test_total_mismatch(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Poste', 'Montant'])
    for i in range(1, 9):
        ws.append([f'P{i}', 100])
    ws.append(['TOTAL', 750])  # vrai total = 800
    p = _save(wb, tmp_path)
    issues = analyze_quality(p)
    assert any(i.kind == 'total_mismatch' for i in issues)


def test_doublons_et_lignes_masquees(tmp_path):
    wb = _wb_simple()
    ws = wb.active
    ws.append(['Ligne3', 103, '2024-01-03'])  # doublon strict de la ligne 4
    ws.row_dimensions[5].hidden = True
    p = _save(wb, tmp_path)
    kinds = {i.kind for i in analyze_quality(p)}
    assert 'duplicate_row' in kinds
    assert 'hidden_data' in kinds


def test_formule_ecrasee(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['x', 'double'])
    for i in range(2, 12):
        ws.cell(i, 1, i)
        ws.cell(i, 2, f'=A{i}*2')
    ws['B6'] = 999  # constante au milieu d'une colonne de formules cohérentes
    p = _save(wb, tmp_path)
    issues = analyze_quality(p)
    assert any(i.kind == 'overwritten_formula' and i.cell == 'B6' for i in issues)


# ------------------------------------------------------------ T9 : pipeline bout-en-bout
@pytest.fixture
def store(tmp_path):
    return JobStore(root=tmp_path / 'jobs')


def _run(store: JobStore, src: Path, name: str):
    job = store.create_job(name)
    store.input_path(job).write_bytes(src.read_bytes())
    return run_pipeline(store, job)


def test_pipeline_sain_certifie(tmp_path, store):
    p = _save(_wb_simple(), tmp_path)
    job = _run(store, p, 'sain.xlsx')
    assert job.status is JobStatus.STATIC_DONE
    report = json.loads(Path(job.report_json_path).read_text(encoding='utf-8'))
    assert report['preflight_summary']['passed'] is True
    assert report['verdict']['verdict'] in ('CERTIFIE', 'A_VERIFIER')
    assert Path(job.report_md_path).exists()


def test_pipeline_macro_rejected(tmp_path, store):
    p = _save(_wb_simple(), tmp_path)
    with zipfile.ZipFile(p, 'a') as zf:
        zf.writestr('xl/vbaProject.bin', b'\x00' * 64)
    job = _run(store, p, 'macro.xlsx')
    assert job.status is JobStatus.REJECTED
    report = json.loads(Path(job.report_json_path).read_text(encoding='utf-8'))
    assert report['preflight_summary']['reason_code'] == 'MACRO_VBAPROJECT'


def test_pipeline_nir_bloque(tmp_path, store):
    wb = Workbook()
    ws = wb.active
    ws.append(['Agent', 'NIR'])
    for i in range(6):
        ws.append([f'A{i}', '185057800608436'])
    p = _save(wb, tmp_path)
    job = _run(store, p, 'rh.xlsx')
    assert job.status is JobStatus.STATIC_DONE
    raw = Path(job.report_json_path).read_text(encoding='utf-8')
    report = json.loads(raw)
    assert report['verdict']['verdict'] == 'BLOQUE'
    assert '185057800608436' not in raw  # redaction de bout en bout


def test_pipeline_extension_rejetee(tmp_path, store):
    p = _save(_wb_simple(), tmp_path)
    job = _run(store, p, 'donnees.xlsm')
    assert job.status is JobStatus.REJECTED


# ------------------------------------------------------------------------ T10 : API
@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setenv('XLS_JOBS_DIR', str(tmp_path / 'jobs'))
    from fastapi.testclient import TestClient
    from app.main import app
    return TestClient(app)


def test_api_upload_et_rapports(client, tmp_path):
    p = _save(_wb_simple(), tmp_path)
    r = client.post('/upload', files={'file': ('ok.xlsx', p.read_bytes(),
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')})
    assert r.status_code == 200
    job_id = r.json()['job_id']
    assert client.get(f'/jobs/{job_id}').status_code == 200
    assert client.get(f'/reports/{job_id}.json').status_code == 200
    assert client.get(f'/reports/{job_id}.md').status_code == 200
    csv_resp = client.get(f'/reports/{job_id}.csv')
    assert csv_resp.status_code == 200
    assert csv_resp.content.startswith(b'\xef\xbb\xbf')  # BOM UTF-8


def test_api_csv_anti_injection(client, tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = '=HYPERLINK(evil)'  # titre hostile mais valide pour openpyxl
    ws.append(['Nom', 'Montant'])
    for i in range(8):
        ws.append([f'L{i}', 10])
    ws.append(['TOTAL', 70])  # faux total -> au moins une anomalie dans le CSV
    p = _save(wb, tmp_path)
    r = client.post('/upload', files={'file': ('inj.xlsx', p.read_bytes(), 'application/octet-stream')})
    job_id = r.json()['job_id']
    body = client.get(f'/reports/{job_id}.csv').content.decode('utf-8-sig')
    for line in body.splitlines()[1:]:
        for field in next(iter([line.split(',')]), []):
            cell = field.strip('"')
            if cell:
                assert cell[0] not in '=+-@', f'champ CSV injectable : {cell!r}'


def test_api_erreurs(client):
    assert client.get('/jobs/zzz').status_code == 400          # id non conforme
    assert client.get('/jobs/' + 'a' * 32).status_code == 404  # id conforme, absent
    assert client.get('/reports/' + 'a' * 32 + '.json').status_code == 404


def test_api_content_length_413(client):
    from app.config import MAX_UPLOAD_BYTES
    r = client.post('/upload', content=b'x',
                    headers={'content-length': str(MAX_UPLOAD_BYTES + 1)})
    assert r.status_code == 413
