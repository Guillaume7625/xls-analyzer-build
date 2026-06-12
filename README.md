# xls-analyzer v2 — source reconstruite, corrigée et testée

Reconstruction auditée du binaire `PC_XLS_Analyser.exe` (analyse locale .xlsx :
sécurité, qualité de données, PII/classification). Voir `RAPPORT_AUDIT.md` pour la
provenance fichier par fichier et les constats F-01…F-23.

## Lancer
    pip install -r requirements.txt
    python -m uvicorn app.main:app --host 127.0.0.1 --port 8000
    # ou : python app_main.py   (lanceur PyInstaller corrigé)

## Tester
    python -m pytest tests/ -q        # 44 tests

## Variables d'environnement
    XLS_JOBS_DIR                 répertoire des jobs (def. /run/xls-analyzer/jobs)
    XLS_MAX_ANALYSIS_CELLS       borne d'analyse complète (def. 200000)
    XLS_ANALYSIS_TIMEOUT_SECONDS budget temps contenu (def. 10)
    XLS_RULES_PATH               rules.json optionnel (cf. rules.example.json)
