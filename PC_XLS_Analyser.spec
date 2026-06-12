# -*- mode: python ; coding: utf-8 -*-
r"""PyInstaller spec — PC_XLS_Analyser (Windows x64, 100 % hors-ligne).

À LANCER SUR UNE MACHINE WINDOWS x64. PyInstaller ne fait PAS de compilation
croisée : un MacBook (ou ce conteneur Linux) ne peut pas produire un .exe Windows.
Le binaire fourni à l'origine a d'ailleurs bien été gelé sur Windows x64
(python312.dll win_amd64 présent dans le bundle).

Construit en mode --onedir (dossier) plutôt que --onefile : sur master durci
(AppLocker/WDAC interdisant l'exécution depuis %TEMP%), le onefile échoue car il
s'extrait dans le temp. Le onedir s'exécute en place → plus robuste, démarrage
plus rapide, et nettement moins de faux positifs antivirus.

Build :
    # 1) passe DEBUG (console visible pour lire les erreurs)
    #    -> mettez DEBUG = True ci-dessous, puis :
    pyinstaller PC_XLS_Analyser.spec --noconfirm
    # 2) passe RELEASE (sans console)
    #    -> remettez DEBUG = False, puis re-lancez la même commande.
    # Résultat : dist\PC_XLS_Analyser\PC_XLS_Analyser.exe
"""
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# --------------------------------------------------------------------------- #
DEBUG = False          # True = build console (diagnostic) ; False = build silencieux (prod)
APP_NAME = "PC_XLS_Analyser"
ENTRY = "app_main.py"
PROJECT_ROOT = os.path.abspath(os.getcwd())
# --------------------------------------------------------------------------- #


def _opt_tree(rel_dir: str):
    """Ajoute un dossier de données au bundle s'il existe (sinon, ignoré sans erreur).

    Couvre les ressources éventuelles lues à l'exécution : règles métier JSON
    (variable XLS_RULES_FILE / rules_engine.load_rules), schémas, fichiers
    statiques. Adaptez `rel_dir` à l'arborescence réelle de votre projet.
    """
    src = os.path.join(PROJECT_ROOT, rel_dir)
    return [(src, rel_dir)] if os.path.isdir(src) else []


datas = []
datas += collect_data_files("webview")          # DLLs WebView2 (loader) + ressources pywebview
datas += _opt_tree("rules")                     # ex. rules\\*.json  (moteur de règles)
datas += _opt_tree("schemas")                   # ex. schemas\\report.schema.json
datas += _opt_tree("static")                    # ex. static\\ (si servi depuis le disque)
datas += _opt_tree(os.path.join("app", "rules"))
datas += _opt_tree(os.path.join("app", "schemas"))
datas += _opt_tree(os.path.join("app", "static"))

hiddenimports = []
hiddenimports += collect_submodules("app")      # logique métier complète
hiddenimports += collect_submodules("uvicorn")  # protocoles/lifespan chargés dynamiquement
hiddenimports += [
    "anyio",
    "h11",
    "openpyxl",
    "clr_loader",                               # backend Windows de pywebview (pythonnet)
]

# On retire le superflu : dépendances de dev gelées par erreur dans le bundle
# d'origine (mypy, pytest, PyInstaller, pip...). Gain de poids + moins de faux
# positifs AV. On NE retire PAS tkinter (fenêtre 'Quitter' du mode navigateur).
excludes = [
    "mypy", "mypy_extensions", "mypyc",
    "pytest", "_pytest", "py",
    "PyInstaller",
    "IPython", "jedi", "black",
]


a = Analysis(
    [ENTRY],
    pathex=[PROJECT_ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,     # ← onedir
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX déconseillé : déclenche des AV, casse parfois WebView2
    console=DEBUG,             # piloté par le drapeau DEBUG ci-dessus
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,          # suit l'architecture de l'hôte de build (visez x64)
    codesign_identity=None,
    entitlements_file=None,
    # icon="assets\\app.ico",  # décommentez si vous avez une icône
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)
