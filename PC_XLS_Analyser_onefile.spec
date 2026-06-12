# -*- mode: python ; coding: utf-8 -*-
r"""PyInstaller spec — PC_XLS_Analyser en FICHIER UNIQUE (--onefile, Windows x64).

Produit UN SEUL fichier : dist\PC_XLS_Analyser.exe (tout embarqué : interpréteur
Python 3.12, bibliothèques, données). À LANCER SUR WINDOWS x64 — PyInstaller ne
fait pas de compilation croisée.

Build :
    pyinstaller PC_XLS_Analyser_onefile.spec --noconfirm
    # Résultat : dist\PC_XLS_Analyser.exe   (fichier unique)

Fonctionnement à l'exécution : le bootloader extrait le bundle dans un dossier
temporaire (par défaut %TEMP%\_MEIxxxxxx) puis lance l'app. Conséquences :
  - démarrage plus lent que le onedir (extraction à chaque lancement) ;
  - si la politique du poste (AppLocker/WDAC) interdit l'exécution depuis %TEMP%,
    décommenter `runtime_tmpdir` ci-dessous pour extraire dans un chemin autorisé.
Pour un déploiement parc durci, la variante onedir (PC_XLS_Analyser.spec) reste
disponible ; les deux specs partagent la même Analysis.
"""
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# --------------------------------------------------------------------------- #
DEBUG = False          # True = console visible (diagnostic) ; False = silencieux (prod)
APP_NAME = "PC_XLS_Analyser"
ENTRY = "app_main.py"
PROJECT_ROOT = os.path.abspath(os.getcwd())
# --------------------------------------------------------------------------- #


def _opt_tree(rel_dir: str):
    """Ajoute un dossier de données au bundle s'il existe (sinon ignoré)."""
    src = os.path.join(PROJECT_ROOT, rel_dir)
    return [(src, rel_dir)] if os.path.isdir(src) else []


datas = []
datas += collect_data_files("webview")          # loader WebView2 + ressources pywebview
datas += _opt_tree("rules")
datas += _opt_tree("schemas")
datas += _opt_tree("static")
datas += _opt_tree(os.path.join("app", "rules"))
datas += _opt_tree(os.path.join("app", "schemas"))
datas += _opt_tree(os.path.join("app", "static"))

hiddenimports = []
hiddenimports += collect_submodules("app")
hiddenimports += collect_submodules("uvicorn")
hiddenimports += [
    "anyio",
    "h11",
    "openpyxl",
    "clr_loader",
]

excludes = [
    "mypy", "mypy_extensions", "mypyc",
    "pytest", "_pytest", "py",
    "PyInstaller",
    "pip", "wheel", "setuptools._distutils",
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
    a.binaries,                # ← onefile : binaires ET données embarqués dans l'exe
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX déconseillé : faux positifs AV, casse parfois WebView2
    upx_exclude=[],
    # runtime_tmpdir=r"C:\xls-analyzer\run",  # ← décommenter si %TEMP% est interdit
    console=DEBUG,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="assets\\app.ico",
)
