@echo off
REM ============================================================================
REM  build_onefile.bat — fabrique dist\PC_XLS_Analyser.exe (FICHIER UNIQUE)
REM  Prerequis : Windows x64 + Python 3.12 installe ("Add python.exe to PATH").
REM  Hors-ligne : lancer d'abord prepare_offline_bundle.bat sur un poste connecte,
REM  puis ce script detectera ..\wheels et installera sans reseau.
REM ============================================================================
setlocal
cd /d "%~dp0"

python --version || (echo [ERREUR] Python introuvable dans le PATH & pause & exit /b 1)

if not exist .venv (
  python -m venv .venv || (echo [ERREUR] creation venv & pause & exit /b 1)
)
call .venv\Scripts\activate.bat

if exist ..\wheels (
  echo [MODE HORS-LIGNE] installation depuis ..\wheels
  python -m pip install --no-index --find-links ..\wheels -r requirements.txt || goto :err
) else (
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt || goto :err
)

pyinstaller PC_XLS_Analyser_onefile.spec --noconfirm || goto :err

echo.
echo ============================================================
echo  OK : dist\PC_XLS_Analyser.exe   (fichier unique, autonome)
echo ============================================================
pause
exit /b 0

:err
echo.
echo [ECHEC] voir les messages ci-dessus. Pour diagnostiquer un crash au
echo lancement de l'exe : DEBUG = True dans PC_XLS_Analyser_onefile.spec,
echo rebuild, lancer l'exe depuis une console.
pause
exit /b 1
