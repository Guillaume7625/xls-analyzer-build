#!/usr/bin/env bash
# =============================================================================
#  get_exe.sh — obtient PC_XLS_Analyser.exe en UNE commande, depuis le Mac.
#
#  Ce que fait le script, tout seul :
#    1. crée (ou réutilise) un dépôt GitHub PRIVÉ "xls-analyzer-build"
#    2. y pousse le code (dossier v2/)
#    3. déclenche le build sur un runner Windows x64 GitHub
#    4. attend la fin (~4-5 min), puis télécharge l'exe dans ~/Downloads
#
#  Prérequis (UNE seule fois) :
#    brew install gh
#    gh auth login        # identité GitHub — la seule chose que je ne peux pas fournir
#
#  Usage :
#    cd v2/  &&  bash get_exe.sh
# =============================================================================
set -euo pipefail
cd "$(dirname "$0")"

REPO="xls-analyzer-build"
WF="build-exe.yml"
OUT="$HOME/Downloads/PC_XLS_Analyser"

err() { printf '\n[ERREUR] %s\n' "$*" >&2; exit 1; }

command -v gh  >/dev/null || err "gh manquant. Faire : brew install gh"
command -v git >/dev/null || err "git manquant (xcode-select --install)"
gh auth status >/dev/null 2>&1 || err "Non connecté à GitHub. Faire : gh auth login"
[ -f "$WF" ] || [ -f ".github/workflows/$WF" ] || err "À lancer depuis le dossier v2/ (workflow introuvable)"

USER="$(gh api user -q .login)"
echo "==> Compte GitHub : $USER"

# --- 1. dépôt privé -----------------------------------------------------------
if ! gh repo view "$USER/$REPO" >/dev/null 2>&1; then
  echo "==> Création du dépôt privé $USER/$REPO"
  gh repo create "$REPO" --private >/dev/null
else
  echo "==> Dépôt $USER/$REPO déjà présent (réutilisé)"
fi

# --- 2. push du code -----------------------------------------------------------
[ -f .gitignore ] || printf '.venv/\ndist/\nbuild/\n__pycache__/\n*.pyc\n.pytest_cache/\n' > .gitignore
git init -q -b main 2>/dev/null || true
git add -A
git commit -q -m "build $(date +%Y-%m-%d-%H%M%S)" 2>/dev/null || true
git remote remove origin 2>/dev/null || true
git remote add origin "https://github.com/$USER/$REPO.git"
echo "==> Push du code"
git push -q -f origin main

# --- 3. déclenchement du build (retry : le workflow met qq s à être enregistré) -
echo "==> Déclenchement du build Windows"
ok=""
for i in 1 2 3 4 5 6 7 8 9 10; do
  if gh workflow run "$WF" --repo "$USER/$REPO" --ref main 2>/dev/null; then ok=1; break; fi
  sleep 4
done
[ -n "$ok" ] || err "Impossible de déclencher le workflow (réessayer dans 1 min)"

# --- 4. identification du run, attente, téléchargement -------------------------
echo "==> Attente du démarrage du run…"
RUN_ID=""
for i in 1 2 3 4 5 6 7 8 9 10; do
  sleep 5
  RUN_ID="$(gh run list --repo "$USER/$REPO" --workflow "$WF" -L1 \
            --json databaseId -q '.[0].databaseId' 2>/dev/null || true)"
  [ -n "$RUN_ID" ] && break
done
[ -n "$RUN_ID" ] || err "Run introuvable — vérifier l'onglet Actions du dépôt"

echo "==> Build en cours (run $RUN_ID) — ~4 à 5 minutes (tests + gel PyInstaller)"
gh run watch "$RUN_ID" --repo "$USER/$REPO" --exit-status \
  || err "Le build a échoué. Journal : gh run view $RUN_ID --repo $USER/$REPO --log"

rm -rf "$OUT"; mkdir -p "$OUT"
gh run download "$RUN_ID" --repo "$USER/$REPO" -n PC_XLS_Analyser-onefile -D "$OUT"

EXE="$OUT/PC_XLS_Analyser.exe"
[ -f "$EXE" ] || err "Artefact téléchargé mais exe introuvable dans $OUT"
echo
echo "============================================================="
echo "  ✅  $EXE"
echo "  SHA-256 : $(shasum -a 256 "$EXE" | awk '{print $1}')"
echo "============================================================="
