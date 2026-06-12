# -*- coding: utf-8 -*-
"""Lanceur desktop de xls-analyzer — Windows, 100 % hors-ligne.

Démarre l'app FastAPI existante (app.main:app) sur un port loopback libre, puis
affiche l'interface. Corrige la cause RÉELLE du non-démarrage, observée dans la
trace du binaire d'origine :

    AttributeError: 'NoneType' object has no attribute 'isatty'
        ... uvicorn\\config.py, configure_logging

CAUSE. Build PyInstaller en sous-système GUI (--windowed) : il n'y a AUCUNE
console attachée, donc sys.stdout et sys.stderr valent None. La config de
logging par défaut d'uvicorn instancie un formateur colorisé qui appelle
sys.stdout.isatty() pour décider d'activer la couleur → crash immédiat, AVANT
même l'ouverture de la fenêtre. Tout code (uvicorn ou autre) qui écrit sur
stdout dans un build windowed plante de la même façon.

CORRECTIF (deux verrous complémentaires) :
  1. On garantit que sys.stdout / sys.stderr ne sont jamais None (redirection
     vers un fichier de capture) → neutralise isatty() ET tout print() parasite.
  2. On passe log_config=None à uvicorn.Config → uvicorn n'exécute plus son
     dictConfig (la ligne exacte qui plante). Nos logs passent par notre propre
     FileHandler.

Plus : journalisation complète + affichage tolérant aux pannes (fenêtre native
pywebview/WebView2, et bascule AUTOMATIQUE sur le navigateur si WebView2 absent,
fréquent en air-gap). Forçage via XLS_UI = auto | webview | browser.

Aucune logique d'analyse ici. Aucun Docker. Racine des jobs = dossier temporaire
éphémère, effacé à la fermeture (fail-closed : pas de résidu).
"""
from __future__ import annotations

import logging
import os
import shutil
import socket
import sys
import tempfile
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

APP_TITLE = "XLS Analyzer"


# --------------------------------------------------------------------------- #
# Dossier de journalisation (calculé en premier : tout en dépend)
# --------------------------------------------------------------------------- #
def _log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    try:
        d = Path(base) / "xls-analyzer"
        d.mkdir(parents=True, exist_ok=True)
        return d
    except Exception:
        d = Path(tempfile.gettempdir()) / "xls-analyzer"
        d.mkdir(parents=True, exist_ok=True)
        return d


LOG_DIR = _log_dir()
LOG_PATH = LOG_DIR / "launcher.log"


# --------------------------------------------------------------------------- #
# VERROU 1 — garantir des flux standard non-None AVANT toute config de logging
# et avant l'import d'uvicorn. C'est LE correctif de l'AttributeError isatty.
# --------------------------------------------------------------------------- #
def _ensure_std_streams() -> None:
    if sys.stdout is not None and sys.stderr is not None:
        return
    try:
        cap = open(LOG_DIR / "stdout.log", "a", encoding="utf-8", buffering=1)
    except Exception:
        cap = open(os.devnull, "w")
    if sys.stdout is None:
        sys.stdout = cap
    if sys.stderr is None:
        sys.stderr = cap


_ensure_std_streams()

# FileHandler uniquement : pas de StreamHandler(sys.stdout) — inutile en windowed
# et source d'un crash similaire si le flux est None.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8")],
)
log = logging.getLogger("xls-analyzer.launcher")


# --------------------------------------------------------------------------- #
# Serveur local
# --------------------------------------------------------------------------- #
def _free_port() -> int:
    """Réserve un port loopback libre et le rend (bind 127.0.0.1:0)."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_ready(url: str, timeout: float = 30.0) -> bool:
    """Sonde l'URL racine jusqu'à HTTP 200 ou expiration."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1.0) as resp:
                if resp.status == 200:
                    return True
        except Exception:
            time.sleep(0.2)
    return False


def start_server(host: str, port: int, jobs_dir: str):
    """Lance uvicorn(app.main:app) dans un thread démon. Renvoie le serveur."""
    os.environ["XLS_JOBS_DIR"] = jobs_dir

    import uvicorn
    from app.main import app  # app FastAPI existante, inchangée

    # VERROU 2 — log_config=None : uvicorn n'exécute plus son dictConfig (la ligne
    # qui appelle isatty et plante en windowed). Nos logs passent par root/FileHandler.
    config = uvicorn.Config(
        app,
        host=host,
        port=port,
        log_level="warning",
        log_config=None,
    )
    server = uvicorn.Server(config)
    # On n'est pas dans le thread principal d'uvicorn : on neutralise ses handlers.
    server.install_signal_handlers = lambda: None
    threading.Thread(target=server.run, daemon=True).start()
    return server


# --------------------------------------------------------------------------- #
# Affichage de l'interface
# --------------------------------------------------------------------------- #
def _try_native_window(url: str) -> bool:
    """Tente la fenêtre native pywebview/WebView2. True si affichée (bloque
    jusqu'à fermeture). Toute panne (pywebview absent, runtime WebView2 manquant,
    échec de création) renvoie False → fallback navigateur."""
    try:
        import webview
    except Exception as exc:  # noqa: BLE001
        log.warning("pywebview indisponible (%s) — bascule navigateur", exc)
        return False
    try:
        webview.create_window(APP_TITLE, url, width=1100, height=820)
        webview.start()
        return True
    except Exception as exc:  # noqa: BLE001 — typiquement runtime WebView2 absent
        log.warning("Fenetre native impossible (%s) — bascule navigateur", exc)
        return False


def _tk_quit_window(url: str) -> bool:
    """Petite fenêtre 'Quitter' (tkinter) pour le mode navigateur. False si absent."""
    try:
        import tkinter as tk
    except Exception:  # noqa: BLE001
        return False
    root = tk.Tk()
    root.title(APP_TITLE)
    try:
        root.attributes("-topmost", True)
    except Exception:  # noqa: BLE001
        pass
    tk.Label(root, text=f"{APP_TITLE} est ouvert dans votre navigateur :",
             padx=16, pady=(16, 4)).pack()
    entry = tk.Entry(root, width=46)
    entry.insert(0, url)
    entry.configure(state="readonly")
    entry.pack(padx=16)
    tk.Button(root, text="Quitter XLS Analyzer", padx=12, pady=6,
              command=root.destroy).pack(pady=14)
    root.geometry("440x160")
    root.mainloop()
    return True


def _open_browser_and_block(url: str) -> None:
    """Ouvre le navigateur par défaut et maintient le serveur en vie."""
    try:
        webbrowser.open(url, new=1)
        log.info("Interface ouverte dans le navigateur : %s", url)
    except Exception as exc:  # noqa: BLE001
        log.error("Ouverture navigateur impossible (%s)", exc)
    try:
        (LOG_DIR / "URL.txt").write_text(url, encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    if not _tk_quit_window(url):
        log.info("Serveur actif. Pour quitter : fermez ce processus.")
        try:
            threading.Event().wait()
        except KeyboardInterrupt:
            pass


def show_ui(url: str) -> None:
    mode = os.environ.get("XLS_UI", "auto").strip().lower()
    if mode in ("auto", "webview"):
        if _try_native_window(url):
            return
        if mode == "webview":
            log.warning("XLS_UI=webview demande mais indisponible — fallback navigateur.")
    _open_browser_and_block(url)


# --------------------------------------------------------------------------- #
# Point d'entrée
# --------------------------------------------------------------------------- #
def main() -> int:
    log.info("=== Demarrage %s (journal : %s) ===", APP_TITLE, LOG_PATH)
    jobs_dir = None
    server = None
    try:
        host = "127.0.0.1"
        port = _free_port()
        jobs_dir = tempfile.mkdtemp(prefix="xls-analyzer-jobs-")
        server = start_server(host, port, jobs_dir)

        url = f"http://{host}:{port}/"
        if not _wait_ready(url, timeout=30.0):
            raise RuntimeError("Le serveur local n'a pas demarre a temps (30 s).")
        log.info("Serveur pret : %s", url)

        show_ui(url)
        return 0

    except Exception:  # noqa: BLE001 — on CAPTURE l'echec silencieux d'origine
        log.exception("Echec du lanceur %s", APP_TITLE)
        try:
            import tkinter
            import tkinter.messagebox as mb
            r = tkinter.Tk()
            r.withdraw()
            mb.showerror(
                APP_TITLE,
                f"Echec au demarrage.\n\nConsultez le journal :\n{LOG_PATH}",
            )
            r.destroy()
        except Exception:  # noqa: BLE001
            pass
        return 1

    finally:
        if server is not None:
            server.should_exit = True
            time.sleep(0.3)
        if jobs_dir:
            shutil.rmtree(jobs_dir, ignore_errors=True)  # fail-closed : pas de residu


if __name__ == "__main__":
    raise SystemExit(main())
