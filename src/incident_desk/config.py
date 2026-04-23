"""Process-wide constants: paths, display formats, app identity.

Importing this module also triggers Windows DPI-awareness setup (must run
before any tkinter window is created)."""
from __future__ import annotations

import ctypes
import os
import sys
from pathlib import Path

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


try:
    if getattr(sys, "frozen", False):
        BASE_DIR = Path(sys.executable).parent
    else:
        BASE_DIR = Path(__file__).resolve().parent
except Exception:
    BASE_DIR = Path(__file__).resolve().parent

DB_DIR = BASE_DIR
try:
    testfile = DB_DIR / ".__writetest"
    with open(testfile, "w", encoding="utf-8") as _f:
        _f.write("ok")
    testfile.unlink(missing_ok=True)
except Exception:
    # Program Files isn't writable; fall back to per-user data dir
    LOCAL = Path(os.environ.get("LOCALAPPDATA", str(BASE_DIR))) / "IncidentDesk"
    LOCAL.mkdir(parents=True, exist_ok=True)
    DB_DIR = LOCAL

DB_PATH = DB_DIR / "incidentdesk.db"

APP_TITLE = "Road America – Race Control - Incident Desk"
APP_VERSION = "1.5"                       # bump before cutting a new GitHub release
GITHUB_REPO = "ryderjsmith/IncidentDesk"  # owner/repo for update checks

# Dev layout: project_root/src/incident_desk/config.py → project_root/img/favicon.ico
# is three parents up.
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
_ico_candidates = [
    Path(getattr(sys, "_MEIPASS", "")) / "img" / "favicon.ico",  # PyInstaller bundle
    _PROJECT_ROOT / "img" / "favicon.ico",                       # dev: project_root/img/
    BASE_DIR / "img" / "favicon.ico",                            # fallback
]
ICO_PATH = next((p for p in _ico_candidates if p.exists()), _ico_candidates[-1])

DT_FMT = "%Y-%m-%d %H:%M:%S"           # ISO — DB storage / queries
DATE_FMT = "%Y-%m-%d"                  # ISO — DB date queries
DISPLAY_DT_FMT = "%m-%d-%Y %H:%M:%S"   # MM-DD-YYYY — UI display
DISPLAY_DATE_FMT = "%m-%d-%Y"          # MM-DD-YYYY — UI date fields
