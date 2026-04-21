"""
Incident Desk - simple offline incident logging app

Developed by Ryder Smith 2025-2026
-----------------------------------------------------------------
• Offline-first: SQLite file stored next to the script (incidentdesk.db)
• Modern-ish Tkinter/ttk UI; color‑coded incident board
• Add/Edit incidents with units and time stamps
• Filter/search by date, location, and incident type
• Editable & Importable/Exportable pick‑lists for Locations, Units, and Incident Types
• Notes per incident with automatic timestamps
• Export board to Excel (xlsx) or CSV; export/print to PDF

"""
from __future__ import annotations
import os
import csv
import json
import sqlite3
import subprocess
import tempfile
import threading
import urllib.error
import urllib.request
import webbrowser
from datetime import datetime, date
from pathlib import Path
from typing import List, Optional, Tuple
import sys
import ctypes

# Tell Windows this process is DPI-aware so it renders text crisply
# instead of bitmap-scaling the entire window (which causes blur).
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk, filedialog
from tkcalendar import DateEntry as _DateEntry


class DateEntry(_DateEntry):
    """DateEntry that applies the dark title bar to its calendar dropdown."""
    def drop_down(self):
        super().drop_down()
        try:
            apply_dark_titlebar(self._top_cal)
        except Exception:
            pass



# Ensure DB is writable: prefer exe folder; fallback to %LOCALAPPDATA%\RA_TrackIncident when needed
try:
    if getattr(sys, 'frozen', False):
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
APP_VERSION = "1.3"                     # bump before cutting a new GitHub release
GITHUB_REPO = "ryderjsmith/IncidentDesk"  # owner/repo for update checks
_ico_candidates = [
    Path(getattr(sys, "_MEIPASS", "")) / "img" / "favicon.ico",  # PyInstaller bundle
    BASE_DIR.parent / "img" / "favicon.ico",                      # dev: src/../img/
    BASE_DIR / "img" / "favicon.ico",                             # fallback
]
ICO_PATH = next((p for p in _ico_candidates if p.exists()), _ico_candidates[-1])
DT_FMT = "%Y-%m-%d %H:%M:%S"        # ISO — DB storage / queries
DATE_FMT = "%Y-%m-%d"               # ISO — DB date queries
DISPLAY_DT_FMT = "%m-%d-%Y %H:%M:%S"   # MM-DD-YYYY — UI display
DISPLAY_DATE_FMT = "%m-%d-%Y"          # MM-DD-YYYY — UI date fields

def set_icon(window) -> None:
    """Apply favicon.ico to any Tk or Toplevel window."""
    try:
        window.iconbitmap(str(ICO_PATH))
    except Exception:
        pass


_WINDOW_ICONS = {
    "incident_form": ("🚨", "#c0392b"),
    "notes":         ("📝", "#2980b9"),
    "billables":     ("🧾", "#d4a74a"),
    "units":         ("🚗", "#27ae60"),
    "locations":     ("📍", "#8e44ad"),
    "types":         ("🏷", "#e67e22"),
    "info":          ("ℹ",  "#2980b9"),
    "confirm":       ("⚠",  "#e67e22"),
    "input":         ("✏",  "#16a085"),
    "guide":         ("📖", "#2c3e50"),
    "export":        ("💾", "#546e7a"),
}
_icon_cache: dict = {}

def _build_icon(char: str) -> list:
    """Return a list of PhotoImages at [16, 32, 48, 256] px so Windows picks the best size per context."""
    try:
        from PIL import Image, ImageDraw, ImageFont, ImageTk
        font_cache: dict = {}
        photos = []
        for size in (16, 32, 48, 256):
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            fs = max(8, int(size * 0.72))
            if fs not in font_cache:
                font = None
                for name in [
                    r"C:\Windows\Fonts\seguiemj.ttf",
                    r"C:\Windows\Fonts\seguisym.ttf",
                    "Segoe UI Emoji", "Segoe UI Symbol", "Arial",
                ]:
                    try:
                        font = ImageFont.truetype(name, fs)
                        break
                    except Exception:
                        pass
                font_cache[fs] = font or ImageFont.load_default()
            font = font_cache[fs]
            bb = d.textbbox((0, 0), char, font=font)
            x = (size - (bb[2] - bb[0])) // 2 - bb[0]
            y = (size - (bb[3] - bb[1])) // 2 - bb[1]
            d.text((x, y), char, fill="white", font=font)
            photos.append(ImageTk.PhotoImage(img))
        return photos
    except Exception:
        return []

def set_window_icon(window, key: str) -> None:
    """Apply a type-specific icon to a Toplevel. Falls back to favicon.ico on failure."""
    if key not in _icon_cache:
        spec = _WINDOW_ICONS.get(key)
        _icon_cache[key] = _build_icon(spec[0]) if spec else []
    photos = _icon_cache.get(key) or []
    if photos:
        window.iconphoto(False, *photos)
        window._window_icons = photos  # prevent GC
    else:
        set_icon(window)


_cascade_count = 0
_CASCADE_OFFSET = 28


def position_on_parent(window, parent) -> None:
    """Center window over parent, cascading by _CASCADE_OFFSET per open window."""
    global _cascade_count
    window.update_idletasks()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    ww = window.winfo_reqwidth()
    wh = window.winfo_reqheight()
    offset = _cascade_count * _CASCADE_OFFSET
    x = px + (pw - ww) // 2 + offset
    y = py + (ph - wh) // 2 + offset
    window.geometry(f"+{x}+{y}")
    window.deiconify()
    _cascade_count += 1

    def _on_destroy(e):
        global _cascade_count
        if e.widget is window:
            _cascade_count = max(0, _cascade_count - 1)

    window.bind("<Destroy>", _on_destroy, add="+")


def dark_info(parent, title: str, message: str) -> None:
    """Info dialog styled to match the rest of the app."""
    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title(title)
    dlg.resizable(False, False)
    set_window_icon(dlg, "info")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=message, wraplength=320, justify="left").pack(anchor="w")
    ttk.Separator(outer).pack(fill="x", pady=(16, 12))
    ttk.Button(outer, text="OK", style="Manage.TButton", command=dlg.destroy).pack(side="right")

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()


def dark_confirm(parent, title: str, message: str) -> bool:
    """Styled Yes/No confirmation dialog. Returns True if the user clicked Yes."""
    result = [False]

    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title(title)
    dlg.resizable(False, False)
    set_window_icon(dlg, "confirm")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=message, wraplength=320, justify="left").pack(anchor="w")
    ttk.Separator(outer).pack(fill="x", pady=(16, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _yes():
        result[0] = True
        dlg.destroy()

    ttk.Button(btn_row, text="No",  command=dlg.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="Yes", style="Danger.TButton", command=_yes).pack(side="right")

    dlg.bind("<Return>", lambda e: _yes())
    dlg.bind("<Escape>", lambda e: dlg.destroy())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()

    return result[0]


def apply_dark_titlebar(window) -> None:
    """Enable dark title bar on Windows 10 (build 18985+) / Windows 11."""
    try:
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        value = ctypes.c_int(1)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value), ctypes.sizeof(value)
        )
    except Exception:
        pass


# -----------------------------
# Update checking (GitHub Releases)
# -----------------------------
def _parse_version(v: str) -> tuple:
    """Turn 'v1.2.3' / '1.2.3-beta' into a comparable tuple (1, 2, 3)."""
    s = v.lstrip("vV").split("-")[0].split("+")[0]
    out = []
    for p in s.split("."):
        try:
            out.append(int(p))
        except ValueError:
            out.append(0)
    return tuple(out)


def check_for_update() -> Optional[Tuple[str, str]]:
    """Query GitHub for the latest release. Returns (tag, installer_url) if
    newer than APP_VERSION, else None. Fails silently on network errors."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(url, headers={
        "User-Agent": "IncidentDesk-Updater",
        "Accept": "application/vnd.github+json",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError):
        return None

    tag = data.get("tag_name", "")
    if not tag or _parse_version(tag) <= _parse_version(APP_VERSION):
        return None

    installer = next(
        (a for a in data.get("assets", []) if a.get("name", "").lower().endswith("-setup.exe")),
        None,
    )
    if not installer:
        return None
    return (tag, installer["browser_download_url"])


def download_file(url: str, dest: Path) -> bool:
    """Stream download to dest. Returns True on success."""
    req = urllib.request.Request(url, headers={"User-Agent": "IncidentDesk-Updater"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
        return True
    except (urllib.error.URLError, urllib.error.HTTPError, OSError):
        return False


def update_prompt(parent, current: str, latest: str) -> bool:
    """Styled update-available dialog. Returns True if user chose Update Now."""
    result = [False]

    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title("Update Available")
    dlg.resizable(False, False)
    set_window_icon(dlg, "export")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    message = (
        f"Incident Desk {latest} is available.\n"
        f"You are currently running {current}.\n\n"
        "Would you like to download and install the update now?\n\n"
        "Your incident data will be preserved automatically."
    )
    ttk.Label(outer, text=message, wraplength=380, justify="left").pack(anchor="w")
    ttk.Separator(outer).pack(fill="x", pady=(16, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _yes():
        result[0] = True
        dlg.destroy()

    ttk.Button(btn_row, text="Later", command=dlg.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="Update Now", style="Manage.TButton", command=_yes).pack(side="right")

    dlg.bind("<Return>", lambda e: _yes())
    dlg.bind("<Escape>", lambda e: dlg.destroy())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()

    return result[0]


# -----------------------------
# Database
# -----------------------------
class DB:
    def __init__(self, path: Path = DB_PATH):
        self.path = path
        self.conn = sqlite3.connect(self.path)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        cur = self.conn.cursor()
        cur.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS locations (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS units (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE,
                category TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS incident_types (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE
            );

            CREATE TABLE IF NOT EXISTS incidents (
                id INTEGER PRIMARY KEY,
                location_id INTEGER,
                type TEXT NOT NULL,
                reported_at TEXT NOT NULL,
                dispatched_at TEXT DEFAULT '',
                arrived_at TEXT DEFAULT '',
                cleared_at TEXT DEFAULT '',
                disposition TEXT DEFAULT '',
                car_number TEXT DEFAULT '',
                is_cleared INTEGER DEFAULT 0,
                FOREIGN KEY(location_id) REFERENCES locations(id)
            );

            CREATE TABLE IF NOT EXISTS incident_units (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                unit_id INTEGER NOT NULL,
                role TEXT NOT NULL CHECK(role in ('primary','backup')),
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE,
                FOREIGN KEY(unit_id) REFERENCES units(id)
            );

            CREATE TABLE IF NOT EXISTS notes (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                ts TEXT NOT NULL,
                body TEXT NOT NULL,
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS billables (
                id INTEGER PRIMARY KEY,
                incident_id INTEGER NOT NULL,
                body TEXT NOT NULL,
                FOREIGN KEY(incident_id) REFERENCES incidents(id) ON DELETE CASCADE
            );
            """
        )
        self.conn.commit()
        self._migrate()

    def _migrate(self):
        cur = self.conn.cursor()
        for stmt in [
            "ALTER TABLE locations ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE incident_types ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE units ADD COLUMN sort_order INTEGER DEFAULT 0",
            "ALTER TABLE incidents ADD COLUMN car_number TEXT DEFAULT ''",
        ]:
            try:
                cur.execute(stmt)
            except Exception:
                pass  # column already exists
        # Initialise sort_order for any rows that still have 0 (existing data)
        cur.execute("UPDATE locations SET sort_order = id WHERE sort_order = 0")
        cur.execute("UPDATE incident_types SET sort_order = id WHERE sort_order = 0")
        cur.execute("UPDATE units SET sort_order = id WHERE sort_order = 0")
        self.conn.commit()

    # ---- CRUD helpers
    def list_locations(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM locations ORDER BY sort_order, id").fetchall()

    def add_location(self, name: str):
        self.conn.execute(
            "INSERT OR IGNORE INTO locations(name, sort_order) "
            "VALUES(?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM locations))",
            (name.strip(),),
        )
        self.conn.commit()

    def rename_location(self, loc_id: int, new_name: str):
        self.conn.execute("UPDATE locations SET name=? WHERE id=?", (new_name.strip(), loc_id))
        self.conn.commit()

    def location_incident_count(self, loc_id: int) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incidents WHERE location_id=?", (loc_id,)
        ).fetchone()[0]

    def delete_location(self, loc_id: int):
        self.conn.execute("DELETE FROM locations WHERE id=?", (loc_id,))
        self.conn.commit()

    def list_units(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM units ORDER BY sort_order, id").fetchall()

    def list_units_with_availability(self) -> List[sqlite3.Row]:
        """Returns all units with a computed `available` column (1=available, 0=unavailable)."""
        return self.conn.execute("""
            SELECT u.*,
                CASE WHEN EXISTS (
                    SELECT 1 FROM incident_units iu
                    JOIN incidents i ON i.id = iu.incident_id
                    WHERE iu.unit_id = u.id AND iu.role = 'primary' AND i.is_cleared = 0
                ) THEN 0 ELSE 1 END AS available
            FROM units u ORDER BY u.sort_order, u.id
        """).fetchall()

    def list_available_units(self, exclude_incident_id: Optional[int] = None) -> List[sqlite3.Row]:
        """Returns units not currently assigned as primary on any active incident.
        exclude_incident_id: ignore assignments belonging to this incident (used when editing)."""
        if exclude_incident_id:
            return self.conn.execute("""
                SELECT u.* FROM units u
                WHERE NOT EXISTS (
                    SELECT 1 FROM incident_units iu
                    JOIN incidents i ON i.id = iu.incident_id
                    WHERE iu.unit_id = u.id AND iu.role = 'primary'
                      AND i.is_cleared = 0 AND i.id != ?
                ) ORDER BY u.sort_order, u.id
            """, (exclude_incident_id,)).fetchall()
        return self.conn.execute("""
            SELECT u.* FROM units u
            WHERE NOT EXISTS (
                SELECT 1 FROM incident_units iu
                JOIN incidents i ON i.id = iu.incident_id
                WHERE iu.unit_id = u.id AND iu.role = 'primary' AND i.is_cleared = 0
            ) ORDER BY u.sort_order, u.id
        """).fetchall()

    def swap_unit_order(self, id1: int, id2: int):
        r1 = self.conn.execute("SELECT sort_order FROM units WHERE id=?", (id1,)).fetchone()
        r2 = self.conn.execute("SELECT sort_order FROM units WHERE id=?", (id2,)).fetchone()
        if r1 and r2:
            self.conn.execute("UPDATE units SET sort_order=? WHERE id=?", (r2["sort_order"], id1))
            self.conn.execute("UPDATE units SET sort_order=? WHERE id=?", (r1["sort_order"], id2))
            self.conn.commit()

    def add_unit(self, name: str, category: str = ""):
        self.conn.execute(
            "INSERT OR IGNORE INTO units(name, category, sort_order) "
            "VALUES(?,?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM units))",
            (name.strip(), category.strip()),
        )
        self.conn.commit()

    def update_unit(self, unit_id: int, name: str, category: str):
        self.conn.execute("UPDATE units SET name=?, category=? WHERE id=?", (name.strip(), category.strip(), unit_id))
        self.conn.commit()

    def unit_incident_count(self, unit_id: int) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incident_units WHERE unit_id=?", (unit_id,)
        ).fetchone()[0]

    def delete_unit(self, unit_id: int):
        self.conn.execute("DELETE FROM units WHERE id=?", (unit_id,))
        self.conn.commit()

    def list_incident_types(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM incident_types ORDER BY sort_order, id").fetchall()

    def add_incident_type(self, name: str):
        self.conn.execute(
            "INSERT OR IGNORE INTO incident_types(name, sort_order) "
            "VALUES(?, (SELECT COALESCE(MAX(sort_order), 0) + 1 FROM incident_types))",
            (name.strip(),),
        )
        self.conn.commit()

    def swap_location_order(self, id1: int, id2: int):
        r1 = self.conn.execute("SELECT sort_order FROM locations WHERE id=?", (id1,)).fetchone()
        r2 = self.conn.execute("SELECT sort_order FROM locations WHERE id=?", (id2,)).fetchone()
        if r1 and r2:
            self.conn.execute("UPDATE locations SET sort_order=? WHERE id=?", (r2["sort_order"], id1))
            self.conn.execute("UPDATE locations SET sort_order=? WHERE id=?", (r1["sort_order"], id2))
            self.conn.commit()

    def swap_incident_type_order(self, id1: int, id2: int):
        r1 = self.conn.execute("SELECT sort_order FROM incident_types WHERE id=?", (id1,)).fetchone()
        r2 = self.conn.execute("SELECT sort_order FROM incident_types WHERE id=?", (id2,)).fetchone()
        if r1 and r2:
            self.conn.execute("UPDATE incident_types SET sort_order=? WHERE id=?", (r2["sort_order"], id1))
            self.conn.execute("UPDATE incident_types SET sort_order=? WHERE id=?", (r1["sort_order"], id2))
            self.conn.commit()

    def rename_incident_type(self, type_id: int, new_name: str):
        self.conn.execute("UPDATE incident_types SET name=? WHERE id=?", (new_name.strip(), type_id))
        self.conn.commit()

    def incident_type_incident_count(self, type_name: str) -> int:
        return self.conn.execute(
            "SELECT COUNT(*) FROM incidents WHERE type=?", (type_name,)
        ).fetchone()[0]

    def delete_incident_type(self, type_id: int):
        self.conn.execute("DELETE FROM incident_types WHERE id=?", (type_id,))
        self.conn.commit()

    def create_incident(self, location_id: Optional[int], type_name: str, reported_at: str,
                         dispatched_at: str = "", arrived_at: str = "", cleared_at: str = "",
                         disposition: str = "", is_cleared: int = 0,
                         primary_unit_id: Optional[int] = None, backup_unit_ids: Optional[List[int]] = None,
                         car_number: str = "") -> int:
        cur = self.conn.cursor()
        cur.execute(
            """
            INSERT INTO incidents(location_id, type, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, is_cleared)
            VALUES(?,?,?,?,?,?,?,?,?)
            """,
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, is_cleared),
        )
        inc_id = cur.lastrowid
        if primary_unit_id:
            cur.execute(
                "INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'primary')",
                (inc_id, primary_unit_id),
            )
        if backup_unit_ids:
            for uid in backup_unit_ids:
                cur.execute(
                    "INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'backup')",
                    (inc_id, uid),
                )
        self.conn.commit()
        return inc_id

    def update_incident(self, inc_id: int, location_id: Optional[int], type_name: str, reported_at: str,
                         dispatched_at: str, arrived_at: str, cleared_at: str, disposition: str, is_cleared: int,
                         primary_unit_id: Optional[int], backup_unit_ids: List[int],
                         car_number: str = ""):
        cur = self.conn.cursor()
        cur.execute(
            "UPDATE incidents SET location_id=?, type=?, reported_at=?, dispatched_at=?, arrived_at=?, cleared_at=?, disposition=?, car_number=?, is_cleared=? WHERE id=?",
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, car_number, is_cleared, inc_id),
        )
        # reset assignments
        cur.execute("DELETE FROM incident_units WHERE incident_id=?", (inc_id,))
        if primary_unit_id:
            cur.execute("INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'primary')", (inc_id, primary_unit_id))
        for uid in backup_unit_ids:
            cur.execute("INSERT INTO incident_units(incident_id, unit_id, role) VALUES(?,?, 'backup')", (inc_id, uid))
        self.conn.commit()

    def delete_incident(self, inc_id: int):
        self.conn.execute("DELETE FROM incidents WHERE id=?", (inc_id,))
        self.conn.commit()

    def add_note(self, inc_id: int, ts: str, body: str):
        self.conn.execute("INSERT INTO notes(incident_id, ts, body) VALUES(?,?,?)", (inc_id, ts, body))
        self.conn.commit()

    def list_notes(self, inc_id: int) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM notes WHERE incident_id=? ORDER BY ts", (inc_id,)).fetchall()

    def add_billable(self, inc_id: int, body: str):
        self.conn.execute("INSERT INTO billables(incident_id, body) VALUES(?,?)", (inc_id, body))
        self.conn.commit()

    def list_billables(self, inc_id: int) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM billables WHERE incident_id=? ORDER BY id", (inc_id,)).fetchall()

    def set_cleared(self, inc_id: int, cleared: bool, cleared_at: Optional[str] = None):
        cur = self.conn.cursor()
        cur.execute(
            "UPDATE incidents SET is_cleared=?, cleared_at=? WHERE id=?",
            (1 if cleared else 0, cleared_at or "", inc_id),
        )
        self.conn.commit()

    def fetch_board(self, loc_filter: Optional[int], type_filter: Optional[str],
                    start_date: Optional[str], end_date: Optional[str]) -> List[sqlite3.Row]:
        # Build query dynamically
        q = [
            "SELECT i.*, l.name as location_name,",
            "GROUP_CONCAT(CASE iu.role WHEN 'primary' THEN u.name END) as primary_units,",
            "GROUP_CONCAT(CASE iu.role WHEN 'backup' THEN u.name END) as backup_units",
            "FROM incidents i",
            "LEFT JOIN locations l ON l.id=i.location_id",
            "LEFT JOIN incident_units iu ON iu.incident_id=i.id",
            "LEFT JOIN units u ON u.id=iu.unit_id",
        ]
        where = []
        params: List[object] = []
        if loc_filter:
            where.append("i.location_id=?")
            params.append(loc_filter)
        if type_filter:
            where.append("i.type=?")
            params.append(type_filter)
        if start_date:
            where.append("date(i.reported_at) >= date(?)")
            params.append(start_date)
        if end_date:
            where.append("date(i.reported_at) <= date(?)")
            params.append(end_date)
        if where:
            q.append("WHERE " + " AND ".join(where))
        q.append("GROUP BY i.id ORDER BY i.is_cleared ASC, i.reported_at DESC")
        sql = "\n".join(q)
        return self.conn.execute(sql, params).fetchall()

    def get_incident(self, inc_id: int) -> sqlite3.Row:
        return self.conn.execute("SELECT * FROM incidents WHERE id=?", (inc_id,)).fetchone()

    def get_incident_assignments(self, inc_id: int) -> Tuple[Optional[int], List[int]]:
        primary = self.conn.execute(
            "SELECT unit_id FROM incident_units WHERE incident_id=? AND role='primary'",
            (inc_id,),
        ).fetchone()
        backups = [r[0] for r in self.conn.execute(
            "SELECT unit_id FROM incident_units WHERE incident_id=? AND role='backup'",
            (inc_id,),
        ).fetchall()]
        return (primary[0] if primary else None, backups)


# -----------------------------
# Utilities
# -----------------------------
def now_dt() -> str:
    """Current datetime in ISO format (for DB storage)."""
    return datetime.now().strftime(DT_FMT)


def now_dt_display() -> str:
    """Current datetime in MM:DD:YYYY display format (for UI)."""
    return datetime.now().strftime(DISPLAY_DT_FMT)


def fmt_dt(s: str) -> str:
    """Convert a stored ISO datetime string to MM:DD:YYYY display format."""
    if not s:
        return s
    for fmt in (DT_FMT, "%d:%m:%Y %H:%M:%S", "%m:%d:%Y %H:%M:%S", "%m-%d-%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime(DISPLAY_DT_FMT)
        except ValueError:
            pass
    return s


def parse_display_dt(s: str) -> str:
    """Convert display datetime (MM:DD:YYYY HH:MM:SS) to ISO for DB storage."""
    if not s:
        return s
    for fmt in (DISPLAY_DT_FMT, DT_FMT, "%d:%m:%Y %H:%M:%S", "%m:%d:%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime(DT_FMT)
        except ValueError:
            pass
    return s


def parse_display_date(s: str) -> str:
    """Convert display date (MM:DD:YYYY) to ISO for DB queries."""
    if not s:
        return s
    for fmt in (DISPLAY_DATE_FMT, DATE_FMT, "%d:%m:%Y", "%m:%d:%Y"):
        try:
            return datetime.strptime(s, fmt).strftime(DATE_FMT)
        except ValueError:
            pass
    return s


def ask_for_text(parent, title: str, initial: str = "") -> Optional[str]:
    """Styled text-input dialog matching the rest of the app."""
    result: list[Optional[str]] = [None]

    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title(title)
    dlg.resizable(False, False)
    set_window_icon(dlg, "input")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=f"{title}:").pack(anchor="w", pady=(0, 6))

    entry = ttk.Entry(outer, width=36)
    entry.insert(0, initial)
    entry.pack(fill="x")
    entry.focus_set()
    entry.select_range(0, "end")

    ttk.Separator(outer).pack(fill="x", pady=(16, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _ok():
        result[0] = entry.get().strip()
        dlg.destroy()

    def _cancel():
        dlg.destroy()

    ttk.Button(btn_row, text="Cancel", command=_cancel).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="OK", style="Manage.TButton", command=_ok).pack(side="right")

    dlg.bind("<Return>", lambda e: _ok())
    dlg.bind("<Escape>", lambda e: _cancel())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()

    return result[0]


# -----------------------------
# Generic manager windows (Locations, Units, Types)
# -----------------------------
class ListManager(tk.Toplevel):
    def __init__(self, master, db: DB, table: str):
        super().__init__(master)
        self.withdraw()
        set_window_icon(self, {"units": "units", "locations": "locations", "incident_types": "types"}.get(table, "units"))
        self.after(0, lambda: (apply_dark_titlebar(self), position_on_parent(self, master)))
        self.db = db
        self.table = table  # 'locations' | 'units' | 'incident_types'
        self.title(f"Manage {table.replace('_', ' ').title()}")
        self.geometry("600x460")
        self.resizable(False, False)
        self.configure(padx=12, pady=12)

        self.tree = ttk.Treeview(self, columns=("name",), show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.column("name", width=420)
        self.tree.tag_configure("available",   background="#c8e6c9", foreground="#1a1a1a")
        self.tree.tag_configure("unavailable", background="#ffcdd2", foreground="#1a1a1a")
        self.tree.grid(row=0, column=0, columnspan=6, sticky="nsew")

        btn_add   = ttk.Button(self, text="Add",    style="New.TButton",    command=self.add)
        btn_edit  = ttk.Button(self, text="Edit",   style="Manage.TButton", command=self.edit)
        btn_up    = ttk.Button(self, text="▲ Up",   command=self.move_up)
        btn_down  = ttk.Button(self, text="▼ Down", command=self.move_down)
        btn_del   = ttk.Button(self, text="Delete", style="Danger.TButton", command=self.delete)
        btn_close = ttk.Button(self, text="Close",  command=self.destroy)
        btn_add.grid(  row=1, column=0, pady=10, sticky="w")
        btn_edit.grid( row=1, column=1, pady=10, sticky="w", padx=(6, 0))
        btn_up.grid(   row=1, column=2, pady=10, sticky="w", padx=(6, 0))
        btn_down.grid( row=1, column=3, pady=10, sticky="w", padx=(6, 0))
        btn_del.grid(  row=1, column=4, pady=10, sticky="w", padx=(6, 0))
        btn_close.grid(row=1, column=5, pady=10, sticky="e", padx=(6, 0))

        self.refresh()
        if self.table == "units":
            self._poll_availability()

    def _poll_availability(self):
        """Re-colour unit rows every 5 seconds so status stays current."""
        if self.winfo_exists():
            self.refresh()
            self.after(5000, self._poll_availability)

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.table == "locations":
            for r in self.db.list_locations():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"],))
        elif self.table == "incident_types":
            for r in self.db.list_incident_types():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"],))
        else:  # units
            for r in self.db.list_units_with_availability():
                tag = "available" if r["available"] else "unavailable"
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"],), tags=(tag,))

    def add(self):
        if self.table == "units":
            name = ask_for_text(self, "Unit name")
            if not name:
                return
            self.db.add_unit(name)
        else:
            name = ask_for_text(self, f"New {self.table[:-1].replace('_', ' ')} name")
            if not name:
                return
            if self.table == "locations":
                self.db.add_location(name)
            else:
                self.db.add_incident_type(name)
        self.refresh()

    def edit(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = int(sel[0])
        if self.table == "units":
            old = self.tree.item(sel, "values")[0]
            new_name = ask_for_text(self, "Edit unit name", old)
            if new_name is None:
                return
            self.db.update_unit(iid, new_name, "")
        elif self.table == "locations":
            old = self.tree.item(sel, "values")[0]
            new = ask_for_text(self, "Rename location", old)
            if new is None:
                return
            self.db.rename_location(iid, new)
        else:
            old = self.tree.item(sel, "values")[0]
            new = ask_for_text(self, "Rename incident type", old)
            if new is None:
                return
            self.db.rename_incident_type(iid, new)
        self.refresh()

    def _swap(self, id1: int, id2: int):
        if self.table == "locations":
            self.db.swap_location_order(id1, id2)
        elif self.table == "units":
            self.db.swap_unit_order(id1, id2)
        else:
            self.db.swap_incident_type_order(id1, id2)

    def move_up(self):
        sel = self.tree.selection()
        if not sel:
            return
        prev = self.tree.prev(sel[0])
        if not prev:
            return
        iid, prev_iid = int(sel[0]), int(prev)
        self._swap(iid, prev_iid)
        self.refresh()
        self.tree.selection_set(str(iid))

    def move_down(self):
        sel = self.tree.selection()
        if not sel:
            return
        nxt = self.tree.next(sel[0])
        if not nxt:
            return
        iid, nxt_iid = int(sel[0]), int(nxt)
        self._swap(iid, nxt_iid)
        self.refresh()
        self.tree.selection_set(str(iid))

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = int(sel[0])
        name = self.tree.item(sel[0], "values")[0]

        # Block deletion if the item is referenced by existing incidents
        if self.table == "units":
            count = self.db.unit_incident_count(iid)
        elif self.table == "locations":
            count = self.db.location_incident_count(iid)
        else:
            count = self.db.incident_type_incident_count(name)

        if count:
            noun = "incident" if count == 1 else "incidents"
            dark_info(self, "Cannot Delete",
                      f'"{name}" cannot be deleted — it is attached to {count} {noun}.')
            return

        if not dark_confirm(self, "Confirm", f'Delete "{name}"?'):
            return

        if self.table == "units":
            self.db.delete_unit(iid)
        elif self.table == "locations":
            self.db.delete_location(iid)
        else:
            self.db.delete_incident_type(iid)
        self.refresh()


# -----------------------------
# Incident Form
# -----------------------------
class IncidentForm(tk.Toplevel):
    def __init__(self, master, db: DB, inc_id: Optional[int] = None, on_saved=None):
        super().__init__(master)
        self.withdraw()
        set_window_icon(self, "incident_form")
        self.after(0, lambda: (apply_dark_titlebar(self), position_on_parent(self, master)))
        self.db = db
        self.inc_id = inc_id
        self.on_saved = on_saved
        self.title("Incident Entry")
        self.geometry("820x570")
        self.transient(master)
        self.configure(padx=12, pady=12)
        self.resizable(False, False)

        # Layout grid
        for i in range(7):
            self.grid_rowconfigure(i, pad=4)
        self.grid_columnconfigure(1, weight=1)

        # Location & Type
        ttk.Label(self, text="Location").grid(row=0, column=0, sticky="w")
        self.loc_var = tk.StringVar()
        self.loc_cb = ttk.Combobox(self, textvariable=self.loc_var, values=[r["name"] for r in self.db.list_locations()], state="readonly")
        self.loc_cb.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Manage", command=lambda: self._manage("locations")).grid(row=0, column=2, sticky="w")

        ttk.Label(self, text="Incident Type").grid(row=1, column=0, sticky="w")
        self.type_var = tk.StringVar()
        self.type_cb = ttk.Combobox(self, textvariable=self.type_var, values=[r["name"] for r in self.db.list_incident_types()], state="readonly")
        self.type_cb.grid(row=1, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Manage", command=lambda: self._manage("incident_types")).grid(row=1, column=2, sticky="w")

        ttk.Label(self, text="Car #").grid(row=2, column=0, sticky="w")
        self.car_number_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.car_number_var).grid(row=2, column=1, sticky="ew", padx=(0, 6))

        # Times
        self.reported_var = tk.StringVar(value=now_dt_display())
        self.dispatched_var = tk.StringVar()
        self.arrived_var = tk.StringVar()
        self.cleared_var = tk.StringVar()

        self._time_row("Reported", self.reported_var, row=3)

        # Primary Unit (directly under Reported)
        self.unit_map = {u["name"]: u["id"] for u in self.db.list_units()}

        ttk.Label(self, text="Primary Unit").grid(row=4, column=0, sticky="w")
        self.primary_var = tk.StringVar()
        self.primary_cb = ttk.Combobox(self, textvariable=self.primary_var,
                                       values=self._available_unit_names(), state="readonly")
        self.primary_cb.grid(row=4, column=1, sticky="ew")

        self._time_row("Dispatched", self.dispatched_var, row=5)
        self._time_row("Arrived", self.arrived_var, row=6)
        self._time_row("Cleared", self.cleared_var, row=7)

        # Notes live entry
        frame_notes = ttk.LabelFrame(self, text="Add Note (time-stamped)")
        frame_notes.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame_notes.grid_columnconfigure(0, weight=1)
        self.note_text = tk.Text(frame_notes, height=4)
        self.note_text.grid(row=0, column=0, sticky="ew")
        btns = ttk.Frame(frame_notes)
        btns.grid(row=0, column=1, sticky="ns")
        ttk.Button(btns, text="Insert Time", command=lambda: self.note_text.insert(tk.END, now_dt_display()+" ")).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(btns, text="Save Note", command=self._save_quick_note).grid(row=1, column=0, padx=4)

        # Billables live entry
        frame_bill = ttk.LabelFrame(self, text="Add Billable")
        frame_bill.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame_bill.grid_columnconfigure(0, weight=1)
        self.bill_text = tk.Text(frame_bill, height=3)
        self.bill_text.grid(row=0, column=0, sticky="ew")
        btns_b = ttk.Frame(frame_bill)
        btns_b.grid(row=0, column=1, sticky="ns")
        ttk.Button(btns_b, text="Save Billable", command=self._save_quick_billable).grid(row=0, column=0, padx=4, pady=4)

        # Bottom actions
        sep = ttk.Separator(self)
        sep.grid(row=10, column=0, columnspan=3, sticky="ew", pady=6)
        self.cleared_var_bool = tk.IntVar(value=0)
        ttk.Checkbutton(self, text="Mark as Cleared", variable=self.cleared_var_bool,
                        onvalue=1, offvalue=0).grid(row=11, column=0, sticky="w")
        ttk.Button(self, text="Close", command=self.destroy).grid(row=11, column=1, sticky="e", padx=(0, 6))
        ttk.Button(self, text="Save", command=self.save).grid(row=11, column=2, sticky="e")

        # Load existing
        if self.inc_id:
            self._load_existing()

    def _manage(self, table: str):
        ListManager(self, self.db, table)
        # refresh combos after window closes
        self.after(200, self._refresh_combos)

    def _available_unit_names(self) -> list:
        """Available units for this form — excludes units on other active incidents."""
        return [u["name"] for u in self.db.list_available_units(self.inc_id)]

    def _refresh_combos(self):
        self.loc_cb["values"] = [r["name"] for r in self.db.list_locations()]
        self.type_cb["values"] = [r["name"] for r in self.db.list_incident_types()]
        self.unit_map = {u["name"]: u["id"] for u in self.db.list_units()}
        self.primary_cb["values"] = self._available_unit_names()

    def _time_row(self, label: str, var: tk.StringVar, row: int):
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w")
        e = ttk.Entry(self, textvariable=var)
        e.grid(row=row, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Now", command=lambda v=var: v.set(now_dt_display())).grid(row=row, column=2, sticky="w")

    def _load_existing(self):
        inc = self.db.get_incident(self.inc_id)
        if inc["location_id"]:
            # set by name
            loc = next((r["name"] for r in self.db.list_locations() if r["id"] == inc["location_id"]), None)
            if loc:
                self.loc_var.set(loc)
        self.type_var.set(inc["type"])
        self.car_number_var.set(inc["car_number"] or "")
        self.reported_var.set(fmt_dt(inc["reported_at"]))
        self.dispatched_var.set(fmt_dt(inc["dispatched_at"]))
        self.arrived_var.set(fmt_dt(inc["arrived_at"]))
        self.cleared_var.set(fmt_dt(inc["cleared_at"]))
        self.cleared_var_bool.set(1 if inc["is_cleared"] else 0)
        primary, _ = self.db.get_incident_assignments(self.inc_id)
        if primary:
            nm = next((u["name"] for u in self.db.list_units() if u["id"] == primary), None)
            if nm:
                self.primary_var.set(nm)

    def _flush_note(self):
        """Save any text currently in the note area. inc_id must already be set."""
        text = self.note_text.get("1.0", tk.END).strip()
        if text:
            self.db.add_note(self.inc_id, now_dt(), text)
            self.note_text.delete("1.0", tk.END)

    def _save_quick_note(self):
        text = self.note_text.get("1.0", tk.END).strip()
        if not text:
            return
        if not self.inc_id:
            # Incident not yet saved — save it first, then attach the note
            self.save()
            return
        self.db.add_note(self.inc_id, now_dt(), text)
        self.note_text.delete("1.0", tk.END)

    def _flush_billable(self):
        """Save any text currently in the billable area. inc_id must already be set."""
        text = self.bill_text.get("1.0", tk.END).strip()
        if text:
            self.db.add_billable(self.inc_id, text)
            self.bill_text.delete("1.0", tk.END)

    def _save_quick_billable(self):
        text = self.bill_text.get("1.0", tk.END).strip()
        if not text:
            return
        if not self.inc_id:
            self.save()
            return
        self.db.add_billable(self.inc_id, text)
        self.bill_text.delete("1.0", tk.END)

    def save(self):
        # Map selections
        loc_name = self.loc_var.get().strip()
        loc_id = None
        if loc_name:
            for r in self.db.list_locations():
                if r["name"] == loc_name:
                    loc_id = r["id"]
                    break
        t_name = self.type_var.get().strip() or "Other"
        reported = parse_display_dt(self.reported_var.get().strip()) or now_dt()
        disp = parse_display_dt(self.dispatched_var.get().strip())
        arr = parse_display_dt(self.arrived_var.get().strip())
        clr = parse_display_dt(self.cleared_var.get().strip())
        cleared_flag = 1 if self.cleared_var_bool.get() else 0

        # Units
        primary_name = self.primary_var.get().strip()
        primary_id = self.unit_map.get(primary_name)

        car_number = self.car_number_var.get().strip()

        if self.inc_id:
            self.db.update_incident(self.inc_id, loc_id, t_name, reported, disp, arr, clr, "", cleared_flag, primary_id, [], car_number)
        else:
            self.inc_id = self.db.create_incident(loc_id, t_name, reported, disp, arr, clr, "", cleared_flag, primary_id, [], car_number)

        # Flush any pending note text now that inc_id is guaranteed to exist
        self._flush_note()
        self._flush_billable()

        if self.on_saved:
            self.on_saved()


# -----------------------------
# Notes viewer window
# -----------------------------
class NotesWindow(tk.Toplevel):
    def __init__(self, master, db: DB, inc_id: int):
        super().__init__(master)
        self.withdraw()
        set_window_icon(self, "notes")
        self.after(0, lambda: (apply_dark_titlebar(self), position_on_parent(self, master)))
        self.db = db
        self.inc_id = inc_id
        self.title("Incident Notes")
        self.geometry("640x420")
        self.configure(padx=12, pady=12)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.entry = tk.Text(self, height=3)
        self.entry.grid(row=0, column=0, sticky="ew")
        btnbar = ttk.Frame(self)
        btnbar.grid(row=0, column=1, sticky="ns")
        ttk.Button(btnbar, text="Insert Time", command=lambda: self.entry.insert(tk.END, now_dt_display()+" ")).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(btnbar, text="Add Note", command=self.add_note).grid(row=1, column=0, padx=4)

        self.tree = ttk.Treeview(self, columns=("ts", "body"), show="headings")
        self.tree.heading("ts", text="Timestamp")
        self.tree.heading("body", text="Note")
        self.tree.column("ts", width=160)
        self.tree.column("body", width=420)
        self.tree.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8,0))
        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for n in self.db.list_notes(self.inc_id):
            self.tree.insert("", "end", values=(n["ts"], n["body"]))

    def add_note(self):
        text = self.entry.get("1.0", tk.END).strip()
        if not text:
            return
        self.db.add_note(self.inc_id, now_dt(), text)
        self.entry.delete("1.0", tk.END)
        self.refresh()


# -----------------------------
# Billables viewer window
# -----------------------------
class BillablesWindow(tk.Toplevel):
    def __init__(self, master, db: DB, inc_id: int):
        super().__init__(master)
        self.withdraw()
        set_window_icon(self, "billables")
        self.after(0, lambda: (apply_dark_titlebar(self), position_on_parent(self, master)))
        self.db = db
        self.inc_id = inc_id
        self.title("Incident Billables")
        self.geometry("640x420")
        self.configure(padx=12, pady=12)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.entry = tk.Text(self, height=3)
        self.entry.grid(row=0, column=0, sticky="ew")
        btnbar = ttk.Frame(self)
        btnbar.grid(row=0, column=1, sticky="ns")
        ttk.Button(btnbar, text="Add Billable", command=self.add_billable).grid(row=0, column=0, padx=4, pady=4)

        self.tree = ttk.Treeview(self, columns=("body",), show="headings")
        self.tree.heading("body", text="Billable")
        self.tree.column("body", width=580)
        self.tree.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for b in self.db.list_billables(self.inc_id):
            self.tree.insert("", "end", values=(b["body"],))

    def add_billable(self):
        text = self.entry.get("1.0", tk.END).strip()
        if not text:
            return
        self.db.add_billable(self.inc_id, text)
        self.entry.delete("1.0", tk.END)
        self.refresh()


# -----------------------------
# Export helpers
# -----------------------------
class Exporter:
    def __init__(self, db: DB):
        self.db = db

    def _notes_text(self, inc_id: int) -> str:
        """Return all notes for an incident as plain text, one per line."""
        notes = self.db.list_notes(inc_id)
        return "\n".join(f"{fmt_dt(n['ts'])}: {n['body']}" for n in notes)

    def _billables_text(self, inc_id: int) -> str:
        """Return all billables for an incident as plain text, one per line."""
        return "\n".join(b["body"] for b in self.db.list_billables(inc_id))

    def export_excel(self, rows: List[sqlite3.Row], path: Path, parent=None):
        headers = ["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Car #", "Unit", "Status", "Notes", "Billables"]
        try:
            import xlsxwriter  # type: ignore
        except Exception:
            # fallback to CSV
            csv_path = path.with_suffix('.csv')
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(headers)
                for r in rows:
                    notes = self._notes_text(r["id"])
                    billables = self._billables_text(r["id"])
                    w.writerow([fmt_dt(r["reported_at"]), fmt_dt(r["dispatched_at"]), fmt_dt(r["arrived_at"]), fmt_dt(r["cleared_at"]),
                                 r["type"], r["location_name"] or "", r["car_number"] or "", r["primary_units"] or "",
                                 "Cleared" if r["is_cleared"] else "Active", notes, billables])
            dark_info(parent, "Exported CSV", f"xlsxwriter not installed. Saved CSV instead to\n{csv_path}")
            return

        wb = xlsxwriter.Workbook(str(path))
        ws = wb.add_worksheet("Incidents")

        hdr_fmt   = wb.add_format({'bold': True, 'bg_color': '#D0D0D0', 'border': 1, 'valign': 'vcenter'})
        odd_fmt   = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'top'})
        even_fmt  = wb.add_format({'bg_color': '#EBEBEB', 'border': 1, 'valign': 'top'})
        odd_notes_fmt  = wb.add_format({'bg_color': '#FFFFFF', 'border': 1, 'valign': 'top', 'text_wrap': True})
        even_notes_fmt = wb.add_format({'bg_color': '#EBEBEB', 'border': 1, 'valign': 'top', 'text_wrap': True})

        col_widths = [20, 20, 20, 20, 18, 22, 10, 18, 10, 45, 35]
        for c, (h, w) in enumerate(zip(headers, col_widths)):
            ws.write(0, c, h, hdr_fmt)
            ws.set_column(c, c, w)

        for r_idx, r in enumerate(rows, start=1):
            cell_fmt  = odd_fmt        if r_idx % 2 else even_fmt
            notes_fmt = odd_notes_fmt  if r_idx % 2 else even_notes_fmt
            notes = self._notes_text(r["id"])
            billables = self._billables_text(r["id"])
            values = [fmt_dt(r["reported_at"]), fmt_dt(r["dispatched_at"]), fmt_dt(r["arrived_at"]), fmt_dt(r["cleared_at"]),
                      r["type"], r["location_name"] or "", r["car_number"] or "", r["primary_units"] or "",
                      "Cleared" if r["is_cleared"] else "Active"]
            for c, v in enumerate(values):
                ws.write(r_idx, c, v, cell_fmt)
            ws.write(r_idx, len(values), notes, notes_fmt)
            ws.write(r_idx, len(values) + 1, billables, notes_fmt)
            if notes or billables:
                line_count = max(notes.count("\n"), billables.count("\n")) + 1
                ws.set_row(r_idx, max(20, min(line_count * 15, 120)))

        ws.autofilter(0, 0, len(rows), len(headers) - 1)
        wb.close()
        dark_info(parent, "Exported", f"Saved Excel file to\n{path}")

    def export_pdf(self, rows: List[sqlite3.Row], path: Path, parent=None, title: str = "Incident Board"):
        try:
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
        except Exception:
            dark_info(parent, "Missing dependency",
                      "reportlab is not installed. Run\n  pip install reportlab\n\nAlternatively, export as Excel/CSV and print to PDF.")
            return

        doc = SimpleDocTemplate(str(path), pagesize=landscape(letter),
                                leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36,
                                title=title)
        styles = getSampleStyleSheet()

        from reportlab.lib.styles import ParagraphStyle
        cell_style = ParagraphStyle('cell', fontSize=7, leading=9, wordWrap='LTR')
        hdr_style  = ParagraphStyle('hdr',  fontSize=7, leading=9, fontName='Helvetica-Bold')

        def P(text, style=cell_style):
            return Paragraph(str(text).replace("\n", "<br/>"), style)

        headers = ["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Car #", "Unit", "Status", "Notes", "Billables"]
        data = [[P(h, hdr_style) for h in headers]]
        for r in rows:
            notes_text = self._notes_text(r["id"])
            billables_text = self._billables_text(r["id"])
            data.append([
                P(fmt_dt(r["reported_at"])), P(fmt_dt(r["dispatched_at"])), P(fmt_dt(r["arrived_at"])), P(fmt_dt(r["cleared_at"])),
                P(r["type"]), P(r["location_name"] or ""), P(r["car_number"] or ""), P(r["primary_units"] or ""),
                P("Cleared" if r["is_cleared"] else "Active"),
                P(notes_text),
                P(billables_text),
            ])

        # Landscape letter usable width ≈ 720pt (11in × 72 − 2×36 margins)
        col_widths = [70, 66, 66, 66, 56, 72, 38, 56, 40, 100, 90]  # sum = 720
        table = Table(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d0d0d0')),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#ebebeb')]),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ]))
        story = [Paragraph(title, styles['Title']), table]
        doc.build(story)
        dark_info(parent, "Exported", f"Saved PDF to\n{path}")

    def export_printable_html(self, rows: List[sqlite3.Row], path: Path, title: str = "Incident Board"):
        import html as html_lib
        html_path = path.with_suffix('.html')
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(f"""\
<!doctype html><html><head><meta charset='utf-8'>
<title>{title}</title>
<style>
body{{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px}}
.table{{border-collapse:collapse;width:100%}}
.table th,.table td{{border:1px solid #ddd;padding:8px;font-size:13px;vertical-align:top}}
.table th{{background:#d0d0d0;text-align:left}}
.table tbody tr:nth-child(even){{background:#ebebeb}}
.table tbody tr:nth-child(odd){{background:#ffffff}}
.notes{{font-size:12px;color:#444;white-space:pre-line}}
</style></head><body>
<h2>{title}</h2>
<table class='table'>
<thead><tr><th>Reported</th><th>Dispatched</th><th>Arrived</th><th>Cleared</th><th>Type</th><th>Location</th><th>Car #</th><th>Unit</th><th>Status</th><th>Notes</th><th>Billables</th></tr></thead>
<tbody>
""")
            for r in rows:
                notes_text = self._notes_text(r["id"])
                billables_text = self._billables_text(r["id"])
                notes_html = f"<span class='notes'>{html_lib.escape(notes_text)}</span>" if notes_text else ""
                billables_html = f"<span class='notes'>{html_lib.escape(billables_text)}</span>" if billables_text else ""
                f.write(
                    f"<tr>"
                    f"<td>{fmt_dt(r['reported_at'])}</td>"
                    f"<td>{fmt_dt(r['dispatched_at'])}</td>"
                    f"<td>{fmt_dt(r['arrived_at'])}</td>"
                    f"<td>{fmt_dt(r['cleared_at'])}</td>"
                    f"<td>{html_lib.escape(r['type'])}</td>"
                    f"<td>{html_lib.escape(r['location_name'] or '')}</td>"
                    f"<td>{html_lib.escape(r['car_number'] or '')}</td>"
                    f"<td>{html_lib.escape(r['primary_units'] or '')}</td>"
                    f"<td>{'Cleared' if r['is_cleared'] else 'Active'}</td>"
                    f"<td>{notes_html}</td>"
                    f"<td>{billables_html}</td>"
                    f"</tr>\n"
                )
            f.write("</tbody></table></body></html>\n")
        webbrowser.open(html_path.as_uri())


# -----------------------------
# Main application
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        set_icon(self)
        self.after(0, lambda: apply_dark_titlebar(self))
        self.db = DB()
        self.title(APP_TITLE)
        self.geometry("1400x740")
        self.minsize(1050, 620)

        try:
            self.style = ttk.Style(self)
            self.style.theme_use("clam")
            self.style.configure("TButton", padding=6)
            self.style.configure("Treeview", rowheight=26)

            # New Incident — vivid green
            self.style.configure("New.TButton", padding=6, background="#2e7d32",
                                 foreground="white", font=("TkDefaultFont", 9, "bold"))
            self.style.map("New.TButton",
                           background=[("active", "#388e3c"), ("pressed", "#1b5e20")],
                           foreground=[("active", "white"), ("pressed", "white")])

            # Manage Units — vivid blue
            self.style.configure("Manage.TButton", padding=6, background="#1565c0",
                                 foreground="white", font=("TkDefaultFont", 9, "bold"))
            self.style.map("Manage.TButton",
                           background=[("active", "#1976d2"), ("pressed", "#0d47a1")],
                           foreground=[("active", "white"), ("pressed", "white")])

            # Delete / danger — vivid red
            self.style.configure("Danger.TButton", padding=6, background="#c62828",
                                 foreground="white", font=("TkDefaultFont", 9, "bold"))
            self.style.map("Danger.TButton",
                           background=[("active", "#e53935"), ("pressed", "#b71c1c")],
                           foreground=[("active", "white"), ("pressed", "white")])
        except Exception:
            pass

        self.exporter = Exporter(self.db)

        self._build_menu()
        self._build_filters()
        self._build_board()
        self.refresh_board()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._session_date = date.today()
        self.after(60000, self._check_date_rollover)

        # Auto-check for updates (only when installed, not in dev)
        if getattr(sys, "frozen", False):
            self.after(1500, self._start_update_check)

    # ----- UI builders
    def _build_menu(self):
        menubar = tk.Menu(self)
        # File
        m_file = tk.Menu(menubar, tearoff=False)
        m_file.add_command(label="New Incident", command=self.new_incident, accelerator="Ctrl+N")
        m_file.add_separator()
        m_file.add_command(label="Export to Excel", command=self.export_excel)
        m_file.add_command(label="Export to PDF", command=self.export_pdf)
        m_file.add_command(label="Open printable HTML", command=self.open_printable_html)
        m_file.add_separator()
        m_file.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=m_file)

        # Manage
        m_mng = tk.Menu(menubar, tearoff=False)
        m_mng.add_command(label="Locations", command=lambda: ListManager(self, self.db, "locations"))
        m_mng.add_command(label="Units", command=lambda: ListManager(self, self.db, "units"))
        m_mng.add_command(label="Incident Types", command=lambda: ListManager(self, self.db, "incident_types"))
        m_mng.add_separator()
        m_mng.add_command(label="Export Lists...", command=self.export_lists)
        m_mng.add_command(label="Import Lists...", command=self.import_lists)
        menubar.add_cascade(label="Manage", menu=m_mng)

        # Help
        m_help = tk.Menu(menubar, tearoff=False)
        m_help.add_command(label="User Guide", command=self._show_tutorial)
        m_help.add_command(label="Check for Updates", command=self._manual_update_check)
        m_help.add_separator()
        m_help.add_command(label="About", command=lambda: dark_info(self, "About", f"Incident Desk {APP_VERSION}\nSimple offline incident board for race control\nBuilt with Python (Tkinter) + SQLite by Ryder Smith 2025-2026"))
        menubar.add_cascade(label="Help", menu=m_help)
        self.config(menu=menubar)

        # keybinds
        self.bind_all("<Control-n>", lambda e: self.new_incident())
        self.bind_all("<Delete>", lambda e: self.delete_selected())

    def _build_filters(self):
        bar = ttk.Frame(self, padding=(12, 10))
        bar.pack(fill="x")

        # Date range
        ttk.Label(bar, text="Date from").pack(side="left")
        self.from_picker = DateEntry(bar, width=12, date_pattern="mm-dd-yyyy",
                                     selectmode="day", firstweekday="sunday")
        self.from_picker.set_date(date.today())
        self.from_picker.pack(side="left", padx=(4, 10))
        ttk.Label(bar, text="to").pack(side="left")
        self.to_picker = DateEntry(bar, width=12, date_pattern="mm-dd-yyyy",
                                   selectmode="day", firstweekday="sunday")
        self.to_picker.set_date(date.today())
        self.to_picker.pack(side="left", padx=(4, 12))

        # Location filter
        ttk.Label(bar, text="Location").pack(side="left")
        self.loc_filter_var = tk.StringVar(value="All")
        locs = ["All"] + [r["name"] for r in self.db.list_locations()]
        self.loc_filter_cb = ttk.Combobox(bar, textvariable=self.loc_filter_var, values=locs, width=20, state="readonly")
        self.loc_filter_cb.pack(side="left", padx=(4,12))

        # Type filter
        ttk.Label(bar, text="Type").pack(side="left")
        self.type_filter_var = tk.StringVar(value="All")
        types = ["All"] + [r["name"] for r in self.db.list_incident_types()]
        self.type_filter_cb = ttk.Combobox(bar, textvariable=self.type_filter_var, values=types, width=16, state="readonly")
        self.type_filter_cb.pack(side="left", padx=(4,12))

        ttk.Button(bar, text="Apply", command=self.refresh_board).pack(side="left")
        ttk.Button(bar, text="Reset", command=self.reset_filters).pack(side="left", padx=(6,0))

        # Spacer
        ttk.Label(bar, text="").pack(side="left", expand=True)

        ttk.Button(bar, text="New Incident", style="New.TButton", command=self.new_incident).pack(side="right")

    def _build_board(self):
        frame = ttk.Frame(self, padding=(12, 0))
        frame.pack(fill="both", expand=True)

        cols = ("reported", "dispatched", "arrived", "cleared", "type", "location", "car", "units", "status")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("reported", text="Rec'd")
        self.tree.heading("dispatched", text="Disp")
        self.tree.heading("arrived", text="Arrv'd")
        self.tree.heading("cleared", text="Done")
        self.tree.heading("type", text="Type")
        self.tree.heading("location", text="Location")
        self.tree.heading("car", text="Car #")
        self.tree.heading("units", text="Unit(s)")
        self.tree.heading("status", text="Status")

        self.tree.column("reported",   width=140, anchor="center")
        self.tree.column("dispatched", width=120, anchor="center")
        self.tree.column("arrived",    width=120, anchor="center")
        self.tree.column("cleared",    width=120, anchor="center")
        self.tree.column("type",       width=120, anchor="center")
        self.tree.column("location",   width=160, anchor="center")
        self.tree.column("car",        width=70,  anchor="center")
        self.tree.column("units",      width=140, anchor="center")
        self.tree.column("status",     width=90,  anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")

        # Row colors via tags
        self.tree.tag_configure('cleared', background="#c8e6c9", foreground="#1a1a1a")
        self.tree.tag_configure('active',  background="#ffcdd2", foreground="#1a1a1a")

        # right-side actions
        side = ttk.Frame(frame)
        side.pack(side="right", fill="y", padx=(8,0))
        ttk.Button(side, text="Manage Units", style="Manage.TButton", command=lambda: ListManager(self, self.db, "units")).pack(fill="x", pady=6)
        ttk.Separator(side).pack(fill="x", pady=4)
        ttk.Button(side, text="View/Edit", command=self.edit_selected).pack(fill="x", pady=6)
        ttk.Button(side, text="Notes", command=self.open_notes).pack(fill="x", pady=6)
        ttk.Button(side, text="Billables", command=self.open_billables).pack(fill="x", pady=6)
        ttk.Button(side, text="Toggle Cleared", command=self.mark_cleared).pack(fill="x", pady=6)
        ttk.Button(side, text="Delete", style="Danger.TButton", command=self.delete_selected).pack(fill="x", pady=6)
        ttk.Separator(side).pack(fill="x", pady=8)
        ttk.Button(side, text="Export Excel", command=self.export_excel).pack(fill="x", pady=6)
        ttk.Button(side, text="Export PDF", command=self.export_pdf).pack(fill="x", pady=6)
        ttk.Button(side, text="Printable HTML", command=self.open_printable_html).pack(fill="x", pady=6)

        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

    # ----- Actions
    def _filters(self):
        # map location name to id
        loc_name = self.loc_filter_var.get()
        loc_id = None
        if loc_name and loc_name != "All":
            for r in self.db.list_locations():
                if r["name"] == loc_name:
                    loc_id = r["id"]
                    break
        tname = self.type_filter_var.get()
        if tname == "All":
            tname = None
        start = self.from_picker.get_date().strftime(DATE_FMT)
        end = self.to_picker.get_date().strftime(DATE_FMT)
        return loc_id, tname, start, end

    def refresh_board(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        loc_id, tname, start, end = self._filters()
        rows = self.db.fetch_board(loc_id, tname, start, end)
        for r in rows:
            units_label = ", ".join([x for x in [r["primary_units"], r["backup_units"]] if x])
            tag = 'cleared' if r["is_cleared"] else 'active'
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                fmt_dt(r["reported_at"]), fmt_dt(r["dispatched_at"]), fmt_dt(r["arrived_at"]), fmt_dt(r["cleared_at"]),
                r["type"], r["location_name"] or "", r["car_number"] or "", units_label, "Cleared" if r["is_cleared"] else "Active"
            ), tags=(tag,))

    def reset_filters(self):
        self.from_picker.set_date(date.today())
        self.to_picker.set_date(date.today())
        self.loc_filter_var.set("All")
        self.type_filter_var.set("All")
        self.refresh_board()

    def new_incident(self):
        IncidentForm(self, self.db, on_saved=self.refresh_board)

    def get_selected_incident_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            dark_info(self, "Select", "Please select an incident from the board.")
            return None
        return int(sel[0])

    def edit_selected(self):
        iid = self.get_selected_incident_id()
        if iid:
            IncidentForm(self, self.db, inc_id=iid, on_saved=self.refresh_board)

    def open_notes(self):
        iid = self.get_selected_incident_id()
        if iid:
            NotesWindow(self, self.db, inc_id=iid)

    def open_billables(self):
        iid = self.get_selected_incident_id()
        if iid:
            BillablesWindow(self, self.db, inc_id=iid)

    def mark_cleared(self):
        iid = self.get_selected_incident_id()
        if not iid:
            return
        inc = self.db.get_incident(iid)
        currently_cleared = bool(inc["is_cleared"])
        self.db.set_cleared(iid, not currently_cleared, now_dt() if not currently_cleared else "")
        self.refresh_board()

    def delete_selected(self):
        iid = self.get_selected_incident_id()
        if not iid:
            return
        if dark_confirm(self, "Confirm", "Delete the selected incident?"):
            self.db.delete_incident(iid)
            self.refresh_board()

    def _current_rows_for_export(self) -> List[sqlite3.Row]:
        loc_id, tname, start, end = self._filters()
        return self.db.fetch_board(loc_id, tname, start, end)

    def export_excel(self):
        rows = self._current_rows_for_export()
        if not rows:
            dark_info(self, "Nothing to export", "No incidents match the current filters.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", ".xlsx"), ("CSV", ".csv")])
        if not path:
            return
        self.exporter.export_excel(rows, Path(path), self)

    def _export_title(self) -> str:
        d1 = self.from_picker.get_date().strftime(DISPLAY_DATE_FMT)
        d2 = self.to_picker.get_date().strftime(DISPLAY_DATE_FMT)
        return d1 if d1 == d2 else f"{d1} \u2013 {d2}"

    def export_pdf(self):
        rows = self._current_rows_for_export()
        if not rows:
            dark_info(self, "Nothing to export", "No incidents match the current filters.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", ".pdf")])
        if not path:
            return
        self.exporter.export_pdf(rows, Path(path), self, title=self._export_title())

    def open_printable_html(self):
        rows = self._current_rows_for_export()
        if not rows:
            dark_info(self, "Nothing to show", "No incidents match the current filters.")
            return
        out = DB_DIR / "incident_board.html"
        self.exporter.export_printable_html(rows, out, title=self._export_title())


    def export_lists(self):
        path = filedialog.asksaveasfilename(
            title="Export Lists",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialfile="incidentdesk_lists.json",
        )
        if not path:
            return
        data = {
            "locations": [r["name"] for r in self.db.list_locations()],
            "incident_types": [r["name"] for r in self.db.list_incident_types()],
            "units": [r["name"] for r in self.db.list_units()],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        dark_info(self, "Exported", f"Lists saved to\n{path}")

    def import_lists(self):
        path = filedialog.askopenfilename(
            title="Import Lists",
            filetypes=[("JSON", "*.json")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            dark_info(self, "Import Failed", f"Could not read file:\n{e}")
            return

        added = {"locations": 0, "incident_types": 0, "units": 0}
        for name in data.get("locations", []):
            if isinstance(name, str) and name.strip():
                before = len(self.db.list_locations())
                self.db.add_location(name.strip())
                if len(self.db.list_locations()) > before:
                    added["locations"] += 1
        for name in data.get("incident_types", []):
            if isinstance(name, str) and name.strip():
                before = len(self.db.list_incident_types())
                self.db.add_incident_type(name.strip())
                if len(self.db.list_incident_types()) > before:
                    added["incident_types"] += 1
        for u in data.get("units", []):
            name = (u if isinstance(u, str) else u.get("name", "")).strip()
            if name:
                before = len(self.db.list_units())
                self.db.add_unit(name)
                if len(self.db.list_units()) > before:
                    added["units"] += 1

        dark_info(self, "Import Complete",
                  f"Added:\n  {added['locations']} location(s)\n  {added['incident_types']} incident type(s)\n  {added['units']} unit(s)\n\nExisting entries were not duplicated.")

    def _check_date_rollover(self):
        if not self.winfo_exists():
            return
        today = date.today()
        if today != self._session_date:
            self._session_date = today
            self.reset_filters()
        self.after(60000, self._check_date_rollover)

    # ----- Update checking
    def _start_update_check(self, show_no_update: bool = False):
        """Kick off a GitHub version check on a background thread."""
        def worker():
            result = check_for_update()
            self.after(0, lambda: self._handle_update_result(result, show_no_update))
        threading.Thread(target=worker, daemon=True).start()

    def _handle_update_result(self, result, show_no_update: bool):
        if not self.winfo_exists():
            return
        if result is None:
            if show_no_update:
                dark_info(self, "No Updates",
                          f"You are running the latest version (Incident Desk {APP_VERSION}).")
            return
        latest_tag, installer_url = result
        if update_prompt(self, APP_VERSION, latest_tag):
            self._download_and_install(installer_url)

    def _manual_update_check(self):
        """Triggered from Help > Check for Updates."""
        self._start_update_check(show_no_update=True)

    def _download_and_install(self, url: str):
        """Show a modal 'downloading' dialog, fetch the installer, then run it."""
        dlg = tk.Toplevel(self)
        dlg.withdraw()
        dlg.title("Updating")
        dlg.resizable(False, False)
        set_window_icon(dlg, "export")
        dlg.after(0, lambda: apply_dark_titlebar(dlg))
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)  # disable close during download

        outer = ttk.Frame(dlg, padding=24)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Downloading update...\nThe app will restart when finished.",
                  wraplength=320, justify="left").pack(anchor="w")

        dlg.transient(self)
        position_on_parent(dlg, self)
        dlg.grab_set()

        dest = Path(tempfile.gettempdir()) / "IncidentDesk-Setup.exe"

        def worker():
            ok = download_file(url, dest)
            self.after(0, lambda: _finish(ok))

        def _finish(ok: bool):
            dlg.destroy()
            if not ok:
                dark_info(self, "Update Failed",
                          "Could not download the update. Please check your connection and try again.")
                return
            # Launch installer detached so it survives this process exiting,
            # then shut down so the installer can replace the exe.
            DETACHED_PROCESS = 0x00000008
            try:
                subprocess.Popen([str(dest), "/SILENT"], creationflags=DETACHED_PROCESS)
            except OSError as ex:
                dark_info(self, "Update Failed", f"Could not launch installer:\n{ex}")
                return
            self.destroy()

        threading.Thread(target=worker, daemon=True).start()

    def _on_close(self):
        dlg = tk.Toplevel(self)
        dlg.withdraw()
        dlg.title("Export Before Closing?")
        dlg.resizable(False, False)
        dlg.grab_set()
        set_window_icon(dlg, "export")
        dlg.after(0, lambda: apply_dark_titlebar(dlg))

        outer = ttk.Frame(dlg, padding=20)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text=(
                "Would you like to export today's board before closing?\n\n"
                "All incident data is saved — when reopened, today's date and\n"
                "incidents will be included and displayed automatically. Use the\n"
                "Date From / To pickers to retrieve incidents from any past\n"
                "date or time period."
            ),
            wraplength=380, justify="left"
        ).pack(anchor="w")

        ttk.Separator(outer).pack(fill="x", pady=(16, 12))

        export_row = ttk.Frame(outer)
        export_row.pack(fill="x", pady=(0, 10))
        ttk.Label(export_row, text="Export as:").pack(side="left", padx=(0, 8))
        ttk.Button(export_row, text="Excel",          command=lambda: self.export_excel()).pack(side="left", padx=(0, 6))
        ttk.Button(export_row, text="PDF",            command=lambda: self.export_pdf()).pack(side="left", padx=(0, 6))
        ttk.Button(export_row, text="Printable HTML", command=lambda: self.open_printable_html()).pack(side="left")

        ttk.Separator(outer).pack(fill="x", pady=(4, 12))

        action_row = ttk.Frame(outer)
        action_row.pack(fill="x")
        ttk.Button(action_row, text="Cancel",           command=dlg.destroy).pack(side="left")
        ttk.Button(action_row, text="Exit Without Exporting", style="Danger.TButton",
                   command=lambda: [dlg.destroy(), self.destroy()]).pack(side="right")

        dlg.transient(self)
        position_on_parent(dlg, self)

    def _show_tutorial(self):
        win = tk.Toplevel(self)
        win.withdraw()
        win.title("User Guide")
        win.geometry("860x780")
        win.minsize(700, 600)
        set_window_icon(win, "guide")
        win.after(0, lambda: apply_dark_titlebar(win))
        win.transient(self)

        outer = ttk.Frame(win, padding=16)
        outer.pack(fill="both", expand=True)

        txt = tk.Text(
            outer, wrap="word", relief="flat", borderwidth=0,
            font=("Segoe UI", 10), padx=8, pady=8,
            state="normal", cursor="arrow",
        )
        sb = ttk.Scrollbar(outer, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        txt.tag_configure("h1", font=("Segoe UI", 13, "bold"), spacing3=4)
        txt.tag_configure("h2", font=("Segoe UI", 10, "bold"), spacing3=2)
        txt.tag_configure("body", font=("Segoe UI", 10), spacing3=2)
        txt.tag_configure("key",  font=("Consolas", 9), relief="groove",
                          borderwidth=1, lmargin1=4, lmargin2=4)

        def h1(t): txt.insert("end", t + "\n", "h1")
        def h2(t): txt.insert("end", t + "\n", "h2")
        def body(t): txt.insert("end", t + "\n", "body")

        h1("Incident Desk — User Guide")
        body("")

        h2("Incident Board")
        body("The main board displays all incidents for the selected date range. Each row is colour-coded:")
        body("  • Red background  — Active incident (not yet cleared)")
        body("  • Green background — Cleared incident")
        body("")

        h2("Creating an Incident")
        body("Click New Incident (top-right of the filter bar) or press Ctrl+N.")
        body("Fill in the Type, Location, Unit(s), and timestamps, then click Save.")
        body("")

        h2("Viewing & Editing")
        body("Double-click any row, or select a row and click View/Edit on the right panel.")
        body("Make your changes in the form and click Save to update.")
        body("")

        h2("Notes")
        body("Select an incident and click Notes to open a dedicated per-incident notes log.")
        body("Each entry is timestamped automatically.")
        body("")

        h2("Clearing an Incident")
        body("Select a row and click Toggle Cleared to mark it cleared (green) or revert it to active (red).")
        body("The cleared timestamp is recorded automatically.")
        body("")

        h2("Deleting an Incident")
        body("Select a row and click Delete. You will be asked to confirm before the record is removed.")
        body("")

        h2("Filtering the Board")
        body("Use the filter bar at the top to narrow the board by date range, location, or incident type.")
        body("  • Date From / To — click the calendar icon to pick dates.")
        body("  • Location / Type — select from the dropdowns.")
        body("  • Apply — refresh the board with the current filters.")
        body("  • Reset — restore today's date and clear all filters.")
        body("")

        h2("Data Saving & Reopening")
        body("All incident data is saved automatically to a local database — nothing is lost when the app is closed.")
        body("When reopened, the board defaults to today's date and will display any incidents already")
        body("logged for today alongside new ones.")
        body("")
        body("To review incidents from a previous day or time period, use the Date From / To pickers")
        body("in the filter bar to select any past date or range, then click Apply.")
        body("")

        h2("Automatic Date Rollover")
        body("If the app is left open past midnight, the board will automatically switch to the new date")
        body("and refresh — no restart required.")
        body("")
        body("When closing, a reminder dialog will appear offering the option to export today's board")
        body("before exiting. All data remains saved regardless of whether you export.")
        body("")

        h2("Managing Units, Locations & Types")
        body("Click Manage Units (right panel) or use the Manage menu to open list managers for:")
        body("  • Units — vehicles or personnel that can be assigned to incidents.")
        body("  • Locations — named locations selectable on the incident form.")
        body("  • Incident Types — categories used to classify incidents.")
        body("In each manager you can Add, Edit, reorder with ▲ Up / ▼ Down, or Delete entries.")
        body("Items cannot be deleted while attached to existing incidents.")
        body("")

        h2("Exporting")
        body("Use the right panel or File menu to export the currently filtered board:")
        body("  • Export Excel — saves an .xlsx spreadsheet.")
        body("  • Export PDF — saves a formatted .pdf report.")
        body("  • Printable HTML — opens a print-ready page in your browser.")
        body("")

        h2("Import / Export Lists")
        body("Use Manage › Export Lists to save your Units, Locations, and Types to a JSON file.")
        body("Use Manage › Import Lists to load them back, for example when setting up a new device.")
        body("")

        h2("Keyboard Shortcuts")
        body("  Ctrl+N — New Incident")
        body("")

        txt.configure(state="disabled")

        ttk.Separator(outer).pack(fill="x", pady=(12, 8), side="bottom")
        ttk.Button(outer, text="Close", style="Manage.TButton",
                   command=win.destroy).pack(side="bottom", anchor="e")
        position_on_parent(win, self)


if __name__ == "__main__":
    App().mainloop()
