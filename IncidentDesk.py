"""
Incident Desk - simple offline incident logging app for racetracks
-----------------------------------------------------------------
• Offline-first: SQLite file stored next to the script (trackincident.db)
• Modern-ish Tkinter/ttk UI; color‑coded incident board
• Add/Edit incidents with units (primary + backups) and time stamps
• Filter/search by date, location, and incident type
• Editable pick‑lists for Locations, Units, and Incident Types
• Notes per incident with automatic timestamps
• Export board to Excel (xlsx) or CSV; export/print to PDF (via reportlab if installed)

Optional exports (install if you want PDF/XLSX):
py -m pip install reportlab xlsxwriter

"""
from __future__ import annotations
import os
import csv
import sqlite3
import webbrowser
from datetime import datetime, date
from pathlib import Path
from typing import List, Optional, Tuple
import sys

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog



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
DT_FMT = "%Y-%m-%d %H:%M"  # 24h local time
DATE_FMT = "%Y-%m-%d"

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
            """
        )
        self.conn.commit()

    # ---- CRUD helpers
    def list_locations(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM locations ORDER BY name").fetchall()

    def add_location(self, name: str):
        self.conn.execute("INSERT OR IGNORE INTO locations(name) VALUES(?)", (name.strip(),))
        self.conn.commit()

    def rename_location(self, loc_id: int, new_name: str):
        self.conn.execute("UPDATE locations SET name=? WHERE id=?", (new_name.strip(), loc_id))
        self.conn.commit()

    def delete_location(self, loc_id: int):
        self.conn.execute("DELETE FROM locations WHERE id=?", (loc_id,))
        self.conn.commit()

    def list_units(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM units ORDER BY name").fetchall()

    def add_unit(self, name: str, category: str = ""):
        self.conn.execute("INSERT OR IGNORE INTO units(name, category) VALUES(?,?)", (name.strip(), category.strip()))
        self.conn.commit()

    def update_unit(self, unit_id: int, name: str, category: str):
        self.conn.execute("UPDATE units SET name=?, category=? WHERE id=?", (name.strip(), category.strip(), unit_id))
        self.conn.commit()

    def delete_unit(self, unit_id: int):
        self.conn.execute("DELETE FROM units WHERE id=?", (unit_id,))
        self.conn.commit()

    def list_incident_types(self) -> List[sqlite3.Row]:
        return self.conn.execute("SELECT * FROM incident_types ORDER BY name").fetchall()

    def add_incident_type(self, name: str):
        self.conn.execute("INSERT OR IGNORE INTO incident_types(name) VALUES(?)", (name.strip(),))
        self.conn.commit()

    def rename_incident_type(self, type_id: int, new_name: str):
        self.conn.execute("UPDATE incident_types SET name=? WHERE id=?", (new_name.strip(), type_id))
        self.conn.commit()

    def delete_incident_type(self, type_id: int):
        self.conn.execute("DELETE FROM incident_types WHERE id=?", (type_id,))
        self.conn.commit()

    def create_incident(self, location_id: Optional[int], type_name: str, reported_at: str,
                         dispatched_at: str = "", arrived_at: str = "", cleared_at: str = "",
                         disposition: str = "", is_cleared: int = 0,
                         primary_unit_id: Optional[int] = None, backup_unit_ids: Optional[List[int]] = None) -> int:
        cur = self.conn.cursor()
        cur.execute(
            """
            INSERT INTO incidents(location_id, type, reported_at, dispatched_at, arrived_at, cleared_at, disposition, is_cleared)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, is_cleared),
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
                         primary_unit_id: Optional[int], backup_unit_ids: List[int]):
        cur = self.conn.cursor()
        cur.execute(
            "UPDATE incidents SET location_id=?, type=?, reported_at=?, dispatched_at=?, arrived_at=?, cleared_at=?, disposition=?, is_cleared=? WHERE id=?",
            (location_id, type_name, reported_at, dispatched_at, arrived_at, cleared_at, disposition, is_cleared, inc_id),
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
    return datetime.now().strftime(DT_FMT)


def ask_for_text(parent, title: str, initial: str = "") -> Optional[str]:
    return simpledialog.askstring(title, f"{title}:", initialvalue=initial, parent=parent)


# -----------------------------
# Generic manager windows (Locations, Units, Types)
# -----------------------------
class ListManager(tk.Toplevel):
    def __init__(self, master, db: DB, table: str):
        super().__init__(master)
        self.db = db
        self.table = table  # 'locations' | 'units' | 'incident_types'
        self.title(f"Manage {table.replace('_', ' ').title()}")
        self.geometry("470x460")
        self.resizable(False, False)
        self.configure(padx=12, pady=12)

        self.tree = ttk.Treeview(self, columns=("name", "category"), show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.column("name", width=280)
        if table == "units":
            self.tree.heading("category", text="Category")
            self.tree.column("category", width=160)
        else:
            self.tree.heading("category", text="")
            self.tree.column("category", width=0)
        self.tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

        btn_add = ttk.Button(self, text="Add", command=self.add)
        btn_edit = ttk.Button(self, text="Edit", command=self.edit)
        btn_del = ttk.Button(self, text="Delete", command=self.delete)
        btn_close = ttk.Button(self, text="Close", command=self.destroy)
        btn_add.grid(row=1, column=0, pady=10, sticky="w")
        btn_edit.grid(row=1, column=1, pady=10, sticky="w")
        btn_del.grid(row=1, column=2, pady=10, sticky="w")
        btn_close.grid(row=1, column=3, pady=10, sticky="e")

        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.table == "locations":
            for r in self.db.list_locations():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], ""))
        elif self.table == "incident_types":
            for r in self.db.list_incident_types():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], ""))
        else:  # units
            for r in self.db.list_units():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], r["category"]))

    def add(self):
        if self.table == "units":
            name = ask_for_text(self, "Unit name")
            if not name:
                return
            cat = ask_for_text(self, "Category (e.g. Fire, Safety Truck, Medical)", "") or ""
            self.db.add_unit(name, cat)
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
            old = self.tree.item(sel, "values")
            new_name = ask_for_text(self, "Edit unit name", old[0])
            if new_name is None:
                return
            new_cat = ask_for_text(self, "Edit category", old[1])
            if new_cat is None:
                return
            self.db.update_unit(iid, new_name, new_cat)
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

    def delete(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = int(sel[0])
        if not messagebox.askyesno("Confirm", "Delete the selected item?"):
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
        self.db = db
        self.inc_id = inc_id
        self.on_saved = on_saved
        self.title("Incident Entry")
        self.geometry("820x560")
        self.transient(master)
        self.configure(padx=12, pady=12)
        self.resizable(False, False)

        # Layout grid
        for i in range(6):
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

        # Times
        self.reported_var = tk.StringVar(value=now_dt())
        self.dispatched_var = tk.StringVar()
        self.arrived_var = tk.StringVar()
        self.cleared_var = tk.StringVar()

        self._time_row("Reported", self.reported_var, row=2)
        self._time_row("Dispatched", self.dispatched_var, row=3)
        self._time_row("Arrived", self.arrived_var, row=4)
        self._time_row("Cleared", self.cleared_var, row=5)

        # Units: primary + backups
        units = self.db.list_units()
        unit_names = [u["name"] for u in units]
        unit_ids = [u["id"] for u in units]
        self.unit_map = dict(zip(unit_names, unit_ids))

        ttk.Label(self, text="Primary Unit").grid(row=6, column=0, sticky="w", pady=(6,0))
        self.primary_var = tk.StringVar()
        self.primary_cb = ttk.Combobox(self, textvariable=self.primary_var, values=unit_names, state="readonly")
        self.primary_cb.grid(row=6, column=1, sticky="ew")

        ttk.Label(self, text="Backup Units").grid(row=7, column=0, sticky="nw")
        self.backup_lb = tk.Listbox(self, selectmode=tk.MULTIPLE, height=6, exportselection=False)
        for n in unit_names:
            self.backup_lb.insert(tk.END, n)
        self.backup_lb.grid(row=7, column=1, sticky="ew")

        # Disposition / notes helper
        ttk.Label(self, text="Disposition").grid(row=8, column=0, sticky="w")
        self.disp_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.disp_var).grid(row=8, column=1, columnspan=2, sticky="ew")

        # Notes live entry
        frame_notes = ttk.LabelFrame(self, text="Add Note (time-stamped)")
        frame_notes.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame_notes.grid_columnconfigure(0, weight=1)
        self.note_text = tk.Text(frame_notes, height=4)
        self.note_text.grid(row=0, column=0, sticky="ew")
        btns = ttk.Frame(frame_notes)
        btns.grid(row=0, column=1, sticky="ns")
        ttk.Button(btns, text="Insert Time", command=lambda: self.note_text.insert(tk.END, now_dt()+" ")).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(btns, text="Save Note", command=self._save_quick_note).grid(row=1, column=0, padx=4)

        # Bottom actions
        sep = ttk.Separator(self)
        sep.grid(row=10, column=0, columnspan=3, sticky="ew", pady=6)
        self.cleared_var_bool = tk.BooleanVar(value=False)
        ttk.Checkbutton(self, text="Mark as Cleared", variable=self.cleared_var_bool).grid(row=11, column=0, sticky="w")
        ttk.Button(self, text="Save", command=self.save).grid(row=11, column=2, sticky="e")

        # Load existing
        if self.inc_id:
            self._load_existing()

    def _manage(self, table: str):
        ListManager(self, self.db, table)
        # refresh combos after window closes
        self.after(200, self._refresh_combos)

    def _refresh_combos(self):
        self.loc_cb["values"] = [r["name"] for r in self.db.list_locations()]
        self.type_cb["values"] = [r["name"] for r in self.db.list_incident_types()]
        units = self.db.list_units()
        names = [u["name"] for u in units]
        self.unit_map = {u["name"]: u["id"] for u in units}
        self.primary_cb["values"] = names
        self.backup_lb.delete(0, tk.END)
        for n in names:
            self.backup_lb.insert(tk.END, n)

    def _time_row(self, label: str, var: tk.StringVar, row: int):
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w")
        e = ttk.Entry(self, textvariable=var)
        e.grid(row=row, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Now", command=lambda v=var: v.set(now_dt())).grid(row=row, column=2, sticky="w")

    def _load_existing(self):
        inc = self.db.get_incident(self.inc_id)
        if inc["location_id"]:
            # set by name
            loc = next((r["name"] for r in self.db.list_locations() if r["id"] == inc["location_id"]), None)
            if loc:
                self.loc_var.set(loc)
        self.type_var.set(inc["type"])
        self.reported_var.set(inc["reported_at"])
        self.dispatched_var.set(inc["dispatched_at"])
        self.arrived_var.set(inc["arrived_at"])
        self.cleared_var.set(inc["cleared_at"])
        self.disp_var.set(inc["disposition"] or "")
        self.cleared_var_bool.set(bool(inc["is_cleared"]))
        primary, backups = self.db.get_incident_assignments(self.inc_id)
        if primary:
            nm = next((u["name"] for u in self.db.list_units() if u["id"] == primary), None)
            if nm:
                self.primary_var.set(nm)
        if backups:
            names = [u["name"] for u in self.db.list_units() if u["id"] in backups]
            # select in listbox
            for i in range(self.backup_lb.size()):
                if self.backup_lb.get(i) in names:
                    self.backup_lb.selection_set(i)

    def _save_quick_note(self):
        text = self.note_text.get("1.0", tk.END).strip()
        if not text:
            return
        if not self.inc_id:
            messagebox.showinfo("Save incident first", "Please save the incident before adding notes.")
            return
        self.db.add_note(self.inc_id, now_dt(), text)
        self.note_text.delete("1.0", tk.END)
        messagebox.showinfo("Saved", "Note added to incident.")

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
        reported = self.reported_var.get().strip() or now_dt()
        disp = self.dispatched_var.get().strip()
        arr = self.arrived_var.get().strip()
        clr = self.cleared_var.get().strip()
        dispn = self.disp_var.get().strip()
        cleared_flag = 1 if self.cleared_var_bool.get() or bool(clr) else 0

        # Units
        primary_name = self.primary_var.get().strip()
        primary_id = self.unit_map.get(primary_name)
        backup_names = [self.backup_lb.get(i) for i in self.backup_lb.curselection()]
        backup_ids = [self.unit_map[n] for n in backup_names if n in self.unit_map]

        if self.inc_id:
            self.db.update_incident(self.inc_id, loc_id, t_name, reported, disp, arr, clr, dispn, cleared_flag, primary_id, backup_ids)
        else:
            self.inc_id = self.db.create_incident(loc_id, t_name, reported, disp, arr, clr, dispn, cleared_flag, primary_id, backup_ids)
        if self.on_saved:
            self.on_saved()
        self.destroy()


# -----------------------------
# Notes viewer window
# -----------------------------
class NotesWindow(tk.Toplevel):
    def __init__(self, master, db: DB, inc_id: int):
        super().__init__(master)
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
        ttk.Button(btnbar, text="Insert Time", command=lambda: self.entry.insert(tk.END, now_dt()+" ")).grid(row=0, column=0, padx=4, pady=4)
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
# Export helpers
# -----------------------------
class Exporter:
    def __init__(self, db: DB):
        self.db = db

    def export_excel(self, rows: List[sqlite3.Row], path: Path):
        try:
            import xlsxwriter  # type: ignore
        except Exception as e:
            # fallback to CSV
            csv_path = path.with_suffix('.csv')
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Primary", "Backups", "Disposition", "Status"]) 
                for r in rows:
                    w.writerow([r["reported_at"], r["dispatched_at"], r["arrived_at"], r["cleared_at"], r["type"], r["location_name"] or "", r["primary_units"] or "", r["backup_units"] or "", r["disposition"] or "", "Cleared" if r["is_cleared"] else "Active"]) 
            messagebox.showinfo("Exported CSV", f"xlsxwriter not installed. Saved CSV instead to\n{csv_path}")
            return

        wb = xlsxwriter.Workbook(path)
        ws = wb.add_worksheet("Incidents")
        headers = ["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Primary", "Backups", "Disposition", "Status"]
        for c, h in enumerate(headers):
            ws.write(0, c, h)
        for r_idx, r in enumerate(rows, start=1):
            values = [r["reported_at"], r["dispatched_at"], r["arrived_at"], r["cleared_at"], r["type"], r["location_name"] or "", r["primary_units"] or "", r["backup_units"] or "", r["disposition"] or "", "Cleared" if r["is_cleared"] else "Active"]
            for c, v in enumerate(values):
                ws.write(r_idx, c, v)
        ws.autofilter(0, 0, r_idx, len(headers)-1)
        wb.close()
        messagebox.showinfo("Exported", f"Saved Excel file to\n{path}")

    def export_pdf(self, rows: List[sqlite3.Row], path: Path):
        try:
            from reportlab.lib.pagesizes import letter, landscape
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
        except Exception:
            messagebox.showwarning(
                "Missing dependency",
                "reportlab is not installed. Run\n  pip install reportlab\n\nAlternatively, export as Excel/CSV and print to PDF.")
            return

        doc = SimpleDocTemplate(str(path), pagesize=landscape(letter), title="Incident Board")
        styles = getSampleStyleSheet()
        data = [["Reported", "Dispatched", "Arrived", "Cleared", "Type", "Location", "Primary", "Backups", "Disposition", "Status"]]
        for r in rows:
            data.append([
                r["reported_at"], r["dispatched_at"], r["arrived_at"], r["cleared_at"], r["type"], r["location_name"] or "",
                r["primary_units"] or "", r["backup_units"] or "", r["disposition"] or "", "Cleared" if r["is_cleared"] else "Active"
            ])
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('TEXTCOLOR', (0,0), (-1,0), colors.black),
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.whitesmoke, colors.beige])
        ]))
        story = [Paragraph("Incident Board", styles['Title']), table]
        doc.build(story)
        messagebox.showinfo("Exported", f"Saved PDF to\n{path}")

    def export_printable_html(self, rows: List[sqlite3.Row], path: Path):
        html_path = path.with_suffix('.html')
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write("""
<!doctype html><html><head><meta charset='utf-8'>
<title>Incident Board</title>
<style>
body{font-family:system-ui,Segoe UI,Arial,sans-serif;margin:24px}
.table{border-collapse:collapse;width:100%}
.table th,.table td{border:1px solid #ddd;padding:8px;font-size:13px}
.table th{background:#f3f4f6;text-align:left}
tr.active{background:#fff7f7}
tr.cleared{background:#e8f8ea}
</style></head><body>
<h2>Incident Board</h2>
<table class='table'>
<thead><tr><th>Reported</th><th>Dispatched</th><th>Arrived</th><th>Cleared</th><th>Type</th><th>Location</th><th>Primary</th><th>Backups</th><th>Disposition</th><th>Status</th></tr></thead>
<tbody>
""")
            for r in rows:
                cls = 'cleared' if r["is_cleared"] else 'active'
                f.write(
                    f"<tr class='{cls}'><td>{r['reported_at']}</td><td>{r['dispatched_at']}</td><td>{r['arrived_at']}</td><td>{r['cleared_at']}</td><td>{r['type']}</td><td>{r['location_name'] or ''}</td><td>{r['primary_units'] or ''}</td><td>{r['backup_units'] or ''}</td><td>{(r['disposition'] or '').replace('&','&amp;')}</td><td>{'Cleared' if r['is_cleared'] else 'Active'}</td></tr>"
                )
            f.write("""
</tbody></table></body></html>
""")
        webbrowser.open(html_path.as_uri())


# -----------------------------
# Main application
# -----------------------------
class App(tk.Tk):    
    def __init__(self):
        super().__init__()
        self.db = DB()
        self.title(APP_TITLE)
        self.geometry("1300x740")
        self.minsize(980, 620)

        try:
            self.style = ttk.Style(self)
            self.style.theme_use("clam")
            self.style.configure("TButton", padding=6)
            self.style.configure("Treeview", rowheight=26)
        except Exception:
            pass

        self.exporter = Exporter(self.db)

        self._build_menu()
        self._build_filters()
        self._build_board()
        self.refresh_board()

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
        m_file.add_command(label="Exit", command=self.destroy)
        menubar.add_cascade(label="File", menu=m_file)

        # Manage
        m_mng = tk.Menu(menubar, tearoff=False)
        m_mng.add_command(label="Locations", command=lambda: ListManager(self, self.db, "locations"))
        m_mng.add_command(label="Units", command=lambda: ListManager(self, self.db, "units"))
        m_mng.add_command(label="Incident Types", command=lambda: ListManager(self, self.db, "incident_types"))
        menubar.add_cascade(label="Manage", menu=m_mng)

        # Help
        m_help = tk.Menu(menubar, tearoff=False)
        m_help.add_command(label="About", command=lambda: messagebox.showinfo("About", "Simple offline incident board for race control\nBuilt with Python (Tkinter) + SQLite by Ryder Smith 2025-2026"))
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
        self.from_var = tk.StringVar(value=date.today().strftime(DATE_FMT))
        ttk.Entry(bar, textvariable=self.from_var, width=12).pack(side="left", padx=(4, 10))
        ttk.Label(bar, text="to").pack(side="left")
        self.to_var = tk.StringVar(value=date.today().strftime(DATE_FMT))
        ttk.Entry(bar, textvariable=self.to_var, width=12).pack(side="left", padx=(4, 12))

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

        ttk.Button(bar, text="New Incident", command=self.new_incident).pack(side="right")

    def _build_board(self):
        frame = ttk.Frame(self, padding=(12, 0))
        frame.pack(fill="both", expand=True)

        cols = ("reported", "dispatched", "arrived", "cleared", "type", "location", "units", "status")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("reported", text="Rec'd")
        self.tree.heading("dispatched", text="Disp")
        self.tree.heading("arrived", text="Arrv'd")
        self.tree.heading("cleared", text="Done")
        self.tree.heading("type", text="Type")
        self.tree.heading("location", text="Location")
        self.tree.heading("units", text="Unit(s)")
        self.tree.heading("status", text="Status")

        self.tree.column("reported", width=140)
        self.tree.column("dispatched", width=120)
        self.tree.column("arrived", width=120)
        self.tree.column("cleared", width=120)
        self.tree.column("type", width=120)
        self.tree.column("location", width=160)
        self.tree.column("units", width=220)
        self.tree.column("status", width=90)
        self.tree.pack(fill="both", expand=True, side="left")

        # Row colors via tags
        self.tree.tag_configure('cleared', background="#e8f8ea")
        self.tree.tag_configure('active', background="#fff5f5")

        # right-side actions
        side = ttk.Frame(frame)
        side.pack(side="right", fill="y", padx=(8,0))
        ttk.Button(side, text="View/Edit", command=self.edit_selected).pack(fill="x", pady=6)
        ttk.Button(side, text="Notes", command=self.open_notes).pack(fill="x", pady=6)
        ttk.Button(side, text="Mark Cleared", command=self.mark_cleared).pack(fill="x", pady=6)
        ttk.Button(side, text="Delete", command=self.delete_selected).pack(fill="x", pady=6)
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
        start = (self.from_var.get().strip() or None)
        end = (self.to_var.get().strip() or None)
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
                r["reported_at"], r["dispatched_at"], r["arrived_at"], r["cleared_at"], r["type"], r["location_name"] or "", units_label, "Cleared" if r["is_cleared"] else "Active"
            ), tags=(tag,))

    def reset_filters(self):
        today = date.today().strftime(DATE_FMT)
        self.from_var.set(today)
        self.to_var.set(today)
        self.loc_filter_var.set("All")
        self.type_filter_var.set("All")
        self.refresh_board()

    def new_incident(self):
        IncidentForm(self, self.db, on_saved=self.refresh_board)

    def get_selected_incident_id(self) -> Optional[int]:
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Please select an incident from the board.")
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

    def mark_cleared(self):
        iid = self.get_selected_incident_id()
        if not iid:
            return
        self.db.set_cleared(iid, True, now_dt())
        self.refresh_board()

    def delete_selected(self):
        iid = self.get_selected_incident_id()
        if not iid:
            return
        if messagebox.askyesno("Confirm", "Delete the selected incident?"):
            self.db.delete_incident(iid)
            self.refresh_board()

    def _current_rows_for_export(self) -> List[sqlite3.Row]:
        loc_id, tname, start, end = self._filters()
        return self.db.fetch_board(loc_id, tname, start, end)

    def export_excel(self):
        rows = self._current_rows_for_export()
        if not rows:
            messagebox.showinfo("Nothing to export", "No incidents match the current filters.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", ".xlsx"), ("CSV", ".csv")])
        if not path:
            return
        self.exporter.export_excel(rows, Path(path))

    def export_pdf(self):
        rows = self._current_rows_for_export()
        if not rows:
            messagebox.showinfo("Nothing to export", "No incidents match the current filters.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", ".pdf")])
        if not path:
            return
        self.exporter.export_pdf(rows, Path(path))

    def open_printable_html(self):
        rows = self._current_rows_for_export()
        if not rows:
            messagebox.showinfo("Nothing to show", "No incidents match the current filters.")
            return
        # Save alongside DB
        out = Path.cwd() / "incident_board.html"
        self.exporter.export_printable_html(rows, out)


if __name__ == "__main__":
    App().mainloop()
