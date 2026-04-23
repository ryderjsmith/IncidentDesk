"""Main application: root Tk window, menu, filter bar, incident board, and lifecycle."""
from __future__ import annotations

import json
import os
import sqlite3
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, ttk
from typing import List, Optional

from .config import APP_TITLE, APP_VERSION, DATE_FMT, DISPLAY_DATE_FMT
from .dates import fmt_dt, now_dt
from .db import DB
from .dialogs import dark_confirm, dark_info
from .exporter import Exporter
from .icons import set_icon, set_window_icon
from .incident_form import IncidentForm
from .list_manager import ListManager
from .notes_windows import BillablesWindow, NotesWindow
from .updater import check_for_update, download_file, update_prompt
from .user_guide import show_user_guide
from .window_utils import DateEntry, apply_dark_titlebar, position_on_parent


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
        m_file.add_command(label="Export to PDF", command=self.export_pdf)
        m_file.add_separator()
        m_file.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=m_file)

        # Manage
        m_mng = tk.Menu(menubar, tearoff=False)
        m_mng.add_command(label="Locations", command=lambda: ListManager(self, self.db, "locations"))
        m_mng.add_command(label="Units", command=lambda: ListManager(self, self.db, "units"))
        m_mng.add_command(label="Incident Types", command=lambda: ListManager(self, self.db, "incident_types"))
        m_mng.add_command(label="Driver Codes", command=lambda: ListManager(self, self.db, "driver_codes"))
        m_mng.add_separator()
        m_mng.add_command(label="Export Lists...", command=self.export_lists)
        m_mng.add_command(label="Import Lists...", command=self.import_lists)
        menubar.add_cascade(label="Manage", menu=m_mng)

        # Help
        m_help = tk.Menu(menubar, tearoff=False)
        m_help.add_command(label="User Guide", command=lambda: show_user_guide(self))
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
        self.loc_filter_cb.pack(side="left", padx=(4, 12))

        # Type filter
        ttk.Label(bar, text="Type").pack(side="left")
        self.type_filter_var = tk.StringVar(value="All")
        types = ["All"] + [r["name"] for r in self.db.list_incident_types()]
        self.type_filter_cb = ttk.Combobox(bar, textvariable=self.type_filter_var, values=types, width=16, state="readonly")
        self.type_filter_cb.pack(side="left", padx=(4, 12))

        ttk.Button(bar, text="Apply", command=self.refresh_board).pack(side="left")
        ttk.Button(bar, text="Reset", command=self.reset_filters).pack(side="left", padx=(6, 0))

        ttk.Label(bar, text="").pack(side="left", expand=True)

        ttk.Button(bar, text="New Incident", style="New.TButton", command=self.new_incident).pack(side="right")

    def _build_board(self):
        frame = ttk.Frame(self, padding=(12, 0))
        frame.pack(fill="both", expand=True)

        cols = ("reported", "dispatched", "arrived", "cleared", "type", "location", "car", "driver", "units", "status")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        self.tree.heading("reported", text="Rec'd")
        self.tree.heading("dispatched", text="Disp")
        self.tree.heading("arrived", text="Arrv'd")
        self.tree.heading("cleared", text="Done")
        self.tree.heading("type", text="Type")
        self.tree.heading("location", text="Location")
        self.tree.heading("car", text="Car #")
        self.tree.heading("driver", text="Driver Code")
        self.tree.heading("units", text="Unit(s)")
        self.tree.heading("status", text="Status")

        self.tree.column("reported",   width=140, anchor="center")
        self.tree.column("dispatched", width=120, anchor="center")
        self.tree.column("arrived",    width=120, anchor="center")
        self.tree.column("cleared",    width=120, anchor="center")
        self.tree.column("type",       width=120, anchor="center")
        self.tree.column("location",   width=160, anchor="center")
        self.tree.column("car",        width=70,  anchor="center")
        self.tree.column("driver",     width=100, anchor="center")
        self.tree.column("units",      width=140, anchor="center")
        self.tree.column("status",     width=90,  anchor="center")
        self.tree.pack(fill="both", expand=True, side="left")

        self.tree.tag_configure('cleared', background="#c8e6c9", foreground="#1a1a1a")
        self.tree.tag_configure('active',  background="#ffcdd2", foreground="#1a1a1a")

        side = ttk.Frame(frame)
        side.pack(side="right", fill="y", padx=(8, 0))
        ttk.Button(side, text="Manage Units", style="Manage.TButton", command=lambda: ListManager(self, self.db, "units")).pack(fill="x", pady=6)
        ttk.Separator(side).pack(fill="x", pady=4)
        ttk.Button(side, text="View/Edit", command=self.edit_selected).pack(fill="x", pady=6)
        ttk.Button(side, text="Notes", command=self.open_notes).pack(fill="x", pady=6)
        ttk.Button(side, text="Billables", command=self.open_billables).pack(fill="x", pady=6)
        ttk.Button(side, text="Toggle Cleared", command=self.mark_cleared).pack(fill="x", pady=6)
        ttk.Button(side, text="Delete", style="Danger.TButton", command=self.delete_selected).pack(fill="x", pady=6)
        ttk.Separator(side).pack(fill="x", pady=8)
        ttk.Button(side, text="Export PDF", command=self.export_pdf).pack(fill="x", pady=6)

        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

    # ----- Actions
    def _filters(self):
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
                r["type"], r["location_name"] or "", r["car_number"] or "", r["driver_code"] or "", units_label, "Cleared" if r["is_cleared"] else "Active"
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

    def _export_title(self) -> str:
        d1 = self.from_picker.get_date().strftime(DISPLAY_DATE_FMT)
        d2 = self.to_picker.get_date().strftime(DISPLAY_DATE_FMT)
        return d1 if d1 == d2 else f"{d1} – {d2}"

    def export_pdf(self):
        rows = self._current_rows_for_export()
        if not rows:
            dark_info(self, "Nothing to export", "No incidents match the current filters.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", ".pdf")])
        if not path:
            return
        self.exporter.export_pdf(rows, Path(path), self, title=self._export_title())

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
            "driver_codes": [r["name"] for r in self.db.list_driver_codes()],
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

        added = {"locations": 0, "incident_types": 0, "units": 0, "driver_codes": 0}
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
        for name in data.get("driver_codes", []):
            if isinstance(name, str) and name.strip():
                before = len(self.db.list_driver_codes())
                self.db.add_driver_code(name.strip())
                if len(self.db.list_driver_codes()) > before:
                    added["driver_codes"] += 1

        dark_info(self, "Import Complete",
                  f"Added:\n  {added['locations']} location(s)\n  {added['incident_types']} incident type(s)\n  {added['units']} unit(s)\n  {added['driver_codes']} driver code(s)\n\nExisting entries were not duplicated.")

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
        dlg.protocol("WM_DELETE_WINDOW", lambda: None)

        outer = ttk.Frame(dlg, padding=24)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Downloading update...\nThe app will restart when finished.",
                  wraplength=320, justify="left").pack(anchor="w")

        dlg.transient(self)
        position_on_parent(dlg, self)
        dlg.grab_set()

        dest = Path(tempfile.gettempdir()) / "IncidentDesk-Setup.exe"
        part = dest.with_suffix(dest.suffix + ".part")

        def worker():
            ok = download_file(url, part)
            if ok:
                try:
                    os.replace(part, dest)
                except OSError:
                    ok = False
            self.after(0, lambda: _finish(ok))

        def _finish(ok: bool):
            dlg.destroy()
            if not ok:
                dark_info(self, "Update Failed",
                          "Could not download the update. Please check your connection and try again.")
                return
            # Launch installer detached so it survives this process exiting,
            # then shut down so the installer can replace the exe.
            try:
                subprocess.Popen([str(dest), "/SILENT"],
                                 creationflags=subprocess.DETACHED_PROCESS)
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
        ttk.Button(export_row, text="Export to PDF", style="Manage.TButton",
                   command=lambda: self.export_pdf()).pack(side="left")

        ttk.Separator(outer).pack(fill="x", pady=(4, 12))

        action_row = ttk.Frame(outer)
        action_row.pack(fill="x")
        ttk.Button(action_row, text="Cancel",           command=dlg.destroy).pack(side="left")
        ttk.Button(action_row, text="Exit Without Exporting", style="Danger.TButton",
                   command=lambda: [dlg.destroy(), self.destroy()]).pack(side="right")

        dlg.transient(self)
        position_on_parent(dlg, self)
