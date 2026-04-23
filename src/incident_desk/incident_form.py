"""New / edit incident entry form with inline notes and billables."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional

from .dates import fmt_dt, now_dt, now_dt_display, parse_display_dt
from .db import DB
from .icons import set_window_icon
from .list_manager import ListManager
from .window_utils import apply_dark_titlebar, position_on_parent


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
        self.geometry("820x605")
        self.transient(master)
        self.configure(padx=12, pady=12)
        self.resizable(False, False)

        # Layout grid
        for i in range(8):
            self.grid_rowconfigure(i, pad=4)
        self.grid_columnconfigure(1, weight=1)

        ttk.Label(self, text="Driver Code").grid(row=0, column=0, sticky="w")
        self.driver_code_var = tk.StringVar()
        self.driver_code_cb = ttk.Combobox(self, textvariable=self.driver_code_var,
                                           values=[r["name"] for r in self.db.list_driver_codes()], state="readonly")
        self.driver_code_cb.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Manage", command=lambda: self._manage("driver_codes")).grid(row=0, column=2, sticky="w")

        ttk.Label(self, text="Location").grid(row=1, column=0, sticky="w")
        self.loc_var = tk.StringVar()
        self.loc_cb = ttk.Combobox(self, textvariable=self.loc_var, values=[r["name"] for r in self.db.list_locations()], state="readonly")
        self.loc_cb.grid(row=1, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Manage", command=lambda: self._manage("locations")).grid(row=1, column=2, sticky="w")

        ttk.Label(self, text="Incident Type").grid(row=2, column=0, sticky="w")
        self.type_var = tk.StringVar()
        self.type_cb = ttk.Combobox(self, textvariable=self.type_var, values=[r["name"] for r in self.db.list_incident_types()], state="readonly")
        self.type_cb.grid(row=2, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(self, text="Manage", command=lambda: self._manage("incident_types")).grid(row=2, column=2, sticky="w")

        ttk.Label(self, text="Car #").grid(row=3, column=0, sticky="w")
        self.car_number_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.car_number_var).grid(row=3, column=1, sticky="ew", padx=(0, 6))

        self.reported_var = tk.StringVar(value=now_dt_display())
        self.dispatched_var = tk.StringVar()
        self.arrived_var = tk.StringVar()
        self.cleared_var = tk.StringVar()

        self._time_row("Reported", self.reported_var, row=4)

        self.unit_map = {u["name"]: u["id"] for u in self.db.list_units()}

        ttk.Label(self, text="Primary Unit").grid(row=5, column=0, sticky="w")
        self.primary_var = tk.StringVar()
        self.primary_cb = ttk.Combobox(self, textvariable=self.primary_var,
                                       values=self._available_unit_names(), state="readonly")
        self.primary_cb.grid(row=5, column=1, sticky="ew")

        self._time_row("Dispatched", self.dispatched_var, row=6)
        self._time_row("Arrived", self.arrived_var, row=7)
        self._time_row("Cleared", self.cleared_var, row=8)

        frame_notes = ttk.LabelFrame(self, text="Add Note (time-stamped)")
        frame_notes.grid(row=9, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame_notes.grid_columnconfigure(0, weight=1)
        self.note_text = tk.Text(frame_notes, height=4)
        self.note_text.grid(row=0, column=0, sticky="ew")
        btns = ttk.Frame(frame_notes)
        btns.grid(row=0, column=1, sticky="ns")
        ttk.Button(btns, text="Insert Time", command=lambda: self.note_text.insert(tk.END, now_dt_display()+" ")).grid(row=0, column=0, padx=4, pady=4)
        ttk.Button(btns, text="Save Note", command=self._save_quick_note).grid(row=1, column=0, padx=4)

        frame_bill = ttk.LabelFrame(self, text="Add Billable")
        frame_bill.grid(row=10, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        frame_bill.grid_columnconfigure(0, weight=1)
        self.bill_text = tk.Text(frame_bill, height=3)
        self.bill_text.grid(row=0, column=0, sticky="ew")
        btns_b = ttk.Frame(frame_bill)
        btns_b.grid(row=0, column=1, sticky="ns")
        ttk.Button(btns_b, text="Save Billable", command=self._save_quick_billable).grid(row=0, column=0, padx=4, pady=4)

        sep = ttk.Separator(self)
        sep.grid(row=11, column=0, columnspan=3, sticky="ew", pady=6)
        self.cleared_var_bool = tk.IntVar(value=0)
        ttk.Checkbutton(self, text="Mark as Cleared", variable=self.cleared_var_bool,
                        onvalue=1, offvalue=0).grid(row=12, column=0, sticky="w")
        ttk.Button(self, text="Close", command=self.destroy).grid(row=12, column=1, sticky="e", padx=(0, 6))
        ttk.Button(self, text="Save", command=self.save).grid(row=12, column=2, sticky="e")

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
        self.driver_code_cb["values"] = [r["name"] for r in self.db.list_driver_codes()]
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
            loc = next((r["name"] for r in self.db.list_locations() if r["id"] == inc["location_id"]), None)
            if loc:
                self.loc_var.set(loc)
        self.type_var.set(inc["type"])
        self.car_number_var.set(inc["car_number"] or "")
        self.driver_code_var.set(inc["driver_code"] or "")
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

        primary_name = self.primary_var.get().strip()
        primary_id = self.unit_map.get(primary_name)

        car_number = self.car_number_var.get().strip()
        driver_code = self.driver_code_var.get().strip()

        if self.inc_id:
            self.db.update_incident(self.inc_id, loc_id, t_name, reported, disp, arr, clr, "", cleared_flag, primary_id, [], car_number, driver_code)
        else:
            self.inc_id = self.db.create_incident(loc_id, t_name, reported, disp, arr, clr, "", cleared_flag, primary_id, [], car_number, driver_code)

        # Flush any pending note text now that inc_id is guaranteed to exist
        self._flush_note()
        self._flush_billable()

        if self.on_saved:
            self.on_saved()
