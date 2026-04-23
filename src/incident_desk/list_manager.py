"""Manage-window for the four editable pick-lists (locations, units, incident_types, driver_codes)."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional

from .db import DB
from .dialogs import ask_for_text, dark_confirm, dark_info
from .icons import set_window_icon
from .window_utils import apply_dark_titlebar, position_on_parent


class ListManager(tk.Toplevel):
    def __init__(self, master, db: DB, table: str):
        super().__init__(master)
        self.withdraw()
        set_window_icon(self, {"units": "units", "locations": "locations", "incident_types": "types", "driver_codes": "driver_codes"}.get(table, "units"))
        self.after(0, lambda: (apply_dark_titlebar(self), position_on_parent(self, master)))
        self.db = db
        self.table = table
        self.title(f"Manage {table.replace('_', ' ').title()}")
        self.geometry("600x460")
        self.resizable(False, False)
        self.configure(padx=12, pady=12)

        self.tree = ttk.Treeview(self, columns=("name", "handle"), show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.heading("handle", text="")
        self.tree.column("name", width=420)
        self.tree.column("handle", width=40, anchor="center", stretch=False)
        self.tree.tag_configure("available",   background="#c8e6c9", foreground="#1a1a1a")
        self.tree.tag_configure("unavailable", background="#ffcdd2", foreground="#1a1a1a")
        self.tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

        self._drag_source: Optional[str] = None
        self.tree.bind("<ButtonPress-1>",   self._on_drag_start)
        self.tree.bind("<B1-Motion>",       self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_end)

        btn_add   = ttk.Button(self, text="Add",    style="New.TButton",    command=self.add)
        btn_edit  = ttk.Button(self, text="Edit",   style="Manage.TButton", command=self.edit)
        btn_del   = ttk.Button(self, text="Delete", style="Danger.TButton", command=self.delete)
        btn_close = ttk.Button(self, text="Close",  command=self.destroy)
        btn_add.grid(  row=1, column=0, pady=10, sticky="w")
        btn_edit.grid( row=1, column=1, pady=10, sticky="w", padx=(6, 0))
        btn_del.grid(  row=1, column=2, pady=10, sticky="w", padx=(6, 0))
        btn_close.grid(row=1, column=3, pady=10, sticky="e", padx=(6, 0))
        self.grid_columnconfigure(2, weight=1)

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
        handle = "⠿"  # Braille pattern dots-123456 — 6-dot drag handle
        if self.table == "locations":
            for r in self.db.list_locations():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], handle))
        elif self.table == "incident_types":
            for r in self.db.list_incident_types():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], handle))
        elif self.table == "driver_codes":
            for r in self.db.list_driver_codes():
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], handle))
        else:  # units
            for r in self.db.list_units_with_availability():
                tag = "available" if r["available"] else "unavailable"
                self.tree.insert("", "end", iid=str(r["id"]), values=(r["name"], handle), tags=(tag,))

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
            elif self.table == "driver_codes":
                self.db.add_driver_code(name)
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
        elif self.table == "driver_codes":
            old = self.tree.item(sel, "values")[0]
            new = ask_for_text(self, "Rename driver code", old)
            if new is None:
                return
            self.db.rename_driver_code(iid, new)
        else:
            old = self.tree.item(sel, "values")[0]
            new = ask_for_text(self, "Rename incident type", old)
            if new is None:
                return
            self.db.rename_incident_type(iid, new)
        self.refresh()

    def _on_drag_start(self, event):
        # Only start a drag when the user clicks the 6-dot handle cell (column #2).
        if (self.tree.identify_region(event.x, event.y) != "cell"
                or self.tree.identify_column(event.x) != "#2"):
            self._drag_source = None
            return
        row = self.tree.identify_row(event.y)
        if not row:
            self._drag_source = None
            return
        self._drag_source = row
        self.tree.selection_set(row)
        self.tree.config(cursor="fleur")

    def _on_drag_motion(self, event):
        if not self._drag_source:
            return
        target = self.tree.identify_row(event.y)
        if not target or target == self._drag_source:
            return
        self.tree.move(self._drag_source, "", self.tree.index(target))

    def _on_drag_end(self, event):
        if not self._drag_source:
            return
        self._drag_source = None
        self.tree.config(cursor="")
        ordered_ids = [int(iid) for iid in self.tree.get_children()]
        self.db.set_sort_order(self.table, ordered_ids)

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
        elif self.table == "driver_codes":
            count = self.db.driver_code_incident_count(name)
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
        elif self.table == "driver_codes":
            self.db.delete_driver_code(iid)
        else:
            self.db.delete_incident_type(iid)
        self.refresh()
