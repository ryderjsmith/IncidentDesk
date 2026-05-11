"""Per-incident notes (time-stamped) and billables (free-form) viewer windows."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .dates import now_dt, now_dt_display
from .db import DB
from .icons import set_window_icon
from .window_utils import apply_dark_titlebar, position_on_parent


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

        view = ttk.Frame(self)
        view.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        view.grid_rowconfigure(0, weight=1)
        view.grid_columnconfigure(0, weight=1)

        self.view = tk.Text(view, wrap="word", state="disabled",
                            padx=10, pady=10, relief="solid", bd=1, cursor="arrow")
        sb = ttk.Scrollbar(view, orient="vertical", command=self.view.yview)
        self.view.configure(yscrollcommand=sb.set)
        self.view.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.view.tag_configure("ts", font=("TkDefaultFont", 9, "bold"), foreground="#1565c0")
        self.view.tag_configure("body", font=("TkDefaultFont", 10),
                                lmargin1=10, lmargin2=10, spacing3=10)

        self.refresh()

    def refresh(self):
        self.view.configure(state="normal")
        self.view.delete("1.0", tk.END)
        for n in self.db.list_notes(self.inc_id):
            self.view.insert(tk.END, n["ts"] + "\n", "ts")
            self.view.insert(tk.END, n["body"] + "\n", "body")
        self.view.configure(state="disabled")

    def add_note(self):
        text = self.entry.get("1.0", tk.END).strip()
        if not text:
            return
        self.db.add_note(self.inc_id, now_dt(), text)
        self.entry.delete("1.0", tk.END)
        self.refresh()


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

        view = ttk.Frame(self)
        view.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        view.grid_rowconfigure(0, weight=1)
        view.grid_columnconfigure(0, weight=1)

        self.view = tk.Text(view, wrap="word", state="disabled",
                            padx=10, pady=10, relief="solid", bd=1, cursor="arrow")
        sb = ttk.Scrollbar(view, orient="vertical", command=self.view.yview)
        self.view.configure(yscrollcommand=sb.set)
        self.view.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.view.tag_configure("body", font=("TkDefaultFont", 10), spacing3=10)

        self.refresh()

    def refresh(self):
        self.view.configure(state="normal")
        self.view.delete("1.0", tk.END)
        for b in self.db.list_billables(self.inc_id):
            self.view.insert(tk.END, b["body"] + "\n", "body")
        self.view.configure(state="disabled")

    def add_billable(self):
        text = self.entry.get("1.0", tk.END).strip()
        if not text:
            return
        self.db.add_billable(self.inc_id, text)
        self.entry.delete("1.0", tk.END)
        self.refresh()
