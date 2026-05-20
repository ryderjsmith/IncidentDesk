"""License activation and info dialogs — styled to match the rest of the app."""
from __future__ import annotations

import tkinter as tk
from datetime import date
from tkinter import ttk
from typing import Optional

from .icons import set_window_icon
from .license import LICENSE_PATH, save_license, validate_license
from .window_utils import apply_dark_titlebar, position_on_parent


def show_activation_dialog(parent) -> Optional[str]:
    """Modal license-entry dialog. Returns the customer_id on success, None if
    the user dismisses without activating. Saves the validated key to disk."""
    result: list[Optional[str]] = [None]

    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title("License Required")
    dlg.resizable(False, False)
    set_window_icon(dlg, "input")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    ttk.Label(
        outer,
        text=(
            "Incident Desk requires a license key to run.\n\n"
            "Paste the key your vendor provided, then click Activate."
        ),
        wraplength=380, justify="left",
    ).pack(anchor="w")

    ttk.Label(outer, text="License key:").pack(anchor="w", pady=(14, 4))
    entry = ttk.Entry(outer, width=44)
    entry.pack(fill="x")
    entry.focus_set()

    status_var = tk.StringVar(value="")
    status = ttk.Label(outer, textvariable=status_var, foreground="#c62828")
    status.pack(anchor="w", pady=(6, 0))

    ttk.Separator(outer).pack(fill="x", pady=(16, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _activate():
        key = entry.get().strip()
        validated = validate_license(key)
        if validated is None:
            status_var.set("Invalid license key.")
            entry.focus_set()
            entry.select_range(0, "end")
            return
        try:
            save_license(key)
        except OSError as ex:
            status_var.set(f"Could not save license: {ex}")
            return
        result[0] = validated[0]
        dlg.destroy()

    def _exit():
        dlg.destroy()

    ttk.Button(btn_row, text="Exit", command=_exit).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="Activate", style="Manage.TButton",
               command=_activate).pack(side="right")

    dlg.bind("<Return>", lambda e: _activate())
    dlg.bind("<Escape>", lambda e: _exit())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()

    return result[0]


def show_license_info(parent, key: str, customer_id: str, issued: date) -> None:
    """Read-only dialog showing the current license details with a copy button.
    Lets the user retrieve their key for transferring to another machine."""
    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title("License Info")
    dlg.resizable(False, False)
    set_window_icon(dlg, "info")
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    info = ttk.Frame(outer)
    info.pack(fill="x")
    ttk.Label(info, text="Licensed to:", foreground="#555").grid(row=0, column=0, sticky="w", pady=(0, 4))
    ttk.Label(info, text=customer_id, font=("TkDefaultFont", 9, "bold")).grid(row=0, column=1, sticky="w", padx=(8, 0))
    ttk.Label(info, text="Issued:", foreground="#555").grid(row=1, column=0, sticky="w")
    ttk.Label(info, text=issued.isoformat()).grid(row=1, column=1, sticky="w", padx=(8, 0))

    ttk.Label(outer, text="License key:", foreground="#555").pack(anchor="w", pady=(14, 4))
    key_entry = ttk.Entry(outer, width=46)
    key_entry.insert(0, key)
    key_entry.configure(state="readonly")
    key_entry.pack(fill="x")

    ttk.Label(
        outer, wraplength=400, justify="left", foreground="#555",
        text=(
            "Save this key somewhere safe. To install on a new machine, "
            f"enter it in the activation dialog on first launch.\n\n"
            f"Stored at:\n{LICENSE_PATH}"
        ),
    ).pack(anchor="w", pady=(10, 0))

    status_var = tk.StringVar(value="")
    ttk.Label(outer, textvariable=status_var, foreground="#2e7d32").pack(anchor="w", pady=(6, 0))

    ttk.Separator(outer).pack(fill="x", pady=(14, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _copy():
        dlg.clipboard_clear()
        dlg.clipboard_append(key)
        dlg.update()  # required for clipboard to persist after dialog closes
        status_var.set("Copied to clipboard.")

    ttk.Button(btn_row, text="Close", command=dlg.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text="Copy License Key", style="Manage.TButton",
               command=_copy).pack(side="right")

    dlg.bind("<Escape>", lambda e: dlg.destroy())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()
