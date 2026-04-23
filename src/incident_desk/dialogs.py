"""Styled modal dialogs that match the rest of the app: info, yes/no confirm, text input."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from typing import Optional

from .icons import set_window_icon
from .window_utils import apply_dark_titlebar, position_on_parent


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


def dark_confirm(parent, title: str, message: str, *,
                 yes_text: str = "Yes", no_text: str = "No",
                 yes_style: str = "Danger.TButton",
                 icon_key: str = "confirm",
                 wraplength: int = 320) -> bool:
    """Styled Yes/No confirmation dialog. Returns True if the user clicked Yes."""
    result = [False]

    dlg = tk.Toplevel(parent)
    dlg.withdraw()
    dlg.title(title)
    dlg.resizable(False, False)
    set_window_icon(dlg, icon_key)
    dlg.after(0, lambda: apply_dark_titlebar(dlg))

    outer = ttk.Frame(dlg, padding=20)
    outer.pack(fill="both", expand=True)

    ttk.Label(outer, text=message, wraplength=wraplength, justify="left").pack(anchor="w")
    ttk.Separator(outer).pack(fill="x", pady=(16, 12))

    btn_row = ttk.Frame(outer)
    btn_row.pack(fill="x")

    def _yes():
        result[0] = True
        dlg.destroy()

    ttk.Button(btn_row, text=no_text,  command=dlg.destroy).pack(side="right", padx=(6, 0))
    ttk.Button(btn_row, text=yes_text, style=yes_style, command=_yes).pack(side="right")

    dlg.bind("<Return>", lambda e: _yes())
    dlg.bind("<Escape>", lambda e: dlg.destroy())

    dlg.transient(parent)
    position_on_parent(dlg, parent)
    dlg.grab_set()
    dlg.wait_window()

    return result[0]


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
