"""In-app User Guide window."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .icons import set_window_icon
from .window_utils import apply_dark_titlebar, position_on_parent


def show_user_guide(parent) -> None:
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("User Guide")
    win.geometry("860x780")
    win.minsize(700, 600)
    set_window_icon(win, "guide")
    win.after(0, lambda: apply_dark_titlebar(win))
    win.transient(parent)

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
    txt.tag_configure("key", font=("Consolas", 9), relief="groove",
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

    h2("Managing Units, Locations, Types & Driver Codes")
    body("Click Manage Units (right panel) or use the Manage menu to open list managers for:")
    body("  • Units — vehicles or personnel that can be assigned to incidents.")
    body("  • Locations — named locations selectable on the incident form.")
    body("  • Incident Types — categories used to classify incidents.")
    body("  • Driver Codes — driver identifiers selectable on the incident form.")
    body("In each manager you can Add, Edit, or Delete entries.")
    body("To reorder, grab the ⠿ handle on the right side of a row and drag it up or down.")
    body("Items cannot be deleted while attached to existing incidents.")
    body("")

    h2("Exporting")
    body("Use the Export PDF button on the right panel or File → Export to PDF to save the")
    body("currently filtered board as a formatted PDF report.")
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
    position_on_parent(win, parent)
