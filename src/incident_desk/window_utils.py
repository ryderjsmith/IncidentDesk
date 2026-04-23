"""Window-level helpers: dark titlebar (Windows DWM), cascading child positioning,
and a DateEntry subclass that carries the dark titlebar into its calendar popup."""
from __future__ import annotations

import ctypes

from tkcalendar import DateEntry as _DateEntry


class DateEntry(_DateEntry):
    """DateEntry that applies the dark title bar to its calendar dropdown."""
    def drop_down(self):
        super().drop_down()
        try:
            apply_dark_titlebar(self._top_cal)
        except Exception:
            pass


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
