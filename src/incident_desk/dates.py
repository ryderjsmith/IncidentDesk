"""Datetime formatting helpers bridging ISO DB storage and MM-DD-YYYY UI display."""
from __future__ import annotations

from datetime import datetime

from .config import DATE_FMT, DISPLAY_DATE_FMT, DISPLAY_DT_FMT, DT_FMT


def now_dt() -> str:
    """Current datetime in ISO format (for DB storage)."""
    return datetime.now().strftime(DT_FMT)


def now_dt_display() -> str:
    """Current datetime in MM-DD-YYYY display format (for UI)."""
    return datetime.now().strftime(DISPLAY_DT_FMT)


def fmt_dt(s: str) -> str:
    """Convert a stored ISO datetime string to MM-DD-YYYY display format."""
    if not s:
        return s
    for fmt in (DT_FMT, "%d:%m:%Y %H:%M:%S", "%m:%d:%Y %H:%M:%S", "%m-%d-%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime(DISPLAY_DT_FMT)
        except ValueError:
            pass
    return s


def parse_display_dt(s: str) -> str:
    """Convert display datetime (MM-DD-YYYY HH:MM:SS) to ISO for DB storage."""
    if not s:
        return s
    for fmt in (DISPLAY_DT_FMT, DT_FMT, "%d:%m:%Y %H:%M:%S", "%m:%d:%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime(DT_FMT)
        except ValueError:
            pass
    return s


def parse_display_date(s: str) -> str:
    """Convert display date (MM-DD-YYYY) to ISO for DB queries."""
    if not s:
        return s
    for fmt in (DISPLAY_DATE_FMT, DATE_FMT, "%d:%m:%Y", "%m:%d:%Y"):
        try:
            return datetime.strptime(s, fmt).strftime(DATE_FMT)
        except ValueError:
            pass
    return s
