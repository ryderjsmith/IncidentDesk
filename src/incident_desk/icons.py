"""Window icons: favicon.ico for the root, emoji-rendered distinct icons per dialog type."""
from __future__ import annotations

from .config import ICO_PATH


def set_icon(window) -> None:
    """Apply favicon.ico to any Tk or Toplevel window."""
    try:
        window.iconbitmap(str(ICO_PATH))
    except Exception:
        pass


_WINDOW_ICONS = {
    "incident_form": ("🚨", "#c0392b"),
    "notes":         ("📝", "#2980b9"),
    "billables":     ("🧾", "#d4a74a"),
    "units":         ("🚒", "#27ae60"),
    "locations":     ("📍", "#8e44ad"),
    "types":         ("🏷", "#e67e22"),
    "driver_codes":  ("🏎", "#16a085"),
    "info":          ("ℹ",  "#2980b9"),
    "confirm":       ("⚠",  "#e67e22"),
    "input":         ("✏",  "#16a085"),
    "guide":         ("📖", "#2c3e50"),
    "export":        ("💾", "#546e7a"),
}
_icon_cache: dict = {}


def _build_icon(char: str) -> list:
    """Return a list of PhotoImages at [16, 32, 48, 256] px so Windows picks the best size per context."""
    try:
        from PIL import Image, ImageDraw, ImageFont, ImageTk
        font_cache: dict = {}
        photos = []
        for size in (16, 32, 48, 256):
            img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
            d = ImageDraw.Draw(img)
            fs = max(8, int(size * 0.72))
            if fs not in font_cache:
                font = None
                for name in [
                    r"C:\Windows\Fonts\seguiemj.ttf",
                    r"C:\Windows\Fonts\seguisym.ttf",
                    "Segoe UI Emoji", "Segoe UI Symbol", "Arial",
                ]:
                    try:
                        font = ImageFont.truetype(name, fs)
                        break
                    except Exception:
                        pass
                font_cache[fs] = font or ImageFont.load_default()
            font = font_cache[fs]
            bb = d.textbbox((0, 0), char, font=font)
            x = (size - (bb[2] - bb[0])) // 2 - bb[0]
            y = (size - (bb[3] - bb[1])) // 2 - bb[1]
            d.text((x, y), char, fill="white", font=font)
            photos.append(ImageTk.PhotoImage(img))
        return photos
    except Exception:
        return []


def set_window_icon(window, key: str) -> None:
    """Apply a type-specific icon to a Toplevel. Falls back to favicon.ico on failure."""
    if key not in _icon_cache:
        spec = _WINDOW_ICONS.get(key)
        _icon_cache[key] = _build_icon(spec[0]) if spec else []
    photos = _icon_cache.get(key) or []
    if photos:
        window.iconphoto(False, *photos)
        window._window_icons = photos  # prevent GC
    else:
        set_icon(window)
