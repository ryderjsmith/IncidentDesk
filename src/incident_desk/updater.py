"""In-app update check against GitHub Releases + installer download + user prompt."""
from __future__ import annotations

import json
import urllib.error
import urllib.request
from pathlib import Path
from typing import Optional, Tuple

from .config import APP_VERSION, GITHUB_REPO
from .dialogs import dark_confirm


def _parse_version(v: str) -> tuple:
    """Turn 'v1.2.3' / '1.2.3-beta' into a comparable tuple (1, 2, 3)."""
    s = v.lstrip("vV").split("-")[0].split("+")[0]
    out = []
    for p in s.split("."):
        try:
            out.append(int(p))
        except ValueError:
            out.append(0)
    return tuple(out)


def check_for_update() -> Optional[Tuple[str, str]]:
    """Query GitHub for the latest release. Returns (tag, installer_url) if
    newer than APP_VERSION, else None. Fails silently on network errors."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(url, headers={
        "User-Agent": "IncidentDesk-Updater",
        "Accept": "application/vnd.github+json",
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, ValueError):
        return None

    tag = data.get("tag_name", "")
    if not tag or _parse_version(tag) <= _parse_version(APP_VERSION):
        return None

    installer = next(
        (a for a in data.get("assets", []) if a.get("name", "").lower().endswith("-setup.exe")),
        None,
    )
    if not installer:
        return None
    return (tag, installer["browser_download_url"])


def download_file(url: str, dest: Path) -> bool:
    """Stream download to dest. Returns True on success."""
    req = urllib.request.Request(url, headers={"User-Agent": "IncidentDesk-Updater"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as f:
            while True:
                chunk = resp.read(65536)
                if not chunk:
                    break
                f.write(chunk)
        return True
    except (urllib.error.URLError, urllib.error.HTTPError, OSError):
        return False


def update_prompt(parent, current: str, latest: str) -> bool:
    """Styled update-available dialog. Returns True if user chose Update Now."""
    message = (
        f"Incident Desk {latest} is available.\n"
        f"You are currently running {current}.\n\n"
        "Would you like to download and install the update now?\n\n"
        "Your incident data will be preserved automatically."
    )
    return dark_confirm(
        parent, "Update Available", message,
        yes_text="Update Now", no_text="Later",
        yes_style="Manage.TButton", icon_key="export",
        wraplength=380,
    )
