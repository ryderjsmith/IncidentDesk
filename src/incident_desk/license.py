"""Offline HMAC-signed perpetual license keys.

Key format: ``CUSTOMER_ID.YYYYMMDD.HMAC``
  - CUSTOMER_ID is a human-readable identifier (uppercase, A-Z 0-9 - _)
  - YYYYMMDD is the issue date (audit only, not enforced)
  - HMAC is 16 base32 chars of HMAC-SHA256 over ``CUSTOMER_ID.YYYYMMDD``
    keyed by _SECRET.

NOTE on the secret: this lives in the bundled exe. A determined attacker
who extracts it can forge keys. Adequate for B2B "polite enforcement";
upgrade to Ed25519 signatures if real DRM ever becomes a requirement.
"""
from __future__ import annotations

import base64
import hmac
import re
from datetime import date
from hashlib import sha256
from pathlib import Path
from typing import Optional, Tuple

from .config import DB_DIR

# REGENERATE THIS BEFORE FIRST SHIPPING TO A NEW CUSTOMER POOL.
# Changing it invalidates every license key already issued.
_SECRET = bytes.fromhex("8c6f9e2c556534a4572d02ece786699a1f5c1f9a3ec27bec9a5d6413423dc39b")

_CUSTOMER_RE = re.compile(r"^[A-Z0-9_\-]{2,32}$")
_SIG_LEN = 16  # base32 chars
LICENSE_PATH = DB_DIR / "license.key"


def _signature(payload: str) -> str:
    digest = hmac.new(_SECRET, payload.encode("utf-8"), sha256).digest()
    return base64.b32encode(digest).decode("ascii")[:_SIG_LEN]


def generate_license(customer_id: str, issued: Optional[date] = None) -> str:
    """Produce a license key for a customer. ``customer_id`` must be uppercase
    A-Z / 0-9 / _ / -, 2-32 chars. Strict input — lowercase is rejected."""
    if not _CUSTOMER_RE.match(customer_id):
        raise ValueError(
            f"customer_id must match {_CUSTOMER_RE.pattern} (got: {customer_id!r})"
        )
    issued = issued or date.today()
    payload = f"{customer_id}.{issued:%Y%m%d}"
    return f"{payload}.{_signature(payload)}"


def validate_license(key: str) -> Optional[Tuple[str, date]]:
    """Returns (customer_id, issue_date) if the key is well-formed and
    signature-verifies. Returns None on any failure."""
    if not key:
        return None
    try:
        customer_id, issued_str, sig = key.strip().upper().split(".")
        if not _CUSTOMER_RE.match(customer_id):
            return None
        issued = date(int(issued_str[:4]), int(issued_str[4:6]), int(issued_str[6:8]))
    except (ValueError, AttributeError):
        return None
    expected = _signature(f"{customer_id}.{issued:%Y%m%d}")
    if not hmac.compare_digest(sig, expected):
        return None
    return customer_id, issued


def load_license() -> Optional[Tuple[str, str, date]]:
    """Read the saved license file and re-validate. Returns (key, customer_id,
    issued) if valid, else None."""
    try:
        key = LICENSE_PATH.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    validated = validate_license(key)
    if validated is None:
        return None
    return (key, validated[0], validated[1])


def save_license(key: str) -> None:
    LICENSE_PATH.parent.mkdir(parents=True, exist_ok=True)
    LICENSE_PATH.write_text(key.strip().upper() + "\n", encoding="utf-8")
