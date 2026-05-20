"""CLI for issuing license keys.

Usage:
    python tools/generate_license.py CUSTOMER_ID [YYYY-MM-DD]

CUSTOMER_ID must be uppercase A-Z / 0-9 / _ / - (2-32 chars).
The optional date overrides today's date (audit field, not enforced).

Examples:
    python tools/generate_license.py ROADAMERICA
    python tools/generate_license.py TRACK-002 2027-01-15

Requires `pip install -e .` in the project root so the package import works.
"""
from __future__ import annotations

import sys
from datetime import date

from incident_desk.license import generate_license, validate_license


def main(argv: list[str]) -> int:
    if not 2 <= len(argv) <= 3:
        print(__doc__, file=sys.stderr)
        return 2

    customer = argv[1]
    issued = None
    if len(argv) == 3:
        try:
            issued = date.fromisoformat(argv[2])
        except ValueError:
            print(f"ERROR: bad date {argv[2]!r} (expected YYYY-MM-DD)", file=sys.stderr)
            return 2

    try:
        key = generate_license(customer, issued)
    except ValueError as ex:
        print(f"ERROR: {ex}", file=sys.stderr)
        return 2

    # Sanity round-trip
    assert validate_license(key) is not None, "generated key failed self-validation"

    print(key)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
