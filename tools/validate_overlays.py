#!/usr/bin/env python3
"""Validate transfers_in.json + transfers_out.json against Squads_Data.xlsx.

Checks:
- Every (division, club) key in overlays exists in DB (no typos like "Sporting CP" vs "Sporting").
- Suggests closest-match fixes for any mismatch.

Exit code 0 = all good, 1 = mismatches found.
"""

from __future__ import annotations
import json, sys
from collections import defaultdict
from pathlib import Path

import openpyxl
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_DB_DIR = Path("/Users/ruimiguelneves/Library/Group Containers/UBF8T346G9.OneDriveSyncClientSuite/OneDrive.noindex/OneDrive/Claude/DB")
DB_PATH = next(
    (p for p in [_DB_DIR / "Squads.xlsm", _DB_DIR / "Squads_Data.xlsm", _DB_DIR / "Squads_Data.xlsx"] if p.exists()),
    _DB_DIR / "Squads.xlsm"
)
REPO = Path("/Users/ruimiguelneves/Code/fpl-dashboard")


def load_db_clubs():
    """Load real divisions AND synthetic '🌐 {Country}' virtual divisions (no-division rows grouped by country)."""
    wb = openpyxl.load_workbook(DB_PATH, read_only=True, data_only=True)
    ws = wb["Players"]
    db = defaultdict(set)
    for row in ws.iter_rows(min_row=2, values_only=True):
        club = row[4]
        country = row[5]
        division = row[6]
        if not club:
            continue
        if division:
            db[division].add(club)
        elif country:
            # Virtual division key — matches what world_data.json uses
            db[f"🌐 {country}"].add(club)
    return db


def main():
    db = load_db_clubs()
    errors = 0

    for fname in ["transfers_in.json", "transfers_out.json"]:
        p = REPO / fname
        if not p.exists():
            continue
        overlay = json.loads(p.read_text())
        print(f"\n=== {fname} ===")
        for div, clubs in overlay.items():
            if div.startswith("_"):
                continue
            if div not in db:
                print(f"  ❌ division '{div}' NOT in DB")
                errors += 1
                continue
            for club in clubs.keys():
                if club not in db[div]:
                    candidates = sorted(
                        c for c in db[div]
                        if club.lower() in c.lower() or c.lower() in club.lower()
                    )
                    print(f"  ❌ '{div}' / '{club}' missing — closest in DB: {candidates}")
                    errors += 1
        if errors == 0:
            print("  ✓ all club keys match the DB")

    sys.exit(0 if errors == 0 else 1)


if __name__ == "__main__":
    main()
