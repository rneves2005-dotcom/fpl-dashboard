#!/usr/bin/env python3
"""Compare the current Squads_Data.xlsx against the most recent .bak file.

Use this when you (Rui) edited the xlsx directly and want to see what changed
before Claude commits the regenerated JSON. Shows added/removed/modified rows
so we can confirm the diff matches your intent.

Usage:
    python3 tools/diff_db.py            # diff against most recent backup
    python3 tools/diff_db.py <bak-name> # diff against a specific backup
"""

from __future__ import annotations
import sys
from pathlib import Path
import openpyxl
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

DB_DIR = Path("/Users/ruimiguelneves/Library/Group Containers/UBF8T346G9.OneDriveSyncClientSuite/OneDrive.noindex/OneDrive/Documents")
DB_PATH = DB_DIR / "Squads_Data.xlsx"

COLS = ["ID", "Shirt", "Player", "Age", "Clubs", "Country", "Division", "Goals", "Games", "International", "Shirt Int", "Previous Club"]


def load_rows(p: Path) -> dict[int, tuple]:
    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb["Players"]
    rows = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows[row[0]] = row
    wb.close()
    return rows


def latest_backup() -> Path | None:
    candidates = sorted(
        DB_DIR.glob("Squads_Data.xlsx.bak_before_*"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def main():
    bak: Path
    if len(sys.argv) > 1:
        bak = DB_DIR / sys.argv[1]
        if not bak.exists():
            print(f"ERROR: backup {bak.name} not found in {DB_DIR}", file=sys.stderr)
            return 1
    else:
        b = latest_backup()
        if not b:
            print("No backups found.", file=sys.stderr)
            return 1
        bak = b
        print(f"Comparing against latest backup: {bak.name}\n")

    cur = load_rows(DB_PATH)
    old = load_rows(bak)

    added = [i for i in cur if i not in old]
    removed = [i for i in old if i not in cur]
    modified = [i for i in cur if i in old and cur[i] != old[i]]

    if not (added or removed or modified):
        print("No differences.")
        return 0

    if added:
        print(f"=== ADDED ({len(added)}) ===")
        for i in added[:20]:
            print(f"  + {cur[i][2]:30}  club={cur[i][4]}  div={cur[i][6]}")
        if len(added) > 20:
            print(f"  …+{len(added)-20} more")
        print()

    if removed:
        print(f"=== REMOVED ({len(removed)}) ===")
        for i in removed[:20]:
            print(f"  - {old[i][2]:30}  was-club={old[i][4]}  was-div={old[i][6]}")
        if len(removed) > 20:
            print(f"  …+{len(removed)-20} more")
        print()

    if modified:
        print(f"=== MODIFIED ({len(modified)}) ===")
        for i in modified[:30]:
            o = old[i]; c = cur[i]
            name = c[2] or o[2]
            diffs = []
            for ci, col in enumerate(COLS):
                if o[ci] != c[ci]:
                    diffs.append(f"{col}: {o[ci]!r} → {c[ci]!r}")
            print(f"  ~ {name}")
            for d in diffs:
                print(f"      {d}")
        if len(modified) > 30:
            print(f"  …+{len(modified)-30} more")

    return 0


if __name__ == "__main__":
    sys.exit(main())
