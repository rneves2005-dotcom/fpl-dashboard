#!/usr/bin/env python3
"""Re-export Squads_Data.xlsx → world_data.json (no transfer applied).

Use this when:
- Rui edited the xlsx directly in Excel (added players, renamed clubs, fixed
  data) and wants those changes reflected in world.html / stats.html
- Claude just opened the project and wants to make sure the JSON matches the
  current xlsx state before doing anything else

Workflow:
1. python3 tools/sync.py            # regenerate JSON
2. python3 tools/validate_overlays.py   # ensure overlay club names still match
3. git add world_data.json && git commit -m "Sync from xlsx" && git push

No xlsx mutation. No backup. Pure read → JSON write.
"""

from __future__ import annotations
import sys
from pathlib import Path

# Reuse the regenerator from apply_transfer (single source of truth for the logic).
sys.path.insert(0, str(Path(__file__).parent))
from apply_transfer import regenerate_world_json, DB_PATH


def main():
    if not DB_PATH.exists():
        print(f"ERROR: DB not found at {DB_PATH}", file=sys.stderr)
        return 1
    print(f"Reading {DB_PATH.name}…")
    regenerate_world_json()
    return 0


if __name__ == "__main__":
    sys.exit(main())
