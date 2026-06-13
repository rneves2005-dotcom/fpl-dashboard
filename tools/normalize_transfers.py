#!/usr/bin/env python3
"""Normalize transfers_in.json + transfers_out.json against DB canonical club + division names.

Strategy:
  1. Load world_data.json — source of truth for division + club names (from Squads_Data.xlsx).
  2. Load tools/club_aliases.json — manual alias map (e.g. "Sporting CP" → "Sporting").
  3. For each entry in the overlays:
     - Normalize the division name (alias → canonical) and verify it exists in world_data
     - Normalize the club name within that division (alias → canonical)
     - Merge entries if multiple alias keys collapse to the same canonical club
     - Normalize 'from'/'to' string fields where the leading token matches a known club
  4. Write back normalized overlays.
  5. Warn (don't fail) when a club is NOT in world_data — flags potential DB-overlay drift.

Idempotent — running twice produces no changes. Run after editing overlays manually
or as a post-step in poll_transfers.py for new entries.

Usage:
  python3 tools/normalize_transfers.py            # apply + report
  python3 tools/normalize_transfers.py --dry-run  # report only, no writes
"""

from __future__ import annotations
import json
import sys
import re
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
WORLD_DATA = ROOT / "world_data.json"
ALIASES = ROOT / "tools" / "club_aliases.json"
TRANSFERS_IN = ROOT / "transfers_in.json"
TRANSFERS_OUT = ROOT / "transfers_out.json"

def load_json(p: Path) -> Any:
    with open(p) as f:
        return json.load(f)

def save_json(p: Path, data: Any):
    with open(p, "w") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def normalize_division(name: str, div_aliases: dict, world_data: dict) -> str:
    """Map division alias → canonical. Pass-through if already canonical."""
    if name in world_data:
        return name
    return div_aliases.get(name, name)

def normalize_club(div: str, club: str, club_aliases: dict, world_data: dict) -> str:
    """Map club alias → canonical within division. Falls back to original if no alias."""
    div_map = club_aliases.get(div, {})
    canonical = div_map.get(club, club)
    # Verify canonical exists in DB
    if canonical not in world_data.get(div, {}):
        # Cross-division check — maybe the entry belongs to a different division
        pass
    return canonical

def normalize_fromto(value: str, club_aliases: dict) -> str:
    """For 'from'/'to' free-text fields, swap any known club alias with its canonical name.
    Skip when alias is a prefix/substring of canonical AND canonical already appears
    (prevents "Tottenham" → "Tottenham Hotspur" from doubling "Tottenham Hotspur" → "Tottenham Hotspur Hotspur").
    """
    if not value or value == "—":
        return value
    out = value
    for div, mapping in club_aliases.items():
        if div.startswith("_"):
            continue
        # Sort longest-alias-first so multi-word aliases match before substring aliases
        for alias, canonical in sorted(mapping.items(), key=lambda kv: -len(kv[0])):
            if alias == canonical:
                continue
            # Idempotency guard: if canonical already in text, don't replace alias (it's already canonical)
            if canonical in out:
                continue
            # Only standalone tokens
            pat = r"\b" + re.escape(alias) + r"\b"
            out = re.sub(pat, canonical, out)
    return out

def normalize_overlay(overlay: dict, world_data: dict, aliases: dict, label: str, dry_run: bool) -> tuple[dict, list, list]:
    """Returns (normalized_overlay, warnings, changes_log)."""
    div_aliases = aliases.get("_divisions", {})
    club_aliases = aliases.get("_clubs", {})

    warnings = []
    changes = []
    new_overlay: dict = {}

    for div, clubs in list(overlay.items()):
        if div.startswith("_") or not isinstance(clubs, dict):
            new_overlay[div] = clubs
            continue

        new_div = normalize_division(div, div_aliases, world_data)
        if new_div != div:
            changes.append(f"DIV  {label}: '{div}' → '{new_div}'")

        if new_div not in new_overlay:
            new_overlay[new_div] = {}

        for club, entries in clubs.items():
            new_club = normalize_club(new_div, club, club_aliases, world_data)
            # Cross-division correction (e.g. PL → Championship for relegated)
            if new_club not in world_data.get(new_div, {}):
                # Try every other division in case the club moved
                for alt_div in world_data:
                    if alt_div == new_div:
                        continue
                    if new_club in world_data[alt_div]:
                        warnings.append(f"  ⚠️  '{new_club}' not in {new_div} but found in {alt_div} · keeping in {new_div} but flag for review")
                        break
                else:
                    if not new_club.startswith("🪦"):  # skip retirement markers
                        warnings.append(f"  ⚠️  '{new_club}' (in {new_div}) is NOT in world_data — check spelling or add to aliases")

            if new_club != club:
                changes.append(f"CLUB {label}: '{div}/{club}' → '{new_div}/{new_club}'")

            # Normalize the from/to strings inside each entry too
            normalized_entries = []
            for e in entries:
                e2 = dict(e)
                if "from" in e2:
                    e2["from"] = normalize_fromto(e2["from"], club_aliases)
                if "to" in e2:
                    e2["to"] = normalize_fromto(e2["to"], club_aliases)
                normalized_entries.append(e2)

            # Merge if canonical key already exists (alias collision)
            existing = new_overlay[new_div].get(new_club, [])
            # Dedup by player name within merged list
            seen_names = {x.get("n") for x in existing}
            for e in normalized_entries:
                if e.get("n") in seen_names:
                    continue
                existing.append(e)
                seen_names.add(e.get("n"))
            new_overlay[new_div][new_club] = existing

    return new_overlay, warnings, changes

def main():
    dry_run = "--dry-run" in sys.argv

    world_data = load_json(WORLD_DATA)
    aliases = load_json(ALIASES)
    ti = load_json(TRANSFERS_IN)
    to = load_json(TRANSFERS_OUT)

    new_ti, warn_in, changes_in = normalize_overlay(ti, world_data, aliases, "IN", dry_run)
    new_to, warn_out, changes_out = normalize_overlay(to, world_data, aliases, "OUT", dry_run)

    all_changes = changes_in + changes_out
    all_warnings = warn_in + warn_out

    print(f"=== Normalization {'(DRY RUN)' if dry_run else 'APPLIED'} ===")
    if all_changes:
        print(f"\n{len(all_changes)} canonicalisations:")
        for c in all_changes:
            print(f"  ✓ {c}")
    else:
        print("\n(no changes needed — overlays already canonical)")

    if all_warnings:
        print(f"\n{len(all_warnings)} warnings (clubs not in DB):")
        for w in all_warnings:
            print(w)

    if not dry_run and all_changes:
        save_json(TRANSFERS_IN, new_ti)
        save_json(TRANSFERS_OUT, new_to)
        print(f"\n✅ Wrote {TRANSFERS_IN.name} + {TRANSFERS_OUT.name}")
    elif not dry_run:
        print("\n(no writes — overlays unchanged)")
    else:
        print("\n(dry run · no writes)")

if __name__ == "__main__":
    main()
