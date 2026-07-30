#!/usr/bin/env python3
"""Repopulate transfers_in.json / transfers_out.json from the poller's OFFICIAL
transfer entries in Transfer_Tracker.xlsx — the pipeline that froze mid-June.

The overlays drive squads.html's per-club "Transfers IN/OUT" panels. They were
hand-curated through June, then stopped. The poller keeps logging club-announced
OFFICIAL moves; this maps those (PL clubs only) into the overlay schema so the
dashboard reflects reality — sourced from club statements, not memory.

Schema: {"_meta": str, "Premier League": {"<Club>": [{"n","from"/"to","note","c"}]}}

Usage:
  python3 tools/refresh_overlays.py --dry-run   # preview counts + samples, NO writes
  python3 tools/refresh_overlays.py             # apply (then run normalize_transfers.py)
"""
from __future__ import annotations
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "Transfer_Tracker.xlsx"
TIN = ROOT / "transfers_in.json"
TOUT = ROOT / "transfers_out.json"

# poller/PL-Official club spellings -> overlay canonical keys (20 PL clubs only)
CLUB_MAP = {
    "arsenal": "Arsenal", "aston villa": "Aston Villa",
    "afc bournemouth": "Bournemouth", "bournemouth": "Bournemouth",
    "brentford": "Brentford",
    "brighton & hove albion": "Brighton & Hove Albion F.C.", "brighton": "Brighton & Hove Albion F.C.",
    "chelsea": "Chelsea", "coventry city": "Coventry City", "coventry": "Coventry City",
    "crystal palace": "Crystal Palace F.C", "everton": "Everton", "fulham": "Fulham",
    "hull city": "Hull City A.F.C.", "hull": "Hull City A.F.C.",
    "ipswich town": "Ipswich Town", "ipswich": "Ipswich Town",
    "leeds united": "Leeds United", "leeds": "Leeds United",
    "liverpool": "Liverpool",
    "man city": "Manchester City", "manchester city": "Manchester City",
    "man utd": "Manchester United", "manchester united": "Manchester United", "man united": "Manchester United",
    "newcastle united": "Newcastle United", "newcastle": "Newcastle United",
    "nottingham forest": "Nottingham Forest", "nott'm forest": "Nottingham Forest",
    "sunderland": "Sunderland A.F.C.",
    "tottenham hotspur": "Tottenham Hotspur", "tottenham": "Tottenham Hotspur", "spurs": "Tottenham Hotspur",
}

# OFFICIAL: <club> · <player> — Transfer In|Out[ · <counterparty>]
PAT = re.compile(r"^OFFICIAL:\s*(.+?)\s*·\s*(.+?)\s*[—-]\s*Transfer\s+(In|Out)\b(?:\s*·\s*-?\s*(.+))?", re.I)


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip().lower()


def official_rows():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    seen = set()
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for r in rows:
            when = r[1] if len(r) > 1 else ""
            for c in r:
                if isinstance(c, str) and c.startswith("OFFICIAL:") and "Transfer" in c:
                    key = c.strip()
                    if key in seen:
                        continue
                    seen.add(key)
                    yield str(when), c.strip()
    wb.close()


def parse(title: str):
    m = PAT.match(title)
    if not m:
        return None
    club_raw, player, direction, counter = m.group(1), m.group(2), m.group(3), m.group(4)
    club = CLUB_MAP.get(norm(club_raw))
    if not club:
        return None  # not a PL club
    return club, player.strip(), direction.capitalize(), (counter or "").strip(" -·")


def build():
    tin = json.load(open(TIN)); tout = json.load(open(TOUT))
    pl_in = tin.setdefault("Premier League", {})
    pl_out = tout.setdefault("Premier League", {})

    def existing_names(store, club):
        return {norm(e.get("n", "")) for e in store.get(club, [])}

    added = {"In": {}, "Out": {}}
    for when, title in official_rows():
        p = parse(title)
        if not p:
            continue
        club, player, direction, counter = p
        store = pl_in if direction == "In" else pl_out
        store.setdefault(club, [])
        if norm(player) in existing_names(store, club):
            continue
        entry = {"n": player}
        if direction == "In":
            entry["from"] = counter or "?"
        else:
            entry["to"] = counter or "?"
        entry["note"] = f"OFFICIAL (club statement){' · ' + when[:10] if when else ''}"
        entry["c"] = ""
        store[club].append(entry)
        added[direction].setdefault(club, []).append(f"{player}" + (f" ({counter})" if counter else ""))
    return tin, tout, added


def main():
    dry = "--dry-run" in sys.argv
    tin, tout, added = build()
    tot_in = sum(len(v) for v in added["In"].values())
    tot_out = sum(len(v) for v in added["Out"].values())
    print(f"=== refresh_overlays {'(DRY-RUN)' if dry else '(APPLY)'} — new PL entries: {tot_in} IN, {tot_out} OUT ===\n")
    for direction in ("In", "Out"):
        print(f"--- Transfers {direction} ---")
        for club in sorted(added[direction]):
            names = added[direction][club]
            print(f"  {club} (+{len(names)}): " + "; ".join(names[:8]) + (" …" if len(names) > 8 else ""))
        print()
    if dry:
        print("DRY-RUN — nothing written. Re-run without --dry-run to apply.")
        return 0
    json.dump(tin, open(TIN, "w"), indent=2, ensure_ascii=False)
    json.dump(tout, open(TOUT, "w"), indent=2, ensure_ascii=False)
    print(f"Wrote {TIN.name} + {TOUT.name}. Now run: python3 tools/normalize_transfers.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
