#!/usr/bin/env python3
"""
Pre-season xMin tracker.

Takes a friendly's lineup and cross-references every name against FPL 25/26
minutes/starts, so a pre-season goal from a fringe player never reads as signal.

Usage:
  python3 tools/preseason_xmin.py CRY "Swindon" 2026-07-18 \
      --xi "Benitez,King,Mingueza,Mitchell,Rak-Sakyi,Ozoh,Nketiah,Hughes,Walker-Smith,Adaramola,Esse" \
      --subs "Jemide,Matthews,Devenny,Cardines,Benamar,Canvot,Johnson,Matheus Franca,Drakes-Thomas,Sosa,Doucoure,Imray" \
      --scorers "Nketiah:2,Matheus Franca:3"

Appends to preseason_xmin.json and prints a triage table.
"""
import json
import os
import sys
import argparse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
STORE = os.path.join(ROOT, "preseason_xmin.json")
CACHE = os.path.join(HERE, ".fpl_bootstrap_cache.json")
BOOTSTRAP = "https://fantasy.premierleague.com/api/bootstrap-static/"

POS = {1: "GK", 2: "DEF", 3: "MID", 4: "FWD"}


def load_bootstrap(refresh=False):
    if not refresh and os.path.exists(CACHE):
        try:
            with open(CACHE) as f:
                return json.load(f)
        except Exception:
            pass
    with urllib.request.urlopen(BOOTSTRAP, timeout=30) as r:
        d = json.load(r)
    with open(CACHE, "w") as f:
        json.dump(d, f)
    return d


def norm(s):
    """Strip trailing initials/periods and lowercase for matching."""
    s = s.strip().rstrip(".")
    parts = [p for p in s.split() if not (len(p.rstrip(".")) == 1)]
    return " ".join(parts).lower() if parts else s.lower()


def classify(starts, mins, found):
    """Classify by STARTS first (more robust than raw minutes)."""
    if not found:
        return "🔴 NEW/ACADEMY", "no PL record — new signing or academy"
    if starts >= 28:
        return "🟢 NAILED", "first-choice starter"
    if starts >= 15:
        return "🟡 ROTATION", "squad rotation"
    if starts >= 3:
        return "🟠 FRINGE", "fringe — pre-season output is NOISE"
    return "🔴 MINIMAL", "academy/deep bench — ignore for FPL"


def find_player(elements, teams, name, club_id=None):
    n = norm(name)
    cands = []
    for e in elements:
        pool = [
            norm(e.get("second_name", "")),
            norm(e.get("web_name", "")),
            norm(f"{e.get('first_name','')} {e.get('second_name','')}"),
        ]
        if any(n == p for p in pool) or any(n in p and len(n) > 4 for p in pool):
            cands.append(e)
    if not cands:
        return None
    if club_id:
        same = [e for e in cands if e["team"] == club_id]
        if same:
            cands = same
    return max(cands, key=lambda e: e["minutes"])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("club")
    ap.add_argument("opponent")
    ap.add_argument("date")
    ap.add_argument("--xi", default="")
    ap.add_argument("--subs", default="")
    ap.add_argument("--scorers", default="")
    ap.add_argument("--refresh", action="store_true")
    a = ap.parse_args()

    d = load_bootstrap(a.refresh)
    elements, teams = d["elements"], {t["id"]: t["name"] for t in d["teams"]}

    # resolve club id loosely from the code/name
    club_id = None
    for t in d["teams"]:
        if a.club.lower() in t["name"].lower() or t["name"].lower().startswith(a.club.lower()[:3]):
            club_id = t["id"]
            break

    scorers = {}
    for tok in [s for s in a.scorers.split(",") if s.strip()]:
        nm, _, g = tok.partition(":")
        scorers[norm(nm)] = int(g or 1)

    rows = []
    for group, raw in (("XI", a.xi), ("SUB", a.subs)):
        for name in [x.strip() for x in raw.split(",") if x.strip()]:
            e = find_player(elements, teams, name, club_id)
            found = e is not None
            starts = e["starts"] if found else 0
            mins = e["minutes"] if found else 0
            tag, why = classify(starts, mins, found)
            rows.append({
                "name": name, "group": group,
                "pos": POS.get(e["element_type"]) if found else "?",
                "mins_2526": mins, "starts_2526": starts,
                "goals": scorers.get(norm(name), 0),
                "tag": tag, "note": why,
            })

    print(f"\n=== {a.club} vs {a.opponent} · {a.date} — pre-season xMin triage ===")
    print(f"{'player':20}{'grp':5}{'pos':5}{'mins':>6}{'st':>4}{'G':>3}  verdict")
    print("-" * 78)
    order = {"🟢 NAILED": 0, "🟡 ROTATION": 1, "🟠 FRINGE": 2, "🔴 MINIMAL": 3, "🔴 NEW/ACADEMY": 4}
    for r in sorted(rows, key=lambda r: (order.get(r["tag"], 9), -r["starts_2526"])):
        g = str(r["goals"]) if r["goals"] else "-"
        print(f"{r['name'][:19]:20}{r['group']:5}{r['pos']:5}{r['mins_2526']:>6}{r['starts_2526']:>4}{g:>3}  {r['tag']} · {r['note']}")

    nailed = [r for r in rows if r["tag"] == "🟢 NAILED"]
    noise = [r for r in rows if r["goals"] and r["tag"] in ("🟠 FRINGE", "🔴 MINIMAL", "🔴 NEW/ACADEMY")]
    print("-" * 78)
    print(f"Nailed first-teamers involved: {len(nailed)}/{len(rows)}"
          + (f" → {', '.join(r['name'] for r in nailed)}" if nailed else " → NONE (reserve run-out)"))
    if noise:
        print("⚠️  TRAP WATCH — goals from non-nailed players (ignore for FPL): "
              + ", ".join(f"{r['name']} ({r['goals']}g, {r['starts_2526']} starts)" for r in noise))

    store = []
    if os.path.exists(STORE):
        try:
            with open(STORE) as f:
                store = json.load(f)
        except Exception:
            store = []
    store.append({"club": a.club, "opponent": a.opponent, "date": a.date, "players": rows})
    with open(STORE, "w") as f:
        json.dump(store, f, indent=2, ensure_ascii=False)
    print(f"\nLogged → {STORE} ({len(store)} fixtures tracked)")


if __name__ == "__main__":
    sys.exit(main())
