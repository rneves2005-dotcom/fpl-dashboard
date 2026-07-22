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
import unicodedata
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


def strip_accents(s):
    """Milenković -> Milenkovic, Sangaré -> Sangare."""
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def norm(s):
    """Strip accents + trailing initials/periods, lowercase, for matching."""
    s = strip_accents(s).strip().rstrip(".")
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
    """Resolve a name to an FPL player.

    STRICT on club: if club_id is known we ONLY match within that club. Matching
    across clubs silently produced false positives (e.g. Forest's Omar Richards
    resolving to Palace's Chris Richards, 0 mins vs 2825) — never do that again.
    """
    n = norm(name)

    def scan(pool_elements):
        out = []
        for e in pool_elements:
            pool = [
                norm(e.get("second_name", "")),
                norm(e.get("web_name", "")),
                norm(f"{e.get('first_name','')} {e.get('second_name','')}"),
            ]
            if any(n == p for p in pool) or any(n in p and len(n) > 4 for p in pool):
                out.append(e)
        return out

    if club_id:
        cands = scan([e for e in elements if e["team"] == club_id])
        if cands:
            return max(cands, key=lambda e: e["minutes"]), False
        # Same-club surname fallback: the graphic's first name may differ from
        # FPL's ("Dan Ballard" vs "Daniel Ballard"), so the full-name scan misses
        # and the player is wrongly tagged NEW/ACADEMY. Within ONE club a bare
        # surname is low-collision — retry the last token against this squad only
        # (max-by-minutes picks the established player if several share it).
        surname = n.split()[-1]
        if len(surname) > 3:
            sc = [e for e in elements if e["team"] == club_id
                  and (norm(e.get("web_name", "")) == surname
                       or surname in norm(e.get("second_name", "")).split())]
            if sc:
                return max(sc, key=lambda e: e["minutes"]), False
        # No one at this club matches. During a transfer window the player may be
        # a new signing still listed at his OLD club in the 25/26 dataset, so fall
        # back league-wide — but FLAG it so it's never read as a same-club fact.
        #
        # GUARD: only attempt this for MULTI-TOKEN names ("Andrey Santos"). A bare
        # surname is far too ambiguous across 20 squads — "Williams" wrongly
        # resolved to Forest's Neco Williams when the real player was a United
        # academy kid. Single-surname misses stay unresolved by design.
        if len(n.split()) < 2:
            return None, False
        cands = [e for e in scan(elements)
                 if norm(f"{e.get('first_name','')} {e.get('second_name','')}") == n
                 or norm(e.get("web_name", "")) == n]
        if cands:
            return max(cands, key=lambda e: e["minutes"]), True
        return None, False

    cands = scan(elements)
    return (max(cands, key=lambda e: e["minutes"]), False) if cands else (None, False)


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

    # Resolve club via FPL's own short_name (CRY/EVE/NFO/...) — fuzzy name
    # matching silently failed for "NFO" vs "Nott'm Forest", which disabled the
    # club filter entirely and produced cross-club false matches.
    club_id = None
    for t in d["teams"]:
        if t.get("short_name", "").lower() == a.club.lower():
            club_id = t["id"]
            break
    if club_id is None:
        for t in d["teams"]:
            if a.club.lower() in t["name"].lower():
                club_id = t["id"]
                break
    if club_id is None:
        codes = ", ".join(sorted(t.get("short_name", "") for t in d["teams"]))
        sys.exit(f"ERROR: club '{a.club}' not recognised. Valid codes: {codes}")
    print(f"[club resolved: {teams[club_id]} — matching restricted to this squad]")

    scorers = {}
    for tok in [s for s in a.scorers.split(",") if s.strip()]:
        nm, _, g = tok.partition(":")
        scorers[norm(nm)] = int(g or 1)

    rows = []
    for group, raw in (("XI", a.xi), ("SUB", a.subs)):
        for name in [x.strip() for x in raw.split(",") if x.strip()]:
            e, other_club = find_player(elements, teams, name, club_id)
            found = e is not None
            starts = e["starts"] if found else 0
            mins = e["minutes"] if found else 0
            tag, why = classify(starts, mins, found)
            if found and other_club:
                why = f"⚠️ NEW SIGNING — 25/26 stats are from {teams[e['team']]}"
            rows.append({
                "name": name, "group": group,
                "pos": POS.get(e["element_type"]) if found else "?",
                "mins_2526": mins, "starts_2526": starts,
                "goals": scorers.get(norm(name), 0),
                "tag": tag, "note": why,
                "from_other_club": teams[e["team"]] if (found and other_club) else None,
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
