#!/usr/bin/env python3
"""Weekly FPL Hall-of-Fame elite tracker (the '17-elite' process).

Each GW after the deadline it:
  1. Pulls the 17 verified elite managers' squads for the current GW
  2. Diffs vs the previous GW -> transfers-IN per manager, ownership swings,
     and flags EARLY CLUSTERS (2+ elites onto the same new name = the buy signal)
  3. Prints the current consensus template + the fixture-swing calendar (next 5 GWs)

Dollimore (leading indicator) and Crellin (deadline tripwire) are highlighted.
Snapshots (elite_snap_gwN.json) and the report (elite_scan_gwN.md) land beside this file.

Usage:  python3 elite_scan.py
"""
import json, os, urllib.request

BASE = "https://fantasy.premierleague.com/api"
HERE = os.path.dirname(os.path.abspath(__file__))
IDS_FILE = os.path.normpath(os.path.join(HERE, "..", "hof_manager_ids.json"))
POS = {1: "GK", 2: "DEF", 3: "MID", 4: "FWD"}
LEADING = {"Tom Dollimore": "🧭 leading-indicator", "Ben Crellin": "⏰ deadline-tripwire"}
WATCH = ["ARS", "LIV", "CHE", "MCI", "MUN", "TOT", "BHA", "NEW", "EVE"]


def get(path):
    req = urllib.request.Request(f"{BASE}/{path}", headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def main():
    boot = get("bootstrap-static/")
    els = {e["id"]: e for e in boot["elements"]}
    teams = {t["id"]: t["short_name"] for t in boot["teams"]}
    cur = next((e for e in boot["events"] if e["is_current"]),
               next((e for e in boot["events"] if e["is_next"]), boot["events"][0]))
    gw = cur["id"]
    ids = json.load(open(IDS_FILE))["verified"]

    squads, own = {}, {}
    for name, eid in ids.items():
        try:
            d = get(f"entry/{eid}/event/{gw}/picks/")
        except Exception as ex:
            print(f"  skip {name} ({eid}): {ex}")
            continue
        s = set(p["element"] for p in d["picks"])
        squads[name] = s
        for el in s:
            own[el] = own.get(el, 0) + 1
    N = len(squads)

    json.dump({"gw": gw, "squads": {k: sorted(v) for k, v in squads.items()},
               "own": {str(k): v for k, v in own.items()}},
              open(os.path.join(HERE, f"elite_snap_gw{gw}.json"), "w"))

    def tag(el):
        e = els[el]
        return f"{e['web_name']} ({POS[e['element_type']]} {teams[e['team']]}, £{e['now_cost']/10})"

    L = [f"# Elite Scan — GW{gw}   ({N}/17 managers)\n"]

    prev_path = os.path.join(HERE, f"elite_snap_gw{gw-1}.json")
    if os.path.exists(prev_path):
        prev = json.load(open(prev_path))
        pown = {int(k): v for k, v in prev["own"].items()}
        psq = {k: set(v) for k, v in prev["squads"].items()}
        moves = {}
        for name, s in squads.items():
            if name in psq:
                for el in s - psq[name]:
                    moves.setdefault(el, []).append(name)

        L.append("## 🔥 EARLY CLUSTERS — 2+ elites bought this GW (ACT)\n")
        clusters = sorted([(el, m) for el, m in moves.items() if len(m) >= 2], key=lambda x: -len(x[1]))
        L += [f"- **{tag(el)}** — +{len(m)}: {', '.join(m)} → now **{own.get(el,0)}/{N}**" for el, m in clusters] or ["- none this week"]

        L.append("\n## 👀 Single movers — watch (1 elite; act only if a fixture swing + a 2nd mover confirm)\n")
        singles = sorted([(el, m[0]) for el, m in moves.items() if len(m) == 1], key=lambda x: els[x[0]]["element_type"])
        L += [f"- {tag(el)} ← {who}{'  ⟵ ' + LEADING[who] if who in LEADING else ''}" for el, who in singles] or ["- none"]

        risers = sorted(((el, own.get(el, 0) - pown.get(el, 0)) for el in set(own) | set(pown)), key=lambda x: -x[1])
        L.append("\n## 📈 Ownership swings\n")
        L += [f"- {tag(el)}  {'+' if d>0 else ''}{d} → {own.get(el,0)}/{N}" for el, d in risers[:8] if d]
        L += [f"- {tag(el)}  {d} → {own.get(el,0)}/{N}" for el, d in risers[-5:] if d < 0]
    else:
        L.append("_Baseline GW — no prior snapshot to diff. Consensus + fixtures below; transfer deltas begin next GW._\n")

    L.append("\n## 🧭 Current consensus (owned by ≥ half of the elites)\n")
    for el, c in sorted(own.items(), key=lambda x: -x[1]):
        if c >= N / 2:
            L.append(f"- {tag(el)} — **{c}/{N}**")

    L.append("\n## 📅 Fixture swing — next 5 GWs (position early, Dollimore-style)\n")
    tid = {v: k for k, v in teams.items()}
    gws = [gw + i for i in range(1, 6)]
    grid = {c: {} for c in WATCH}
    for g in gws:
        try:
            fx = get(f"fixtures/?event={g}")
        except Exception:
            continue
        for f in fx:
            h, a = teams[f["team_h"]], teams[f["team_a"]]
            if h in WATCH:
                grid[h][g] = (f"{a}(H)", f["team_h_difficulty"])
            if a in WATCH:
                grid[a][g] = (f"{h}(A)", f["team_a_difficulty"])
    L.append("| Club | " + " | ".join(f"GW{g}" for g in gws) + " | ΣFDR |")
    L.append("|" + "---|" * (len(gws) + 2))
    rows = []
    for c in WATCH:
        cells, s = [], 0
        for g in gws:
            if g in grid[c]:
                opp, fdr = grid[c][g]
                cells.append(f"{opp}{fdr}")
                s += fdr
            else:
                cells.append("-")
        rows.append((c, cells, s))
    for c, cells, s in sorted(rows, key=lambda x: x[2]):
        L.append(f"| {c} | " + " | ".join(cells) + f" | **{s}** |")
    L.append("\n_Lowest ΣFDR = best upcoming run. Cross with the clusters above: an early mover onto a club whose run turns green = the buy._")

    out = os.path.join(HERE, f"elite_scan_gw{gw}.md")
    open(out, "w").write("\n".join(L))
    print(f"Wrote {out}   ({N}/17 managers, GW{gw})")


if __name__ == "__main__":
    main()
