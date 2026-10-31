#!/usr/bin/env python3
"""Weekly FPL elite tracker — the 27 ("Mag 27" = 17 HoF + 10 strong10_last_year).

Each GW after the deadline it:
  1. Pulls all 27 managers' GW picks (detecting Free Hit) + their GW transfers.
  2. 27-MANAGER TRANSFER BOARD — consensus BUYS / SELLS + EARLY CLUSTERS (2+ onto
     the same new name = the buy signal), each split [HoF / strong10].
  3. Current consensus template (owned by >= half of the 27; Free-Hit squads are
     unwound to the manager's prior-GW team so a 1-week rental doesn't distort it).
  4. Fixture-swing calendar (next 5 GWs) to cross against the clusters.

Dollimore (leading indicator) + Crellin (deadline tripwire) are flagged among buyers.
The 10 are the FIXED hof_manager_ids.json["strong10_last_year"] set — NOT the live
league-940453 standings (those fluctuate). Transfers are only visible AFTER the
deadline locks: this is a LAGGING view of the current GW, a LEADING signal for next.

Snapshots (elite_snap_gwN.json) + report (elite_scan_gwN.md) land beside this file.

Usage:  python3 elite_scan.py
"""
import json, os, urllib.request

BASE = "https://fantasy.premierleague.com/api"
HERE = os.path.dirname(os.path.abspath(__file__))
IDS_FILE = os.path.normpath(os.path.join(HERE, "..", "hof_manager_ids.json"))
POS = {1: "GK", 2: "DEF", 3: "MID", 4: "FWD"}
LEADING = {179777: "🧭 Dollimore (leading)", 53517: "⏰ Crellin (deadline)"}
WATCH = ["ARS", "LIV", "CHE", "MCI", "MUN", "TOT", "BHA", "NEW", "EVE"]


def get(path):
    req = urllib.request.Request(f"{BASE}/{path}", headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def load_managers():
    d = json.load(open(IDS_FILE))
    hof = d["verified"]                      # 17
    strong = {k: v for k, v in d["strong10_last_year"].items() if not k.startswith("_")}  # 10
    mgrs = {}
    for n, i in hof.items():
        mgrs[i] = {"name": n, "group": "HoF"}
    for n, i in strong.items():
        mgrs.setdefault(i, {"name": n, "group": "S10"})  # don't double-count an overlap
    return mgrs, len(hof), len(strong)


def main():
    boot = get("bootstrap-static/")
    els = {e["id"]: e for e in boot["elements"]}
    teams = {t["id"]: t["short_name"] for t in boot["teams"]}
    cur = next((e for e in boot["events"] if e["is_current"]),
               next((e for e in boot["events"] if e["is_next"]), boot["events"][0]))
    gw = cur["id"]
    mgrs, n_hof, n_s10 = load_managers()

    squads, own = {}, {}
    buys, sells = {}, {}          # element_id -> list of entry_ids
    fh = []                        # managers on Free Hit this GW
    done = 0
    for eid, meta in mgrs.items():
        try:
            picks = get(f"entry/{eid}/event/{gw}/picks/")
        except Exception as ex:
            print(f"  skip picks {meta['name']} ({eid}): {ex}")
            picks = None
        if picks is not None:
            durable = picks
            if picks.get("active_chip") == "freehit":
                fh.append(eid)
                try:                                   # unwind the 1-week FH rental
                    durable = get(f"entry/{eid}/event/{gw-1}/picks/")
                except Exception:
                    durable = picks
            s = set(p["element"] for p in durable["picks"])
            squads[eid] = s
            for el in s:
                own[el] = own.get(el, 0) + 1
            done += 1
        try:
            tr = [t for t in get(f"entry/{eid}/transfers/") if t["event"] == gw]
        except Exception as ex:
            print(f"  skip transfers {meta['name']} ({eid}): {ex}")
            tr = []
        for t in tr:
            buys.setdefault(t["element_in"], []).append(eid)
            sells.setdefault(t["element_out"], []).append(eid)

    N = len(mgrs)
    json.dump({"gw": gw, "n": N, "members": sorted(mgrs.keys()),
               "groups": {str(i): m["group"] for i, m in mgrs.items()},
               "squads": {str(i): sorted(v) for i, v in squads.items()},
               "own": {str(k): v for k, v in own.items()}},
              open(os.path.join(HERE, f"elite_snap_gw{gw}.json"), "w"))

    def tag(el):
        e = els[el]
        return f"{e['web_name']} ({POS[e['element_type']]} {teams[e['team']]}, £{e['now_cost']/10})"

    def split(ids):
        h = sum(1 for i in ids if mgrs[i]["group"] == "HoF")
        return f"{h} HoF / {len(ids)-h} S10"

    def leaders(ids):
        fl = [LEADING[i] for i in ids if i in LEADING]
        return ("  ⟵ " + ", ".join(fl)) if fl else ""

    moved = len(set([i for v in buys.values() for i in v]))
    L = [f"# Elite Scan — GW{gw}   (the 27: {n_hof} HoF + {n_s10} strong10)\n",
         f"_{moved}/{N} managers made at least one GW{gw} transfer"
         + (f" · {len(fh)} on Free Hit (squad unwound to GW{gw-1})" if fh else "") + "._\n"]

    L.append("## 🔥 Early clusters — 2+ of the 27 bought this GW (the buy signal)\n")
    clusters = sorted([(el, m) for el, m in buys.items() if len(m) >= 2], key=lambda x: -len(x[1]))
    L += [f"- **{tag(el)}** — **{len(m)}×** [{split(m)}] → now owned by {own.get(el,0)}/{N}{leaders(m)}"
          for el, m in clusters] or ["- none this week"]

    L.append("\n## 💸 Consensus sells\n")
    sell_sorted = sorted(sells.items(), key=lambda x: -len(x[1]))
    L += [f"- {tag(el)} — **{len(m)}×** [{split(m)}]" for el, m in sell_sorted[:10] if len(m) >= 2] \
         or ["- none with 2+ sellers"]

    L.append("\n## 👀 Single buys — watch (1 manager; act only on a fixture swing + a 2nd mover)\n")
    singles = sorted([(el, m[0]) for el, m in buys.items() if len(m) == 1],
                     key=lambda x: els[x[0]]["element_type"])
    L += [f"- {tag(el)} ← {mgrs[who]['name']} [{mgrs[who]['group']}]{leaders([who])}" for el, who in singles] or ["- none"]

    # ownership swings only vs a prior snapshot with the SAME membership (else denominators differ)
    prev_path = os.path.join(HERE, f"elite_snap_gw{gw-1}.json")
    if os.path.exists(prev_path):
        prev = json.load(open(prev_path))
        if prev.get("n") == N and set(prev.get("members", [])) == set(mgrs.keys()):
            pown = {int(k): v for k, v in prev["own"].items()}
            risers = sorted(((el, own.get(el, 0) - pown.get(el, 0)) for el in set(own) | set(pown)),
                            key=lambda x: -x[1])
            L.append("\n## 📈 Ownership swings (durable, FH-unwound)\n")
            L += [f"- {tag(el)}  +{d} → {own.get(el,0)}/{N}" for el, d in risers[:8] if d > 0]
            L += [f"- {tag(el)}  {d} → {own.get(el,0)}/{N}" for el, d in risers[-5:] if d < 0]

    L.append("\n## 🧭 Current consensus template (owned by ≥ half of the 27)\n")
    for el, c in sorted(own.items(), key=lambda x: -x[1]):
        if c >= N / 2:
            hofc = sum(1 for i in squads if el in squads[i] and mgrs[i]["group"] == "HoF")
            L.append(f"- {tag(el)} — **{c}/{N}** ({hofc}/{n_hof} HoF)")

    L.append("\n## 📅 Fixture swing — next 5 GWs (position early, Dollimore-style)\n")
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
    print(f"Wrote {out}   ({done}/{N} managers scanned, GW{gw}, {len(clusters)} clusters)")


if __name__ == "__main__":
    main()
