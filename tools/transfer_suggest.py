#!/usr/bin/env python3
"""Weekly transfer suggester — standalone port of the Process board (process.html) Section 1.

Reproduces the exact "My Squad · Weakest Link" method:
  ppg = minutes gate (recent-4-GW mins) x per-90 output (xG/xA + CS + DefCon) x fixture ease,
  pens COUNTED as real points, availability + curated minutes-risk haircut.
The CUT = the STARTER with the biggest gap to its best affordable upgrade (classic) /
best free agent (draft). Horizon = next 5, season basis, opponent+venue weighted — the page defaults.

Classic = entry 37889. Draft = league 6079, Rui = entry_id 27317 (owner==27317 his squad, owner==null FA).
Run locally (hits FPL + draft APIs directly, no proxy needed). Report-only — suggests, never executes.

Usage:  python3 tools/transfer_suggest.py
"""
import urllib.request, json, math, sys

UA = "Mozilla/5.0"
FPL = "https://fantasy.premierleague.com/api"
DRAFT = "https://draft.premierleague.com/api"
CLASSIC_ENTRY = 37889
DRAFT_LEAGUE = 6079
DRAFT_ENTRY_ID = 27317
HORIZON = 5
POS = {1: "GK", 2: "DEF", 3: "MID", 4: "FWD"}
MINUTES_RISK = {
    "Konsa": "Saliba returning — may lose the Arsenal CB spot (or shift to RB). Watch the teamsheet.",
}


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=40) as r:
        return json.load(r)


def clamp(v, lo, hi):
    return max(lo, min(v, hi))


def pen_taker(e):
    po = e.get("penalties_order")
    return po is not None and po <= 2


def build_model():
    boot = get(f"{FPL}/bootstrap-static/")
    teams = {t["id"]: {"id": t["id"], "name": t["short_name"], "full": t["name"]} for t in boot["teams"]}
    els = boot["elements"]
    by_id = {e["id"]: e for e in els}
    by_code = {e["code"]: e for e in els}
    pteam = {e["id"]: e["team"] for e in els}

    fixtures = get(f"{FPL}/fixtures/")
    played_gws = sorted({f["event"] for f in fixtures if f["event"] and (f["finished"] or f["started"])})

    # team xG per gw + player minutes per gw
    txg, pmin = {}, {}
    for gw in played_gws:
        try:
            d = get(f"{FPL}/event/{gw}/live/")
        except Exception:
            continue
        m = {}
        for x in d["elements"]:
            t = pteam.get(x["id"])
            mn = float(x["stats"].get("minutes", 0) or 0)
            if t:
                m[t] = m.get(t, 0) + float(x["stats"].get("expected_goals", 0) or 0)
            pmin.setdefault(x["id"], []).append((gw, mn))
        txg[gw] = m

    # recent minutes: last 4 GW average (current role, not season-avg dragged by old weeks)
    rec_mpg = {}
    for pid, arr in pmin.items():
        a = sorted(arr)[-4:]
        rec_mpg[pid] = sum(r[1] for r in a) / (len(a) or 1)

    # per-team season model from fixtures
    T = {tid: {**teams[tid], "pl": 0, "xgf": 0.0, "xga": 0.0, "next": []} for tid in teams}
    played = sorted([f for f in fixtures if (f["finished"] or f["started"]) and f["event"]], key=lambda x: x["event"])
    for f in played:
        gw, hg, ag = f["event"], f["team_h_score"], f["team_a_score"]
        if hg is None or ag is None:
            continue
        hx = (txg.get(gw) or {}).get(f["team_h"], 0)
        ax = (txg.get(gw) or {}).get(f["team_a"], 0)
        for tid, xf, xa in [(f["team_h"], hx, ax), (f["team_a"], ax, hx)]:
            o = T[tid]
            o["pl"] += 1
            o["xgf"] += xf
            o["xga"] += xa
    for o in T.values():
        o["xgfpg"] = o["xgf"] / o["pl"] if o["pl"] else 0
        o["xgapg"] = o["xga"] / o["pl"] if o["pl"] else 0

    # next 5 fixtures
    upcoming = sorted([f for f in fixtures if not f["finished"] and not f["started"] and f["event"]], key=lambda x: x["event"])
    for f in upcoming:
        if len(T[f["team_h"]]["next"]) < 5:
            T[f["team_h"]]["next"].append({"opp": f["team_a"], "home": True})
        if len(T[f["team_a"]]["next"]) < 5:
            T[f["team_a"]]["next"].append({"opp": f["team_h"], "home": False})

    vals = list(T.values())
    avg_xgf = sum(o["xgfpg"] for o in vals) / len(vals)
    avg_xga = sum(o["xgapg"] for o in vals) / len(vals)

    # projections (season basis, opponent + venue weighted)
    for o in vals:
        nx = o["next"][:HORIZON]
        n = len(nx) or 1
        pa = pd = 0.0
        for fx in nx:
            op = T[fx["opp"]]
            vA = 1.08 if fx["home"] else 0.92
            vD = 0.92 if fx["home"] else 1.08
            pa += o["xgfpg"] * (op["xgapg"] / avg_xga) * vA
            pd += o["xgapg"] * (op["xgfpg"] / avg_xgf) * vD
        o["projAtt"] = pa / n
        o["projDef"] = pd / n
        o["fixAttMult"] = clamp(o["projAtt"] / (o["xgfpg"] or 1), 0.5, 1.7)

    cur_event = next((e for e in boot["events"] if e["is_current"]), None)
    cur_gw = cur_event["id"] if cur_event else (played_gws[-1] if played_gws else 1)

    return {"T": T, "teams": teams, "by_id": by_id, "by_code": by_code, "elements": els,
            "rec_mpg": rec_mpg, "cur_gw": cur_gw, "last_played": played_gws[-1] if played_gws else None}


def exp_ppg(M, e):
    tm = M["T"].get(e["team"])
    if not tm or not e["minutes"]:
        return {"ppg": 0.0, "mpg": 0, "avail": 0, "risk": None, "flag": "no mins"}
    gp = tm["pl"] or 1
    rec = M["rec_mpg"].get(e["id"])
    mpg = min(rec, 90) if rec is not None else min(e["minutes"] / gp, 90)
    play_share = mpg / 90
    app = min(mpg / 60, 1) * 2
    per90 = lambda k: float(e.get(k, 0) or 0) / e["minutes"] * 90
    xg90, xa90 = per90("expected_goals"), per90("expected_assists")
    gv = {1: 6, 2: 6, 3: 5, 4: 4}[e["element_type"]]
    cs_val = {1: 4, 2: 4, 3: 1, 4: 0}[e["element_type"]]
    att_mult = tm.get("fixAttMult", 1)
    p_cs = math.exp(-(tm.get("projDef") or tm.get("xgapg") or 1.4))
    dc90 = float(e.get("defensive_contribution_per_90", 0) or 0)
    if e["element_type"] == 2:
        p_defcon = clamp((dc90 - 2) / 10, 0, 0.85)
    elif e["element_type"] == 3:
        p_defcon = clamp((dc90 - 4) / 14, 0, 0.5)
    else:
        p_defcon = 0
    cop = e.get("chance_of_playing_next_round")
    avail = cop / 100 if cop is not None else (1 if e["status"] == "a" else 0.75 if e["status"] == "d" else 0.2)
    attack_pts = (xg90 * gv + xa90 * 3) * att_mult * play_share
    cs_pts = cs_val * p_cs * min(mpg / 70, 1)
    dc_pts = 2 * p_defcon * play_share
    ppg = (app + attack_pts + cs_pts + dc_pts) * avail
    risk = MINUTES_RISK.get(e["web_name"])
    if risk:
        ppg *= 0.85
    flag = "mins" if risk else ("fit?" if avail < 0.6 else "")
    return {"ppg": ppg, "mpg": mpg, "avail": avail, "risk": risk, "flag": flag}


def best_upgrade(M, own, pool, budget):
    cur = exp_ppg(M, own)["ppg"]
    best = None
    for c in pool:
        if c["element_type"] != own["element_type"] or c["id"] == own["id"]:
            continue
        if budget is not None and c["now_cost"] > budget:
            continue
        av = c.get("chance_of_playing_next_round")
        avail = av / 100 if av is not None else (1 if c["status"] == "a" else 0.3)
        if avail < 0.75:
            continue
        tm = M["T"].get(c["team"])
        if not tm or not c["minutes"]:
            continue
        rec = M["rec_mpg"].get(c["id"])
        mpg_c = rec if rec is not None else c["minutes"] / (tm["pl"] or 1)
        if mpg_c < 60:
            continue
        p = exp_ppg(M, c)["ppg"]
        if p > cur and (best is None or p > best["ppg"]):
            best = {"e": c, "ppg": p}
    if best:
        return {**best, "gap": best["ppg"] - cur, "pen": pen_taker(best["e"])}
    return None


def load_squads(M):
    sq = {"classic": [], "draft": [], "bank": 0, "fa": [], "classicErr": None, "draftErr": None}
    try:
        pk = get(f"{FPL}/entry/{CLASSIC_ENTRY}/event/{M['cur_gw']}/picks/")
        sq["bank"] = (pk.get("entry_history", {}).get("bank", 0)) / 10
        sq["classic"] = [{"e": M["by_id"][p["element"]], "isC": p["is_captain"], "bench": p["position"] > 11}
                         for p in pk["picks"] if p["element"] in M["by_id"]]
    except Exception as ex:
        sq["classicErr"] = str(ex)
    try:
        db = get(f"{DRAFT}/bootstrap-static")
        d_by_id = {x["id"]: x for x in db["elements"]}
        es = get(f"{DRAFT}/league/{DRAFT_LEAGUE}/element-status")
        rows = es.get("element_status", [])
        to_cl = lambda de: M["by_code"].get(de["code"]) if de else None
        sq["draft"] = [{"e": to_cl(d_by_id[s["element"]]), "bench": False}
                       for s in rows if s["owner"] == DRAFT_ENTRY_ID and to_cl(d_by_id.get(s["element"]))]
        sq["fa"] = [to_cl(d_by_id[s["element"]]) for s in rows if s["owner"] is None and to_cl(d_by_id.get(s["element"]))]
    except Exception as ex:
        sq["draftErr"] = str(ex)
    return sq


def squad_rows(M, picks, pool, budget_mode):
    rows = []
    for x in picks:
        e = x["e"]
        r = exp_ppg(M, e)
        if budget_mode == "classic":
            budget = (M["sq"]["bank"] or 0) * 10 + e["now_cost"]
            up = best_upgrade(M, e, pool, budget)
        else:
            up = best_upgrade(M, e, pool, None)
        rows.append({**x, **r, "up": up})
    return rows


def weakest(rows):
    weak = None
    for r in rows:
        if r.get("bench"):
            continue
        if r["up"] and (weak is None or r["up"]["gap"] > weak["up"]["gap"]):
            weak = r
    return weak


def fmt_table(M, rows, is_classic, err):
    if err:
        return f"  couldn't load — {err}"
    weak = weakest(rows)
    rows = sorted(rows, key=lambda r: ((1 if r.get("bench") else 0), -r["ppg"]))
    out = []
    out.append(f"  {'Player':<15}{'Pos':<4}{'£':>5} {'Min/g':>6} {'ppg':>6}  {'Sig':<7} Best {'affordable upgrade' if is_classic else 'free agent'}")
    for r in rows:
        e = r["e"]
        cut = weak and e["id"] == weak["e"]["id"]
        sig = "CUT" if cut else ("MINS" if r["flag"] == "mins" else "FIT?" if r["flag"] == "fit?" else "BENCH" if r.get("bench") else "")
        cap = " ©" if r.get("isC") else ""
        price = f"£{e['now_cost']/10:.1f}" if is_classic else "—"
        if r["up"]:
            u = r["up"]
            pen = " ⚠" if u["pen"] else ""
            upstr = f"{u['e']['web_name']} ({M['teams'][u['e']['team']]['name']}){pen}  +{u['gap']:.2f}"
        else:
            upstr = "— none better"
        out.append(f"  {e['web_name'][:14]+cap:<15}{POS[e['element_type']]:<4}{price:>5} {round(r['mpg']):>5}' {r['ppg']:>6.2f}  {sig:<7} {upstr}")
    return "\n".join(out), weak


def main():
    M = build_model()
    M["sq"] = load_squads(M)
    owned = {x["e"]["id"] for x in M["sq"]["classic"]}
    classic_pool = [c for c in M["elements"] if c["id"] not in owned]
    crows = squad_rows(M, M["sq"]["classic"], classic_pool, "classic")
    drows = squad_rows(M, M["sq"]["draft"], M["sq"]["fa"], "draft")

    print(f"TRANSFER SUGGESTER · process-board method · GW{M['cur_gw']} (last played GW{M['last_played']})")
    print(f"Horizon: next {HORIZON} · season basis · bank £{M['sq']['bank']:.1f}\n")

    print("== CLASSIC (entry 37889) ==")
    ct, cweak = fmt_table(M, crows, True, M["sq"]["classicErr"])
    print(ct)
    if cweak:
        u = cweak["up"]
        print(f"\n  >> CUT: {cweak['e']['web_name']} ({POS[cweak['e']['element_type']]}, ppg {cweak['ppg']:.2f})"
              f"  ->  {u['e']['web_name']} ({M['teams'][u['e']['team']]['name']}, +{u['gap']:.2f} ppg)")

    print("\n== DRAFT (league 6079) ==")
    dt, dweak = fmt_table(M, drows, False, M["sq"]["draftErr"])
    print(dt)
    if dweak:
        u = dweak["up"]
        print(f"\n  >> CUT: {dweak['e']['web_name']} ({POS[dweak['e']['element_type']]}, ppg {dweak['ppg']:.2f})"
              f"  ->  FA {u['e']['web_name']} ({M['teams'][u['e']['team']]['name']}, +{u['gap']:.2f} ppg)")

    print("\n(ppg = projected FPL points/game over the horizon. CUT = biggest starter upgrade-gap. "
          "pens counted; ⚠ = pen-taker. Report-only — suggestions, never executed.)")


if __name__ == "__main__":
    main()
