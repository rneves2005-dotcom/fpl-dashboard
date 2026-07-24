#!/usr/bin/env python3
"""Club fact lookup — RUN THIS BEFORE ANY manager/role/system claim.

Three times on 2026-07-23 I asserted a manager fact that was already written in
team_meta.json (Iraola@LIV, Glasner@NFO, Sage@CRY). Remembering to "check the
notes" demonstrably does not work. This makes it one command.

Usage:
  python3 tools/club.py CRY            # manager + system + full note
  python3 tools/club.py CRY --managers # just the manager/system lines
  python3 tools/club.py --all          # every club's manager line
"""
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
META = os.path.join(ROOT, "team_meta.json")

# words that signal a manager/system fact worth surfacing first
MGR = re.compile(
    r"manager|head coach|\bHC\b|boss|appointed|succeeds|system|back-3|back-4|"
    r"wing-back|formation|style|profile =",
    re.I,
)


def load():
    with open(META) as f:
        return json.load(f)


def segs(note):
    return [s.strip() for s in note.split(" · ") if s.strip()]


def show(code, meta, managers_only=False):
    if code not in meta:
        sys.exit(f"ERROR: '{code}' not found. Valid: {', '.join(sorted(meta))}")
    note = meta[code].get("note", "")
    parts = segs(note)
    mgr = [s for s in parts if MGR.search(s)]

    print(f"\n=== {code} ===")
    mgr_field = meta[code].get("manager", "(not recorded)")
    stale = "RUMOR" in mgr_field.upper() or "TBD" in mgr_field.upper()
    print(f"{'⚠️ ' if stale else '✅ '}MANAGER: {mgr_field}")
    if meta[code].get("system"):
        print(f"   system: {meta[code]['system']}")
    if meta[code].get("discount"):
        print(f"   discount: {meta[code]['discount']}")
    print()
    if mgr:
        print("🧑‍💼 MANAGER / SYSTEM:")
        for s in mgr:
            print(f"   • {s}")
    else:
        print("🧑‍💼 MANAGER / SYSTEM: (nothing recorded — VERIFY before claiming)")
    if managers_only:
        return
    print(f"\n📋 FULL NOTE ({len(parts)} entries):")
    for s in parts:
        print(f"   - {s}")


def main():
    args = [a for a in sys.argv[1:]]
    meta = load()
    if not args or "--all" in args:
        print("=== MANAGER / SYSTEM line per club ===")
        for code in sorted(meta):
            mgr = meta[code].get("manager", "(not recorded)")
            flag = "⚠️" if ("RUMOR" in mgr.upper() or "TBD" in mgr.upper()) else "✅"
            print(f"{code:5} {flag} {mgr[:60]:62}{str(meta[code].get('discount',''))}")
        return 0
    managers_only = "--managers" in args
    for code in [a.upper() for a in args if not a.startswith("--")]:
        show(code, meta, managers_only)
    return 0


if __name__ == "__main__":
    sys.exit(main())
