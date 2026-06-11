#!/usr/bin/env python3
"""Hourly transfer poller — feeds Word tracker doc + JSON cache.

Sources:
  - xcancel.com RSS for ~25 X accounts (FabrizioRomano, Ornstein, club official)
  - maisfutebol cronologia RSS (Portuguese transfer timeline)
  - premierleague.com /en/transfers/2026-27/summer (HTML scrape)

Output:
  - /Users/ruimiguelneves/Code/fpl-dashboard/Transfer_Tracker.docx
    (append-only Word doc; new section per run, table of new entries)
  - tools/cache/last_run.json (dedup state per source)

Run hourly via launchd plist `com.fpl.transfer-poll.plist`.
"""

from __future__ import annotations
import os
import re
import sys
import json
import time
from pathlib import Path
from datetime import datetime, timezone
from typing import Any

try:
    import feedparser
    import requests
    from bs4 import BeautifulSoup
    from docx import Document
    from docx.shared import Inches, RGBColor, Pt
    from docx.enum.table import WD_ALIGN_VERTICAL
except ImportError as e:
    print(f"Missing dep: {e}\nInstall: pip3 install feedparser requests beautifulsoup4 python-docx", file=sys.stderr)
    sys.exit(1)

# ─── Config ──────────────────────────────────────────────────────────────

ROOT = Path(__file__).resolve().parent.parent
CACHE_DIR = ROOT / "tools" / "cache"
CACHE_DIR.mkdir(parents=True, exist_ok=True)
LAST_RUN_PATH = CACHE_DIR / "last_run.json"
TRACKER_DOCX = ROOT / "Transfer_Tracker.docx"
ONEDRIVE_TRACKER = Path("/Users/ruimiguelneves/Library/Group Containers/UBF8T346G9.OneDriveSyncClientSuite/OneDrive.noindex/OneDrive/Claude/FPL/Transfer_Tracker.docx")

XCANCEL = "https://nitter.net"  # primary Nitter mirror · fallback to others if needed
NITTER_FALLBACKS = ["https://nitter.net", "https://xcancel.com", "https://lightbrd.com"]

X_ACCOUNTS = [
    # Journos
    "FabrizioRomano", "David_Ornstein",
    # Official
    "premierleague",
    # Top-6 clubs
    "NUFC", "LFC", "ManUtd", "ManCity", "Arsenal",
    # Top-12
    "SpursOfficial", "ChelseaFC", "CPFC", "Everton", "AVFCOfficial",
    # Mid
    "LCFC", "Wolves", "SouthamptonFC", "OfficialBHAFC", "NFFC",
    # Promoted + relegated
    "LUFC", "HullCity", "IpswichTown", "Coventry_City",
    "WestHam", "BurnleyOfficial",
    # Other PL
    "BrentfordFC", "afcbournemouth", "FulhamFC", "SunderlandAFC",
]

MAISFUTEBOL_RSS = "http://feeds.feedburner.com/iol/maisfutebol"
PL_TRANSFERS_URL = "https://www.premierleague.com/en/transfers/2026-27/summer"

# Keyword filters — case-insensitive substring match
# Filter logic: HIGH requires (CONFIRM_WORD AND TRANSFER_WORD) or (HERE_WE_GO / OFICIAL standalone)
#               LOW requires RUMOR_WORD AND TRANSFER_WORD
# This filters out hashtag/promo noise like "#OFFICIAL Twitter video"

CONFIRM_WORDS = [
    "OFICIAL", "OFICIALMENTE", "confirma", "anuncia", "anunciou", "anunciaram",
    "OFFICIAL:", "officially", "confirmed",
    "we are delighted to announce", "have signed", "we have signed",
    "are delighted to announce", "agreed", "deal", "permanent move",
]

TRANSFER_WORDS = [
    # Portuguese
    "transferência", "contratação", "contrato", "saída", "deixa", "chega",
    "regressa", "empréstimo", "transferiu", "abandona", "novo treinador",
    # English
    "signed", "signs", "joins", "joining", "joined", "leaves", "leaving", "departure",
    "departs", "loan", "loaned", "release", "released", "free transfer", "free agent",
    "contract expires", "contract expiry", "appointed", "step down", "steps down",
    "stepping down", "new head coach", "new manager", "thank you", "farewell",
]

# Standalone triggers — single keywords that ALWAYS qualify regardless of context
STANDALONE_HIGH = [
    "HERE WE GO", "Here we go!", "🚨 OFICIAL", "🚨 OFFICIAL",
]

KEYWORDS_LOW_TRANSFER = TRANSFER_WORDS  # transfer word reqd for rumor too
RUMOR_WORDS = [
    "rumor", "rumour", "talks", "discussions", "talks with", "in advanced negotiations",
    "verbal agreement", "negotiations underway", "interested in", "close to signing",
    "set to sign", "set to join", "wants to sign",
]

# Anti-keywords: presence reduces confidence (likely social/marketing not transfer)
ANTI_WORDS = [
    "match preview", "matchday", "ticket", "kit launch", "behind the scenes",
    "🎶", "🎤", "podcast", "interview", "feature", "matchcam", "highlight",
    "wallpaper", "buy now", "save", "discount",
]

# ─── HTTP helpers ────────────────────────────────────────────────────────

USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"

def fetch(url: str, timeout: int = 20) -> str | None:
    try:
        r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print(f"FETCH FAIL {url}: {e}", file=sys.stderr)
        return None

# ─── Source: xcancel.com RSS ─────────────────────────────────────────────

def poll_x_accounts() -> list[dict[str, Any]]:
    """Pull RSS for each X account via Nitter · falls back through mirrors."""
    out = []
    for handle in X_ACCOUNTS:
        body = None
        used_mirror = None
        for mirror in NITTER_FALLBACKS:
            url = f"{mirror}/{handle}/rss"
            body = fetch(url, timeout=15)
            if body and "whitelist" not in body.lower()[:2000] and "<item>" in body:
                used_mirror = mirror
                break
            body = None
        if not body:
            continue
        feed = feedparser.parse(body)
        for e in feed.entries[:20]:  # cap last 20 per account
            entry = {
                "source": f"X:@{handle}",
                "title": e.get("title", "")[:300],
                "summary": e.get("summary", "")[:1500],
                "link": e.get("link", ""),
                "published": e.get("published", ""),
                "id": e.get("id") or e.get("link"),
            }
            out.append(entry)
    return out

# ─── Source: maisfutebol RSS ─────────────────────────────────────────────

def poll_maisfutebol() -> list[dict[str, Any]]:
    body = fetch(MAISFUTEBOL_RSS, timeout=20)
    if not body:
        return []
    feed = feedparser.parse(body)
    out = []
    for e in feed.entries[:50]:
        out.append({
            "source": "maisfutebol",
            "title": e.get("title", "")[:300],
            "summary": e.get("summary", "")[:1500],
            "link": e.get("link", ""),
            "published": e.get("published", ""),
            "id": e.get("id") or e.get("link"),
        })
    return out

# ─── Source: Premier League HTML ─────────────────────────────────────────

def poll_premier_league() -> list[dict[str, Any]]:
    body = fetch(PL_TRANSFERS_URL, timeout=20)
    if not body:
        return []
    soup = BeautifulSoup(body, "html.parser")
    out = []
    # PL transfers are usually in a list with player + direction
    # Generic crawl — find any link with /transfers/ in href
    for a in soup.find_all("a", href=True):
        txt = a.get_text(strip=True)
        if not txt or len(txt) > 200:
            continue
        href = a["href"]
        if "transfer" in href.lower() or "signing" in txt.lower() or "joins" in txt.lower():
            out.append({
                "source": "PL Official",
                "title": txt,
                "summary": "",
                "link": href if href.startswith("http") else f"https://www.premierleague.com{href}",
                "published": datetime.now(timezone.utc).isoformat(),
                "id": f"PL:{txt}:{href}",
            })
    return out

# ─── Filter ──────────────────────────────────────────────────────────────

def classify_confidence(text: str) -> str:
    """HIGH requires (CONFIRM word AND TRANSFER word) or standalone trigger.
    LOW requires (RUMOR word AND TRANSFER word).
    Anti-words downgrade or skip.
    """
    tl = text.lower()

    # Anti-words = likely social/marketing noise · skip
    for w in ANTI_WORDS:
        if w.lower() in tl:
            return "SKIP"

    # Standalone HIGH triggers (HERE WE GO, 🚨 OFICIAL)
    for w in STANDALONE_HIGH:
        if w.lower() in tl:
            return "HIGH"

    has_transfer = any(w.lower() in tl for w in TRANSFER_WORDS)
    has_confirm = any(w.lower() in tl for w in CONFIRM_WORDS)
    has_rumor = any(w.lower() in tl for w in RUMOR_WORDS)

    if has_confirm and has_transfer:
        return "HIGH"
    if has_rumor and has_transfer:
        return "LOW"
    return "SKIP"

def matched_keywords(text: str) -> list[str]:
    tl = text.lower()
    matched = []
    for w in CONFIRM_WORDS + TRANSFER_WORDS + RUMOR_WORDS + STANDALONE_HIGH:
        if w.lower() in tl:
            matched.append(w)
    return matched[:8]  # cap shown

# ─── Dedup ───────────────────────────────────────────────────────────────

def load_seen() -> set[str]:
    if not LAST_RUN_PATH.exists():
        return set()
    try:
        return set(json.loads(LAST_RUN_PATH.read_text()).get("seen_ids", []))
    except Exception:
        return set()

def save_seen(seen: set[str]):
    # Cap at 5000 to keep file manageable
    if len(seen) > 5000:
        seen = set(list(seen)[-5000:])
    LAST_RUN_PATH.write_text(json.dumps({"seen_ids": sorted(seen), "updated": datetime.now(timezone.utc).isoformat()}, indent=2))

# ─── Word doc output ─────────────────────────────────────────────────────

def ensure_doc() -> Document:
    """Open existing tracker or create new."""
    if TRACKER_DOCX.exists():
        return Document(str(TRACKER_DOCX))
    doc = Document()
    # Header
    h = doc.add_heading("FPL 26/27 Transfer Tracker", 0)
    p = doc.add_paragraph()
    p.add_run("Append-only log · review entries · mark with ✓ when applied to Squads_2026.xlsm DB.\n").italic = True
    p.add_run(f"Created: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n").italic = True
    p.add_run("Sources: xcancel.com (Twitter mirror) · maisfutebol cronologia · premierleague.com official\n").italic = True
    return doc

def append_run_section(doc: Document, entries: list[dict[str, Any]]):
    """Add a new section with table for this run's new entries."""
    if not entries:
        return
    doc.add_heading(f"Run · {datetime.now().strftime('%Y-%m-%d %H:%M')}", level=2)
    p = doc.add_paragraph()
    p.add_run(f"{len(entries)} new entries found · highest-confidence first").italic = True

    # Table
    tbl = doc.add_table(rows=1, cols=6)
    tbl.style = "Light Grid Accent 1"
    hdr = tbl.rows[0].cells
    hdr[0].text = "✓"
    hdr[1].text = "Confidence"
    hdr[2].text = "Source"
    hdr[3].text = "Title / Headline"
    hdr[4].text = "Keywords matched"
    hdr[5].text = "Link"
    for c in hdr:
        for run in c.paragraphs[0].runs:
            run.font.bold = True
            run.font.size = Pt(9)
    # Sort entries: HIGH first, then LOW
    sorted_entries = sorted(entries, key=lambda e: (0 if e["confidence"] == "HIGH" else 1, e["source"]))
    for e in sorted_entries:
        row = tbl.add_row().cells
        row[0].text = ""  # checkbox column · user marks manually
        row[1].text = e["confidence"]
        row[2].text = e["source"]
        title = e["title"]
        # Strip HTML tags from summary if present in title
        title = re.sub(r"<[^>]+>", "", title)
        row[3].text = title[:200] + ("…" if len(title) > 200 else "")
        row[4].text = ", ".join(e["keywords"][:5])
        row[5].text = e["link"][:120]
        # Color code by confidence
        if e["confidence"] == "HIGH":
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(9)
        else:
            for cell in row:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(8)
                        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

# ─── Main ────────────────────────────────────────────────────────────────

def main():
    started = time.time()
    print(f"[{datetime.now().isoformat()}] poll_transfers START")

    seen = load_seen()

    all_entries = []
    counts = {"X": 0, "maisfutebol": 0, "PL": 0}

    print("Polling X accounts via xcancel…")
    x_entries = poll_x_accounts()
    counts["X"] = len(x_entries)
    all_entries.extend(x_entries)

    print("Polling maisfutebol RSS…")
    mf_entries = poll_maisfutebol()
    counts["maisfutebol"] = len(mf_entries)
    all_entries.extend(mf_entries)

    print("Polling PL Official transfers page…")
    pl_entries = poll_premier_league()
    counts["PL"] = len(pl_entries)
    all_entries.extend(pl_entries)

    print(f"Total raw entries: {len(all_entries)}")

    # Filter + dedup
    keep = []
    new_seen = set(seen)
    for e in all_entries:
        eid = e.get("id") or e.get("link") or e.get("title")
        if not eid or eid in seen:
            continue
        new_seen.add(eid)
        text = f"{e.get('title','')} {e.get('summary','')}"
        conf = classify_confidence(text)
        if conf == "SKIP":
            continue
        e["confidence"] = conf
        e["keywords"] = matched_keywords(text)
        keep.append(e)

    print(f"After filter+dedup: {len(keep)} new entries")
    save_seen(new_seen)

    if not keep:
        print(f"No new entries this run · elapsed {time.time()-started:.1f}s")
        return 0

    # Append to Word doc
    doc = ensure_doc()
    append_run_section(doc, keep)
    doc.save(str(TRACKER_DOCX))
    print(f"Wrote {TRACKER_DOCX} ({TRACKER_DOCX.stat().st_size:,} bytes)")

    # Sync to OneDrive
    try:
        ONEDRIVE_TRACKER.parent.mkdir(parents=True, exist_ok=True)
        ONEDRIVE_TRACKER.write_bytes(TRACKER_DOCX.read_bytes())
        print(f"Synced to OneDrive: {ONEDRIVE_TRACKER}")
    except Exception as e:
        print(f"OneDrive sync skipped: {e}", file=sys.stderr)

    print(f"Sources: X={counts['X']} maisfutebol={counts['maisfutebol']} PL={counts['PL']}")
    print(f"Elapsed: {time.time()-started:.1f}s")
    return 0

if __name__ == "__main__":
    sys.exit(main())
