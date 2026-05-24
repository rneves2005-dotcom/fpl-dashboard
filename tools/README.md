# Player Database Workflow

`Squads_Data.xlsx` is the **single source of truth** for player → club → division mappings.
Both Rui (manual Excel edits) and Claude (via these scripts) can update it. `world_data.json`
is the derived artifact that powers `world.html` and `stats.html` — always regenerated from
the latest xlsx state.

## File locations

User-facing paths in Finder:

| File | Finder path |
|---|---|
| Source DB (xlsx) | `OneDrive → Documents → Squads_Data.xlsx` |
| Backups | same folder, `Squads_Data.xlsx.bak_before_<change>_<timestamp>` |
| OneDrive JSON mirror | `OneDrive → Claude → FPL → world_data.json` |
| Repo JSON (deployed) | `~/Code/fpl-dashboard/world_data.json` |
| Transfer overlays | `~/Code/fpl-dashboard/transfers_{in,out}.json` |

Under the hood, scripts read/write through OneDrive's local sync cache
(`~/Library/Group Containers/UBF8T346G9.OneDriveSyncClientSuite/OneDrive.noindex/OneDrive/…`)
because macOS TCC blocks shell access to the user-facing path under
`~/Library/CloudStorage/OneDrive-Personal/Documents/`. OneDrive propagates
changes from the cache back to the user-facing file automatically.

## Schema

The `Players` sheet has 12 columns (do not reorder):

| Col | Field | Notes |
|---|---|---|
| A | ID | unique key |
| B | Shirt | club shirt #. **Clear on transfer** (unknown until announced) |
| C | Player | full name |
| D | Age | |
| E | Clubs | current club. **Use exact DB form** (e.g. `Sporting`, not `Sporting CP`) |
| F | Country | country of the **club's league**, not player nationality |
| G | Division | one of 28 tracked names (or empty → falls into `🌐 {Country}` virtual division) |
| H | Goals | season goals |
| I | Games | season apps |
| J | International | player's national team (this is where nationality lives) |
| K | Shirt Int | national team shirt # |
| L | Previous Club | most recent past club |

## Workflows

### Rui edits xlsx directly in Excel

1. Open `Squads_Data.xlsx`, make changes, save & close.
2. Tell Claude "I updated the file" (or run yourself):
   ```bash
   python3 tools/sync.py                 # regenerate world_data.json
   python3 tools/validate_overlays.py    # confirm overlay clubs still match
   ```
3. Optionally inspect the diff:
   ```bash
   python3 tools/diff_db.py              # compare against most recent .bak
   ```
4. Commit + push:
   ```bash
   git add world_data.json && git commit -m "Sync from xlsx" && git push
   ```

### Claude applies a transfer (via chat)

```bash
# Standard move
python3 tools/apply_transfer.py --player "Geovany Quenda" --to-club "Chelsea" --to-div "Premier League"

# Departure without known destination
python3 tools/apply_transfer.py --player "Luuk de Jong" --leaving

# Disambiguate when two players share a name
python3 tools/apply_transfer.py --player "Thiago Silva" --current-club "FC Porto" --leaving
```

The script automatically:
1. Creates a `.bak_before_<player>_<timestamp>` snapshot of the xlsx
2. Mutates Clubs / Division / Country (derived from new Division's host country) / Shirt (cleared) / Previous Club (← old club)
3. Regenerates `world_data.json` in both repo + OneDrive
4. Prints a one-line summary

Claude then edits `transfers_in.json` and `transfers_out.json` overlays to annotate the transfer
(destination, "from", note like "Confirmed 26/27"), runs the validator, commits + pushes.

### Hard rules

- **Country column** = country of the **league**, not the player. Derived automatically from
  Division via `DIVISION_TO_COUNTRY` map in `apply_transfer.py`.
- **Shirt** is cleared on every transfer (unknown at announce time).
- **Player nationality** lives in `International` (J) — never modified by transfer scripts.
- **Club names must match DB exactly** — overlay validator catches typos like `Sporting CP`
  vs `Sporting`.

## Scripts

| Script | Purpose |
|---|---|
| `apply_transfer.py` | Apply one transfer to DB + JSON. Backs up xlsx automatically. |
| `sync.py` | Regenerate `world_data.json` from current xlsx (no mutation). |
| `diff_db.py` | Compare current xlsx against most recent backup (or a specified one). |
| `validate_overlays.py` | Check `transfers_in/out.json` club keys against DB. Run after any overlay edit. |
