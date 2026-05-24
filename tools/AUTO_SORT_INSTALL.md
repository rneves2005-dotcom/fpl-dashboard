# Auto-Sort Setup for Squads_Data

The Players table will auto-sort every 5 minutes by:
1. **Clubs** (A → Z)
2. **Shirt** (low → high)

…with a 5-minute idle reset so it never fires while you're typing.

## Prerequisites

- **Excel for Mac with macros enabled**
- The file must be saved as `.xlsm` (macro-enabled). `.xlsx` cannot store VBA.

## Step 1 — Save as .xlsm

1. Open `Squads_Data.xlsx` in Excel
2. **File → Save As…**
3. Format: **Excel Macro-Enabled Workbook (.xlsm)**
4. Save as `Squads_Data.xlsm` in the SAME folder (`OneDrive → Claude → DB`)
5. (Optional) Delete `Squads_Data.xlsx` afterwards, or keep as a "data-only" snapshot
6. **IMPORTANT for Claude scripts** — let me know once renamed; I'll update DB_PATH from `.xlsx` to `.xlsm` in the sync tools

## Step 2 — Open the VBA editor

1. With the .xlsm file open, go to **Tools → Macro → Visual Basic Editor**
   (or **⌥+F11** on Mac)
2. The VBA editor opens

## Step 3 — Add the AutoSort module

1. In the VBA editor, find the **Project** panel on the left (shows `VBAProject (Squads_Data.xlsm)`)
2. Right-click `VBAProject (Squads_Data.xlsm)` → **Insert → Module**
3. A new empty module appears (e.g., "Module1")
4. **File → Import File…** → select `auto_sort.bas` from `~/Code/fpl-dashboard/tools/`
5. The new module is now named **AutoSort** with all the sort code

(Alternative: open `auto_sort.bas` in any text editor, copy ALL its contents,
and paste into the empty Module1 in the VBA editor.)

## Step 4 — Add the workbook event hooks

1. In the Project panel, double-click **ThisWorkbook**
2. A code window opens — likely empty
3. Open `auto_sort_thisworkbook.bas` in a text editor, copy the contents
   (everything except the first comment line) and paste into the ThisWorkbook
   code window

## Step 5 — Save the macro-enabled workbook

1. **Cmd+S**
2. If Excel asks about saving as macro-enabled, confirm
3. Close and reopen the file (so the `Workbook_Open` event fires)

## Step 6 — Verify

1. After re-opening, Excel's status bar (bottom) should say:
   `AutoSort ON — sorts every 5 min by Clubs → Shirt`
2. Type a small change in the Players sheet, wait 5 minutes idle — the table will
   sort automatically
3. To force sort immediately: **Tools → Macro → Macros…** → run **SortNow**
4. To turn off: run **StopAutoSort**
5. To turn back on: run **StartAutoSort**

## Mac security prompts

The first time you open the file after adding macros, macOS / Excel may show:
> "This file contains macros. Do you want to enable them?"

Click **Enable Macros**. You may also need to:
- **Excel → Preferences → Security & Privacy → Enable all macros** (if Excel
  refuses to enable per-file)
- macOS may also prompt **System Settings → Privacy & Security → Files & Folders**
  to grant Excel access to the OneDrive folder

## Settings you can tweak (in the AutoSort module)

```vba
Public Const SORT_INTERVAL_MINUTES As Integer = 5   ' change to 10, 30, etc.
Public Const SORT_SHEET As String = "Players"
Public Const SORT_TABLE As String = "Players"
```

## Things that pause / break the sort

- ✅ Editing a cell — sort waits another 5 min after your last edit (idle reset)
- ⚠️ Filter applied to the table — Excel may sort within visible rows only;
  clear filter before relying on global sort
- ⚠️ Frozen pane — fine, sort still works
- ❌ File closed — sort timer stops; resumes when reopened
- ❌ Excel quit unexpectedly — `Application.OnTime` is lost; reopen the file

## Removal

To remove auto-sort: in the VBA editor delete the `AutoSort` module and clear
the `ThisWorkbook` code → save → reopen.
