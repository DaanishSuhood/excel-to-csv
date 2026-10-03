# Excel → CSV Converter

Exports every worksheet of every selected workbook as its own UTF-8 CSV file. Works on a hand-picked
list of files, or on a whole folder tree.

Script: [`excel-to-csv.ahk`](excel-to-csv.ahk)

---

## Requirements

| Requirement | Notes |
| --- | --- |
| [AutoHotkey v2.0+](https://www.autohotkey.com/) | Declares `#Requires AutoHotkey v2.0` |
| Microsoft Excel | Excel performs the export; the script drives it over COM |

## Getting started

1. Run `excel-to-csv.ahk`.
2. Pick an input mode:
   - **Convert individual files** — **Add Files…** to build a list (`.xlsx`, `.xlsm`, `.xls`).
   - **Convert a whole folder** — **Browse…** to a folder; it is scanned **recursively**, including
     subfolders.
3. Choose an **Output Folder**.
4. Click **Convert**.

## Output naming

One CSV per worksheet, all written flat into the output folder:

```
<workbook name>_<sheet name>.csv
```

Both parts are sanitised — `\ / : * ? " < > |` become `_`, leading and trailing spaces and dots are
trimmed, and an empty result becomes `Sheet`. If that name is already taken, ` (2)`, ` (3)`, … is
appended, so nothing is ever silently overwritten.

## The window

| Control | Purpose |
| --- | --- |
| **Convert individual files** / **Convert a whole folder** | Input mode; the controls below swap to match |
| File list + **Add Files… / Remove / Clear All** | The individual-file list. Duplicates are ignored |
| Folder box + **Browse…** | The folder to scan recursively |
| **Output Folder** + **Browse…** | Where the CSVs go |
| **Convert** | Starts the run; input controls are disabled while it works |
| Progress bar + status line | `Converting (n/total): <path>` |
| Log box | `Processing: <path>` then `-> OK (n sheet(s) exported)` or `-> FAILED: <message>`, ending with a totals block |
| **Open Output Folder** | Enabled after a run |

## How it works

The script is one `ExcelToCsvApp` class. On **Convert** it starts a fresh hidden Excel instance
(`Visible := false`, `DisplayAlerts := false`, `ScreenUpdating := false`) — it does **not** attach to
a running Excel — and opens each workbook read-only with `Workbooks.Open(path, 0, true)`.

For each worksheet, `sheet.Copy()` creates a temporary single-sheet workbook which is saved with
`SaveAs` using format **62** (`xlCSVUTF8`), falling back to **6** (`xlCSV`, ANSI) if the running Excel
is too old to know 62. The temporary workbook is closed without saving, and the source workbook is
closed in a `finally` block.

Per-file failures are logged and counted rather than aborting the run, and Excel is quit in a
`finally` block so it does not linger after an error.

## Limitations

- **Every sheet is exported**, including hidden sheets and sheets you did not want. There is no
  filter.
- **Values only.** CSV cannot carry formulas, formatting, merged cells, multiple sheets, or charts —
  each cell is written as Excel currently displays it.
- The output is flat: a recursive folder scan puts CSVs from all subfolders into the one output
  folder. Two workbooks with the same name in different subfolders both convert, distinguished only
  by the ` (2)` suffix.
- The GUI freezes for the duration; there is no cancel button, and the progress bar only advances
  between files, not between sheets.
- Password-protected workbooks will stall on Excel's password prompt (`DisplayAlerts` does not
  suppress it) or fail.
- The log box auto-scrolls by focusing itself and sending `Ctrl+End`, so avoid typing elsewhere while
  a batch is running.
- If format 62 is unavailable, the fallback writes ANSI — non-Latin text (Sinhala, Tamil) will be
  mangled in those files.
