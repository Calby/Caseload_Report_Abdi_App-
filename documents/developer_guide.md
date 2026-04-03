# Developer Guide

How to run, edit, and build the Caseload Report Generator.

---

## Prerequisites

- Python 3.9 or higher
- Windows (for the GUI app and .exe build)

## Setup

```powershell
# Clone the repo
git clone https://github.com/Calby/Caseload_Report_Abdi_App-.git
cd Caseload_Report_Abdi_App-

# Create a virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

---

## Running

### GUI App

```powershell
python app.py
```

The app opens with file pickers for the Data Report Card, Legal Referral,
and output folder. Drop your CaseWorthy exports in the `input/` folders
and they'll be auto-detected on startup.

### Command Line (no GUI)

```powershell
# Auto-detect files from input folders
python caseload_report.py

# Specify files directly
python caseload_report.py --drc path\to\report.xlsx --legal path\to\referral.xlsx -o output.xlsx
```

---

## Project Structure

```
Caseload_Report_Abdi_App-/
├── app.py                      # GUI entry point (tkinter)
├── caseload_report.py          # Processing engine (standalone, no GUI dependency)
├── build_exe.py                # PyInstaller build script
├── requirements.txt            # Python dependencies
├── CLAUDE.md                   # Build specification / AI context
│
├── config/
│   ├── staff_roster/           # Drop CaseWorthy User Export .xlsx here
│   └── site_tab_mapping.csv    # Program-to-tab mapping (editable CSV)
│
├── input/
│   ├── data_report_card/       # Drop Data Report Card .xlsx here
│   └── legal_referral/         # Drop Legal Services Referral .xlsx here
│
├── output/                     # Generated reports saved here
│
└── documents/
    ├── developer_guide.md      # This file
    ├── gui_app_plan.md         # Original GUI planning doc
    └── gui_design_reference.md # Reusable design spec for other apps
```

---

## Making Changes

### Processing Logic (`caseload_report.py`)

The processing engine is a single file with clear sections:

1. **Constants** (top of file) — Column names, program lists, tab order, formatting.
   Most changes go here. For example:
   - Add a non-SSVF program: edit `NON_SSVF_PROGRAMS`
   - Change column order: edit `ALL_COLUMNS_ORDERED`
   - Add/remove FOX columns: edit `FOX_DROP_COLUMNS`

2. **Loaders** — `load_data_report_card()`, `load_legal_referral()`, `load_staff_roster()`

3. **Processing** — `process_main_sheet()` runs Steps 1-9 in order

4. **Site Tabs** — `create_site_tabs()` uses filters from `config/site_tab_mapping.csv`

5. **Formatting** — `apply_formatting()` handles Excel styling

### Adding a New Site Tab

Option A — Edit `config/site_tab_mapping.csv` (no code changes):
```csv
NewTab,prefix,NewProgram
```

Option B — Add to the hardcoded filters in `create_site_tabs()` as a fallback.

Also add the tab name to `SITE_TAB_ORDER` in the constants section.

### Adding a New Output Column

1. Add the column name to `ALL_COLUMNS_ORDERED` at the desired position
2. Add the source column to `KEEP_COLUMNS_FROM_DRC` if it comes from the Data Report Card
3. Add the calculation/derivation logic in `process_main_sheet()`

### GUI Changes (`app.py`)

The GUI is a single `CaseloadReportApp` class. Key methods:
- `_build_header()` / `_build_body()` / `_build_footer()` — Layout
- `_browse_*()` — File/folder picker handlers
- `_on_generate()` — Validates inputs, launches background thread
- `_process()` — Runs in background, calls `caseload_report.main()` logic
- `_on_success()` / `_on_error()` — Callbacks on GUI thread

### Config Changes

- **Staff roster**: Drop a new `.xlsx` in `config/staff_roster/` or pick it via the GUI
- **Site tab mapping**: Edit `config/site_tab_mapping.csv` directly. Match types:
  - `prefix` — Program Name starts with value
  - `contains` — Program Name contains value (case-insensitive)
  - `location_contains` — Office Location contains value
  - `exclude_contains` — Exclude programs containing value
  - `require_contains` — Require programs containing value
- **Program exclusions**: Edit `config/program_exclusions.xlsx` (see below)

### Program Exclusions (`config/program_exclusions.xlsx`)

This Excel file controls which programs get SSVF-related fields blanked out
(SOAR, ShallowSub, HUDVASH, Recert) and which programs should be silenced
from heuristic warnings. **No code changes or rebuild needed** — just edit
the Excel and re-run.

#### Columns

| Column | Required | Description |
|--------|----------|-------------|
| Program Name | Yes | The program name or keyword to match |
| Exclusion Type | Yes | What action to take (see below) |
| Match Type | No | `exact` (default) or `contains` |
| Notes | No | Free-text description (informational only) |

#### Exclusion Types

| Exclusion Type | Effect |
|----------------|--------|
| `non_ssvf` | Blanks out SOAR, ShallowSub, HUDVASH, Last 90 Day Recert, and Days since Last Recert/Update for matching programs |
| `skip_ssvf_warning` | Suppresses the log warning for programs that don't match SSVF keywords. Does not change any data. |

#### Match Types

| Match Type | How it works | Example |
|------------|-------------|---------|
| `exact` | Program Name must match the value exactly | `Tampa-THHI-CDBG DAP CES 1107` matches only that one program |
| `contains` | Program Name contains the value (case-insensitive) | `EHA` matches `Charlotte-VA Supportive Services-SSVF-EHA`, `Orlando-SSVF-EHA`, etc. |

If the Match Type column is missing or blank, it defaults to `exact`.

#### Examples

| Program Name | Exclusion Type | Match Type | Notes |
|---|---|---|---|
| All-County-VA-Suicide-Prevention 1114 | non_ssvf | exact | Specific program |
| Charlotte County-SHIP-RRH 1305 | non_ssvf | exact | SHIP program |
| Tampa-THHI-CDBG DAP CES 1107 | non_ssvf | exact | CDBG program |
| EHA | non_ssvf | contains | All EHA programs |
| Some Known Program | skip_ssvf_warning | exact | Suppress warning only |

#### Common Tasks

**Add a specific program to exclude:**
1. Open `config/program_exclusions.xlsx` in Excel
2. Add a row: Program Name = full program name, Exclusion Type = `non_ssvf`, Match Type = `exact`
3. Save and re-run

**Exclude all programs matching a keyword:**
1. Add a row: Program Name = the keyword (e.g., `SHIP`), Exclusion Type = `non_ssvf`, Match Type = `contains`
2. This will match any program with that keyword anywhere in its name

**Stop a warning from appearing:**
1. Add a row: Program Name = the program name, Exclusion Type = `skip_ssvf_warning`, Match Type = `exact`
2. The program will still be processed normally, but the warning log will be suppressed

**Fallback:** If the Excel file is missing or has the wrong columns, the script falls back to a hardcoded list of 3 programs built into the code.

---

## Building the .exe

```powershell
# Make sure pyinstaller is installed
pip install pyinstaller

# Build
python build_exe.py
```

This creates `dist/CaseloadReport.exe`.

### Distribution Package

Create a folder to distribute:

```
CaseloadReport/
├── CaseloadReport.exe
├── config/
│   ├── staff_roster/           # Include roster .xlsx if available
│   └── site_tab_mapping.csv
├── input/
│   ├── data_report_card/
│   └── legal_referral/
└── output/
```

Copy the `config/`, `input/`, and `output/` folders from the repo next to the `.exe`.
Zip the whole folder and share.

Users just:
1. Unzip
2. Drop their CaseWorthy exports in the `input/` folders
3. Double-click `CaseloadReport.exe`
4. Click "Generate Report"

No Python installation needed on the user's machine.

---

## Testing

### Quick Smoke Test

```powershell
# Verify CLI works
python caseload_report.py --help

# Verify imports
python -c "import caseload_report; print('OK')"

# Verify GUI syntax (headless)
python -c "import ast; ast.parse(open('app.py').read()); print('OK')"
```

### Full Test

1. Place test exports in `input/data_report_card/` and `input/legal_referral/`
2. Run `python app.py`
3. Click "Generate Report"
4. Open the output file and verify:
   - All tabs present and correctly filtered
   - Column counts correct (All=20, FOX=12, Lake/Citrus=21)
   - Formatting applied (headers, dates, alternating rows, red >30 days)
   - Staff names resolved (if roster provided)
