# Caseload Report App — GUI Build Plan

## Overview

Package the existing `caseload_report.py` processing engine into a standalone
tkinter GUI app distributed as a single `.exe` file via PyInstaller.

---

## App Structure

```
Caseload_Report_Abdi_App-/
├── app.py                      — GUI entry point (tkinter)
├── caseload_report.py          — Processing engine (existing, no changes)
├── build_exe.py                — PyInstaller build script
├── config/
│   ├── staff_roster/           — Staff name mapping (optional .xlsx)
│   └── site_tab_mapping.csv    — Program-to-tab mapping
├── input/
│   ├── data_report_card/       — Data Report Card export
│   └── legal_referral/         — Legal Services Referral export
├── output/                     — Generated reports
└── requirements.txt            — + pyinstaller
```

**Key separation:** `app.py` handles the GUI. `caseload_report.py` stays as the
pure processing module — called by the GUI but also still runnable from CLI.

---

## GUI Layout

```
┌─────────────────────────────────────────────────┐
│  HEADER  (#1F4E79 dark blue banner)             │
│    "Caseload Report Generator"  (16pt bold)     │
│    "SVdP CARES Data Systems"    (10pt light)    │
├─────────────────────────────────────────────────┤
│  BODY  (#F0F4F8 light background)               │
│                                                 │
│  Data Report Card (bold label)                  │
│  [ /path/to/file.xlsx        ] [Browse...]      │
│                                                 │
│  Legal Services Referral (bold label)           │
│  [ /path/to/file.xlsx        ] [Browse...]      │
│                                                 │
│  Output Folder (bold label)                     │
│  [ /path/to/output/          ] [Browse...]      │
│                                                 │
│  ─── Config (collapsible/section) ───           │
│  Staff Roster (optional)                        │
│  [ /path/to/roster.xlsx      ] [Browse...]      │
│  ✓ Staff roster loaded (42 entries)             │
│                                                 │
│  [========= progress bar =========]             │
│  "Ready"                                        │
│                                                 │
│                       [ Generate Report ]       │
├─────────────────────────────────────────────────┤
│  FOOTER  (#E8EDF2)                              │
│    "SVdP CARES · Data Systems · v1.0"           │
└─────────────────────────────────────────────────┘
```

### Browse Fields (4 total)

| # | Field | Type | Default | Required |
|---|-------|------|---------|----------|
| 1 | Data Report Card | File (.xlsx) | Last used / `input/data_report_card/` | Yes |
| 2 | Legal Services Referral | File (.xlsx) | Last used / `input/legal_referral/` | Yes |
| 3 | Output Folder | Directory | `output/` | Yes |
| 4 | Staff Roster | File (.xlsx) | `config/staff_roster/` auto-detect | No (optional) |

### Status Indicators

- **Green check (✓)** — File found/valid, roster loaded, report generated
- **Yellow warning (⚠)** — Optional file missing (e.g., no staff roster)
- **Red X (✗)** — Required file missing, processing error

### Buttons

| Button | Action | State |
|--------|--------|-------|
| Browse... (×4) | Open file/folder dialog | Always enabled |
| Generate Report | Run processing pipeline | Disabled during processing |

---

## Color Palette (from design reference)

| Element | Color | Hex |
|---------|-------|-----|
| Header / primary button | Dark blue | `#1F4E79` |
| Button hover | Medium blue | `#2E75B6` |
| Body background | Light gray-blue | `#F0F4F8` |
| Footer background | Gray | `#E8EDF2` |
| Success | Green | `#2E7D32` |
| Error | Red | `#C62828` |
| Muted text | Gray | `#555` |

## Fonts (Segoe UI throughout)

| Element | Spec |
|---------|------|
| Title | 16pt bold |
| Subtitle | 10pt regular |
| Section labels | 10pt bold |
| Inputs / body | 9pt regular |
| Button | 11pt bold |
| Footer | 8pt italic |

---

## App Behavior

### Startup
1. Window opens, non-resizable
2. Auto-detect files in default input folders if present
3. Pre-fill Output Folder with `output/`
4. Auto-detect staff roster from `config/staff_roster/`
5. Show status indicators for each field (found/missing)
6. Load `config/site_tab_mapping.csv` silently (no UI — power user config)

### Generate Report (click)
1. Validate required fields — show error dialog if missing
2. Disable Generate button
3. Start indeterminate progress bar
4. Status text: "Processing..."
5. Run `caseload_report.main()` in a **background thread** (keeps UI responsive)
6. On success:
   - Stop progress bar
   - Status: "✓ Report generated: Current_Caseload_2026-04-03.xlsx"
   - Dialog: "Report generated successfully. Open the output folder?"
   - If yes → `os.startfile(output_folder)` (Windows)
7. On error:
   - Stop progress bar
   - Status: "✗ Error: [message]"
   - Error dialog with details
   - Re-enable Generate button

### Threading Pattern
```
User clicks "Generate"
  → GUI thread: disable button, start progress bar
  → Background thread: call caseload_report.main(...)
  → On completion: root.after(0, callback)  # back to GUI thread
  → GUI thread: stop progress, show result
```

---

## Path Resolution (PyInstaller compatible)

```python
def get_app_dir():
    """Get the directory where the app/exe lives."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)  # .exe mode
    return os.path.dirname(os.path.abspath(__file__))  # script mode
```

All relative paths (config/, input/, output/) resolve from `get_app_dir()`.
This means the folder structure ships alongside the .exe.

### Distribution Layout (what the user gets)
```
Caseload_Report/
├── CaseloadReport.exe          — The app
├── config/
│   ├── staff_roster/           — Drop roster .xlsx here
│   └── site_tab_mapping.csv    — Edit to configure tabs
├── input/
│   ├── data_report_card/       — Drop DRC export here
│   └── legal_referral/         — Drop Legal export here
└── output/                     — Reports saved here
```

---

## PyInstaller Build

### build_exe.py
```python
PyInstaller flags:
  --onefile         → single .exe
  --windowed        → no console window
  --name CaseloadReport
  --add-data "config;config"       → bundle config files
  --icon icon.ico                  → app icon (optional)
```

### Post-build
- Copy `config/` folder next to the .exe (for user-editable configs)
- The .exe reads configs from its own directory, not the bundled ones

---

## Implementation Steps

### Phase 1: GUI Shell
- [ ] Create `app.py` with tkinter layout (header, body, footer)
- [ ] Add 4 browse fields with file/folder dialogs
- [ ] Add progress bar and status label
- [ ] Add Generate Report button (disabled state)
- [ ] Wire up file validation (green/red indicators)

### Phase 2: Integration
- [ ] Connect Generate button to `caseload_report.main()` via threading
- [ ] Pass GUI-selected paths to the processing engine
- [ ] Capture and display log output in status area
- [ ] Handle success/error callbacks

### Phase 3: Polish
- [ ] Remember last-used paths (save to a small JSON preferences file)
- [ ] Auto-detect files in default folders on startup
- [ ] Add "Open Output Folder" after successful generation
- [ ] Window icon

### Phase 4: Build & Package
- [ ] Create `build_exe.py` with PyInstaller config
- [ ] Test .exe on a clean Windows machine
- [ ] Create a zip distribution with exe + config folders
- [ ] Document for end users

---

## Considerations

- **Config files must be editable** — they live outside the .exe, not bundled
- **site_tab_mapping.csv is power-user** — no GUI needed, just edit the CSV
- **Staff roster could change monthly** — GUI file picker makes this easy
- **No admin rights needed** — the .exe runs from any folder
- **Input files contain PII** — app should never upload/transmit anything
- **Future: could add a log viewer** — scrollable text area showing processing log
