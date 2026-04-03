# Caseload Report Automation

Python script that automates the SVdP CARES Caseload Report — replaces the manual process of combining CaseWorthy Excel exports into a formatted, multi-tab workbook.

## Setup

```bash
pip install -r requirements.txt
```

Requires Python 3.9+ with `pandas` and `openpyxl`.

## Input Folder Structure

Create the following folders and drop your CaseWorthy exports into them:

```
input/
├── data_report_card/          — Data Report Card export (.xlsx)
└── legal_referral/            — Legal Services Referral export (.xlsx)
```

Each folder should contain exactly one `.xlsx` file. The script auto-detects the file by extension, so the filename from CaseWorthy doesn't matter.

## Usage

```bash
# Auto-detect files from input folders (recommended)
python caseload_report.py

# Specify output filename
python caseload_report.py -o Current_Caseload_2026-04-01.xlsx

# Override input file paths directly
python caseload_report.py --drc path/to/report.xlsx --legal path/to/referral.xlsx
```

### Arguments

| Argument | Description |
|----------|-------------|
| `--drc` | Path to Data Report Card `.xlsx` (default: auto-detect from `input/data_report_card/`) |
| `--legal` | Path to Legal Services Referral `.xlsx` (default: auto-detect from `input/legal_referral/`) |
| `-o, --output` | Output file path (default: `Current_Caseload_{YYYY-MM-DD}.xlsx`) |

## Output

The script produces an Excel workbook with 16 tabs:

| Tab | Description |
|-----|-------------|
| All | Every client across all programs (19 columns) |
| Charlotte | Charlotte County non-shelter programs |
| Charlotte Shelter | Charlotte Care Center shelter programs |
| FOX | SSG Fox / VA Suicide Prevention (12-column reduced layout) |
| GPD | Grant Per Diem programs |
| MidFlorida | Lake/Citrus/Hernando/Sumter programs (19 + Location = 20 columns) |
| Orlando | Orlando / Orange / Seminole County |
| Pasco | Pasco County programs |
| Pinellas | Pinellas County programs |
| Polk | Polk County programs |
| PSH | All PSH programs (cross-site rollup) |
| San Juan | San Juan / Puerto Rico programs |
| Sarasota | Sarasota / Manatee County |
| Sebring | Sebring / Highlands County |
| SouthWest | Fort Myers / Lee County |
| Tampa | Tampa / Hillsborough County |

## Key Details

- **Days With no Service/Contact** is calculated from `Last Case Note Date Per Prog` in the Data Report Card (`TODAY() - Last Case Note Date Per Prog`). No separate Client Not Served report needed.
- **PQI Review / Peer Review** columns show `Yes`, `No`, or blank.

## Configuration

Edit the constants at the top of `caseload_report.py` to update:

- `NON_SSVF_PROGRAMS` — programs that get N/A for SOAR, ShallowSub, HUDVASH fields
- `INPUT_DIR_DATA_REPORT_CARD` / `INPUT_DIR_LEGAL_REFERRAL` — input folder paths
- Site tab filter rules (in `create_site_tabs()`)
