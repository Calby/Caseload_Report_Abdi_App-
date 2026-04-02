# Caseload Report Automation

Python script that automates the SVdP CARES Caseload Report — replaces the manual process of combining three CaseWorthy Excel exports into a formatted, multi-tab workbook.

## Setup

```bash
pip install -r requirements.txt
```

Requires Python 3.9+ with `pandas` and `openpyxl`.

## Usage

```bash
python caseload_report.py <data_report_card> <client_not_served> <legal_referral> [-o output.xlsx]
```

### Arguments

| Argument | Description |
|----------|-------------|
| `data_report_card` | Path to `Data_Report_Card_*.xlsx` (68-column export) |
| `client_not_served` | Path to `Client_Not_Served_*.xlsx` (merged-cell report) |
| `legal_referral` | Path to `Legal_Services_Referral_*.xlsx` |
| `-o, --output` | Output file path (default: `Current_Caseload_{YYYY-MM-DD}.xlsx`) |

### Example

```bash
python caseload_report.py \
  Data_Report_Card_2026-04-01.xlsx \
  Client_Not_Served_2026-04-01.xlsx \
  Legal_Services_Referral_2026-04-01.xlsx \
  -o Current_Caseload_2026-04-01.xlsx
```

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

## Configuration

Edit the constants at the top of `caseload_report.py` to update:

- `NON_SSVF_PROGRAMS` — programs that get N/A for SOAR, ShallowSub, HUDVASH fields
- Site tab filter rules (in `create_site_tabs()`)
