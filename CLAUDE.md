# Caseload Report Automation

## Overview

Python automation script that replaces a manual 30-45 minute Excel workflow for generating the St. Vincent de Paul CARES Caseload Report. The script takes two CaseWorthy Excel exports as input and produces a formatted, multi-tab workbook as output.

## Tech Stack

- **Language:** Python
- **Libraries:** `pandas`, `openpyxl`
- **Install:** `pip install pandas openpyxl`

## Input Files

The script accepts two `.xlsx` files exported from CaseWorthy, placed in separate input folders:

```
input/
├── data_report_card/          — Drop Data Report Card export here
│   └── Data_Report_Card_*.xlsx
└── legal_referral/            — Drop Legal Services Referral export here
    └── Legal_Services_Referral_*.xlsx
```

| Input | Folder | Description |
|-------|--------|-------------|
| Data Report Card | `input/data_report_card/` | Primary caseload data (68+ columns, one row per assessment event per client per program). Also provides `Last Case Note Date Per Prog` for service gap calculation. |
| Legal Services Referral | `input/legal_referral/` | Legal referral status — filter for 'Approved' only, convert Client ID to integer |

**Note:** The Client Not Served report is no longer needed. Days With no Service/Contact is now calculated directly from the `Last Case Note Date Per Prog` column in the Data Report Card as `TODAY() - Last Case Note Date Per Prog`.

## Output

- Filename pattern: `Current_Caseload_{YYYY-MM-DD}.xlsx`
- 16 tabs: `All` + 15 site-specific tabs (see Site Tabs section)
- 19-column standard layout (with exceptions for FOX and MidFlorida tabs)

## Script Structure

```
caseload_report.py
├── find_xlsx_in_folder(folder_path) → str
├── load_data_report_card(filepath) → (DataFrame, Series)
├── load_legal_referral(filepath) → DataFrame
├── process_main_sheet(drc, office_location, legal) → (DataFrame, Series)
├── create_site_tabs(df_all, office_location) → dict[str, DataFrame]
├── apply_formatting(workbook) → None
└── main(data_report_card, legal_referral, output_path)
```

## Processing Steps (Must Execute in Order)

1. **Load Data Report Card** — Read all columns. Rename: `Case Manager` → `Assigned Staff`, `Receiving Shallow Subsidy` → `Current Receive ShallowSub`, `Referred From HUD-VASH` → `Referred From HUDVASH`. Replace `At Exit` → `ZAT Exit` in Event column.
2. **Sort & Deduplicate** — Sort by Event, Client ID, Program Name (all A→Z). Drop duplicates on `[Client ID, Program Name]` keeping first (entry over exit).
3. **Drop Unused Columns** — Keep only 15 direct columns from Data Report Card (including `Last Case Note Date Per Prog`). Also drop Event, Position Type, End Date, Assign Begin/End Date.
4. **Calculate Last 90 Day Recert** — `TODAY() - Days since Last Recert/Update`. Format as MM/DD/YYYY. Insert at column 10.
5. **N/A Replacement for Non-SSVF** — Replace SOAR, ShallowSub, HUDVASH with `N/A`. Clear Recert fields. See Non-SSVF list below.
6. **VLOOKUP — Legal Assistance** — Match on Client ID. Write `Received` if found, else `N/A`.
7. **Calculate Days With no Service/Contact** — `TODAY() - Last Case Note Date Per Prog`. Blank if no case note date exists.
8. **Housed/Not Housed** — If Move-In Date blank → `Not Housed`, else `Housed`. Override to `N/A` for non-RRH programs (no `RRH` in Program Name).
9. **PQI Review & Peer Review** — `Yes` if value is present/yes, `No` if value is "No", blank if null/empty. These are the last two columns (18 & 19).
10. **Create Site Tabs** — Filter All sheet into 15 site tabs per mapping rules. FOX uses 12-column reduced layout. MidFlorida adds Location column at position 10.
11. **Apply Formatting** — Dark blue headers (#1B3A5C), white text, freeze row 1, auto-fit columns, MM/DD/YYYY dates, alternating row shading (#F2F6FA).

## Final Output Columns (19 Columns — "All" Tab)

| # | Column | Source |
|---|--------|--------|
| 1 | Client ID | Direct |
| 2 | First Name | Direct |
| 3 | Last Name | Direct |
| 4 | # Enrolled Family Members | Direct |
| 5 | Program Name | Direct |
| 6 | Begin Date | Direct |
| 7 | Days Enrolled | Direct |
| 8 | Move-In Date | Direct |
| 9 | Assigned Staff | Renamed from Case Manager |
| 10 | Last 90 Day Recert | Calculated: TODAY() - Days since Last Recert/Update |
| 11 | Days since Last Recert/Update | Direct |
| 12 | Current Receive ShallowSub | Renamed from Receiving Shallow Subsidy |
| 13 | Referred From HUDVASH | Renamed from Referred From HUD-VASH |
| 14 | Connection With SOAR | Direct |
| 15 | Received Legal Assistance | VLOOKUP from Legal Services Referral |
| 16 | Days With no Service/Contact | Calculated: TODAY() - Last Case Note Date Per Prog |
| 17 | Housed Not Housed | Calculated from Move-In Date |
| 18 | PQI Review | Yes, No, or blank |
| 19 | Peer Review | Yes, No, or blank |

## Non-SSVF Program List (Configurable)

These programs get `N/A` for SOAR, ShallowSub, HUDVASH, and Recert columns:

- `All-County-VA-Suicide-Prevention 1114`
- `Charlotte County-SHIP-RRH 1305`
- `Tampa-THHI-CDBG DAP CES 1107`

**Heuristic:** If Program Name does NOT contain `SSVF`, `EHA`, `GPD`, `HCHV`, `HUD-CoC`, `CoC`, `ESG`, or `PSH`, it is likely non-SSVF. This list should be maintained as a configurable constant at the top of the script.

## Site Tab Definitions

Tab order: All, Charlotte, Charlotte Shelter, FOX, GPD, MidFlorida, Orlando, Pasco, Pinellas, Polk, PSH, San Juan, Sarasota, Sebring, SouthWest, Tampa.

| Tab | Filter | Layout |
|-----|--------|--------|
| All | Every client | Full 19 columns |
| Charlotte | Starts with `Charlotte` AND NOT `Care Center` | Full 19 columns |
| Charlotte Shelter | Starts with `Charlotte` AND contains `Care Center` | Full 19 columns |
| FOX | `All-County-VA-Suicide` OR `Bob Woodruff` | **Reduced 12 columns** — drop Move-In, Recert, ShallowSub, HUDVASH, SOAR, Housed; keep PQI & Peer Review |
| GPD | Starts with `Pre-Housing` OR `Retention` | Full 19 columns |
| MidFlorida | Starts with `Lake Mid` OR `MidFlorida` OR `Citrus` OR `Hernando` OR `Sumter` | **Full 19 + Location column at position 10** (20 total) |
| Orlando | Starts with `Orlando` | Full 19 columns |
| Pasco | Starts with `Pasco` | Full 19 columns |
| Pinellas | Starts with `Pinellas` | Full 19 columns |
| Polk | Starts with `Polk` | Full 19 columns |
| PSH | Program Name contains `PSH` (any site) | Full 19 columns |
| San Juan | Starts with `San Juan` | Full 19 columns |
| Sarasota | Starts with `Sarasota` | Full 19 columns |
| Sebring | Starts with `Sebring` | Full 19 columns |
| SouthWest | Starts with `SouthWest` | Full 19 columns |
| Tampa | Starts with `Tampa` | Full 19 columns |

**Note:** PSH and FOX are cross-site rollups — clients appear on both their office tab and the rollup tab. Sum of site tab rows may exceed the All tab total.

## Key Implementation Notes

- **Days With no Service/Contact** is derived from `Last Case Note Date Per Prog` in the Data Report Card — no separate Client Not Served report needed
- **Deduplication** keeps first occurrence after sorting (entry records beat exit records due to ZAT Exit rename)
- **Legal Services join** is a left merge on Client ID (not all clients have referrals)
- **Site tab mapping** should be configurable via a dictionary of `{tab_name: filter_function}`
- **MidFlorida** needs the `Current Office Location` column from raw data (renamed to `Location`)
- **PQI/Peer Review** values are normalized to `Yes`, `No`, or blank
- **PII warning:** Output contains Client ID, First Name, Last Name — handle appropriately
- **Input folders** keep exports organized and prevent filename collisions from CaseWorthy

## Formatting Spec

| Element | Style |
|---------|-------|
| Header row | Bold, white text (#FFFFFF), dark blue background (#1B3A5C), frozen |
| Data rows | Black text (#333333), white background |
| Alternating rows | Light gray (#F2F6FA) |
| Date columns | MM/DD/YYYY format (Begin Date, Move-In Date, Last 90 Day Recert) |
| Column widths | Auto-fit, minimum 8 chars; Program Name ~40 chars |

## Error Handling

- Validate critical columns in Data Report Card (Event, Client ID, Program Name, Case Manager, Last Case Note Date Per Prog) — raise clear error if missing
- Log unmatched legal referral lookup counts for QA
- Log row count per site tab for verification
- Log any unmatched program name prefixes (clients only on All tab) as warnings
- Auto-detect .xlsx files in input folders; error if none found or multiple found

## Validation Checklist

- [ ] All tab contains all active clients after dedup
- [ ] No duplicate Client ID + Program Name on any tab
- [ ] Last 90 Day Recert dates are realistic (not future, not before 2020)
- [ ] Non-SSVF programs show N/A for SOAR, ShallowSub, HUDVASH
- [ ] Non-RRH programs show N/A for Housed/Not Housed
- [ ] FOX tab has exactly 12 columns (includes PQI Review and Peer Review)
- [ ] PQI Review: only 'Yes', 'No', or blank
- [ ] Peer Review: only 'Yes', 'No', or blank
- [ ] MidFlorida tab has 20 columns (19 + Location at position 10)
- [ ] Legal Assistance shows 'Received' only for approved referrals
- [ ] Days With no Service/Contact = TODAY() - Last Case Note Date Per Prog
- [ ] All dates formatted MM/DD/YYYY (no time components)
- [ ] Headers frozen on every tab

## Open Items

1. **New office/program detection** — Log unmatched prefixes as warnings (Medium priority)
2. **Conditional formatting for Recert** — Red (>=90 days), Yellow (70-89), Green (<70), Gray (none) (Low priority)
3. **Non-SSVF list maintenance** — Consider deriving from program name patterns vs hardcoded list (Medium priority)
4. **SSRS replacement (Phase 2)** — Long-term goal to replace with a single SSRS report (Future)

## GUI App

The app is built with tkinter and packaged via PyInstaller. See `documents/developer_guide.md` for build/run instructions and `documents/gui_design_reference.md` for the reusable design spec.

### App Footer

```
SVdP CARES · Data Systems · v1.0
⭐ Crafted by the legendary James Calby — Data Systems Analyst Extraordinaire ⭐
```
