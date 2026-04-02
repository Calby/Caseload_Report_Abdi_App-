# Caseload Report Automation

## Overview

Python automation script that replaces a manual 30-45 minute Excel workflow for generating the St. Vincent de Paul CARES Caseload Report. The script takes three CaseWorthy Excel exports as input and produces a formatted, multi-tab workbook as output.

## Tech Stack

- **Language:** Python
- **Libraries:** `pandas`, `openpyxl`
- **Install:** `pip install pandas openpyxl`

## Input Files

The script accepts three `.xlsx` files exported from CaseWorthy:

| Input | Pattern | Description |
|-------|---------|-------------|
| Input 1 | `Data_Report_Card_*.xlsx` | Primary caseload data (68 columns, one row per assessment event per client per program) |
| Input 2 | `Client_Not_Served_*.xlsx` | Days since last activity per client/program. Has merged cells (rows 1-25), hidden columns — requires preprocessing |
| Input 3 | `Legal_Services_Referral_*.xlsx` | Legal referral status — filter for 'Approved' only, convert Client ID to integer |

## Output

- Filename pattern: `Current_Caseload_{YYYY-MM-DD}.xlsx`
- 16 tabs: `All` + 15 site-specific tabs (see Site Tabs section)
- 19-column standard layout (with exceptions for FOX and MidFlorida tabs)

## Script Structure

```
caseload_report.py
├── load_data_report_card(filepath) → DataFrame
├── load_client_not_served(filepath) → DataFrame
├── load_legal_referral(filepath) → DataFrame
├── process_main_sheet(drc, cns, legal) → DataFrame
├── create_site_tabs(df_all) → dict[str, DataFrame]
├── apply_formatting(workbook) → None
└── main(input1, input2, input3, output_path)
```

## Processing Steps (Must Execute in Order)

1. **Load Data Report Card** — Read all 68 columns. Rename: `Case Manager` → `Assigned Staff`, `Receiving Shallow Subsidy` → `Current Receive ShallowSub`, `Referred From HUD-VASH` → `Referred From HUDVASH`. Replace `At Exit` → `ZAT Exit` in Event column.
2. **Sort & Deduplicate** — Sort by Event, Client ID, Program Name (all A→Z). Drop duplicates on `[Client ID, Program Name]` keeping first (entry over exit).
3. **Drop Unused Columns** — Keep only 14 direct columns from Data Report Card. Also drop Event, Position Type, End Date, Assign Begin/End Date.
4. **Calculate Last 90 Day Recert** — `TODAY() - Days since Last Recert/Update`. Format as MM/DD/YYYY. Insert at column 10.
5. **N/A Replacement for Non-SSVF** — Replace SOAR, ShallowSub, HUDVASH with `N/A`. Clear Recert fields. See Non-SSVF list below.
6. **Load Client Not Served** — Unmerge cells, strip header rows (1-25), unhide columns. Keep: Client ID, Program Name, Days since Last Activity, Relationship to HoH.
7. **Load Legal Services Referral** — Keep: CW Client ID, Referral Status. Filter for `Approved`, change label to `Received`.
8. **VLOOKUP — Legal Assistance** — Match on Client ID. Write `Received` if found, else `N/A`.
9. **INDEX/MATCH — Days With no Service/Contact** — Match on BOTH Client ID AND Program Name (compound key). Return Days since Last Activity.
10. **Housed/Not Housed** — If Move-In Date blank → `Not Housed`, else `Housed`. Override to `N/A` for non-RRH programs (no `RRH` in Program Name).
11. **PQI Review & Peer Review** — If any non-blank value → `Yes`, else leave blank. These are the last two columns (18 & 19).
12. **Create Site Tabs** — Filter All sheet into 15 site tabs per mapping rules. FOX uses 12-column reduced layout. MidFlorida adds Location column at position 10.
13. **Apply Formatting** — Dark blue headers (#1B3A5C), white text, freeze row 1, auto-fit columns, MM/DD/YYYY dates, optional alternating row shading (#F2F6FA).

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
| 16 | Days With no Service/Contact | INDEX/MATCH from Client Not Served |
| 17 | Housed Not Housed | Calculated from Move-In Date |
| 18 | PQI Review | Yes if value present, else blank |
| 19 | Peer Review | Yes if value present, else blank |

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

- **Client Not Served** has merged cells and hidden columns — use `openpyxl` to unmerge before loading into pandas
- **Deduplication** keeps first occurrence after sorting (entry records beat exit records due to ZAT Exit rename)
- **Legal Services join** is a left merge on Client ID (not all clients have referrals)
- **Client Not Served join** is on compound key: Client ID + Program Name
- **Site tab mapping** should be configurable via a dictionary of `{tab_name: filter_function}`
- **MidFlorida** needs the `Current Office Location` column from raw data (renamed to `Location`)
- **PII warning:** Output contains Client ID, First Name, Last Name — handle appropriately

## Formatting Spec

| Element | Style |
|---------|-------|
| Header row | Bold, white text (#FFFFFF), dark blue background (#1B3A5C), frozen |
| Data rows | Black text (#333333), white background |
| Alternating rows | Optional light gray (#F2F6FA) |
| Date columns | MM/DD/YYYY format (Begin Date, Move-In Date, Last 90 Day Recert) |
| Column widths | Auto-fit, minimum 8 chars; Program Name ~40 chars |

## Error Handling

- Validate all 68 expected columns in Data Report Card — raise clear error if missing/renamed
- Warn (don't fail) if Client Not Served has fewer than 25 header rows
- Log unmatched legal referral and client-not-served lookup counts for QA
- Log row count per site tab for verification
- Log any unmatched program name prefixes (clients only on All tab) as warnings

## Validation Checklist

- [ ] All tab contains all active clients after dedup
- [ ] No duplicate Client ID + Program Name on any tab
- [ ] Last 90 Day Recert dates are realistic (not future, not before 2020)
- [ ] Non-SSVF programs show N/A for SOAR, ShallowSub, HUDVASH
- [ ] Non-RRH programs show N/A for Housed/Not Housed
- [ ] FOX tab has exactly 12 columns (includes PQI Review and Peer Review)
- [ ] PQI Review: only 'Yes' or blank
- [ ] Peer Review: only 'Yes' or blank
- [ ] MidFlorida tab has 20 columns (19 + Location at position 10)
- [ ] Legal Assistance shows 'Received' only for approved referrals
- [ ] Days With no Service/Contact matches source data
- [ ] All dates formatted MM/DD/YYYY (no time components)
- [ ] Headers frozen on every tab

## Open Items

1. **New office/program detection** — Log unmatched prefixes as warnings (Medium priority)
2. **Conditional formatting for Recert** — Red (>=90 days), Yellow (70-89), Green (<70), Gray (none) (Low priority)
3. **Auto-detect Client Not Served header row** — Find 'Client ID' text instead of assuming row 25 (High priority)
4. **Non-SSVF list maintenance** — Consider deriving from program name patterns vs hardcoded list (Medium priority)
5. **SSRS replacement (Phase 2)** — Long-term goal to replace with a single SSRS report (Future)
