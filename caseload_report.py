"""
Caseload Report Automation Script
St. Vincent de Paul CARES — Data Systems Team

Replaces the manual process of combining three CaseWorthy Excel exports
into a formatted, multi-tab Caseload Report workbook.

Usage:
    python caseload_report.py <data_report_card> <client_not_served> <legal_referral> [-o output.xlsx]
"""

import argparse
import logging
import sys
from collections import OrderedDict
from datetime import date, timedelta

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# CONFIGURABLE CONSTANTS
# ---------------------------------------------------------------------------

# Column renames applied to the Data Report Card
COLUMN_RENAMES = {
    "Case Manager": "Assigned Staff",
    "Receiving Shallow Subsidy": "Current Receive ShallowSub",
    "Referred From HUD-VASH": "Referred From HUDVASH",
}

# Columns kept from the Data Report Card after dropping unused ones (Step 3).
# These are the 14 direct columns that survive into the final output.
# "Current Office Location" is extracted separately for the MidFlorida tab.
KEEP_COLUMNS_FROM_DRC = [
    "Client ID",
    "First Name",
    "Last Name",
    "# Enrolled Family Members",
    "Program Name",
    "Begin Date",
    "Days Enrolled",
    "Move-In Date",
    "Assigned Staff",
    "Days since Last Recert/Update",
    "Current Receive ShallowSub",
    "Referred From HUDVASH",
    "Connection With SOAR",
    "PQI Review",
    "Peer Review",
]

# The 19 output columns in exact order for the "All" tab
ALL_COLUMNS_ORDERED = [
    "Client ID",
    "First Name",
    "Last Name",
    "# Enrolled Family Members",
    "Program Name",
    "Begin Date",
    "Days Enrolled",
    "Move-In Date",
    "Assigned Staff",
    "Last 90 Day Recert",
    "Days since Last Recert/Update",
    "Current Receive ShallowSub",
    "Referred From HUDVASH",
    "Connection With SOAR",
    "Received Legal Assistance",
    "Days With no Service/Contact",
    "Housed Not Housed",
    "PQI Review",
    "Peer Review",
]

# Critical columns that MUST exist in the Data Report Card
CRITICAL_DRC_COLUMNS = ["Event", "Client ID", "Program Name", "Case Manager"]

# Non-SSVF programs — these get N/A for SOAR, ShallowSub, HUDVASH, and Recert.
# Update this list as new non-SSVF programs are added to CaseWorthy.
NON_SSVF_PROGRAMS = [
    "All-County-VA-Suicide-Prevention 1114",
    "Charlotte County-SHIP-RRH 1305",
    "Tampa-THHI-CDBG DAP CES 1107",
]

# Keywords that indicate a program IS SSVF-funded (used for heuristic warnings)
SSVF_KEYWORDS = ["SSVF", "EHA", "GPD", "HCHV", "HUD-CoC", "CoC", "ESG", "PSH"]

# Tab order for the output workbook
SITE_TAB_ORDER = [
    "All",
    "Charlotte",
    "Charlotte Shelter",
    "FOX",
    "GPD",
    "MidFlorida",
    "Orlando",
    "Pasco",
    "Pinellas",
    "Polk",
    "PSH",
    "San Juan",
    "Sarasota",
    "Sebring",
    "SouthWest",
    "Tampa",
]

# Columns dropped from the FOX tab (reduced 12-column layout)
FOX_DROP_COLUMNS = [
    "Move-In Date",
    "Last 90 Day Recert",
    "Days since Last Recert/Update",
    "Current Receive ShallowSub",
    "Referred From HUDVASH",
    "Connection With SOAR",
    "Housed Not Housed",
]

# Formatting constants
HEADER_FILL = PatternFill(start_color="1B3A5C", end_color="1B3A5C", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ALT_ROW_FILL = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
DATE_FORMAT = "MM/DD/YYYY"
DATE_COLUMNS = {"Begin Date", "Move-In Date", "Last 90 Day Recert"}


# ---------------------------------------------------------------------------
# LOADER FUNCTIONS
# ---------------------------------------------------------------------------


def load_data_report_card(filepath):
    """Load and validate the Data Report Card export (Input 1).

    Returns:
        tuple: (DataFrame with all columns, Series of Current Office Location)
    """
    logger.info("Reading Data Report Card: %s", filepath)
    df = pd.read_excel(filepath, engine="openpyxl")
    logger.info("  Loaded %d rows, %d columns", len(df), len(df.columns))

    # Validate critical columns
    missing = [c for c in CRITICAL_DRC_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Data Report Card is missing critical columns: {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    # Extract Current Office Location before any processing (for MidFlorida tab)
    if "Current Office Location" in df.columns:
        office_location = df["Current Office Location"].copy()
    else:
        logger.warning("'Current Office Location' column not found — MidFlorida Location will be blank")
        office_location = pd.Series("", index=df.index)

    return df, office_location


def load_client_not_served(filepath):
    """Load the Client Not Served report (Input 2).

    Handles merged cells and variable header row position by using openpyxl
    to unmerge cells and auto-detect the header row containing 'Client ID'.

    Returns:
        DataFrame with columns: Client ID, Program Name, Days since Last Activity,
        Relationship to HoH
    """
    logger.info("Reading Client Not Served: %s", filepath)
    wb = load_workbook(filepath)
    ws = wb.active

    # Unmerge all merged cell ranges
    for merge_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merge_range))

    # Auto-detect header row by scanning for "Client ID"
    header_row = None
    for row in ws.iter_rows(min_row=1, max_row=50, max_col=20):
        for cell in row:
            if cell.value and str(cell.value).strip() == "Client ID":
                header_row = cell.row
                break
        if header_row:
            break

    if header_row is None:
        raise ValueError(
            "Could not find 'Client ID' header in Client Not Served report "
            "(searched rows 1-50). The file format may have changed."
        )

    logger.info("  Detected header at row %d", header_row)

    # Extract headers from the detected row
    headers = [cell.value for cell in ws[header_row]]

    # Extract data rows below the header
    data = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        # Stop at completely empty rows
        if all(v is None for v in row):
            continue
        data.append(row)

    df = pd.DataFrame(data, columns=headers)
    logger.info("  Loaded %d data rows", len(df))

    # Keep only required columns
    keep_cols = ["Client ID", "Program Name", "Days since Last Activity", "Relationship to HoH"]
    available = [c for c in keep_cols if c in df.columns]
    missing = [c for c in keep_cols if c not in df.columns]
    if missing:
        logger.warning("Client Not Served missing columns: %s", missing)

    df = df[available]

    # Convert Client ID to numeric
    if "Client ID" in df.columns:
        df["Client ID"] = pd.to_numeric(df["Client ID"], errors="coerce")
        df = df.dropna(subset=["Client ID"])
        df["Client ID"] = df["Client ID"].astype("Int64")

    wb.close()
    return df


def load_legal_referral(filepath):
    """Load the Legal Services Referral report (Input 3).

    Filters for Approved referrals, converts Client ID to integer,
    and relabels status as 'Received'.

    Returns:
        DataFrame with columns: CW Client ID, Referral Status
    """
    logger.info("Reading Legal Services Referral: %s", filepath)
    df = pd.read_excel(filepath, engine="openpyxl")
    logger.info("  Loaded %d rows", len(df))

    # Validate required columns
    required = ["CW Client ID", "Referral Status"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Legal Services Referral missing columns: {missing}")

    # Keep only needed columns
    df = df[required].copy()

    # Convert Client ID to numeric
    df["CW Client ID"] = pd.to_numeric(df["CW Client ID"], errors="coerce")
    df = df.dropna(subset=["CW Client ID"])
    df["CW Client ID"] = df["CW Client ID"].astype("Int64")

    # Filter for Approved only, then relabel
    df = df[df["Referral Status"] == "Approved"].copy()
    df["Referral Status"] = "Received"

    # Deduplicate — one marker per client is enough
    df = df.drop_duplicates(subset=["CW Client ID"], keep="first")

    logger.info("  %d clients with approved legal referrals", len(df))
    return df


# ---------------------------------------------------------------------------
# PROCESSING FUNCTIONS
# ---------------------------------------------------------------------------


def process_main_sheet(drc, office_location, cns, legal):
    """Process the main 'All' sheet through Steps 1-11 of the build doc.

    Args:
        drc: Raw Data Report Card DataFrame
        office_location: Series of Current Office Location values (aligned to drc index)
        cns: Cleaned Client Not Served DataFrame
        legal: Cleaned Legal Services Referral DataFrame

    Returns:
        tuple: (Processed All-tab DataFrame with 19 columns, aligned office_location Series)
    """
    df = drc.copy()

    # ── Step 1: Rename columns and replace Event values ──
    df.rename(columns=COLUMN_RENAMES, inplace=True)
    df["Event"] = df["Event"].replace("At Exit", "ZAT Exit")

    # ── Step 2: Sort and deduplicate ──
    df.sort_values(by=["Event", "Client ID", "Program Name"], inplace=True)
    before_count = len(df)
    df.drop_duplicates(subset=["Client ID", "Program Name"], keep="first", inplace=True)
    df.reset_index(drop=False, inplace=True)  # preserve original index for office_location alignment
    logger.info(
        "Dedup: %d -> %d rows (%d duplicates removed)",
        before_count,
        len(df),
        before_count - len(df),
    )

    # Align office_location to the surviving rows
    if "index" in df.columns:
        office_loc_aligned = office_location.reindex(df["index"]).reset_index(drop=True)
        df.drop(columns=["index"], inplace=True)
    else:
        office_loc_aligned = pd.Series("", index=df.index)

    # ── Step 3: Drop unused columns ──
    available_keep = [c for c in KEEP_COLUMNS_FROM_DRC if c in df.columns]
    df = df[available_keep].copy()
    df.reset_index(drop=True, inplace=True)

    # ── Step 4: Calculate Last 90 Day Recert ──
    today = date.today()

    def calc_recert(days_val):
        if pd.isna(days_val):
            return pd.NaT
        try:
            return today - timedelta(days=int(days_val))
        except (ValueError, TypeError):
            return pd.NaT

    df["Last 90 Day Recert"] = df["Days since Last Recert/Update"].apply(calc_recert)

    # Insert at position 9 (after Assigned Staff, before Days since Last Recert/Update)
    cols = list(df.columns)
    cols.remove("Last 90 Day Recert")
    cols.insert(9, "Last 90 Day Recert")
    df = df[cols]

    # ── Step 5: N/A replacement for non-SSVF programs ──
    non_ssvf_mask = df["Program Name"].isin(NON_SSVF_PROGRAMS)

    na_columns = ["Current Receive ShallowSub", "Referred From HUDVASH", "Connection With SOAR"]
    for col in na_columns:
        if col in df.columns:
            df.loc[non_ssvf_mask, col] = "N/A"

    df.loc[non_ssvf_mask, "Last 90 Day Recert"] = pd.NaT
    df.loc[non_ssvf_mask, "Days since Last Recert/Update"] = pd.NA

    # Heuristic warning: flag programs that don't match any known SSVF keyword
    # and aren't in the explicit non-SSVF list
    all_programs = df["Program Name"].unique()
    for prog in all_programs:
        if prog in NON_SSVF_PROGRAMS:
            continue
        has_keyword = any(kw in str(prog) for kw in SSVF_KEYWORDS)
        if not has_keyword:
            logger.warning(
                "Program '%s' does not match any SSVF keyword and is not in "
                "NON_SSVF_PROGRAMS — verify if N/A replacement is needed",
                prog,
            )

    # ── Step 8: VLOOKUP — Received Legal Assistance ──
    legal_lookup = legal.rename(
        columns={"CW Client ID": "Client ID", "Referral Status": "Received Legal Assistance"}
    )
    df = df.merge(legal_lookup, on="Client ID", how="left")

    unmatched_legal = df["Received Legal Assistance"].isna().sum()
    matched_legal = len(df) - unmatched_legal
    logger.info("Legal referral: %d matched, %d unmatched -> N/A", matched_legal, unmatched_legal)
    df["Received Legal Assistance"] = df["Received Legal Assistance"].fillna("N/A")

    # ── Step 9: INDEX/MATCH — Days With no Service/Contact ──
    cns_lookup = cns.rename(
        columns={"Days since Last Activity": "Days With no Service/Contact"}
    )
    # Deduplicate CNS on compound key before merging
    if "Days With no Service/Contact" in cns_lookup.columns:
        cns_lookup = cns_lookup.drop_duplicates(subset=["Client ID", "Program Name"], keep="first")
        df = df.merge(
            cns_lookup[["Client ID", "Program Name", "Days With no Service/Contact"]],
            on=["Client ID", "Program Name"],
            how="left",
        )
        unmatched_cns = df["Days With no Service/Contact"].isna().sum()
        logger.info(
            "Client Not Served: %d matched, %d unmatched",
            len(df) - unmatched_cns,
            unmatched_cns,
        )
    else:
        df["Days With no Service/Contact"] = pd.NA
        logger.warning("Could not match Days With no Service/Contact — column missing from CNS data")

    # ── Step 10: Housed / Not Housed ──
    df["Housed Not Housed"] = df["Move-In Date"].apply(
        lambda x: "Housed" if pd.notna(x) else "Not Housed"
    )
    non_rrh_mask = ~df["Program Name"].str.contains("RRH", case=False, na=False)
    df.loc[non_rrh_mask, "Housed Not Housed"] = "N/A"

    # ── Step 11: PQI Review and Peer Review ──
    for col in ["PQI Review", "Peer Review"]:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: "Yes" if pd.notna(x) and str(x).strip() != "" else ""
            )

    # ── Final column ordering ──
    final_cols = [c for c in ALL_COLUMNS_ORDERED if c in df.columns]
    df = df[final_cols]

    logger.info("All tab: %d rows, %d columns", len(df), len(df.columns))
    return df, office_loc_aligned


def create_site_tabs(df_all, office_location):
    """Create the 16 site tabs from the All sheet data (Step 12).

    Args:
        df_all: Processed All-tab DataFrame (19 columns)
        office_location: Series of Location values aligned to df_all

    Returns:
        OrderedDict of {tab_name: DataFrame} in SITE_TAB_ORDER
    """
    pn = df_all["Program Name"].fillna("")

    # Define filter masks for each site tab
    filters = {
        "All": pd.Series(True, index=df_all.index),
        "Charlotte": pn.str.startswith("Charlotte") & ~pn.str.contains("Care Center"),
        "Charlotte Shelter": pn.str.startswith("Charlotte") & pn.str.contains("Care Center"),
        "FOX": pn.str.startswith("All-County-VA-Suicide") | pn.str.startswith("Bob Woodruff"),
        "GPD": pn.str.startswith("Pre-Housing") | pn.str.startswith("Retention"),
        "MidFlorida": (
            pn.str.startswith("Lake Mid")
            | pn.str.startswith("MidFlorida")
            | pn.str.startswith("Citrus")
            | pn.str.startswith("Hernando")
            | pn.str.startswith("Sumter")
        ),
        "Orlando": pn.str.startswith("Orlando"),
        "Pasco": pn.str.startswith("Pasco"),
        "Pinellas": pn.str.startswith("Pinellas"),
        "Polk": pn.str.startswith("Polk"),
        "PSH": pn.str.contains("PSH", case=False, na=False),
        "San Juan": pn.str.startswith("San Juan"),
        "Sarasota": pn.str.startswith("Sarasota"),
        "Sebring": pn.str.startswith("Sebring"),
        "SouthWest": pn.str.startswith("SouthWest"),
        "Tampa": pn.str.startswith("Tampa"),
    }

    tabs = OrderedDict()

    for tab_name in SITE_TAB_ORDER:
        mask = filters[tab_name]
        tab_df = df_all[mask].copy()
        tab_df.sort_values(by="Program Name", inplace=True)
        tab_df.reset_index(drop=True, inplace=True)

        if tab_name == "FOX":
            # Reduced 12-column layout — drop 7 columns
            drop_cols = [c for c in FOX_DROP_COLUMNS if c in tab_df.columns]
            tab_df = tab_df.drop(columns=drop_cols)

        elif tab_name == "MidFlorida":
            # Add Location column at position 9 (column 10 in 1-indexed)
            loc_values = office_location.reindex(df_all[mask].index).reset_index(drop=True)
            tab_df.insert(9, "Location", loc_values)

        tabs[tab_name] = tab_df
        logger.info("  %-20s %d rows, %d columns", tab_name, len(tab_df), len(tab_df.columns))

    # Warn about programs that don't match any site filter (only on All tab)
    matched = pd.Series(False, index=df_all.index)
    for name, mask in filters.items():
        if name != "All":
            matched = matched | mask
    unmatched = df_all[~matched]
    if len(unmatched) > 0:
        unmatched_programs = unmatched["Program Name"].unique()
        logger.warning(
            "%d program(s) appear only on the All tab (no site match): %s",
            len(unmatched_programs),
            list(unmatched_programs),
        )

    return tabs


def apply_formatting(wb):
    """Apply visual formatting to the output workbook (Step 13).

    Args:
        wb: openpyxl Workbook object (already written with data)
    """
    for ws in wb.worksheets:
        # Format header row
        for cell in ws[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center")

        # Freeze top row
        ws.freeze_panes = "A2"

        # Auto-fit column widths and format date columns
        for col_idx, col_cells in enumerate(ws.columns, 1):
            col_cells = list(col_cells)
            header_text = str(col_cells[0].value or "")

            # Calculate width
            max_len = max(len(str(cell.value or "")) for cell in col_cells)
            max_len = max(max_len, 8)  # minimum 8 characters
            if header_text == "Program Name":
                max_len = max(max_len, 40)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 50)

            # Format date columns
            if header_text in DATE_COLUMNS:
                for cell in col_cells[1:]:
                    if cell.value is not None:
                        cell.number_format = DATE_FORMAT

        # Alternating row shading
        for row_idx in range(2, ws.max_row + 1):
            if row_idx % 2 == 0:
                for cell in ws[row_idx]:
                    cell.fill = ALT_ROW_FILL


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------


def main(input1, input2, input3, output_path=None):
    """Run the full Caseload Report pipeline.

    Args:
        input1: Path to Data_Report_Card_*.xlsx
        input2: Path to Client_Not_Served_*.xlsx
        input3: Path to Legal_Services_Referral_*.xlsx
        output_path: Output file path (default: Current_Caseload_{date}.xlsx)
    """
    if output_path is None:
        output_path = f"Current_Caseload_{date.today().isoformat()}.xlsx"

    # Load all three input files
    drc, office_location = load_data_report_card(input1)
    cns = load_client_not_served(input2)
    legal = load_legal_referral(input3)

    # Process the main sheet (Steps 1-11)
    logger.info("Processing main sheet...")
    df_all, office_loc_aligned = process_main_sheet(drc, office_location, cns, legal)

    # Create site tabs (Step 12)
    logger.info("Creating site tabs...")
    tabs = create_site_tabs(df_all, office_loc_aligned)

    # Write to Excel
    logger.info("Writing output to %s", output_path)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for tab_name, tab_df in tabs.items():
            tab_df.to_excel(writer, sheet_name=tab_name, index=False)

    # Reopen for formatting (Step 13)
    logger.info("Applying formatting...")
    wb = load_workbook(output_path)
    apply_formatting(wb)
    wb.save(output_path)

    logger.info("Done. Output: %s", output_path)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Generate SVdP CARES Caseload Report from CaseWorthy exports.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python caseload_report.py Data_Report_Card.xlsx Client_Not_Served.xlsx Legal_Referral.xlsx\n"
            "  python caseload_report.py input1.xlsx input2.xlsx input3.xlsx -o My_Report.xlsx\n"
        ),
    )
    parser.add_argument("data_report_card", help="Path to Data_Report_Card_*.xlsx")
    parser.add_argument("client_not_served", help="Path to Client_Not_Served_*.xlsx")
    parser.add_argument("legal_referral", help="Path to Legal_Services_Referral_*.xlsx")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Output file path (default: Current_Caseload_{YYYY-MM-DD}.xlsx)",
    )
    args = parser.parse_args()

    main(args.data_report_card, args.client_not_served, args.legal_referral, args.output)
