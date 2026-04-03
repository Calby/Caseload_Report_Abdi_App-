# Configuration Files

This folder contains lookup/mapping files used by the Caseload Report script.

## Staff Roster (`staff_roster/`)

Drop the CaseWorthy User Export (.xlsx) here. The script maps login names
to full staff names and job types.

Required columns: `Login Name`, `Last Name`, `First Name`, `Job Type`

## Program-to-Site Mapping (`site_tab_mapping.csv`)

Maps program name patterns to site tabs. Edit this CSV to add new programs
or change which tab a program appears on.

Columns: `tab_name`, `match_type`, `match_value`

- `match_type: prefix` — matches programs starting with `match_value`
- `match_type: contains` — matches programs containing `match_value`
- `match_type: location_contains` — matches Current Office Location containing `match_value`
