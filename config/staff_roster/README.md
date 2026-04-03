Place your **CaseWorthy User Export** (.xlsx) here.

This file maps staff login names to their actual names and job types.

**Required columns:** `Login Name`, `Last Name`, `First Name`, `Job Type`

The script uses `Login Name` to look up the assigned staff on the report
and replaces the abbreviation with their full name (Last Name, First Name)
and adds their Job Type.

If no file is placed here, the script will use the raw staff names from
the Data Report Card as-is.
