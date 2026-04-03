"""
PyInstaller build script for Caseload Report Generator.

Usage:
    pip install pyinstaller
    python build_exe.py

Output:
    dist/CaseloadReport.exe
"""

import os
import subprocess
import sys


def build():
    script_dir = os.path.dirname(os.path.abspath(__file__))

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", "CaseloadReport",
        # Add the processing module as hidden import
        "--hidden-import", "caseload_report",
        # Include config files
        "--add-data", f"{os.path.join(script_dir, 'config')};config",
        # Entry point
        os.path.join(script_dir, "app.py"),
    ]

    print("Building CaseloadReport.exe...")
    print(f"Command: {' '.join(cmd)}")
    subprocess.run(cmd, check=True)

    # Post-build: remind user to copy folders
    print()
    print("=" * 60)
    print("BUILD COMPLETE")
    print("=" * 60)
    print()
    print("Output: dist/CaseloadReport.exe")
    print()
    print("To distribute, create a folder with:")
    print()
    print("  CaseloadReport/")
    print("  ├── CaseloadReport.exe")
    print("  ├── config/")
    print("  │   ├── staff_roster/       <- drop roster .xlsx here")
    print("  │   └── site_tab_mapping.csv")
    print("  ├── input/")
    print("  │   ├── data_report_card/   <- drop DRC export here")
    print("  │   └── legal_referral/     <- drop Legal export here")
    print("  └── output/                 <- reports saved here")
    print()
    print("Copy the config/, input/, and output/ folders next to the .exe.")


if __name__ == "__main__":
    build()
