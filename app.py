"""
Caseload Report Generator — GUI Application
St. Vincent de Paul CARES — Data Systems Team

Tkinter GUI wrapper around caseload_report.py processing engine.
Allows users to select input files, configure options, and generate
the Caseload Report workbook with a single click.

Usage:
    python app.py
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Path resolution (PyInstaller compatible)
# ---------------------------------------------------------------------------


def get_app_dir():
    """Get the directory where the app/exe lives."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = get_app_dir()

# ---------------------------------------------------------------------------
# Colors & Fonts
# ---------------------------------------------------------------------------

HEADER_BG = "#1F4E79"
HEADER_FG = "white"
SUBTITLE_FG = "#B0C4DE"
BODY_BG = "#F0F4F8"
FOOTER_BG = "#E8EDF2"
BTN_BG = "#1F4E79"
BTN_HOVER = "#2E75B6"
BTN_FG = "white"
SUCCESS_FG = "#2E7D32"
ERROR_FG = "#C62828"
MUTED_FG = "#555"

FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SUBTITLE = ("Segoe UI", 10)
FONT_LABEL = ("Segoe UI", 10, "bold")
FONT_ENTRY = ("Segoe UI", 9)
FONT_BUTTON = ("Segoe UI", 11, "bold")
FONT_SMALL_BTN = ("Segoe UI", 9)
FONT_STATUS = ("Segoe UI", 9)
FONT_FOOTER = ("Segoe UI", 8, "italic")
FONT_INDICATOR = ("Segoe UI", 9)

APP_VERSION = "1.0"


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------


class CaseloadReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Caseload Report Generator")
        self.root.resizable(False, False)
        self.root.configure(bg=BODY_BG)

        # Variables
        self.drc_path = tk.StringVar()
        self.legal_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.path.join(APP_DIR, "output"))
        self.roster_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")
        self.processing = False

        # Auto-detect defaults
        self._auto_detect_defaults()

        # Build UI
        self._build_header()
        self._build_body()
        self._build_footer()

        # Center window on screen
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"+{x}+{y}")

    # ── Auto-detect ──────────────────────────────────────────────────────

    def _auto_detect_defaults(self):
        """Pre-fill paths from default folders if files exist."""
        # Data Report Card
        drc_dir = os.path.join(APP_DIR, "input", "data_report_card")
        found = self._find_xlsx(drc_dir)
        if found:
            self.drc_path.set(found)

        # Legal Referral
        legal_dir = os.path.join(APP_DIR, "input", "legal_referral")
        found = self._find_xlsx(legal_dir)
        if found:
            self.legal_path.set(found)

        # Staff Roster (optional)
        roster_dir = os.path.join(APP_DIR, "config", "staff_roster")
        found = self._find_xlsx(roster_dir)
        if found:
            self.roster_path.set(found)

    @staticmethod
    def _find_xlsx(folder):
        """Find a single .xlsx in a folder, ignoring temp files."""
        if not os.path.isdir(folder):
            return None
        import glob
        files = glob.glob(os.path.join(folder, "*.xlsx"))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        if len(files) == 1:
            return files[0]
        return None

    # ── Header ───────────────────────────────────────────────────────────

    def _build_header(self):
        header = tk.Frame(self.root, bg=HEADER_BG, padx=20, pady=14)
        header.pack(fill="x")

        tk.Label(
            header, text="Caseload Report Generator",
            font=FONT_TITLE, fg=HEADER_FG, bg=HEADER_BG
        ).pack()

        tk.Label(
            header, text="SVdP CARES \u2014 Data Systems Team",
            font=FONT_SUBTITLE, fg=SUBTITLE_FG, bg=HEADER_BG
        ).pack()

    # ── Body ─────────────────────────────────────────────────────────────

    def _build_body(self):
        body = tk.Frame(self.root, bg=BODY_BG, padx=24, pady=16)
        body.pack(fill="both", expand=True)

        row = 0

        # --- Data Report Card ---
        row = self._add_file_row(
            body, row, "Data Report Card", self.drc_path,
            self._browse_drc, required=True
        )

        # --- Legal Services Referral ---
        row = self._add_file_row(
            body, row, "Legal Services Referral", self.legal_path,
            self._browse_legal, required=True
        )

        # --- Output Folder ---
        tk.Label(
            body, text="Output Folder", font=FONT_LABEL,
            bg=BODY_BG, anchor="w"
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(12, 2))
        row += 1

        out_frame = tk.Frame(body, bg=BODY_BG)
        out_frame.grid(row=row, column=0, columnspan=2, sticky="ew")
        tk.Entry(
            out_frame, textvariable=self.output_dir, width=52,
            font=FONT_ENTRY
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            out_frame, text="Browse...", command=self._browse_output,
            font=FONT_SMALL_BTN
        ).pack(side="left")
        row += 1

        # --- Separator ---
        sep = ttk.Separator(body, orient="horizontal")
        sep.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(16, 8))
        row += 1

        # --- Staff Roster (optional) ---
        tk.Label(
            body, text="Staff Roster (optional)", font=FONT_LABEL,
            bg=BODY_BG, anchor="w"
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(4, 2))
        row += 1

        roster_frame = tk.Frame(body, bg=BODY_BG)
        roster_frame.grid(row=row, column=0, columnspan=2, sticky="ew")
        tk.Entry(
            roster_frame, textvariable=self.roster_path, width=52,
            font=FONT_ENTRY
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            roster_frame, text="Browse...", command=self._browse_roster,
            font=FONT_SMALL_BTN
        ).pack(side="left")
        row += 1

        # Roster status indicator
        self.roster_indicator = tk.Label(
            body, text="", font=FONT_INDICATOR, bg=BODY_BG, anchor="w"
        )
        self.roster_indicator.grid(row=row, column=0, columnspan=2, sticky="w", pady=(2, 0))
        self._update_roster_indicator()
        self.roster_path.trace_add("write", lambda *_: self._update_roster_indicator())
        row += 1

        # --- Progress bar ---
        self.progress = ttk.Progressbar(body, mode="indeterminate", length=420)
        self.progress.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(16, 4))
        row += 1

        # Status text
        self.status_label = tk.Label(
            body, textvariable=self.status_var, font=FONT_STATUS,
            bg=BODY_BG, fg=MUTED_FG, anchor="w"
        )
        self.status_label.grid(row=row, column=0, columnspan=2, sticky="w")
        row += 1

        # --- Generate button ---
        btn_frame = tk.Frame(body, bg=BODY_BG)
        btn_frame.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))

        self.run_btn = tk.Button(
            btn_frame, text="Generate Report",
            command=self._on_generate, font=FONT_BUTTON,
            bg=BTN_BG, fg=BTN_FG, padx=20, pady=6,
            activebackground=BTN_HOVER, activeforeground=BTN_FG,
            cursor="hand2"
        )
        self.run_btn.pack(side="right")

        # Hover effect
        self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg=BTN_HOVER))
        self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg=BTN_BG))

    def _add_file_row(self, parent, row, label_text, var, browse_fn, required=False):
        """Add a label + entry + browse button row. Returns next row number."""
        suffix = "" if not required else ""
        tk.Label(
            parent, text=f"{label_text}{suffix}", font=FONT_LABEL,
            bg=BODY_BG, anchor="w"
        ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(12, 2))
        row += 1

        frame = tk.Frame(parent, bg=BODY_BG)
        frame.grid(row=row, column=0, columnspan=2, sticky="ew")
        tk.Entry(
            frame, textvariable=var, width=52, font=FONT_ENTRY
        ).pack(side="left", padx=(0, 8))
        tk.Button(
            frame, text="Browse...", command=browse_fn, font=FONT_SMALL_BTN
        ).pack(side="left")
        row += 1

        return row

    # ── Footer ───────────────────────────────────────────────────────────

    def _build_footer(self):
        footer = tk.Frame(self.root, bg=FOOTER_BG, padx=20, pady=8)
        footer.pack(fill="x", side="bottom")

        tk.Label(
            footer,
            text=f"SVdP CARES \u00b7 Data Systems \u00b7 v{APP_VERSION}",
            font=FONT_FOOTER, fg="#666", bg=FOOTER_BG
        ).pack()

    # ── Browse handlers ──────────────────────────────────────────────────

    def _browse_drc(self):
        path = filedialog.askopenfilename(
            title="Select Data Report Card",
            initialdir=os.path.join(APP_DIR, "input", "data_report_card"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.drc_path.set(path)

    def _browse_legal(self):
        path = filedialog.askopenfilename(
            title="Select Legal Services Referral",
            initialdir=os.path.join(APP_DIR, "input", "legal_referral"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.legal_path.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_dir.get() or APP_DIR
        )
        if path:
            self.output_dir.set(path)

    def _browse_roster(self):
        path = filedialog.askopenfilename(
            title="Select Staff Roster (optional)",
            initialdir=os.path.join(APP_DIR, "config", "staff_roster"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.roster_path.set(path)

    # ── Indicators ───────────────────────────────────────────────────────

    def _update_roster_indicator(self):
        path = self.roster_path.get()
        if path and os.path.isfile(path):
            self.roster_indicator.config(
                text="\u2713 Staff roster loaded",
                fg=SUCCESS_FG
            )
        else:
            self.roster_indicator.config(
                text="\u26a0 No staff roster \u2014 using raw staff names",
                fg=MUTED_FG
            )

    # ── Generate ─────────────────────────────────────────────────────────

    def _on_generate(self):
        """Validate inputs and launch processing in a background thread."""
        # Validate required fields
        drc = self.drc_path.get().strip()
        legal = self.legal_path.get().strip()
        output = self.output_dir.get().strip()

        errors = []
        if not drc or not os.path.isfile(drc):
            errors.append("Data Report Card file is required.")
        if not legal or not os.path.isfile(legal):
            errors.append("Legal Services Referral file is required.")
        if not output:
            errors.append("Output folder is required.")

        if errors:
            messagebox.showerror("Missing Input", "\n".join(errors))
            return

        # Create output folder if needed
        os.makedirs(output, exist_ok=True)

        # Disable button and start progress
        self.run_btn.config(state="disabled")
        self.run_btn.unbind("<Enter>")
        self.run_btn.unbind("<Leave>")
        self.progress.start(15)
        self.status_var.set("Processing...")
        self.status_label.config(fg=MUTED_FG)
        self.processing = True

        # Build output path
        from datetime import date
        output_path = os.path.join(output, f"Current_Caseload_{date.today().isoformat()}.xlsx")

        # Staff roster path (optional)
        roster = self.roster_path.get().strip()
        if not roster or not os.path.isfile(roster):
            roster = None

        # Run in background thread
        thread = threading.Thread(
            target=self._process,
            args=(drc, legal, output_path, roster),
            daemon=True
        )
        thread.start()

    def _process(self, drc_path, legal_path, output_path, roster_path):
        """Background thread: run the report processing engine."""
        try:
            # Import here to avoid circular imports and keep startup fast
            import caseload_report as cr

            # Load inputs
            drc, office_location = cr.load_data_report_card(drc_path)
            legal = cr.load_legal_referral(legal_path)

            # Load config
            if roster_path:
                # Temporarily override the config dir to load from the selected file
                staff_roster = self._load_roster_from_path(roster_path)
            else:
                staff_roster = cr.load_staff_roster()

            site_tab_mapping = cr.load_site_tab_mapping()

            # Process
            df_all, office_loc_aligned = cr.process_main_sheet(
                drc, office_location, legal, staff_roster
            )

            # Create tabs
            tabs = cr.create_site_tabs(df_all, office_loc_aligned, site_tab_mapping)

            # Write output
            import pandas as pd
            from openpyxl import load_workbook

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for tab_name, tab_df in tabs.items():
                    tab_df.to_excel(writer, sheet_name=tab_name, index=False)

            wb = load_workbook(output_path)
            cr.apply_formatting(wb)
            wb.save(output_path)

            self.root.after(0, self._on_success, output_path)

        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    @staticmethod
    def _load_roster_from_path(filepath):
        """Load staff roster from a specific file path."""
        import pandas as pd

        df = pd.read_excel(filepath, engine="openpyxl")
        df.columns = df.columns.str.strip()

        required = ["Login Name", "Last Name", "First Name"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            return None

        roster = {}
        for _, row in df.iterrows():
            login = str(row["Login Name"]).strip()
            if not login or login == "nan":
                continue
            last = str(row.get("Last Name", "")).strip()
            first = str(row.get("First Name", "")).strip()
            job_type = str(row.get("Job Type", "")).strip() if "Job Type" in df.columns else ""
            if job_type == "nan":
                job_type = ""
            roster[login] = {
                "full_name": f"{last}, {first}" if last and first else login,
                "job_type": job_type,
            }
        return roster

    def _on_success(self, output_path):
        """Called on GUI thread after successful processing."""
        self.progress.stop()
        self.processing = False
        filename = os.path.basename(output_path)
        self.status_var.set(f"\u2713 Report generated: {filename}")
        self.status_label.config(fg=SUCCESS_FG)
        self._reset_button()

        result = messagebox.askyesno(
            "Report Generated",
            f"Report saved to:\n{output_path}\n\nOpen the output folder?"
        )
        if result:
            output_dir = os.path.dirname(output_path)
            if sys.platform == "win32":
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", output_dir])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", output_dir])

    def _on_error(self, error_msg):
        """Called on GUI thread after processing error."""
        self.progress.stop()
        self.processing = False
        self.status_var.set(f"\u2717 Error occurred")
        self.status_label.config(fg=ERROR_FG)
        self._reset_button()
        messagebox.showerror("Processing Error", f"An error occurred:\n\n{error_msg}")

    def _reset_button(self):
        """Re-enable the Generate button and restore hover effects."""
        self.run_btn.config(state="normal", bg=BTN_BG)
        self.run_btn.bind("<Enter>", lambda e: self.run_btn.config(bg=BTN_HOVER))
        self.run_btn.bind("<Leave>", lambda e: self.run_btn.config(bg=BTN_BG))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main():
    root = tk.Tk()
    app = CaseloadReportApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
