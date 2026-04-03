# GUI Design Reference

Reusable design spec for building tkinter data-processing apps.
Copy this structure and adapt the fields/processing logic for new applications.

> *Created by: James Calby — Data Systems*

---

## Layout Structure

```
┌─────────────────────────────────────────────────┐
│  HEADER (accent color banner)                   │
│    App Title — large bold white text             │
│    Subtitle — lighter, smaller text             │
├─────────────────────────────────────────────────┤
│  BODY (light background, padded)                │
│                                                 │
│  Input File Label (bold)                        │
│  [ text entry field          ] [Browse...]      │
│                                                 │
│  Output Folder Label (bold)                     │
│  [ text entry field          ] [Browse...]      │
│                                                 │
│  ─── Optional Section (separator) ───           │
│  Optional File Label                            │
│  [ text entry field          ] [Browse...]      │
│  Status indicator (✓ loaded / ⚠ not found)     │
│                                                 │
│  [========= progress bar =========]             │
│  Status text ("Ready" / "Processing...")        │
│                                                 │
│                        [ Action Button ]        │
├─────────────────────────────────────────────────┤
│  FOOTER (subtle background)                     │
│    Organization · Team · Version                │
│    ⭐ Credit line ⭐                             │
└─────────────────────────────────────────────────┘
```

---

## Color Palette

| Element | Color | Hex | Usage |
|---------|-------|-----|-------|
| Header / accent / primary button | Dark blue | `#1F4E79` | App identity, action buttons |
| Button hover / active | Medium blue | `#2E75B6` | Interactive feedback |
| Body background | Light gray-blue | `#F0F4F8` | Main content area |
| Footer background | Slightly darker gray | `#E8EDF2` | Subtle separation |
| Success text / indicators | Green | `#2E7D32` | Checkmarks, completion |
| Error / warning text | Red | `#C62828` | Errors, failures |
| Muted text (status, footer) | Gray | `#555` / `#666` | Secondary information |
| Credit line text | Light gray | `#888` | Subtle attribution |

---

## Font Specifications

All fonts use **Segoe UI** (Windows system font — clean, modern, professional).

| Element | Family | Size | Weight | Code |
|---------|--------|------|--------|------|
| App title | Segoe UI | 16 | Bold | `("Segoe UI", 16, "bold")` |
| Subtitle | Segoe UI | 10 | Regular | `("Segoe UI", 10)` |
| Section labels | Segoe UI | 10 | Bold | `("Segoe UI", 10, "bold")` |
| Input fields / body text | Segoe UI | 9 | Regular | `("Segoe UI", 9)` |
| Primary button | Segoe UI | 11 | Bold | `("Segoe UI", 11, "bold")` |
| Small buttons (Browse) | Segoe UI | 9 | Regular | `("Segoe UI", 9)` |
| Status text | Segoe UI | 9 | Regular | `("Segoe UI", 9)` |
| Footer | Segoe UI | 8 | Italic | `("Segoe UI", 8, "italic")` |

---

## Component Patterns

### Window Setup

```python
root = tk.Tk()
root.title("App Title")
root.resizable(False, False)
root.configure(bg="#F0F4F8")

# Center on screen
root.update_idletasks()
w, h = root.winfo_width(), root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (w // 2)
y = (root.winfo_screenheight() // 2) - (h // 2)
root.geometry(f"+{x}+{y}")
```

### Header Banner

```python
header = tk.Frame(root, bg="#1F4E79", padx=20, pady=14)
header.pack(fill="x")

tk.Label(header, text="App Title",
         font=("Segoe UI", 16, "bold"), fg="white", bg="#1F4E79").pack()
tk.Label(header, text="Subtitle — Description",
         font=("Segoe UI", 10), fg="#B0C4DE", bg="#1F4E79").pack()
```

### File Browse Row

```python
# Label
tk.Label(body, text="Input File:", font=("Segoe UI", 10, "bold"),
         bg="#F0F4F8", anchor="w").grid(row=R, column=0, columnspan=2,
         sticky="w", pady=(12, 2))

# Entry + Button
frame = tk.Frame(body, bg="#F0F4F8")
frame.grid(row=R+1, column=0, columnspan=2, sticky="ew")

file_var = tk.StringVar()
tk.Entry(frame, textvariable=file_var, width=52,
         font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
tk.Button(frame, text="Browse...", command=browse_fn,
          font=("Segoe UI", 9)).pack(side="left")
```

### Folder Browse Row

```python
# Same pattern, but use filedialog.askdirectory() instead of askopenfilename()
def browse_folder():
    path = filedialog.askdirectory(title="Select Folder", initialdir=default_dir)
    if path:
        folder_var.set(path)
```

### Status Indicator

```python
indicator = tk.Label(body, text="", font=("Segoe UI", 9), bg="#F0F4F8")
indicator.grid(...)

# Success state
indicator.config(text="\u2713 File loaded successfully", fg="#2E7D32")

# Warning state
indicator.config(text="\u26a0 Optional file not found", fg="#555")

# Error state
indicator.config(text="\u2717 Required file missing", fg="#C62828")
```

### Progress Bar + Status Text

```python
progress = ttk.Progressbar(body, mode="indeterminate", length=420)
progress.grid(row=R, column=0, columnspan=2, sticky="ew", pady=(16, 4))

status_var = tk.StringVar(value="Ready")
tk.Label(body, textvariable=status_var, font=("Segoe UI", 9),
         bg="#F0F4F8", fg="#555").grid(row=R+1, column=0, columnspan=2,
         sticky="w")

# To start:  progress.start(15)
# To stop:   progress.stop()
```

### Primary Action Button (right-aligned, with hover)

```python
btn = tk.Button(frame, text="Generate Report",
                command=run_fn, font=("Segoe UI", 11, "bold"),
                bg="#1F4E79", fg="white", padx=20, pady=6,
                activebackground="#2E75B6", activeforeground="white",
                cursor="hand2")
btn.pack(side="right")

# Hover effect
btn.bind("<Enter>", lambda e: btn.config(bg="#2E75B6"))
btn.bind("<Leave>", lambda e: btn.config(bg="#1F4E79"))
```

### Footer

```python
footer = tk.Frame(root, bg="#E8EDF2", padx=20, pady=8)
footer.pack(fill="x", side="bottom")

tk.Label(footer, text="Organization \u00b7 Team \u00b7 v1.0",
         font=("Segoe UI", 8, "italic"), fg="#666", bg="#E8EDF2").pack()

tk.Label(footer,
         text="\u2b50 Crafted by the legendary James Calby "
              "\u2014 Data Systems Analyst Extraordinaire \u2b50",
         font=("Segoe UI", 8, "italic"), fg="#888", bg="#E8EDF2").pack()
```

### Separator (between sections)

```python
sep = ttk.Separator(body, orient="horizontal")
sep.grid(row=R, column=0, columnspan=2, sticky="ew", pady=(16, 8))
```

---

## Threading Pattern (keeps UI responsive)

Never run long-running tasks on the GUI thread. Use this pattern:

```python
def _on_generate(self):
    """User clicked the action button."""
    # 1. Validate inputs
    if not self._validate():
        return

    # 2. Disable button, start progress
    self.run_btn.config(state="disabled")
    self.progress.start(15)
    self.status_var.set("Processing...")

    # 3. Launch background thread
    thread = threading.Thread(target=self._process, args=(...,), daemon=True)
    thread.start()

def _process(self, ...):
    """Runs in background thread — NO tkinter calls here."""
    try:
        # ... do the actual work ...
        result = do_heavy_processing()

        # Signal success back to GUI thread
        self.root.after(0, self._on_success, result)
    except Exception as e:
        # Signal error back to GUI thread
        self.root.after(0, self._on_error, str(e))

def _on_success(self, result):
    """Called on GUI thread after success."""
    self.progress.stop()
    self.status_var.set("\u2713 Done!")
    self.run_btn.config(state="normal")

    if messagebox.askyesno("Done", "Open the result?"):
        os.startfile(result)  # Windows

def _on_error(self, msg):
    """Called on GUI thread after error."""
    self.progress.stop()
    self.status_var.set("\u2717 Error occurred")
    self.run_btn.config(state="normal")
    messagebox.showerror("Error", msg)
```

**Key rules:**
- `self.root.after(0, callback)` — marshals back to GUI thread
- `daemon=True` — thread dies if the app is closed
- Never call `tk.Label()`, `messagebox`, etc. from the background thread

---

## PyInstaller Packaging

### Build Script

```python
import subprocess, sys

subprocess.run([
    sys.executable, "-m", "PyInstaller",
    "--onefile",          # Single .exe
    "--windowed",         # No console window
    "--name", "AppName",
    "app.py",
], check=True)
```

### Path Resolution (PyInstaller compatible)

```python
def get_app_dir():
    """Works both as script and as frozen .exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))
```

Use `get_app_dir()` as the base for all relative paths (config files, input
folders, output folders). This ensures the app works identically whether run
as `python app.py` or as a compiled `.exe`.

### Distribution Folder

```
MyApp/
├── MyApp.exe
├── config/         # User-editable configuration files
├── input/          # User drops input files here
└── output/         # Generated output goes here
```

Config files must live **outside** the .exe so users can edit them.
The .exe reads them from its own directory via `get_app_dir()`.

---

## Adapting for a New App

1. Copy `app.py` as your starting point
2. Update the header title and subtitle
3. Replace the browse fields with your app's inputs
4. Replace `_process()` with your app's processing logic
5. Update the footer credit/version
6. Update `build_exe.py` with the new app name
7. Keep the same color palette, fonts, and threading pattern for consistency
