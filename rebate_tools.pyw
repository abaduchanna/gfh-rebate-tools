#!/usr/bin/env python3
"""
GFH Rebate Folder Tools — Professional Edition
==============================================
Complete rebate processing suite with:
  • Store management (add/edit/delete stores dynamically)
  • Rebate operations (rename, add/remove year suffix, delete rows)
  • XLS to XLSX conversion
  • Professional UI with tabs, logging, and GFH branding

Developed by Abad Umair Channa | Copyright © 2026 | All rights reserved.run(
        [sys.executable, "-m", "pip", "install", pkg_name, "--quiet", "--disable-pip-version-check"],
        capture_output=True,
    )

for _pkg, _pip in [
    ("openpyxl", "openpyxl"),
    ("xlrd", "xlrd==1.2.0"),
    ("xlwt", "xlwt"),
    ("xlutils", "xlutils"),
    ("pillow", "pillow"),
]:
    try:
        __import__(_pkg)
    except ImportError:
        _pip_install(_pip)

try:
    from openpyxl import load_workbook as _load_xlsx
except ImportError:
    _load_xlsx = None

try:
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
except ImportError:
    xlrd = xlwt = xl_copy = None

# win32com removed - not needed for basic rebate tools

# ── LOGO HANDLER ──
def load_header_logo():
    """Load GFH logo for header"""
    logo_files = ["GFH_Telecom_Logo.png", "logo.png", "gfh_logo.png"]
    for logo_file in logo_files:
        if Path(logo_file).exists():
            try:
                from PIL import Image
                img = Image.open(logo_file)
                img.thumbnail((100, 50), Image.Resampling.LANCZOS)
                return img
            except Exception:
                pass
    return None

# ═══════════════════════════════════════════════════════════════════════════
# GFH BRANDING & CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════
BRAND_NAVY = "#090d26"
BRAND_RED = "#f0541c"
BRAND_WHITE = "#ffffff"
STORE_CONFIG_FILE = Path(os.path.expanduser("~")) / ".gfh_rebate_stores.json"

# Default store renames (editable via UI)
DEFAULT_STORES = [
    ("Eliff", "E Iliff"),
    ("FM 1960", "FM1960"),
    ("Mckellips", "McKellips"),
    ("N 19th", "N 19"),
    ("N 35th", "N 35"),
    ("N 51st", "N 51"),
    ("New Thomas", "W Thomas"),
    ("Mt View RD", "Mount View"),
    ("4363 W Fuqua St", "Fuqua"),
]

# ═══════════════════════════════════════════════════════════════════════════
# STORE CONFIGURATION MANAGER
# ═══════════════════════════════════════════════════════════════════════════
class StoreConfigManager:
    """Manage store rename rules - load/save from JSON"""
    
    @staticmethod
    def load():
        """Load store config from file, fallback to defaults"""
        if STORE_CONFIG_FILE.exists():
            try:
                with open(STORE_CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    return data.get('stores', DEFAULT_STORES)
            except Exception:
                pass
        return DEFAULT_STORES
    
    @staticmethod
    def save(stores):
        """Save store config to JSON file"""
        try:
            with open(STORE_CONFIG_FILE, 'w') as f:
                json.dump({'stores': stores}, f, indent=2)
            return True
        except Exception:
            return False

# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — STORE RENAME (using loaded config)
# ═══════════════════════════════════════════════════════════════════════════
def step1_store_rename(folder: Path, stores_config, log):
    """Rename files based on store config"""
    log("\n" + "─" * 60)
    log("STEP 1: Store Rename")
    log("─" * 60)
    renamed = skipped = 0
    
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        new_name = f.name
        for old, new in stores_config:
            new_name = new_name.replace(old, new)
        
        if new_name != f.name:
            dest = f.parent / new_name
            try:
                f.rename(dest)
                log(f"✓ Renamed: {f.name} → {new_name}")
                renamed += 1
            except Exception as e:
                log(f"✗ ERROR renaming {f.name}: {e}")
        else:
            skipped += 1
    
    log(f"\nDone — {renamed} renamed, {skipped} unchanged.")
    return renamed

# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — ADD " 2026" SUFFIX
# ═══════════════════════════════════════════════════════════════════════════
def step2_add_2026(folder: Path, log):
    """Append ' 2026' to Excel filenames"""
    log("\n" + "─" * 60)
    log("STEP 2: Add ' 2026' Suffix")
    log("─" * 60)
    renamed = skipped = 0
    
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".xlsx", ".xls"):
            continue
        
        stem = f.stem
        if " 2026" in stem or " 2025" in stem or stem.endswith("2026"):
            log(f"⊘ Skipped (already has year): {f.name}")
            skipped += 1
            continue
        
        new_stem = f"{stem} 2026"
        new_name = f"{new_stem}{f.suffix}"
        dest = f.parent / new_name
        
        try:
            f.rename(dest)
            log(f"✓ Renamed: {f.name} → {new_name}")
            renamed += 1
        except Exception as e:
            log(f"✗ ERROR renaming {f.name}: {e}")
    
    log(f"\nDone — {renamed} renamed, {skipped} skipped.")
    return renamed

# ═══════════════════════════════════════════════════════════════════════════
# STEP 3 — DELETE 2025 ROWS
# ═══════════════════════════════════════════════════════════════════════════
def step3_delete_2025_rows(folder: Path, log):
    """Remove rows where column A contains '2025'"""
    log("\n" + "─" * 60)
    log("STEP 3: Delete 2025 Rows")
    log("─" * 60)
    
    modified = skipped = 0
    
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() != ".xls":
            continue
        
        try:
            workbook = xlrd.open_workbook(str(f))
            sheet = workbook.sheet_by_index(0)
            
            rows_to_delete = []
            for row_idx in range(sheet.nrows):
                cell_a = sheet.cell_value(row_idx, 0)
                if "2025" in str(cell_a):
                    rows_to_delete.append(row_idx)
            
            if rows_to_delete:
                new_workbook = xl_copy(workbook)
                new_sheet = new_workbook.get_sheet(0)
                
                for row_idx in sorted(rows_to_delete, reverse=True):
                    new_sheet.delete_rows(row_idx + 1)
                
                new_workbook.save(str(f))
                log(f"✓ Modified: {f.name} ({len(rows_to_delete)} rows deleted)")
                modified += 1
            else:
                log(f"⊘ Skipped: {f.name} (no 2025 rows)")
                skipped += 1
        
        except Exception as e:
            log(f"✗ ERROR processing {f.name}: {e}")
    
    log(f"\nDone — {modified} modified, {skipped} unchanged.")
    return modified

# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — XLS TO XLSX CONVERSION
# ═══════════════════════════════════════════════════════════════════════════
XLSX_FORMAT = 51

def step4_xls_to_xlsx(folder: Path, recurse=False, delete_originals=False, log=None):
    """Convert legacy XLS files to XLSX using Excel COM"""
    if log is None:
        log = print
    
    log("\n" + "─" * 60)
    log("STEP 4: Convert XLS to XLSX")
    log("─" * 60)
    
    if not win32com:
        log("✗ ERROR: pywin32 not available (requires Excel COM)")
        return 0
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception as e:
        log(f"✗ ERROR: Cannot launch Excel: {e}")
        return 0
    
    # Find legacy files
    legacy_exts = (".xls", ".xlsm", ".xlt", ".xlsb", ".xlc")
    files = []
    
    if recurse:
        for root, _, filenames in os.walk(str(folder)):
            for name in filenames:
                if name.startswith("~$"):
                    continue
                if os.path.splitext(name)[1].lower() in legacy_exts:
                    files.append(os.path.join(root, name))
    else:
        for name in os.listdir(str(folder)):
            if name.startswith("~$"):
                continue
            path = os.path.join(str(folder), name)
            if os.path.isfile(path) and os.path.splitext(name)[1].lower() in legacy_exts:
                files.append(path)
    
    converted = skipped = 0
    
    for i, filepath in enumerate(sorted(files)):
        try:
            filename = os.path.basename(filepath)
            xlsx_path = os.path.splitext(filepath)[0] + ".xlsx"
            
            if os.path.exists(xlsx_path) and not delete_originals:
                log(f"⊘ Skipped: {filename} (XLSX already exists)")
                skipped += 1
                continue
            
            workbook = excel.Workbooks.Open(os.path.abspath(filepath), ReadOnly=False)
            workbook.SaveAs(os.path.abspath(xlsx_path), XLSX_FORMAT)
            workbook.Close()
            
            if delete_originals:
                os.remove(filepath)
                log(f"✓ Converted: {filename} → {os.path.basename(xlsx_path)} (original deleted)")
            else:
                log(f"✓ Converted: {filename} → {os.path.basename(xlsx_path)}")
            
            converted += 1
            
            # Restart Excel periodically
            if i % 60 == 59:
                excel.Quit()
                time.sleep(1)
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
        
        except Exception as e:
            log(f"✗ ERROR converting {filename}: {e}")
    
    try:
        excel.Quit()
    except:
        pass
    
    log(f"\nDone — {converted} converted, {skipped} skipped.")
    return converted

# ═══════════════════════════════════════════════════════════════════════════
# MAIN GUI APPLICATION
# ═══════════════════════════════════════════════════════════════════════════
class RebateToolsApp:
    """Unified Rebate Tools GUI"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("GFH Rebate Tools Suite")
        self.root.geometry("900x700")
        self.root.configure(bg=BRAND_NAVY)
        
        try:
            if os.path.exists("app_icon.ico"):
                self.root.iconbitmap("app_icon.ico")
        except:
            pass
        
        self.stores_config = StoreConfigManager.load()
        self.is_dark = True
        self.folder_path = None
        self.log_queue = []
        
        self.build_ui()
    

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        self.is_dark = not self.is_dark
        if self.is_dark:
            self.root.configure(bg=BRAND_NAVY)
        else:
            self.root.configure(bg=BRAND_LIGHT_BG)
        self.log(f"✓ Theme: {'Dark' if self.is_dark else 'Light'}")

    def build_ui(self):
        """Build professional 4-tab UI"""
        # ── HEADER ──
        header = tk.Frame(self.root, bg=BRAND_NAVY, height=100)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        # Logo + Title
        logo_title = tk.Frame(header, bg=BRAND_NAVY)
        logo_title.pack(fill=tk.X, padx=20, pady=10)
        
        # Try to load logo image
        try:
            from PIL import Image, ImageTk
            logo_files = ["GFH_Telecom_Logo.png", "logo.png"]
            for logo_file in logo_files:
                if os.path.exists(logo_file):
                    img = Image.open(logo_file)
                    img.thumbnail((80, 50), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    logo_label = tk.Label(logo_title, image=photo, bg=BRAND_NAVY)
                    logo_label.image = photo
                    logo_label.pack(side=tk.LEFT, padx=(0, 15))
                    break
        except Exception:
            pass
        
        # Title text
        tk.Label(
            logo_title,
            text="GFH Telecom LLC",
            font=("Segoe UI", 20, "bold"),
            fg=BRAND_RED,
            bg=BRAND_NAVY
        ).pack(side=tk.LEFT)
        
        tk.Label(
            logo_title,
            text="Rebate Folder Tools",
            font=("Segoe UI", 14, "bold"),
            fg="#ffffff",
            bg=BRAND_NAVY
        ).pack(side=tk.LEFT, padx=20)
        
        # ── NOTEBOOK (TABS) ──
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Operations
        self.build_tab_operations()
        
        # Tab 2: Store Management
        self.build_tab_store_management()
        
        # Tab 3: Settings
        self.build_tab_settings()
        
        # Tab 4: Logs
        self.build_tab_logs()
    
    def build_tab_operations(self):
        """Rebate operations tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Rebate Operations  ")
        
        # Folder selection
        folder_frame = ttk.LabelFrame(frame, text="Select Rebate Folder", padding=15)
        folder_frame.pack(fill=tk.X, padx=15, pady=10)
        
        btn_frame = ttk.Frame(folder_frame)
        btn_frame.pack(fill=tk.X)
        
        ttk.Button(btn_frame, text="📁 Browse", command=self.select_folder).pack(side=tk.LEFT, padx=5)
        self.folder_label = ttk.Label(btn_frame, text="No folder selected", foreground="#7a8a99")
        self.folder_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # Operations checkboxes
        ops_frame = ttk.LabelFrame(frame, text="Select Operations", padding=15)
        ops_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.op_rename = tk.BooleanVar(value=True)
        self.op_add_year = tk.BooleanVar(value=True)
        self.op_remove_2025 = tk.BooleanVar(value=True)
        self.op_convert = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(ops_frame, text="1. Store Rename", variable=self.op_rename).pack(anchor=tk.W, pady=5)
        ttk.Checkbutton(ops_frame, text="2. Add '2026' Suffix", variable=self.op_add_year).pack(anchor=tk.W, pady=5)
        ttk.Checkbutton(ops_frame, text="3. Delete 2025 Rows", variable=self.op_remove_2025).pack(anchor=tk.W, pady=5)
        ttk.Checkbutton(ops_frame, text="4. Convert XLS → XLSX", variable=self.op_convert).pack(anchor=tk.W, pady=5)
        
        # Options frame
        opts_frame = ttk.LabelFrame(frame, text="Step 4 Options (XLS Conversion)", padding=15)
        opts_frame.pack(fill=tk.X, padx=15, pady=10)
        
        self.opt_subfolders = tk.BooleanVar(value=True)
        self.opt_overwrite = tk.BooleanVar(value=False)
        self.opt_delete_orig = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(opts_frame, text="Include subfolders", variable=self.opt_subfolders).pack(anchor=tk.W, pady=3)
        ttk.Checkbutton(opts_frame, text="Overwrite existing .xlsx", variable=self.opt_overwrite).pack(anchor=tk.W, pady=3)
        ttk.Checkbutton(opts_frame, text="Delete original .xls", variable=self.opt_delete_orig).pack(anchor=tk.W, pady=3)
        
        # Action buttons
        action_frame = ttk.Frame(frame)
        action_frame.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Button(action_frame, text="▶ RUN ALL STEPS", command=self.run_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Step 1 Only", command=lambda: self.run_step(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Step 2 Only", command=lambda: self.run_step(2)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Step 3 Only", command=lambda: self.run_step(3)).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="Step 4 Only", command=lambda: self.run_step(4)).pack(side=tk.LEFT, padx=5)
    
    def build_tab_store_management(self):
        """Store management tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Store Management  ")
        
        info_label = ttk.Label(
            frame,
            text="Add, edit, or delete store rename rules. Changes are saved automatically.",
            foreground="#7a8a99"
        )
        info_label.pack(padx=15, pady=10)
        
        # Treeview
        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        cols = ("Old Name", "New Name")
        self.stores_tree = ttk.Treeview(tree_frame, columns=cols, height=15, show="headings")
        self.stores_tree.column("Old Name", width=300)
        self.stores_tree.column("New Name", width=300)
        self.stores_tree.heading("Old Name", text="Old Name (find in filenames)")
        self.stores_tree.heading("New Name", text="New Name (replace with)")
        self.stores_tree.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.stores_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.stores_tree.configure(yscrollcommand=scrollbar.set)
        
        self.refresh_stores_tree()
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=15, pady=15)
        
        ttk.Button(btn_frame, text="➕ Add Store", command=self.add_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✏️ Edit Selected", command=self.edit_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✕ Delete Selected", command=self.delete_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Save Changes", command=self.save_stores).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="↻ Reset to Defaults", command=self.reset_stores).pack(side=tk.LEFT, padx=5)
    
    def build_tab_settings(self):
        """Settings tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Settings  ")
        
        info = ttk.Label(
            frame,
            text="Configure application settings and behavior",
            foreground="#7a8a99"
        )
        info.pack(padx=15, pady=10)
        
        # Settings
        settings_frame = ttk.LabelFrame(frame, text="General Settings", padding=15)
        settings_frame.pack(fill=tk.X, padx=15, pady=10)
        
        # Store config location
        config_frame = ttk.Frame(settings_frame)
        config_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(config_frame, text="Store config location:").pack(side=tk.LEFT)
        ttk.Label(config_frame, text=str(STORE_CONFIG_FILE), foreground="#5a8acc").pack(side=tk.LEFT, padx=10)
        
        # Version
        version_frame = ttk.Frame(settings_frame)
        version_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(version_frame, text="App version:").pack(side=tk.LEFT)
        ttk.Label(version_frame, text="1.0.0", foreground="#5a8acc").pack(side=tk.LEFT, padx=10)
        
        # Help text
        help_frame = ttk.LabelFrame(frame, text="Help", padding=15)
        help_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        help_text = tk.Text(help_frame, height=12, width=70, wrap=tk.WORD, bg="#f6f7fb")
        help_text.pack(fill=tk.BOTH, expand=True)
        
        help_text.insert(tk.END, """REBATE FOLDER TOOLS - User Guide

Step 1: Store Rename
Renames files based on configured store mapping rules.
Example: "Eliff" → "E Iliff"

Step 2: Add '2026' Suffix
Adds " 2026" suffix to Excel files that don't have a year.
Example: "Store Report" → "Store Report 2026"

Step 3: Delete 2025 Rows
Removes rows containing "2025" from XLS files.
Keeps all 2026 data intact.

Step 4: Convert XLS → XLSX
Converts legacy Excel files (.xls) to modern format (.xlsx).
Options:
  • Include subfolders: Process folders recursively
  • Overwrite existing: Replace existing XLSX files
  • Delete original: Remove .xls after conversion

Store Management:
Add or edit store rename rules. All changes are saved to:
~/.gfh_rebate_stores.json""")
        help_text.config(state=tk.DISABLED)
    
    def build_tab_logs(self):
        """Logs tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="  Logs  ")
        
        self.log_text = scrolledtext.ScrolledText(
            frame,
            height=20,
            width=100,
            font=("Consolas", 9),
            bg="#0f1830",
            fg="#e2e8f0"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Button(btn_frame, text="Clear", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Copy", command=self.copy_log).pack(side=tk.LEFT, padx=5)
    def build_tab_stores(self):
        """Store Management Tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Store Management")
        
        ttk.Label(frame, text="Store Rename Rules", font=("Segoe UI", 11, "bold")).pack(padx=10, pady=10)
        
        # Treeview
        cols = ("Old Name", "New Name")
        self.stores_tree = ttk.Treeview(frame, columns=cols, height=12, show="headings")
        self.stores_tree.column("Old Name", width=200)
        self.stores_tree.column("New Name", width=200)
        self.stores_tree.heading("Old Name", text="Old Name")
        self.stores_tree.heading("New Name", text="New Name")
        self.stores_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.refresh_stores_tree()
        
        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Add Store", command=self.add_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit Selected", command=self.edit_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Delete Selected", command=self.delete_store).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save Changes", command=self.save_stores).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_stores).pack(side=tk.LEFT, padx=5)
    
    def build_tab_rebate(self):
        """Rebate Tools Tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Rebate Folder Operations")
        
        ttk.Label(frame, text="Select Folder to Process", font=("Segoe UI", 11, "bold")).pack(padx=10, pady=10)
        
        folder_btn_frame = ttk.Frame(frame)
        folder_btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(folder_btn_frame, text="Browse Folder", command=self.select_folder).pack(side=tk.LEFT, padx=5)
        self.folder_label = ttk.Label(folder_btn_frame, text="No folder selected", font=("Segoe UI", 9))
        self.folder_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        ttk.Label(frame, text="Operations:", font=("Segoe UI", 11, "bold")).pack(padx=10, pady=10)
        
        ops_frame = ttk.Frame(frame)
        ops_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(ops_frame, text="Run All Steps", command=self.run_all_rebate).pack(side=tk.LEFT, padx=5)
        ttk.Button(ops_frame, text="Step 1: Rename Stores", command=lambda: self.run_step(1)).pack(side=tk.LEFT, padx=5)
        ttk.Button(ops_frame, text="Step 2: Add 2026", command=lambda: self.run_step(2)).pack(side=tk.LEFT, padx=5)
        ttk.Button(ops_frame, text="Step 3: Delete 2025", command=lambda: self.run_step(3)).pack(side=tk.LEFT, padx=5)
    
    def build_tab_convert(self):
        """XLS to XLSX Conversion Tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="XLS to XLSX Conversion")
        
        ttk.Label(frame, text="Convert Legacy Excel Files", font=("Segoe UI", 11, "bold")).pack(padx=10, pady=10)
        
        folder_btn_frame = ttk.Frame(frame)
        folder_btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(folder_btn_frame, text="Browse Folder", command=self.select_convert_folder).pack(side=tk.LEFT, padx=5)
        self.convert_folder_label = ttk.Label(folder_btn_frame, text="No folder selected", font=("Segoe UI", 9))
        self.convert_folder_label.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        options_frame = ttk.Frame(frame)
        options_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.recurse_var = tk.BooleanVar()
        self.delete_var = tk.BooleanVar()
        
        ttk.Checkbutton(options_frame, text="Include Subfolders", variable=self.recurse_var).pack(anchor=tk.W)
        ttk.Checkbutton(options_frame, text="Delete Original Files", variable=self.delete_var).pack(anchor=tk.W)
        
        ttk.Button(frame, text="Start Conversion", command=self.run_conversion).pack(padx=10, pady=10)
    
    def build_tab_log(self):
        """Log Output Tab"""
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Log Output")
        
        self.log_text = scrolledtext.ScrolledText(frame, height=20, width=80, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(btn_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Copy Log", command=self.copy_log).pack(side=tk.LEFT, padx=5)
    
    # ── Store Management ──
    def refresh_stores_tree(self):
        """Refresh the stores treeview"""
        for item in self.stores_tree.get_children():
            self.stores_tree.delete(item)
        
        for old, new in self.stores_config:
            self.stores_tree.insert("", tk.END, values=(old, new))
    
    def add_store(self):
        """Add new store rename rule"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Store")
        dialog.geometry("400x150")
        
        ttk.Label(dialog, text="Old Name:").pack(padx=10, pady=5)
        old_entry = ttk.Entry(dialog, width=40)
        old_entry.pack(padx=10, pady=5)
        
        ttk.Label(dialog, text="New Name:").pack(padx=10, pady=5)
        new_entry = ttk.Entry(dialog, width=40)
        new_entry.pack(padx=10, pady=5)
        
        def save():
            old = old_entry.get().strip()
            new = new_entry.get().strip()
            if old and new:
                self.stores_config.append((old, new))
                self.refresh_stores_tree()
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Both fields required")
        
        ttk.Button(dialog, text="Save", command=save).pack(pady=10)
    
    def edit_store(self):
        """Edit selected store rule"""
        selection = self.stores_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Select a store to edit")
            return
        
        item = selection[0]
        old, new = self.stores_tree.item(item, 'values')
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Store")
        dialog.geometry("400x150")
        
        ttk.Label(dialog, text="Old Name:").pack(padx=10, pady=5)
        old_entry = ttk.Entry(dialog, width=40)
        old_entry.insert(0, old)
        old_entry.pack(padx=10, pady=5)
        
        ttk.Label(dialog, text="New Name:").pack(padx=10, pady=5)
        new_entry = ttk.Entry(dialog, width=40)
        new_entry.insert(0, new)
        new_entry.pack(padx=10, pady=5)
        
        def save():
            new_old = old_entry.get().strip()
            new_new = new_entry.get().strip()
            if new_old and new_new:
                idx = list(self.stores_tree.get_children()).index(item)
                self.stores_config[idx] = (new_old, new_new)
                self.refresh_stores_tree()
                dialog.destroy()
        
        ttk.Button(dialog, text="Save", command=save).pack(pady=10)
    
    def delete_store(self):
        """Delete selected store rule"""
        selection = self.stores_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Select a store to delete")
            return
        
        item = selection[0]
        idx = list(self.stores_tree.get_children()).index(item)
        self.stores_config.pop(idx)
        self.refresh_stores_tree()
    
    def save_stores(self):
        """Save stores to file"""
        if StoreConfigManager.save(self.stores_config):
            messagebox.showinfo("Success", "Store configuration saved")
            self.log(f"✓ Saved {len(self.stores_config)} store rules")
        else:
            messagebox.showerror("Error", "Failed to save configuration")
    
    def reset_stores(self):
        """Reset to default stores"""
        if messagebox.askyesno("Confirm", "Reset to default stores?"):
            self.stores_config = DEFAULT_STORES.copy()
            self.refresh_stores_tree()
            self.log("Stores reset to defaults")
    
    # ── Rebate Operations ──
    def select_folder(self):
        """Select folder for rebate operations"""
        folder = filedialog.askdirectory(title="Select Rebate Folder")
        if folder:
            self.folder_path = Path(folder)
            self.folder_label.config(text=str(self.folder_path))
    
    def run_step(self, step):
        """Run specific rebate step"""
        if not self.folder_path:
            messagebox.showwarning("Warning", "Select a folder first")
            return
        
        threading.Thread(target=self._run_step_thread, args=(step,), daemon=True).start()
    
    def _run_step_thread(self, step):
        """Run step in thread"""
        try:
            if step == 1:
                step1_store_rename(self.folder_path, self.stores_config, self.log)
            elif step == 2:
                step2_add_2026(self.folder_path, self.log)
            elif step == 3:
                step3_delete_2025_rows(self.folder_path, self.log)
        except Exception as e:
            self.log(f"✗ ERROR: {e}")
    
    def run_all_rebate(self):
        """Run all rebate steps"""
        if not self.folder_path:
            messagebox.showwarning("Warning", "Select a folder first")
            return
        
        threading.Thread(target=self._run_all_thread, daemon=True).start()
    
    def _run_all_thread(self):
        """Run all steps in thread"""
        try:
            self.log("=" * 60)
            self.log("STARTING REBATE PROCESSING")
            self.log("=" * 60)
            step1_store_rename(self.folder_path, self.stores_config, self.log)
            step2_add_2026(self.folder_path, self.log)
            step3_delete_2025_rows(self.folder_path, self.log)
            self.log("\n" + "=" * 60)
            self.log("ALL STEPS COMPLETED")
            self.log("=" * 60)
        except Exception as e:
            self.log(f"✗ ERROR: {e}")
    
    # ── XLS to XLSX Conversion ──
    def select_convert_folder(self):
        """Select folder for conversion"""
        folder = filedialog.askdirectory(title="Select Folder to Convert")
        if folder:
            self.convert_folder_path = Path(folder)
            self.convert_folder_label.config(text=str(self.convert_folder_path))
    
    def run_conversion(self):
        """Run XLS to XLSX conversion"""
        if not hasattr(self, 'convert_folder_path'):
            messagebox.showwarning("Warning", "Select a folder first")
            return
        
        threading.Thread(
            target=self._run_conversion_thread,
            daemon=True
        ).start()
    
    def _run_conversion_thread(self):
        """Run conversion in thread"""
        try:
            step4_xls_to_xlsx(
                self.convert_folder_path,
                recurse=self.recurse_var.get(),
                delete_originals=self.delete_var.get(),
                log=self.log
            )
        except Exception as e:
            self.log(f"✗ ERROR: {e}")
    
    # ── Logging ──
    def log(self, msg):
        """Add message to log"""
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        """Clear log output"""
        self.log_text.delete(1.0, tk.END)
    
    def copy_log(self):
        """Copy log to clipboard"""
        content = self.log_text.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        messagebox.showinfo("Success", "Log copied to clipboard")

# ═══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    root = tk.Tk()
    app = RebateToolsApp(root)
    root.mainloop()
