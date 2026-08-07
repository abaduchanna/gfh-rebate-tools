import datetime as _doc_dt
_DOC_YEAR = _doc_dt.date.today().year

f"""
Rebate Folder Tools
===================
Runs four operations in sequence (each can be toggled on/off):

  1. Store Rename          — fix store name typos in filenames
  2. Add " 2026" Suffix    — append " 2026" to Excel filenames that don't have it
  3. Delete 2025 Rows      — remove rows where column A contains "2025"
  4. Convert Legacy Excel  — convert every .xls / .xlsm / .xlt / .xlsb in the
                             folder to modern .xlsx using real Excel (COM).
                             Originals can be kept or deleted.

Ship this file together with GFH_Telecom_TBLogo.ico and GFH_Telecom_Logo.png
in the same folder for the window/taskbar icon and header logo.

Developed by Abad Umair Channa  |  Copyright © {_DOC_YEAR}
"""

import os
import sys
import time
import subprocess
import threading
import queue
import traceback
import tkinter as tk
from theme_manager import ThemeManager, apply_theme_to_window, get_copyright_year, create_theme_toggle_button
from tkinter import ttk, filedialog, scrolledtext, messagebox
from pathlib import Path
from datetime import datetime

# ── Auto-install required packages BEFORE importing them ──────────────────
def _pip_install(pkg_name):
    # Never pip-install from inside a frozen EXE: all deps are bundled by
    # PyInstaller, and re-launching sys.executable would spawn another
    # instance of the app itself (window flooding).
    if getattr(sys, "frozen", False):
        return
    subprocess.run(
        [sys.executable, "-m", "pip", "install", pkg_name,
         "--quiet", "--disable-pip-version-check"],
        capture_output=True,
    )

# Check each dependency by its real import module name (pywin32 -> win32com,
# pillow -> PIL), not the pip package name, so this only runs pip when a
# package is genuinely missing.
for _mod, _pip in [
    ("openpyxl",  "openpyxl"),
    ("xlrd",      "xlrd==1.2.0"),
    ("xlwt",      "xlwt"),
    ("xlutils",   "xlutils"),
    ("win32com",  "pywin32"),
    ("PIL",       "pillow"),
]:
    try:
        __import__(_mod)
    except ImportError:
        _pip_install(_pip)

# ── Now import after ensuring packages are installed ──────────────────────
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

try:
    import win32com.client as _win32com_client
except ImportError:
    _win32com_client = None

try:
    from PIL import Image as _PI, ImageTk as _PIT
except ImportError:
    _PI = _PIT = None


# ═══════════════════════════════════════════════════════════════════════════
# BRAND / WINDOW CONFIG — kept in sync with GFH_Inventory_Aging_Processor.pyw
# ═══════════════════════════════════════════════════════════════════════════
NAVY  = "#090d26"
RED   = "#e8212a"
WHITE = "#ffffff"
LIGHT = "#f0f4fa"
LOG_BG = "#10182e"
LOG_FG = "#a8d8ff"

ICON_ICO_NAME = "GFH_Telecom_TBLogo.ico"
LOGO_PNG_NAME = "GFH_Telecom_Logo.png"
COPYRIGHT_TEXT = f"Developed by Abad Umair Channa  |  Copyright © {get_copyright_year()}  |  All rights reserved."
ICON_ICO_B64 = "AAABAAEAICAAAAEAIACoEAAAFgAAACgAAAAgAAAAQAAAAAEAIAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAA0GRv/NBkb/zQZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zQZG/81GRv/NBgb/zUZG/81GRv/NBkb/zQZG/80GBv/NRkb/zQYG/80GRv/NBkb/zQZG/80GRv/NBkb/zQZG/81GRv/NRkb/zQZG/80GRv/NBgb/zUZG/80GRv/NBkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GBv/NRgb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zYZHP81GRv/NRkb/zUYG/81GRv/NRkb/zUZG/81GRv/NRkc/zUZG/81GRz/Nhkc/zYZG/82GRz/Nhkc/zYZHP82GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUZG/81GRv/NRkb/zUYG/8/JCf/TDM2/zYaHP81GRv/NRgb/zUZG/81GRv/NRkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZHP81GRv/NRkb/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkb/zYZG/82GRz/NRkb/zoeIP8/JCf/Nhkb/zUZG/81GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zUYG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRv/Nhkb/zYZG/82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZG/82GRr/NRkh/zMZKv8zGS3/Mxkn/zUZHv82GRr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Mxoy/yseZv8lIIn/IyCS/yIfkf8iHoz/JB16/ywbUP80GSX/Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRv/Nhke/y8eWf8mI53/JyGL/y0dW/8wG0H/MBo7/y4aRf8pHGT/IR6O/yMdg/8wGjr/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHf8uIGj/JyWn/y8eWf80Giv/MRxH/ywfaP8rH3D/LR1e/zIaOf8zGS3/Jxxr/yEekv8wGj//Nhka/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRr/MR9S/ykorP8xHk3/Mxw4/yojjf8nJKD/KSGD/ysgd/8oIIj/JSKb/yoecP8vG0f/KB1t/yMfjv8zGi7/Nhkb/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP83GRz/Nxkb/zUaKP8sKKL/MCJu/zQbM/8qJqD/LCOE/zQaMv83GRv/Nxkb/zYZHf8uHVn/IyOp/yQhnP8sHWX/Jh+G/yodav82GRv/Nxkc/zYZHP82GRz/Nhkc/zYZHP82GRz/Nhkc/zcZHP82GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP83GRn/NCBT/y4rsP81GzL/LyWD/y0mkv82GiP/NRsu/y8hcP8sI4T/MB5Z/zAdSv8mI6H/JiOc/ywfbP8vHEz/JiGW/zQaKv83GRv/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GRz/Nxkc/zYZHP83GRz/Nxkc/zYZHP83GRz/Nxkc/zcZGv8yJX3/MSeR/zUbNf8uKrL/NB5K/zYaJv8tJpb/LCaY/y4hdv8pJqT/LiFy/y4eX/8sH3D/NBou/zQaK/8mIpn/MRxC/zcZGv83GRz/Nxkc/zcZHP83GRz/Nxkc/zcZHP82GBv/Nhgb/zYYG/82GRz/Nhkc/zYZHP82GRz/Nxkb/zIplP8zJnr/NB5I/y8ss/81Giz/Mx9P/y0qrv81GzD/NxgX/zEeUf8pKLD/LSN//y0iev8tInn/LCF9/yclq/8wHU3/NxgZ/zcZHP83GRz/Nxkc/zcZHP83GRz/Nhkc/zYYG/82GBv/Nhgc/zYZHP82GBz/Nhgb/zYYG/82GR//Miyi/zIrnf80IVv/MS62/zUbLP8zH0//Lyyy/zUbMf82GBf/Mh9U/ysqtP8uI4D/LiJ7/y0iev8tIXn/Lh9r/zQaK/82GBv/Nhgb/zYYHP82GRz/Nhkc/zYYHP82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NhgZ/zUfSf8yM8X/MjHC/zIrmv8xMLv/NCBO/zUZJv8xKp//MCul/zElgP8uK7H/MSNz/zYYGf82GBf/NhgY/zUaK/8zHUL/Nhgc/zYYG/82GBv/Nhgc/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NSBL/zM0yf8yM8X/MyiB/zMqkf8yLaP/Nhok/zUbL/8yJXr/MSiQ/zIiY/81GSH/NRst/zMfUv80HDX/Mxw+/zEiaP82GSD/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR//NSJb/zQpiv81I1z/NRw5/zIwuf8zLJr/NRw3/zYYGv80H0n/MimS/zMmfP8wK6j/Lyy4/y8qrv80HT7/Nhke/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBb/NSFS/zMzyP81Ilf/NR09/zIuqP8yMsT/Myh//zEtpf8xLrP/MDLP/zInh/80HDn/MCyx/zMiZv82GBj/Nhgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nxgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GR7/NSh//zQ0zf81JWr/NRsu/zQhVf80I2L/MyeE/zExxf8yLar/NB1D/zInh/8wL7v/NB1A/zYYGf82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GR//NSZv/zQ0yv80MLL/NSZy/zUgT/81IVP/MyiI/zIsnf8xMLz/MS6v/zQfSv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBr/NR07/zUogv80MLX/NDLE/zQyxf8zMb3/Myyj/zQjZ/81Gin/NhgZ/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBn/NhgZ/zUZIv82Gy//NRw0/zUaK/82GB7/NhgY/zYYGv82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/83GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYGv82GBr/Nhga/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/NRgb/zUYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zUYG/81GBv/NRgb/zYYG/81GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/82GBv/Nhgb/zYYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/82GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/NRgb/zYZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUZG/81GBv/NRgb/zUYG/81GBv/NRgb/zUYG/81GBv/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA="


def _script_dir() -> str:
    """Directory containing this .pyw (or .exe when frozen)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(name: str) -> str:
    """Resolve a bundled resource (logo PNG) whether running from source or
    from a PyInstaller one-file EXE (extra files extract to _MEIPASS)."""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", _script_dir())
        return os.path.join(base, name)
    return os.path.join(_script_dir(), name)


def _set_window_icon(root):
    """Set taskbar + titlebar icon from the embedded GFH_Telecom_TBLogo.ico."""
    try:
        import base64, tempfile, atexit
        data = base64.b64decode(ICON_ICO_B64.strip())
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".ico")
        tmp.write(data); tmp.close()
        atexit.register(lambda p=tmp.name: os.path.exists(p) and os.unlink(p))
        root.iconbitmap(default=False, bitmap=tmp.name)
        root.iconbitmap(tmp.name)
        return
    except Exception:
        pass
    # Fallback: use the brand PNG as the window icon
    png_path = _resource_path(LOGO_PNG_NAME)
    try:
        if os.path.exists(png_path) and _PIT is not None:
            root.iconphoto(True, _PIT.PhotoImage(_PI.open(png_path)))
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════════
# STEP 1 — STORE RENAME
# ═══════════════════════════════════════════════════════════════════════════
STORE_RENAMES = [
    ("Eliff",       "E Iliff"),
    ("FM 1960",     "FM1960"),
    ("Mckellips",   "McKellips"),
    ("N 19th",      "N 19"),
    ("N 35th",      "N 35"),
    ("N 51st",      "N 51"),
    ("New Thomas",  "W Thomas"),
    ("Mt View RD",  "Mount View"),
    ("4363 W Fuqua St", "Fuqua"),
]

def step1_store_rename(folder: Path, log):
    log("\n── STEP 1: Store Rename ──────────────────────────────────")
    renamed = skipped = 0
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        new_name = f.name
        for old, new in STORE_RENAMES:
            new_name = new_name.replace(old, new)
        if new_name != f.name:
            dest = f.parent / new_name
            try:
                f.rename(dest)
                log(f"  Renamed: {f.name}  →  {new_name}")
                renamed += 1
            except Exception as e:
                log(f"  ERROR renaming {f.name}: {e}")
        else:
            skipped += 1
    log(f"  Done — {renamed} renamed, {skipped} unchanged.")
    return renamed


# ═══════════════════════════════════════════════════════════════════════════
# STEP 2 — ADD " 2026" SUFFIX
# ═══════════════════════════════════════════════════════════════════════════
def step2_add_2026(folder: Path, log):
    log("\n── STEP 2: Add ' 2026' Suffix ───────────────────────────")
    renamed = skipped = 0
    for f in sorted(folder.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in (".xlsx", ".xls"):
            continue
        stem = f.stem
        if " 2026" in stem:
            log(f"  Skipped (already has ' 2026'): {f.name}")
            skipped += 1
            continue
        new_name = f"{stem} 2026{f.suffix}"
        dest = f.parent / new_name
        try:
            f.rename(dest)
            log(f"  Renamed: {f.name}  →  {new_name}")
            renamed += 1
        except Exception as e:
            log(f"  ERROR renaming {f.name}: {e}")
    log(f"  Done — {renamed} renamed, {skipped} skipped.")
    return renamed


# ═══════════════════════════════════════════════════════════════════════════
# STEP 3 — DELETE 2025 ROWS
# ═══════════════════════════════════════════════════════════════════════════
def _delete_2025_xlsx(path: Path, log) -> str:
    """Delete rows where col A contains '2025' from an .xlsx file."""
    if _load_xlsx is None:
        return "SKIP (openpyxl not installed)"
    wb = _load_xlsx(str(path))
    ws = wb.active
    rows_to_delete = [
        r for r in range(1, ws.max_row + 1)
        if "2025" in str(ws.cell(r, 1).value or "")
    ]
    if not rows_to_delete:
        return "SKIP (no 2025 rows)"
    # Delete bottom-up so row indices stay valid
    for r in reversed(rows_to_delete):
        ws.delete_rows(r)
    wb.save(str(path))
    return f"OK — deleted {len(rows_to_delete)} row(s)"


def _delete_2025_xls(path: Path, log) -> str:
    """
    Delete rows where col A contains '2025' from a legacy .xls file.
    Saves the result as .xlsx (modern format) so Excel / Power Query can open it.
    The original .xls file is removed after successful conversion.
    """
    if xlrd is None:
        return "SKIP (xlrd not installed)"
    if _load_xlsx is None:
        return "SKIP (openpyxl not installed)"

    # Read source with xlrd (only library that reads legacy .xls properly)
    rb = xlrd.open_workbook(str(path), formatting_info=False)
    rs = rb.sheet_by_index(0)

    keep_rows = [
        r for r in range(rs.nrows)
        if "2025" not in str(rs.cell_value(r, 0))
    ]
    deleted = rs.nrows - len(keep_rows)
    if deleted == 0:
        return "SKIP (no 2025 rows)"

    # Write kept rows into a new .xlsx workbook via openpyxl
    import openpyxl
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = rs.name

    for new_r_idx, old_r in enumerate(keep_rows, start=1):
        for c in range(rs.ncols):
            cell_val = rs.cell_value(old_r, c)
            # Convert xlrd floats that look like integers back to int
            if isinstance(cell_val, float) and cell_val == int(cell_val):
                cell_val = int(cell_val)
            new_ws.cell(row=new_r_idx, column=c + 1, value=cell_val)

    # Save as .xlsx next to the original .xls file
    xlsx_path = path.with_suffix(".xlsx")
    new_wb.save(str(xlsx_path))

    # Remove the old .xls so only the clean .xlsx remains
    try:
        path.unlink()
    except Exception:
        pass

    return f"OK — deleted {deleted} row(s), saved as {xlsx_path.name}"


def step3_delete_2025(folder: Path, log):
    log("\n── STEP 3: Delete 2025 Rows ─────────────────────────────")
    modified = skipped = errors = 0
    # Process both .xls (convert to .xlsx) and .xlsx (edit in-place)
    xls_files = [f for f in sorted(folder.iterdir())
                 if f.is_file() and f.suffix.lower() in (".xls", ".xlsx")]
    if not xls_files:
        log("  No Excel files found.")
        return 0
    for f in xls_files:
        log(f"  Processing: {f.name}")
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                ext = f.suffix.lower()
                if ext == ".xlsx":
                    result = _delete_2025_xlsx(f, log)
                else:
                    result = _delete_2025_xls(f, log)
                log(f"    {result}")
                if result.startswith("OK"):
                    modified += 1
                else:
                    skipped += 1
                break
            except Exception as e:
                log(f"    Attempt {attempt}/{max_retries} failed: {e}")
                if attempt < max_retries:
                    time.sleep(2)
                else:
                    log(f"    FINAL FAILURE — skipping file.")
                    errors += 1
    log(f"  Done — {modified} modified, {skipped} skipped, {errors} errors.")
    return modified


# ═══════════════════════════════════════════════════════════════════════════
# STEP 4 — CONVERT LEGACY EXCEL (.xls / .xlsm / .xlt / .xlsb → .xlsx)
# Merged in from gfh_xls_to_xlsx.pyw so the rebate folder can be fully
# modernised in one pass. Uses real Microsoft Excel via COM so formatting,
# formulas and data are preserved exactly.
# ═══════════════════════════════════════════════════════════════════════════
XLSX_FILE_FORMAT = 51        # xlOpenXMLWorkbook
LEGACY_EXTS = (".xls", ".xlsm", ".xlt", ".xlsb", ".xlc")   # convert these → .xlsx
EXCEL_RESTART_EVERY = 60     # restart Excel periodically to avoid memory bloat

_CANCEL_CONVERT = threading.Event()


def _find_legacy_files(folder: Path, recurse: bool):
    out = []
    if recurse:
        for root, _, files in os.walk(folder):
            for f in files:
                out.append(Path(root) / f)
    else:
        out = [f for f in folder.iterdir() if f.is_file()]
    res = []
    for p in out:
        if p.name.startswith("~$"):
            continue                    # Excel lock files
        if p.suffix.lower() in LEGACY_EXTS:
            res.append(p)
    return sorted(res)


class _ExcelConverter:
    """Thin wrapper around the Excel COM instance that converts legacy
    Excel files to .xlsx. Recycles the Excel process periodically."""

    def __init__(self, log):
        self.log = log
        self.xl = None
        self._opened = 0

    def _start_excel(self):
        if _win32com_client is None:
            raise RuntimeError("pywin32 not installed — cannot drive Excel.")
        self.xl = _win32com_client.DispatchEx("Excel.Application")
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        try: self.xl.AutomationSecurity = 3   # block macros from prompting
        except Exception: pass
        try: self.xl.AskToUpdateLinks = False
        except Exception: pass

    def _stop_excel(self):
        if self.xl is not None:
            try: self.xl.Quit()
            except Exception: pass
        self.xl = None

    def _recycle_if_needed(self):
        self._opened += 1
        if self._opened % EXCEL_RESTART_EVERY == 0:
            self._stop_excel(); time.sleep(1); self._start_excel()

    def convert_one(self, path: Path, overwrite: bool, delete_original: bool) -> str:
        out = path.with_suffix(".xlsx")
        if out.exists() and not overwrite:
            return "skip"
        wb = None
        try:
            try:
                wb = self.xl.Workbooks.Open(os.fspath(path.resolve()),
                                            UpdateLinks=0, ReadOnly=True)
            except Exception:
                # corrupt/odd file → try Excel's repair-open
                wb = self.xl.Workbooks.Open(os.fspath(path.resolve()),
                                            UpdateLinks=0, ReadOnly=True,
                                            CorruptLoad=1)
            wb.SaveAs(os.fspath(out.resolve()), FileFormat=XLSX_FILE_FORMAT)
            wb.Close(False); wb = None
            self._recycle_if_needed()
            if delete_original:
                try: path.unlink()
                except Exception as e:
                    self.log(f"      (kept original — delete failed: {e})")
            return "ok"
        except Exception as e:
            if wb is not None:
                try: wb.Close(False)
                except Exception: pass
            # a bad file can wedge the instance — recycle it
            try: self._stop_excel(); self._start_excel()
            except Exception: pass
            return f"error: {e}"


def step4_convert_legacy_excel(folder: Path, log,
                               recurse: bool = False,
                               overwrite: bool = False,
                               delete_original: bool = False):
    """Convert every legacy .xls/.xlsm/.xlt/.xlsb in `folder` to .xlsx."""
    log("\n── STEP 4: Convert Legacy Excel → .xlsx ────────────────")
    if _win32com_client is None:
        log("  SKIP — pywin32 not installed. Install with: pip install pywin32")
        return 0

    _CANCEL_CONVERT.clear()
    files = _find_legacy_files(folder, recurse)
    if not files:
        log("  No legacy Excel files found in folder.")
        return 0

    log(f"  Found {len(files)} legacy Excel file(s). Starting Excel...")
    conv = _ExcelConverter(log)
    try:
        conv._start_excel()
    except Exception as e:
        log(f"  ERROR starting Excel: {e}")
        log("  Make sure Microsoft Excel is installed on this machine.")
        return 0

    ok = skip = err = 0
    try:
        for i, p in enumerate(files, 1):
            if _CANCEL_CONVERT.is_set():
                log("  ⏹ Cancelled by user.")
                break
            log(f"  [{i}/{len(files)}] {p.name}")
            r = conv.convert_one(p, overwrite, delete_original)
            if r == "ok":
                ok += 1
                log(f"      ✅ converted → {p.stem}.xlsx")
            elif r == "skip":
                skip += 1
                log(f"      ↷ skipped (.xlsx already exists)")
            else:
                err += 1
                log(f"      ❌ {r}")
    finally:
        conv._stop_excel()

    log(f"  Done — {ok} converted, {skip} skipped, {err} errors.")
    return ok


# ═══════════════════════════════════════════════════════════════════════════
# GUI  (styled to match GFH_Inventory_Aging_Processor.pyw)
# ═══════════════════════════════════════════════════════════════════════════
class App:
    def __init__(self, root):
        self.root = root
        self._q = queue.Queue()
        self._running = False
        self._logo_img = None

        root.title("GFH Telecom — Rebate Folder Tools")
        # Dynamic screen resolution support: size to 90% of the screen and
        # center it (DPI-aware), then stay a normal resizable top-level so
        # Windows Snap (50% left/right, corners, Win+arrow) keeps working.
        self._apply_dynamic_geometry()
        root.configure(bg=LIGHT)
        _set_window_icon(root)

        self.theme_manager = ThemeManager("GFH Rebate Folder Tools")
        self._styles()
        self._header()
        self._body()
        self._copyright_bar()
        apply_theme_to_window(self.root, self.theme_manager)
        self._poll()

    def _apply_dynamic_geometry(self) -> None:
        """Size the window to 90% of the screen and center it.

        Works on any laptop/monitor/PC (1080p, 1440p, 2K, 4K) and respects
        Windows DPI scaling (run after _enable_dpi_awareness()). The window
        stays resizable so Windows Snap gestures keep working — it centers
        on launch, then snaps normally to 50% left/right, corners or via
        Win+arrow shortcuts.
        """
        try:
            root = self.root
            root.update_idletasks()
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            w = max(640, min(int(sw * 0.90), sw - 20))
            h = max(480, min(int(sh * 0.90), sh - 40))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            root.geometry(f"{w}x{h}+{x}+{y}")
            # minsize <= half the screen so 50% / corner snap is never blocked
            root.minsize(min(660, max(480, sw // 2)),
                         min(540, max(400, sh // 2)))
            root.resizable(True, True)
        except Exception:
            pass

    # ── styles ─────────────────────────────────────────────────────────────
    def _styles(self):
        s = ttk.Style(); s.theme_use("clam")
        s.configure("Run.TButton", background=RED, foreground=WHITE,
                    font=("Calibri", 11, "bold"), padding=(16, 9), borderwidth=0)
        s.map("Run.TButton",
              background=[("active", "#c01820"), ("disabled", "#aaa")])
        s.configure("Browse.TButton", background=NAVY, foreground=WHITE,
                    font=("Calibri", 10), padding=(10, 6), borderwidth=0)
        s.map("Browse.TButton", background=[("active", "#1a2550")])
        s.configure("Cancel.TButton", background="#1a2550", foreground=WHITE,
                    font=("Calibri", 10), padding=(10, 6), borderwidth=0)
        s.map("Cancel.TButton", background=[("active", "#2a3560")])
        s.configure("Accent.Horizontal.TProgressbar",
                    troughcolor="#dde6f0", background=RED, borderwidth=0)

    # ── header (matches Aging Processor: NAVY 108px, logo left, title center) ──
    def _header(self):
        hdr = tk.Frame(self.root, bg=NAVY, height=108)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        hdr._tag = "header"

        # Logo on the left — load GFH_Telecom_Logo.png next to this script,
        # composite on NAVY (so transparent regions render correctly), and
        # thumbnail to 260x82 (same recipe as the Aging Processor).
        logo_path = _resource_path(LOGO_PNG_NAME)
        if os.path.exists(logo_path) and _PI is not None:
            try:
                img = _PI.open(logo_path).convert("RGBA")
                bg2 = _PI.new("RGBA", img.size, (9, 13, 38, 255))
                bg2.paste(img, mask=img.split()[3])
                img = bg2.convert("RGB")
                img.thumbnail((260, 82), _PI.Resampling.LANCZOS)
                self._logo_img = _PIT.PhotoImage(img)
            except Exception:
                self._logo_img = None

        lf = tk.Frame(hdr, bg=NAVY)
        lf.place(relx=0, rely=0.5, anchor="w", x=24)
        lf._tag = "header"
        if self._logo_img:
            tk.Label(lf, image=self._logo_img, bg=NAVY).pack()
        else:
            tk.Label(lf, text="GFH TELECOM", font=("Calibri", 16, "bold"),
                     fg=RED, bg=NAVY).pack()

        tf = tk.Frame(hdr, bg=NAVY)
        tf.place(relx=0.58, rely=0.5, anchor="center")
        tf._tag = "header"
        tk.Label(tf, text="REBATE FOLDER TOOLS",
                 font=("Calibri", 18, "bold"), fg=WHITE, bg=NAVY).pack()
        tk.Label(tf, text="Rename · Suffix · Delete 2025 · Convert Legacy Excel",
                 font=("Calibri", 9), fg=WHITE, bg=NAVY).pack()

        theme_btn = create_theme_toggle_button(hdr, self.theme_manager)
        theme_btn.place(relx=0.98, rely=0.5, anchor="e")

    def _apply_theme(self, colors=None):
        apply_theme_to_window(self.root, self.theme_manager)

    # ── body ───────────────────────────────────────────────────────────────
    def _body(self):
        body = tk.Frame(self.root, bg=LIGHT)
        body.pack(fill="both", expand=True, padx=24, pady=18)

        # Folder row
        row = tk.Frame(body, bg=LIGHT)
        row.pack(fill="x", pady=(0, 14))
        row.columnconfigure(0, weight=1)
        self.folder_var = tk.StringVar(value="")
        tk.Entry(row, textvariable=self.folder_var,
                 font=("Calibri", 9), relief="flat", bg="#e8eff8", fg=NAVY,
                 readonlybackground="#e8eff8",
                 highlightbackground="#b0c4de", highlightthickness=1
                 ).grid(row=0, column=0, sticky="ew", ipady=5, padx=(0, 8))
        ttk.Button(row, text="Browse", style="Browse.TButton",
                   command=self._browse).grid(row=0, column=1)

        # Step checkboxes — 4 options now (Step 4 = Convert Legacy Excel)
        opts = tk.Frame(body, bg=LIGHT)
        opts.pack(fill="x", pady=(0, 10))
        self.do_step1 = tk.BooleanVar(value=True)
        self.do_step2 = tk.BooleanVar(value=True)
        self.do_step3 = tk.BooleanVar(value=True)
        self.do_step4 = tk.BooleanVar(value=True)
        for var, label in [
            (self.do_step1, "1. Store Rename"),
            (self.do_step2, "2. Add ' 2026' Suffix"),
            (self.do_step3, "3. Delete 2025 Rows"),
            (self.do_step4, "4. Convert Legacy Excel → .xlsx"),
        ]:
            tk.Checkbutton(opts, text=label, variable=var,
                           bg=LIGHT, fg=NAVY, selectcolor=WHITE,
                           activebackground=LIGHT, activeforeground=NAVY,
                           font=("Calibri", 10)).pack(side="left", padx=(0, 16))

        # Step-4 sub-options (only relevant when Step 4 is ticked)
        sub = tk.Frame(body, bg=LIGHT)
        sub.pack(fill="x", pady=(0, 10))
        self.recurse_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=True)
        self.delete_original_var = tk.BooleanVar(value=True)
        for txt, var in [
            ("Include subfolders (Step 4)", self.recurse_var),
            ("Overwrite existing .xlsx (Step 4)", self.overwrite_var),
            ("Delete original after converting (Step 4)", self.delete_original_var),
        ]:
            tk.Checkbutton(sub, text=txt, variable=var,
                           bg=LIGHT, fg="#4a6080", selectcolor=WHITE,
                           activebackground=LIGHT, activeforeground=NAVY,
                           font=("Calibri", 9)).pack(side="left", padx=(0, 14))

        # Run + Cancel buttons
        act = tk.Frame(body, bg=LIGHT)
        act.pack(fill="x", pady=(0, 10))
        self.run_btn = ttk.Button(act, text="▶  Run Selected Steps",
                                  style="Run.TButton", command=self._start)
        self.run_btn.pack(side="left")
        self.cancel_btn = ttk.Button(act, text="⏹  Cancel Step 4",
                                     style="Cancel.TButton",
                                     command=lambda: _CANCEL_CONVERT.set(),
                                     state="disabled")
        self.cancel_btn.pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(act, textvariable=self.status_var, bg=LIGHT, fg=NAVY,
                 font=("Calibri", 9)).pack(side="left", padx=12)

        # Progress bar
        self.progress = ttk.Progressbar(body, mode="indeterminate",
                                        style="Accent.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 10))

        # Log
        tk.Label(body, text="Activity Log", font=("Calibri", 9, "bold"),
                 fg=NAVY, bg=LIGHT).pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            body, height=12, font=("Consolas", 8),
            bg=LOG_BG, fg=LOG_FG, relief="flat", state="disabled", wrap="word"
        )
        self.log_box.pack(fill="both", expand=True)

    def _copyright_bar(self):
        bar = tk.Frame(self.root, bg=NAVY, height=26)
        bar.pack(fill="x", side="bottom"); bar.pack_propagate(False)
        tk.Label(bar, text=COPYRIGHT_TEXT, bg=NAVY, fg="#9d9db8",
                 font=("Calibri", 8)).pack(pady=4)

    # ── helpers ────────────────────────────────────────────────────────────
    def _browse(self):
        d = filedialog.askdirectory(title="Select Rebate folder")
        if d:
            self.folder_var.set(d)

    def _log(self, msg):
        self._q.put(msg)

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                self.log_box.config(state="normal")
                self.log_box.insert("end", msg + "\n")
                self.log_box.see("end")
                self.log_box.config(state="disabled")
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _start(self):
        if self._running:
            return
        folder = Path(self.folder_var.get().strip())
        if not folder.exists():
            messagebox.showerror("Folder not found", str(folder))
            return
        # If Step 4 is ticked but pywin32 isn't available, warn early.
        if self.do_step4.get() and _win32com_client is None:
            if not messagebox.askyesno(
                "Excel driver missing",
                "Step 4 needs pywin32 + Microsoft Excel, which doesn't appear to be installed.\n\n"
                "Run Step 4 anyway? (It will skip itself if Excel can't start.)"
            ):
                return
        self._running = True
        self.run_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.progress.start(12)
        threading.Thread(target=self._worker, args=(folder,), daemon=True).start()

    def _worker(self, folder: Path):
        try:
            self.status_var.set("Running…")
            self._log("═" * 55)
            self._log(f"Folder: {folder}")
            self._log(f"Steps: "
                      f"{'1 ' if self.do_step1.get() else ''}"
                      f"{'2 ' if self.do_step2.get() else ''}"
                      f"{'3 ' if self.do_step3.get() else ''}"
                      f"{'4' if self.do_step4.get() else ''}".rstrip())
            self._log("═" * 55)

            if self.do_step1.get():
                step1_store_rename(folder, self._log)
            if self.do_step2.get():
                step2_add_2026(folder, self._log)
            if self.do_step3.get():
                step3_delete_2025(folder, self._log)
            if self.do_step4.get():
                step4_convert_legacy_excel(
                    folder, self._log,
                    recurse=self.recurse_var.get(),
                    overwrite=self.overwrite_var.get(),
                    delete_original=self.delete_original_var.get(),
                )

            self._log("")
            self._log("✓ COMPLETE — all selected steps finished.")
            self.status_var.set("Done.")
        except Exception as e:
            self._log(f"\nCRITICAL ERROR: {e}")
            self._log(traceback.format_exc())
            self.status_var.set("Error — see log.")
        finally:
            self.root.after(0, self._done)

    def _done(self):
        self.progress.stop()
        self.run_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self._running = False


# ═══════════════════════════════════════════════════════════════════════════
def _enable_dpi_awareness() -> None:
    """Make Windows report physical pixels so winfo_screen* is accurate on
    high-DPI displays (1080p, 1440p, 2K, 4K, DPI-scaled laptops)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


if __name__ == "__main__":
    _enable_dpi_awareness()
    try:
        root = tk.Tk()
        App(root)
        root.mainloop()
    except Exception:
        traceback.print_exc()
        try:
            from tkinter import messagebox as _mb
            _mb.showerror("Fatal Error", traceback.format_exc())
        except Exception:
            pass
