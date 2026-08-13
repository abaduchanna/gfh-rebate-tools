# -*- mode: python ; coding: utf-8 -*-
import datetime as _dt
_year = _dt.date.today().year

SPEC_DOC = f"""PyInstaller spec
Developed by Abad Umair Channa \u00a9 {_year}
Build command: pyinstaller rebate_tools.spec
"""


block_cipher = None

a = Analysis(
    ['rebate_tools.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('gfh_icon.ico', '.'),
        ('gfh_icon_white.ico', '.'),
        ('gfh_icon.png', '.'),
        ('gfh_wordmark.png', '.'),
        ('GFH_Telecom_Logo.png', '.'),
        ('stores.json', '.'),
        ('theme_manager.py', '.'),
        ('logo_handler.py', '.'),
        ('header_manager.py', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlutils',
        'PIL',
        'theme_manager',
        'logo_handler',
        'header_manager',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[
        'doctest',
        'pdb',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='rebate_tools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gfh_icon.ico',
)
