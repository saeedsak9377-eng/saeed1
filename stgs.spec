# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for STGS — Smart Test Generation System
# Run:  pyinstaller stgs.spec

import sys
from pathlib import Path

block_cipher = None

# ── Collect data files that third-party libraries need ───────────────────────
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# openpyxl ships internal XML templates / style data
openpyxl_datas = collect_data_files('openpyxl')

# matplotlib needs its mpl-data folder (fonts, style sheets, backends)
mpl_datas = collect_data_files('matplotlib')

# Merge all data tuples
all_datas = openpyxl_datas + mpl_datas

# ── Hidden imports that PyInstaller misses via static analysis ────────────────
hidden = [
    # openpyxl
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.styles.stylesheet',
    'openpyxl.chart',
    'openpyxl.utils',
    'openpyxl.utils.dataframe',
    'openpyxl.writer',
    'openpyxl.reader',
    # xlsxwriter
    'xlsxwriter',
    'xlsxwriter.workbook',
    'xlsxwriter.worksheet',
    'xlsxwriter.chart',
    'xlsxwriter.chartsheet',
    'xlsxwriter.format',
    # pandas
    'pandas',
    'pandas.core.dtypes',
    'pandas.io.formats.style',
    'pandas._libs.tslibs.base',
    # matplotlib
    'matplotlib',
    'matplotlib.backends.backend_agg',
    'matplotlib.backends.backend_tkagg',
    'matplotlib.pyplot',
    'matplotlib.cm',
    # numpy
    'numpy',
    'numpy.core._multiarray_umath',
    'numpy.random',
    # stdlib
    'tkinter',
    'tkinter.ttk',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'json',
    'math',
    'queue',
    'threading',
    'subprocess',
    'pathlib',
    're',
    'datetime',
    'io',
    'logging',
    'collections',
    'dataclasses',
]

a = Analysis(
    ['stgs_unified.py'],
    pathex=[],
    binaries=[],
    datas=all_datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude heavy unused packages to reduce .exe size
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'PIL',
        'cv2',
        'sklearn',
        'tensorflow',
        'torch',
        'wx',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher,
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='STGS',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,            # compress with UPX if available (reduces size ~30%)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # no black console window — GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='stgs_icon.ico',  # uncomment and add your .ico file here
)
