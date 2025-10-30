# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller Spec File for Task Scheduler
# Builds a single-file executable with Windows 11 app container icon
# Usage: pyinstaller TaskScheduler.spec
#

import sys
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# Hidden imports required for the application
hiddenimports = [
    'customtkinter',
    'apscheduler',
    'apscheduler.schedulers',
    'apscheduler.schedulers.background',
    'apscheduler.triggers',
    'apscheduler.triggers.interval',
    'psutil',
    'json',
    'subprocess',
    'threading',
    'datetime',
]

# Collect data files for CustomTkinter theme files
datas = []
try:
    datas += collect_data_files('customtkinter')
except Exception as e:
    print(f"Warning: Could not collect CustomTkinter data files: {e}")

# Build Analysis
a = Analysis(
    ['index.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=['matplotlib', 'numpy', 'scipy'],  # Exclude unnecessary packages
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Create PYZ archive
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Create EXE - Single file, no console window
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='TaskScheduler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window - GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='scheduler_icon.ico',  # Windows 11 app container icon
)
