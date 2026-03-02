# -*- mode: python ; coding: utf-8 -*-
# Simplified spec file to avoid pandas hook issues

import os
import sys

block_cipher = None

# Get paths
venv_path = os.path.join(os.getcwd(), '.venv')
site_packages = os.path.join(venv_path, 'Lib','site-packages')

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.py', '.'),
    ],
    hiddenimports=[
        'streamlit',
        'streamlit.web.cli',
        'streamlit.runtime',
        'streamlit.runtime.scriptrunner',
        'pandas',
        'numpy',
        'openpyxl',
        'pyarrow',
        'plotly',
        'altair',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
       'pandas.tests',
        'numpy.tests',
        'matplotlib',
        'scipy',
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
    [],
    exclude_binaries=True,
    name='iCenterDashboard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='iCenterDashboard',
)
