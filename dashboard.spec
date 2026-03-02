# -*- mode: python ; coding: utf-8 -*-

import os
import importlib
import importlib.metadata

# Locate package directories
streamlit_dir = os.path.dirname(importlib.import_module('streamlit').__file__)
plotly_dir = os.path.dirname(importlib.import_module('plotly').__file__)
site_packages = os.path.dirname(streamlit_dir)

# Collect dist-info metadata for packages that use importlib.metadata.version()
metadata_pkgs = [
    'streamlit', 'plotly', 'pandas', 'numpy', 'pyarrow', 'openpyxl',
    'altair', 'narwhals', 'pydeck', 'tornado', 'click', 'toml',
    'packaging', 'pytz', 'tzdata', 'Pillow', 'protobuf', 'requests',
    'tenacity', 'charset-normalizer', 'certifi', 'urllib3', 'idna',
    'jinja2', 'MarkupSafe', 'jsonschema', 'attrs', 'gitdb', 'GitPython',
    'blinker', 'cachetools', 'watchdog',
]
metadata_datas = []
for pkg in metadata_pkgs:
    try:
        dist = importlib.metadata.distribution(pkg)
        dist_info = str(dist._path)
        dest_name = os.path.basename(dist_info)
        metadata_datas.append((dist_info, dest_name))
    except importlib.metadata.PackageNotFoundError:
        pass

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.py', '.'),
        (streamlit_dir, 'streamlit'),
        (plotly_dir, 'plotly'),
    ] + metadata_datas,
    hiddenimports=[
        'streamlit',
        'streamlit.web.cli',
        'streamlit.runtime',
        'streamlit.runtime.scriptrunner',
        'streamlit.runtime.caching',
        'plotly',
        'plotly.express',
        'plotly.graph_objects',
        'plotly.io',
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.offsets',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.tslibs.tzconversion',
        'pandas._libs.hashtable',
        'pandas._libs.lib',
        'pandas._libs.missing',
        'pandas._libs.algos',
        'pandas._libs.groupby',
        'pandas._libs.ops',
        'pandas._libs.join',
        'pandas._libs.index',
        'pandas._libs.internals',
        'pandas._libs.writers',
        'pandas.io.formats.style',
        'numpy',
        'openpyxl',
        'pyarrow',
        'altair',
        'pydeck',
    ],
    hookspath=[],
    hooksconfig={
        'pandas': {
            'use_isolated': False,  # Disable isolated subprocess for pandas
        }
    },
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

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
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='iCenterDashboard',
)
