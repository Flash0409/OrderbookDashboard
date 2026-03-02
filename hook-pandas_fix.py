# Custom pandas hook to avoid isolated subprocess issues
# This replaces the default hook-pandas.py that causes OSError

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Manually list the hidden imports instead of using collect_submodules with isolated subprocess
hiddenimports = [
    'pandas._libs',
    'pandas._libs.tslibs',
    'pandas._libs.tslibs.base',
    'pandas._libs.tslibs.parsing',
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.offsets',
    'pandas._libs.tslibs.timedeltas',
    'pandas._libs.tslibs.timestamps',
    'pandas._libs.tslibs.tzconversion',
    'pandas._libs.hashtable',
    'pandas._libs.lib',
    'pandas._libs.missing',
    'pandas._libs.properties',
    'pandas._libs.reshape',
    'pandas._libs.algos',
    'pandas._libs.groupby',
    'pandas._libs.ops',
    'pandas._libs.join',
    'pandas._libs.index',
    'pandas._libs.internals',
    'pandas._libs.writers',
    'pandas._libs.testing',
    'pandas._libs.sparse',
    'pandas._libs.ops_dispatch',
    'pandas._libs.arrays',
    'pandas.io.formats.style',
]

# Collect data files (without isolated subprocess)
try:
    datas = collect_data_files('pandas', include_py_files=False)
except:
    datas = []
