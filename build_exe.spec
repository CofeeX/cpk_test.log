# -*- mode: python ; coding: utf-8 -*-
# PyInstaller 打包配置: CPK-PCBA

import os
from datetime import datetime

block_cipher = None
project_dir = os.path.abspath('.')

# 生成带时间后缀的 exe 文件名 (区分版本): CPK-PCBA_YYYYMMDD_HHMM
_exe_timestamp = datetime.now().strftime("%Y%m%d_%H%M")
_exe_name = f'CPK-PCBA_{_exe_timestamp}'

a = Analysis(
    ['app.py'],
    pathex=[project_dir],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
        ('configs', 'configs'),
    ],
    hiddenimports=[
        'flask_sqlalchemy',
        'openpyxl',
        'openpyxl.cell._writer',
        'xlrd',
        'openpyxl.workbook',
        'sqlalchemy',
        'sqlalchemy.sql.default_comparator',
        'sqlalchemy.dialects.sqlite',
        'pandas',
        'pandas._libs',
        'pandas._libs.tslibs',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.timestamps',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'PyQt5',
        'PySide2',
        'IPython',
        'notebook',
        'jupyter',
        'pytest',
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
    name=_exe_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        'vcruntime140.dll',
        'python3.dll',
        'python313.dll',
    ],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
