# -*- mode: python ; coding: utf-8 -*-

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 收集 Streamlit 相关的数据文件
streamlit_datas = collect_data_files('streamlit')
streamlit_modules = collect_submodules('streamlit')

# 收集其他依赖的数据文件
datas = [
    ('app', 'app'),
    *streamlit_datas,
]

# 收集所有需要的模块
hiddenimports = [
    'streamlit',
    'streamlit.runtime.scriptrunner',
    'streamlit.runtime.scriptrunner.script_runner',
    'streamlit.runtime.caching',
    'streamlit.components',
    'pandas',
    'numpy',
    'docx',
    'openpyxl',
    'lxml',
    'packaging',
    'watchdog',
    'tornado',
    'altair',
    'jsonschema',
    'gitpython',
    'requests',
    'click',
    'plotly',
    'PIL',
    'PIL._tkinter_finder',
]

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'sklearn',
        'tensorflow',
        'torch',
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
    name='WordReplace',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if sys.platform == 'win32' else None,
)
