# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_all
from pathlib import Path

project_dir = Path(SPEC).resolve().parent

datas = []
binaries = []
hiddenimports = ['customtkinter', 'PIL._tkinter_finder', 'openpyxl', 'xlrd', 'pandas', 'matplotlib', 'matplotlib.backends.backend_tkagg', 'ui', 'core', 'ingestion', 'export', 'utils', 'app_info']
datas += collect_data_files('matplotlib')
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

for folder in ('assets', 'data'):
    resource_dir = project_dir / folder
    if resource_dir.is_dir():
        datas.append((str(resource_dir), folder))

icon_path = project_dir / 'assets' / 'icono.ico'


a = Analysis(
    [str(project_dir / 'main.py')],
    pathex=[str(project_dir)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SGAM',
    icon=str(icon_path) if icon_path.exists() else None,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
