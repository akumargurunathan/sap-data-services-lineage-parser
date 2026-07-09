# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for the DS XML Lineage Extractor GUI.

Builds a single, self-contained Windows .exe that a non-technical user can run
by double-clicking — no Python install, no virtual environment, no pip.

Build:   pyinstaller DS_Lineage_Extractor.spec --noconfirm
Output:  dist/DS_Lineage_Extractor.exe
"""
from PyInstaller.utils.hooks import collect_submodules

# The engine reaches these modules through dynamic ``import`` statements inside
# functions; list them explicitly so PyInstaller never drops one.
hiddenimports = (
    # pandas imports openpyxl lazily (engine='openpyxl') for .xlsx export, so
    # PyInstaller's static analysis can miss it — name it explicitly.
    ['config', 'ds_xml_engine', 'openpyxl']
    + collect_submodules('ds_engine')
    + collect_submodules('semantic_extension')
)

a = Analysis(
    ['ds_ui_launcher.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'PyInstaller'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='DS_Lineage_Extractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # GUI app: no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
