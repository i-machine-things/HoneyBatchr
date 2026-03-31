# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Honey Batchr - one-dir build (fast startup)

from PyInstaller.utils.hooks import collect_data_files

qt_data = collect_data_files('PyQt6')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=qt_data + [('resources', 'resources')],
    hiddenimports=[
        'PyQt6.sip',
        'PyQt6.QtPrintSupport',
        'PyQt6.QtGui',
        'PyQt6.QtCore',
        'PyQt6.QtWidgets',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name='HoneyBatchr',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['resources/badger.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='HoneyBatchr',
)
