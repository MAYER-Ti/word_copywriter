# word_copywriter.spec
# -*- mode: python ; coding: utf-8 -*-

import pathlib
root = pathlib.Path.cwd()

a = Analysis(
    ['main.py'],
    pathex=[str(root)],
    binaries=[],
    datas=[
        (root / 'resources' / 'icon.png',     'resources'),
        (root / 'templates' / 'act.xlsx',     'templates'),
        (root / 'templates' / 'invoice.xlsx', 'templates'),
    ],
    hiddenimports=['sip'],
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
    name='word_copywriter',
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
    icon=str(root / 'resources' / 'icon.png'),
)
