# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[('./dlls/libiconv.dll', 'pyzbar'), ('./dlls/libzbar-64.dll', 'pyzbar'),('./dlls/qwindows.dll','qwindows')],
    datas=[
        ('./meiryo.ttc','./'),
        ('./scan_logo.png','./'),
        ('./loading.gif','./'),
        ],
    hiddenimports=['cv2', 'pyzbar','qwindows'],
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
    name='QRCode',
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
