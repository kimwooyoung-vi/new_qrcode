# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['core\\qr_reader\\main.py'],
    pathex=[],
    binaries=[('./dlls/libiconv.dll', 'pyzbar'), ('./dlls/libzbar-64.dll', 'pyzbar'),('./dlls/qwindows.dll','qwindows')],
    datas=[
        ('frontend', 'frontend'),
        ('core','core'),
        ('./meiryo.ttc','./'),
        ('./scan_logo.png','./'),
        ('./loading.gif','./'),
        ('./beep.wav','./'),
    ],
    hiddenimports=['qwindows','pandas','json','qrcode','cv2','pyzbar.pyzbar','PyQt6.QtMultimedia', 'PyQt6.QtCore', 'PyQt6.QtGui', 'PyQt6.QtWidgets'],
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
    name='main',
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
