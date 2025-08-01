# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['file_report_app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['tkcalendar','pandas','openpyxl','win32com.client'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FileReportApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False  # <-- hides the console
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='FileReportApp'
)