# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['generate_reports.py'],
    pathex=[],
    binaries=[],
    datas=[('report_core.py', '.'), ('模板.docx', '.')],
    hiddenimports=['tkinterdnd2', 'xlrd', 'openpyxl'],
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
    name='混铁车中夜班安全检查报告生成工具',
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
