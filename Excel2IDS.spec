# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Excel2IDS.py'],
    pathex=[],
    binaries=[],
    datas=[
		('.\\.venv\\Lib\\site-packages\\xmlschema\\schemas\\XSD_1.0', 'xmlschema/schemas/XSD_1.0'), 
		('.\\.venv\\Lib\\site-packages\\xmlschema\\schemas\\XML', 'xmlschema/schemas/XML'), 
		('.\\.venv\\Lib\\site-packages\\xmlschema\\schemas\\XSI', 'xmlschema/schemas/XSI'), 
		('.\\.venv\\Lib\\site-packages\\xmlschema\\schemas\\VC', 'xmlschema/schemas/VC'), 
		('.\\.venv\\Lib\\site-packages\\xmlschema\\schemas\\XSD_1.1', 'xmlschema/schemas/XSD_1.1'),
		('.\\.venv\\Lib\\site-packages\\ifctester', 'ifctester'),
		],
    hiddenimports=[],
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
    name='Excel2IDS',
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
	icon='C:\\Code\\Excel2IDS\\ids-logo.ico',
)
