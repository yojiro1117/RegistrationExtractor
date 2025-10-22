# -*- mode: python ; coding: utf-8 -*-

# PyInstaller spec file for RegistrationExtractor
#
# この spec は Windows 環境での単一 EXE 化を目的としており、必要に応じてバイナリやデータファイルを同梱します。

block_cipher = None


a = Analysis(
    ['app/cli.py'],
    pathex=[],
    binaries=[],
    datas=[('app/settings.json', 'app'), ('docs/screenshot.png', 'docs')],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='RegistrationExtractor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='RegistrationExtractor',
)