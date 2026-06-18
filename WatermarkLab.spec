# PyInstaller spec for Watermark Lab — onedir build
#
# Build:   pyinstaller WatermarkLab.spec
# Output:  dist\WatermarkLab\   (folder)
# Release: dist\WatermarkLab.zip (created by release.yml)
#
# Why onedir instead of onefile:
#   Onefile extracts a _MEI<pid> folder at every launch. When the exe lives
#   on OneDrive the sync engine locks those files causing a TclError crash.
#   Extracting to %TEMP% instead triggers WDAC which blocks DLL loads.
#   Onedir has no runtime extraction — all files are pre-extracted at build
#   time, so neither problem applies.

import os

block_cipher = None

a = Analysis(
    ['Watermark_Lab.pyw'],
    pathex=[],
    binaries=[],
    datas=[
        ('SplashLab.png', '.'),
        ('Watermark.png', '.'),
        ('Watermark.ico', '.'),
    ],
    hiddenimports=[
        'packaging', 'packaging.version',
        '_ffmpeg',
        'PIL._imaging',      # Pillow core C extension — must be explicit
        'PIL._imagingft',    # FreeType font renderer
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        'PIL.PngImagePlugin',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'unittest', 'doctest', 'pdb', 'pydoc',
        'sqlite3', '_sqlite3',
        'multiprocessing',
        'xml.etree',
    ],
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
    name='WatermarkLab',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Watermark.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='WatermarkLab',
)
