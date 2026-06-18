# PyInstaller spec for Watermark Lab
#
# Build:   pyinstaller WatermarkLab.spec
# Output:  dist\WatermarkLab.exe   (single, self-contained, portable)
#
# Size strategy:
#   - ffmpeg is NOT bundled. On first video use the app downloads the latest
#     minimal ffmpeg build from GitHub (GyanD/codexffmpeg) and caches it in
#     %LOCALAPPDATA%\WatermarkLab\ffmpeg.exe. Subsequent runs use the cache.
#   - Unused Pillow binary extensions (_avif, _webp, _imagingcms) are stripped.
#   - Heavy stdlib modules that are never imported are excluded.
#
# Portability:
#   runtime_tmpdir='.' extracts the _MEI<pid> folder next to the exe so it
#   works from USB sticks, network shares, and WDAC-locked machines.

import os

block_cipher = None

# Pillow extensions we actively use: _imaging (core), _imagingft (FreeType).
# _avif (~7.5 MB), _webp (~0.4 MB), _imagingcms (~0.3 MB) are unused.
_PILLOW_STRIP = {"_avif", "_webp", "_imagingcms"}

a = Analysis(
    ['Watermark_Lab.pyw'],
    pathex=[],
    binaries=[],
    datas=[
        ('SplashLab.png', '.'),
        ('Watermark.png', '.'),
        ('Watermark.ico', '.'),
    ],
    hiddenimports=['packaging', 'packaging.version', '_ffmpeg'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Test / dev tools — never needed at runtime
        'unittest', 'doctest', 'pdb', 'pydoc',
        # Truly unused — no database, no multiprocessing, no XML
        'sqlite3', '_sqlite3',
        'multiprocessing',
        'xml.etree',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Strip unused Pillow binary extensions
a.binaries = [
    (name, src, kind)
    for name, src, kind in a.binaries
    if not any(name.lower().startswith(f"pil/{stub}") or
               name.lower().startswith(stub)
               for stub in _PILLOW_STRIP)
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='WatermarkLab',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir='.',    # extract next to the exe, not %TEMP%
    console=False,         # windowed GUI app
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='Watermark.ico',
)

