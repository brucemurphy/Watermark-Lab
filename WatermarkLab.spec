# PyInstaller spec for Watermark Lab
#
# Build:   pyinstaller WatermarkLab.spec
# Output:  dist\WatermarkLab.exe   (single, self-contained, portable)
#
# Size strategy (target: <60 MB):
#   - ffmpeg is supplied by imageio-ffmpeg (~30 MB minimal build).
#     No user-supplied ffmpeg.exe / ffprobe.exe needed at build time.
#   - Unused Pillow binary extensions (_avif, _webp, _imagingcms) are
#     stripped from the binary list after analysis (~10 MB saving).
#   - Heavy stdlib modules that are never imported are excluded.
#   - UPX compression is intentionally left OFF: it triggers false-positive
#     AV alerts and slows startup without meaningful size gains on already-
#     compressed data.
#
# Portability:
#   runtime_tmpdir='.' extracts the _MEI<pid> folder next to the exe instead
#   of %TEMP%, so it works from USB sticks, network shares, and machines
#   running Windows Application Control / WDAC.

import os
import sys

# ---------------------------------------------------------------------------
# Locate the imageio-ffmpeg binary at spec-evaluation time so PyInstaller
# can embed it. Falls back to a user-supplied ffmpeg.exe beside the spec.
# ---------------------------------------------------------------------------
try:
    import imageio_ffmpeg
    _ffmpeg_src = imageio_ffmpeg.get_ffmpeg_exe()
except Exception:
    _ffmpeg_src = None

if not _ffmpeg_src or not os.path.isfile(_ffmpeg_src):
    _ffmpeg_src = os.path.join(SPECPATH, "ffmpeg.exe")

if not os.path.isfile(_ffmpeg_src):
    raise FileNotFoundError(
        "ffmpeg not found. Run: pip install imageio-ffmpeg\n"
        "or place ffmpeg.exe next to WatermarkLab.spec."
    )

block_cipher = None

# Pillow extensions we actively use: _imaging (core), _imagingft (FreeType fonts).
# _avif (~7.5 MB), _webp (~0.4 MB), _imagingcms (~0.3 MB) are unused.
_PILLOW_STRIP = {"_avif", "_webp", "_imagingcms"}

a = Analysis(
    ['Watermark_Lab.pyw'],
    pathex=[],
    binaries=[
        (_ffmpeg_src, '.'),
    ],
    datas=[
        ('SplashLab.png', '.'),
        ('Watermark.png', '.'),
        ('Watermark.ico', '.'),
    ],
    hiddenimports=['packaging', 'packaging.version', 'imageio_ffmpeg'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Test frameworks
        'unittest', 'doctest', 'pdb', 'pydoc',
        # Unused stdlib heavyweights
        'email', 'html', 'http', 'urllib.robotparser',
        'xmlrpc', 'ftplib', 'imaplib', 'poplib', 'smtplib', 'telnetlib',
        'sqlite3', '_sqlite3',
        'multiprocessing',
        'concurrent.futures',
        'asyncio',
        'xml.etree',
        'difflib',
        'calendar',
        'csv',
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

