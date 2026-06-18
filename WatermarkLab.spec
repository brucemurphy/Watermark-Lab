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
        # Dev tools
        'unittest', 'doctest', 'pdb', 'pydoc',
        # No database
        'sqlite3', '_sqlite3',
        # No multiprocessing
        'multiprocessing',
        # No XML
        'xml.etree',
        # numpy / scipy — pulled in by Pillow on Python 3.14 but not needed
        # for PNG rendering (only _imaging + _imagingft are required)
        'numpy', 'scipy',
        # Timezone data — not needed (we do no tz-aware datetime work)
        'tzdata',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------------------------------------------------------------------------
# Strip unused Pillow binary extensions from the collected binaries.
# _avif (~7.5 MB), _webp (~0.4 MB), _imagingcms (~0.3 MB) are format
# decoders we never use — only _imaging (core) and _imagingft (FreeType)
# are needed for PNG rendering and text drawing.
# ---------------------------------------------------------------------------
_PILLOW_STRIP = {'_avif', '_webp', '_imagingcms'}
a.binaries = [
    (name, src, kind) for name, src, kind in a.binaries
    if not any(
        os.path.basename(src).lower().startswith(stub)
        for stub in _PILLOW_STRIP
    )
]

# ---------------------------------------------------------------------------
# Strip the bulk of the Tcl/Tk data tree.  The app uses dark-themed ttk
# widgets and a PNG splash — it needs core Tk, ttk themes, and the default
# encoding.  Everything else (locales, demos, extra encodings, old http
# packages, opt packages) is dead weight.
# ---------------------------------------------------------------------------
_TCL_KEEP = {
    'init.tcl', 'tk.tcl', 'ttk', 'button.tcl', 'entry.tcl',
    'scale.tcl', 'scrollbar.tcl', 'listbox.tcl', 'dialog.tcl',
    'msgbox.tcl', 'panedwindow.tcl', 'tearoff.tcl', 'menu.tcl',
}

def _keep_tcl(dst_path: str, src_path: str) -> bool:
    """Return True if this Tcl/Tk data file should be kept."""
    # Check destination path (what ends up in the output folder)
    parts = dst_path.replace('\\', '/').lower().split('/')
    # Always keep files not in the tcl/tk data trees
    if '_tcl_data' not in parts and '_tk_data' not in parts:
        return True
    # Drop the entire Tcl tzdata tree (America, Europe, Asia etc.)
    if 'tzdata' in parts:
        return False
    # Keep ttk theme directory entirely
    if 'ttk' in parts:
        return True
    # Keep individual core files
    fname = os.path.basename(dst_path)
    if fname in _TCL_KEEP:
        return True
    return False

a.datas = [(dst, src, kind) for dst, src, kind in a.datas if _keep_tcl(dst, src)]

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
