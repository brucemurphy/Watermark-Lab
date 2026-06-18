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
        'numpy', 'scipy',
        # Timezone data
        'tzdata',
        # Async — we use threads only
        'asyncio', '_asyncio',
        # Compression formats we never open (zipfile only needs zlib)
        '_zstd',
        # Queue — not used
        'queue', '_queue',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------------------------------------------------------------------------
# Strip unused binaries
# ---------------------------------------------------------------------------
_BIN_STRIP = {
    # Pillow extensions we never call
    '_avif', '_webp', '_imagingcms', '_imagingmath', '_imagingtk',
    # win32ui / MFC — we use Tkinter, not the Pythonwin GUI toolkit
    'win32ui', 'mfc140u',
    # win32 internals not needed at runtime
    'win32trace', '_win32sysloader',
}
a.binaries = [
    (name, src, kind) for name, src, kind in a.binaries
    if not any(
        os.path.splitext(os.path.basename(src))[0].lower().startswith(stub)
        for stub in _BIN_STRIP
    )
]

# ---------------------------------------------------------------------------
# Strip the bulk of the Tcl/Tk data tree.  The app uses dark-themed ttk
# widgets and a PNG splash — it needs core Tk, ttk themes, and the default
# encoding.  Everything else (locales, demos, extra encodings, old http
# packages, opt packages) is dead weight.
# ---------------------------------------------------------------------------
_TCL_TM_STRIP = {'http', 'tcltest', 'msgcat', 'platform'}

def _keep_tcl(dst_path: str, src_path: str) -> bool:
    """Return True if this Tcl/Tk data file should be kept."""
    parts = dst_path.replace('\\', '/').lower().split('/')

    # --- _tcl_data (core Tcl library) ---
    # Keep everything EXCEPT tzdata (timezone db) and unused .tm packages.
    # The full Tcl library is required for initialisation — do not strip
    # individual .tcl files or tcl_findLibrary / other core procs will fail.
    if '_tcl_data' in parts:
        if 'tzdata' in parts:
            return False
        fname = os.path.basename(dst_path)
        if fname.endswith('.tm'):
            stem = fname.split('-')[0].lower()
            if stem in _TCL_TM_STRIP:
                return False
        return True

    # --- _tk_data (Tk widget library) ---
    # Keep ttk themes (needed for dark theme) and core widget scripts.
    if '_tk_data' in parts:
        if 'ttk' in parts:
            if 'aquatheme.tcl' in dst_path.lower():
                return False   # macOS only
            return True
        fname = os.path.basename(dst_path)
        return fname in {
            'tk.tcl', 'button.tcl', 'entry.tcl', 'scale.tcl',
            'scrollbar.tcl', 'listbox.tcl', 'dialog.tcl',
            'msgbox.tcl', 'panedwindow.tcl', 'tearoff.tcl', 'menu.tcl',
        }

    # --- tcl8 .tm packages ---
    if 'tcl8' in parts:
        fname = os.path.basename(dst_path)
        if fname.endswith('.tm'):
            stem = fname.split('-')[0].lower()
            if stem in _TCL_TM_STRIP:
                return False
        return True

    return True

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
