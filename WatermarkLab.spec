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
from PyInstaller.utils.hooks import collect_all

block_cipher = None

# Collect tkinterdnd2 (pure-python wrapper + native tkdnd library) so the
# bundled exe supports drag-and-drop. collect_all returns (datas, binaries,
# hiddenimports) tuples we merge into the Analysis below.
_dnd_datas, _dnd_binaries, _dnd_hiddenimports = collect_all('tkinterdnd2')

a = Analysis(
    ['Watermark_Lab.pyw'],
    pathex=[SPECPATH],
    binaries=_dnd_binaries,
    datas=[
        ('SplashLab.png',  '.'),
        ('Watermark.png',  '.'),
        ('Watermark.ico',  '.'),
    ] + _dnd_datas,
    hiddenimports=[
        'packaging', 'packaging.version',
        '_ffmpeg', '_updater', '_version', '_word', '_prefs',
        'docx', 'docx.oxml.ns', 'lxml', 'lxml.etree',
        'tkinterdnd2',
        'PIL._imaging',
        'PIL._imagingft',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
        'PIL.PngImagePlugin',
    ] + _dnd_hiddenimports,
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
    """Return True if this Tcl/Tk data file should be kept.

    Keep the full _tcl_data and _tk_data trees — individual .tcl files
    have deep internal dependencies that are not safe to whitelist.
    Only strip: Tcl tzdata (timezone db) and unused .tm packages.
    """
    parts = dst_path.replace('\\', '/').lower().split('/')

    # Drop Tcl timezone database — continent folders like America/, Europe/
    if 'tzdata' in parts:
        return False

    # Drop specific unused .tm packages
    fname = os.path.basename(dst_path)
    if fname.endswith('.tm'):
        stem = fname.split('-')[0].lower()
        if stem in _TCL_TM_STRIP:
            return False

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
