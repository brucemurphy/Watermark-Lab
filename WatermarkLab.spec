# PyInstaller spec for Watermark Lab
#
# Build:   pyinstaller WatermarkLab.spec
# Output:  dist\WatermarkLab.exe   (single, self-contained, portable)
#
# How it stays portable:
#   This is a one-file build (`exclude_binaries` is not set, all binaries are
#   embedded in the EXE). At launch the PyInstaller bootloader unpacks the
#   payload into a temporary `_MEI<pid>` folder. `runtime_tmpdir='.'` directs
#   that extraction to a subdirectory **next to the executable** instead of
#   `%TEMP%`. This lets the exe run from any location (USB stick, Desktop,
#   network share) and also works on machines running Windows Application
#   Control / WDAC, which can block DLL loads from `%TEMP%`.
#
# Requirements before building:
#   - pip install pyinstaller pywin32 Pillow
#   - ffmpeg.exe and ffprobe.exe must be present in this folder
#     (see README -> "FFmpeg setup")

block_cipher = None

a = Analysis(
    ['Watermark_Lab.pyw'],
    pathex=[],
    binaries=[
        ('ffmpeg.exe',  '.'),
        ('ffprobe.exe', '.'),
    ],
    datas=[
        ('SplashLab.png',  '.'),
        ('Watermark.png',  '.'),
        ('Watermark.ico',  '.'),
    ],
    hiddenimports=['packaging', 'packaging.version'],
    hookspath=[],
    hooksconfig={},
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
