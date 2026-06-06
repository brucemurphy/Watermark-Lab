# Watermark Lab

A lightweight Windows desktop tool for applying tiled, diagonal, semi-transparent watermarks to **PowerPoint presentations** and **video files**.

It ships as a dark-themed Tkinter GUI, drives PowerPoint via COM for native slide edits, and uses **ffmpeg** with a Pillow-generated overlay for video.

---

## Features

- **PowerPoint watermarking** (`.pptx`, `.ppt`) — tiled text added as a real text shape on every slide, rotated -30°, with adjustable color and transparency. Saved as a new `*_watermarked.pptx`.
- **Video watermarking** (`.mp4`, `.mov`, `.m4v`, `.mkv`, `.avi`, `.webm`) — generates a tiled diagonal PNG overlay and re-encodes via ffmpeg (libx264, CRF 20). Audio is stream-copied so quality is preserved. Saved as a new `*_watermarked.<ext>`.
- **Dark-mode GUI** with file picker, color picker, transparency slider, and live status / progress.
- **Background processing** — encoding runs on a worker thread so the UI stays responsive.
- **Auto-open** the result when finished.

---

## Requirements

- **Windows 10 / 11**
- **Python 3.10+**
- **Microsoft PowerPoint** installed locally (required for `.pptx` / `.ppt` — driven through COM)
- **FFmpeg** (`ffmpeg.exe` and `ffprobe.exe`) — required for video watermarking. **Not included** in this repository; see [FFmpeg setup](#ffmpeg-setup) below.

### Python packages

```powershell
pip install pywin32 Pillow
```

---

## FFmpeg setup

FFmpeg is **not bundled** with this repository — you need to download it yourself and place the executables next to `Watermark_Lab.pyw`.

1. Download a Windows **"essentials"** build of FFmpeg. Recommended sources:
   - https://www.gyan.dev/ffmpeg/builds/ (look for `ffmpeg-release-essentials.zip`)
   - https://github.com/BtbN/FFmpeg-Builds/releases (look for a `win64-gpl` build)
2. Extract the archive. Inside you'll find a `bin/` folder containing `ffmpeg.exe`, `ffprobe.exe`, and `ffplay.exe`.
3. Copy **`ffmpeg.exe`** and **`ffprobe.exe`** (and optionally `ffplay.exe`) into the **same folder as `Watermark_Lab.pyw`**.

Alternatively, install FFmpeg system-wide (e.g., `winget install Gyan.FFmpeg`) and make sure `ffmpeg.exe` / `ffprobe.exe` are on your `PATH`. The app checks the script folder first, then falls back to `PATH`.

If neither location has the binaries, video watermarking will fail with a clear error message; PowerPoint watermarking will continue to work without FFmpeg.

---

## Project layout

| File | Purpose |
|---|---|
| [Watermark_Lab.pyw](Watermark_Lab.pyw) | Tkinter GUI entry point |
| [_powerpoint.py](_powerpoint.py) | PowerPoint COM watermarking |
| [_video.py](_video.py) | Video watermarking via ffmpeg + Pillow |
| `ffmpeg.exe` / `ffprobe.exe` | **User-supplied** — see [FFmpeg setup](#ffmpeg-setup) |
| `Watermark.ico` / `Watermark.png` | Window icon |
| `SplashLab.png` | Splash screen shown at launch |
| [WatermarkLab.spec](WatermarkLab.spec) | PyInstaller spec for the portable single-file build |

---

## Building a portable executable

You can package Watermark Lab into a **single self-contained `WatermarkLab.exe`** using [PyInstaller](https://pyinstaller.org/). The resulting executable is fully portable — copy it to a USB stick, network share, or any folder and double-click to run. No Python install required on the target machine.

### How the portable exe works

The build embeds the Python runtime, the application code, the image assets, and the FFmpeg binaries into a single `.exe`. At launch, PyInstaller's bootloader unpacks that payload into a temporary `_MEI<pid>` subfolder **next to the executable** (not in `%TEMP%`), runs the app from there, and cleans the folder up on a normal exit.

Extracting beside the exe (rather than in `%TEMP%`) is what keeps it portable and also lets it run on locked-down corporate machines where **Windows Application Control / WDAC** blocks DLL loads from `%TEMP%`. This is configured in [WatermarkLab.spec](WatermarkLab.spec) via `runtime_tmpdir='.'`.

### Build steps

1. Complete the [FFmpeg setup](#ffmpeg-setup) — `ffmpeg.exe` and `ffprobe.exe` must be sitting next to `WatermarkLab.spec`, because the build embeds them into the output.
2. Install the build tooling:

   ```powershell
   pip install pyinstaller pywin32 Pillow
   ```

3. Build the executable:

   ```powershell
   pyinstaller WatermarkLab.spec
   ```

4. The portable executable is written to **`dist\WatermarkLab.exe`**. Copy that single file wherever you want to run it.

### What gets embedded

- Application code: `Watermark_Lab.pyw`, `_powerpoint.py`, `_video.py`
- Image assets: `SplashLab.png`, `Watermark.png`, `Watermark.ico`
- FFmpeg binaries: `ffmpeg.exe`, `ffprobe.exe` (`ffplay.exe` is intentionally omitted — the app doesn't use it)

The resulting exe is windowed (`console=False`) and uses `Watermark.ico` as its taskbar / file icon.

### Notes

- Place the exe somewhere user-writable (Desktop, Documents, a USB stick). It cannot run from a folder where the user has no write permission, because it needs to create the `_MEI<pid>` extraction folder beside itself.
- The first launch is slightly slower than subsequent launches due to the unpack step.
- PyInstaller writes intermediate files to `build\` and finished output to `dist\`. Both are ignored by [.gitignore](.gitignore). Clean them anytime with:

  ```powershell
  Remove-Item -Recurse -Force build, dist
  ```

---

## Usage

### GUI

Double-click `Watermark_Lab.pyw`, or from a terminal:

```powershell
python Watermark_Lab.pyw
```

1. **Browse…** and pick a `.pptx`, `.ppt`, or video file.
2. Enter the **watermark text** (default: `CONFIDENTIAL`).
3. Pick a **color** and adjust **transparency** (0% = opaque, 100% = invisible).
4. Click **Apply Watermark**. The output file opens automatically when finished.

### Command line

The two backend modules can also be invoked directly:

```powershell
# PowerPoint
python _powerpoint.py "deck.pptx" "CONFIDENTIAL"

# Video
python _video.py "clip.mp4" "CONFIDENTIAL"
```

Both write a sibling file with a `_watermarked` suffix.

---

## How it works

- **PowerPoint** — uses `pywin32` early-bound COM (`gencache.EnsureDispatch`) to add an oversized `TextBox` shape per slide, fills it with a tiled grid of text, rotates it -30°, and applies color + transparency via `Font2.Fill`. PowerPoint's `ForeColor.RGB` uses BGR byte order, so the input `0xRRGGBB` value is byte-swapped before assignment. PowerPoint runs hidden (`WithWindow=False`) and is force-terminated after save to release file locks.
- **Video** — Pillow renders a tiled, rotated RGBA PNG sized to the video frame, then ffmpeg overlays it with `-filter_complex "[0:v][1:v]overlay=0:0"`. Font size auto-scales to roughly `height / 30`. Progress is parsed from ffmpeg's `time=` output and pushed back to the GUI.

---

## Notes & limitations

- PowerPoint integration is **Windows-only** (uses COM).
- Closing the GUI while a job is running will not kill the background ffmpeg / PowerPoint process — let it finish first.
- Output files are always written next to the source with a `_watermarked` suffix; existing files with that name are overwritten.

---

## Third-party software

This repository does **not** redistribute any third-party binaries. The following components are used at runtime and remain the property of their respective authors:

- **FFmpeg** — https://ffmpeg.org/ — user-supplied at runtime (see [FFmpeg setup](#ffmpeg-setup)). Licensed under the LGPL or GPL depending on the build you download. Consult the license bundled with your FFmpeg download for the exact terms.
- **Pillow** — https://python-pillow.org/ — MIT-CMU / HPND license.
- **pywin32** — https://github.com/mhammond/pywin32 — PSF License.

---

## License

The source code in this repository is © its respective authors. See [`LICENSE`](LICENSE) for terms.
