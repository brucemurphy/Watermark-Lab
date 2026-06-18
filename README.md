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
pip install pywin32 Pillow packaging
```

---

## FFmpeg setup

No manual setup required. On the **first time you watermark a video**, the app will prompt you to download ffmpeg automatically. It fetches the latest minimal build from [GyanD/codexffmpeg](https://github.com/GyanD/codexffmpeg/releases) (~30 MB) and saves it to:

```
%LOCALAPPDATA%\WatermarkLab\ffmpeg.exe
```

This cache persists across app updates so the download only happens once. PowerPoint watermarking works without ffmpeg at all.

---

## Project layout

| File | Purpose |
|---|---|
| [Watermark_Lab.pyw](Watermark_Lab.pyw) | Tkinter GUI entry point |
| [_powerpoint.py](_powerpoint.py) | PowerPoint COM watermarking |
| [_video.py](_video.py) | Video watermarking via ffmpeg + Pillow |
| [_ffmpeg.py](_ffmpeg.py) | ffmpeg cache + on-demand download from GitHub |
| [_version.py](_version.py) | App version constant (stamped by CI) |
| [_updater.py](_updater.py) | Auto-update via GitHub Releases API |
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
   pip install pyinstaller pywin32 Pillow packaging
   ```

3. Build the executable:

   ```powershell
   pyinstaller WatermarkLab.spec
   ```

4. The portable executable is written to **`dist\WatermarkLab.exe`**. Copy that single file wherever you want to run it.

### What gets embedded

- Application code: `Watermark_Lab.pyw`, `_powerpoint.py`, `_video.py`, `_ffmpeg.py`, `_version.py`, `_updater.py`
- Image assets: `SplashLab.png`, `Watermark.png`, `Watermark.ico`

ffmpeg is **not embedded** — it is downloaded on first video use and cached in `%LOCALAPPDATA%\WatermarkLab\`.

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

## Release & Auto-Update

Watermark Lab uses **GitHub Releases** as the single source of truth for both distribution and updates. No external services or secrets required.

### How releases are published

1. Push a version tag:
   ```powershell
   git tag v1.2.0
   git push origin v1.2.0
   ```
2. The [release workflow](.github/workflows/release.yml) runs automatically:
   - Stamps `_version.py` with the tag version.
   - Builds `WatermarkLab.exe` via PyInstaller.
   - Creates a GitHub Release and attaches `WatermarkLab.exe` as a release asset.

### How the in-app update works

3 seconds after launch the app silently calls the GitHub Releases API:
```
https://api.github.com/repos/brucemurphy/Watermark-Lab/releases/latest
```
If the `tag_name` version is newer than the running build, the user is prompted. On confirmation:

1. Streams the new `WatermarkLab.exe` from GitHub's CDN to `WatermarkLab.exe.tmp`.
2. Renames the running exe to `WatermarkLab.exe.old`.
3. Moves `.tmp` to `WatermarkLab.exe`.
4. Launches the new exe and exits.

The `.old` file is cleaned up on the next startup. Update checks are a no-op when running from Python source.

### No secrets or external services needed

GitHub Releases are public. No Azure storage account, no manifest file, no API keys.

---

## License

The source code in this repository is © its respective authors. See [`LICENSE`](LICENSE) for terms.
