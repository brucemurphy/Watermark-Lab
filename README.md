# Watermark Lab

> **Version 2.0.2** — PySide6/Qt rewrite with true-render live preview  
> A lightweight Windows desktop tool for applying professional, tiled, diagonal watermarks to PowerPoint, Word and video files.

---

## Overview

Watermark Lab lets you stamp any PowerPoint, Word or video file with a customisable diagonal text watermark in seconds. It runs entirely on your Windows machine — no cloud, no subscription, no data leaves your device.

- PowerPoint and Word files are watermarked natively — the result is a real, editable document, not a flattened image.
- A **true-render live preview** shows the genuine watermarked output as you type and adjust settings.
- Video files are processed with ffmpeg, which the app downloads automatically on first use.
- The app ships as a self-contained portable folder — unzip and run, no installer needed.

> **New in 2.0.0:** the interface is a full **PySide6/Qt** rewrite — a three-pane layout with drag-and-drop, an editable transparency value, a colour palette + eyedropper, and a faithful live preview. The previous Tkinter UI is preserved under [`previous_version/`](previous_version/README.md).
>
> **New in 2.0.1:** switch between the new and classic interface from inside the app (your choice is remembered), browse dialogs reopen at your last folder, the window comes to the front on launch/switch, and a new **Export PDF Only** button leaves just the PDF behind.

---

## Features

| Feature | Detail |
|---|---|
| **True-render live preview** | A faithful, pixel-accurate preview of the actual watermarked output — rendered via Qt's PDF engine — updates as you edit text, colour and transparency. |
| **PowerPoint watermarking** | `.pptx` / `.ppt` — tiled diagonal text shape added to every slide. Fully editable output. OneDrive-safe saving. |
| **Word watermarking** | `.docx` / `.doc` — native diagonal watermark via Word's own VML format. Snug, auto-wrapped text. |
| **Video watermarking** | `.mp4` `.mov` `.m4v` `.mkv` `.avi` `.webm` — PNG overlay composited via ffmpeg. Audio stream-copied, no quality loss. |
| **Drag & drop** | Drop a file straight onto the app, or browse for a file or a whole folder. |
| **Custom watermark text** | Any text — default is `CONFIDENTIAL`. Wraps automatically; up to 100 characters. |
| **Colour picker** | Palette swatches, an editable hex value, and an eyedropper for any custom colour. |
| **Transparency** | Slider plus an **editable value** — type an exact 0–100 percentage. |
| **Switchable interface** | Flip between the modern and classic interface from inside the app — your choice is remembered for next launch. |
| **Remembered browse folder** | File / folder dialogs reopen at your last-used folder (Documents by default). |
| **Export PDF Only** | One-click export of just the PDF — the intermediate watermarked Office file is cleaned up for you. |
| **Batch processing** | Pick a folder and watermark every supported file in one click. |
| **Presets** | Save and recall named configurations (text + colour + transparency). |
| **Recent files** | Quick access to recently used files and presets. |
| **PDF export** | Optionally export a PDF alongside the watermarked PowerPoint / Word file. |
| **Open file after save** | Toggle to auto-open the output file(s) when done. |
| **Dark theme** | Full dark Qt UI throughout — custom title bar, controls, status panel. |
| **Auto-updater** | Checks GitHub Releases on launch; one-click install of new versions. |
| **ffmpeg auto-download** | Downloads ffmpeg automatically on first video use (~30 MB). No manual setup. |
| **Portable** | Runs from any user-writable folder — Desktop, USB stick, network share. |

---

## Requirements

- **Windows 10 or 11** (64-bit)
- **Microsoft PowerPoint** — required for `.pptx` / `.ppt` watermarking (uses COM automation). Not needed for Word or video.
- **Microsoft Word** — required for `.docx` / `.doc` PDF export. Not needed for the watermark itself.
- No Python install required when running the pre-built `.exe` from a release.

### Running from source

```powershell
pip install PySide6 pywin32 Pillow packaging python-docx
python Watermark_Lab.pyw
```

Python 3.10 or later required. The live preview uses **PySide6** (Qt for Python), including its bundled `QtPdf` module.

---

## Getting Started

### Option A — Pre-built release (recommended)

1. Download `WatermarkLab.zip` from the [latest release](https://github.com/brucemurphy/Watermark-Lab/releases/latest).
2. Extract the zip to any folder (Desktop, USB stick, etc.).
3. Double-click `WatermarkLab.exe`.

No installer, no admin rights needed.

### Option B — Run from source

```powershell
git clone https://github.com/brucemurphy/Watermark-Lab.git
cd Watermark-Lab
pip install PySide6 pywin32 Pillow packaging python-docx
python Watermark_Lab.pyw
```

---

## Usage

1. **Drag a file onto the app**, or click **Browse File…** / **Browse Folder…** to select a `.pptx`, `.ppt`, `.docx`, `.doc`, or video file.
2. Enter your **watermark text** (default: `CONFIDENTIAL`).
3. Pick a **colour** (palette, hex, or eyedropper) and set **transparency** — drag the slider or type an exact value in the editable pill.
4. Watch the **live preview** update to show the real watermarked result.
5. Toggle **Export PDF** (PowerPoint / Word) and **Open file(s) after processing** as needed.
6. Click **Apply Watermark**.

The output file is saved next to the source with a `_watermarked` suffix:

```
my_presentation.pptx  →  my_presentation_watermarked.pptx
recording.mp4         →  recording_watermarked.mp4
```

When complete, the 📂 icon appears in the status bar — click it to open the output folder in Explorer.

---

## FFmpeg

Video watermarking requires ffmpeg. **No manual setup is needed.**

On first video use the app prompts you to download ffmpeg automatically. It fetches the latest minimal essentials build from [GyanD/codexffmpeg](https://github.com/GyanD/codexffmpeg) (~30 MB) and saves it alongside the app:

```
WatermarkLab\ffmpeg.exe
```

ffmpeg is cached — subsequent launches find it instantly with no network call. Copy the whole `WatermarkLab` folder to another machine and ffmpeg comes with it.

> **FFmpeg copyright notice**  
> FFmpeg is © the FFmpeg developers and other contributors.  
> Licensed under the [GNU Lesser General Public License (LGPL) v2.1+](https://ffmpeg.org/legal.html) or the GNU GPL v2+ depending on the build configuration.  
> Watermark Lab does **not** redistribute FFmpeg binaries. FFmpeg is downloaded separately by the user at runtime.  
> Source code: https://ffmpeg.org/

---

## Auto-Update

Three seconds after launch the app silently checks the GitHub Releases API for a newer version. If one is found you'll see a prompt:

> **Watermark Lab x.y.z is available.**  
> What's new: …  
> Install now and restart?

On **Yes**:
1. The new `WatermarkLab.zip` is downloaded from GitHub's CDN.
2. The updated files are copied over the existing app folder via a background PowerShell script.
3. The updated app launches automatically.

Update checks are a no-op when running from Python source (`python Watermark_Lab.pyw`).

---

## Upgrading from 1.4.x

The upgrade path is automatic — **no reinstall required**.

- **Pre-built app users:** launch your existing 1.4.x app. A few seconds after start it checks GitHub Releases, sees that **2.0.0 > 1.4.x**, and prompts you to install. Click **Yes** and the new Qt app downloads, replaces the old folder, and relaunches. Your presets and recent files are preserved.
- **Running from source:** pull the latest and install the one new dependency:

  ```powershell
  git pull
  pip install PySide6
  python Watermark_Lab.pyw
  ```

- **Prefer the old UI?** The legacy Tkinter front-end is archived under [`previous_version/`](previous_version/README.md) and still works against the shared backend modules in the repository root.

---

## Project Layout

| File | Purpose |
|---|---|
| `Watermark_Lab.pyw` | Main GUI entry point (PySide6/Qt) |
| `_uiswitch.py` | Classic/Modern UI switch + shared asset paths |
| `_xpreview.py` | True-render live preview engine (Qt `QtPdf` + ffmpeg) |
| `_xpowerpoint.py` | OneDrive-safe PowerPoint save wrapper |
| `_xword.py` | Snug, auto-wrapped Word watermark |
| `_powerpoint.py` | PowerPoint COM watermarking engine (shared) |
| `_word.py` | Word VML watermarking engine (shared) |
| `_video.py` | Video watermarking via ffmpeg + Pillow (shared) |
| `_ffmpeg.py` | ffmpeg auto-download and path resolution |
| `_updater.py` | GitHub Releases auto-update logic |
| `_version.py` | Version constant — stamped by CI at build time |
| `version.json` | Release metadata and in-app update notes |
| `WatermarkLab.spec` | PyInstaller build spec |
| `SplashLab.png` | Splash screen image |
| `Watermark.ico` / `Watermark.png` | App icon |
| `previous_version/` | Archived legacy 1.4.x Tkinter app |
| `.github/workflows/release.yml` | CI: build, zip, publish GitHub Release on tag push |

---

## Building from Source

The release workflow builds automatically via GitHub Actions on every version tag push. To build locally:

```powershell
pip install pyinstaller PySide6 pywin32 Pillow packaging python-docx
pyinstaller WatermarkLab.spec
```

Output is in `dist\WatermarkLab\`. Zip it for distribution:

```powershell
Compress-Archive -Path dist\WatermarkLab -DestinationPath WatermarkLab.zip
```

> **Why a folder (onedir) and not a single exe (onefile)?**  
> Onefile mode extracts to `%TEMP%` at every launch. This causes crashes on OneDrive-synced folders (sync engine locks the files) and is blocked by WDAC/App Control policies on managed Windows machines. Onedir has no runtime extraction — all files are in place from the start.

---

## Third-Party Notices

This repository does **not** redistribute any third-party binaries.

| Component | License | Notes |
|---|---|---|
| **FFmpeg** | LGPL v2.1+ / GPL v2+ | Downloaded at runtime by the user. © FFmpeg developers. https://ffmpeg.org |
| **PySide6 (Qt for Python)** | LGPL v3 / GPL | Qt UI and `QtPdf` preview rendering. © The Qt Company. https://www.qt.io |
| **Pillow** | HPND / MIT-CMU | Used for watermark image generation. https://python-pillow.org |
| **pywin32** | PSF License | Used for PowerPoint / Word COM automation. https://github.com/mhammond/pywin32 |
| **python-docx** | MIT | Used for Word VML watermark injection. https://github.com/python-openxml/python-docx |
| **packaging** | Apache 2.0 / BSD | Used for version comparison. https://github.com/pypa/packaging |

---

## License

Copyright © 2026 Bruce Murphy.  
Released under the [MIT License](LICENSE).

---

## Release History

See [Releases](https://github.com/brucemurphy/Watermark-Lab/releases) on GitHub.
