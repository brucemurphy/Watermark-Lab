# Watermark Lab

> **Version 1.2.0** — First official release  
> A lightweight Windows desktop tool for applying professional, tiled, diagonal watermarks to PowerPoint presentations and video files.

---

## Overview

Watermark Lab lets you stamp any `.pptx` or video file with a customisable diagonal text watermark in seconds. It runs entirely on your Windows machine — no cloud, no subscription, no data leaves your device.

- PowerPoint files are watermarked natively via COM — the result is a real editable `.pptx`, not a flattened image.
- Video files are processed with ffmpeg, which the app downloads automatically on first use.
- The app ships as a self-contained portable folder — unzip and run, no installer needed.

---

## Features

| Feature | Detail |
|---|---|
| **PowerPoint watermarking** | `.pptx` / `.ppt` — tiled diagonal text shape added to every slide. Fully editable output. |
| **Word watermarking** | `.docx` / `.doc` — native diagonal watermark via Word's own VML format. Identical to Design → Watermark in the ribbon. |
| **Video watermarking** | `.mp4` `.mov` `.m4v` `.mkv` `.avi` `.webm` — PNG overlay composited via ffmpeg. Audio stream-copied, no quality loss. |
| **Custom watermark text** | Any text — default is `CONFIDENTIAL`. |
| **Color picker** | Full RGB color chooser with live swatch preview. |
| **Transparency slider** | 0% (fully opaque) → 100% (invisible). |
| **PDF export** | Optionally export a PDF alongside the watermarked PowerPoint. |
| **Open file after save** | Checkbox to auto-open the output file when done (on by default). |
| **Folder shortcut** | 📂 icon in the status bar — click to open the output folder in Explorer. |
| **Splash screen** | Branded splash shown at launch. |
| **Dark theme** | Full dark UI throughout — title bar, controls, status bar. |
| **Auto-updater** | Checks GitHub Releases on launch; one-click install of new versions. |
| **ffmpeg auto-download** | Downloads ffmpeg automatically on first video use (~30 MB). No manual setup. |
| **Portable** | Runs from any user-writable folder — Desktop, USB stick, network share. |

---

## Requirements

- **Windows 10 or 11** (64-bit)
- **Microsoft PowerPoint** — required for `.pptx` / `.ppt` watermarking (uses COM automation). Not needed for video.
- No Python install required when running the pre-built `.exe` from a release.

### Running from source

```powershell
pip install pywin32 Pillow packaging
python Watermark_Lab.pyw
```

Python 3.10 or later required.

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
pip install pywin32 Pillow packaging
python Watermark_Lab.pyw
```

---

## Usage

1. Click **Browse…** and select a `.pptx`, `.ppt`, or video file.
2. Enter your **watermark text** (default: `CONFIDENTIAL`).
3. Choose a **color** and set **transparency** with the slider.
4. Optionally check **Also export PDF** (PowerPoint only).
5. Optionally uncheck **Open file after watermarking** if you don't want it to auto-open.
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

## Project Layout

| File | Purpose |
|---|---|
| `Watermark_Lab.pyw` | Main GUI entry point (Tkinter) |
| `_powerpoint.py` | PowerPoint COM watermarking engine |
| `_video.py` | Video watermarking via ffmpeg + Pillow |
| `_ffmpeg.py` | ffmpeg auto-download and path resolution |
| `_updater.py` | GitHub Releases auto-update logic |
| `_version.py` | Version constant — stamped by CI at build time |
| `version.json` | Release metadata and in-app update notes |
| `WatermarkLab.spec` | PyInstaller build spec |
| `SplashLab.png` | Splash screen image |
| `Watermark.ico` / `Watermark.png` | App icon |
| `.github/workflows/release.yml` | CI: build, zip, publish GitHub Release on tag push |

---

## Building from Source

The release workflow builds automatically via GitHub Actions on every version tag push. To build locally:

```powershell
pip install pyinstaller pywin32 Pillow packaging
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
| **Pillow** | HPND / MIT-CMU | Used for watermark image generation. https://python-pillow.org |
| **pywin32** | PSF License | Used for PowerPoint COM automation. https://github.com/mhammond/pywin32 |
| **packaging** | Apache 2.0 / BSD | Used for version comparison. https://github.com/pypa/packaging |

---

## License

Copyright © 2026 Bruce Murphy.  
Released under the [MIT License](LICENSE).

---

## Release History

See [Releases](https://github.com/brucemurphy/Watermark-Lab/releases) on GitHub.
