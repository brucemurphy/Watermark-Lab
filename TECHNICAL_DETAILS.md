# Watermark Lab — Technical Details

> Engineering reference for Watermark Lab. For the user-facing guide, see [README.md](README.md).

**Current version:** 2.1.0
**Platform:** Windows 10 / 11 (64-bit)
**Language / UI:** Python 3.10+ with PySide6 (Qt for Python)

---

## Architecture overview

Watermark Lab is a single-window PySide6/Qt desktop app built around a three-pane layout (file selection + presets / watermark text + live preview + colour-transparency-preset / actions + status). The UI is one entry point (`Watermark_Lab.pyw`) that orchestrates a set of focused backend modules — one per file format — plus Qt-specific helpers for preview and OneDrive-safe saving.

Design principles:

- **Native, editable output.** Office and PDF files are watermarked at the document level, not flattened to images. A PowerPoint stays a real editable PowerPoint.
- **True-render preview.** The preview rasterises the *actual* watermarked output (real backend → PDF → Qt `QtPdf`) rather than approximating it, so what you see is what you get. A fast live-composite path handles slider/colour/text changes without re-running the backend.
- **Everything local.** No telemetry, no cloud. The only network calls are the optional one-time ffmpeg download and the GitHub Releases update check.
- **Portable by design.** Onedir build, no installer, no registry, no admin rights. Preferences are stored as JSON alongside the app.

---

## Watermarking engines

| Format | Engine | Approach |
|---|---|---|
| **PowerPoint** (`.pptx` / `.ppt`) | `_powerpoint.py` + `_xpowerpoint.py` | Tiled diagonal text shape added to every slide via PowerPoint COM automation. `_xpowerpoint` wraps the save in a temp-then-copy step so SaveAs doesn't silently fail on OneDrive-synced folders. |
| **Word** (`.docx` / `.doc`) | `_word.py` + `_xword.py` | Native diagonal watermark injected as Word's own VML header XML (`python-docx` + `lxml`) — modern `.docx` needs no Word install. Legacy `.doc` is first converted to `.docx` via Word COM; PDF export also uses Word COM. `_xword` handles snug, auto-wrapped text sizing. |
| **PDF** (`.pdf`) | `_pdf.py` | Pillow renders the tiled diagonal text to an image with an explicit soft mask (`/SMask`); `pypdf` overlays it on every page. The original page content is preserved (no re-rasterising) and shows through the semi-transparent text. |
| **Video** (`.mp4` `.mov` `.m4v` `.mkv` `.avi` `.webm`) | `_video.py` + `_ffmpeg.py` | Pillow generates a tiled PNG overlay; ffmpeg composites it over the video. Audio is stream-copied (no re-encode, no quality loss). |

The default watermark style is Segoe UI, gray (`#A6A6A6`), 70% transparency, tiled at −30°.

---

## Live preview

`_xpreview.py` drives the preview:

- **Background load (COM-light).** `.pptx` uses the embedded thumbnail (no COM); `.doc/.docx` and legacy `.ppt` make one short Office export trip; PDF rasterises its first page directly via Qt `QtPdf`; video pulls the first frame via ffmpeg.
- **Live composite.** Once a clean backdrop is cached, watermark text/colour/transparency changes are composited in the UI thread instantly — no backend, no COM.
- **Threading & cancellation.** Heavy work runs on a `QThreadPool` with a monotonic token, so stale results are discarded and turning the preview off cancels in-flight work at the next checkpoint.
- **Protected files.** Encrypted / sensitivity-labelled (MIP) Office files are detected before launching Office and reported clearly instead of failing mid-render.
- **Render notice.** Because Word/PowerPoint backdrops need a full high-resolution render, the preview shows a heads-up message while that first render is generated.
- **Zoom control.** The preview canvas zooms from 10% to 800% (`_MIN_ZOOM` / `_MAX_ZOOM`). The toolbar percentage is an editable combo: the +/- buttons step at 1.25×, preset items jump to common levels, and a typed value is parsed and clamped to range. A fixed width keeps three-digit readouts (up to `800%`) from clipping.

---

## Status feedback

The Status panel reflects job state live:

- **Working** — an animated gold spinner (`SpinnerWidget`, a `QPropertyAnimation`-driven arc).
- **Success** — a green tick.
- **Failure** — a red cross (`CrossIcon`) on a red-tinted badge, plus a plain-language error. Backend/COM errors are translated by `_friendly_error()` (e.g. "file is open in another program", "file is password-protected", "protected or restricted for editing").

The red "Failed" state is shown only when every file in a run fails; mixed batches report "Completed with errors" with the per-file detail surfaced in a dialog.

---

## Requirements (from source)

- **Windows 10 or 11** (64-bit)
- **Python 3.10+**
- **Microsoft PowerPoint** — for `.pptx` / `.ppt` watermarking (COM automation). Not needed for Word, PDF, or video.
- **Microsoft Word** — for legacy `.doc` watermarking (converted to `.docx` via COM) and for `.docx` / `.doc` PDF export. Modern `.docx` watermarking uses `python-docx` only and does **not** require Word.

```powershell
git clone https://github.com/brucemurphy/Watermark-Lab.git
cd Watermark-Lab
pip install PySide6 pywin32 Pillow packaging python-docx pypdf
python Watermark_Lab.pyw
```

The live preview uses **PySide6** (Qt for Python), including its bundled `QtPdf` module.

---

## Project layout

| File | Purpose |
|---|---|
| `Watermark_Lab.pyw` | Main GUI entry point (PySide6/Qt) |
| `_xpreview.py` | True-render live preview engine (Qt `QtPdf` + ffmpeg) |
| `_xpowerpoint.py` | OneDrive-safe PowerPoint save wrapper |
| `_xword.py` | Snug, auto-wrapped Word watermark |
| `_powerpoint.py` | PowerPoint COM watermarking engine (shared) |
| `_word.py` | Word VML watermarking engine (shared) |
| `_pdf.py` | PDF watermarking via Pillow + pypdf (shared) |
| `_video.py` | Video watermarking via ffmpeg + Pillow (shared) |
| `_ffmpeg.py` | ffmpeg auto-download and path resolution |
| `_prefs.py` | Presets, recent files, and last-used folder storage |
| `_updater.py` | GitHub Releases auto-update logic |
| `_version.py` | Version constant — stamped by CI at build time |
| `version.json` | Release metadata and in-app update notes |
| `WatermarkLab.spec` | PyInstaller build spec |
| `SplashLab.png` | Splash screen image |
| `Watermark.ico` / `Watermark.png` | App icon |
| `.github/workflows/release.yml` | CI: build, zip, publish GitHub Release on tag push |

---

## ffmpeg auto-download

Video watermarking requires ffmpeg, but nothing is bundled or installed manually. On first video use the app fetches the latest minimal essentials build from [GyanD/codexffmpeg](https://github.com/GyanD/codexffmpeg) (~30 MB) and caches it beside the app:

```
WatermarkLab\ffmpeg.exe
```

Subsequent launches find it instantly with no network call. Copying the whole `WatermarkLab` folder carries ffmpeg with it.

> **FFmpeg copyright notice**
> FFmpeg is © the FFmpeg developers and other contributors.
> Licensed under the [GNU Lesser General Public License (LGPL) v2.1+](https://ffmpeg.org/legal.html) or the GNU GPL v2+ depending on the build configuration.
> Watermark Lab does **not** redistribute FFmpeg binaries. FFmpeg is downloaded separately by the user at runtime.
> Source code: https://ffmpeg.org/

---

## Auto-update

About three seconds after launch, a `_UpdateChecker` `QObject` runs on a `QThread`, reuses the `_updater` logic (`_fetch_latest_release` / `_is_newer`), and emits `updateAvailable(version, notes)` to a `QMessageBox` prompt. On **Yes**:

1. The new `WatermarkLab.zip` is downloaded from GitHub's CDN.
2. A background PowerShell script copies the updated files over the existing app folder.
3. The updated app relaunches automatically.

`cleanup_old_exe()` runs at startup to remove leftovers from a previous in-place update. Update checks are frozen-only and silent on failure — a no-op when running from source.

---

## Building from source

The release workflow builds automatically via GitHub Actions on every version tag push (`v[0-9]+.[0-9]+.[0-9]+`). To build locally:

```powershell
pip install pyinstaller PySide6 pywin32 Pillow packaging python-docx pypdf
pyinstaller WatermarkLab.spec
```

Output is in `dist\WatermarkLab\`. Zip it for distribution:

```powershell
Compress-Archive -Path dist\WatermarkLab -DestinationPath WatermarkLab.zip
```

### Why onedir, not onefile?

Onefile mode extracts to `%TEMP%` at every launch. This causes crashes on OneDrive-synced folders (the sync engine locks the extracted files) and is blocked by WDAC / App Control policies on managed Windows machines. Onedir has no runtime extraction — all files are in place from build time, so neither problem applies. The spec also strips unused binaries (select Pillow extensions, win32ui/MFC, Tcl/Tk data) to keep the build lean.

---

## Third-party notices

This repository does **not** redistribute any third-party binaries.

| Component | License | Notes |
|---|---|---|
| **FFmpeg** | LGPL v2.1+ / GPL v2+ | Downloaded at runtime by the user. © FFmpeg developers. https://ffmpeg.org |
| **PySide6 (Qt for Python)** | LGPL v3 / GPL | Qt UI and `QtPdf` preview rendering. © The Qt Company. https://www.qt.io |
| **Pillow** | HPND / MIT-CMU | Watermark image generation. https://python-pillow.org |
| **pywin32** | PSF License | PowerPoint / Word COM automation. https://github.com/mhammond/pywin32 |
| **python-docx** | MIT | Word VML watermark injection. https://github.com/python-openxml/python-docx |
| **pypdf** | BSD | Composites watermarks onto existing PDFs. https://github.com/py-pdf/pypdf |
| **packaging** | Apache 2.0 / BSD | Version comparison. https://github.com/pypa/packaging |

---

## License

Copyright © 2026 Bruce Murphy.
Released under the [MIT License](LICENSE).

---

## Release history

See [Releases](https://github.com/brucemurphy/Watermark-Lab/releases) on GitHub.
