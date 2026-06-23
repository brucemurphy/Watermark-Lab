# Watermark Lab — Legacy 1.4.x (Tkinter)

> **Archived for reference.** This folder contains the original Tkinter
> front-end that shipped through **v1.4.2**. It has been superseded by the
> PySide6/Qt application in the repository root (**v2.0.0+**).

This file is kept so the previous user experience remains available and
auditable. It is **not** built or published by the release pipeline.

---

## What this is

`Watermark_Lab.pyw` (in this folder) is the legacy Tkinter GUI. It provides the
same core watermarking abilities that 1.4.x users know:

- PowerPoint (`.pptx` / `.ppt`), Word (`.docx` / `.doc`) and video watermarking
- Custom text, RGB colour picker, transparency slider
- Batch folder processing, saved presets, recent files
- Optional PDF export, open-after-save, ffmpeg auto-download, auto-update

The 2.0.0 Qt app keeps all of these and adds a true-render live preview,
drag-and-drop, an editable transparency value, a colour palette + eyedropper,
and OneDrive-safe PowerPoint saving. See the [root README](../README.md).

---

## Dependency note (important)

This legacy entry point imports the **shared backend modules that still live in
the repository root**, by bare module name:

```
_powerpoint.py  _word.py  _video.py
_prefs.py       _ffmpeg.py  _updater.py  _version.py
```

Those modules were intentionally **left in the root** (the 2.0.0 app uses them
too) and were **not** copied into this folder to avoid divergence. As a result,
this file cannot run standalone from inside `previous_version/`.

### Running the legacy app from source

Run it from the **repository root** so the shared modules resolve, pointing
Python at this archived file:

```powershell
pip install pywin32 Pillow packaging python-docx
python previous_version/Watermark_Lab.pyw
```

Python 3.10 or later is required.

---

## Why it was replaced

Version 2.0.0 is a full PySide6/Qt rewrite. The Tkinter UI is retained here only
as a historical reference; new development happens in the root application.
