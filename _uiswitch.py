"""UI switching for Watermark Lab (source runs).

Lets the user flip between the modern PySide6/Qt UI and the legacy Tkinter UI.
Because Qt and Tkinter each own a blocking event loop, switching is done by
relaunching the chosen front-end as a fresh process and exiting the current
one. The choice is remembered via _prefs (settings.json), so the next cold
start reopens in the last-used UI.

This module only orchestrates *which front-end process runs*. Both UIs share
the exact same backend engines (_powerpoint, _word, _video, _ffmpeg, ...), so
nothing here touches watermarking logic.

The relaunch passes a marker flag (FORCED_FLAG) so the spawned process opens
the requested UI directly and never re-dispatches — preventing any loop.
"""
from __future__ import annotations

import os
import subprocess
import sys

import _prefs

# Passed to the spawned process so it opens the requested UI without consulting
# the saved preference again (avoids a dispatch loop on cold start).
FORCED_FLAG = "--ui-forced"

MODE_MODERN = "modern"
MODE_CLASSIC = "classic"

# Repo root = folder containing this module (and the modern entry point).
ROOT = os.path.dirname(os.path.abspath(__file__))

# Entry points. The modern Qt UI lives in the root; the legacy Tk UI was
# archived under previous_version/ during the 2.0.0 switch.
MODERN_ENTRY = os.path.join(ROOT, "Watermark_Lab.pyw")
CLASSIC_ENTRY = os.path.join(ROOT, "previous_version", "Watermark_Lab.pyw")

_ENTRIES = {MODE_MODERN: MODERN_ENTRY, MODE_CLASSIC: CLASSIC_ENTRY}

# Shared brand assets — both UIs resolve these from the repo root so the paths
# never depend on where each entry file happens to live (the classic UI was
# moved into previous_version/ but the assets stay in the root).
ICON_ICO = os.path.join(ROOT, "Watermark.ico")
ICON_PNG = os.path.join(ROOT, "Watermark.png")
SPLASH_PNG = os.path.join(ROOT, "SplashLab.png")

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def is_forced() -> bool:
	"""True if this process was launched with the relaunch marker flag."""
	return FORCED_FLAG in sys.argv


def _python_for_gui() -> str:
	"""Prefer pythonw.exe (no console) over python.exe when running from source.

	Frozen builds run from the exe itself; there we just reuse sys.executable.
	"""
	exe = sys.executable
	if getattr(sys, "frozen", False):
		return exe
	base = os.path.dirname(exe)
	pythonw = os.path.join(base, "pythonw.exe")
	return pythonw if os.path.isfile(pythonw) else exe


def relaunch(mode: str) -> bool:
	"""Persist the chosen UI mode and spawn that front-end as a new process.

	Returns True if the new process was launched (the caller should then exit
	its own event loop), False if the target entry file is missing.
	"""
	if mode not in _ENTRIES:
		return False
	entry = _ENTRIES[mode]
	if not os.path.isfile(entry):
		return False

	_prefs.save_ui_mode(mode)

	# Ensure the child can import the shared root modules regardless of cwd.
	env = dict(os.environ)
	existing = env.get("PYTHONPATH", "")
	env["PYTHONPATH"] = ROOT + (os.pathsep + existing if existing else "")

	cmd = [_python_for_gui(), entry, FORCED_FLAG]
	try:
		subprocess.Popen(cmd, cwd=ROOT, env=env, creationflags=_NO_WINDOW)
		return True
	except Exception:
		return False
