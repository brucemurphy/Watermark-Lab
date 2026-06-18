"""ffmpeg resolution and on-demand download for Watermark Lab.

On first use the app downloads a minimal ffmpeg build from the official
GyanD/codexffmpeg GitHub releases and saves it beside the running exe
(or beside this script when running from source):

	<app folder>\ffmpeg.exe

This keeps the app fully self-contained and portable — copy the folder
to a USB stick or another machine and ffmpeg comes with it.

Subsequent launches find the binary instantly — no network call.

Public API
----------
get_ffmpeg_exe() -> str
	Return the path to a ready ffmpeg binary.
	Raises FfmpegNotReadyError if ffmpeg has not been downloaded yet.

is_ffmpeg_cached() -> bool
	Quick check — no network call.

download_ffmpeg(progress_cb=None) -> str
	Download, extract, and save ffmpeg.exe beside the app.
	progress_cb(bytes_done: int, total: int | None) is called during download.
	Returns the path to the binary.
"""
import io
import json
import os
import shutil
import sys
import urllib.request
import zipfile
from typing import Callable, Optional

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
_FFMPEG_REPO      = "GyanD/codexffmpeg"
_FFMPEG_API_URL   = f"https://api.github.com/repos/{_FFMPEG_REPO}/releases/latest"
_ASSET_SUFFIX     = "essentials_build.zip"
_UA               = "WatermarkLab/ffmpeg-downloader"
_TIMEOUT_API      = 15
_TIMEOUT_DOWNLOAD = 120


def _app_dir() -> str:
	"""Return the folder that contains the running exe (or this script)."""
	if getattr(sys, "frozen", False):
		# PyInstaller: sys.executable is the .exe path
		return os.path.dirname(os.path.abspath(sys.executable))
	# Running from source
	return os.path.dirname(os.path.abspath(__file__))


def _cached_exe() -> str:
	return os.path.join(_app_dir(), "ffmpeg.exe")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class FfmpegNotReadyError(RuntimeError):
	"""Raised by get_ffmpeg_exe() when ffmpeg has not been downloaded yet."""


def is_ffmpeg_cached() -> bool:
	"""Return True if a cached ffmpeg.exe already exists."""
	return os.path.isfile(_cached_exe())


def get_ffmpeg_exe() -> str:
	"""Return path to the cached ffmpeg binary.

	Raises FfmpegNotReadyError if it has not been downloaded yet — the
	caller (GUI) should then trigger download_ffmpeg() first.
	"""
	path = _cached_exe()
	if os.path.isfile(path):
		return path
	raise FfmpegNotReadyError(
		"ffmpeg is not available yet.\n"
		"It will be downloaded automatically before your first video job."
	)


def download_ffmpeg(
	progress_cb: Optional[Callable[[int, Optional[int]], None]] = None,
) -> str:
	"""Download the latest essentials ffmpeg build and cache ffmpeg.exe.

	Steps
	-----
	1. Query the GyanD/codexffmpeg GitHub Releases API for the latest release.
	2. Find the ``*essentials_build.zip`` asset.
	3. Stream-download the zip with progress callbacks.
	4. Extract only ``*/bin/ffmpeg.exe`` from the archive.
	5. Write it to the cache directory and make it executable.

	Returns the path to the cached ffmpeg.exe.
	Raises RuntimeError on any failure.
	"""
	# --- 1. Resolve the download URL ---
	req = urllib.request.Request(_FFMPEG_API_URL, headers={"User-Agent": _UA})
	with urllib.request.urlopen(req, timeout=_TIMEOUT_API) as resp:
		release = json.loads(resp.read().decode("utf-8"))

	asset_url = None
	for asset in release.get("assets", []):
		name = asset.get("name", "")
		if name.endswith(_ASSET_SUFFIX):
			asset_url = asset["browser_download_url"]
			break

	if not asset_url:
		raise RuntimeError(
			f"Could not find an *{_ASSET_SUFFIX} asset in the latest "
			f"GyanD/codexffmpeg release ({release.get('tag_name', '?')}).\n"
			"Please check https://github.com/GyanD/codexffmpeg/releases and "
			"download ffmpeg.exe manually into:\n"
			f"  {_cache_dir()}"
		)

	# --- 2. Stream-download ---
	req2 = urllib.request.Request(asset_url, headers={"User-Agent": _UA})
	with urllib.request.urlopen(req2, timeout=_TIMEOUT_DOWNLOAD) as resp:
		total = int(resp.headers.get("Content-Length", 0)) or None
		done  = 0
		buf   = io.BytesIO()
		while True:
			chunk = resp.read(65536)
			if not chunk:
				break
			buf.write(chunk)
			done += len(chunk)
			if progress_cb:
				progress_cb(done, total)

	# --- 3. Extract ffmpeg.exe ---
	buf.seek(0)
	dest = _cached_exe()
	tmp  = dest + ".tmp"

	with zipfile.ZipFile(buf) as zf:
		# The zip layout is: ffmpeg-<ver>-essentials_build/bin/ffmpeg.exe
		ffmpeg_entry = next(
			(n for n in zf.namelist() if n.endswith("/bin/ffmpeg.exe")),
			None,
		)
		if not ffmpeg_entry:
			raise RuntimeError(
				"ffmpeg.exe not found inside the downloaded zip.\n"
				"Zip contents: " + ", ".join(zf.namelist()[:10])
			)

		with zf.open(ffmpeg_entry) as src, open(tmp, "wb") as fh:
			shutil.copyfileobj(src, fh)

	# Atomic replace
	if os.path.exists(dest):
		os.remove(dest)
	os.rename(tmp, dest)

	return dest
