"""Auto-update helper for Watermark Lab.

How it works
------------
On each launch the GUI calls ``check_for_update_async`` which spins up a
background thread that:

1. Calls the GitHub Releases API to find the latest published release.
2. Compares the remote version tag against APP_VERSION (PEP-440).
3. If a newer build is available, invokes ``on_update_available`` on the
   GUI thread via tkinter's ``after``.
4. When the user confirms, ``apply_update`` streams the new exe from
   GitHub's CDN, swaps in-place, and restarts.

GitHub is the single source of truth — no external manifest file or cloud
storage account is required.

The module is a no-op when running from Python source (``sys.frozen`` not
set), so developers are never interrupted by update prompts.
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import threading
import urllib.request
from typing import Callable, Optional

from packaging.version import Version

from _version import APP_VERSION

# ---------------------------------------------------------------------------
# GitHub repository coordinates
# ---------------------------------------------------------------------------
GITHUB_REPO      = "brucemurphy/Watermark-Lab"
RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
EXE_ASSET_NAME   = "WatermarkLab.exe"

_TIMEOUT = 15
_UA      = f"WatermarkLab/{APP_VERSION} (update-check)"


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _api_get(url: str) -> dict:
	req = urllib.request.Request(url, headers={"User-Agent": _UA})
	with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
		return json.loads(resp.read().decode("utf-8"))


def _fetch_latest_release() -> dict:
	return _api_get(RELEASES_API_URL)


def _exe_download_url(release: dict) -> str:
	for asset in release.get("assets", []):
		if asset.get("name", "").lower() == EXE_ASSET_NAME.lower():
			return asset["browser_download_url"]
	raise RuntimeError(
		f"Release {release.get('tag_name')} does not contain {EXE_ASSET_NAME}.\n"
		"The CI build may still be running — try again in a few minutes."
	)


def _is_newer(remote_tag: str) -> bool:
	try:
		return Version(remote_tag.lstrip("v")) > Version(APP_VERSION)
	except Exception:
		return False


def _stream_download(
	url: str,
	dest: str,
	progress_cb: Optional[Callable[[int, Optional[int]], None]] = None,
) -> None:
	req = urllib.request.Request(url, headers={"User-Agent": _UA})
	with urllib.request.urlopen(req, timeout=60) as resp:
		total = int(resp.headers.get("Content-Length", 0)) or None
		done  = 0
		with open(dest, "wb") as fh:
			while True:
				chunk = resp.read(65536)
				if not chunk:
					break
				fh.write(chunk)
				done += len(chunk)
				if progress_cb:
					progress_cb(done, total)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def check_for_update_async(
	tk_root,
	on_update_available: Callable[[str, str], None],
) -> None:
	"""Spawn a daemon thread to check GitHub for a newer release.

	``on_update_available(remote_version, release_notes)`` is scheduled on
	the GUI thread via ``tk_root.after`` if a newer build is found.

	Silent no-op when running from source, network unavailable, or rate-limited.
	"""
	if not getattr(sys, "frozen", False):
		return

	def _worker():
		try:
			release = _fetch_latest_release()
			tag     = release.get("tag_name", "")
			if _is_newer(tag):
				notes   = release.get("body", "").strip()
				version = tag.lstrip("v")
				tk_root.after(0, lambda: on_update_available(version, notes))
		except Exception:
			pass

	threading.Thread(target=_worker, daemon=True).start()


def apply_update(
	progress_cb: Optional[Callable[[int, Optional[int]], None]] = None,
) -> None:
	"""Download the latest GitHub release exe, swap in-place, and restart.

	Raises ``RuntimeError`` on any failure so the caller can show an error.
	On failure the original exe is restored if possible.
	"""
	if not getattr(sys, "frozen", False):
		raise RuntimeError("apply_update: not running as a frozen exe.")

	release     = _fetch_latest_release()
	exe_url     = _exe_download_url(release)
	current_exe = sys.executable
	tmp_exe     = current_exe + ".tmp"
	old_exe     = current_exe + ".old"

	try:
		_stream_download(exe_url, tmp_exe, progress_cb)

		if os.path.exists(old_exe):
			try:
				os.remove(old_exe)
			except OSError:
				pass

		os.rename(current_exe, old_exe)
		shutil.move(tmp_exe, current_exe)

		subprocess.Popen([current_exe], close_fds=True)
		sys.exit(0)

	except Exception:
		if not os.path.exists(current_exe) and os.path.exists(old_exe):
			try:
				os.rename(old_exe, current_exe)
			except OSError:
				pass
		if os.path.exists(tmp_exe):
			try:
				os.remove(tmp_exe)
			except OSError:
				pass
		raise


def cleanup_old_exe() -> None:
	"""Remove the ``<exe>.old`` leftover from a previous update if present."""
	if not getattr(sys, "frozen", False):
		return
	old_exe = sys.executable + ".old"
	if os.path.exists(old_exe):
		try:
			os.remove(old_exe)
		except OSError:
			pass
