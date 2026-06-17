"""Auto-update helper for Watermark Lab.

How it works
------------
On each launch the GUI calls ``check_for_update_async`` which spins up a
background thread that:

1. Fetches ``version.json`` from the public Azure Blob URL.
2. Compares the remote version against APP_VERSION using PEP-440 semantics.
3. If a newer build exists, invokes the supplied ``on_update_available``
   callback on the calling thread (via tkinter's ``after``).
4. When the user confirms, ``apply_update`` downloads the new exe, swaps it
   in-place, and restarts the process.

The module is a no-op when the app is running from Python source (i.e.
``sys.frozen`` is not set), so developers are never bothered by update prompts.
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
from typing import Callable, Optional
from packaging.version import Version

from _version import APP_VERSION

# ---------------------------------------------------------------------------
# Public constants — must match the container / blob names created by CI.
# ---------------------------------------------------------------------------
STORAGE_ACCOUNT = "watermarklab"
CONTAINER       = "releases"
VERSION_BLOB    = "version.json"
VERSION_URL     = (
	f"https://{STORAGE_ACCOUNT}.blob.core.windows.net/{CONTAINER}/{VERSION_BLOB}"
)

_TIMEOUT = 10  # seconds for network requests


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _fetch_version_manifest() -> dict:
	"""Download and parse version.json from blob storage."""
	with urllib.request.urlopen(VERSION_URL, timeout=_TIMEOUT) as resp:
		return json.loads(resp.read().decode("utf-8"))


def _is_newer(remote: str) -> bool:
	try:
		return Version(remote) > Version(APP_VERSION)
	except Exception:
		return False


def _sha256_file(path: str) -> str:
	h = hashlib.sha256()
	with open(path, "rb") as fh:
		for chunk in iter(lambda: fh.read(65536), b""):
			h.update(chunk)
	return h.hexdigest()


def _download_exe(url: str, dest: str, progress_cb: Optional[Callable] = None) -> None:
	"""Stream-download *url* to *dest*, calling progress_cb(bytes_done, total)."""
	with urllib.request.urlopen(url, timeout=60) as resp:
		total = int(resp.headers.get("Content-Length", 0)) or None
		done = 0
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
	"""Spawn a daemon thread to check for updates.

	``on_update_available(remote_version, release_notes)`` is called via
	``tk_root.after`` so it runs on the main/GUI thread.

	Does nothing when running from Python source (not a frozen exe).
	"""
	if not getattr(sys, "frozen", False):
		return  # running from source — skip update check

	def _worker():
		try:
			manifest = _fetch_version_manifest()
			remote = manifest.get("version", "")
			if _is_newer(remote):
				notes = manifest.get("notes", "")
				tk_root.after(0, lambda: on_update_available(remote, notes))
		except Exception:
			pass  # network unavailable or malformed manifest — silent fail

	threading.Thread(target=_worker, daemon=True).start()


def apply_update(
	manifest_url: str = VERSION_URL,
	progress_cb: Optional[Callable] = None,
) -> None:
	"""Download the new exe, swap it in-place, and restart.

	Sequence
	--------
	1. Fetch version.json to get the download URL and expected SHA-256.
	2. Download the new exe to a sibling ``.tmp`` file.
	3. Verify SHA-256 if the manifest supplies one.
	4. Rename the running exe to ``.old`` (Windows lets you rename an
	   open file; it is deleted on next clean run by ``_cleanup_old``).
	5. Move the ``.tmp`` file to the original exe name.
	6. Launch the new exe and exit this process.

	Raises ``RuntimeError`` on any failure so the caller can show an error.
	"""
	if not getattr(sys, "frozen", False):
		raise RuntimeError("apply_update called outside of a frozen exe — nothing to do.")

	manifest = _fetch_version_manifest()
	exe_url: str = manifest.get("url", "")
	expected_sha: str = manifest.get("sha256", "")

	if not exe_url:
		raise RuntimeError("version.json is missing the 'url' field.")

	current_exe = sys.executable
	tmp_exe     = current_exe + ".tmp"
	old_exe     = current_exe + ".old"

	try:
		_download_exe(exe_url, tmp_exe, progress_cb)

		if expected_sha:
			actual = _sha256_file(tmp_exe)
			if actual.lower() != expected_sha.lower():
				raise RuntimeError(
					f"SHA-256 mismatch after download.\n"
					f"  expected: {expected_sha}\n"
					f"  actual:   {actual}"
				)

		# Remove a leftover .old from a previous update, if any.
		if os.path.exists(old_exe):
			try:
				os.remove(old_exe)
			except OSError:
				pass

		os.rename(current_exe, old_exe)
		shutil.move(tmp_exe, current_exe)

		# Launch the updated exe and exit cleanly.
		subprocess.Popen([current_exe], close_fds=True)
		sys.exit(0)

	except Exception:
		# Roll back: restore the original if the rename succeeded.
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
	"""Remove the ``<exe>.old`` leftover from a previous update, if present.

	Call once at startup so stale files do not accumulate.
	Does nothing when running from source.
	"""
	if not getattr(sys, "frozen", False):
		return
	old_exe = sys.executable + ".old"
	if os.path.exists(old_exe):
		try:
			os.remove(old_exe)
		except OSError:
			pass
