"""Auto-update helper for Watermark Lab (onedir build).

How it works
------------
On each launch the GUI calls ``check_for_update_async`` which spins up a
background thread that:

1. Calls the GitHub Releases API to find the latest published release.
2. Compares the remote version tag against APP_VERSION (PEP-440).
3. If a newer build is available, invokes ``on_update_available`` on the
   GUI thread via tkinter after.
4. When the user confirms apply_update:
     a. Downloads WatermarkLab.zip from the release assets.
     b. Extracts it to a sibling folder WatermarkLab_new.
     c. Launches a hidden PowerShell script that waits for this process to
        exit, renames WatermarkLab to WatermarkLab_old, renames
        WatermarkLab_new to WatermarkLab, relaunches WatermarkLab.exe
        and deletes WatermarkLab_old.

GitHub is the single source of truth. No-op when running from source.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import urllib.request
import zipfile
from typing import Callable, Optional

from packaging.version import Version

from _version import APP_VERSION

GITHUB_REPO      = "brucemurphy/Watermark-Lab"
RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
ZIP_ASSET_NAME   = "WatermarkLab.zip"

_TIMEOUT = 15
_UA      = f"WatermarkLab/{APP_VERSION} (update-check)"

_SWAP_PS1 = """
param($OldDir, $NewDir, $Exe)
$ErrorActionPreference = 'Stop'
for ($i = 0; $i -lt 60; $i++) {
    if (-not (Get-Process -Name 'WatermarkLab' -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Seconds 1
}
$backup = "$OldDir`_old"
if (Test-Path $backup) { Remove-Item $backup -Recurse -Force }
Rename-Item -LiteralPath $OldDir -NewName $backup
Rename-Item -LiteralPath $NewDir -NewName $OldDir
Start-Process -FilePath $Exe
Start-Sleep -Seconds 5
Remove-Item $backup -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
"""


def _api_get(url):
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _fetch_latest_release():
    return _api_get(RELEASES_API_URL)


def _zip_download_url(release):
    for asset in release.get("assets", []):
        if asset.get("name", "").lower() == ZIP_ASSET_NAME.lower():
            return asset["browser_download_url"]
    raise RuntimeError(f"Release {release.get('tag_name')} does not contain {ZIP_ASSET_NAME}.")


def _is_newer(remote_tag):
    try:
        return Version(remote_tag.lstrip("v")) > Version(APP_VERSION)
    except Exception:
        return False


def _stream_download(url, dest, progress_cb=None):
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=120) as resp:
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


def check_for_update_async(tk_root, on_update_available):
    """Spawn a daemon thread to check GitHub for a newer release."""
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


def apply_update(progress_cb=None):
    """Download zip, extract to sibling folder, launch swap script, exit."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("apply_update: not running as a frozen exe.")

    exe_path = os.path.abspath(sys.executable)
    app_dir  = os.path.dirname(exe_path)
    parent   = os.path.dirname(app_dir)
    new_dir  = os.path.join(parent, "WatermarkLab_new")

    release  = _fetch_latest_release()
    zip_url  = _zip_download_url(release)

    fd, zip_tmp = tempfile.mkstemp(suffix=".zip", prefix="wml_update_")
    os.close(fd)
    try:
        _stream_download(zip_url, zip_tmp, progress_cb)

        if os.path.exists(new_dir):
            shutil.rmtree(new_dir, ignore_errors=True)

        with zipfile.ZipFile(zip_tmp) as zf:
            zf.extractall(parent)

        if not os.path.isdir(new_dir):
            candidates = [
                d for d in os.listdir(parent)
                if d.lower().startswith("watermarklab") and
                os.path.isdir(os.path.join(parent, d)) and
                d != os.path.basename(app_dir)
            ]
            if not candidates:
                raise RuntimeError("Could not find extracted folder after unzip.")
            os.rename(os.path.join(parent, candidates[0]), new_dir)
    finally:
        try:
            os.remove(zip_tmp)
        except OSError:
            pass

    fd2, ps1 = tempfile.mkstemp(suffix=".ps1", prefix="wml_swap_")
    os.close(fd2)
    new_exe = os.path.join(new_dir, "WatermarkLab.exe")
    with open(ps1, "w", encoding="utf-8") as f:
        f.write(_SWAP_PS1)

    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-NonInteractive",
         "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
         "-File", ps1, "-OldDir", app_dir, "-NewDir", new_dir, "-Exe", new_exe],
        creationflags=subprocess.CREATE_NO_WINDOW,
        startupinfo=si,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL, close_fds=True,
    )
    sys.exit(0)


def cleanup_old_exe():
    """Remove WatermarkLab_old left by a previous update swap."""
    if not getattr(sys, "frozen", False):
        return
    app_dir = os.path.dirname(os.path.abspath(sys.executable))
    parent  = os.path.dirname(app_dir)
    old_dir = os.path.join(parent, "WatermarkLab_old")
    if os.path.isdir(old_dir):
        shutil.rmtree(old_dir, ignore_errors=True)
