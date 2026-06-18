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

GITHUB_REPO    = "brucemurphy/Watermark-Lab"
ZIP_ASSET_NAME = "WatermarkLab.zip"

_TIMEOUT = 15
_UA      = f"WatermarkLab/{APP_VERSION} (update-check)"

_SWAP_PS1 = """
param($OldDir, $NewDir, $ExeName)
$ErrorActionPreference = 'SilentlyContinue'

# 1. Wait for the running process to exit (up to 60 s)
for ($i = 0; $i -lt 60; $i++) {
    if (-not (Get-Process -Name 'WatermarkLab' -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Seconds 1
}

# 2. Rename current app folder to _old (backup)
$backup = $OldDir + '_old'
if (Test-Path -LiteralPath $backup) {
    Remove-Item -LiteralPath $backup -Recurse -Force
}
Move-Item -LiteralPath $OldDir -Destination $backup -Force

# 3. Move new folder into place as the app folder
Move-Item -LiteralPath $NewDir -Destination $OldDir -Force

# 4. Launch the new exe from its final location
$newExe = Join-Path $OldDir $ExeName
Start-Process -FilePath $newExe

# 5. Clean up old folder and this script
Start-Sleep -Seconds 5
Remove-Item -LiteralPath $backup -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
"""


def _api_get(url):
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _fetch_latest_release() -> dict:
    """Return the highest-versioned published release by semver, not publish date."""
    releases = _api_get(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases?per_page=20"
    )
    best = None
    best_ver = None
    for r in releases:
        if r.get("draft") or r.get("prerelease"):
            continue
        tag = r.get("tag_name", "")
        try:
            v = Version(tag.lstrip("v"))
        except Exception:
            continue
        if best_ver is None or v > best_ver:
            best_ver = v
            best = r
    if not best:
        raise RuntimeError("No published releases found.")
    return best


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

    # Extract to a private temp dir first — never touch the running app_dir
    extract_tmp = os.path.join(parent, "WatermarkLab_extract_tmp")

    try:
        _stream_download(zip_url, zip_tmp, progress_cb)

        # Clean up any leftover staging dirs from a previous failed update
        for d in (new_dir, extract_tmp):
            if os.path.exists(d):
                shutil.rmtree(d, ignore_errors=True)

        # Extract into a private staging folder, never into parent directly
        os.makedirs(extract_tmp, exist_ok=True)
        with zipfile.ZipFile(zip_tmp) as zf:
            zf.extractall(extract_tmp)

        # The zip contains a single top-level folder (WatermarkLab or similar)
        entries = [
            e for e in os.listdir(extract_tmp)
            if os.path.isdir(os.path.join(extract_tmp, e))
        ]
        if not entries:
            raise RuntimeError("Zip contained no top-level folder after extraction.")

        extracted = os.path.join(extract_tmp, entries[0])
        os.rename(extracted, new_dir)

    finally:
        try:
            os.remove(zip_tmp)
        except OSError:
            pass
        # Clean up staging dir whether we succeeded or failed
        if os.path.exists(extract_tmp):
            shutil.rmtree(extract_tmp, ignore_errors=True)

    fd2, ps1 = tempfile.mkstemp(suffix=".ps1", prefix="wml_swap_")
    os.close(fd2)
    with open(ps1, "w", encoding="utf-8") as f:
        f.write(_SWAP_PS1)

    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    si.wShowWindow = 0
    subprocess.Popen(
        ["powershell.exe", "-NoProfile", "-NonInteractive",
         "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
         "-File", ps1,
         "-OldDir",  app_dir,
         "-NewDir",  new_dir,
         "-ExeName", "WatermarkLab.exe"],
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
