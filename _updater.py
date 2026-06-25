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

from packaging.version import Version

from _version import APP_VERSION

GITHUB_REPO    = "brucemurphy/Watermark-Lab"
ZIP_ASSET_NAME = "WatermarkLab.zip"

_TIMEOUT = 15
_UA      = f"WatermarkLab/{APP_VERSION} (update-check)"

_SWAP_PS1 = """
param($SrcDir, $DestDir)

# 1. Wait for WatermarkLab to fully exit (up to 60s)
for ($i = 0; $i -lt 60; $i++) {
    if (-not (Get-Process -Name 'WatermarkLab' -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Seconds 1
}
Start-Sleep -Seconds 2

# 2. Overwrite the existing app folder in-place - no renaming, no sibling folders.
#    robocopy copies every file from $SrcDir into $DestDir, overwriting existing files.
#    Exit codes 0-7 are success for robocopy (8+ are errors).
robocopy $SrcDir $DestDir /E /IS /IT /NFL /NDL /NJH /NJS /NC /NS /NP
if ($LASTEXITCODE -ge 8) {
    exit 1
}

# 3. Launch the updated exe from the same folder it always lived in
Start-Process -FilePath (Join-Path $DestDir 'WatermarkLab.exe')

# 4. Clean up temp source and this script
Start-Sleep -Seconds 3
Remove-Item -LiteralPath $SrcDir -Recurse -Force -ErrorAction SilentlyContinue
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
    """Download zip, extract to %TEMP%, robocopy over existing folder after exit."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("apply_update: not running as a frozen exe.")

    app_dir  = os.path.dirname(os.path.abspath(sys.executable))
    tmp_root = tempfile.gettempdir()
    stage    = os.path.join(tmp_root, "WatermarkLab_update")

    release = _fetch_latest_release()
    zip_url = _zip_download_url(release)

    fd, zip_tmp = tempfile.mkstemp(suffix=".zip", prefix="wml_", dir=tmp_root)
    os.close(fd)

    try:
        _stream_download(zip_url, zip_tmp, progress_cb)

        if os.path.exists(stage):
            shutil.rmtree(stage, ignore_errors=True)
        os.makedirs(stage)

        with zipfile.ZipFile(zip_tmp) as zf:
            zf.extractall(stage)

        # Zip has one top-level folder - find it
        entries = [e for e in os.listdir(stage)
                   if os.path.isdir(os.path.join(stage, e))]
        if not entries:
            raise RuntimeError("No folder found in downloaded zip.")
        src_dir = os.path.join(stage, entries[0])

    finally:
        try:
            os.remove(zip_tmp)
        except OSError:
            pass

    # Write PS1 to %TEMP%
    fd2, ps1 = tempfile.mkstemp(suffix=".ps1", prefix="wml_swap_", dir=tmp_root)
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
         "-SrcDir",  src_dir,
         "-DestDir", app_dir],
        creationflags=subprocess.CREATE_NO_WINDOW,
        startupinfo=si,
        stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL, close_fds=True,
    )
    sys.exit(0)


def cleanup_old_exe():
    """Remove leftover update staging from previous runs."""
    if not getattr(sys, "frozen", False):
        return
    app_dir  = os.path.dirname(os.path.abspath(sys.executable))
    parent   = os.path.dirname(app_dir)
    tmp_root = tempfile.gettempdir()
    # Sweep both beside-the-app (legacy) and %TEMP% (current) locations
    for base in (parent, tmp_root):
        for name in ("WatermarkLab_new", "WatermarkLab_old",
                     "WatermarkLab_update", "WatermarkLab_extract_tmp"):
            d = os.path.join(base, name)
            if os.path.isdir(d):
                shutil.rmtree(d, ignore_errors=True)
