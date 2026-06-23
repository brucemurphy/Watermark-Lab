"""Persistent preferences for Watermark Lab.

Stores presets.json and recent.json in the app's data folder:
  * Frozen (PyInstaller onedir) build -> <app>/_internal  (keeps the exe
    folder clean; only WatermarkLab.exe and ffmpeg.exe live in the root).
  * Running from source              -> the folder of this script.

Reads degrade to empty defaults and writes are atomic. If the folder is not
writable the write is silently ignored - values are simply not persisted.

A startup migration moves any stray prefs files (or leftover .tmp files) that
end up in the exe root into the data folder and deletes the root copies, so
the root stays clean even if an older build wrote there.
"""
import json
import os
import shutil
import sys
import tempfile

_MAX_RECENT = 8
_FILES = ("presets.json", "recent.json", "settings.json")

_VALID_UI_MODES = ("modern", "classic")
_DEFAULT_UI_MODE = "modern"


def _is_frozen():
    return getattr(sys, "frozen", False)


def _exe_root():
    """Folder containing the running exe (frozen) or this script (source)."""
    if _is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _data_dir():
    """Preferred storage folder.

    Frozen onedir: <root>/_internal when present (keeps the exe folder tidy).
    Falls back to the exe root if _internal is missing. Source runs use the
    script folder.
    """
    root = _exe_root()
    if _is_frozen():
        internal = os.path.join(root, "_internal")
        if os.path.isdir(internal):
            return internal
    return root


def _migrate_from_root():
    """Move any prefs files left in the exe root into the data folder, then
    remove the root copies so the folder stays clean. Also sweeps stray
    atomic-write .tmp leftovers. No-op when data dir == root (source runs).
    """
    if not _is_frozen():
        return
    root = _exe_root()
    data = _data_dir()
    if os.path.normcase(os.path.abspath(root)) == os.path.normcase(os.path.abspath(data)):
        return
    try:
        for name in _FILES:
            root_path = os.path.join(root, name)
            if os.path.isfile(root_path):
                data_path = os.path.join(data, name)
                # Only copy across if the data folder doesn't already have it
                # (the data-folder copy is authoritative once migrated).
                if not os.path.isfile(data_path):
                    try:
                        shutil.copy2(root_path, data_path)
                    except Exception:
                        pass
                try:
                    os.remove(root_path)
                except Exception:
                    pass
        # Sweep stray atomic-write temp files left in the root by old builds.
        for entry in os.listdir(root):
            if entry.endswith(".tmp") and (
                entry.startswith("presets.json.") or entry.startswith("recent.json.")
            ):
                try:
                    os.remove(os.path.join(root, entry))
                except Exception:
                    pass
    except Exception:
        pass


# Run once at import (i.e. at app startup) so the reads below see migrated
# files and the root is cleaned on every launch.
_migrate_from_root()


def _load(filename):
    path = os.path.join(_data_dir(), filename)
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


def _save(filename, data):
    """Atomically write JSON to the data folder; silently ignore any failure."""
    folder = _data_dir()
    target = os.path.join(folder, filename)
    tmp_fd = tmp_path = None
    try:
        tmp_fd, tmp_path = tempfile.mkstemp(prefix=filename + ".", suffix=".tmp", dir=folder)
        with os.fdopen(tmp_fd, "w", encoding="utf-8") as fh:
            tmp_fd = None
            json.dump(data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp_path, target)
        tmp_path = None
    except Exception:
        pass
    finally:
        if tmp_fd is not None:
            try:
                os.close(tmp_fd)
            except Exception:
                pass
        if tmp_path is not None and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def load_presets():
    data = _load("presets.json")
    return data if isinstance(data, dict) else {}


def save_preset(name, text, color, transparency):
    presets = load_presets()
    presets[name] = {"text": text, "color": color, "transparency": transparency}
    _save("presets.json", presets)


def delete_preset(name):
    presets = load_presets()
    presets.pop(name, None)
    _save("presets.json", presets)


def load_recent():
    data = _load("recent.json")
    if isinstance(data, list):
        return [p for p in data if isinstance(p, str) and os.path.isfile(p)]
    return []


def add_recent(path):
    recent = load_recent()
    path = os.path.abspath(path)
    if path in recent:
        recent.remove(path)
    recent.insert(0, path)
    _save("recent.json", recent[:_MAX_RECENT])


def load_ui_mode():
    """Return the remembered UI mode: "modern" (Qt) or "classic" (Tkinter).

    Falls back to the default if settings.json is missing or invalid.
    """
    data = _load("settings.json")
    if isinstance(data, dict):
        mode = data.get("ui_mode")
        if mode in _VALID_UI_MODES:
            return mode
    return _DEFAULT_UI_MODE


def save_ui_mode(mode):
    """Persist the chosen UI mode; ignores anything outside the valid set."""
    if mode not in _VALID_UI_MODES:
        return
    data = _load("settings.json")
    if not isinstance(data, dict):
        data = {}
    data["ui_mode"] = mode
    _save("settings.json", data)


def _default_browse_dir():
    """Best default directory for a first-time browse: the user's Documents
    folder, falling back to the home folder, then the current directory."""
    home = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    docs = os.path.join(home, "Documents")
    if os.path.isdir(docs):
        return docs
    if os.path.isdir(home):
        return home
    return os.getcwd()


def load_last_dir():
    """Return the last directory the user browsed from.

    Falls back to the Documents folder when unset or no longer present.
    """
    data = _load("settings.json")
    if isinstance(data, dict):
        last = data.get("last_dir")
        if isinstance(last, str) and os.path.isdir(last):
            return last
    return _default_browse_dir()


def save_last_dir(path):
    """Remember the directory of the user's most recent file/folder pick."""
    if not isinstance(path, str) or not path:
        return
    directory = path if os.path.isdir(path) else os.path.dirname(path)
    if not os.path.isdir(directory):
        return
    data = _load("settings.json")
    if not isinstance(data, dict):
        data = {}
    data["last_dir"] = directory
    _save("settings.json", data)
