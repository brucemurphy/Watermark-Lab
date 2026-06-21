"""Persistent preferences for Watermark Lab.

Stores presets.json and recent.json next to the exe (or this script). Reads
degrade to empty defaults and writes are atomic. If the app folder is not
writable the write is silently ignored - values are simply not persisted.
"""
import json
import os
import sys
import tempfile

_MAX_RECENT = 8


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _load(filename):
    path = os.path.join(_app_dir(), filename)
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


def _save(filename, data):
    """Atomically write JSON to the app folder; silently ignore any failure."""
    folder = _app_dir()
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
