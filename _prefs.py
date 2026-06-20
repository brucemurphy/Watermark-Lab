"""Persistent preferences for Watermark Lab.

Stores two JSON files in the same folder as the running exe (or script):
  presets.json  — named watermark configs (text, color, transparency)
  recent.json   — last N file paths used

Keeping them alongside the app (not in %APPDATA%) means the whole
WatermarkLab folder stays self-contained and portable.
"""
import json
import os
import sys

_MAX_RECENT = 8


def _prefs_dir() -> str:
	"""Return the folder containing the running exe or this script."""
	if getattr(sys, "frozen", False):
		return os.path.dirname(os.path.abspath(sys.executable))
	return os.path.dirname(os.path.abspath(__file__))


def _load(filename: str) -> object:
	path = os.path.join(_prefs_dir(), filename)
	try:
		with open(path, "r", encoding="utf-8") as f:
			return json.load(f)
	except Exception:
		return None


def _save(filename: str, data: object) -> None:
	path = os.path.join(_prefs_dir(), filename)
	try:
		with open(path, "w", encoding="utf-8") as f:
			json.dump(data, f, indent=2, ensure_ascii=False)
	except Exception:
		pass


# ---------------------------------------------------------------------------
# Presets
# ---------------------------------------------------------------------------

def load_presets() -> dict:
	"""Return {name: {text, color, transparency}} or {}."""
	data = _load("presets.json")
	return data if isinstance(data, dict) else {}


def save_preset(name: str, text: str, color: str, transparency: float) -> None:
	presets = load_presets()
	presets[name] = {"text": text, "color": color, "transparency": transparency}
	_save("presets.json", presets)


def delete_preset(name: str) -> None:
	presets = load_presets()
	presets.pop(name, None)
	_save("presets.json", presets)


# ---------------------------------------------------------------------------
# Recent files
# ---------------------------------------------------------------------------

def load_recent() -> list:
	"""Return list of recent file paths (newest first)."""
	data = _load("recent.json")
	if isinstance(data, list):
		return [p for p in data if os.path.isfile(p)]
	return []


def add_recent(path: str) -> None:
	recent = load_recent()
	path = os.path.abspath(path)
	if path in recent:
		recent.remove(path)
	recent.insert(0, path)
	_save("recent.json", recent[:_MAX_RECENT])
