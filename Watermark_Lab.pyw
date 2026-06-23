"""Watermark Lab — PySide6/Qt desktop UI with true-render preview (v2.0.0+).

This is the primary application entry point. It reuses the shared backend
modules (_powerpoint, _word, _video, _prefs, _ffmpeg, _updater, _version)
together with the Qt-specific helpers (_xpowerpoint, _xword, _xpreview) to
deliver a three-pane interface with a live, pixel-accurate watermark preview.

The previous Tkinter front-end is preserved for reference under
``previous_version/Watermark_Lab.pyw``.
"""
from __future__ import annotations

import os
import subprocess
import sys

from PySide6.QtCore import Qt, QPoint, QSize, Signal, QEvent, QPropertyAnimation, Property, QRectF, QThread, QObject, QTimer, QPointF
from PySide6.QtGui import QIcon, QPixmap, QFont, QPainter, QColor, QBrush, QPen, QPainterPath, QImage
from PySide6.QtWidgets import (
	QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
	QHBoxLayout, QVBoxLayout, QFileDialog, QScrollArea, QMenu, QSizePolicy,
	QLineEdit, QSlider, QComboBox, QColorDialog, QInputDialog, QGridLayout,
	QSplashScreen,
)

from _version import APP_VERSION
from _prefs import (
	load_presets, save_preset, delete_preset, load_recent, add_recent,
)
from _xpowerpoint import add_watermark  # experimental: OneDrive-safe save
from _word import WORD_EXTS
from _xword import add_word_watermark  # experimental: snug, wrapped Word watermark
from _video import add_video_watermark, VIDEO_EXTS, is_video_file
from _ffmpeg import is_ffmpeg_cached, download_ffmpeg, FfmpegNotReadyError
from _xpreview import PreviewController, PreviewCanvas, file_kind

PPT_EXTS = {".pptx", ".ppt"}
ALL_SUPPORTED = PPT_EXTS | WORD_EXTS | VIDEO_EXTS

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_ICO   = os.path.join(SCRIPT_DIR, "Watermark.ico")
ICON_PNG   = os.path.join(SCRIPT_DIR, "Watermark.png")
SPLASH_PNG = os.path.join(SCRIPT_DIR, "SplashLab.png")

# Help & Support opens the project's GitHub repo. Derived from the updater's
# repo constant so the two never drift apart.
try:
	from _updater import GITHUB_REPO as _GITHUB_REPO
except Exception:
	_GITHUB_REPO = "brucemurphy/Watermark-Lab"
HELP_URL = f"https://github.com/{_GITHUB_REPO}"

# Segoe Fluent Icons (Win11) / Segoe MDL2 Assets (Win10) glyphs
# Vector icon names (drawn by vector_icon — no icon-font dependency)
ICON_FILE     = "doc_add"
ICON_OPENFILE = "folder"
ICON_FOLDER   = "folder"
ICON_ADD      = "plus"
ICON_ADDTO    = "doc_add"
ICON_SETTINGS = "settings"
ICON_SAVE     = "save"
ICON_DELETE   = "trash"
ICON_INFO     = "info"
ICON_TEXT     = "text"
ICON_ZOOM_IN  = "plus"
ICON_ZOOM_OUT = "minus"
ICON_PAN      = "fit"
ICON_FULL     = "fit"
ICON_STAMP    = "stamp"
ICON_EXPORT   = "export"
ICON_EYEDROP  = "eyedropper"
ICON_CHECK    = "check"
ICON_HELP     = "help"
ICON_MORE     = "more"
ICON_CHEVRON  = "chevron"
ICON_EYE      = "eye"
ICON_EYE_OFF  = "eye_off"

# ─────────────────────────────────────────────────────────────────────────────
# Palette — tuned to the mockup (darker, bluer than the legacy Tkinter theme)
# while keeping the signature gold accent.
# ─────────────────────────────────────────────────────────────────────────────
BG            = "#0d1117"   # app background
PANEL         = "#161b22"   # card background
PANEL_HI      = "#1c2230"   # raised element (inputs, hover)
FIELD         = "#0b0e14"   # sunken field background
BORDER        = "#272e3a"   # subtle card / control border
BORDER_HI     = "#39414f"   # hover / focus border
FG            = "#e6edf3"   # primary text
FG_SOFT       = "#c4ccd6"   # secondary text
MUTED         = "#8b949e"   # tertiary / hint text
ACCENT        = "#FDC301"   # signature gold
ACCENT_HOVER  = "#ffce33"
ACCENT_ACTIVE = "#d9a600"
ACCENT_FG     = "#1b1b1b"   # text/icon on top of gold
SUCCESS       = "#3fb950"   # ready / ok green
DANGER        = "#f0556b"   # delete / error red
PREVIEW_BG    = "#0a0d12"   # canvas backdrop behind the document


def icon_font(size: int = 12) -> QFont:
    """Return the best available Windows icon font at the given pixel size."""
    for family in ("Segoe Fluent Icons", "Segoe MDL2 Assets"):
        f = QFont(family)
        f.setPixelSize(size)
        return f
    return QFont("Segoe UI", size)


def _draw_icon(p: QPainter, name: str, s: float, c: QColor) -> None:
    """Paint a single vector icon into a logical s-by-s box at the origin."""
    cx, cy = s / 2.0, s / 2.0
    p.setBrush(Qt.NoBrush)

    if name in ("minus", "plus"):
        half = s * 0.30
        p.drawLine(QPointF(cx - half, cy), QPointF(cx + half, cy))
        if name == "plus":
            p.drawLine(QPointF(cx, cy - half), QPointF(cx, cy + half))

    elif name in ("eye", "eye_off"):
        w, h = s * 0.40, s * 0.26
        path = QPainterPath()
        path.moveTo(cx - w, cy)
        path.quadTo(cx, cy - h * 2, cx + w, cy)
        path.quadTo(cx, cy + h * 2, cx - w, cy)
        p.drawPath(path)
        r = s * 0.11
        p.setBrush(QBrush(c)); p.drawEllipse(QPointF(cx, cy), r, r); p.setBrush(Qt.NoBrush)
        if name == "eye_off":
            d = s * 0.34
            p.drawLine(QPointF(cx - d, cy - d), QPointF(cx + d, cy + d))

    elif name == "fit":
        a, head = s * 0.30, s * 0.12
        for sx, sy in ((-1, -1), (1, -1), (-1, 1), (1, 1)):
            x2, y2 = cx + sx * a, cy + sy * a
            x1, y1 = cx + sx * (a * 0.45), cy + sy * (a * 0.45)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
            p.drawLine(QPointF(x2, y2), QPointF(x2 - sx * head, y2))
            p.drawLine(QPointF(x2, y2), QPointF(x2, y2 - sy * head))

    elif name == "folder":
        x, y, w, h = s * 0.16, s * 0.30, s * 0.68, s * 0.44
        tab = QPainterPath()
        tab.moveTo(x, y)
        tab.lineTo(x + w * 0.34, y)
        tab.lineTo(x + w * 0.46, y - s * 0.09)
        tab.lineTo(x + w, y - s * 0.09)
        p.drawPath(tab)
        p.drawRoundedRect(QRectF(x, y, w, h), s * 0.06, s * 0.06)

    elif name == "folder_add":
        x, y, w, h = s * 0.14, s * 0.30, s * 0.62, s * 0.44
        p.drawRoundedRect(QRectF(x, y, w, h), s * 0.06, s * 0.06)
        # plus badge bottom-right
        bx, by, hp = x + w, y + h, s * 0.13
        p.drawLine(QPointF(bx - hp, by), QPointF(bx + hp, by))
        p.drawLine(QPointF(bx, by - hp), QPointF(bx, by + hp))

    elif name == "doc_add":
        # Document with a plus — the drop zone glyph.
        x, y, w, h = s * 0.26, s * 0.16, s * 0.40, s * 0.58
        fold = s * 0.13
        path = QPainterPath()
        path.moveTo(x, y)
        path.lineTo(x + w - fold, y)
        path.lineTo(x + w, y + fold)
        path.lineTo(x + w, y + h)
        path.lineTo(x, y + h)
        path.closeSubpath()
        p.drawPath(path)
        p.drawLine(QPointF(x + w - fold, y), QPointF(x + w - fold, y + fold))
        p.drawLine(QPointF(x + w - fold, y + fold), QPointF(x + w, y + fold))
        bx, by, hp = x + w * 0.5, y + h * 0.62, s * 0.11
        p.drawLine(QPointF(bx - hp, by), QPointF(bx + hp, by))
        p.drawLine(QPointF(bx, by - hp), QPointF(bx, by + hp))

    elif name == "settings":
        # Gear: outer ring + teeth + hub.
        import math
        rout, rin, hub = s * 0.34, s * 0.22, s * 0.10
        teeth = QPainterPath()
        for k in range(8):
            ang = math.radians(k * 45)
            x1 = cx + rout * math.cos(ang)
            y1 = cy + rout * math.sin(ang)
            if k == 0:
                teeth.moveTo(x1, y1)
            else:
                teeth.lineTo(x1, y1)
        p.drawEllipse(QPointF(cx, cy), rin, rin)
        for k in range(8):
            ang = math.radians(k * 45)
            x1 = cx + rin * math.cos(ang); y1 = cy + rin * math.sin(ang)
            x2 = cx + rout * math.cos(ang); y2 = cy + rout * math.sin(ang)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
        p.setBrush(QBrush(c)); p.drawEllipse(QPointF(cx, cy), hub, hub); p.setBrush(Qt.NoBrush)

    elif name == "info":
        r = s * 0.36
        p.drawEllipse(QPointF(cx, cy), r, r)
        p.setBrush(QBrush(c))
        p.drawEllipse(QPointF(cx, cy - s * 0.17), s * 0.045, s * 0.045)
        p.setBrush(Qt.NoBrush)
        p.drawLine(QPointF(cx, cy - s * 0.04), QPointF(cx, cy + s * 0.17))

    elif name == "text":
        # Capital "A" mark for preset rows.
        p.drawLine(QPointF(cx - s * 0.20, cy + s * 0.22), QPointF(cx, cy - s * 0.24))
        p.drawLine(QPointF(cx, cy - s * 0.24), QPointF(cx + s * 0.20, cy + s * 0.22))
        p.drawLine(QPointF(cx - s * 0.11, cy + s * 0.04), QPointF(cx + s * 0.11, cy + s * 0.04))

    elif name == "more":
        r = s * 0.045
        p.setBrush(QBrush(c))
        for dy in (-s * 0.20, 0, s * 0.20):
            p.drawEllipse(QPointF(cx, cy + dy), r, r)
        p.setBrush(Qt.NoBrush)

    elif name == "chevron":
        w, h = s * 0.16, s * 0.22
        p.drawLine(QPointF(cx - w * 0.5, cy - h), QPointF(cx + w, cy))
        p.drawLine(QPointF(cx + w, cy), QPointF(cx - w * 0.5, cy + h))

    elif name == "save":
        x, y, w, h = s * 0.20, s * 0.20, s * 0.60, s * 0.60
        path = QPainterPath()
        path.moveTo(x, y); path.lineTo(x + w - s * 0.12, y)
        path.lineTo(x + w, y + s * 0.12); path.lineTo(x + w, y + h)
        path.lineTo(x, y + h); path.closeSubpath()
        p.drawPath(path)
        p.drawRect(QRectF(x + s * 0.14, y + h - s * 0.22, w - s * 0.28, s * 0.22))
        p.drawRect(QRectF(x + s * 0.16, y, w - s * 0.40, s * 0.14))

    elif name == "trash":
        x, w = s * 0.26, s * 0.48
        top, bot = s * 0.32, s * 0.78
        p.drawLine(QPointF(x - s * 0.04, top), QPointF(x + w + s * 0.04, top))   # lid
        p.drawLine(QPointF(cx - s * 0.08, top - s * 0.06), QPointF(cx + s * 0.08, top - s * 0.06))  # handle
        body = QPainterPath()
        body.moveTo(x + s * 0.02, top)
        body.lineTo(x + s * 0.06, bot)
        body.lineTo(x + w - s * 0.06, bot)
        body.lineTo(x + w - s * 0.02, top)
        p.drawPath(body)
        for dx in (-s * 0.10, 0, s * 0.10):
            p.drawLine(QPointF(cx + dx, top + s * 0.08), QPointF(cx + dx, bot - s * 0.06))

    elif name == "eyedropper":
        # Diagonal dropper.
        p.drawLine(QPointF(s * 0.30, s * 0.70), QPointF(s * 0.62, s * 0.38))
        p.drawEllipse(QRectF(s * 0.58, s * 0.16, s * 0.22, s * 0.22))
        p.setBrush(QBrush(c)); p.drawEllipse(QPointF(s * 0.30, s * 0.70), s * 0.05, s * 0.05); p.setBrush(Qt.NoBrush)

    elif name == "stamp":
        # Approval stamp.
        p.drawLine(QPointF(cx, cy - s * 0.30), QPointF(cx, cy + s * 0.02))
        p.drawEllipse(QRectF(cx - s * 0.16, cy - s * 0.34, s * 0.32, s * 0.20))
        p.drawLine(QPointF(s * 0.24, cy + s * 0.10), QPointF(s * 0.76, cy + s * 0.10))
        p.drawLine(QPointF(s * 0.20, cy + s * 0.28), QPointF(s * 0.80, cy + s * 0.28))

    elif name == "export":
        # Up arrow out of a tray.
        p.drawLine(QPointF(cx, cy - s * 0.30), QPointF(cx, cy + s * 0.12))
        p.drawLine(QPointF(cx, cy - s * 0.30), QPointF(cx - s * 0.14, cy - s * 0.14))
        p.drawLine(QPointF(cx, cy - s * 0.30), QPointF(cx + s * 0.14, cy - s * 0.14))
        p.drawLine(QPointF(s * 0.22, cy + s * 0.20), QPointF(s * 0.22, cy + s * 0.34))
        p.drawLine(QPointF(s * 0.22, cy + s * 0.34), QPointF(s * 0.78, cy + s * 0.34))
        p.drawLine(QPointF(s * 0.78, cy + s * 0.34), QPointF(s * 0.78, cy + s * 0.20))

    elif name == "check":
        p.drawLine(QPointF(s * 0.26, cy + s * 0.02), QPointF(s * 0.44, cy + s * 0.20))
        p.drawLine(QPointF(s * 0.44, cy + s * 0.20), QPointF(s * 0.76, cy - s * 0.20))

    elif name == "help":
        r = s * 0.34
        p.drawEllipse(QPointF(cx, cy), r, r)
        q = QPainterPath()
        q.moveTo(cx - s * 0.10, cy - s * 0.08)
        q.quadTo(cx - s * 0.10, cy - s * 0.22, cx, cy - s * 0.22)
        q.quadTo(cx + s * 0.14, cy - s * 0.22, cx + s * 0.12, cy - s * 0.05)
        q.quadTo(cx + s * 0.10, cy + s * 0.03, cx, cy + s * 0.06)
        p.drawPath(q)
        p.setBrush(QBrush(c)); p.drawEllipse(QPointF(cx, cy + s * 0.20), s * 0.04, s * 0.04); p.setBrush(Qt.NoBrush)

    elif name == "win_min":
        y = cy + s * 0.02
        p.drawLine(QPointF(s * 0.30, y), QPointF(s * 0.70, y))

    elif name == "win_max":
        r = s * 0.20
        p.drawRect(QRectF(cx - r, cy - r, r * 2, r * 2))

    elif name == "win_restore":
        r = s * 0.17
        off = s * 0.07
        # Back square (top-right)
        p.drawRect(QRectF(cx - r + off, cy - r - off, r * 2, r * 2))
        # Front square (bottom-left), drawn over to look layered
        p.fillRect(QRectF(cx - r - off + 0.5, cy - r + off + 0.5,
                          r * 2 - 1, r * 2 - 1), QColor(BG))
        p.drawRect(QRectF(cx - r - off, cy - r + off, r * 2, r * 2))

    elif name == "win_close":
        d = s * 0.22
        p.drawLine(QPointF(cx - d, cy - d), QPointF(cx + d, cy + d))
        p.drawLine(QPointF(cx + d, cy - d), QPointF(cx - d, cy + d))


def vector_icon(name: str, size: int = 16, color: str = FG_SOFT,
                dpr: float = 2.0) -> QIcon:
    """Draw a crisp toolbar/UI icon with QPainter (no icon-font dependency)."""
    px = max(8, int(size * dpr))
    pm = QPixmap(px, px)
    pm.setDevicePixelRatio(dpr)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    c = QColor(color)
    pen = QPen(c, max(1.4, size * 0.095))
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    _draw_icon(p, name, float(size), c)
    p.end()
    return QIcon(pm)


def vector_pixmap(name: str, size: int = 16, color: str = FG_SOFT,
                  dpr: float = 2.0) -> QPixmap:
    """Like vector_icon but returns a QPixmap for QLabel.setPixmap."""
    return vector_icon(name, size, color, dpr).pixmap(size, size)


def make_icon_label(name: str, size: int = 14, color: str | None = None) -> QLabel:
    """A QLabel showing a vector icon (kept name for call-site compatibility)."""
    lbl = QLabel()
    lbl.setPixmap(vector_pixmap(name, size, color or FG_SOFT))
    lbl.setStyleSheet("background: transparent;")
    return lbl


def glyph_icon(name: str, size: int = 16, color: str = FG_SOFT) -> QIcon:
    """Compatibility shim — now renders a vector icon by name."""
    return vector_icon(name, size, color)


def build_qss() -> str:
	"""Return the full application stylesheet."""
	return f"""
	* {{
		font-family: 'Segoe UI', 'Segoe UI Variable', Arial, sans-serif;
		font-size: 13px;
		color: {FG};
		outline: none;
	}}

	#Root {{
		background: {BG};
		border: 1px solid {BORDER};
		border-radius: 10px;
	}}

	/* ── Title bar ─────────────────────────────────────────────────── */
	#TitleBar {{ background: transparent; }}
	#TitleText {{ font-size: 14px; font-weight: 600; color: {FG}; }}
	#TitleVersion {{ font-size: 12px; color: {MUTED}; }}

	#WinBtn {{
		background: transparent;
		border: none;
		border-radius: 6px;
		color: {FG_SOFT};
		font-size: 14px;
		padding: 0;
	}}
	#WinBtn:hover {{ background: {PANEL_HI}; color: {FG}; }}
	#WinClose:hover {{ background: #e81123; color: white; }}

	/* ── Cards ─────────────────────────────────────────────────────── */
	QFrame[card="true"] {{
		background: {PANEL};
		border: 1px solid {BORDER};
		border-radius: 12px;
	}}

	QLabel[sectionTitle="true"] {{
		color: {MUTED};
		font-size: 11px;
		font-weight: 700;
		letter-spacing: 1px;
	}}
	QLabel[heading="true"] {{ font-size: 14px; font-weight: 600; color: {FG}; }}
	QLabel[muted="true"] {{ color: {MUTED}; }}
	QLabel[soft="true"]  {{ color: {FG_SOFT}; }}

	/* ── Buttons ───────────────────────────────────────────────────── */
	QPushButton {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 8px;
		color: {FG};
		padding: 8px 14px;
	}}
	QPushButton:hover {{ background: {BORDER}; border-color: {BORDER_HI}; }}
	QPushButton:pressed {{ background: {FIELD}; }}
	QPushButton:disabled {{ color: {MUTED}; background: {PANEL}; }}

	QPushButton[accent="true"] {{
		background: {ACCENT};
		color: {ACCENT_FG};
		border: 1px solid {ACCENT};
		font-weight: 700;
	}}
	QPushButton[accent="true"]:hover {{ background: {ACCENT_HOVER}; border-color: {ACCENT_HOVER}; }}
	QPushButton[accent="true"]:pressed {{ background: {ACCENT_ACTIVE}; border-color: {ACCENT_ACTIVE}; }}
	QPushButton[accent="true"]:disabled {{ background: {BORDER}; color: {MUTED}; border-color: {BORDER}; }}

	QPushButton[ghost="true"] {{ background: transparent; border: none; color: {MUTED}; }}
	QPushButton[ghost="true"]:hover {{ color: {FG}; background: {PANEL_HI}; }}

	QPushButton[danger="true"] {{ color: {DANGER}; }}
	QPushButton[danger="true"]:hover {{ background: rgba(240,85,107,0.12); border-color: {DANGER}; }}

	/* ── Inputs ────────────────────────────────────────────────────── */
	QLineEdit, QComboBox, QSpinBox, QPlainTextEdit {{
		background: {FIELD};
		border: 1px solid {BORDER};
		border-radius: 8px;
		padding: 8px 10px;
		color: {FG};
		selection-background-color: {ACCENT};
		selection-color: {ACCENT_FG};
	}}
	QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{ border-color: {ACCENT}; }}
	QComboBox::drop-down {{ border: none; width: 22px; }}
	QComboBox QAbstractItemView {{
		background: {PANEL};
		border: 1px solid {BORDER};
		selection-background-color: {ACCENT};
		selection-color: {ACCENT_FG};
		outline: none;
	}}

	/* ── Slider ────────────────────────────────────────────────────── */
	QSlider::groove:horizontal {{
		height: 6px; border-radius: 3px; background: {FIELD};
	}}
	QSlider::sub-page:horizontal {{ background: {ACCENT}; border-radius: 3px; }}
	QSlider::handle:horizontal {{
		background: {ACCENT}; width: 16px; height: 16px;
		margin: -6px 0; border-radius: 8px; border: 2px solid {BG};
	}}
	QSlider::handle:horizontal:hover {{ background: {ACCENT_HOVER}; }}

	/* ── Scrollbars ────────────────────────────────────────────────── */
	QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
	QScrollBar::handle:vertical {{ background: {BORDER}; border-radius: 5px; min-height: 30px; }}
	QScrollBar::handle:vertical:hover {{ background: {BORDER_HI}; }}
	QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
	QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
	QScrollBar::handle:horizontal {{ background: {BORDER}; border-radius: 5px; min-width: 30px; }}

	QToolTip {{
		background: {PANEL}; color: {FG};
		border: 1px solid {BORDER_HI}; padding: 6px 8px; border-radius: 6px;
	}}

	/* ── Drag & drop zone ──────────────────────────────────────────── */
	#DropZone {{
		background: {FIELD};
		border: 2px dashed {BORDER_HI};
		border-radius: 10px;
	}}
	#DropZone[hover="true"] {{
		border-color: {ACCENT};
		background: rgba(253,195,1,0.06);
	}}
	#DropHint {{ color: {MUTED}; }}
	#DropBrowse {{ color: {ACCENT}; font-weight: 600; }}

	/* ── Recent preset list rows ───────────────────────────────────── */
	#PresetRow {{
		background: transparent;
		border: 1px solid transparent;
		border-radius: 8px;
	}}
	#PresetRow:hover {{ background: {PANEL_HI}; border-color: {BORDER}; }}
	#PresetIcon {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 7px;
		color: {ACCENT};
	}}
	#PresetName {{ font-size: 13px; font-weight: 600; color: {FG}; }}
	#PresetDate {{ font-size: 11px; color: {MUTED}; }}

	/* ── Numbered step badge ───────────────────────────────────────── */
	#StepBadge {{
		background: {ACCENT};
		color: {ACCENT_FG};
		border-radius: 9px;
		font-size: 11px;
		font-weight: 700;
	}}
	#StepTitle {{ font-size: 11px; font-weight: 700; letter-spacing: 1px; color: {FG_SOFT}; }}
	#OptionalTag {{
		font-size: 9px; font-weight: 700; letter-spacing: 1px;
		color: {MUTED}; background: {PANEL_HI};
		border: 1px solid {BORDER}; border-radius: 4px;
		padding: 1px 5px;
	}}

	/* ── Char counter ──────────────────────────────────────────────── */
	#CharCount {{ color: {MUTED}; font-size: 11px; }}

	/* ── Big watermark text field ──────────────────────────────────── */
	#WmText {{ font-size: 22px; font-weight: 600; padding: 12px 14px; }}

	/* ── Color hex chip ────────────────────────────────────────────── */
	#HexValue {{
		background: {FIELD}; border: 1px solid {BORDER}; border-radius: 8px;
		padding: 7px 10px; color: {FG};
	}}
	#ColorSwatch {{ border: 1px solid {BORDER}; border-radius: 6px; }}

	/* ── Slider value pill ─────────────────────────────────────────── */
	#TransPill {{
		background: {ACCENT}; color: {ACCENT_FG};
		border: 1px solid {ACCENT}; border-radius: 6px;
		font-weight: 700; padding: 3px 8px;
		selection-background-color: {ACCENT_FG};
		selection-color: {ACCENT};
	}}
	#TransPill:focus {{ border-color: {ACCENT_FG}; }}
	#TransEnd {{ color: {MUTED}; font-size: 11px; }}

	/* ── Sub-card (inner cells for color/trans/preset) ─────────────── */
	QFrame[subcard="true"] {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 10px;
	}}

	/* ── Advanced options bar ──────────────────────────────────────── */
	#AdvancedBar {{
		background: {PANEL};
		border: 1px solid {BORDER};
		border-radius: 12px;
	}}
	#AdvancedBar:hover {{ border-color: {BORDER_HI}; background: {PANEL_HI}; }}
	#AdvTitle {{ font-size: 13px; font-weight: 600; color: {FG}; }}
	#AdvSub {{ font-size: 11px; color: {MUTED}; }}

	/* ── Right column: toggle rows & status ────────────────────────── */
	#ToggleRow {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 10px;
	}}
	#ToggleLabel {{ font-size: 13px; color: {FG}; }}
	#ToggleSub {{ font-size: 11px; color: {MUTED}; }}

	#StatusBadge {{
		background: rgba(63,185,80,0.14);
		border-radius: 18px;
	}}
	#StatusReady {{ font-size: 14px; font-weight: 700; color: {FG}; }}
	#StatusReadySub {{ font-size: 11px; color: {MUTED}; }}
	#StatRowKey {{ color: {FG_SOFT}; font-size: 13px; }}
	#StatRowVal {{ color: {FG}; font-size: 13px; font-weight: 600; }}
	QFrame[divider="true"] {{ background: {BORDER}; max-height: 1px; min-height: 1px; border: none; }}

	#HelpLink {{ color: {MUTED}; }}
	#HelpLink:hover {{ color: {ACCENT}; }}

	/* ── Preview toolbar ───────────────────────────────────────────── */
	#PvTool {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 7px;
		color: {FG_SOFT};
		padding: 0;
	}}
	#PvTool:hover {{ background: {BORDER}; color: {FG}; }}
	#PvTool:checked {{ background: {ACCENT}; color: {ACCENT_FG}; border-color: {ACCENT}; }}
	#PvZoom {{
		background: {PANEL_HI};
		border: 1px solid {BORDER};
		border-radius: 7px;
		color: {FG};
		padding: 4px 6px;
	}}
	#PvZoom::drop-down {{ width: 14px; border: none; }}

	/* ── Status bar ────────────────────────────────────────────────── */
	#AppStatusBar {{ background: transparent; }}
	"""


def _win_round_and_shadow(win: QWidget) -> None:
	"""Apply Windows 11 rounded corners + drop shadow to a frameless window."""
	if sys.platform != "win32":
		return
	try:
		from ctypes import windll, byref, c_int, sizeof
		hwnd = int(win.winId())
		# DWMWA_WINDOW_CORNER_PREFERENCE = 33, DWMWCP_ROUND = 2
		windll.dwmapi.DwmSetWindowAttribute(hwnd, 33, byref(c_int(2)), sizeof(c_int))
		# DWMWA_USE_IMMERSIVE_DARK_MODE = 20
		windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, byref(c_int(1)), sizeof(c_int))
		# Extend a 1px frame so DWM paints the native drop shadow.
		from ctypes import Structure
		class MARGINS(Structure):
			_fields_ = [("l", c_int), ("r", c_int), ("t", c_int), ("b", c_int)]
		m = MARGINS(1, 1, 1, 1)
		windll.dwmapi.DwmExtendFrameIntoClientArea(hwnd, byref(m))
	except Exception:
		pass


class DropZone(QFrame):
	"""Drag-and-drop target that also acts as a click-to-browse button."""

	filesDropped = Signal(list)
	browseClicked = Signal()

	def __init__(self, parent=None):
		super().__init__(parent)
		self.setObjectName("DropZone")
		self.setAcceptDrops(True)
		self.setCursor(Qt.PointingHandCursor)
		self.setMinimumHeight(150)

		v = QVBoxLayout(self)
		v.setContentsMargins(12, 18, 12, 18)
		v.setSpacing(8)
		v.setAlignment(Qt.AlignCenter)

		glyph = make_icon_label(ICON_ADDTO, 30, ACCENT)
		glyph.setAlignment(Qt.AlignCenter)
		v.addWidget(glyph, 0, Qt.AlignCenter)

		hint = QLabel("Drag & drop files here")
		hint.setObjectName("DropHint")
		hint.setAlignment(Qt.AlignCenter)
		v.addWidget(hint)

		browse_row = QLabel(
			'or click to <span style="color:%s;font-weight:600">browse</span>' % ACCENT)
		browse_row.setObjectName("DropHint")
		browse_row.setAlignment(Qt.AlignCenter)
		v.addWidget(browse_row)

	def _set_hover(self, on: bool):
		self.setProperty("hover", "true" if on else "false")
		self.style().unpolish(self)
		self.style().polish(self)

	def mouseReleaseEvent(self, e):
		if e.button() == Qt.LeftButton:
			self.browseClicked.emit()

	def dragEnterEvent(self, e):
		if e.mimeData().hasUrls():
			e.acceptProposedAction()
			self._set_hover(True)

	def dragLeaveEvent(self, e):
		self._set_hover(False)

	def dropEvent(self, e):
		self._set_hover(False)
		paths = []
		for url in e.mimeData().urls():
			p = url.toLocalFile()
			if p and os.path.splitext(p)[1].lower() in ALL_SUPPORTED:
				paths.append(os.path.normpath(p))
		if paths:
			self.filesDropped.emit(paths)


class PresetRow(QFrame):
	"""A single recent-preset row: icon + name + timestamp + overflow menu."""

	clicked = Signal(str)
	deleteRequested = Signal(str)

	def __init__(self, name: str, subtitle: str, color: str | None = None,
				 data: dict | None = None, parent=None):
		super().__init__(parent)
		self.setObjectName("PresetRow")
		self.setCursor(Qt.PointingHandCursor)
		self._name = name

		# Rich hover tooltip with the full, ungranulated detail.
		self.setToolTip(self._build_tooltip(name, data or {}, color))

		h = QHBoxLayout(self)
		h.setContentsMargins(8, 7, 6, 7)
		h.setSpacing(10)

		icon = QLabel()
		icon.setObjectName("PresetIcon")
		icon.setPixmap(vector_pixmap(ICON_TEXT, 15, color or ACCENT))
		icon.setFixedSize(32, 32)
		icon.setAlignment(Qt.AlignCenter)
		if color:
			# Tint the tile background subtly; the glyph already uses the colour.
			icon.setStyleSheet(
				f"#PresetIcon {{ background: {PANEL_HI}; border: 1px solid {BORDER};"
				f" border-radius: 7px; }}")
		h.addWidget(icon)

		col = QVBoxLayout()
		col.setSpacing(1)
		name_lbl = QLabel(name)
		name_lbl.setObjectName("PresetName")
		sub_lbl = QLabel(subtitle)
		sub_lbl.setObjectName("PresetDate")
		col.addWidget(name_lbl)
		col.addWidget(sub_lbl)
		h.addLayout(col, 1)

		more = QPushButton()
		more.setIcon(vector_icon(ICON_MORE, 16, MUTED))
		more.setIconSize(QSize(16, 16))
		more.setProperty("ghost", True)
		more.setFixedSize(26, 26)
		more.setCursor(Qt.PointingHandCursor)
		more.clicked.connect(self._show_menu)
		h.addWidget(more)
		self._more = more

	def _build_tooltip(self, name: str, data: dict, color: str | None) -> str:
		"""Rich HTML tooltip exposing the full preset detail."""
		text = data.get("text", "") or ""
		hexc = (color or data.get("color", "#A6A6A6")).upper()
		trans = data.get("transparency")
		trans_str = f"{int(trans)}%" if isinstance(trans, (int, float)) else "—"
		# Escape angle brackets in user text.
		safe = (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
		swatch = (
			f'<span style="background:{hexc};">&nbsp;&nbsp;&nbsp;</span>')

		rows = ""
		# Only show the Text row when it adds something beyond the name heading.
		if text and text.strip() != name.strip():
			rows += (
				f'<tr><td style="color:#8b949e">Text</td>'
				f'<td>{safe}</td></tr>')
		rows += (
			f'<tr><td style="color:#8b949e">Colour</td>'
			f'<td>{swatch} {hexc}</td></tr>'
			f'<tr><td style="color:#8b949e">Transparency</td>'
			f'<td>{trans_str}</td></tr>')
		return (
			f'<div style="min-width:180px">'
			f'<b>{name}</b><br>'
			f'<table cellspacing="0" cellpadding="2">{rows}</table>'
			f'</div>'
		)

	def _show_menu(self):
		menu = QMenu(self)
		act_load = menu.addAction("Load preset")
		act_del = menu.addAction("Delete preset")
		chosen = menu.exec(self._more.mapToGlobal(QPoint(0, self._more.height())))
		if chosen == act_load:
			self.clicked.emit(self._name)
		elif chosen == act_del:
			self.deleteRequested.emit(self._name)

	def mouseReleaseEvent(self, e):
		if e.button() == Qt.LeftButton:
			self.clicked.emit(self._name)


class ToggleSwitch(QWidget):
	"""An animated iOS-style on/off toggle."""

	toggled = Signal(bool)

	def __init__(self, checked: bool = False, parent=None):
		super().__init__(parent)
		self._checked = checked
		self._offset = 1.0 if checked else 0.0
		self.setFixedSize(42, 24)
		self.setCursor(Qt.PointingHandCursor)
		self._anim = QPropertyAnimation(self, b"offset", self)
		self._anim.setDuration(140)

	def isChecked(self) -> bool:
		return self._checked

	def setChecked(self, value: bool):
		value = bool(value)
		if value == self._checked:
			return
		self._checked = value
		self._anim.stop()
		self._anim.setStartValue(self._offset)
		self._anim.setEndValue(1.0 if value else 0.0)
		self._anim.start()
		self.toggled.emit(value)

	def mouseReleaseEvent(self, e):
		if e.button() == Qt.LeftButton:
			self.setChecked(not self._checked)

	def _get_offset(self):
		return self._offset

	def _set_offset(self, v):
		self._offset = v
		self.update()

	offset = Property(float, _get_offset, _set_offset)

	def paintEvent(self, _e):
		p = QPainter(self)
		p.setRenderHint(QPainter.Antialiasing, True)
		r = self.rect().adjusted(1, 1, -1, -1)
		track_off = QColor(BORDER)
		track_on = QColor(ACCENT)
		track = QColor(
			int(track_off.red()   + (track_on.red()   - track_off.red())   * self._offset),
			int(track_off.green() + (track_on.green() - track_off.green()) * self._offset),
			int(track_off.blue()  + (track_on.blue()  - track_off.blue())  * self._offset),
		)
		p.setPen(Qt.NoPen)
		p.setBrush(QBrush(track))
		radius = r.height() / 2
		p.drawRoundedRect(r, radius, radius)
		d = r.height() - 6
		x = r.left() + 3 + (r.width() - d - 6) * self._offset
		knob = QColor(ACCENT_FG) if self._checked else QColor("#e6edf3")
		p.setBrush(QBrush(knob))
		p.drawEllipse(QRectF(x, r.top() + 3, d, d))
		p.end()


def _friendly_error(exc: Exception) -> str:
	"""Translate noisy COM / backend errors into plain language for the user."""
	msg = str(exc)
	low = msg.lower()
	# PowerPoint/Word edit-blocked: Restrict Editing, read-only-by-author,
	# rights management, or a sensitivity label that enforces protection.
	if ("access denied" in low or "enough privileges" in low
			or "-2147024891" in msg or "0x80070005" in low):
		return ("file is protected or restricted for editing "
				"(read-only by author, Restrict Editing, or a sensitivity label)")
	if "password" in low:
		return "file is password-protected"
	if "being used by another process" in low or "in use" in low:
		return "file is open in another program — close it and retry"
	# Trim a very long COM traceback string down to something readable.
	if len(msg) > 140:
		msg = msg[:137] + "…"
	return msg


class ProcessWorker(QObject):
	"""Runs the watermark backend over a list of files on a worker thread.

	Lives in its own QThread; COM is initialised per-thread inside the backend
	modules. Communicates with the UI exclusively via signals.
	"""

	progress   = Signal(int, int, str)   # done, total, current basename
	fileDone   = Signal(str)             # output path of a finished file
	encode     = Signal(float)           # video encode seconds processed
	finished   = Signal(int, list, str)  # done_count, errors, last_output_path
	needFfmpeg = Signal()                # ffmpeg missing for a video job

	def __init__(self, files, text, color_rgb, transparency, export_pdf):
		super().__init__()
		self._files = list(files)
		self._text = text
		self._color_rgb = color_rgb
		self._transparency = transparency
		self._export_pdf = export_pdf
		self._cancel = False

	def cancel(self):
		self._cancel = True

	def run(self):
		done, errors, last_out = 0, [], None
		total = len(self._files)
		for i, path in enumerate(self._files, 1):
			if self._cancel:
				break
			self.progress.emit(i, total, os.path.basename(path))
			kind = file_kind(path)
			try:
				if kind == "ppt":
					out = add_watermark(path, self._text, color_rgb=self._color_rgb,
										transparency=self._transparency,
										export_pdf=self._export_pdf)
				elif kind == "word":
					out = add_word_watermark(path, self._text, color_rgb=self._color_rgb,
											 transparency=self._transparency,
											 export_pdf=self._export_pdf)
				elif kind == "video":
					out = add_video_watermark(path, self._text, color_rgb=self._color_rgb,
											  transparency=self._transparency,
											  progress_cb=lambda s, _t: self.encode.emit(s))
				else:
					raise ValueError(f"Unsupported file type: {path}")
				last_out = out
				done += 1
				self.fileDone.emit(out)
			except FfmpegNotReadyError:
				self.needFfmpeg.emit()
				errors.append(f"{os.path.basename(path)}: ffmpeg not available")
			except Exception as exc:  # noqa: BLE001
				errors.append(f"{os.path.basename(path)}: {_friendly_error(exc)}")
		self.finished.emit(done, errors, last_out or "")


class _FfmpegDownloadWorker(QObject):
	"""Downloads ffmpeg on a worker thread, reporting MB progress."""

	progress = Signal(float)        # MB downloaded
	done     = Signal(bool, str)    # success, error message

	def run(self):
		try:
			def cb(done_bytes, _total):
				self.progress.emit(done_bytes / 1_048_576)
			download_ffmpeg(progress_cb=cb)
			self.done.emit(True, "")
		except Exception as exc:  # noqa: BLE001
			self.done.emit(False, str(exc))


class TitleBar(QFrame):
	"""Custom draggable title bar with window controls."""

	def __init__(self, parent: "WatermarkLabX"):
		super().__init__(parent)
		self.setObjectName("TitleBar")
		self.setFixedHeight(46)
		self._win = parent
		self._drag_offset: QPoint | None = None

		lay = QHBoxLayout(self)
		lay.setContentsMargins(14, 0, 8, 0)
		lay.setSpacing(10)

		logo = QLabel()
		if os.path.isfile(ICON_PNG):
			pm = QPixmap(ICON_PNG).scaled(24, 24, Qt.KeepAspectRatio,
										  Qt.SmoothTransformation)
			logo.setPixmap(pm)
		logo.setFixedSize(26, 26)
		lay.addWidget(logo)

		title = QLabel("Watermark Lab")
		title.setObjectName("TitleText")
		lay.addWidget(title)

		ver = QLabel(f"v{APP_VERSION}")
		ver.setObjectName("TitleVersion")
		lay.addWidget(ver)

		lay.addStretch(1)

		self.btn_min   = self._win_button("win_min", "WinMin")
		self.btn_max   = self._win_button("win_max", "WinMax")
		self.btn_close = self._win_button("win_close", "WinClose")
		self.btn_min.clicked.connect(self._win.showMinimized)
		self.btn_max.clicked.connect(self._toggle_max)
		self.btn_close.clicked.connect(self._win.close)
		for b in (self.btn_min, self.btn_max, self.btn_close):
			lay.addWidget(b)

	def _win_button(self, icon_name: str, name: str) -> QPushButton:
		b = QPushButton()
		b.setObjectName("WinClose" if name == "WinClose" else "WinBtn")
		b.setProperty("winbtn", name)
		col = FG_SOFT
		b.setIcon(vector_icon(icon_name, 12, col))
		b.setIconSize(QSize(12, 12))
		b.setFixedSize(38, 30)
		b.setCursor(Qt.PointingHandCursor)
		return b

	def _toggle_max(self):
		if self._win.isMaximized():
			self._win.showNormal()
		else:
			self._win.showMaximized()
		# Swap the glyph to reflect the new state.
		self.btn_max.setIcon(vector_icon(
			"win_restore" if self._win.isMaximized() else "win_max", 12, FG_SOFT))

	# Drag-to-move
	def mousePressEvent(self, e):
		if e.button() == Qt.LeftButton and not self._win.isMaximized():
			self._drag_offset = e.globalPosition().toPoint() - self._win.frameGeometry().topLeft()
			e.accept()

	def mouseMoveEvent(self, e):
		if self._drag_offset is not None and e.buttons() & Qt.LeftButton:
			self._win.move(e.globalPosition().toPoint() - self._drag_offset)
			e.accept()

	def mouseReleaseEvent(self, e):
		self._drag_offset = None

	def mouseDoubleClickEvent(self, e):
		self._toggle_max()


class WatermarkLabX(QMainWindow):
	"""Main application window — frameless with custom chrome."""

	_RESIZE_MARGIN = 7

	def __init__(self):
		super().__init__()
		self.setWindowTitle(f"Watermark Lab  v{APP_VERSION}")
		if os.path.isfile(ICON_ICO):
			self.setWindowIcon(QIcon(ICON_ICO))
		self.setWindowFlag(Qt.FramelessWindowHint, True)
		self.setMinimumSize(1240, 760)
		self.resize(1320, 820)

		self._resize_edge = Qt.Edges()
		self.setMouseTracking(True)

		# ── Application state ──────────────────────────────────────────
		self._files: list[str] = []          # current selection (1 file, or a folder batch)
		self._presets: dict = {}             # name -> preset dict
		self._files_processed = 0
		self._last_output_path: str | None = None
		self._color_hex = "#A6A6A6"          # current watermark colour
		self._MAX_CHARS = 100
		self._export_pdf_only = False        # "Export PDF (PowerPoint/Word only)" toggle
		self._open_after = True              # "Open file(s) after processing" toggle
		self._preview_enabled = False        # live preview off by default (COM is costly)
		self._proc_thread = None             # active processing QThread
		self._proc_worker = None
		self._dl_thread = None               # ffmpeg download QThread

		root = QFrame()
		root.setObjectName("Root")
		root.setMouseTracking(True)
		self.setCentralWidget(root)

		outer = QVBoxLayout(root)
		outer.setContentsMargins(1, 1, 1, 1)
		outer.setSpacing(0)

		self.title_bar = TitleBar(self)
		outer.addWidget(self.title_bar)

		# ── Body: three columns ────────────────────────────────────────
		body = QWidget()
		body.setMouseTracking(True)
		body_lay = QHBoxLayout(body)
		body_lay.setContentsMargins(14, 6, 14, 14)
		body_lay.setSpacing(14)

		self.left_col   = self._build_left_column()
		self.center_col = self._build_center_column()
		self.right_col  = self._build_right_column()

		body_lay.addWidget(self.left_col)
		body_lay.addWidget(self.center_col, 1)
		body_lay.addWidget(self.right_col)
		outer.addWidget(body, 1)

		self.statusbar_widget = self._build_statusbar()
		outer.addWidget(self.statusbar_widget)

		# Populate dynamic content
		self._refresh_recent()
		self._refresh_preset_combo()
		self._init_preview()
		self._on_files_changed()

	# ── Column builders (placeholder content filled in later steps) ─────
	def _card(self, title: str | None = None) -> QFrame:
		card = QFrame()
		card.setProperty("card", True)
		card.setMouseTracking(True)
		v = QVBoxLayout(card)
		v.setContentsMargins(16, 14, 16, 16)
		v.setSpacing(12)
		if title:
			lbl = QLabel(title.upper())
			lbl.setProperty("sectionTitle", True)
			v.addWidget(lbl)
		return card

	def _build_left_column(self) -> QWidget:
		col = QWidget()
		col.setFixedWidth(264)
		col.setMouseTracking(True)
		v = QVBoxLayout(col)
		v.setContentsMargins(0, 0, 0, 0)
		v.setSpacing(14)

		# ── File Selection card ────────────────────────────────────────
		file_card = self._card("File Selection")
		fv = file_card.layout()

		self.drop_zone = DropZone()
		self.drop_zone.setToolTip(
			"Supported: PowerPoint (.pptx .ppt), Word (.docx .doc),\n"
			"Video (.mp4 .mov .m4v .mkv .avi .webm)")
		self.drop_zone.browseClicked.connect(self._pick_file)
		self.drop_zone.filesDropped.connect(self._set_files)
		fv.addWidget(self.drop_zone)

		btn_files = self._icon_button("Browse File…", ICON_FOLDER, accent=True)
		btn_files.clicked.connect(self._pick_file)
		fv.addWidget(btn_files)

		btn_folder = self._icon_button("Browse Folder…", ICON_FOLDER)
		btn_folder.clicked.connect(self._pick_folder)
		fv.addWidget(btn_folder)

		v.addWidget(file_card)

		# ── Recent Presets card ────────────────────────────────────────
		recent_card = self._card("Recent Presets")
		rv = recent_card.layout()

		scroll = QScrollArea()
		scroll.setWidgetResizable(True)
		scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
		scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
		scroll.setFrameShape(QFrame.NoFrame)
		scroll.setStyleSheet("background: transparent;")
		holder = QWidget()
		holder.setStyleSheet("background: transparent;")
		self.recent_list_lay = QVBoxLayout(holder)
		# Right margin keeps row hover borders off the scrollbar / card edge.
		self.recent_list_lay.setContentsMargins(0, 0, 4, 0)
		self.recent_list_lay.setSpacing(4)
		self.recent_list_lay.addStretch(1)
		scroll.setWidget(holder)
		rv.addWidget(scroll, 1)
		v.addWidget(recent_card, 1)

		# Settings button hidden for the initial release.

		return col

	def _build_center_column(self) -> QWidget:
		col = QWidget()
		col.setMouseTracking(True)
		v = QVBoxLayout(col)
		v.setContentsMargins(0, 0, 0, 0)
		v.setSpacing(14)

		# ── Watermark Text card (step 1) ───────────────────────────────
		text_card = self._card()
		tv = text_card.layout()
		tv.addLayout(self._step_header(
			1, "WATERMARK TEXT", info=True,
			tip="The text stamped diagonally across every page, slide or video "
				"frame. Long text wraps automatically. Up to 100 characters."))

		self.wm_text = QLineEdit("CONFIDENTIAL")
		self.wm_text.setObjectName("WmText")
		self.wm_text.setMaxLength(self._MAX_CHARS)
		self.wm_text.textChanged.connect(self._on_text_changed)
		# Char counter overlaid at the right edge
		self.char_count = QLabel(parent=self.wm_text)
		self.char_count.setObjectName("CharCount")
		self.char_count.setAttribute(Qt.WA_TransparentForMouseEvents, True)
		self.wm_text.installEventFilter(self)
		tv.addWidget(self.wm_text)
		v.addWidget(text_card)

		# ── Preview card ───────────────────────────────────────────────
		self.preview_card = self._card()
		pv = self.preview_card.layout()
		preview_header = QHBoxLayout()
		title = QLabel("WATERMARK PREVIEW")
		title.setProperty("sectionTitle", True)
		preview_header.addWidget(title)
		preview_header.addStretch(1)
		self._build_preview_toolbar(preview_header)
		pv.addLayout(preview_header)

		self.preview_host = QFrame()
		self.preview_host.setObjectName("PreviewHost")
		self.preview_host.setStyleSheet(
			f"#PreviewHost {{ background: {PREVIEW_BG}; border: 1px solid {BORDER};"
			f" border-radius: 8px; }}")
		self.preview_host.setMinimumHeight(300)
		ph_lay = QVBoxLayout(self.preview_host)
		ph_lay.setContentsMargins(1, 1, 1, 1)
		self.preview_canvas = PreviewCanvas()
		self.preview_canvas.set_placeholder("Preview is off — click the Preview button to turn it on")
		self.preview_canvas.zoomChanged.connect(self._on_zoom_changed)
		ph_lay.addWidget(self.preview_canvas)
		pv.addWidget(self.preview_host, 1)

		v.addWidget(self.preview_card, 1)

		# ── Row of sub-cards: Color / Transparency (primary) + Preset (optional)
		row = QHBoxLayout()
		row.setSpacing(14)
		# Give the primary controls more room; the optional Preset recedes.
		row.addWidget(self._build_color_cell(), 6)
		row.addWidget(self._build_transparency_cell(), 5)
		row.addWidget(self._build_preset_cell(), 4)
		v.addLayout(row)

		# Advanced options bar hidden for the initial release.

		self._update_char_count()
		return col

	# ── Center-column sub-builders ─────────────────────────────────────
	def _step_header(self, num: int, title: str, info: bool = False,
					 tip: str | None = None) -> QHBoxLayout:
		h = QHBoxLayout()
		h.setSpacing(8)
		badge = QLabel(str(num))
		badge.setObjectName("StepBadge")
		badge.setFixedSize(18, 18)
		badge.setAlignment(Qt.AlignCenter)
		h.addWidget(badge)
		lbl = QLabel(title)
		lbl.setObjectName("StepTitle")
		h.addWidget(lbl)
		if info and tip:
			info_lbl = make_icon_label(ICON_INFO, 12, MUTED)
			info_lbl.setToolTip(tip)
			info_lbl.setCursor(Qt.WhatsThisCursor)
			h.addWidget(info_lbl)
		h.addStretch(1)
		return h

	def _optional_header(self, title: str, tag: str | None, tip: str | None = None) -> QHBoxLayout:
		"""Header for a non-numbered section (e.g. Preset).

		Has no step number so it doesn't read as a required part of the
		1-2-3 watermarking flow. An optional small tag pill can be shown.
		"""
		h = QHBoxLayout()
		h.setSpacing(8)
		lbl = QLabel(title)
		lbl.setObjectName("StepTitle")
		h.addWidget(lbl)
		if tag:
			pill = QLabel(tag)
			pill.setObjectName("OptionalTag")
			h.addWidget(pill)
		if tip:
			info_lbl = make_icon_label(ICON_INFO, 12, MUTED)
			info_lbl.setToolTip(tip)
			info_lbl.setCursor(Qt.WhatsThisCursor)
			h.addWidget(info_lbl)
		h.addStretch(1)
		return h

	def _build_color_cell(self) -> QFrame:
		cell = QFrame()
		cell.setProperty("subcard", True)
		v = QVBoxLayout(cell)
		v.setContentsMargins(14, 12, 14, 14)
		v.setSpacing(10)
		v.addLayout(self._step_header(
			2, "TEXT COLOR", info=True,
			tip="Colour of the watermark text. Pick a swatch, type a hex value, "
				"or use the eyedropper for a custom colour."))

		# Hex chip row: swatch + hex value + eyedropper
		chip = QHBoxLayout()
		chip.setSpacing(8)
		self.color_swatch = QLabel()
		self.color_swatch.setObjectName("ColorSwatch")
		self.color_swatch.setFixedSize(28, 28)
		self._apply_swatch_color(self._color_hex)
		chip.addWidget(self.color_swatch)

		self.hex_value = QLabel(self._color_hex)
		self.hex_value.setObjectName("HexValue")
		chip.addWidget(self.hex_value, 1)

		eyedrop = QPushButton()
		eyedrop.setIcon(vector_icon(ICON_EYEDROP, 16, FG_SOFT))
		eyedrop.setIconSize(QSize(18, 18))
		eyedrop.setProperty("ghost", True)
		eyedrop.setFixedSize(32, 32)
		eyedrop.setCursor(Qt.PointingHandCursor)
		eyedrop.setToolTip("Pick a custom colour")
		eyedrop.clicked.connect(self._pick_color)
		chip.addWidget(eyedrop)
		v.addLayout(chip)

		# Two rows of preset swatches
		grid = QGridLayout()
		grid.setSpacing(6)
		palette = [
			"#A6A6A6", "#1b1b1b", "#E53935", "#FB8C00", "#FDD835", "#43A047",
			"#1E88E5", "#00ACC1", "#26A69A", "#5E35B1", "#8E24AA", "#EC407A",
		]
		for i, hexc in enumerate(palette):
			sw = QPushButton()
			sw.setFixedSize(26, 22)
			sw.setCursor(Qt.PointingHandCursor)
			selected = hexc.lower() == self._color_hex.lower()
			sw.setStyleSheet(self._swatch_btn_css(hexc, selected))
			sw.clicked.connect(lambda _=False, c=hexc: self._set_color(c))
			grid.addWidget(sw, i // 6, i % 6)
		self._palette_buttons = grid
		v.addLayout(grid)
		v.addStretch(1)
		return cell

	def _swatch_btn_css(self, hexc: str, selected: bool) -> str:
		border = ACCENT if selected else BORDER
		width = 2 if selected else 1
		return (f"background: {hexc}; border: {width}px solid {border};"
				f"border-radius: 5px;")

	def _build_transparency_cell(self) -> QFrame:
		cell = QFrame()
		cell.setProperty("subcard", True)
		v = QVBoxLayout(cell)
		v.setContentsMargins(14, 12, 14, 14)
		v.setSpacing(10)

		head = self._step_header(
			3, "TRANSPARENCY", info=True,
			tip="How see-through the watermark is. 0% is solid; 100% is invisible. "
				"70% is a good default that stays readable without hiding content.")
		self.trans_pill = QLineEdit("70%")
		self.trans_pill.setObjectName("TransPill")
		self.trans_pill.setAlignment(Qt.AlignCenter)
		self.trans_pill.setMaxLength(4)
		self.trans_pill.setFixedWidth(58)
		self.trans_pill.setToolTip("Type a value from 0 to 100")
		self.trans_pill.editingFinished.connect(self._on_trans_pill_edited)
		head.addWidget(self.trans_pill)
		v.addLayout(head)

		v.addStretch(1)
		self.trans_slider = QSlider(Qt.Horizontal)
		self.trans_slider.setRange(0, 100)
		self.trans_slider.setValue(70)
		self.trans_slider.valueChanged.connect(self._on_trans_changed)
		v.addWidget(self.trans_slider)

		ends = QHBoxLayout()
		lo = QLabel("0%"); lo.setObjectName("TransEnd")
		hi = QLabel("100%"); hi.setObjectName("TransEnd")
		ends.addWidget(lo)
		ends.addStretch(1)
		ends.addWidget(hi)
		v.addLayout(ends)
		v.addStretch(1)
		return cell

	def _build_preset_cell(self) -> QFrame:
		cell = QFrame()
		cell.setProperty("subcard", True)
		v = QVBoxLayout(cell)
		v.setContentsMargins(14, 12, 14, 14)
		v.setSpacing(10)
		v.addLayout(self._optional_header(
			"PRESET", None,
			tip="Optional. Save the current text, colour and transparency as a "
				"named preset to reuse later — or load one from the list. You "
				"can watermark without ever touching presets."))

		self.preset_combo = QComboBox()
		self.preset_combo.setMinimumHeight(34)
		self.preset_combo.currentTextChanged.connect(self._on_preset_combo)
		v.addWidget(self.preset_combo)

		btns = QHBoxLayout()
		btns.setSpacing(8)
		self.btn_save_preset = QPushButton("  Save")
		self.btn_save_preset.setIcon(glyph_icon(ICON_SAVE, 14, FG_SOFT))
		self.btn_save_preset.setCursor(Qt.PointingHandCursor)
		self.btn_save_preset.clicked.connect(self._save_preset)
		btns.addWidget(self.btn_save_preset, 1)

		self.btn_del_preset = QPushButton("  Delete")
		self.btn_del_preset.setProperty("danger", True)
		self.btn_del_preset.setIcon(glyph_icon(ICON_DELETE, 14, DANGER))
		self.btn_del_preset.setCursor(Qt.PointingHandCursor)
		self.btn_del_preset.clicked.connect(self._delete_current_preset)
		btns.addWidget(self.btn_del_preset, 1)
		v.addLayout(btns)
		v.addStretch(1)
		return cell

	def _build_advanced_bar(self) -> QFrame:
		bar = QFrame()
		bar.setObjectName("AdvancedBar")
		bar.setCursor(Qt.PointingHandCursor)
		bar.setMinimumHeight(58)
		h = QHBoxLayout(bar)
		h.setContentsMargins(16, 10, 16, 10)
		h.setSpacing(12)

		gear = make_icon_label(ICON_SETTINGS, 18, ACCENT)
		h.addWidget(gear)
		col = QVBoxLayout()
		col.setSpacing(1)
		t = QLabel("ADVANCED OPTIONS"); t.setObjectName("AdvTitle")
		s = QLabel("Font, Size, Position, Angle, Tiling, and more…"); s.setObjectName("AdvSub")
		col.addWidget(t); col.addWidget(s)
		h.addLayout(col, 1)
		chevron = make_icon_label(ICON_CHEVRON, 16, MUTED)
		h.addWidget(chevron)

		bar.mouseReleaseEvent = lambda e: self._toggle_advanced()
		return bar

	# ── Preview toolbar & controller wiring ────────────────────────────
	def _pv_button(self, icon: QIcon, tip: str, checkable: bool = False) -> QPushButton:
		b = QPushButton()
		b.setObjectName("PvTool")
		b.setIcon(icon)
		b.setIconSize(QSize(16, 16))
		b.setFixedSize(30, 28)
		b.setCursor(Qt.PointingHandCursor)
		b.setToolTip(tip)
		b.setCheckable(checkable)
		return b

	def _build_preview_toolbar(self, header: QHBoxLayout) -> None:
		group = QHBoxLayout()
		group.setSpacing(6)

		# Preview on/off (default off — gates the costly COM/render pipeline).
		self._eye_on_icon  = vector_icon("eye", 16, FG)
		self._eye_off_icon = vector_icon("eye_off", 16, FG_SOFT)
		self.pv_toggle = self._pv_button(self._eye_off_icon, "Turn preview on", checkable=True)
		self.pv_toggle.setChecked(self._preview_enabled)
		self.pv_toggle.toggled.connect(self._on_preview_toggle)
		group.addWidget(self.pv_toggle)

		group.addSpacing(6)

		self.pv_zoom_out = self._pv_button(vector_icon("minus", 16, FG_SOFT), "Zoom out")
		self.pv_zoom_out.clicked.connect(lambda: self.preview_canvas.zoom_out())
		group.addWidget(self.pv_zoom_out)

		self.pv_zoom_in = self._pv_button(vector_icon("plus", 16, FG_SOFT), "Zoom in")
		self.pv_zoom_in.clicked.connect(lambda: self.preview_canvas.zoom_in())
		group.addWidget(self.pv_zoom_in)

		self.pv_zoom_combo = QComboBox()
		self.pv_zoom_combo.setObjectName("PvZoom")
		self.pv_zoom_combo.setFixedHeight(28)
		self.pv_zoom_combo.addItems(["50%", "75%", "100%", "125%", "150%", "200%"])
		self.pv_zoom_combo.setCurrentText("100%")
		self.pv_zoom_combo.setEditable(False)
		self.pv_zoom_combo.activated.connect(self._on_zoom_combo)
		group.addWidget(self.pv_zoom_combo)

		group.addSpacing(6)

		self.pv_full = self._pv_button(vector_icon("fit", 16, FG_SOFT), "Fit to window")
		self.pv_full.clicked.connect(lambda: self.preview_canvas.fit())
		group.addWidget(self.pv_full)

		# View controls are inert until preview is switched on.
		self._pv_view_controls = [
			self.pv_zoom_out, self.pv_zoom_in, self.pv_zoom_combo, self.pv_full,
		]
		self._update_preview_controls_enabled()

		header.addLayout(group)

	def _update_preview_controls_enabled(self):
		on = self._preview_enabled
		for w in getattr(self, "_pv_view_controls", []):
			w.setEnabled(on)

	def _on_preview_toggle(self, on: bool):
		# Never allow preview for multi-file selections.
		if on and not self._preview_allowed():
			self.pv_toggle.setChecked(False)
			return
		self._preview_enabled = on
		self.pv_toggle.setIcon(self._eye_on_icon if on else self._eye_off_icon)
		self.pv_toggle.setToolTip("Turn preview off" if on else "Turn preview on")
		self._update_preview_controls_enabled()
		if on:
			self._preview_current_file()
		else:
			# Stop any in-flight render and clear the canvas.
			if hasattr(self, "preview"):
				self.preview.clear()
			self.preview_canvas.set_placeholder("Preview is off")

	def _on_zoom_changed(self, percent: float):
		txt = f"{int(round(percent))}%"
		if hasattr(self, "pv_zoom_combo"):
			self.pv_zoom_combo.blockSignals(True)
			self.pv_zoom_combo.setCurrentText(txt)
			self.pv_zoom_combo.blockSignals(False)

	def _on_zoom_combo(self, _idx: int):
		txt = self.pv_zoom_combo.currentText().rstrip("%")
		try:
			pct = float(txt)
		except ValueError:
			return
		self.preview_canvas._set_zoom(pct / 100.0)

	def _init_preview(self):
		"""Create the preview controller and connect signals (called from __init__)."""
		self.preview = PreviewController(self)
		self.preview.imageReady.connect(self.preview_canvas.set_image)
		self.preview.statusChanged.connect(self._on_preview_status)
		self.preview.busyChanged.connect(self._on_preview_busy)
		self.preview.protectedFile.connect(self._on_preview_protected)

	def _request_preview(self):
		"""Push current settings to the controller and refresh the preview."""
		if not hasattr(self, "preview") or not self._preview_enabled:
			return
		self.preview.set_settings(
			self.wm_text.text(), self._color_hex,
			self.trans_slider.value() / 100.0)
		if self._files:
			self.preview.refresh()

	def _preview_current_file(self):
		"""Load the selected file into the preview (first of a folder batch)."""
		if not hasattr(self, "preview") or not self._preview_enabled:
			return
		if not self._files:
			self.preview.clear()
			self.preview_canvas.set_placeholder("Select a file to preview")
			return
		self.preview.set_settings(
			self.wm_text.text(), self._color_hex,
			self.trans_slider.value() / 100.0)
		self.preview.load(self._files[0])

	def _on_preview_status(self, text: str):
		if text:
			self.status_lbl.setText(text)

	def _on_preview_busy(self, busy: bool):
		if busy:
			self.preview_canvas.set_placeholder("Loading preview…")

	def _on_preview_protected(self, path: str):
		"""A sensitivity-labelled / encrypted file can't be rendered for preview."""
		self.preview_canvas.set_image(QImage())
		self.preview_canvas.set_placeholder(
			"🔒  Can't preview a protected file\n\n"
			"This document has a sensitivity label with encryption.\n"
			"Watermarking still works — just turn preview off and apply.")
		self.status_lbl.setText("Protected file — preview unavailable.")

	def _build_right_column(self) -> QWidget:
		col = QWidget()
		col.setFixedWidth(236)
		col.setMouseTracking(True)
		v = QVBoxLayout(col)
		v.setContentsMargins(0, 0, 0, 0)
		v.setSpacing(14)

		# ── Actions card ───────────────────────────────────────────────
		actions = self._card("Actions")
		av = actions.layout()

		self.btn_apply = QPushButton("  Apply Watermark")
		self.btn_apply.setProperty("accent", True)
		self.btn_apply.setIcon(glyph_icon(ICON_STAMP, 18, ACCENT_FG))
		self.btn_apply.setIconSize(QSize(20, 20))
		self.btn_apply.setMinimumHeight(48)
		self.btn_apply.setCursor(Qt.PointingHandCursor)
		self.btn_apply.setToolTip("Watermark the selected file and save a copy")
		self.btn_apply.clicked.connect(self._on_apply_clicked)
		av.addWidget(self.btn_apply)

		self.btn_export = QPushButton("  Export PDF")
		self.btn_export.setIcon(glyph_icon(ICON_EXPORT, 16, FG_SOFT))
		self.btn_export.setIconSize(QSize(18, 18))
		self.btn_export.setMinimumHeight(44)
		self.btn_export.setCursor(Qt.PointingHandCursor)
		self.btn_export.setToolTip("Watermark and also export a PDF (PowerPoint / Word)")
		self.btn_export.clicked.connect(self._on_export_clicked)
		av.addWidget(self.btn_export)

		# Toggle rows
		self.toggle_export = ToggleSwitch(self._export_pdf_only)
		self.toggle_export.toggled.connect(self._on_export_toggle)
		av.addWidget(self._toggle_row(
			"Export PDF", "(PowerPoint / Word only)", self.toggle_export))

		self.toggle_open = ToggleSwitch(self._open_after)
		self.toggle_open.toggled.connect(self._on_open_toggle)
		av.addWidget(self._toggle_row(
			"Open file(s) after", "processing", self.toggle_open))

		v.addWidget(actions)

		# ── Status card ────────────────────────────────────────────────
		status = self._card("Status")
		sv = status.layout()

		ready_row = QHBoxLayout()
		ready_row.setSpacing(12)
		badge = QFrame()
		badge.setObjectName("StatusBadge")
		badge.setFixedSize(36, 36)
		bl = QVBoxLayout(badge)
		bl.setContentsMargins(0, 0, 0, 0)
		check = make_icon_label(ICON_CHECK, 16, SUCCESS)
		check.setAlignment(Qt.AlignCenter)
		bl.addWidget(check)
		ready_row.addWidget(badge)

		ready_col = QVBoxLayout()
		ready_col.setSpacing(1)
		self.status_title = QLabel("Ready")
		self.status_title.setObjectName("StatusReady")
		self.status_sub = QLabel("No files processed yet.")
		self.status_sub.setObjectName("StatusReadySub")
		self.status_sub.setWordWrap(True)
		ready_col.addWidget(self.status_title)
		ready_col.addWidget(self.status_sub)
		ready_row.addLayout(ready_col, 1)
		sv.addLayout(ready_row)

		div = QFrame(); div.setProperty("divider", True)
		sv.addWidget(div)

		self.stat_selected_val = QLabel("0")
		self.stat_processed_val = QLabel("0")
		self.stat_lastrun_val = QLabel("—")
		sv.addLayout(self._stat_row("Files selected", self.stat_selected_val))
		sv.addLayout(self._stat_row("Files processed", self.stat_processed_val))
		sv.addLayout(self._stat_row("Last run", self.stat_lastrun_val))

		v.addWidget(status)
		v.addStretch(1)

		# ── Help & Support ─────────────────────────────────────────────
		help_btn = QPushButton("  Help & Support")
		help_btn.setObjectName("HelpLink")
		help_btn.setProperty("ghost", True)
		help_btn.setIcon(glyph_icon(ICON_HELP, 15, MUTED))
		help_btn.setCursor(Qt.PointingHandCursor)
		help_btn.clicked.connect(self._open_help)
		v.addWidget(help_btn, 0, Qt.AlignRight)

		return col

	def _toggle_row(self, label: str, sub: str, toggle: ToggleSwitch) -> QFrame:
		row = QFrame()
		row.setObjectName("ToggleRow")
		row.setMinimumHeight(56)
		h = QHBoxLayout(row)
		h.setContentsMargins(14, 8, 12, 8)
		h.setSpacing(8)
		col = QVBoxLayout()
		col.setSpacing(1)
		t = QLabel(label); t.setObjectName("ToggleLabel")
		col.addWidget(t)
		if sub:
			s = QLabel(sub); s.setObjectName("ToggleSub")
			col.addWidget(s)
		h.addLayout(col, 1)
		h.addWidget(toggle, 0, Qt.AlignVCenter)
		return row

	def _stat_row(self, key: str, value_lbl: QLabel) -> QHBoxLayout:
		h = QHBoxLayout()
		k = QLabel(key); k.setObjectName("StatRowKey")
		value_lbl.setObjectName("StatRowVal")
		h.addWidget(k)
		h.addStretch(1)
		h.addWidget(value_lbl)
		return h

	def _build_statusbar(self) -> QWidget:
		bar = QFrame()
		bar.setObjectName("AppStatusBar")
		bar.setFixedHeight(30)
		bar.setMouseTracking(True)
		h = QHBoxLayout(bar)
		h.setContentsMargins(16, 0, 16, 0)
		dot = QLabel("\u25CF")
		dot.setStyleSheet(f"color: {SUCCESS};")
		h.addWidget(dot)
		self.status_lbl = QLabel("Ready")
		self.status_lbl.setProperty("muted", True)
		h.addWidget(self.status_lbl)
		h.addSpacing(12)
		self.sel_lbl = QLabel("No file selected")
		self.sel_lbl.setProperty("muted", True)
		h.addWidget(self.sel_lbl)
		h.addStretch(1)
		return bar

	# ── Left-rail helpers & logic ──────────────────────────────────────
	def _icon_button(self, text: str, glyph: str, accent: bool = False) -> QPushButton:
		"""Create a left-aligned button with an icon-font glyph."""
		btn = QPushButton("  " + text)
		btn.setCursor(Qt.PointingHandCursor)
		if accent:
			btn.setProperty("accent", True)
		color = ACCENT_FG if accent else FG_SOFT
		btn.setIcon(glyph_icon(glyph, 16, color))
		btn.setIconSize(QSize(18, 18))
		btn.setMinimumHeight(38)
		btn.setStyleSheet("text-align: left; padding-left: 12px;")
		return btn

	def _pick_file(self):
		path, _ = QFileDialog.getOpenFileName(
			self, "Select a file to watermark", "",
			"Supported files (*.pptx *.ppt *.docx *.doc *.mp4 *.mov *.m4v *.mkv *.avi *.webm);;"
			"PowerPoint (*.pptx *.ppt);;Word (*.docx *.doc);;"
			"Video (*.mp4 *.mov *.m4v *.mkv *.avi *.webm);;All files (*.*)",
		)
		if path:
			self._set_files([os.path.normpath(path)])

	def _pick_folder(self):
		folder = QFileDialog.getExistingDirectory(self, "Select folder to batch watermark")
		if not folder:
			return
		folder = os.path.normpath(folder)
		files = [
			os.path.join(folder, f) for f in os.listdir(folder)
			if os.path.splitext(f)[1].lower() in ALL_SUPPORTED
		]
		if files:
			self._set_files(files)
			self.status_lbl.setText(f"Folder selected — {len(files)} file(s) ready.")
		else:
			self.status_lbl.setText("No supported files found in that folder.")

	def _set_files(self, paths: list[str]):
		"""Replace the current selection with the given file(s).

		One-off picks and drops replace the selection (no queue); folder picks
		pass the whole batch. Duplicates within a single selection are removed
		but nothing accumulates across selections.
		"""
		seen = set()
		selected = []
		for p in paths:
			if os.path.isfile(p) and p not in seen:
				seen.add(p)
				selected.append(p)
				add_recent(p)
		self._files = selected
		# Preview is opt-in per file: every new selection starts with it OFF
		# so opening a document never launches Word/PowerPoint unasked.
		self._force_preview_off()
		self._on_files_changed()
		self._refresh_recent()

	def _force_preview_off(self):
		"""Reset the preview toggle to off (used on every new file open)."""
		if not hasattr(self, "pv_toggle"):
			return
		if self.pv_toggle.isChecked():
			self.pv_toggle.setChecked(False)   # fires _on_preview_toggle(False)
		else:
			self._preview_enabled = False
			if hasattr(self, "preview"):
				self.preview.clear()
			if hasattr(self, "preview_canvas"):
				self.preview_canvas.set_placeholder(
					"Preview is off — click the Preview icon to turn it on")

	def _preview_allowed(self) -> bool:
		"""Preview only makes sense for a single selected file."""
		return len(self._files) == 1

	def _update_preview_availability(self):
		"""Enable the Preview toggle only for a single file; disable for batches.

		Folder picks and multi-file selections can't be meaningfully previewed
		(which one?), so the preview is switched off, every preview control is
		disabled, and the status is updated to say so.
		"""
		if not hasattr(self, "pv_toggle"):
			return
		allowed = self._preview_allowed()
		# Turning a disabled toggle off first avoids a stuck-on state.
		if not allowed and self.pv_toggle.isChecked():
			self.pv_toggle.setChecked(False)
		self.pv_toggle.setEnabled(allowed)
		# View controls (zoom/fit) follow preview state; force them off in a batch.
		for ctl in getattr(self, "_pv_view_controls", []):
			ctl.setEnabled(allowed and self._preview_enabled)

		if allowed:
			self.pv_toggle.setToolTip(
				"Turn preview off" if self._preview_enabled else "Turn preview on")
		else:
			self.pv_toggle.setToolTip("Preview is available for a single file only")
			if len(self._files) > 1:
				if hasattr(self, "preview_canvas"):
					self.preview_canvas.set_placeholder(
						"Preview isn't available for multiple files.\n"
						"Select a single file to preview it.")
				self.status_lbl.setText(
					f"{len(self._files)} files selected — preview disabled for batches.")
			# (0 files keeps whatever placeholder / status is already shown.)

	def _on_files_changed(self):
		"""Hook for selection changes — updates status bar, status card, preview."""
		n = len(self._files)
		if n == 0:
			self.sel_lbl.setText("No file selected")
		elif n == 1:
			self.sel_lbl.setText(os.path.basename(self._files[0]))
		else:
			self.sel_lbl.setText(f"{n} files selected")
		if hasattr(self, "stat_selected_val"):
			self.stat_selected_val.setText(str(n))
		# Gate preview availability to single-file selections.
		self._update_preview_availability()
		# Preview only loads when the user has switched it on (opt-in per file).
		if hasattr(self, "preview") and self._preview_enabled and self._preview_allowed():
			self._preview_current_file()

	def _open_settings(self):
		self.status_lbl.setText("Settings — coming soon.")

	# ── Recent presets ─────────────────────────────────────────────────
	def _clear_layout(self, lay):
		while lay.count():
			item = lay.takeAt(0)
			w = item.widget()
			if w is not None:
				w.deleteLater()

	def _refresh_recent(self):
		"""Populate the recent-presets list from saved presets."""
		self._presets = load_presets()
		# Rebuild rows
		lay = self.recent_list_lay
		self._clear_layout(lay)
		if not self._presets:
			empty = QLabel("No saved presets yet.")
			empty.setProperty("muted", True)
			empty.setAlignment(Qt.AlignCenter)
			lay.addWidget(empty)
			lay.addStretch(1)
			return
		for name, data in self._presets.items():
			subtitle = self._preset_subtitle(name, data)
			color = data.get("color", "#A6A6A6")
			row = PresetRow(name, subtitle, color=color, data=data)
			row.clicked.connect(self._apply_preset_by_name)
			row.deleteRequested.connect(self._delete_preset_by_name)
			lay.addWidget(row)
		lay.addStretch(1)

	def _refresh_preset_combo(self):
		"""Sync the step-4 preset dropdown with saved presets."""
		if not hasattr(self, "preset_combo"):
			return
		self.preset_combo.blockSignals(True)
		current = self.preset_combo.currentText()
		self.preset_combo.clear()
		self.preset_combo.addItem("")
		for name in self._presets.keys():
			self.preset_combo.addItem(name)
		idx = self.preset_combo.findText(current)
		if idx >= 0:
			self.preset_combo.setCurrentIndex(idx)
		self.preset_combo.blockSignals(False)

	def _preset_subtitle(self, name: str, data: dict) -> str:
		"""Second line for a preset row.

		Avoids repeating the watermark text when the preset name already
		conveys it (e.g. name 'CONFIDENTIAL yellow', text 'CONFIDENTIAL').
		In that case we show just the transparency.
		"""
		txt = (data.get("text") or "").strip()
		trans = data.get("transparency")
		trans_str = f"{int(trans)}%" if isinstance(trans, (int, float)) else None

		name_l = name.strip().lower()
		redundant = bool(txt) and (txt.lower() in name_l or name_l in txt.lower())

		if redundant:
			return f"Transparency {trans_str}" if trans_str else "—"

		shown = txt if len(txt) <= 18 else txt[:17] + "…"
		if trans_str:
			return f"{shown or '—'}  ·  {trans_str}"
		return shown or "—"

	def _apply_preset_by_name(self, name: str):
		data = self._presets.get(name)
		if not data:
			return
		self._apply_preset_data(data)
		if hasattr(self, "preset_combo"):
			idx = self.preset_combo.findText(name)
			if idx >= 0:
				self.preset_combo.blockSignals(True)
				self.preset_combo.setCurrentIndex(idx)
				self.preset_combo.blockSignals(False)
		self.status_lbl.setText(f'Loaded preset "{name}".')

	def _delete_preset_by_name(self, name: str):
		delete_preset(name)
		self._refresh_recent()
		self._refresh_preset_combo()
		self.status_lbl.setText(f'Deleted preset "{name}".')

	# ── Center controls: text / colour / transparency ──────────────────
	def _on_text_changed(self, _txt: str):
		self._update_char_count()
		if hasattr(self, "_request_preview"):
			self._request_preview()

	def _update_char_count(self):
		n = len(self.wm_text.text())
		self.char_count.setText(f"{n} / {self._MAX_CHARS}")
		self.char_count.adjustSize()
		# Right-align inside the line edit
		x = self.wm_text.width() - self.char_count.width() - 14
		y = (self.wm_text.height() - self.char_count.height()) // 2
		self.char_count.move(max(0, x), max(0, y))

	def eventFilter(self, obj, event):
		if obj is getattr(self, "wm_text", None) and event.type() == QEvent.Resize:
			self._update_char_count()
		return super().eventFilter(obj, event)

	def _apply_swatch_color(self, hexc: str):
		self.color_swatch.setStyleSheet(
			f"background: {hexc}; border: 1px solid {BORDER}; border-radius: 6px;")

	def _set_color(self, hexc: str):
		self._color_hex = hexc.upper()
		self._apply_swatch_color(self._color_hex)
		self.hex_value.setText(self._color_hex)
		self._refresh_palette_selection()
		if hasattr(self, "_request_preview"):
			self._request_preview()

	def _refresh_palette_selection(self):
		"""Re-render swatch borders to reflect the current selection."""
		grid = self._palette_buttons
		for i in range(grid.count()):
			btn = grid.itemAt(i).widget()
			if btn is None:
				continue
			css = btn.styleSheet()
			# Extract the background colour we stored
			try:
				hexc = css.split("background:")[1].split(";")[0].strip()
			except Exception:
				continue
			selected = hexc.lower() == self._color_hex.lower()
			btn.setStyleSheet(self._swatch_btn_css(hexc, selected))

	def _pick_color(self):
		col = QColorDialog.getColor(QColor(self._color_hex), self, "Pick watermark colour")
		if col.isValid():
			self._set_color(col.name())

	def _on_trans_changed(self, value: int):
		self.trans_pill.setText(f"{value}%")
		if hasattr(self, "_request_preview"):
			self._request_preview()

	def _on_trans_pill_edited(self):
		"""Parse a typed transparency value and sync it back to the slider."""
		raw = "".join(ch for ch in self.trans_pill.text() if ch.isdigit())
		value = int(raw) if raw else self.trans_slider.value()
		value = max(0, min(100, value))
		# Always normalise the displayed text; setValue only fires
		# _on_trans_changed when the value actually changes.
		if value == self.trans_slider.value():
			self.trans_pill.setText(f"{value}%")
		else:
			self.trans_slider.setValue(value)

	# ── Center controls: presets ───────────────────────────────────────
	def _apply_preset_data(self, data: dict):
		self.wm_text.setText(data.get("text", "CONFIDENTIAL"))
		self._set_color(data.get("color", "#A6A6A6"))
		self.trans_slider.setValue(int(data.get("transparency", 70.0)))

	def _on_preset_combo(self, name: str):
		if name and name in self._presets:
			self._apply_preset_data(self._presets[name])

	def _save_preset(self):
		default = self.preset_combo.currentText() or self.wm_text.text()
		name, ok = QInputDialog.getText(self, "Save Preset", "Preset name:", text=default)
		if not ok or not name.strip():
			return
		name = name.strip()
		save_preset(name, self.wm_text.text(), self._color_hex,
					float(self.trans_slider.value()))
		self._refresh_recent()
		self._refresh_preset_combo()
		idx = self.preset_combo.findText(name)
		if idx >= 0:
			self.preset_combo.setCurrentIndex(idx)
		self.status_lbl.setText(f'Saved preset "{name}".')

	def _delete_current_preset(self):
		name = self.preset_combo.currentText()
		if not name:
			self.status_lbl.setText("Select a preset to delete.")
			return
		self._delete_preset_by_name(name)
		self.preset_combo.setCurrentIndex(0)

	def _toggle_advanced(self):
		self.status_lbl.setText("Advanced options — coming soon.")

	# ── Right column: actions, toggles, status ─────────────────────────
	def _on_export_toggle(self, value: bool):
		self._export_pdf_only = value
		self.status_lbl.setText(
			"PDF export enabled." if value else "PDF export disabled.")

	def _on_open_toggle(self, value: bool):
		self._open_after = value

	def _on_apply_clicked(self):
		"""Apply the watermark to the selected file (or folder batch) on a worker thread."""
		if not self._files:
			self.status_lbl.setText("Select a file first.")
			self._pulse_status_sub("No file selected — choose a file to watermark.")
			return
		if not self.wm_text.text().strip():
			self.status_lbl.setText("Enter watermark text.")
			return
		self._start_processing(export_only=False)

	def _on_export_clicked(self):
		"""Export PDF for the selected Office file(s) — forces export_pdf on."""
		if not self._files:
			self.status_lbl.setText("Select a file first.")
			return
		if not self.wm_text.text().strip():
			self.status_lbl.setText("Enter watermark text.")
			return
		# Only Office files can export PDF; warn if the selection is video-only.
		if all(file_kind(f) == "video" for f in self._files):
			self.status_lbl.setText("PDF export applies to PowerPoint / Word only.")
			return
		self._start_processing(export_only=True)

	# ── Processing pipeline (threaded) ─────────────────────────────────
	def _start_processing(self, export_only: bool):
		if getattr(self, "_proc_thread", None) is not None:
			self.status_lbl.setText("Already processing — please wait.")
			return

		try:
			color_rgb = int(self._color_hex.lstrip("#"), 16)
		except ValueError:
			color_rgb = 0xA6A6A6
		transparency = max(0.0, min(1.0, self.trans_slider.value() / 100.0))
		export_pdf = bool(self._export_pdf_only or export_only)
		text = self.wm_text.text().strip()
		files = list(self._files)

		self._set_processing_ui(True)
		self._set_status_state("Working…", f"Processing 0 / {len(files)}…")
		self.status_lbl.setText("Processing…")

		self._proc_thread = QThread(self)
		self._proc_worker = ProcessWorker(files, text, color_rgb, transparency, export_pdf)
		self._proc_worker.moveToThread(self._proc_thread)
		self._proc_thread.started.connect(self._proc_worker.run)
		self._proc_worker.progress.connect(self._on_proc_progress)
		self._proc_worker.encode.connect(self._on_proc_encode)
		self._proc_worker.fileDone.connect(self._on_proc_file_done)
		self._proc_worker.needFfmpeg.connect(self._on_need_ffmpeg)
		self._proc_worker.finished.connect(self._on_proc_finished)
		self._proc_thread.start()

	def _set_processing_ui(self, busy: bool):
		self.btn_apply.setEnabled(not busy)
		self.btn_export.setEnabled(not busy)
		if busy:
			self.btn_apply.setText("  Working…")
		else:
			self.btn_apply.setText("  Apply Watermark")

	def _on_proc_progress(self, done: int, total: int, name: str):
		self._set_status_state("Working…", f"Processing {done} / {total}\n{name}")
		self.status_lbl.setText(f"Processing {done} / {total} — {name}")

	def _on_proc_encode(self, seconds: float):
		self.status_lbl.setText(f"Encoding… {seconds:0.1f}s processed")

	def _on_proc_file_done(self, out_path: str):
		self._last_output_path = out_path

	def _on_need_ffmpeg(self):
		self._prompt_ffmpeg_download()

	def _on_proc_finished(self, done: int, errors: list, last_out: str):
		# Tear down the thread
		if getattr(self, "_proc_thread", None) is not None:
			self._proc_thread.quit()
			self._proc_thread.wait()
			self._proc_thread = None
			self._proc_worker = None

		self._set_processing_ui(False)
		self._files_processed += done
		self.stat_processed_val.setText(str(self._files_processed))

		from datetime import datetime
		self.stat_lastrun_val.setText(datetime.now().strftime("%I:%M %p").lstrip("0"))

		if errors and not done:
			self._set_status_state("Failed", errors[0])
			self.status_lbl.setText(f"Error: {errors[0]}")
		elif errors:
			self._set_status_state("Completed with errors",
								   f"{done} done, {len(errors)} failed.")
			self.status_lbl.setText(f"Done: {done}; {len(errors)} error(s).")
		elif done <= 1:
			self._set_status_state("Success", f"{done} file watermarked.")
			self.status_lbl.setText(f"Done. {done} file processed.")
		else:
			# Batch: make it clear it's finished and where the files are.
			self._set_status_state("Success",
								   f"{done} files watermarked. Opening folder…")
			self.status_lbl.setText(f"Done. {done} files processed — opened the folder.")

		# Never let failures be silent — surface them in a dialog.
		if errors:
			self._show_error_dialog(done, errors)

		if last_out:
			self._last_output_path = last_out
			if self._open_after:
				# Single file: open it. Batch (2+): don't flood the user with
				# windows — open the containing folder instead.
				if done <= 1:
					self._open_output(last_out)
				else:
					self._open_output_folder(last_out)

	def _show_error_dialog(self, done: int, errors: list):
		"""Tell the user clearly which files failed and why."""
		from PySide6.QtWidgets import QMessageBox
		box = QMessageBox(self)
		box.setIcon(QMessageBox.Warning if done else QMessageBox.Critical)
		box.setWindowTitle("Some files could not be watermarked"
						   if done else "Watermarking failed")
		headline = (f"{done} file(s) watermarked, {len(errors)} could not be."
					if done else f"{len(errors)} file(s) could not be watermarked.")
		box.setText(headline)
		box.setInformativeText("\n".join(f"• {e}" for e in errors[:12]))
		box.setStandardButtons(QMessageBox.Ok)
		box.exec()

	def _open_output(self, path: str):
		try:
			os.startfile(path)  # type: ignore[attr-defined]
			# Also open a sibling PDF when export was requested.
			pdf = os.path.splitext(path)[0] + ".pdf"
			if self._export_pdf_only and os.path.isfile(pdf):
				os.startfile(pdf)  # type: ignore[attr-defined]
		except Exception as exc:  # noqa: BLE001
			self.status_lbl.setText(f"Saved, but could not open: {exc}")

	def _open_output_folder(self, path: str):
		"""Open the folder containing the outputs and select the latest file."""
		try:
			folder = os.path.dirname(os.path.abspath(path))
			if os.path.isfile(path):
				# Open Explorer with the file selected.
				subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
			else:
				os.startfile(folder)  # type: ignore[attr-defined]
		except Exception as exc:  # noqa: BLE001
			self.status_lbl.setText(f"Saved to {folder}, but could not open it: {exc}")

	def _prompt_ffmpeg_download(self):
		from PySide6.QtWidgets import QMessageBox
		resp = QMessageBox.question(
			self, "ffmpeg required",
			"Video watermarking needs ffmpeg, which hasn't been downloaded yet.\n\n"
			"Download it once (~30 MB) now?",
			QMessageBox.Yes | QMessageBox.No)
		if resp != QMessageBox.Yes:
			self.status_lbl.setText("Video job cancelled — ffmpeg not downloaded.")
			return
		self.status_lbl.setText("Downloading ffmpeg…")

		dl_thread = QThread(self)
		dl_worker = _FfmpegDownloadWorker()
		dl_worker.moveToThread(dl_thread)
		dl_thread.started.connect(dl_worker.run)
		dl_worker.progress.connect(
			lambda mb: self.status_lbl.setText(f"Downloading ffmpeg… {mb:.1f} MB"))
		dl_worker.done.connect(lambda ok, err: self._on_ffmpeg_downloaded(
			ok, err, dl_thread, dl_worker))
		self._dl_thread = dl_thread
		dl_thread.start()

	def _on_ffmpeg_downloaded(self, ok, err, thread, worker):
		thread.quit(); thread.wait()
		self._dl_thread = None
		if ok:
			self.status_lbl.setText("ffmpeg ready. Re-run the video job.")
		else:
			self.status_lbl.setText(f"ffmpeg download failed: {err}")

	def _open_help(self):
		"""Open the project's GitHub repository in the default browser."""
		from PySide6.QtGui import QDesktopServices
		from PySide6.QtCore import QUrl
		if QDesktopServices.openUrl(QUrl(HELP_URL)):
			self.status_lbl.setText("Opened Help & Support in your browser.")
		else:
			self.status_lbl.setText(f"Could not open {HELP_URL}")

	def _set_status_state(self, title: str, subtitle: str):
		"""Update the right-column status card headline + subtitle."""
		if hasattr(self, "status_title"):
			self.status_title.setText(title)
			self.status_sub.setText(subtitle)

	def _pulse_status_sub(self, text: str):
		if hasattr(self, "status_sub"):
			self.status_sub.setText(text)

	# ── Frameless resize handling ──────────────────────────────────────
	def _edge_at(self, pos: QPoint) -> Qt.Edges:
		m = self._RESIZE_MARGIN
		r = self.rect()
		edges = Qt.Edges()
		if pos.x() <= m:
			edges |= Qt.LeftEdge
		if pos.x() >= r.width() - m:
			edges |= Qt.RightEdge
		if pos.y() <= m:
			edges |= Qt.TopEdge
		if pos.y() >= r.height() - m:
			edges |= Qt.BottomEdge
		return edges

	def _cursor_for(self, edges: Qt.Edges) -> Qt.CursorShape:
		if (edges & Qt.LeftEdge and edges & Qt.TopEdge) or (edges & Qt.RightEdge and edges & Qt.BottomEdge):
			return Qt.SizeFDiagCursor
		if (edges & Qt.RightEdge and edges & Qt.TopEdge) or (edges & Qt.LeftEdge and edges & Qt.BottomEdge):
			return Qt.SizeBDiagCursor
		if edges & (Qt.LeftEdge | Qt.RightEdge):
			return Qt.SizeHorCursor
		if edges & (Qt.TopEdge | Qt.BottomEdge):
			return Qt.SizeVerCursor
		return Qt.ArrowCursor

	def mouseMoveEvent(self, e):
		if not self.isMaximized():
			edges = self._edge_at(e.position().toPoint())
			self.setCursor(self._cursor_for(edges))
		super().mouseMoveEvent(e)

	def mousePressEvent(self, e):
		if e.button() == Qt.LeftButton and not self.isMaximized():
			edges = self._edge_at(e.position().toPoint())
			if edges:
				handle = self.windowHandle()
				if handle is not None:
					handle.startSystemResize(edges)
					e.accept()
					return
		super().mousePressEvent(e)

	def showEvent(self, e):
		super().showEvent(e)
		_win_round_and_shadow(self)

	def closeEvent(self, e):
		# Gracefully stop in-flight work so no signal fires into a dead window.
		try:
			from PySide6.QtCore import QThreadPool
			if getattr(self, "preview", None) is not None:
				self.preview._token += 1  # invalidate pending preview results
			if getattr(self, "_proc_worker", None) is not None:
				self._proc_worker.cancel()
			if getattr(self, "_proc_thread", None) is not None:
				self._proc_thread.quit()
				self._proc_thread.wait(3000)
			QThreadPool.globalInstance().waitForDone(3000)
		except Exception:
			pass
		super().closeEvent(e)


def main() -> None:
	# Qt 6 enables high-DPI scaling by default; use rounded-up rounding policy
	# so 125% / 150% Windows scaling stays crisp.
	QApplication.setHighDpiScaleFactorRoundingPolicy(
		Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
	app = QApplication(sys.argv)
	app.setApplicationName("Watermark Lab")
	if os.path.isfile(ICON_ICO):
		app.setWindowIcon(QIcon(ICON_ICO))
	app.setStyleSheet(build_qss())

	# Ensure the taskbar groups under our own icon (not python.exe) on Windows.
	if sys.platform == "win32":
		try:
			from ctypes import windll
			windll.shell32.SetCurrentProcessExplicitAppUserModelID(
				"BruceMu.WatermarkLab.X")
		except Exception:
			pass

	# Splash screen (parity with the legacy app's SplashLab.png).
	splash = None
	if os.path.isfile(SPLASH_PNG):
		pm = QPixmap(SPLASH_PNG)
		if not pm.isNull():
			splash = QSplashScreen(pm, Qt.WindowStaysOnTopHint)
			splash.show()
			app.processEvents()

	win = WatermarkLabX()
	# Center on screen
	scr = app.primaryScreen().availableGeometry()
	win.move((scr.width() - win.width()) // 2, (scr.height() - win.height()) // 2)

	def _reveal():
		win.show()
		if splash is not None:
			splash.finish(win)

	if splash is not None:
		QTimer.singleShot(750, _reveal)
	else:
		_reveal()

	sys.exit(app.exec())


if __name__ == "__main__":
	main()
