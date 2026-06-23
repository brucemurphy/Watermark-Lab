"""True-render preview engine for Watermark Lab X.

Two-tier strategy for a faithful yet responsive preview:

  1. TRUE RENDER (gold standard)
	 The source file is copied to a temp folder and the *real* backend
	 (_powerpoint.add_watermark / _word.add_word_watermark) is run there with
	 export_pdf=True. The resulting PDF — the genuine watermarked output — is
	 rasterised with Qt's built-in QtPdf (QPdfDocument). For video, a real
	 frame is extracted with ffmpeg and the actual _video watermark PNG is
	 composited. This is pixel-accurate because it IS the production output.

  2. LIVE COMPOSITE (instant feedback)
	 A clean (un-watermarked) page is rendered once per file and cached. While
	 the user drags sliders, a QPainter tiled-diagonal watermark is composited
	 over that cached background so feedback is immediate. When settings settle
	 a true render is kicked off in the background and swapped in.

All heavy work (COM, ffmpeg, QtPdf) runs on a QThreadPool. A monotonic token
discards stale results so only the latest request is shown.
"""
from __future__ import annotations

import io
import math
import os
import shutil
import subprocess
import tempfile
import zipfile

from PySide6.QtCore import (
	Qt, QObject, QRunnable, QThreadPool, Signal, QSize, QPointF, QRect,
)
from PySide6.QtGui import (
	QImage, QPainter, QColor, QFont,
)
from PySide6.QtPdf import QPdfDocument
from PySide6.QtWidgets import QWidget

PPT_EXTS   = {".pptx", ".ppt"}
WORD_EXTS_ = {".docx", ".doc"}
VIDEO_EXTS_ = {".mp4", ".mov", ".m4v", ".mkv", ".avi", ".webm"}

_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
_TARGET_W  = 1100   # rasterisation target width in px (quality vs. speed)


# ─────────────────────────────────────────────────────────────────────────────
# Low-level helpers
# ─────────────────────────────────────────────────────────────────────────────
def _ext(path: str) -> str:
	return os.path.splitext(path)[1].lower()


def file_kind(path: str) -> str:
	e = _ext(path)
	if e in PPT_EXTS:
		return "ppt"
	if e in WORD_EXTS_:
		return "word"
	if e in VIDEO_EXTS_:
		return "video"
	return "other"


class ProtectedFileError(Exception):
	"""Raised when a file is encrypted / sensitivity-label protected and so
	cannot be rendered for preview (Microsoft Information Protection / RMS)."""


class CancelledError(Exception):
	"""Raised internally to unwind a cancelled preview task."""


class _Cancel:
	"""A simple thread-safe cancellation flag shared with a running task.

	The controller flips this when the preview is turned off / cleared so the
	in-flight COM work can bail out at the next checkpoint instead of running
	to completion in the background.
	"""
	__slots__ = ("_flag",)

	def __init__(self):
		self._flag = False

	def cancel(self):
		self._flag = True

	@property
	def cancelled(self) -> bool:
		return self._flag


def _check_cancel(cancel) -> None:
	"""Raise CancelledError if the shared cancel flag is set."""
	if cancel is not None and cancel.cancelled:
		raise CancelledError()


# Modern Office files (.docx/.pptx) are ZIP archives. When a file carries a
# sensitivity label WITH ENCRYPTION (Microsoft Information Protection / Azure
# Rights Management) it is stored as an OLE2 compound file instead, whose
# content we cannot render — Office produces a placeholder "protected" page.
_ZIP_MAGIC  = b"PK\x03\x04"
_OLE2_MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"

# Markers found in the MIP placeholder page Office emits for protected files.
_MIP_MARKERS = (
	"microsoft information protection",
	"azure rights management",
	"uses encryption powered by",
	"unsupported pdf viewer",
)


def is_protected_file(path: str) -> bool:
	"""Fast, pre-COM check: is this Office file encrypted/label-protected?

	A modern .docx/.pptx is a ZIP; an encrypted (RMS/MIP) one is an OLE2
	compound file. Older .doc/.ppt are always OLE2, so we only treat the
	*modern* extensions as protected when they are OLE2. Returns False for
	anything we cannot positively identify as protected.
	"""
	ext = _ext(path)
	# Only the modern zip-based formats give us a clean signal.
	if ext not in (".docx", ".pptx"):
		return False
	try:
		with open(path, "rb") as fh:
			head = fh.read(8)
	except Exception:
		return False
	# Modern format but OLE2 header => encrypted / rights-protected container.
	return head == _OLE2_MAGIC or (not head.startswith(_ZIP_MAGIC) and head[:8] == _OLE2_MAGIC)


def _pdf_is_mip_placeholder(pdf_path: str) -> bool:
	"""Post-export check: did Office emit the MIP 'protected' placeholder PDF
	instead of the real document? Scans page-1 text for known markers."""
	doc = QPdfDocument()
	try:
		doc.load(pdf_path)
		if doc.status() != QPdfDocument.Status.Ready or doc.pageCount() < 1:
			return False
		try:
			text = doc.getAllText(0).text().lower()
		except Exception:
			return False
		return any(marker in text for marker in _MIP_MARKERS)
	finally:
		doc.close()


def rasterize_pdf_page(pdf_path: str, page: int = 0, target_w: int = _TARGET_W) -> QImage:
	"""Rasterise one PDF page to a QImage using QtPdf. Returns a null image on failure.

	QtPdf renders pages with a TRANSPARENT background (ARGB32, alpha 0), so the
	'paper' would show through as black on the dark preview canvas. We flatten
	the render onto opaque white so pages always look like paper.
	"""
	doc = QPdfDocument()
	status = doc.load(pdf_path)
	try:
		# QPdfDocument.load returns an Error enum (None == success on some builds)
		if doc.status() != QPdfDocument.Status.Ready:
			return QImage()
		if doc.pageCount() <= page:
			return QImage()
		pt = doc.pagePointSize(page)  # in points (1/72")
		if pt.width() <= 0:
			return QImage()
		scale = target_w / pt.width()
		size = QSize(int(pt.width() * scale), int(pt.height() * scale))
		rendered = doc.render(page, size)
		if rendered.isNull():
			return rendered
		# Flatten onto white so transparent 'paper' isn't shown as black.
		flat = QImage(rendered.size(), QImage.Format_RGB32)
		flat.fill(QColor("white"))
		p = QPainter(flat)
		p.drawImage(0, 0, rendered)
		p.end()
		return flat
	finally:
		doc.close()


def _pptx_thumbnail(path: str, target_w: int = _TARGET_W) -> QImage:
	"""Read the embedded first-slide thumbnail from a .pptx — NO COM, ~tens of ms.

	Office saves a JPEG preview at docProps/thumbnail.jpeg. It is low-res
	(typically 256x144) so we upscale smoothly to the preview width. Returns a
	null image if the file has no embedded thumbnail (older .ppt never do).
	"""
	if _ext(path) != ".pptx":
		return QImage()
	try:
		with zipfile.ZipFile(path) as z:
			name = None
			for n in z.namelist():
				if n.lower() == "docprops/thumbnail.jpeg":
					name = n
					break
			if name is None:
				return QImage()
			data = z.read(name)
	except Exception:
		return QImage()

	img = QImage()
	if not img.loadFromData(data):
		return QImage()
	if img.width() < target_w:
		img = img.scaledToWidth(target_w, Qt.SmoothTransformation)
	return img


def _word_pids() -> set[int]:
	"""Return the set of currently-running WINWORD.EXE process IDs."""
	try:
		out = subprocess.run(
			["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
			capture_output=True, text=True, creationflags=_NO_WINDOW, timeout=5,
		).stdout
	except Exception:
		return set()
	pids = set()
	for line in out.splitlines():
		parts = [p.strip('"') for p in line.split('","')]
		if len(parts) >= 2 and parts[0].upper().startswith("WINWORD"):
			try:
				pids.add(int(parts[1]))
			except ValueError:
				pass
	return pids


def _kill_pid(pid: int) -> None:
	try:
		subprocess.run(["taskkill", "/F", "/PID", str(pid)], check=False,
					   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
					   creationflags=_NO_WINDOW, timeout=5)
	except Exception:
		pass


def _word_page1_image(path: str, target_w: int = _TARGET_W, cancel=None) -> QImage:
	"""Render page 1 of a Word doc to an image via a single, short COM trip.

	Reliability notes (hard-won by measurement):
	- We use LATE binding (Dispatch), NOT gencache.EnsureDispatch. Early binding
	  regenerates Word's gen_py type-library cache on first use in a process,
	  which takes ~28s — that was the real source of the "25s" preview. Plain
	  Dispatch boots in ~1s. SaveAs2 works fine without early binding.
	- We use the robust SaveAs2(PDF) export (~1.8s) and rasterise only page 1
	  with QtPdf. (ExportAsFixedFormat with a page range crashes with an access
	  violation on real documents, so it is deliberately avoided.)
	- We record exactly which WINWORD PID our automation spawns and force-kill
	  only that one if Quit() doesn't release it — preventing zombie Word
	  processes from colliding and slowing later previews. A Word the user has
	  open is never touched.
	"""
	import pythoncom
	import win32com.client

	tmp_dir = tempfile.mkdtemp(prefix="wlx_w1_")
	local = os.path.join(tmp_dir, os.path.basename(path))
	try:
		shutil.copy2(path, local)
	except Exception:
		shutil.rmtree(tmp_dir, ignore_errors=True)
		return QImage()

	pdf = os.path.splitext(local)[0] + ".pdf"
	pids_before = _word_pids()
	pythoncom.CoInitialize()
	word = doc = None
	our_pid = None
	try:
		_check_cancel(cancel)
		word = win32com.client.Dispatch("Word.Application")
		# Identify the WINWORD process we just started (if any new one appeared).
		new_pids = _word_pids() - pids_before
		if len(new_pids) == 1:
			our_pid = next(iter(new_pids))

		word.Visible = False
		word.DisplayAlerts = False
		_check_cancel(cancel)
		doc = word.Documents.Open(os.path.abspath(local), ReadOnly=True,
								  AddToRecentFiles=False)
		_check_cancel(cancel)
		# Plain, robust full export; QtPdf then rasterises page 1 only.
		doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)  # wdFormatPDF
		if not os.path.isfile(pdf):
			return QImage()
		_check_cancel(cancel)
		# Catch protected docs that exported the MIP placeholder page.
		if _pdf_is_mip_placeholder(pdf):
			raise ProtectedFileError(path)
		return rasterize_pdf_page(pdf, 0, target_w)
	except (CancelledError, ProtectedFileError):
		raise
	except Exception:
		return QImage()
	finally:
		try:
			if doc:
				doc.Close(SaveChanges=False)
		except Exception:
			pass
		try:
			if word:
				word.Quit()
		except Exception:
			pass
		pythoncom.CoUninitialize()
		# If our specific WINWORD instance is still alive, terminate just it so
		# the next preview starts clean (prevents the contention slowdown).
		if our_pid is not None and our_pid in _word_pids():
			_kill_pid(our_pid)
		shutil.rmtree(tmp_dir, ignore_errors=True)


def clean_background(path: str, target_w: int = _TARGET_W, cancel=None) -> QImage:
	"""Fast, COM-light clean backdrop for the preview.

	- PPTX  -> embedded thumbnail (no COM at all)
	- DOC/X -> page-1-only export (one short COM trip, quits immediately)
	- Video -> first frame via ffmpeg
	- PPT (legacy, no thumbnail) -> full export fallback
	Returns a null image if nothing works.

	`cancel` is an optional _Cancel flag; the work bails out (raising
	CancelledError) at each checkpoint when the preview is turned off.

	Raises ProtectedFileError if the Office file is encrypted / sensitivity-label
	protected (Microsoft Information Protection) and therefore cannot be rendered.
	"""
	kind = file_kind(path)
	_check_cancel(cancel)
	# Pre-COM guard: never even launch Office for an encrypted/protected file.
	if kind in ("ppt", "word") and is_protected_file(path):
		raise ProtectedFileError(path)
	if kind == "ppt":
		img = _pptx_thumbnail(path, target_w)
		if not img.isNull():
			return img
		_check_cancel(cancel)
		# Legacy .ppt or thumbnail-less file: fall back to a full clean export.
		pdf = _office_export_pdf(path, False, "", 0, 0.0)
		_check_cancel(cancel)
		if pdf and _pdf_is_mip_placeholder(pdf):
			raise ProtectedFileError(path)
		return rasterize_pdf_page(pdf, 0, target_w) if pdf else QImage()
	if kind == "word":
		return _word_page1_image(path, target_w, cancel=cancel)
	if kind == "video":
		return _video_frame(path, target_w)
	return QImage()


def _office_export_pdf(path: str, with_watermark: bool, text: str,
					   color_rgb: int, transparency: float) -> str | None:
	"""Produce a PDF for the given Office file in a temp folder.

	with_watermark=False -> clean export of the original (for the live background).
	with_watermark=True  -> run the real backend so the PDF is the true output.
	Returns the PDF path or None on failure. Never touches the user's folder.
	"""
	kind = file_kind(path)
	if kind not in ("ppt", "word"):
		return None

	tmp_dir = tempfile.mkdtemp(prefix="wlx_prev_")
	local = os.path.join(tmp_dir, os.path.basename(path))
	try:
		shutil.copy2(path, local)
	except Exception:
		shutil.rmtree(tmp_dir, ignore_errors=True)
		return None

	try:
		if with_watermark:
			if kind == "ppt":
				from _powerpoint import add_watermark
				out = add_watermark(local, text, color_rgb=color_rgb,
									transparency=transparency, export_pdf=True)
			else:
				from _word import add_word_watermark
				out = add_word_watermark(local, text, color_rgb=color_rgb,
										 transparency=transparency, export_pdf=True)
			pdf = os.path.splitext(out)[0] + ".pdf"
			return pdf if os.path.isfile(pdf) else None
		else:
			# Clean export of the original — no watermark.
			if kind == "ppt":
				return _ppt_to_pdf(local)
			return _word_to_pdf(local)
	except Exception:
		return None


def _ppt_to_pdf(path: str) -> str | None:
	import pythoncom
	import win32com.client
	pythoncom.CoInitialize()
	pp = pres = None
	try:
		pp = win32com.client.Dispatch("PowerPoint.Application")
		pres = pp.Presentations.Open(os.path.abspath(path), WithWindow=False)
		pdf = os.path.splitext(path)[0] + "_clean.pdf"
		pres.SaveAs(os.path.abspath(pdf), 32)  # ppSaveAsPDF
		return pdf if os.path.isfile(pdf) else None
	finally:
		try:
			if pres:
				pres.Close()
		except Exception:
			pass
		try:
			if pp:
				pp.Quit()
		except Exception:
			pass
		pythoncom.CoUninitialize()


def _word_to_pdf(path: str) -> str | None:
	import pythoncom
	import win32com.client
	pythoncom.CoInitialize()
	word = doc = None
	try:
		word = win32com.client.Dispatch("Word.Application")
		word.Visible = False
		word.DisplayAlerts = False
		doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True,
								  AddToRecentFiles=False)
		pdf = os.path.splitext(path)[0] + "_clean.pdf"
		doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)  # wdFormatPDF
		return pdf if os.path.isfile(pdf) else None
	finally:
		try:
			if doc:
				doc.Close(SaveChanges=False)
		except Exception:
			pass
		try:
			if word:
				word.Quit()
		except Exception:
			pass
		pythoncom.CoUninitialize()


def _video_frame(path: str, target_w: int = _TARGET_W) -> QImage:
	"""Extract the first representative frame of a video as a QImage."""
	try:
		from _ffmpeg import get_ffmpeg_exe
		ffmpeg = get_ffmpeg_exe()
	except Exception:
		return QImage()
	tmp_dir = tempfile.mkdtemp(prefix="wlx_frame_")
	out_png = os.path.join(tmp_dir, "frame.png")
	try:
		cmd = [
			ffmpeg, "-y", "-hide_banner", "-loglevel", "error",
			"-ss", "1", "-i", path, "-frames:v", "1",
			"-vf", f"scale={target_w}:-1", out_png,
		]
		subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
					   creationflags=_NO_WINDOW, timeout=30)
		if os.path.isfile(out_png):
			img = QImage(out_png)
			return img
		return QImage()
	except Exception:
		return QImage()
	finally:
		shutil.rmtree(tmp_dir, ignore_errors=True)


def composite_watermark(base: QImage, text: str, color_hex: str,
						transparency: float, kind: str) -> QImage:
	"""QPainter watermark over a clean background image.

	- Word: ONE centred, diagonal, wrapped block (matches _xword output).
	- PPT / video: tiled diagonal grid (matches those backends).
	alpha = 1 - transparency. Used for instant slider feedback.
	"""
	if base.isNull():
		return base
	text = text or " "
	w, h = base.width(), base.height()
	out = QImage(base)
	if out.format() != QImage.Format_ARGB32_Premultiplied:
		out = out.convertToFormat(QImage.Format_ARGB32_Premultiplied)

	p = QPainter(out)
	p.setRenderHint(QPainter.Antialiasing, True)
	p.setRenderHint(QPainter.TextAntialiasing, True)

	col = QColor(color_hex)
	alpha = int(round(255 * (1.0 - max(0.0, min(1.0, transparency)))))
	col.setAlpha(alpha)
	p.setPen(col)

	if kind == "word":
		_paint_word_block(p, w, h, text, col)
	else:
		_paint_tiled(p, w, h, text)
	p.end()
	return out


def _wrap_preview_text(text: str) -> list[str]:
	"""Use _xword's exact wrapping so the preview matches the applied output."""
	try:
		from _xword import _wrap_text
		return _wrap_text(text)
	except Exception:
		# Fallback: minimal local wrap if _xword is unavailable.
		text = " ".join(text.split())
		if len(text) <= 20:
			return [text]
		words, lines, cur = text.split(), [], ""
		for word in words:
			cand = word if not cur else f"{cur} {word}"
			if len(cand) <= 20 or not cur:
				cur = cand
			else:
				lines.append(cur); cur = word
		if cur:
			lines.append(cur)
		return lines[:3]


def _paint_word_block(p: QPainter, w: int, h: int, text: str, col: QColor):
	"""One centred, -30° diagonal, wrapped text block (Word style).

	Matches _xword: ~80% printable-width block, Segoe UI Semibold.
	"""
	lines = _wrap_preview_text(text)
	# Block spans ~80% of the page width along the diagonal (matches _xword).
	target_w = w * 0.80
	font_px = max(18, int(w * 0.05))
	font = QFont("Segoe UI Semibold", font_px)
	font.setWeight(QFont.DemiBold)
	font.setPixelSize(font_px)
	p.setFont(font)
	fm = p.fontMetrics()
	widest = max((fm.horizontalAdvance(ln) for ln in lines), default=1)
	# Scale font so the widest line ~= target_w.
	if widest > 0:
		font_px = max(12, int(font_px * target_w / widest))
		font.setPixelSize(font_px)
		p.setFont(font)
		fm = p.fontMetrics()
	line_h = int(fm.height() * 1.15)
	block_h = line_h * len(lines)

	p.save()
	p.translate(w / 2, h / 2)
	p.rotate(-30)
	y0 = -block_h / 2 + fm.ascent()
	for i, ln in enumerate(lines):
		lw = fm.horizontalAdvance(ln)
		p.drawText(QPointF(-lw / 2, y0 + i * line_h), ln)
	p.restore()


def _paint_tiled(p: QPainter, w: int, h: int, text: str):
	"""Tiled diagonal grid (PowerPoint / video style)."""
	font_px = max(16, int(h * 0.035))
	font = QFont("Segoe UI", font_px)
	font.setPixelSize(font_px)
	p.setFont(font)
	fm = p.fontMetrics()
	tw = max(1, fm.horizontalAdvance(text))
	th = max(1, fm.height())

	p.translate(w / 2, h / 2)
	p.rotate(-30)
	diag = int(math.hypot(w, h))
	gap_x = max(font_px, tw // 3)
	gap_y = th * 2
	step_x = tw + gap_x
	step_y = th + gap_y

	row = 0
	y = -diag
	while y < diag:
		offset = (step_x // 2) if (row % 2) else 0
		x = -diag
		while x < diag:
			p.drawText(QPointF(x + offset, y), text)
			x += step_x
		y += step_y
		row += 1


# ─────────────────────────────────────────────────────────────────────────────
# Threaded tasks
# ─────────────────────────────────────────────────────────────────────────────
class _Signals(QObject):
	cleanReady = Signal(int, str, QImage)   # token, path, clean background
	failed     = Signal(int, str, str)      # token, path, message
	protected  = Signal(int, str)           # token, path (encrypted / MIP)


def _safe_emit(signal, *args) -> None:
	"""Emit a signal, swallowing the RuntimeError raised when the underlying
	C++ object has already been deleted (e.g. window closed mid-render)."""
	try:
		signal.emit(*args)
	except RuntimeError:
		pass


class _CleanTask(QRunnable):
	"""Produce the clean (un-watermarked) backdrop quickly and COM-light.

	PPTX uses the embedded thumbnail (no COM); Word does a single page-1 export;
	video grabs a frame. The watermark is composited live in the UI thread, so
	this is the only background work a preview ever does.

	Honours a shared _Cancel flag so toggling preview off stops the work.
	"""

	def __init__(self, token: int, path: str, signals: _Signals, cancel: _Cancel):
		super().__init__()
		self.token, self.path, self.sig = token, path, signals
		self.cancel = cancel

	def run(self):
		try:
			img = clean_background(self.path, cancel=self.cancel)
			if self.cancel.cancelled:
				return
			if img.isNull():
				_safe_emit(self.sig.failed, self.token, self.path,
						   "Could not render preview for this file.")
			else:
				_safe_emit(self.sig.cleanReady, self.token, self.path, img)
		except CancelledError:
			return  # silently drop — the user turned preview off
		except ProtectedFileError:
			if not self.cancel.cancelled:
				_safe_emit(self.sig.protected, self.token, self.path)
		except Exception as exc:  # noqa: BLE001
			if not self.cancel.cancelled:
				_safe_emit(self.sig.failed, self.token, self.path, str(exc))


# ─────────────────────────────────────────────────────────────────────────────
# Controller
# ─────────────────────────────────────────────────────────────────────────────
class PreviewController(QObject):
	"""Orchestrates fast clean-backdrop loads with caching and tokens.

	The watermark is always composited live in the UI thread over a cached
	clean backdrop, so slider/colour/text changes are instant and never touch
	COM. Only loading a new file does background work (and that is COM-light).
	"""

	imageReady   = Signal(QImage)   # composite to display
	statusChanged = Signal(str)     # human-readable preview status
	busyChanged   = Signal(bool)
	protectedFile = Signal(str)     # path — encrypted / sensitivity-labelled

	def __init__(self, parent=None):
		super().__init__(parent)
		self._pool = QThreadPool.globalInstance()
		self._sig = _Signals()
		self._sig.cleanReady.connect(self._on_clean_ready)
		self._sig.failed.connect(self._on_failed)
		self._sig.protected.connect(self._on_protected)

		self._token = 0
		self._path: str | None = None
		self._clean_bg: QImage = QImage()          # cached clean background
		self._clean_path: str | None = None        # which file the bg belongs to
		self._active_cancel: _Cancel | None = None  # flag for the in-flight task

		self._text = "CONFIDENTIAL"
		self._color_hex = "#A6A6A6"
		self._transparency = 0.70

	# ── Public API ─────────────────────────────────────────────────────
	def set_settings(self, text: str, color_hex: str, transparency: float):
		self._text = text
		self._color_hex = color_hex
		self._transparency = transparency

	def _cancel_active(self):
		"""Signal any in-flight preview task to stop at its next checkpoint."""
		if self._active_cancel is not None:
			self._active_cancel.cancel()
			self._active_cancel = None

	def load(self, path: str):
		"""Begin previewing a new file (fast, COM-light backdrop)."""
		# Any previous task is now stale — tell it to stop doing COM work.
		self._cancel_active()
		self._token += 1
		self._path = path
		kind = file_kind(path)
		if kind == "other":
			self.statusChanged.emit("Unsupported file type.")
			self.imageReady.emit(QImage())
			return
		if self._clean_path != path or self._clean_bg.isNull():
			self._clean_bg = QImage()
			self._clean_path = None
			self.statusChanged.emit("Loading preview…")
			self.busyChanged.emit(True)
			cancel = _Cancel()
			self._active_cancel = cancel
			self._pool.start(_CleanTask(self._token, path, self._sig, cancel))
		else:
			self._composite_live()

	def refresh(self):
		"""Re-composite the watermark over the cached backdrop (instant)."""
		if not self._path:
			return
		self._composite_live()

	def clear(self):
		# Stop any in-flight COM work immediately and drop its results.
		self._cancel_active()
		self._token += 1
		self._path = None
		self._clean_bg = QImage()
		self._clean_path = None
		self.busyChanged.emit(False)
		self.imageReady.emit(QImage())
		self.statusChanged.emit("")

	# ── Internal ───────────────────────────────────────────────────────
	def _color_rgb(self) -> int:
		try:
			return int(self._color_hex.lstrip("#"), 16)
		except ValueError:
			return 0xA6A6A6

	def _composite_live(self):
		if self._clean_bg.isNull() or not self._path:
			return
		img = composite_watermark(self._clean_bg, self._text, self._color_hex,
								  self._transparency, file_kind(self._path))
		self.imageReady.emit(img)

	def _on_clean_ready(self, token: int, path: str, img: QImage):
		if token != self._token:
			return
		self._active_cancel = None
		self._clean_bg = img
		self._clean_path = path
		self._composite_live()
		self.busyChanged.emit(False)
		self.statusChanged.emit("Preview ready")

	def _on_failed(self, token: int, path: str, message: str):
		if token != self._token:
			return
		self.busyChanged.emit(False)
		# Keep any live composite already shown; just report.
		self.statusChanged.emit(message)

	def _on_protected(self, token: int, path: str):
		if token != self._token:
			return
		self.busyChanged.emit(False)
		self._clean_bg = QImage()
		self._clean_path = None
		self.imageReady.emit(QImage())
		self.protectedFile.emit(path)
		self.statusChanged.emit("Protected file — preview unavailable.")


# ─────────────────────────────────────────────────────────────────────────────
# Preview canvas widget (zoom / pan / fit)
# ─────────────────────────────────────────────────────────────────────────────
class PreviewCanvas(QWidget):
	"""Displays a QImage with zoom, pan, and fit-to-window support."""

	zoomChanged = Signal(float)   # current zoom factor (1.0 == fit)

	_MIN_ZOOM = 0.1
	_MAX_ZOOM = 8.0

	def __init__(self, parent=None):
		super().__init__(parent)
		self.setMouseTracking(True)
		self._image: QImage = QImage()
		self._zoom = 1.0          # user zoom multiplier on top of fit-scale
		self._fit_scale = 1.0     # scale that fits the image to the viewport
		self._offset = QPointF(0, 0)
		self._panning = False
		self._pan_start = QPointF(0, 0)
		self._pan_mode = False    # explicit pan tool (hand) toggle
		self._placeholder = "Select a file to preview"

	# ── Public API ─────────────────────────────────────────────────────
	def set_image(self, img: QImage):
		first = self._image.isNull()
		self._image = img if img is not None else QImage()
		if self._image.isNull():
			self.update()
			return
		self._recompute_fit()
		if first:
			self.reset_view()
		self.update()

	def set_placeholder(self, text: str):
		self._placeholder = text
		self.update()

	def has_image(self) -> bool:
		return not self._image.isNull()

	def reset_view(self):
		self._zoom = 1.0
		self._offset = QPointF(0, 0)
		self.zoomChanged.emit(self._effective_percent())
		self.update()

	def zoom_in(self):
		self._set_zoom(self._zoom * 1.25)

	def zoom_out(self):
		self._set_zoom(self._zoom / 1.25)

	def fit(self):
		self.reset_view()

	def set_pan_mode(self, on: bool):
		self._pan_mode = on
		self.setCursor(Qt.OpenHandCursor if on else Qt.ArrowCursor)

	def current_percent(self) -> int:
		return int(round(self._effective_percent()))

	# ── Internals ──────────────────────────────────────────────────────
	def _effective_percent(self) -> float:
		# 100% == fit-to-window; user zoom multiplies from there.
		return self._zoom * 100.0

	def _recompute_fit(self):
		if self._image.isNull():
			self._fit_scale = 1.0
			return
		vw, vh = self.width() - 24, self.height() - 24
		iw, ih = self._image.width(), self._image.height()
		if iw <= 0 or ih <= 0 or vw <= 0 or vh <= 0:
			self._fit_scale = 1.0
			return
		self._fit_scale = min(vw / iw, vh / ih)

	def _set_zoom(self, z: float):
		z = max(self._MIN_ZOOM, min(self._MAX_ZOOM, z))
		if abs(z - self._zoom) < 1e-4:
			return
		self._zoom = z
		self.zoomChanged.emit(self._effective_percent())
		self.update()

	def resizeEvent(self, e):
		self._recompute_fit()
		super().resizeEvent(e)

	def wheelEvent(self, e):
		if self._image.isNull():
			return
		delta = e.angleDelta().y()
		if delta == 0:
			return
		factor = 1.0015 ** delta
		self._set_zoom(self._zoom * factor)

	def mousePressEvent(self, e):
		if e.button() == Qt.LeftButton and not self._image.isNull():
			self._panning = True
			self._pan_start = e.position()
			self.setCursor(Qt.ClosedHandCursor)

	def mouseMoveEvent(self, e):
		if self._panning:
			d = e.position() - self._pan_start
			self._pan_start = e.position()
			self._offset += d
			self.update()

	def mouseReleaseEvent(self, e):
		if e.button() == Qt.LeftButton and self._panning:
			self._panning = False
			self.setCursor(Qt.OpenHandCursor if self._pan_mode else Qt.ArrowCursor)

	def mouseDoubleClickEvent(self, e):
		self.reset_view()

	def paintEvent(self, _e):
		p = QPainter(self)
		p.fillRect(self.rect(), QColor("#0a0d12"))
		if self._image.isNull():
			p.setPen(QColor("#8b949e"))
			f = QFont("Segoe UI", 11)
			p.setFont(f)
			flags = Qt.AlignCenter | Qt.TextWordWrap
			# Inset so wrapped text never touches the edges.
			r = self.rect().adjusted(28, 28, -28, -28)
			p.drawText(r, flags, self._placeholder)
			p.end()
			return

		p.setRenderHint(QPainter.SmoothPixmapTransform, True)
		scale = self._fit_scale * self._zoom
		iw, ih = self._image.width(), self._image.height()
		dw, dh = iw * scale, ih * scale
		x = (self.width() - dw) / 2 + self._offset.x()
		y = (self.height() - dh) / 2 + self._offset.y()

		# Soft shadow behind the page
		shadow = QColor(0, 0, 0, 120)
		p.fillRect(int(x + 6), int(y + 8), int(dw), int(dh), shadow)
		target = QRect(int(x), int(y), int(dw), int(dh))
		p.drawImage(target, self._image)
		p.end()

