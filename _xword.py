"""Experimental Word watermarking for Watermark Lab X.

This is a drop-in alternative to _word.add_word_watermark that fixes the
"fat / stretched" watermark on long text. The legacy _word.py is left
untouched; the X app uses this module instead.

Root cause of the fat text (measured): the watermark shape is a FIXED box
(527.85pt x 131.95pt) and the VML <v:textpath fitshape="t"> stretches the text
to fill that box. Long single-line text therefore gets fattened horizontally.

Fix:
  * Wrap long text into a few balanced lines (snug, like a hand-made watermark).
  * Size the shape box to the *natural* proportions of the wrapped text block
	(measured with Pillow), so fitshape scales it uniformly with no distortion.
  * Clamp the on-page width so the diagonal block always fits the page.

Everything else (canonical shapetype, header injection, .doc->.docx, PDF export)
mirrors the proven _word.py implementation.
"""
import os
import subprocess
import sys

import pythoncom
import win32com.client
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

try:
	from PIL import ImageFont
except Exception:  # pragma: no cover
	ImageFont = None

WORD_EXTS = {".docx", ".doc"}
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# ── Geometry constants (points) ────────────────────────────────────────────
# The diagonal block is centred on the page; keep it within the printable area.
_MAX_BLOCK_W_PT = 368.0     # ~80% of printable width — snug, off the margins
_MIN_BLOCK_W_PT = 200.0
_BLOCK_H_CAP_PT = 320.0     # never taller than this once wrapped
_LINE_SPACING   = 1.18      # line height as a multiple of font size

# Wrapping: aim for a roughly square-ish block so the diagonal sits snug.
_TARGET_CHARS_PER_LINE = 20
_MAX_LINES = 3

# Semibold reads present without the heaviness of full bold (matches the
# hand-made HBI watermark look).
_FONT_BOLD_CANDIDATES = [
	r"C:\Windows\Fonts\seguisb.ttf",    # Segoe UI Semibold
	r"C:\Windows\Fonts\segoeui.ttf",
	r"C:\Windows\Fonts\arial.ttf",
]

# ── Canonical shapetype — identical to what Word's Design>Watermark writes ──
_SHAPETYPE_XML = (
	'<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" '
	'adj="10800" path="m@7,0l@8,0m@5,21600l@6,21600&amp;e" '
	'xmlns:v="urn:schemas-microsoft-com:vml" '
	'xmlns:o="urn:schemas-microsoft-com:office:office">'
	'<v:formulas>'
	'<v:f eqn="sum #0 0 10800"/><v:f eqn="prod #0 2 1"/>'
	'<v:f eqn="sum 21600 0 @1"/><v:f eqn="sum 0 0 @2"/>'
	'<v:f eqn="sum 21600 0 @3"/><v:f eqn="if @0 @3 0"/>'
	'<v:f eqn="if @0 21600 @1"/><v:f eqn="if @0 0 @2"/>'
	'<v:f eqn="if @0 @1 21600"/><v:f eqn="mid @5 @6"/>'
	'<v:f eqn="mid @8 @5"/><v:f eqn="mid @7 @8"/>'
	'<v:f eqn="mid @6 @7"/><v:f eqn="sum @6 0 @7"/>'
	'</v:formulas>'
	'<v:path textpathok="t" o:connecttype="custom" '
	'o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800"/>'
	'<v:textpath on="t" fitshape="t"/>'
	'<v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles>'
	'<o:lock v:ext="edit" shapetype="t"/>'
	'</v:shapetype>'
)


def _xe(t: str) -> str:
	return (t.replace("&", "&amp;").replace("<", "&lt;")
			 .replace(">", "&gt;").replace('"', "&quot;"))


def _load_bold_font(size_px: int):
	if ImageFont is None:
		return None
	for path in _FONT_BOLD_CANDIDATES:
		if os.path.isfile(path):
			try:
				return ImageFont.truetype(path, size_px)
			except OSError:
				continue
	try:
		return ImageFont.load_default()
	except Exception:
		return None


def _wrap_text(text: str) -> list[str]:
	"""Wrap long watermark text into a few balanced lines.

	Breaks on existing separators (" - ", " / ", ":") first, then on word
	boundaries, aiming for ~_TARGET_CHARS_PER_LINE per line and at most
	_MAX_LINES lines. Short text is returned as a single line unchanged.
	"""
	text = " ".join(text.split())  # normalise whitespace
	if len(text) <= _TARGET_CHARS_PER_LINE:
		return [text]

	# Prefer breaking on explicit separators the user typed.
	for sep in (" - ", " – ", " / ", ": "):
		if sep in text:
			parts = [p.strip() for p in text.split(sep) if p.strip()]
			if 2 <= len(parts) <= _MAX_LINES:
				# Re-wrap any over-long part by words.
				lines: list[str] = []
				for part in parts:
					lines.extend(_wrap_words(part))
				if len(lines) <= _MAX_LINES:
					return lines

	return _wrap_words(text)


def _wrap_words(text: str) -> list[str]:
	"""Greedy word-wrap toward the target line length, capped at _MAX_LINES."""
	words = text.split()
	lines: list[str] = []
	cur = ""
	for w in words:
		candidate = w if not cur else f"{cur} {w}"
		if len(candidate) <= _TARGET_CHARS_PER_LINE or not cur:
			cur = candidate
		else:
			lines.append(cur)
			cur = w
	if cur:
		lines.append(cur)

	# If we overshot the max number of lines, rebalance by raising the width.
	if len(lines) > _MAX_LINES:
		approx = max(_TARGET_CHARS_PER_LINE,
					 (len(text) + _MAX_LINES - 1) // _MAX_LINES)
		lines, cur = [], ""
		for w in words:
			candidate = w if not cur else f"{cur} {w}"
			if len(candidate) <= approx or not cur:
				cur = candidate
			else:
				lines.append(cur)
				cur = w
		if cur:
			lines.append(cur)
	return lines[:_MAX_LINES] if len(lines) > _MAX_LINES else lines


def _measure_block(lines: list[str]):
	"""Return (block_w_pt, block_h_pt) for the wrapped text block.

	Measures the widest line and the stacked height with Pillow at a reference
	size, then scales to fit within the on-page width budget. The ASPECT RATIO
	is what matters: matching the shape box to it makes fitshape scale text
	uniformly (no horizontal stretching / fattening).
	"""
	ref_px = 100
	font = _load_bold_font(ref_px)
	if font is None:
		# Fallback heuristic if Pillow is unavailable.
		widest = max((len(l) for l in lines), default=1)
		w_units = widest * 0.55 * ref_px
		h_units = len(lines) * ref_px * _LINE_SPACING
	else:
		widths = []
		for ln in lines:
			bbox = font.getbbox(ln or " ")
			widths.append(bbox[2] - bbox[0])
		w_units = max(widths) if widths else ref_px
		# Use the font ascent+descent for a stable per-line height.
		asc, desc = font.getmetrics()
		line_h = (asc + desc)
		h_units = len(lines) * line_h * _LINE_SPACING

	if w_units <= 0:
		w_units = ref_px
	aspect = h_units / w_units  # height per unit width

	# Scale so the block width sits within the page budget.
	block_w = max(_MIN_BLOCK_W_PT, min(_MAX_BLOCK_W_PT, _MAX_BLOCK_W_PT))
	block_h = block_w * aspect
	if block_h > _BLOCK_H_CAP_PT:
		block_h = _BLOCK_H_CAP_PT
		block_w = block_h / aspect if aspect > 0 else block_w
	return round(block_w, 2), round(block_h, 2)


def _shape_xml(text: str, color_hex: str) -> str:
	"""Build a watermark shape whose BOX matches the text's natural aspect so
	fitshape scales it without horizontal distortion; long text is wrapped."""
	lines = _wrap_text(text)
	block_w, block_h = _measure_block(lines)

	# VML textpath uses &#10; (line feed) to separate lines.
	vml_string = "&#10;".join(_xe(ln) for ln in lines)

	# A representative font-size for the textpath style (fitshape still governs
	# the final rendered size, but a sane value avoids extreme initial layout).
	font_pt = max(14.0, min(32.0, block_h / (len(lines) * _LINE_SPACING)))
	font_pt = round(font_pt, 1)

	# Semibold (not bold) — present but lighter, matching the HBI look.
	tp_style = (
		f"font-family:'Segoe UI Semibold';font-size:{font_pt}pt;font-weight:600;"
		f"color:{color_hex}"
	)
	return (
		f'<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2051" '
		f'type="#_x0000_t136" fillcolor="{color_hex}" '
		f'style="position:absolute;margin-left:0;margin-top:0;'
		f'width:{block_w}pt;height:{block_h}pt;z-index:-251654144;'
		'mso-position-horizontal:center;'
		'mso-position-horizontal-relative:margin;'
		'mso-position-vertical:center;'
		'mso-position-vertical-relative:margin;'
		'rotation:315" '
		'o:allowincell="f" stroked="f" '
		'xmlns:v="urn:schemas-microsoft-com:vml" '
		'xmlns:o="urn:schemas-microsoft-com:office:office" '
		'xmlns:w10="urn:schemas-microsoft-com:office:word">'
		f'<v:fill o:detectmouseclick="f" color="{color_hex}" color2="{color_hex}"/>'
		f'<v:textpath style="{tp_style}" string="{vml_string}" fitshape="t"/>'
		'<v:imagedata o:relid="" o:title=""/>'
		'<o:lock v:ext="edit" position="t"/>'
		'<w10:wrap w10:type="none"/>'
		'<w10:anchorlock/>'
		'</v:shape>'
	)


# -----------------------------------------------------------------------
# Public API — signature-compatible with _word.add_word_watermark
# -----------------------------------------------------------------------
def add_word_watermark(doc_path, watermark_text, color_rgb=0xA6A6A6,
					   transparency=0.70, export_pdf=False):
	doc_path = os.path.abspath(doc_path)
	if not os.path.isfile(doc_path):
		raise FileNotFoundError(doc_path)

	pythoncom.CoInitialize()
	try:
		base, ext = os.path.splitext(doc_path)

		if ext.lower() == ".doc":
			doc_path = _convert_doc_to_docx(doc_path)
			base = os.path.splitext(doc_path)[0]

		output_path = _next_available(base, "_watermarked", ".docx")
		color_hex = "#{:06x}".format(color_rgb)

		doc = Document(doc_path)
		_inject(doc, watermark_text, color_hex)
		doc.save(output_path)

		if export_pdf:
			_pdf_via_com(output_path)

		return output_path
	finally:
		pythoncom.CoUninitialize()


def _inject(doc, text, color_hex):
	"""Write watermark VML into the primary header of every section."""
	ns_v = "urn:schemas-microsoft-com:vml"
	seen = set()
	for section in doc.sections:
		hdr = section.header
		pid = id(hdr._element)
		if pid in seen:
			continue
		seen.add(pid)

		for tag in (f"{{{ns_v}}}shape", f"{{{ns_v}}}shapetype"):
			for el in hdr._element.findall(f".//{tag}"):
				p = el.getparent()
				if p is not None:
					p.remove(el)

		paras = hdr._element.findall(qn("w:p"))
		para = paras[0] if paras else etree.SubElement(hdr._element, qn("w:p"))
		run = etree.SubElement(para, qn("w:r"))
		pict = etree.SubElement(run, qn("w:pict"))
		pict.append(etree.fromstring(_SHAPETYPE_XML))
		pict.append(etree.fromstring(_shape_xml(text, color_hex)))


# -----------------------------------------------------------------------
# COM helpers (mirrors _word.py)
# -----------------------------------------------------------------------
def _convert_doc_to_docx(doc_path):
	out = os.path.splitext(doc_path)[0] + "_tmp.docx"
	word = doc = None
	try:
		word = win32com.client.Dispatch("Word.Application")
		word.Visible = False
		word.DisplayAlerts = False
		doc = word.Documents.Open(os.path.abspath(doc_path),
								  ReadOnly=True, AddToRecentFiles=False)
		doc.SaveAs2(os.path.abspath(out), FileFormat=16)
		return out
	finally:
		if doc:
			try: doc.Close(SaveChanges=False)
			except Exception: pass
		if word:
			try: word.Quit()
			except Exception: pass
		_kill_word()


def _pdf_via_com(docx_path):
	pdf = os.path.splitext(docx_path)[0] + ".pdf"
	word = doc = None
	try:
		word = win32com.client.Dispatch("Word.Application")
		word.Visible = False
		word.DisplayAlerts = False
		doc = word.Documents.Open(os.path.abspath(docx_path),
								  ReadOnly=True, AddToRecentFiles=False)
		doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)
	finally:
		if doc:
			try: doc.Close(SaveChanges=False)
			except Exception: pass
		if word:
			try: word.Quit()
			except Exception: pass
		_kill_word()


def _next_available(base, suffix, ext):
	c, p, n = f"{base}{suffix}{ext}", f"{base}{suffix}.pdf", 1
	while os.path.exists(c) or os.path.exists(p):
		c, p = f"{base}{suffix}({n}){ext}", f"{base}{suffix}({n}).pdf"
		n += 1
	return c


def _kill_word():
	try:
		subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], check=False,
					   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
					   creationflags=_NO_WINDOW)
	except Exception:
		pass


if __name__ == "__main__":
	if len(sys.argv) < 3:
		print("Usage: python _xword.py input.docx TEXT"); sys.exit(1)
	print("Saved:", add_word_watermark(sys.argv[1], sys.argv[2]))
