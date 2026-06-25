"""PDF watermarking for Watermark Lab.

Stamps the same tiled, diagonal, semi-transparent text watermark used for
PowerPoint onto every page of an existing PDF — without re-rasterising the page
(the original vector text/content is preserved). The watermark is drawn as an
image XObject carrying an explicit soft mask (/SMask) so the page shows through
the semi-transparent text.

Style parity with _powerpoint.add_watermark:
  * Segoe UI font
  * Default colour #A6A6A6, 70% transparency
  * Tiled block of repeated text, rotated -30 degrees (diagonal)

Pure Python: Pillow renders the overlay, pypdf composites it. No COM / Office.
The output is written via a local temp file then copied to the destination so
OneDrive-synced folders save reliably (mirrors _xpowerpoint).
"""
import os
import shutil
import tempfile
import zlib

from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader, PdfWriter
from pypdf.generic import (
	ArrayObject, DecodedStreamObject, DictionaryObject, NameObject, NumberObject,
)

PDF_EXTS = {".pdf"}

# Render resolution for the overlay (DPI relative to the PDF page points).
# 110 DPI keeps the tiled text crisp on screen and in print while making the
# rasterised overlay (and its BICUBIC rotate) a fraction of the 150-DPI cost.
_OVERLAY_DPI = 110.0

# Baseline: PowerPoint uses Segoe UI 18pt on a ~960pt-wide slide. Scale the
# font to the page width so tiling density looks the same across page sizes.
_PPT_FONT_PT = 18.0
_PPT_SLIDE_W_PT = 960.0


def _find_font(font_size: int) -> ImageFont.ImageFont:
	"""Return Segoe UI (preferred) at the given pixel size, with fallbacks."""
	for path in (
		r"C:\Windows\Fonts\segoeui.ttf",
		r"C:\Windows\Fonts\arial.ttf",
		r"C:\Windows\Fonts\calibri.ttf",
	):
		try:
			if os.path.isfile(path):
				return ImageFont.truetype(path, font_size)
		except Exception:
			pass
	return ImageFont.load_default()


def _next_available(base: str, suffix: str, ext: str) -> str:
	"""base+suffix+ext, appending (1), (2)... if the file already exists."""
	candidate = f"{base}{suffix}{ext}"
	n = 1
	while os.path.exists(candidate):
		candidate = f"{base}{suffix}({n}){ext}"
		n += 1
	return candidate


def _make_overlay_rgba(w_px: int, h_px: int, text: str, color_rgb: int,
					   transparency: float, font_size: int) -> Image.Image:
	"""Render a tiled, diagonal, semi-transparent block of text covering a
	w_px-by-h_px page. Mirrors _video._make_watermark_png so the look matches."""
	diag = int((w_px ** 2 + h_px ** 2) ** 0.5)
	canvas = diag + max(w_px, h_px)

	img = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
	draw = ImageDraw.Draw(img)
	font = _find_font(font_size)

	r = (color_rgb >> 16) & 0xFF
	g = (color_rgb >> 8) & 0xFF
	b = color_rgb & 0xFF
	alpha = int(round(255 * (1.0 - max(0.0, min(1.0, transparency)))))
	fill = (r, g, b, alpha)

	bbox = draw.textbbox((0, 0), text, font=font)
	tw = max(1, bbox[2] - bbox[0])
	th = max(1, bbox[3] - bbox[1])

	gap_x = max(font_size, tw // 3)
	gap_y = th * 3
	step_x = tw + gap_x
	step_y = th + gap_y

	# draw.text() is by far the costliest call here, so render the text once
	# into a single step-sized tile and blit that tile across the canvas
	# instead of re-rendering the glyphs thousands of times. The tiles are
	# spaced (step > text size) so they never overlap, which makes a plain
	# paste a pixel-for-pixel match for the original per-tile draw loop.
	tile = Image.new("RGBA", (step_x, step_y), (0, 0, 0, 0))
	ImageDraw.Draw(tile).text((0, 0), text, font=font, fill=fill)

	row = 0
	for y in range(0, canvas, step_y):
		offset = (step_x // 2) if (row % 2) else 0
		for x in range(-step_x, canvas, step_x):
			img.paste(tile, (x + offset, y))
		row += 1

	rotated = img.rotate(30, resample=Image.BICUBIC, expand=False)
	cx = (canvas - w_px) // 2
	cy = (canvas - h_px) // 2
	return rotated.crop((cx, cy, cx + w_px, cy + h_px))


def _build_overlay_xobject(writer: PdfWriter, w_pt: float, h_pt: float,
						   text: str, color_rgb: int, transparency: float):
	"""Render the watermark for a given page size and add it to the writer as a
	masked image XObject. Returns the XObject reference, which can be shared by
	every page of the same size (the expensive raster work happens here, once).
	"""
	w_px = max(1, int(round(w_pt / 72.0 * _OVERLAY_DPI)))
	h_px = max(1, int(round(h_pt / 72.0 * _OVERLAY_DPI)))

	# Scale the Segoe UI size from the PowerPoint baseline to this page width.
	font_px = max(12, int(round(_PPT_FONT_PT * (w_pt / _PPT_SLIDE_W_PT)
								* (_OVERLAY_DPI / 72.0))))

	rgba = _make_overlay_rgba(w_px, h_px, text, color_rgb, transparency, font_px)
	rgb_bytes = rgba.convert("RGB").tobytes()
	alpha_bytes = rgba.getchannel("A").tobytes()

	# Soft mask (alpha) as a DeviceGray image XObject.
	smask = DecodedStreamObject()
	smask.set_data(zlib.compress(alpha_bytes))
	smask[NameObject("/Type")] = NameObject("/XObject")
	smask[NameObject("/Subtype")] = NameObject("/Image")
	smask[NameObject("/Width")] = NumberObject(w_px)
	smask[NameObject("/Height")] = NumberObject(h_px)
	smask[NameObject("/BitsPerComponent")] = NumberObject(8)
	smask[NameObject("/ColorSpace")] = NameObject("/DeviceGray")
	smask[NameObject("/Filter")] = NameObject("/FlateDecode")
	smask_ref = writer._add_object(smask)

	# The colour image, masked by the alpha SMask.
	image = DecodedStreamObject()
	image.set_data(zlib.compress(rgb_bytes))
	image[NameObject("/Type")] = NameObject("/XObject")
	image[NameObject("/Subtype")] = NameObject("/Image")
	image[NameObject("/Width")] = NumberObject(w_px)
	image[NameObject("/Height")] = NumberObject(h_px)
	image[NameObject("/BitsPerComponent")] = NumberObject(8)
	image[NameObject("/ColorSpace")] = NameObject("/DeviceRGB")
	image[NameObject("/Filter")] = NameObject("/FlateDecode")
	image[NameObject("/SMask")] = smask_ref
	return writer._add_object(image)


def _overlay_page(writer: PdfWriter, page, image_ref, w_pt: float,
				  h_pt: float) -> None:
	"""Paint a pre-built watermark XObject over a single page, preserving its
	content. Cheap: it only registers the shared image and appends a tiny
	content stream, so it can run for every page without re-rendering."""
	# Register the image in the page's resources under a unique name.
	resources = page[NameObject("/Resources")]
	if "/XObject" not in resources:
		resources[NameObject("/XObject")] = DictionaryObject()
	xobjects = resources["/XObject"]
	name = "/WMKwm"
	xobjects[NameObject(name)] = image_ref

	# Content stream that paints the image across the full page box.
	content = (
		f"q {w_pt:.4f} 0 0 {h_pt:.4f} 0 0 cm {name} Do Q"
	).encode("latin-1")
	stream = DecodedStreamObject()
	stream.set_data(content)
	stream_ref = writer._add_object(stream)

	contents = page.get("/Contents")
	if contents is None:
		page[NameObject("/Contents")] = stream_ref
	elif isinstance(contents, ArrayObject):
		contents.append(stream_ref)
	else:
		page[NameObject("/Contents")] = ArrayObject([contents, stream_ref])


def add_pdf_watermark(pdf_path, watermark_text, color_rgb=0xA6A6A6,
					  transparency=0.70, export_pdf=False, progress_cb=None):
	"""Watermark every page of a PDF with a tiled diagonal text overlay.

	color_rgb: integer RGB in 0xRRGGBB form.
	transparency: float 0.0 (opaque) to 1.0 (fully transparent).
	export_pdf: ignored — the source already is a PDF (kept for signature
		parity with the other engines so the GUI can call them uniformly).
	progress_cb: optional callable(pages_done, total_pages) for UI updates so a
		long multi-page document doesn't look frozen.

	Returns the output path: <name>_watermarked.pdf in the source folder.
	"""
	pdf_path = os.path.abspath(pdf_path)
	if not os.path.isfile(pdf_path):
		raise FileNotFoundError(pdf_path)

	base, ext = os.path.splitext(pdf_path)
	final_out = _next_available(base, "_watermarked", ext or ".pdf")

	try:
		reader = PdfReader(pdf_path)
	except Exception as exc:  # noqa: BLE001
		raise RuntimeError(f"Could not read the PDF: {exc}")

	if getattr(reader, "is_encrypted", False):
		# Try an empty-password unlock; bail clearly if it's truly protected.
		try:
			if reader.decrypt("") == 0:
				raise RuntimeError(
					"This PDF is password-protected and can't be watermarked.")
		except Exception:
			raise RuntimeError(
				"This PDF is password-protected and can't be watermarked.")

	writer = PdfWriter()
	writer.append(reader)

	# Rendering the overlay is the expensive part, so build it once per unique
	# page size and reuse the XObject for every page that shares that size —
	# uniform documents (the common case) only pay the raster cost a single
	# time. Keyed by rounded points so near-identical sizes collapse together.
	overlay_cache: dict = {}
	total = len(writer.pages)
	for done, page in enumerate(writer.pages, 1):
		w_pt = float(page.mediabox.width)
		h_pt = float(page.mediabox.height)
		if w_pt > 0 and h_pt > 0:
			key = (round(w_pt, 1), round(h_pt, 1))
			image_ref = overlay_cache.get(key)
			if image_ref is None:
				image_ref = _build_overlay_xobject(
					writer, w_pt, h_pt, watermark_text, color_rgb, transparency)
				overlay_cache[key] = image_ref
			_overlay_page(writer, page, image_ref, w_pt, h_pt)
		if progress_cb is not None:
			try:
				progress_cb(done, total)
			except Exception:  # noqa: BLE001
				pass  # progress reporting must never break the job

	# OneDrive-safe: write to a local temp file, then copy to the destination.
	work_dir = tempfile.mkdtemp(prefix="wlx_pdf_")
	try:
		tmp_out = os.path.join(work_dir, "out.pdf")
		with open(tmp_out, "wb") as fh:
			writer.write(fh)
		shutil.copy2(tmp_out, final_out)
		return final_out
	finally:
		shutil.rmtree(work_dir, ignore_errors=True)
