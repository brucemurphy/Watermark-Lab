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
_OVERLAY_DPI = 150.0

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

	row = 0
	for y in range(0, canvas, step_y):
		offset = (step_x // 2) if (row % 2) else 0
		for x in range(-step_x, canvas, step_x):
			draw.text((x + offset, y), text, font=font, fill=fill)
		row += 1

	rotated = img.rotate(30, resample=Image.BICUBIC, expand=False)
	cx = (canvas - w_px) // 2
	cy = (canvas - h_px) // 2
	return rotated.crop((cx, cy, cx + w_px, cy + h_px))


def _overlay_page(writer: PdfWriter, page, text: str, color_rgb: int,
				  transparency: float) -> None:
	"""Composite the watermark over a single page, preserving its content."""
	w_pt = float(page.mediabox.width)
	h_pt = float(page.mediabox.height)
	if w_pt <= 0 or h_pt <= 0:
		return

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
	image_ref = writer._add_object(image)

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
					  transparency=0.70, export_pdf=False):
	"""Watermark every page of a PDF with a tiled diagonal text overlay.

	color_rgb: integer RGB in 0xRRGGBB form.
	transparency: float 0.0 (opaque) to 1.0 (fully transparent).
	export_pdf: ignored — the source already is a PDF (kept for signature
		parity with the other engines so the GUI can call them uniformly).

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

	for page in writer.pages:
		_overlay_page(writer, page, watermark_text, color_rgb, transparency)

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
