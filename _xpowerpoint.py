"""Experimental PowerPoint watermarking wrapper for Watermark Lab X.

Fixes a real bug: PowerPoint's COM SaveAs SILENTLY FAILS when the target is a
OneDrive-synced folder — it reports success but never writes the file. (Word
and ffmpeg use normal file I/O and are unaffected, which is why .docx/.mp4
worked but .pptx silently produced nothing.)

Strategy: run the proven _powerpoint.add_watermark on a TEMP copy (PowerPoint
saves happily to a local temp path), then copy the watermarked .pptx (and .pdf)
to the real destination with normal Python file I/O, which OneDrive handles
fine. The legacy _powerpoint.py is left completely untouched.
"""
import os
import shutil
import tempfile

from _powerpoint import add_watermark as _add_watermark_real


def _next_available(base: str, suffix: str, ext: str) -> str:
	"""Return base+suffix+ext, appending (1), (2)... if a file (or its .pdf
	companion) already exists — mirrors _powerpoint._next_available."""
	candidate = f"{base}{suffix}{ext}"
	pdf_candidate = f"{base}{suffix}.pdf"
	n = 1
	while os.path.exists(candidate) or os.path.exists(pdf_candidate):
		candidate = f"{base}{suffix}({n}){ext}"
		pdf_candidate = f"{base}{suffix}({n}).pdf"
		n += 1
	return candidate


def add_watermark(ppt_path, watermark_text, color_rgb=0xA6A6A6,
				  transparency=0.70, export_pdf=False):
	"""Watermark a PowerPoint file, saving reliably even on OneDrive paths.

	Signature-compatible with _powerpoint.add_watermark. Returns the final
	output path in the SAME folder as the source.
	"""
	ppt_path = os.path.abspath(ppt_path)
	if not os.path.isfile(ppt_path):
		raise FileNotFoundError(ppt_path)

	dest_dir = os.path.dirname(ppt_path)
	base_name, ext = os.path.splitext(os.path.basename(ppt_path))
	ext = ext or ".pptx"

	# Final destination path(s) in the real (possibly OneDrive) folder.
	final_base = os.path.join(dest_dir, base_name)
	final_out = _next_available(final_base, "_watermarked", ext)
	final_pdf = os.path.splitext(final_out)[0] + ".pdf"

	work_dir = tempfile.mkdtemp(prefix="wlx_ppt_")
	try:
		# Copy the source into a clean local temp folder.
		local_src = os.path.join(work_dir, os.path.basename(ppt_path))
		shutil.copy2(ppt_path, local_src)

		# Run the REAL watermarker against the temp copy — PowerPoint saves to
		# temp without the OneDrive silent-failure problem.
		temp_out = _add_watermark_real(
			local_src, watermark_text, color_rgb=color_rgb,
			transparency=transparency, export_pdf=export_pdf,
		)
		if not temp_out or not os.path.isfile(temp_out):
			raise RuntimeError(
				"PowerPoint did not produce an output file (the source may be "
				"protected or restricted for editing).")

		# Copy the watermarked .pptx to the real destination via normal I/O.
		shutil.copy2(temp_out, final_out)

		# Copy the PDF companion if it was produced.
		if export_pdf:
			temp_pdf = os.path.splitext(temp_out)[0] + ".pdf"
			if os.path.isfile(temp_pdf):
				shutil.copy2(temp_pdf, final_pdf)

		return final_out
	finally:
		shutil.rmtree(work_dir, ignore_errors=True)
