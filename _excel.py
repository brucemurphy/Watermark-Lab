"""Excel watermarking via COM.

Adds a diagonal text watermark to every worksheet's header/footer using
Excel's PageSetup — visible in Print Preview and when printed.
Also supports PDF export via Excel's ExportAsFixedFormat.
"""
import os
import subprocess
import sys
import win32com.client

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Excel COM constants
XL_FORMAT_XLSX = 51   # xlOpenXMLWorkbook
XL_TYPE_PDF    = 0    # xlTypePDF


def add_excel_watermark(file_path: str, watermark_text: str,
						color_rgb: int = 0xA6A6A6,
						transparency: float = 0.70,
						export_pdf: bool = False) -> str:
	"""Add a watermark header to every sheet and save as *_watermarked.xlsx.

	Excel does not support true behind-text watermarks like Word/PPT.
	The watermark is placed in the centre header section of every sheet's
	page setup — it appears in Print Preview, printed output, and PDF export.
	The header uses large bold grey text that mimics a watermark appearance.
	"""
	file_path = os.path.abspath(file_path)
	if not os.path.isfile(file_path):
		raise FileNotFoundError(file_path)

	base, _ = os.path.splitext(file_path)
	output_path = _next_available(base, "_watermarked", ".xlsx")

	# Build Excel header code:
	# &C = center section, &"font,Bold" = font, &XX = font size, &KrrggBB = color
	r = (color_rgb >> 16) & 0xFF
	g = (color_rgb >>  8) & 0xFF
	b =  color_rgb        & 0xFF
	hex_color = f"{r:02X}{g:02X}{b:02X}"
	header_str = f'&C&"Segoe UI,Bold"&36&K{hex_color}{watermark_text}'

	xl = wb = None
	try:
		xl = win32com.client.Dispatch("Excel.Application")
		xl.Visible       = False
		xl.DisplayAlerts = False
		xl.ScreenUpdating = False

		wb = xl.Workbooks.Open(file_path, ReadOnly=False, AddToMru=False)

		for sheet in wb.Worksheets:
			ps = sheet.PageSetup
			ps.CenterHeader = header_str
			ps.CenterFooter = ""

		wb.SaveAs(os.path.abspath(output_path), FileFormat=XL_FORMAT_XLSX)

		if export_pdf:
			pdf_path = os.path.splitext(output_path)[0] + ".pdf"
			wb.ExportAsFixedFormat(
				Type=XL_TYPE_PDF,
				Filename=os.path.abspath(pdf_path),
				IncludeDocProperties=True,
				IgnorePrintAreas=False,
			)

		return output_path

	finally:
		if wb is not None:
			try:
				wb.Close(SaveChanges=False)
			except Exception:
				pass
		if xl is not None:
			try:
				xl.Quit()
			except Exception:
				pass
		_kill_excel()


def _next_available(base: str, suffix: str, ext: str) -> str:
	c = f"{base}{suffix}{ext}"
	p = f"{base}{suffix}.pdf"
	n = 1
	while os.path.exists(c) or os.path.exists(p):
		c = f"{base}{suffix}({n}){ext}"
		p = f"{base}{suffix}({n}).pdf"
		n += 1
	return c


def _kill_excel() -> None:
	try:
		subprocess.run(
			["taskkill", "/F", "/IM", "EXCEL.EXE"],
			check=False,
			stdout=subprocess.DEVNULL,
			stderr=subprocess.DEVNULL,
			creationflags=_NO_WINDOW,
		)
	except Exception:
		pass


if __name__ == "__main__":
	if len(sys.argv) < 3:
		print("Usage: python _excel.py input.xlsx \"TEXT\"")
		sys.exit(1)
	print("Saved:", add_excel_watermark(sys.argv[1], sys.argv[2]))
