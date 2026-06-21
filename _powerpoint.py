import os
import subprocess
import sys
import win32com.client

# PowerPoint constants
PP_SAVE_AS_OPEN_XML_PRESENTATION = 24
PP_SAVE_AS_PDF = 32
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
PP_ALIGN_CENTER = 2
PP_AUTO_SIZE_NONE = 0


def add_watermark(ppt_path, watermark_text, color_rgb=0xA6A6A6, transparency=0.70,
                  export_pdf=False):
    """Add a tiled diagonal watermark to every slide of a PowerPoint file.

    color_rgb: integer RGB in 0xRRGGBB form (note: PowerPoint COM uses BGR ordering).
    transparency: float 0.0 (opaque) to 1.0 (fully transparent).
    export_pdf: if True, also export the watermarked presentation as a PDF with
        the same base name.
    """
    # PowerPoint requires an absolute path
    ppt_path = os.path.abspath(ppt_path)
    if not os.path.isfile(ppt_path):
        raise FileNotFoundError(ppt_path)

    base, ext = os.path.splitext(ppt_path)
    output_path = _next_available(base, "_watermarked", ext or ".pptx")

    # Launch PowerPoint via dynamic (late) COM binding so the watermark flow
    # works in the frozen build without the win32com gen_py cache.
    powerpoint = _ensure_powerpoint()
    # Keep the application window hidden during processing.
    try:
        powerpoint.Visible = 0
    except Exception:
        # Some PowerPoint builds reject Visible=0 unless WithWindow is False on Open.
        pass

    presentation = None
    try:
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)

        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight

        # Build a tiled block of text covering the whole slide.
        # Repeats per line and number of lines are generous so that after
        # rotation the watermark still fills every corner.
        line_text = (watermark_text + "        ") * 10
        block_text = "\r".join([line_text] * 18)

        # Oversize the textbox so rotation never leaves blank corners.
        box_w = slide_width * 1.8
        box_h = slide_height * 1.8

        for slide in presentation.Slides:
            shape = slide.Shapes.AddTextbox(
                Orientation=MSO_TEXT_ORIENTATION_HORIZONTAL,
                Left=(slide_width - box_w) / 2,
                Top=(slide_height - box_h) / 2,
                Width=box_w,
                Height=box_h,
            )

            shape.TextFrame.AutoSize = PP_AUTO_SIZE_NONE
            shape.TextFrame.WordWrap = False
            shape.TextFrame.MarginLeft = 0
            shape.TextFrame.MarginRight = 0
            shape.TextFrame.MarginTop = 0
            shape.TextFrame.MarginBottom = 0

            text_range = shape.TextFrame.TextRange
            text_range.Text = block_text

            # Font: Segoe UI 18, gray #A6A6A6 at 70% transparency.
            # Transparency lives on Font2.Fill (TextFrame2), not on the
            # legacy Font object exposed via TextFrame.TextRange.
            font2 = shape.TextFrame2.TextRange.Font
            font2.Name = "Segoe UI"
            font2.Size = 18
            font2.Fill.Visible = True
            font2.Fill.Solid()
            # PowerPoint's ForeColor.RGB uses OLE BGR byte ordering, so convert
            # the standard 0xRRGGBB value supplied by the caller.
            r = (color_rgb >> 16) & 0xFF
            g = (color_rgb >> 8) & 0xFF
            b = color_rgb & 0xFF
            font2.Fill.ForeColor.RGB = (b << 16) | (g << 8) | r
            font2.Fill.Transparency = float(transparency)
            

            # Generous vertical spacing between rows
            text_range.ParagraphFormat.Alignment = PP_ALIGN_CENTER
            try:
                text_range.ParagraphFormat.SpaceWithin = 3.0  # line spacing
                text_range.ParagraphFormat.SpaceBefore = 12
                text_range.ParagraphFormat.SpaceAfter = 12
            except Exception:
                pass

            shape.Line.Visible = False
            shape.Rotation = 330  # diagonal (-30 degrees)

        # Save output as .pptx
        presentation.SaveAs(output_path, FileFormat=PP_SAVE_AS_OPEN_XML_PRESENTATION)
        print(f"Saved watermarked file as: {output_path}")

        if export_pdf:
            pdf_path = os.path.splitext(output_path)[0] + ".pdf"
            presentation.SaveAs(pdf_path, FileFormat=PP_SAVE_AS_PDF)
            print(f"Saved watermarked PDF as: {pdf_path}")

        return output_path
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            powerpoint.Quit()
        except Exception:
            pass
        # Force-terminate any lingering PowerPoint process so the operation
        # is guaranteed to complete and release file locks.
        _kill_powerpoint()


def _next_available(base, suffix, ext):
    """Return base+suffix+ext, appending (1), (2)... if the file exists.

    Also avoids collision with the matching .pdf companion so the pair
    stays in sync when export_pdf is requested.
    """
    candidate = f"{base}{suffix}{ext}"
    pdf_candidate = f"{base}{suffix}.pdf"
    n = 1
    while os.path.exists(candidate) or os.path.exists(pdf_candidate):
        candidate = f"{base}{suffix}({n}){ext}"
        pdf_candidate = f"{base}{suffix}({n}).pdf"
        n += 1
    return candidate


def _kill_powerpoint():
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", "POWERPNT.EXE"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        pass


def _ensure_powerpoint():
    """Return a PowerPoint.Application COM object using pure dynamic (late)
    binding.

    Dynamic dispatch resolves every property and method by name at runtime and
    never touches the win32com gen_py early-binding cache. This avoids the
    frozen-build failures:
      * ModuleNotFoundError: No module named 'win32com.gen_py'
      * COM errors generating wrapper classes for sub-objects such as
        PowerPoint.Application.Presentations
    The full watermark chain (Presentations, TextFrame2, Fill.Transparency,
    Rotation) works correctly under dynamic dispatch.
    """
    import win32com.client.dynamic
    return win32com.client.dynamic.Dispatch("PowerPoint.Application")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python watermark_com.py input.pptx \"WATERMARK TEXT\"")
        sys.exit(1)

    ppt_file = sys.argv[1]
    text = sys.argv[2]

    add_watermark(ppt_file, text)
