"""Word watermarking via python-docx + canonical ECMA-376 VML injection.

Injects the exact header XML that Word's Design -> Watermark writes,
per the ECMA-376 Office Open XML specification. No COM shape positioning.
"""
import os, subprocess, sys, zipfile, shutil, tempfile
import pythoncom
import win32com.client
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

WORD_EXTS = {".docx", ".doc"}
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# Namespaces used in Word header XML
_NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v":   "urn:schemas-microsoft-com:vml",
    "o":   "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
}

# -----------------------------------------------------------------------
# Canonical shapetype -- identical to what Word writes (ECMA-376 / VML spec)
# -----------------------------------------------------------------------
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

def _shape_xml(text, color_hex):
    """Build the watermark shape with explicit font size - no fitshape recalculation needed."""
    # Shape is 527.85pt wide. Scale font size so text fills ~90% of that width.
    # Empirically: Segoe UI bold ~0.55pt width-per-point per character.
    char_count  = max(len(text), 1)
    font_size   = max(30.0, min(90.0, (527.85 * 0.90) / (char_count * 0.55)))
    font_size   = round(font_size, 1)

    tp_style = (
        f"font-family:'Segoe UI';font-size:{font_size}pt;font-weight:bold;"
        f"color:{color_hex}"
    )
    return (
        f'<v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2051" '
        f'type="#_x0000_t136" fillcolor="{color_hex}" '
        'style="position:absolute;margin-left:0;margin-top:0;'
        'width:527.85pt;height:131.95pt;z-index:-251654144;'
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
        f'<v:textpath style="{tp_style}" string="{_xe(text)}" fitshape="t"/>'
        '<v:imagedata o:relid="" o:title=""/>'
        '<o:lock v:ext="edit" position="t"/>'
        '<w10:wrap w10:type="none"/>'
        '<w10:anchorlock/>'
        '</v:shape>'
    )

def _xe(t):
    return (t.replace("&","&amp;").replace("<","&lt;")
             .replace(">","&gt;").replace('"',"&quot;"))

# -----------------------------------------------------------------------
# Public API
# -----------------------------------------------------------------------

def add_word_watermark(doc_path, watermark_text, color_rgb=0xA6A6A6,
                       transparency=0.70, export_pdf=False):
    doc_path = os.path.abspath(doc_path)
    if not os.path.isfile(doc_path):
        raise FileNotFoundError(doc_path)

    # Initialise COM for this thread. add_word_watermark runs on a background
    # worker thread in the GUI; the .doc conversion and PDF export use Word
    # COM, which must be initialised per-thread.
    pythoncom.CoInitialize()
    try:
        base, ext = os.path.splitext(doc_path)

        # .doc -> convert to docx first via COM
        if ext.lower() == ".doc":
            doc_path = _convert_doc_to_docx(doc_path)
            base = os.path.splitext(doc_path)[0]

        output_path = _next_available(base, "_watermarked", ".docx")
        color_hex   = "#{:06x}".format(color_rgb)

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

        # Remove any previous watermark shapes/shapetypes
        for tag in (f"{{{ns_v}}}shape", f"{{{ns_v}}}shapetype"):
            for el in hdr._element.findall(f".//{tag}"):
                p = el.getparent()
                if p is not None:
                    p.remove(el)

        # Use or create the first paragraph in the header
        paras = hdr._element.findall(qn("w:p"))
        para  = paras[0] if paras else etree.SubElement(hdr._element, qn("w:p"))

        run  = etree.SubElement(para, qn("w:r"))
        pict = etree.SubElement(run,  qn("w:pict"))
        pict.append(etree.fromstring(_SHAPETYPE_XML))
        pict.append(etree.fromstring(_shape_xml(text, color_hex)))


# -----------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------

def _convert_doc_to_docx(doc_path):
    out  = os.path.splitext(doc_path)[0] + "_tmp.docx"
    word = doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(doc_path),
                                  ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(os.path.abspath(out), FileFormat=16)
        return out
    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        _kill_word()

def _normalize_via_com(docx_path):
    """Open the saved .docx in Word and re-save it.

    This forces Word's layout engine to recalculate the watermark's
    fitshape auto-sizing -- equivalent to manually selecting 'Auto'
    in the watermark font size dropdown.
    """
    word = doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(docx_path),
                                  ReadOnly=False, AddToRecentFiles=False)
        doc.Save()
    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        _kill_word()

def _pdf_via_com(docx_path):
    pdf  = os.path.splitext(docx_path)[0] + ".pdf"
    word = doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False; word.DisplayAlerts = False
        doc = word.Documents.Open(os.path.abspath(docx_path),
                                  ReadOnly=True, AddToRecentFiles=False)
        doc.SaveAs2(os.path.abspath(pdf), FileFormat=17)
    finally:
        if doc:
            try: doc.Close(SaveChanges=False)
            except: pass
        if word:
            try: word.Quit()
            except: pass
        _kill_word()

def _next_available(base, suffix, ext):
    c, p, n = f"{base}{suffix}{ext}", f"{base}{suffix}.pdf", 1
    while os.path.exists(c) or os.path.exists(p):
        c, p = f"{base}{suffix}({n}){ext}", f"{base}{suffix}({n}).pdf"
        n += 1
    return c

def _kill_word():
    try:
        subprocess.run(["taskkill","/F","/IM","WINWORD.EXE"], check=False,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=_NO_WINDOW)
    except: pass

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python _word.py input.docx TEXT"); sys.exit(1)
    print("Saved:", add_word_watermark(sys.argv[1], sys.argv[2]))
