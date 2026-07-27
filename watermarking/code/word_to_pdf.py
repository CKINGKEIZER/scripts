"""
any_to_pdf.py  (formerly word_to_pdf.py)
----------------------------------------
Batch-convert a folder (or a selection of files) to PDF.

Supported inputs and the engine used for each:

  Word        .doc .docx .docm .rtf .odt          Microsoft Word      (native, high fidelity)
  Excel       .xls .xlsx .xlsm .xlsb .csv .ods    Microsoft Excel     (native, high fidelity)
  PowerPoint  .ppt .pptx .pptm .odp               Microsoft PowerPoint(native, high fidelity)
  Outlook     .msg .eml                           Outlook + Word      (native, high fidelity)
  Images      .jpg .jpeg .png .gif .bmp .tif      Pillow              (full-resolution embed)
              .tiff .webp
  HTML        .htm .html .mht .mhtml              Microsoft Word      (rendered; no JS, partial CSS)
  Text/code   .txt .css .xml .json .log .md .js   reportlab           (monospace layout of the source)
              .ts .py .ini .yaml .yml .bat .sh
              .sql .java .c .cpp .h and similar
  Unknown     anything else                       text if decodable, else skipped

Office and Outlook formats require the matching Microsoft Office app installed.
Images and text/code do not require Office.

Engine module - GUI is provided by launcher.py Convert to PDF tab.
CLI:  python word_to_pdf.py <file-or-folder> [more ...]
"""

import os
import sys
import tempfile


# ── EXTENSION ROUTING ─────────────────────────────────────────────────────────

WORD_EXT    = {".doc", ".docx", ".docm", ".rtf", ".odt"}
EXCEL_EXT   = {".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods"}
PPT_EXT     = {".ppt", ".pptx", ".pptm", ".odp"}
OUTLOOK_EXT = {".msg", ".eml"}
IMAGE_EXT   = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff", ".webp"}
HTML_EXT    = {".htm", ".html", ".mht", ".mhtml"}
TEXT_EXT    = {
    ".txt", ".css", ".xml", ".json", ".log", ".md", ".js", ".ts", ".py",
    ".ini", ".yaml", ".yml", ".bat", ".sh", ".sql", ".java", ".c", ".cpp",
    ".h", ".hpp", ".cs", ".go", ".rb", ".php", ".pl", ".r", ".vb", ".tsv",
    ".conf", ".cfg", ".env", ".toml", ".rst", ".tex",
}

ALL_KNOWN = (
    WORD_EXT | EXCEL_EXT | PPT_EXT | OUTLOOK_EXT
    | IMAGE_EXT | HTML_EXT | TEXT_EXT
)


def route(ext):
    ext = ext.lower()
    if ext in WORD_EXT:    return "word"
    if ext in EXCEL_EXT:   return "excel"
    if ext in PPT_EXT:     return "ppt"
    if ext in OUTLOOK_EXT: return "outlook"
    if ext in IMAGE_EXT:   return "image"
    if ext in HTML_EXT:    return "html"
    if ext in TEXT_EXT:    return "text"
    return "unknown"


# ── TEXT READING ──────────────────────────────────────────────────────────────

def read_text_file(path):
    """Return decoded text, or None if the file looks binary."""
    with open(path, "rb") as f:
        raw = f.read()
    if b"\x00" in raw[:4096]:
        return None
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return None


# ── NON-OFFICE CONVERTERS ─────────────────────────────────────────────────────

def convert_image(src, pdf):
    """Embed an image at full resolution, one page. Flatten transparency onto white."""
    from PIL import Image
    im = Image.open(src)
    try:
        im.seek(0)  # first frame for multi-frame formats
    except (EOFError, ValueError):
        pass
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        bg = Image.new("RGB", rgba.size, (255, 255, 255))
        bg.paste(rgba, mask=rgba.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    im.save(pdf, "PDF", resolution=150.0)


def convert_text(src, pdf):
    """Render a text/code file as monospace on A4. Returns True, or False if binary."""
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    from reportlab.pdfbase.pdfmetrics import stringWidth

    text = read_text_file(src)
    if text is None:
        return False

    text = text.replace("\t", "    ")
    width, height = A4
    margin = 15 * mm
    font, size = "Courier", 9
    leading = size * 1.25
    usable_w = width - 2 * margin
    char_w = stringWidth("0", font, size)
    max_chars = max(1, int(usable_w / char_w))
    top = height - margin
    bottom = margin

    c = canvas.Canvas(pdf, pagesize=A4)
    c.setFont(font, size)
    y = top
    for raw_line in text.split("\n"):
        raw_line = raw_line.rstrip("\r")
        chunks = [raw_line[i:i + max_chars] for i in range(0, len(raw_line), max_chars)] or [""]
        for chunk in chunks:
            if y < bottom:
                c.showPage()
                c.setFont(font, size)
                y = top
            c.drawString(margin, y, chunk)
            y -= leading
    c.showPage()
    c.save()
    return True


# ── OFFICE APP MANAGER (lazy, one instance per app, reused across the batch) ───

class OfficeApps:
    """Creates Office/Outlook COM instances on demand and reuses them."""

    def __init__(self, log):
        self.log = log
        self._word = None
        self._excel = None
        self._ppt = None
        self._outlook = None

    def word(self):
        if self._word is None:
            import win32com.client
            self._word = win32com.client.DispatchEx("Word.Application")
            self._word.Visible = False
            self._word.DisplayAlerts = 0
        return self._word

    def excel(self):
        if self._excel is None:
            import win32com.client
            self._excel = win32com.client.DispatchEx("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            try:
                self._excel.AskToUpdateLinks = False
            except Exception:
                pass
        return self._excel

    def ppt(self):
        if self._ppt is None:
            import win32com.client
            self._ppt = win32com.client.DispatchEx("PowerPoint.Application")
            # PowerPoint refuses Visible=False on many builds, so leave it and
            # minimise instead.
            try:
                self._ppt.WindowState = 2  # ppWindowMinimized
            except Exception:
                pass
        return self._ppt

    def outlook(self):
        if self._outlook is None:
            import win32com.client
            self._outlook = win32com.client.DispatchEx("Outlook.Application")
        return self._outlook

    def close(self):
        for app, name in (
            (self._word, "Word"),
            (self._excel, "Excel"),
            (self._ppt, "PowerPoint"),
            (self._outlook, "Outlook"),
        ):
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass


# ── OFFICE CONVERTERS ─────────────────────────────────────────────────────────

def convert_word(apps, src, pdf):
    word = apps.word()
    doc = word.Documents.Open(os.path.abspath(src), ReadOnly=True)
    try:
        doc.SaveAs(os.path.abspath(pdf), FileFormat=17)  # 17 = wdFormatPDF
    finally:
        doc.Close(False)


def convert_html(apps, src, pdf):
    # Word opens HTML/MHT and renders it. Same path as Word documents.
    convert_word(apps, src, pdf)


def convert_excel(apps, src, pdf):
    excel = apps.excel()
    wb = excel.Workbooks.Open(os.path.abspath(src), ReadOnly=True, UpdateLinks=0)
    try:
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf))  # 0 = xlTypePDF
    finally:
        wb.Close(False)


def convert_ppt(apps, src, pdf):
    ppt = apps.ppt()
    pres = ppt.Presentations.Open(os.path.abspath(src), ReadOnly=True, WithWindow=False)
    try:
        try:
            pres.ExportAsFixedFormat(os.path.abspath(pdf), 2, 2)  # type=PDF, intent=print
        except Exception:
            pres.SaveAs(os.path.abspath(pdf), 32)  # 32 = ppSaveAsPDF fallback
    finally:
        pres.Close()


def convert_outlook(apps, src, pdf):
    """
    .msg / .eml -> PDF.
    Open the item in Outlook, save it as MHTML, then render that with Word.
    Requires both Outlook and Word.
    """
    outlook = apps.outlook()
    item = outlook.Session.OpenSharedItem(os.path.abspath(src))
    tmp_mht = None
    try:
        fd, tmp_mht = tempfile.mkstemp(suffix=".mht")
        os.close(fd)
        item.SaveAs(tmp_mht, 10)  # 10 = olMHTML
    finally:
        try:
            item.Close(1)  # 1 = olDiscard
        except Exception:
            pass
    try:
        convert_word(apps, tmp_mht, pdf)
    finally:
        if tmp_mht and os.path.isfile(tmp_mht):
            try:
                os.remove(tmp_mht)
            except Exception:
                pass


# ── DISPATCH ──────────────────────────────────────────────────────────────────

def convert_one(apps, src, pdf, log):
    """Route a single file. Returns True on success."""
    ext = os.path.splitext(src)[1].lower()
    kind = route(ext)

    if kind == "image":
        convert_image(src, pdf)
        return True
    if kind == "text":
        return convert_text(src, pdf)
    if kind == "word":
        convert_word(apps, src, pdf); return True
    if kind == "html":
        convert_html(apps, src, pdf); return True
    if kind == "excel":
        convert_excel(apps, src, pdf); return True
    if kind == "ppt":
        convert_ppt(apps, src, pdf); return True
    if kind == "outlook":
        convert_outlook(apps, src, pdf); return True

    # unknown: try text, else skip
    if convert_text(src, pdf):
        log(f"      (unknown extension {ext}, rendered as text)")
        return True
    return False


# ── FILE COLLECTION ───────────────────────────────────────────────────────────

def collect_from_folder(folder, recurse):
    files = []
    if recurse:
        for root, _dirs, names in os.walk(folder):
            for n in names:
                files.append(os.path.join(root, n))
    else:
        for n in os.listdir(folder):
            p = os.path.join(folder, n)
            if os.path.isfile(p):
                files.append(p)
    return files


def pdf_output_path(src, out_dir):
    base = os.path.splitext(os.path.basename(src))[0] + ".pdf"
    if out_dir:
        return os.path.join(out_dir, base)
    return os.path.splitext(os.path.abspath(src))[0] + ".pdf"


# ── BATCH RUNNER ──────────────────────────────────────────────────────────────

def run_batch(files, out_dir, log, progress=None):
    """
    Convert every file in `files`. Office apps are opened once and reused.
    Returns (done, skipped, failed_list).
    """
    import pythoncom
    pythoncom.CoInitialize()

    apps = OfficeApps(log)

    done = 0
    skipped = 0
    failed = []
    total = len(files)

    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    try:
        for i, src in enumerate(files, 1):
            name = os.path.basename(src)
            ext = os.path.splitext(src)[1].lower()
            if ext == ".pdf":
                # No output folder (loose files) -> nothing to tidy, leave it.
                if not out_dir:
                    log(f"  [{i}/{total}] skip (already PDF)  {name}")
                    skipped += 1
                    if progress: progress(i, total)
                    continue
                # Folder run: move the existing PDF into the pdf/ subfolder too,
                # so the source folder ends up clean. Never overwrite.
                dest = pdf_output_path(src, out_dir)
                if os.path.abspath(src) == os.path.abspath(dest):
                    log(f"  [{i}/{total}] already in output  {name}")
                    skipped += 1
                elif os.path.exists(dest):
                    log(f"  [{i}/{total}] skip (name already in pdf/)  {name}")
                    skipped += 1
                else:
                    try:
                        import shutil
                        shutil.move(src, dest)
                        log(f"  [{i}/{total}] moved (already PDF)  {name}  ->  pdf/{os.path.basename(dest)}")
                        done += 1
                    except Exception as e:
                        log(f"  [{i}/{total}] FAIL  {name}  :  {e}")
                        failed.append((name, str(e)))
                if progress: progress(i, total)
                continue

            pdf = pdf_output_path(src, out_dir)
            try:
                ok = convert_one(apps, src, pdf, log)
                if ok:
                    log(f"  [{i}/{total}] ok    {name}  ->  {os.path.basename(pdf)}")
                    done += 1
                else:
                    log(f"  [{i}/{total}] skip (unsupported)  {name}")
                    skipped += 1
            except Exception as e:
                log(f"  [{i}/{total}] FAIL  {name}  :  {e}")
                failed.append((name, str(e)))
            if progress:
                progress(i, total)
    finally:
        apps.close()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return done, skipped, failed


# ── ENTRY POINT (CLI) ────────────────────────────────

def main():
    """Convert files/folders passed on the command line. GUI lives in launcher.py."""
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if not args:
        print("Usage: python word_to_pdf.py <file-or-folder> [more ...]")
        print("Converts Word/Excel/PowerPoint/images/text/Outlook files to PDF.")
        return
    files = []
    for a in args:
        if os.path.isdir(a):
            files += collect_from_folder(a, recurse=True)
        elif os.path.isfile(a):
            files.append(os.path.abspath(a))
    if not files:
        print("No files found.")
        return
    print("Converting %d file(s)...\n" % len(files))
    done, skipped, failed = run_batch(files, None, print)
    print("\nDone. %d converted, %d skipped, %d failed." % (done, skipped, len(failed)))


if __name__ == "__main__":
    main()
