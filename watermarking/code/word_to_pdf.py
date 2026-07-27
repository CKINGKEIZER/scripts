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

Run:  python any_to_pdf.py
"""

import os
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext


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

    needs_office = any(
        route(os.path.splitext(f)[1]) in ("word", "excel", "ppt", "outlook", "html")
        for f in files
    )
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
                log(f"  [{i}/{total}] skip (already PDF)  {name}")
                skipped += 1
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


# ── GUI ───────────────────────────────────────────────────────────────────────

NAVY = "#0B2340"
GOLD = "#C19A50"
BG   = "#F7F7F5"
CARD = "#FFFFFF"
SUB  = "#7A8A98"
TEXT = "#0B1E30"
LOGB = "#0D1B2A"
LOGF = "#C8D8E8"
FF   = "Segoe UI"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Kumulus Partners — Convert to PDF")
        self.geometry("860x680")
        self.minsize(760, 600)
        self.configure(bg=BG)

        self.items = []      # selected file paths
        self.out_dir = None  # optional single output folder

        self._build_header()
        self._build_body()

    # ── header ──
    def _build_header(self):
        hdr = tk.Frame(self, bg=NAVY, height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="KUMULUS PARTNERS", bg=NAVY, fg="white",
                 font=(FF, 14, "bold")).pack(side="left", padx=24, anchor="s", pady=(0, 4))
        tk.Frame(hdr, bg="#2A4A6A", width=1).pack(side="left", fill="y", pady=14, padx=14)
        tk.Label(hdr, text="Convert to PDF", bg=NAVY, fg="#7A9AB8",
                 font=(FF, 10)).pack(side="left", anchor="s", pady=(0, 5))
        tk.Frame(self, bg=GOLD, height=2).pack(fill="x")

    # ── body ──
    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=16)

        # source controls
        src_card = tk.Frame(body, bg=CARD, highlightbackground="#E2E2DC",
                            highlightthickness=1)
        src_card.pack(fill="x")

        tk.Label(src_card, text="SOURCE", bg=CARD, fg=SUB,
                 font=(FF, 7, "bold")).pack(anchor="w", padx=16, pady=(12, 0))
        tk.Label(src_card,
                 text="Add individual files or a whole folder. Every supported file becomes a PDF.",
                 bg=CARD, fg=TEXT, font=(FF, 9)).pack(anchor="w", padx=16, pady=(2, 8))

        btn_row = tk.Frame(src_card, bg=CARD)
        btn_row.pack(fill="x", padx=16, pady=(0, 8))
        tk.Button(btn_row, text="Add files", command=self._add_files,
                  bg=NAVY, fg="white", relief="flat", padx=14, pady=6,
                  font=(FF, 9, "bold"), cursor="hand2").pack(side="left")
        tk.Button(btn_row, text="Add folder", command=self._add_folder,
                  bg=NAVY, fg="white", relief="flat", padx=14, pady=6,
                  font=(FF, 9, "bold"), cursor="hand2").pack(side="left", padx=(8, 0))
        tk.Button(btn_row, text="Clear", command=self._clear,
                  bg="#E2E2DC", fg=TEXT, relief="flat", padx=14, pady=6,
                  font=(FF, 9), cursor="hand2").pack(side="left", padx=(8, 0))

        self._recurse = tk.BooleanVar(value=True)
        tk.Checkbutton(btn_row, text="Include subfolders", variable=self._recurse,
                       bg=CARD, fg=TEXT, activebackground=CARD, selectcolor=CARD,
                       font=(FF, 9)).pack(side="left", padx=(16, 0))

        self._count_var = tk.StringVar(value="No files selected")
        tk.Label(src_card, textvariable=self._count_var, bg=CARD, fg=SUB,
                 font=(FF, 9)).pack(anchor="w", padx=16, pady=(0, 12))

        # output controls
        out_row = tk.Frame(body, bg=BG)
        out_row.pack(fill="x", pady=(12, 0))
        tk.Button(out_row, text="Output folder…", command=self._pick_out,
                  bg="#E2E2DC", fg=TEXT, relief="flat", padx=12, pady=5,
                  font=(FF, 9), cursor="hand2").pack(side="left")
        self._out_var = tk.StringVar(value="PDFs saved next to each source file")
        tk.Label(out_row, textvariable=self._out_var, bg=BG, fg=SUB,
                 font=(FF, 9)).pack(side="left", padx=(12, 0))

        # action bar
        act = tk.Frame(body, bg=BG)
        act.pack(fill="x", pady=(14, 10))
        self._run_btn = tk.Button(act, text="Convert to PDF", command=self._run,
                                  bg=NAVY, fg="white", relief="flat", padx=22, pady=10,
                                  font=(FF, 10, "bold"), cursor="hand2")
        self._run_btn.pack(side="left")
        self._bar = ttk.Progressbar(act, mode="determinate", length=200)
        self._status = tk.StringVar(value="")
        tk.Label(act, textvariable=self._status, bg=BG, fg=SUB,
                 font=(FF, 9)).pack(side="left", padx=14)

        # log
        log_card = tk.Frame(body, bg=CARD, highlightbackground="#E2E2DC",
                           highlightthickness=1)
        log_card.pack(fill="both", expand=True)
        tk.Label(log_card, text="OUTPUT LOG", bg=CARD, fg=SUB,
                 font=(FF, 7, "bold")).pack(anchor="w", padx=16, pady=(12, 4))
        self._log = scrolledtext.ScrolledText(
            log_card, height=12, font=("Consolas", 9),
            bg=LOGB, fg=LOGF, relief="flat", state="disabled",
            insertbackground=LOGF, borderwidth=0)
        self._log.pack(fill="both", expand=True, padx=16, pady=(0, 14))

    # ── selection handlers ──
    def _add_files(self):
        paths = filedialog.askopenfilenames(title="Select files to convert")
        if paths:
            self._merge(paths)

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Select a folder to convert")
        if folder:
            self._merge(collect_from_folder(folder, self._recurse.get()))

    def _merge(self, paths):
        seen = set(self.items)
        for p in paths:
            ap = os.path.abspath(p)
            if os.path.isfile(ap) and ap not in seen:
                self.items.append(ap)
                seen.add(ap)
        self._update_count()

    def _clear(self):
        self.items = []
        self._update_count()

    def _update_count(self):
        n = len(self.items)
        supported = sum(1 for f in self.items
                        if os.path.splitext(f)[1].lower() in ALL_KNOWN)
        self._count_var.set(
            f"{n} file(s) selected  ·  {supported} directly supported, "
            f"{n - supported} will be tried as text or skipped")

    def _pick_out(self):
        folder = filedialog.askdirectory(title="Choose a single output folder")
        if folder:
            self.out_dir = folder
            self._out_var.set(f"Output folder: {folder}")
        else:
            self.out_dir = None
            self._out_var.set("PDFs saved next to each source file")

    # ── run ──
    def _log_write(self, msg):
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n")
        self._log.see("end")
        self._log.configure(state="disabled")

    def _run(self):
        if not self.items:
            messagebox.showwarning("No files", "Add files or a folder first.")
            return
        self._run_btn.configure(state="disabled")
        self._bar.pack(side="left", padx=(14, 0))
        self._bar.configure(maximum=len(self.items), value=0)
        self._status.set("Converting…")
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

        threading.Thread(target=self._execute, daemon=True).start()

    def _execute(self):
        def log(msg):
            self.after(0, lambda m=msg: self._log_write(m))

        def progress(i, total):
            self.after(0, lambda: self._bar.configure(value=i))

        files = list(self.items)
        try:
            done, skipped, failed = run_batch(files, self.out_dir, log, progress)
            log("")
            log(f"  Finished: {done} converted, {skipped} skipped, {len(failed)} failed.")
            summary = f"{done} converted, {skipped} skipped, {len(failed)} failed."
            self.after(0, lambda: self._status.set(summary))
            if failed:
                self.after(0, lambda: messagebox.showwarning(
                    "Finished with errors", summary + "\n\nSee the log for details."))
            else:
                self.after(0, lambda: messagebox.showinfo("Done", summary))
        except Exception as e:
            log(f"  ERROR  {e}")
            self.after(0, lambda: self._status.set("Error — see log"))
        finally:
            self.after(0, lambda: self._run_btn.configure(state="normal"))
            self.after(0, lambda: self._bar.pack_forget())


# ── ENTRY POINT ───────────────────────────────────────────────────────────────

def main():
    # Console mode if paths are passed on the command line.
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        files = []
        for a in args:
            if os.path.isdir(a):
                files += collect_from_folder(a, recurse=True)
            elif os.path.isfile(a):
                files.append(os.path.abspath(a))
        if not files:
            print("No files found.")
            return
        print(f"Converting {len(files)} file(s)...\n")
        done, skipped, failed = run_batch(files, None, print)
        print(f"\nDone. {done} converted, {skipped} skipped, {len(failed)} failed.")
        return

    App().mainloop()


if __name__ == "__main__":
    main()
