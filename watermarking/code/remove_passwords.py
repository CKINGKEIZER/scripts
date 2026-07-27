"""
remove_passwords.py
-------------------
Engine module: decrypt owner/user-password-protected PDFs, writing clean
(unencrypted) copies to an output folder.

Used by:
  * launcher.py "Remove Passwords" tab -> remove_passwords(files, password, output_folder, ...)
  * command line                       -> python remove_passwords.py
        (reads the password from removepass.xlsx cell A1 and decrypts every
         PDF in the removepass/ subfolder — the original stand-alone flow)

Public function:
  remove_passwords(files, password, output_folder, log, progress) -> (ok, total)
"""

import os
import sys

# ── CONFIG ────────────────────────────────────────────────────────────────────
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR    = os.path.dirname(SCRIPT_DIR)
OUTPUT_FOLDER = os.path.join(PARENT_DIR, "output")
# ─────────────────────────────────────────────────────────────────────────────


def remove_passwords(files, password, output_folder=OUTPUT_FOLDER,
                     log=None, progress=None):
    """
    Decrypt a list of PDF file paths using `password` and write clean copies
    to output_folder.

    Returns (ok, total).
    """
    from pypdf import PdfReader, PdfWriter

    log = log or print
    os.makedirs(output_folder, exist_ok=True)
    total = len(files)
    ok = 0

    for i, path in enumerate(files, 1):
        fname = os.path.basename(path)
        try:
            reader = PdfReader(path)
            if reader.is_encrypted and reader.decrypt(password) == 0:
                log(f"  ✗  {fname}  —  wrong password")
                if progress:
                    progress(i, total)
                continue
            writer = PdfWriter()
            writer.clone_reader_document_root(reader)
            with open(os.path.join(output_folder, fname), "wb") as f:
                writer.write(f)
            log(f"  ✓  {fname}")
            ok += 1
        except Exception as e:
            log(f"  ✗  {fname}  —  {e}")

        if progress:
            progress(i, total)

    log("")
    log(f"  Finished: {ok}/{total} files written to output/")
    return ok, total


# ── CLI: original Excel-driven batch flow ─────────────────────────────────────

def _cli():
    """
    Stand-alone flow: read the owner password from removepass.xlsx (cell A1)
    and decrypt every PDF in the removepass/ subfolder.
    """
    import openpyxl

    excel_path = os.path.join(SCRIPT_DIR, "removepass.xlsx")
    input_folder = os.path.join(SCRIPT_DIR, "removepass")

    if not os.path.isfile(excel_path):
        print(f"ERROR: {excel_path} not found.")
        sys.exit(1)
    if not os.path.isdir(input_folder):
        print(f"ERROR: input folder not found: {input_folder}")
        sys.exit(1)

    wb = openpyxl.load_workbook(excel_path)
    password = wb.active["A1"].value
    wb.close()
    print("Password loaded from Excel.")

    files = [os.path.join(input_folder, f)
             for f in os.listdir(input_folder) if f.lower().endswith(".pdf")]
    print(f"Found {len(files)} PDFs to process.")
    remove_passwords(files, password, OUTPUT_FOLDER)
    print(f"\nAll files written to '{OUTPUT_FOLDER}/'.")


if __name__ == "__main__":
    _cli()
