"""
add_passwords.py
----------------
Engine module: apply AES-256 owner-password encryption (edit/copy protection)
to PDF files. The files still open WITHOUT a password; the owner password only
blocks editing and copying.

Used by:
  * launcher.py "Add Passwords" tab  -> protect_files(...)
  * command line                     -> python add_passwords.py --password PW [--folder DIR]

Public functions:
  protect_files(files, owner_password, output_folder, log, progress) -> (ok, total)
  protect_pdfs(input_folder, owner_password, output_folder, ...)      -> ok   (folder wrapper)
"""

import os
import sys

# ── CONFIG ────────────────────────────────────────────────────────────────────
SCRIPT_DIR    = os.path.dirname(os.path.abspath(__file__))
PARENT_DIR    = os.path.dirname(SCRIPT_DIR)
INPUT_FOLDER  = os.path.join(SCRIPT_DIR, "to_protect")
OUTPUT_FOLDER = os.path.join(PARENT_DIR, "output")
# ─────────────────────────────────────────────────────────────────────────────


def check_dependencies():
    missing = []
    for pkg in ("pypdf",):
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"Missing packages: {', '.join(missing)}")
        print(f"Run:  pip install {' '.join(missing)}")
        sys.exit(1)


def _permissions():
    """Owner-password permissions: allow printing + text/graphics extraction."""
    from pypdf.constants import UserAccessPermissions
    return (
        UserAccessPermissions.PRINT
        | UserAccessPermissions.PRINT_TO_REPRESENTATION
        | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
    )


def collect_pdfs(folder):
    """Return sorted list of PDF filenames in folder. Exit if none found."""
    if not os.path.isdir(folder):
        print(f"ERROR: Input folder not found: {folder}")
        sys.exit(1)
    pdfs = sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
    if not pdfs:
        print(f"ERROR: No PDF files found in {folder}")
        sys.exit(1)
    return pdfs


def protect_files(files, owner_password, output_folder=OUTPUT_FOLDER,
                  log=None, progress=None):
    """
    Owner-password-encrypt a list of PDF file paths.

    Parameters
    ----------
    files          : list[str]  – absolute paths to source PDFs
    owner_password : str        – owner password (blocks editing/copying)
    output_folder  : str        – where protected copies are written
    log            : callable(str), optional
    progress       : callable(current, total), optional

    Returns (ok, total).
    """
    from pypdf import PdfReader, PdfWriter

    log = log or print
    perms = _permissions()

    os.makedirs(output_folder, exist_ok=True)
    total = len(files)
    ok = 0

    for i, path in enumerate(files, 1):
        fname = os.path.basename(path)
        try:
            reader = PdfReader(path)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(
                user_password="",
                owner_password=owner_password,
                algorithm="AES-256",
                permissions_flag=perms,
            )

            with open(os.path.join(output_folder, fname), "wb") as f:
                writer.write(f)

            log(f"  [{i}/{total}] ✓  {fname}")
            ok += 1
        except Exception as e:
            log(f"  [{i}/{total}] ✗  {fname}  —  {e}")

        if progress:
            progress(i, total)

    log("")
    log(f"  Finished: {ok}/{total} files written to output/")
    return ok, total


def protect_pdfs(input_folder, owner_password, output_folder=OUTPUT_FOLDER,
                 log_callback=None, progress_callback=None):
    """Folder wrapper around protect_files (kept for the CLI / older callers)."""
    log = log_callback if log_callback else print
    pdfs = sorted(f for f in os.listdir(input_folder) if f.lower().endswith(".pdf"))
    if not pdfs:
        log(f"ERROR: No PDF files found in {input_folder}")
        return 0
    files = [os.path.join(input_folder, f) for f in pdfs]
    ok, _ = protect_files(files, owner_password, output_folder, log, progress_callback)
    return ok


def main():
    check_dependencies()

    owner_password = None
    if "--password" in sys.argv:
        idx = sys.argv.index("--password")
        if idx + 1 < len(sys.argv):
            owner_password = sys.argv[idx + 1]

    if not owner_password:
        print("ERROR: No password provided.")
        print("Usage:  python add_passwords.py --password YOUR_PASSWORD [--folder DIR]")
        sys.exit(1)

    folder = INPUT_FOLDER
    if "--folder" in sys.argv:
        idx = sys.argv.index("--folder")
        if idx + 1 < len(sys.argv):
            folder = sys.argv[idx + 1]

    pdfs = collect_pdfs(folder)
    print(f"Input      : {folder}")
    print(f"Files      : {len(pdfs)}")
    print(f"Output     : {OUTPUT_FOLDER}")
    print()
    protect_pdfs(folder, owner_password)


if __name__ == "__main__":
    main()
