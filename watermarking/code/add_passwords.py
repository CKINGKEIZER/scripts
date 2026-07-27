"""
add_passwords.py
----------------
Reads every PDF in the 'to_protect/' subfolder, applies AES-256 owner-password
encryption (edit/copy protection), and writes the protected files to the output folder.

The files open without a password. The owner password blocks editing and copying.

Usage (standalone):
    python add_passwords.py --password YOUR_PASSWORD

Usage (from launcher):
    Called via protect_pdfs(input_folder, owner_password, log_callback)
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


def protect_pdfs(input_folder, owner_password, log_callback=None, progress_callback=None):
    """
    Apply owner-password protection to every PDF in input_folder.
    Writes output to OUTPUT_FOLDER.

    Parameters
    ----------
    input_folder       : str   – folder containing source PDFs
    owner_password     : str   – owner password for edit protection
    log_callback       : callable(str), optional – receives log lines
    progress_callback  : callable(current, total), optional – progress updates
    """
    from pypdf import PdfReader, PdfWriter
    from pypdf.constants import UserAccessPermissions

    log = log_callback if log_callback else print

    PERMS = (
        UserAccessPermissions.PRINT
        | UserAccessPermissions.PRINT_TO_REPRESENTATION
        | UserAccessPermissions.EXTRACT_TEXT_AND_GRAPHICS
    )

    pdfs = sorted(f for f in os.listdir(input_folder) if f.lower().endswith(".pdf"))
    if not pdfs:
        log(f"ERROR: No PDF files found in {input_folder}")
        return 0

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    total = len(pdfs)
    ok = 0

    for i, filename in enumerate(pdfs, 1):
        input_path  = os.path.join(input_folder, filename)
        output_path = os.path.join(OUTPUT_FOLDER, filename)

        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            writer.encrypt(
                user_password="",
                owner_password=owner_password,
                algorithm="AES-256",
                permissions_flag=PERMS,
            )

            with open(output_path, "wb") as f:
                writer.write(f)

            log(f"  [{i}/{total}] {filename}")
            ok += 1

        except Exception as e:
            log(f"  [{i}/{total}] {filename}  —  ERROR: {e}")

        if progress_callback:
            progress_callback(i, total)

    log(f"\nDone. {ok}/{total} file(s) protected in: {OUTPUT_FOLDER}")
    return ok


def main():
    check_dependencies()

    # Parse --password argument
    owner_password = None
    if "--password" in sys.argv:
        idx = sys.argv.index("--password")
        if idx + 1 < len(sys.argv):
            owner_password = sys.argv[idx + 1]

    if not owner_password:
        print("ERROR: No password provided.")
        print("Usage:  python add_passwords.py --password YOUR_PASSWORD")
        sys.exit(1)

    folder = INPUT_FOLDER
    # Allow optional --folder override
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
