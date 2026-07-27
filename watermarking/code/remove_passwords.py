import os
from pypdf import PdfReader, PdfWriter
import openpyxl

BASE   = os.path.dirname(os.path.abspath(__file__))   # watermarking/code/
PARENT = os.path.dirname(BASE)                         # watermarking/

# removepass.xlsx is in the same folder as this script
EXCEL_PATH = os.path.join(BASE, "removepass.xlsx")
wb = openpyxl.load_workbook(EXCEL_PATH)
ws = wb.active
OWNER_PASSWORD = ws["A1"].value
print(f"Password loaded from Excel.")

INPUT_FOLDER  = os.path.join(BASE,   "removepass")
OUTPUT_FOLDER = os.path.join(PARENT, "output")

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(".pdf")]
print(f"Found {len(files)} PDFs to process.")

for filename in files:
    input_path = os.path.join(INPUT_FOLDER, filename)
    output_path = os.path.join(OUTPUT_FOLDER, filename)

    reader = PdfReader(input_path)

    if reader.is_encrypted:
        result = reader.decrypt(OWNER_PASSWORD)
        if result == 0:
            print(f"SKIP (wrong password): {filename}")
            continue

    writer = PdfWriter()
    writer.clone_reader_document_root(reader)

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"Done: {filename}")

print(f"\nAll files written to '{OUTPUT_FOLDER}/'.")
