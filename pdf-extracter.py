import os
import sys

import pdfplumber as ppmb

pdf_files = []
for _file in os.listdir(r"C:\pdf-extract-to-excel"):
    if _file.endswith(".pdf"):
        pdf_files.append(_file)

for _pdf in range(len(pdf_files)):
    with ppmb.open(pdf_files[_pdf]) as pdf:
        second = pdf.pages[1]
        third = pdf.pages[2]
        sys.stdout = open(f"мазмуны_{_pdf + 1}.txt", "w", encoding="utf-8")
        print(second.extract_text(), "\n", third.extract_text())
        sys.stdout.close()
