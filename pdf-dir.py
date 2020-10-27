import glob


pdf_files = glob.glob("*.pdf")


for pdf in pdf_files:
    print(pdf)
