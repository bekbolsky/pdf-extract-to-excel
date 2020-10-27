import os
import re

import openpyxl as excel
# from openpyxl.worksheet.dimensions import ColumnDimension


txt_files = []
for _file in os.listdir(r"C:\pdf-extract-to-excel"):
    if _file.startswith("мазмуны") and _file.endswith(".txt"):
        txt_files.append(_file)


def text_parse(filename):
    fio = []
    with open(filename, encoding="utf-8") as f:
        # читать информацию в файле построчно
        for line in f.read().splitlines():
            # очистить строку от лишних точек и запятых
            new_line = re.sub("[,]|[.{1}][.\1+]", "", line)
            fio.append(new_line.strip())
        new_fio = []
        for i in range(0, len(fio), 2):
            x1 = re.sub("[0-9]$|[0-9][0-9]$", "", fio[i])
            x2 = re.sub("[0-9]$|[0-9][0-9]$", "", fio[i + 1])
            new_fio.append([x1.strip(), x2.strip()])
            i += 1
    return new_fio


def convert_txt_to_xlsx(list_txt_files):
    for _txt in range(len(list_txt_files)):
        massiv = text_parse(list_txt_files[_txt])
        wb = excel.Workbook()
        dest_filename = f"excel_baza_{_txt + 1}.xlsx"
        ws1 = wb.active
        ws1.title = "База"
        # col_a_dimension = ColumnDimension()
        for row in range(1, len(massiv) + 1):
            for col in range(1, len(massiv[0]) + 1):
                # print(massiv[row - 1][col - 1])
                _ = ws1.cell(
                    column=col,
                    row=row,
                    value="{}".format(massiv[row - 1][col - 1]),
                )
        wb.save(filename=dest_filename)


convert_txt_to_xlsx(txt_files)
