import tkinter as tk
import xlrd
import openpyxl as px
from tkinter import filedialog

select_files = tk.Tk()
select_files.withdraw()

full_path = filedialog.askopenfilenames()

d = {}

wb = px.Workbook()
for f in full_path:
    wb = xlrd.open_workbook(f)
    sheet = wb.sheet_by_index(0)
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) in "部材機種コード":
                for rowx in range(row + 1, sheet.nrows):
                    for colx in range(col, 6):
                        d.setdefault(sheet.cell_value(rowx, col + 1), [])[rowx - (row + 1)]\
                            .append(sheet.cell_value(rowx, colx))

