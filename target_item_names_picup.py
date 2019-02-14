import tkinter as tk
import xlrd
from tkinter import filedialog

select_files = tk.Tk()
select_files.withdraw()

full_path = filedialog.askopenfilenames()

for f in full_path:
    wb = xlrd.open_workbook(f)
    sheet = wb.sheet_by_index(1)
    print(sheet.cell(4, 3).value)
