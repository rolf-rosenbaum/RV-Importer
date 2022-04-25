#! python3
import openpyxl
import os
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

filePath = filedialog.askopenfilename()
index = filePath.rfind('/')
directory = filePath[:index]
wb = openpyxl.load_workbook(filePath)
sheet = wb.active

newSheet = wb.copy_worksheet(sheet)
newSheet.insert_cols(3, 1)

for cell in newSheet['C']:
    cell.value = 'St.'

newSheet['A1'].value = 'code'
newSheet['B1'].value = 'name'
newSheet['C1'].value = 'unit'
newSheet['D1'].value = 'price'

for cell in newSheet['B']:
    newVal = str(cell.value).replace(';', ',')
    cell.value = newVal

for cell in newSheet['D']:
    newVal = str(cell.value).replace('€', '').replace('.', ',')
    cell.value = newVal

wb.save(filePath)

lines = []
for row in newSheet.rows:
    line = ''
    for cell in row:
        line += str(cell.value) + ';'
    line += '\n'
    lines += [line] 

csvFile = open(os.path.join(directory, 'katalog.csv'), 'w', newline = '')
csvFile.writelines(lines)
csvFile.close()
