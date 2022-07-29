#ANURAJ PILANKU
#EXTRACT  A SHEET FROM ONEWORKBOOK ANDPASTING IT IN ANOTHER WORKBOOK

import sys
import openpyxl as xl

path1 = sys.argv[1]
path2 = sys.argv[2]

wb1 = xl.load_workbook(filename=path1)
ws1 = wb1.worksheets[1]

wb2 = xl.Workbook()
ws2 = wb2.active

for row in ws1:
    for cell in row:
        ws2[cell.coordinate].value = cell.value

wb2.save(path2)
print("success")

#python musfira.py C:\Users\2040664\anuraj\SMO\excon.xlsx C:\Users\2040664\anuraj\SMO\exconw.xlsx