#Anuraj Pilanku
from win32com.client import Dispatch
import sys
import openpyxl
xlsm = sys.argv[1]
sql=sys.argv[2]

#wb=openpyxl.load_workbook(wkbk1)
#ws=wb.active
#rowcount=ws.max_row
#wb.close()

try:
    excel = Dispatch("Excel.Application")
    excel.Visible = 1
    source = excel.Workbooks.Open(xlsm)
    #sourcesheet=source.Worksheets(1)
    excel.Range("A1:M800").Select()
    excel.Selection.Copy()
    #excel2=Dispatch("Excel.Application")
    destination = excel.Workbooks.Open(sql)
    #destinationsheet = destination.Worksheets(1)
    excel.Range("A1:M800").Select()
    excel.Selection.PasteSpecial(Paste=-4163)


    destination.Save()
    destination.Close()
    source.Save()  # SaveAs(Filename:=sys.argv[1])
    source.Close()
    excel.Quit()

    print("success")
except:
    import os
    import win32com.shell.shell as shell
    import time

    # commands='taskkill /f /im EXCEL.EXE'
    # shell.ShellExecuteEx(lpVerb='runas',lpFile='cmd.exe',lpParameters='/c'+commands)
    os.system('taskkill /f /im EXCEL.EXE')
    time.sleep(10)
    

    excel = Dispatch("Excel.Application")
    excel.Visible = 1
    source = excel.Workbooks.Open(xlsm)
    # sourcesheet=source.Worksheets(1)
    excel.Range("A1:M800").Select()
    excel.Selection.Copy()
    # excel2=Dispatch("Excel.Application")
    destination = excel.Workbooks.Open(sql)
    # destinationsheet = destination.Worksheets(1)
    excel.Range("A1:M800").Select()
    excel.Selection.PasteSpecial(Paste=-4163)

    destination.Save()
    destination.Close()
    source.Save()  # SaveAs(Filename:=sys.argv[1])
    source.Close()
    excel.Quit()

    print("success")




#python ipi2k.py C:\Users\2040664\anuraj\SMO\sacin.xlsx C:\Users\2040664\anuraj\SMO\sd.xlsx