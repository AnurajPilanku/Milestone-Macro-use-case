import win32com.client
import os, os.path
from pywinauto import Application
import threading
import time
import sys

input=sys.argv[1]


def run_macro():
    try:
        doc = ""
        xl1 = ""
        path =input
        print(path)
        #path = input+"\\\\Milestone.xlsm"
        xl1=win32com.client.Dispatch("Excel.Application")
        
        xl1.DisplayAlerts = False
        doc = xl1.Workbooks.Open(path, ReadOnly=1)
        
        xl1.Application.Run("Extract")
        doc.Save()
        doc.Close()
        xl1.Quit()
        print(path)
    except Exception as e:
        doc.Save()
        doc.Close()
        xl1.Quit()

def close_pop_up_window():
    main_app = Application(backend="uia").connect(title_re="Microsoft Visual Basic", control_type="Window")
    main_app_win1 = main_app.window(title_re="Microsoft Visual Basic", control_type="Window")
    End_button = main_app_win1.child_window(title="End", control_type="Button", visible_only=True, found_index=0)
    End_button.invoke()

t1 = threading.Thread(target=run_macro)
#t2 = threading.Thread(target=close_pop_up_window)

t1.start()
time.sleep(20)
#t2.start()

print("success")