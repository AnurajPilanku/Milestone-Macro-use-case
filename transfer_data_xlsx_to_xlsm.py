import openpyxl
import sys
from datetime import datetime

head = 'portfolio,ad_request_application,AM/RLD,id,ad_request_title,Milestone,Planned_Start_Date,Planned_Completion_Date,planned_units,request_status,status'.split(",")
xlsm = sys.argv[1]
xl = sys.argv[2]
cnt = sys.argv[3]
try:
    xlsmWorkBook = openpyxl.load_workbook(filename=xlsm, read_only=False, keep_vba=True)
    # del xlsmWorkBook['Copy']
    xlsmWorkBook.create_sheet(title='Copy', index=0)
    xlsmSheet = xlsmWorkBook['Copy']
    xlsmSheet.append(head)

    with open(xl, 'r') as d:
        wholeData = d.read().split("\n")
    for i in range(1, len(wholeData)):
        singleList = wholeData[i].split(",")
        for i in range(0, len(singleList)):
            if i in [4, 5]:
                singleList[i] = "".join(list([val for val in singleList[i] if val.isalnum()]))
            elif i in [6, 7]:
                if "Date" not in str(singleList[i]):
                    singleList[i] = datetime.strptime(singleList[i], "%m/%d/%Y %H:%S")
            else:
                singleList[i] = singleList[i]
        xlsmSheet.append(singleList)
    d.close()
    xlsmWorkBook.save(xlsm)
    xlsmWorkBook.close()
    d.close()

    # python laia.py C:\Users\2040664\anuraj\ipi2k\Milestone.xlsm C:\Users\2040664\anuraj\ipi2k\milestone_dump.csv
    print(str(int(cnt) - int(1)))
except Exception as e:
    print(str(int(cnt) - int(1)))



