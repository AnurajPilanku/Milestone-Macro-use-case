# start Excel
$excel = New-Object -comobject Excel.Application
#open file
$FilePath = '\\acprd01\E\3M_CAC\SMO_AMA\Milestone\xlsm_file\Milestone_Enhanced- Nov 19th.xlsm'
$workbook = $excel.Workbooks.Open($FilePath)
##If you will like to check what is happend
$excel.Visible = $true
## Here you can "click" the button
$app = $excel.Application
$app.Run("Extract")
$workbook.save()
$workbook.close()
$excel.quit()