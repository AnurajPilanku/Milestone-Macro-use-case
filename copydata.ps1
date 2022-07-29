##powershell script to copy data from xlsx sheet and paste in an existing xlsm sheet.
##File Paths##
$DumpFile_Path = '\\acprd01\E\3M_CAC\SMO_AMA\Milestone\query_output\milestone_dump.csv' # source's fullpath
$MasterFile_path = '\\acprd01\E\3M_CAC\SMO_AMA\Milestone\xlsm_file\Milestone.xlsm' # destination's fullpath



try
{
##Initializing Com object##
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.displayAlerts = $false



##Opening Excel File##
$Dump_Open = $Excel.Workbooks.open($DumpFile_Path)
$MasterFile_Open = $Excel.Workbooks.open($MasterFile_path)


##Reading source sheet##
$DumpSheet = $Dump_Open.WorkSheets.item("milestone_dump")
$DumpSheet.Activate()
$range = $DumpSheet.Range("A:z")
$range.Copy() | out-null


##Pasting Data to destination sheet##
$MasterCopySheet = $MasterFile_Open.WorkSheets.item("Copy")
$MasterCopySheet.Activate()
$range2 = $MasterCopySheet.Range("A:Z")
$MasterCopySheet.Paste($range2)

##Saving and closing the excel files##
$MasterFile_Open.Save()
$Dump_Open.close()
$MasterFile_Open.close()
return "File copied successfully"#printing

}
catch
{
$ErrorMessage = $_.Exception.Message
return $_.Exception.Message
}
finally
{
$Excel.quit()

}




















