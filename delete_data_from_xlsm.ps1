##powershell script to delete data from a sheet or clear content from an xlsm sheet.
##File Paths##
#ANURAJ PILANKU

$MasterFile_path = '\\acprd01\E\3M_CAC\SMO_AMA\Milestone\xlsm_file\Milestone.xlsm' # destination's fullpath



try
{
##Initializing Com object##
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false
$Excel.displayAlerts = $false



##Opening Excel File##
$MasterFile_Open = $Excel.Workbooks.open($MasterFile_path)


##Pasting Data to destination sheet##
$MasterCopySheet = $Excel.WorkSheets.item("Copy")
$MasterCopySheet.Activate()
$rowMax=$MasterCopySheet.UsedRange.Rows.Count
$range=$MasterCopySheet.Range("d1","d18")
#$range.clear()
$list=@('A','B','C','D','E','F','G','H','I','J','K','L','M','N')
for ($i=0;$i -lt $list.Length ;$i++){$MasterCopySheet.Columns.item($list[$i]).clear()}
#$MasterCopySheet.Columns.item('C').clear()


#$MasterCopySheet.Rows($rowMax).EntireRow.Delete()
##Saving and closing the excel files##
$MasterFile_Open.Save()
$MasterFile_Open.close()



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