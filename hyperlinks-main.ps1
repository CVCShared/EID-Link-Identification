using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"
using module "C:\Users\Nash Ferguson\Desktop\EID-Link-Identification\link-search.psm1"

#Create names for csv's and the log
$Date = Get-Date -Format s
$CsvName = $Date.tostring() -replace ":","-"
$LogName = $CsvName + " log.txt"
$CsvName = $CsvName + " link data.csv"
$Dir = "C:\Users\Nash Ferguson\Desktop\Files to read"
#Read-Host("Please enter the directory to find links in: ")

Start-Transcript -Path "$PsScriptRoot\$LogName"
$CheckForLinks = IdentifyLinks $Dir $CsvName
$CheckForLinks
Stop-Transcript