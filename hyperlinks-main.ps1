using module ".\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"
using module ".\link-search.psm1"

#Create names for csv's and the log
$Date = Get-Date -Format s
$CsvName = $Date.tostring() -replace ":","-"
$LogName = $CsvName + " log.txt"
$CsvName = $CsvName + " link data.csv"
#$Dir = "C:\Users\Nash Ferguson\Desktop\Files to read"
$Dir = Read-Host("Please enter the directory to find links in: ")

Start-Transcript -Path "$PsScriptRoot\$LogName"

try{
    $CheckForLinks = IdentifyLinks $Dir $CsvName
    $CheckForLinks
}
catch{
    write-host("ERROR")
    write-host($_.Exception.message)
}

Stop-Transcript