using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"
using module "C:\Users\Nash Ferguson\Desktop\EID-Link-Identification\link-search.psm1"

$Date = Get-Date -Format s
$CsvName = $Date.tostring() -replace ":","-"
$CsvName = $CsvName + " link data.csv"
$Dir = Read-Host("Please enter the directory to find links in: ")

IdentifyLinks $Dir $CsvName

# Need to add CSV input, and a replacement function. Should do as 2 psm1's