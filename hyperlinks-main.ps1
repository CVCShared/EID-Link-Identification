using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

using module "C:\Users\Nash Ferguson\Desktop\EID-Link-Identification\link-search.psm1"
$Date = Get-Date -Format s
$CsvName = $Date.tostring() -replace ":","-"
$Dir = "C:\Users\Nash Ferguson\Desktop\Small Docx"
IdentifyLinks($Dir)
