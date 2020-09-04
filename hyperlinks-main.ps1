using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"
#using module "C:\Users\Nash Ferguson\Desktop\El Dorado\docx-search.psm1"
#using module "C:\Users\Nash Ferguson\Desktop\El Dorado\xlsx-search.psm1"
using module "C:\Users\Nash Ferguson\Desktop\El Dorado\link-search.psm1"
$Date = Get-Date -Format s
$CsvName = $Date.tostring() -replace ":","-"
$Dir = "C:\Users\Nash Ferguson\Desktop\Target Dir"

IdentifyLinks($Dir)
#IdentifyDocxLinks($Dir)
#IdentifyXlsxLinks($Dir)