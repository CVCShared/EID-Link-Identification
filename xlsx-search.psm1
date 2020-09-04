using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

function IdentifyXlsxLinks($Dir, $CsvName){
    $Files = Invoke-FastFind -Recurse -Path $Dir -Filter "*.xlsx"
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    
    foreach($File in $Files){
        
        $FilePath = $Dir + "\" + $File.name
        $workbook = $excel.Workbooks.Open($FilePath)
        if(($File.Attributes|Out-String) -like "*Hidden*"){continue}
        else{
            $WorksheetNum = 0
            foreach($Worksheet in $workbook.Worksheets){
                $WorksheetNum++
                $Hyperlinks = $workbook.Worksheets($WorksheetNum).Hyperlinks
                write-host($FilePath)
                if($Hyperlinks.count -gt 0){
                $obj = [PSCustomObject]@{
                    'Document Name' = $FilePath
                    'Text' = $null
                    'Target' = $null
                }

                foreach ($Hyperlink in $Hyperlinks){
                        $obj.'Text' = $Hyperlink|Select-Object -ExpandProperty TextToDisplay
                        $obj.'Target' = $Hyperlink|Select-Object -ExpandProperty Address
                        $obj|Export-Csv -Path "$Dir\test-csv60.csv" -NoClobber -Append -NoTypeInformation
                        
                }
            }
        }
        }
    }
    $excel.quit()
}

Export-ModuleMember -Function IdentifyXlsxLinks