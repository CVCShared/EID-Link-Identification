using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

function IdentifyDocxLinks($Dir){
    $Files = Invoke-FastFind -Recurse -Path $Dir -Filter "*.docx"
    $Word = New-Object -ComObject word.application
    $Word.visible = $false
    foreach($File in $Files){

        if(($File.Attributes|Out-String) -like "*Hidden*"){continue}

        else{
        $FilePath = $Dir + "\" + $File.Name
        write-host("File ",$FilePath)
        $Document = $Word.Documents.Open($FilePath)
        $Hyperlinks = $Document.Hyperlinks 
        
        if($Hyperlinks.count -gt 0)
        {
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

Export-ModuleMember -Function IdentifyDocxLinks