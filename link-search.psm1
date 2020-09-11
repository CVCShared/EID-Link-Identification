using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

function IdentifyLinks($Dir){
    #XLSX and PPT needs to be in separate function, containing it in IdentifyLinks caused an error
    write-host("doubng docx")
    Docx($Dir)
    Xlsx($Dir)
    Ppt($Dir)
}

function CheckLocks($Files, $Dir){
    write-host("Checking Locks")
    $OperableFiles = New-Object System.Collections.Generic.List[string]
    foreach($File in $Files){
        try
        {
            $FilePath = "$Dir\$File"
            $Test = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
            $Test.Close() #This line will unlock it (Supposedly)
            $Test.Dispose()
            $OperableFiles.Add[$File]
        }
        
        catch 
        {
            if($_.Exception.Message -Match "being used by another process"){
                $obj = [PSCustomObject]@{
                    'Document Name' = $File
                    'Error' = $_.Exception.Message  
                }
                $obj|Export-Csv -Path "$Dir\error-report.csv" -NoClobber -Append -NoTypeInformation
            }

            else{
                $obj = [PSCustomObject]@{
                    'Document Name' = $File
                    'Error' = $_.Exception.Message    
                }
                $obj|Export-Csv -Path "$Dir\error-report.csv" -NoClobber -Append -NoTypeInformation
            }
        }
    }
    return $OperableFiles
}
function Docx($Dir){
    #Input: Directory
    #Purpose: Find all word files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in doc/x files
    write-host("Doing files")
    $DocxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.doc?"
    Write-Host("Operable File output: ")
   
    $Array = CheckLocks($DocxFiles, $Dir)
    write-host("Array ", $Array)
    $Word = New-Object -ComObject word.application
    $Word.visible = $false

    foreach($File in $DocxFiles){

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
                    $obj|Export-Csv -Path "$Dir\test-csv600.csv" -NoClobber -Append -NoTypeInformation
                    
            }
        }
        
    }
    
    }

}

function Xlsx($Dir) {
    #Input: Directory
    #Purpose: Find all excel files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in xls/x files

    $XlsxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.xls?"
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false

    foreach($File in $XlsxFiles){
        
        if(($File.Attributes|Out-String) -like "*Hidden*"){continue} 
        
        else{
            $FilePath = $Dir + "\" + $File.name
            $workbook = $excel.Workbooks.Open($FilePath)
            $WorksheetNum = 0

            foreach($Worksheet in $workbook.Worksheets){
                $WorksheetNum++
                $Hyperlinks = $workbook.Worksheets($WorksheetNum).Hyperlinks
                write-host("File ", $FilePath)

                if($Hyperlinks.count -gt 0){
                $obj = [PSCustomObject]@{
                    'Document Name' = $FilePath
                    'Text' = $null
                    'Target' = $null
                }

                foreach ($Hyperlink in $Hyperlinks){
                        $obj.'Text' = $Hyperlink|Select-Object -ExpandProperty TextToDisplay
                        $obj.'Target' = $Hyperlink|Select-Object -ExpandProperty Address
                        $obj|Export-Csv -Path "$Dir\test-csv600.csv" -NoClobber -Append -NoTypeInformation
                        
                }
            }
        }
        }
    }
}

function Ppt($Dir){
    #Input: Directory
    #Purpose: Find all powerpoint files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in ppt/x files

    $PptFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.ppt?"

    $PowerPt = New-Object -ComObject powerpoint.application

    foreach($File in $PptFiles){
        if(($File.Attributes|Out-String) -like "*Hidden*"){continue} 

        else{
            $FilePath = $Dir + "\" + $File.name
            $Ppt = $PowerPt.Presentations.Open($FilePath)
            $Slides = $Ppt.Slides

            Foreach ($Slide in $Slides){
                $Hyperlinks = $Slide.Hyperlinks

                if($Hyperlinks.count -gt 0){
                    $obj = [PSCustomObject]@{
                        'Document Name' = $FilePath
                        'Text' = $null
                        'Target' = $null
                    }

                    foreach ($Hyperlink in $Hyperlinks){
                            $obj.'Text' = $Hyperlink|Select-Object -ExpandProperty TextToDisplay
                            $obj.'Target' = $Hyperlink|Select-Object -ExpandProperty Address
                            write-host($obj)
                            $obj|Export-Csv -Path "$Dir\test-csv600.csv" -NoClobber -Append -NoTypeInformation
                            
                    }
                }
            }
            }
    }
    
}

Export-ModuleMember -Function IdentifyLinks
#Get-Process | ?{$_.ProcessName -eq "WINWORD"} | Stop-Process