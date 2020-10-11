using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"
function IdentifyLinks($Dir){
    #XLSX and PPT needs to be in separate function, containing it in IdentifyLinks caused an error
    Docx($Dir)
    Xlsx($Dir)
    Ppt($Dir)
}

# Docx, Xlsx, and Ppt functions all have same basic function: they search for all
# of the files that are their type, clear away the hidden files, check which files
# are locked, and then search for links in the files that remain (inacessible files are logged).
# They then add to a hash table of documents and their links, which is then exported to a csv. 

function Docx($Dir){
    #Input: Directory
    #Purpose: Find all word files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in doc/x files
    
    $DocxLinks = @{}
    $DocxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.doc?" -Hidden -AttributeFilterMode Exclude
    $DocxFiles = [System.Collections.ArrayList] $DocxFiles
    
    #Remove hidden files
    <# [System.Collections.ArrayList] $RemoveFiles = @()
    foreach ($File in $DocxFiles){
        if(($File.Attributes|Out-String) -like "*Hidden*"){
            $RemoveFiles.add($File)
        }
    }

    foreach ($File in $RemoveFiles){
        $DocxFiles.Remove($File)
    } #>

    #Check which files are locked, and keep unlocked files
    $DocxFiles = CheckLocks $DocxFiles $Dir
    $Word = New-Object -ComObject word.application
    $Word.visible = $false

    #Look for links in unlocked files
    foreach($File in $DocxFiles){

        # NEW CODE
        $FilePath = $Dir + "\" + $File
        write-host("File ", $FilePath)
        $Document = $Word.Documents.Open($FilePath)
        $Hyperlinks = $Document.Hyperlinks
        $Shapes = $Document.inlineshapes
        $DocxLinks[$FilePath] = [System.Collections.ArrayList]@()


        foreach($Shape in $Shapes){
            $Value = @{"Shape " = $Shape.linkformat.sourcefullname}
            [void]$DocxLinks[$FilePath].Add($Value)
        }
        
        foreach ($Hyperlink in $Hyperlinks){
            $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
            [void]$DocxLinks[$FilePath].Add($Value)
        } 

        $Document.Close()
        

        # OLD CODE
        <# $FilePath = $Dir + "\" + $File
        write-host("File ",$FilePath)
        
        $Document = $Word.Documents.Open($FilePath)
        $Hyperlinks = $Document.Hyperlinks
        
        #Add document to hash table with links and their text
        if($Hyperlinks.count -gt 0)
        {
            $DocxLinks[$FilePath] = [System.Collections.ArrayList]@()
            foreach ($Hyperlink in $Hyperlinks){
                $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 

                [void]$DocxLinks[$FilePath].Add($Value)
            } 
        }
        $Document.Close() #>
    }
    $Word.Quit()
    ExportToCsv($DocxLinks)
}

function Xlsx($Dir){
    #Input: Directory
    #Purpose: Find all excel files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in xls/x files

    $XlsxLinks = @{}
    $XlsxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.xls?" -Hidden -AttributeFilterMode Exclude
    $XlsxFiles = [System.Collections.ArrayList]$XlsxFiles
    #Remove hidden files
    <# [System.Collections.ArrayList]$RemoveFiles = @()
    foreach ($File in $XlsxFiles){
        if(($File.Attributes|Out-String) -like "*Hidden*"){
            $RemoveFiles.add($File)
        }
    }

    foreach ($File in $RemoveFiles){
        $XlsxFiles.Remove($File)
    } #>

    #Check which files are locked, and keep unlocked files
    $XlsxFiles = CheckLocks $XlsxFiles $Dir
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    Write-Host("XLSX Files : ", $XlsxFiles)
    #Look for links in unlocked files
    foreach($File in $XlsxFiles){
            $FilePath = $Dir + "\" + $File
            $workbook = $excel.Workbooks.Open($FilePath)
            $WorksheetNum = 0

            foreach($Worksheet in $workbook.Worksheets){
                $WorksheetNum++
                $Hyperlinks = $workbook.Worksheets($WorksheetNum).Hyperlinks
                write-host("File ", $FilePath)

                #Add document to hash table with links and their text
                if($Hyperlinks.count -gt 0){
                    $XlsxLinks[$FilePath] = [System.Collections.ArrayList]@()

                    foreach ($Hyperlink in $Hyperlinks){
                        $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                        [void]$XlsxLinks[$FilePath].Add($Value)
                    }

            }
        }
        $workbook.Close()
        
    }
    $Excel.Quit()
    ExportToCsv($XlsxLinks)
}

function Ppt($Dir){
    #Input: Directory
    #Purpose: Find all powerpoint files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in ppt/x files

    $PptLinks = @{}
    $PptFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.ppt?" -Hidden -AttributeFilterMode Exclude
    
    #Remove hidden files
    <# [System.Collections.ArrayList] $RemoveFiles = @()
    foreach ($File in $PptFiles){
        if(($File.Attributes|Out-String) -like "*Hidden*"){
            $RemoveFiles.add($File)
        }
    }

    foreach ($File in $RemoveFiles){
        $PptFiles.Remove($File)
    } #>

    #Check which files are locked, and keep unlocked files
    $PptFiles = CheckLocks $PptFiles $Dir
    $PowerPt = New-Object -ComObject powerpoint.application
    write-host("PPT Files : ", $PptFiles)
    #$PowerPt.visible = $false
    #Look for links in unlocked files
    foreach($File in $PptFiles){

        # NEW CODE, has shape link finding
            $FilePath = $Dir + "\" + $File
            write-host("File ", $FilePath)
            $Ppt = $PowerPt.Presentations.Open($FilePath, [Microsoft.Office.Core.MsoTriState]::msoFalse,[Microsoft.Office.Core.MsoTriState]::msoFalse,[Microsoft.Office.Core.MsoTriState]::msoFalse)
            $Slides = $Ppt.Slides

            Foreach ($Slide in $Slides){
                $Shapes = $Slide.shapes
                $Hyperlinks = $Slide.Hyperlinks
                $PptLinks[$FilePath] = [System.Collections.ArrayList]@()
                foreach($Shape in $Shapes){
                    $Value = @{"Shape" = $Shape.linkformat.sourcefullname}
                    [void]$PptLinks[$FilePath].Add($Value)
                }

                foreach($Hyperlink in $Hyperlinks){
                    $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                    [void]$PptLinks[$FilePath].Add($Value)
                }



            # OLD CODE, uncomment if necessary

                <# $Hyperlinks = $Slide.Hyperlinks
                
                #Add document to hash table with links and their text
                if($Hyperlinks.count -gt 0){
                    $PptLinks[$FilePath] = [System.Collections.ArrayList]@()

                    foreach ($Hyperlink in $Hyperlinks){
                        $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                        [void]$PptLinks[$FilePath].Add($Value)
                    } 

                } #>
            
            
            }
            $Ppt.Close()
    }
    $PowerPt.Quit()
    ExportToCsv($PptLinks)
}

function CheckLocks($Files, $Dir){
    
    write-host("Checking Locks on dir $Dir")
    $OperableFiles = [System.Collections.ArrayList]@()
    foreach($File in $Files){
        try
        {   
            $File = $File.Name
            $FilePath = "$Dir\$File"
            $Test = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
            $Test.Close() #This line will unlock it (Supposedly)
            $Test.Dispose()
            [void]$OperableFiles.Add($File)
        }
        
        catch 
        {
            $obj = [PSCustomObject]@{
                'Document Name' = $File
                'Error' = $_.Exception.Message    
            }
            $obj|Export-Csv -Path "$Dir\error-report.csv" -NoClobber -Append -NoTypeInformation
            
        }
    }

    return $OperableFiles
}

function ExportToCsv($LinkList){
    $LinkList.GetEnumerator() | ForEach-Object {
        $FileName = $_.key
        $Links = $_.value
    
        foreach ($Link in $Links){
            $Link.GetEnumerator()|ForEach-Object{
                $obj = [PSCustomObject]@{
                    'Document Name' = $FileName
                    'Text' = $_.key
                    'Target' = $_.value
                }
                $obj|Export-Csv -Path "$Dir\sunday-testing.csv" -NoClobber -Append -NoTypeInformation
        }
    }
}
}

Export-ModuleMember -Function IdentifyLinks