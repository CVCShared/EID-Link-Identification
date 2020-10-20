using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

function IdentifyLinks($Dir, $CsvName){
    $global:CsvName = $CsvName
    $Dir = $Dir

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


    #Check which files are locked, and keep unlocked files
    $DocxFiles = CheckLocks $DocxFiles $Dir
    $Word = New-Object -ComObject word.application
    $Word.visible = $false
    write-host("Docx Files to read are : $DocxFiles")
    #Look for links in unlocked files
    foreach($File in $DocxFiles){

        # NEW CODE
        $FilePath = $Dir + "\" + $File
        write-host("File currently being read: ", $FilePath)
        $Document = $Word.Documents.Open($FilePath)
        $Hyperlinks = $Document.Hyperlinks
        $Shapes = $Document.inlineshapes
        $DocxLinks[$FilePath] = [System.Collections.ArrayList]@()
        #Check for linked shapes (charts, data tables, stuff like that)
        foreach($Shape in $Shapes){
            $Value = @{"Shape " = $Shape.linkformat.sourcefullname}
            [void]$DocxLinks[$FilePath].Add($Value)
        }
        #Check for linked text
        foreach ($Hyperlink in $Hyperlinks){
            $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
            [void]$DocxLinks[$FilePath].Add($Value)
        } 

        $Document.Close()
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

    #Check which files are locked, and keep unlocked files
    $XlsxFiles = CheckLocks $XlsxFiles $Dir
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    Write-Host("XLSX Files to read are : ", $XlsxFiles)
    #Look for links in unlocked files
    foreach($File in $XlsxFiles){
            $FilePath = $Dir + "\" + $File
            $workbook = $excel.Workbooks.Open($FilePath)
            $WorksheetNum = 0
            $XlsxLinks[$FilePath] = [System.Collections.ArrayList]@()

            foreach($Worksheet in $workbook.Worksheets){
                $WorksheetNum++
                $Hyperlinks = $workbook.Worksheets($WorksheetNum).Hyperlinks
                $Charts = $Worksheet.chartobjects()
                write-host("File currently being read : ", $FilePath)

                #Add hyperlinks to hash table with links and their text
                if($Hyperlinks.count -gt 0){
                    foreach ($Hyperlink in $Hyperlinks){
                        $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                        [void]$XlsxLinks[$FilePath].Add($Value)
                    }
                }

                try{
                # Try to get a list of all charts, catch if there are none
                $charts = $worksheet.chartobjects()
                foreach ($chart in $charts){
                    # Get cell formulas, which contain links to source data
                    $Formulas = $chart.chart.seriescollection()|select-object Formula
                    $FormulasSeen = [System.Collections.ArrayList]@()
                    
                    # Foreach formula, use a regex to take out the link, ignoring duplicate links from the 
                    # same chart
                    foreach ($Formula in $Formulas){
                        $link = $Formula.psobject.properties.value
                        $regex = [regex]::new("'(.+?)'")
                        $link = $regex.matches($link)

                        if ($link.count -eq 0){
                            write-host("No linked charts found in $FilePath")
                            continue
                        }
                        if ($FormulasSeen.contains($link[1].tostring())){
                            write-host("Seen link (",$link[1],") already")
                            continue
                        }
                        else{
                            $Value = @{"Chart" = $link[1].tostring()} 
                            [void]$XlsxLinks[$FilePath].Add($Value)
                            $FormulasSeen.add($link[1].tostring())
                        }
                    }
                }
            }
            

            catch{
                "No charts"
            }
        }
            
        $workbook.Close()
    }
    $Excel.Quit()
    ExportToCsv($XlsxLinks)


     # !!!!! IMPORTANT EXCEL NOTE !!!!!!# 
     # When making anything to change the address, note that excel is stupid
     # and uses "../"^n instead of a full path. Something will have to be built
     # to take a certain portion of the $dir based on whatever "../" is there
     
}

function Ppt($Dir){
    #Input: Directory
    #Purpose: Find all powerpoint files with hyperlinks, and make a record of them 
    #Output: Adds to CSV the file location, text, and destination of all hyperlinks in ppt/x files

    $PptLinks = @{}
    $PptFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.ppt?" -Hidden -AttributeFilterMode Exclude

    #Check which files are locked, and keep unlocked files
    $PptFiles = CheckLocks $PptFiles $Dir
    $PowerPt = New-Object -ComObject powerpoint.application
    write-host("PPT Files to read are : ", $PptFiles)
    #$PowerPt.visible = $false
    #Look for links in unlocked files
    foreach($File in $PptFiles){

        # NEW CODE, has shape link finding
            $FilePath = $Dir + "\" + $File
            write-host("File currently being read : ", $FilePath)
            $Ppt = $PowerPt.Presentations.Open($FilePath, [Microsoft.Office.Core.MsoTriState]::msoFalse,[Microsoft.Office.Core.MsoTriState]::msoFalse,[Microsoft.Office.Core.MsoTriState]::msoFalse)
            $Slides = $Ppt.Slides
            $PptLinks[$FilePath] = [System.Collections.ArrayList]@()

            Foreach ($Slide in $Slides){
                $Shapes = $Slide.shapes
                $Hyperlinks = $Slide.Hyperlinks
                
                #Check for linked shapes (charts, data tables, stuff like that)
                foreach($Shape in $Shapes){
                    if (-not ($null -eq $Shape.linkformat.sourcefullname )){
                        $Value = @{"Shape" = $Shape.linkformat.sourcefullname}
                        [void]$PptLinks[$FilePath].Add($Value)
                    }
                }
                
                #Check for linked text
                foreach($Hyperlink in $Hyperlinks){
                    $Value = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                    [void]$PptLinks[$FilePath].Add($Value)
                }
            
            }
            $Ppt.Close()
    }
    $PowerPt.Quit()
    ExportToCsv($PptLinks)
}

function CheckLocks($Files, $Dir){
    $CsvName = $Global:CsvName
    write-host("Checking Locks on dir $Dir")
    $OperableFiles = [System.Collections.ArrayList]@()
    foreach($File in $Files){
        #Try to open file in IO stream, catch errors and record
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
            $obj|Export-Csv -Path "$Dir\$CsvName error-report.csv" -NoClobber -Append -NoTypeInformation
            
        }
    }

    return $OperableFiles
}
function ExportToCsv($LinkList){
    $CsvName = $Global:CsvName
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
                $obj|Export-Csv -Path "$Dir\$CsvName" -NoClobber -Append -NoTypeInformation
        }
    }
}
}

Export-ModuleMember -Function IdentifyLinks