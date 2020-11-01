using module "C:\Users\Nash Ferguson\Downloads\Communary.FileExtensions-master\Communary.FileExtensions-master\Communary.FileExtensions.psm1"

function IdentifyLinks($Dir, $CsvName){
    $global:CsvName = $CsvName
    $Dir = $Dir
    
    CheckLinksDocx($Dir)
    #CheckLinksXlsx($Dir)
    #CheckLinksPpt($Dir)
}

# Docx, Xlsx, and Ppt functions all have same basic function: they search for all
# of the files that are their type, clear away the hidden files, check which files
# are locked, and then search for links in the files that remain (inacessible files are logged).
# They then add to a hash table of documents and their links, which is then exported to a csv. 

function CheckLinksDocx($Dir){
    <#
    .SYNOPSIS
        A function to extract the links in all doc/x files in a given dir
    .DESCRIPTION
        This function does a FastFind for all doc/x files in the dir, 
        opens each file in a Word COM object, gathers all links (in text and shapes),
        and sends them to a CSV
    .PARAMETER Dir
        Directory to search for doc/x files in
    #>

    $DocxLinks = @{}
    $DocxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.doc?" -Hidden -AttributeFilterMode Exclude
    
    if($null -eq $DocxFiles){
        #Record dirs with no doc/x
        $EmptyDir = [PSCustomObject]@{
            'Directory' = $Dir
            'Doc type that doesnt exist' = "Doc/x"    
        }
        $EmptyDir|Export-Csv -Path "$Dir\$CsvName empty directories.csv" -NoClobber -Append -NoTypeInformation
        continue
    }
    else{
    $DocxFiles = [System.Collections.ArrayList] $DocxFiles

    #Check which files are locked, and keep unlocked files
    $DocxFiles = CheckLocks $DocxFiles $Dir
    $Word = New-Object -ComObject word.application
    $Word.visible = $false

    write-host("Docx Files to read are : $DocxFiles")

    #Look for links in unlocked files
    foreach($File in $DocxFiles){

        $FilePath = $File
        write-host("File currently being read: ", $FilePath)
        $Document = $Word.Documents.Open($FilePath)
        $Hyperlinks = $Document.Hyperlinks
        $Shapes = $Document.inlineshapes
        $DocxLinks[$FilePath] = [System.Collections.ArrayList]@()
        #Check for linked shapes (charts, data tables, stuff like that)
        foreach($Shape in $Shapes){
            $Shape = @{"Shape " = $Shape.linkformat.sourcefullname}
            [void]$DocxLinks[$FilePath].Add($Shape)
        }
        #Check for linked text
        foreach ($Hyperlink in $Hyperlinks){
            $Link = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
            [void]$DocxLinks[$FilePath].Add($Link)
        } 

        $Document.Close()
    }
    $Word.Quit()
    ExportToCsv($DocxLinks)
}
}

function CheckLinksXlsx($Dir){
    <#
    .SYNOPSIS
        A function to extract the links in all xls/x files in a given dir
    .DESCRIPTION
        This function does a FastFind for all xls/x files in the dir, 
        opens each file in an Excel COM object, gathers all links (in text and charts),
        and sends them to a CSV
    .PARAMETER Dir
        Directory to search for xls/x files in
    #>

    $XlsxLinks = @{}
    $XlsxFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.xls?" -Hidden -AttributeFilterMode Exclude

    
    if($null -eq $XlsxFiles){
        #Record dirs with no xls/x
        $EmptyDir = [PSCustomObject]@{
            'Directory' = $Dir
            'Doc type that doesnt exist' = "Xls/x"    
        }
        $EmptyDir|Export-Csv -Path "$Dir\$CsvName empty directories.csv" -NoClobber -Append -NoTypeInformation
        continue
    }

    else{
    $XlsxFiles = [System.Collections.ArrayList]$XlsxFiles
    #Check which files are locked, and keep unlocked files
    $XlsxFiles = CheckLocks $XlsxFiles $Dir
    $excel = New-Object -ComObject excel.application
    $excel.visible = $false
    
    Write-Host("XLSX Files to read are : ", $XlsxFiles)
    
    #Look for links in unlocked files
    foreach($File in $XlsxFiles){
            $FilePath = $File
            $Workbook = $excel.Workbooks.Open($FilePath)
            $WorksheetNum = 0
            $XlsxLinks[$FilePath] = [System.Collections.ArrayList]@()

            foreach($Worksheet in $Workbook.Worksheets){
                $WorksheetNum++
                $Hyperlinks = $Workbook.Worksheets($WorksheetNum).Hyperlinks
                $Charts = $Worksheet.chartobjects()
                
                write-host("File currently being read : ", $FilePath)

                #Add hyperlinks to hash table with links and their text
                if($Hyperlinks.count -gt 0){
                    foreach ($Hyperlink in $Hyperlinks){
                        $Link = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                        [void]$XlsxLinks[$FilePath].Add($Link)
                    }
                }

                try{
                # Try to get a list of all charts, catch if there are none
                $Charts = $Worksheet.chartobjects()

                foreach ($Chart in $Charts){
                    # Get cell formulas, which contain links to source data
                    $Formulas = $Chart.chart.seriescollection()|select-object Formula
                    $FormulasSeen = [System.Collections.ArrayList]@()
                    
                    # Foreach formula, use a regex to take out the link, ignoring duplicate links from the same chart
                    foreach ($Formula in $Formulas){
                        $Link = $Formula.psobject.properties.value
                        $Regex = [regex]::new("'(.+?)'")
                        $Link = $Regex.matches($Link)

                        if ($Link.count -eq 0){
                            write-host("No linked charts found in $FilePath")
                            continue
                        }
                        if ($FormulasSeen.contains($Link[1].tostring())){
                            write-host("Seen link (",$Link[1],") already")
                            continue
                        }
                        else{
                            $Chart = @{"Chart" = $Link[1].tostring()} 
                            [void]$XlsxLinks[$FilePath].Add($Chart)
                            $FormulasSeen.add($Link[1].tostring())
                        }
                    }
                }
            }
            

                catch{
                    "No charts"
                }
        }
            
        $Workbook.Close()
    }
    $Excel.Quit()
    ExportToCsv($XlsxLinks)
}

     # !!!!! IMPORTANT EXCEL NOTE !!!!!!# 
     # When making anything to change the address, note that excel
     # uses "../"^n instead of a full path. Something will have to be built
     # to take a certain portion of the $dir based on whatever "../" is there
     
}

function CheckLinksPpt($Dir){
    <#
    .SYNOPSIS
        A function to extract the links in all ppt/x files in a given dir
    .DESCRIPTION
        This function does a FastFind for all ppt/x files in the dir, 
        opens each file in a Powerpoint COM object, gathers all links (in text and shapes),
        and sends them to a CSV
    .PARAMETER Dir
        Directory to search for ppt/x files in
    #>

    $PptLinks = @{}
    $PptFiles = Invoke-FastFind -Recurse -Path $Dir -Filter "*.ppt?" -Hidden -AttributeFilterMode Exclude

    if($null -eq $PptFiles){
        #Record dirs with no ppt/x
        $EmptyDir = [PSCustomObject]@{
            'Directory' = $Dir
            'Doc type that doesnt exist' = "Ppt/x"    
        }
        $EmptyDir|Export-Csv -Path "$Dir\$CsvName empty directories.csv" -NoClobber -Append -NoTypeInformation
        continue
    }
    else{
    #Check which files are locked, and keep unlocked files
    $PptFiles = CheckLocks $PptFiles $Dir
    $PowerPt = New-Object -ComObject powerpoint.application

    write-host("PPT Files to read are : ", $PptFiles)

    #$PowerPt.visible = $false
    #Look for links in unlocked files
    foreach($File in $PptFiles){

        # NEW CODE, has shape link finding
            $FilePath = $File
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
                        $LinkedShape = @{"Shape" = $Shape.linkformat.sourcefullname}
                        [void]$PptLinks[$FilePath].Add($LinkedShape)
                    }
                }
                
                #Check for linked text
                foreach($Hyperlink in $Hyperlinks){
                    $Link = @{$Hyperlink.TextToDisplay = $Hyperlink.Address} 
                    [void]$PptLinks[$FilePath].Add($Link)
                }
            
            }
            $Ppt.Close()
    }
    $PowerPt.Quit()
    ExportToCsv($PptLinks)
}
}

function CheckLocks($Files, $Dir){
    $CsvName = $Global:CsvName
    write-host("Checking Locks on dir $Dir")
    $OperableFiles = [System.Collections.ArrayList]@()
    foreach($File in $Files){
        #Try to open file in IO stream, catch errors and record
        try
        {   
            #$File = $File.Name
            $FilePath = $File.path
            $Test = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
            $Test.Close() #This line will unlock it (Supposedly)
            $Test.Dispose()
            [void]$OperableFiles.Add($File.path)

            write-host($FilePath, " is unlocked")

        }
        
        catch 
        {
            write-host($FilePath, " is locked")
            $LockedFile = [PSCustomObject]@{
                'Document Name' = $File
                'Error' = $_.Exception.Message    
            }
            $LockedFile|Export-Csv -Path "$Dir\$CsvName error-report.csv" -NoClobber -Append -NoTypeInformation
            
        }
    }

    return $OperableFiles
}

function ExportToCsv($LinkList){
    $CsvName = $Global:CsvName
    #Foreach link in the list, append it to the global CSV
    $LinkList.GetEnumerator() | ForEach-Object {
        $FileName = $_.key
        $Links = $_.value
        
        foreach ($Link in $Links){
            $Link.GetEnumerator()|ForEach-Object{
                $LinkInfo = [PSCustomObject]@{
                    'Document Name' = $FileName
                    'Text' = $_.key
                    'Target' = $_.value
                }
                $LinkInfo|Export-Csv -Path "$Dir\$CsvName" -NoClobber -Append -NoTypeInformation
        }
    }
}
}

Export-ModuleMember -Function IdentifyLinks