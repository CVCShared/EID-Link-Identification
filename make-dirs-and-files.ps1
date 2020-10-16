# Create sample directories and files to find links in 
param(
$MakeFiles,
$MakeDir,
$File,
$WriteDest,
$DirName,
$NumDuplicates,
$NumDirs,
$Int = 0
)

if ($MakeFiles){
    $FileName = (Split-Path $File -Leaf)

while ($Int -lt $NumDuplicates){
    $WriteName = "$WriteDest\$Int" + "-" + $FileName
    write-host("Dest is $WriteName")
    Copy-Item -Path $File -Destination "$WriteName"
    $Int += 1
}
}

if ($MakeDir){
    while ($Int -lt $NumDirs){
        New-Item -ItemType Directory -Path "$DirName-$Int"
        $Int += 1
    }
    
}
