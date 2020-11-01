# read in a csv, and from that, change the links
# csv will contain two columns: old path and new path. There can be varying levels of specificity:
# first, make the specific swaps (determined by number of slashes), then work backwards. Make a queue
# of changes to make 
# if not given full path (like \old folder\ to \new folder\), do these last


# Read in CSV
function Read-Csv($CsvPath){
    $Csv = Import-Csv -Path $CsvPath
    $Csv|Get-Member
}

Read-Csv "C:\Users\Nash Ferguson\Desktop\link-replacement-csv.csv"