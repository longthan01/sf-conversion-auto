function main () {
    Write-Output "test1";

    $content = Get-Content -Path "F:\Steadfast\Migration\PreConversionRecordCount.txt"
    $start_table = 0
    $end_table = 20
    $regex = "[-][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-]"
    $count = 0;
    foreach($line in $content) {
        if($line -match $regex){
            $start_table = $count
        }
       
        $count++
    }
    #     $table= (gc "F:\Steadfast\Migration\PreConversionRecordCount.txt" -raw ) -replace "-{10,100}`r?`n" -replace "`r?`n *`r?`n"
# $table -split "`r?`n" | %{
#     If ($_ -notmatch 'Timed out'){
#         "`"{0}`"" -f ($_ -replace " {2,15}",'","')
#     }


# }|ConvertFrom-Csv | ConvertTo-html | Set-Content .\Sample.html -Encoding UTF8

Write-Output $count
}

main  