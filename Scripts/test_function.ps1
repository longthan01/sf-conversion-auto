

function main () {
    Write-Output "begin";

    $PreConvRecordCountContent = Get-Content -Path "F:\Steadfast\Migration\PreConversionRecordCount.txt"
    $PostConvRecordCountContent = Get-Content -Path "F:\Steadfast\Migration\PostConversionRecordCount.txt"
    $start_table = 0
    $end_table = 20
    $regex = "[-][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-][-]"


    # Open record count spreadsheet and write value
    $objExcel = New-Object -com Excel.Application
    $objExcel.Visible = $True
    $targetFilePath = 'F:\Steadfast\Migration\Conversion Record Counts.xlsx'
    $UserWorkBook = $objExcel.Workbooks.Open($targetFilePath)
    $UserWorksheet = $UserWorkBook.Worksheets.Item(1)


    #Loop pre conversion table
    Write-Output "/******************** Insert to pre conversion table ***************/"
    $intRow = 1

    Do {
        
        $ColumnID = $UserWorksheet.Cells.Item($intRow, 1).Value()
        Write-Output "-------------------$ColumnID-----------------"
        UpdateFields "Entities" "*Entities*entities*" $UserWorksheet "pre"
        UpdateFields "Profiles" "*Profiles*profile*" $UserWorksheet "pre"
        UpdateFields "Contacts" "*Contacts*personnel*" $UserWorksheet "pre"
        UpdateFields "Addresseses" "*Addresses*addresses*" $UserWorksheet "pre"
        UpdateFields "Authorised Reps" "*Authorised Reps*entities*" $UserWorksheet "pre"
        UpdateFields "All Insurers" "*All Insurers*entities*" $UserWorksheet "pre"
        UpdateFields "SVU Insurers" "*SVU Insurers*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "SVU Products" "*SVU Products*" $UserWorksheet "pre"
        UpdateFields "Sunrise Insurers" "*Sunrise Insurers*sunrise_insurers*" $UserWorksheet "pre"
        UpdateFields "Sunrise Products" "*Sunrise Products*sunrise_products*" $UserWorksheet "pre"
        UpdateFields "Client Tasks" "*Client Tasks*tasks*" $UserWorksheet "pre"
        UpdateFields "Client Taks Documents" "*Client Taks Documents*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "Policy Tasks" "*Policy Tasks*journals*" $UserWorksheet "pre"
        UpdateFields "Policy Taks Documents" "*Policy Taks Documents*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "Claims" "*Claims*claims*" $UserWorksheet "pre"
        UpdateFields "Claim Tasks" "*Claim Tasks*claims*" $UserWorksheet "pre"
        UpdateFields "Claim Taks Documents" "*Claim Taks Documents*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "Policies" "*Policies*policies*" $UserWorksheet "pre"
        UpdateFields "Invoices" "*Invoices*policies*" $UserWorksheet "pre"
        UpdateFields "Sunrise Policies" "*Sunrise Policies*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "SVU Policies" "*SVU Policies*N/A (table does not exist)*" $UserWorksheet "pre"
        UpdateFields "SVU Policy Opportunities" "*SVU Policy Opportunities*SVUPolicies*" $UserWorksheet "pre"

        $intRow++
    } While ($ColumnID)

    # Loop for post conversion table
    Write-Output "/******************** Insert to post conversion table ***************/"
    $intRow = 27
    Do {
        
        $ColumnID = $UserWorksheet.Cells.Item($intRow, 1).Value()
        Write-Output "-------------------$ColumnID-----------------"
 
        UpdateFields "Entities" "*Entities*entities*" $UserWorksheet "post"
        UpdateFields "Profiles" "*Profiles*profile*" $UserWorksheet "post"
        UpdateFields "Contacts" "*Contacts*personnel*" $UserWorksheet "post"
        UpdateFields "Addresseses" "*Addresses*addresses*" $UserWorksheet "post"
        UpdateFields "Authorised Reps" "*Authorised Reps*entities*" $UserWorksheet "post"
        UpdateFields "All Insurers" "*All Insurers*entities*" $UserWorksheet "post"
        UpdateFields "SVU Insurers" "*SVU Insurers*svu_insurers*" $UserWorksheet "post"
        UpdateFields "SVU Products" "*SVU Products*svu_products*" $UserWorksheet "post"
        UpdateFields "Sunrise Insurers" "*Sunrise Insurers*sunrise_insurers*" $UserWorksheet "post"
        UpdateFields "Sunrise Products" "*Sunrise Products*sunrise_products*" $UserWorksheet "post"
        UpdateFields "Client Tasks" "*Client Tasks*tasks*" $UserWorksheet "post"
        UpdateFields "Client Taks Documents" "*Client Taks Documents*tasks_sub_tasks*" $UserWorksheet "post"
        UpdateFields "Policy Tasks" "*Policy Tasks*journals*" $UserWorksheet "post"
        UpdateFields "Policy Taks Documents" "*Policy Taks Documents*journal_sub_tasks*" $UserWorksheet "post"
        UpdateFields "Claims" "*Claims*claims*" $UserWorksheet "post"
        UpdateFields "Claim Tasks" "*Claim Tasks*claims*" $UserWorksheet "post"
        UpdateFields "Claim Taks Documents" "*Claim Taks Documents*tasks_sub_tasks*" $UserWorksheet "post"
        UpdateFields "Policies" "*Policies*policies*" $UserWorksheet "post"
        UpdateFields "Invoices" "*Invoices*policies*" $UserWorksheet "post"
        UpdateFields "Sunrise Policies" "*Sunrise Policies*sunrise_policies*" $UserWorksheet "post"
        UpdateFields "SVU Policies" "*SVU Policies*SVUPolicies*" $UserWorksheet "post"
        UpdateFields "SVU Policy Opportunities" "*SVU Policy Opportunities*SVUPolicies*" $UserWorksheet "post"

     
        $intRow++
    } While ($ColumnID)


    Write-Output "end"
    Write-Output  $intRow
}
function UpdateFields ([String]$excelPattern, [String] $textPattern, $worksheet, [String] $type = "pre") {
    If ($ColumnID -contains $excelPattern) {
        # $distinct_count = $worksheet.Cells.Item($intRow, 4).Value()
        # $min_value = $worksheet.Cells.Item($intRow, 5).Value()
        # $max_value = $worksheet.Cells.Item($intRow, 6).Value()
        # Write-Output $distinct_count
        # Write-Output $max_value
        # Write-Output $min_value
        If($type -eq "pre"){
            foreach ($line in $PreConvRecordCountContent) {
                if ($line -like $textPattern) {
                    Write-Output "line $count $line"
                    $splitUp = $line.substring(300) -split "\s+"
                    if ($splitUp) {
                        # distinct count
                        Write-Output "distinct count" $splitUp[1]
                        $worksheet.Cells.Item($intRow, 4) = $splitUp[1]
                        # min value
                        Write-Output "min value" $splitUp[1]
                        $worksheet.Cells.Item($intRow, 5) = $splitUp[2]
                        # max value
                        Write-Output "max value" $splitUp[1]
                        $worksheet.Cells.Item($intRow, 6) = $splitUp[3]
                    } else {
                        Write-Output "No line match"
                    }
                }
               
                $count++
            }
        } else{
            foreach ($line in $PostConvRecordCountContent) {
                if ($line -like $textPattern) {
                    Write-Output "line $count $line"
                    $splitUp = $line.substring(300) -split "\s+"
                    if ($splitUp) {
                        # distinct count
                        Write-Output "distinct count" $splitUp[1]
                        $worksheet.Cells.Item($intRow, 4) = $splitUp[1]
                        
                        if ($line -like "*N/A (different table used)*") {
                            Write-Output "min max value N/A (different table used) " 
                            $worksheet.Cells.Item($intRow, 5)  = "N/A (different table used)"
                            $worksheet.Cells.Item($intRow, 6) = "N/A (different table used)"
                        } elseif ($line -like "*N/A (different logic used)*") {
                            Write-Output "min max value N/A (different logic used) " 
                            $worksheet.Cells.Item($intRow, 5)  = "N/A (different logic used)"
                            $worksheet.Cells.Item($intRow, 6) = "N/A (different logic used)"
                        } else {
                            # min value
                            Write-Output "min value" $splitUp[2]
                            $worksheet.Cells.Item($intRow, 5) = $splitUp[2]
                            # max value
                            Write-Output "max value" $splitUp[3]
                            $worksheet.Cells.Item($intRow, 6) = $splitUp[3]
                        }
                        
                    } else {
                        Write-Output "No line match"
                    }
                }
               
                $count++
            }
        }
        
    }
}
main  