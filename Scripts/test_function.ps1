

function main () {
    Write-Output "begin";

    $PreConvRecordCountContent = Get-Content -Path "F:\Steadfast\Migration\PreConversionRecordCount.txt"
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
    $intRow = 1

    Do {
        
        $ColumnID = $UserWorksheet.Cells.Item($intRow, 1).Value()
        Write-Output "-------------------$ColumnID-----------------"
        # If ($ColumnID -gt 0) {
        #     ForEach ($record in $recordSet) {
        #         #for each address found, concatinate them into one variable
        #         Write-Output $record
        #     }
        # }
        UpdateFields "Entities" "*Entities*entities*" $UserWorksheet
        UpdateFields "Profiles" "*Profiles*profile*" $UserWorksheet
        UpdateFields "Contacts" "*Contacts*personnel*" $UserWorksheet
        UpdateFields "Addresseses" "*Addresses*addresses*" $UserWorksheet
        UpdateFields "Authorised Reps" "*Authorised Reps*entities*" $UserWorksheet
        UpdateFields "All Insurers" "*All Insurers*entities*" $UserWorksheet
        UpdateFields "SVU Insurers" "*SVU Insurers*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "SVU Products" "*SVU Products*" $UserWorksheet
        UpdateFields "Sunrise Insurers" "*Sunrise Insurers*sunrise_insurers*" $UserWorksheet
        UpdateFields "Sunrise Products" "*Sunrise Products*sunrise_products*" $UserWorksheet
        UpdateFields "Client Tasks" "*Client Tasks*tasks*" $UserWorksheet
        UpdateFields "Client Taks Documents" "*Client Taks Documents*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "Policy Tasks" "*Policy Tasks*journals*" $UserWorksheet
        UpdateFields "Policy Taks Documents" "*Policy Taks Documents*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "Claims" "*Claims*claims*" $UserWorksheet
        UpdateFields "Claim Tasks" "*Claim Tasks*claims*" $UserWorksheet
        UpdateFields "Claim Taks Documents" "*Claim Taks Documents*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "Policies" "*Policies*policies*" $UserWorksheet
        UpdateFields "Invoices" "*Invoices*policies*" $UserWorksheet
        UpdateFields "Sunrise Policies" "*Sunrise Policies*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "SVU Policies" "*SVU Policies*N/A (table does not exist)*" $UserWorksheet
        UpdateFields "SVU Policy Opportunities" "*SVU Policy Opportunities*SVUPolicies*" $UserWorksheet

        $intRow++
    } While ($ColumnID)

    # Loop for post conversion table

    # $intRow = 27
    # Do {
        
    #     $ColumnID = $UserWorksheet.Cells.Item($intRow, 1).Value()
    #     Write-Output $ColumnID
 
    #     UpdateFields "Entities" "*Entities*entities*" UserWorksheet
    #     UpdateFields "Profiles" "*Profiles*profile*" UserWorksheet
    #     UpdateFields "Contacts" "*Contacts*personnel*" UserWorksheet
    #     UpdateFields "Addresseses" "*Addresseses*addresses*" UserWorksheet
    #     UpdateFields "Authorised Reps" "*AuthorisedReps*entities*" UserWorksheet
    #     UpdateFields "All Insurers" "*AllInsurers*entities*" UserWorksheet
    #     UpdateFields "SVU Insurers" "*SVUInsurers*entities*" UserWorksheet
    #     UpdateFields "SVU Products" "*SVUProducts*CustomWorkbooks*" UserWorksheet
    #     UpdateFields "Sunrise Insurers" "*SunriseInsurers*sunrise_insurers*" UserWorksheet
    #     UpdateFields "Sunrise Products" "*SunriseProducts*sunrise_products*" UserWorksheet
    #     UpdateFields "Client Tasks" "*ClientTasks*tasks*" UserWorksheet
    #     UpdateFields "Client Taks Documents" "*ClientTaksDocuments*tasks&tasks_sub_tasks*" UserWorksheet
    #     UpdateFields "Policy Tasks" "*PolicyTasks*journals*" UserWorksheet
    #     UpdateFields "Policy Taks Documents" "*PolicyTaksDocuments*journals&journal_sub_tasks*" UserWorksheet
    #     UpdateFields "Claims" "*Claims*claims*" UserWorksheet
    #     UpdateFields "Claim Tasks" "*ClaimTasks*claims*" UserWorksheet
    #     UpdateFields "Claim Taks Documents" "*ClaimTaksDocuments*clatas_id&clasubta_id*" UserWorksheet
    #     UpdateFields "Policies" "*Policies*pol_id*" UserWorksheet
    #     UpdateFields "Invoices" "*Invoices*tran_id*" UserWorksheet
    #     UpdateFields "Sunrise Policies" "*SunrisePolicies*sunwor_id*" UserWorksheet
    #     UpdateFields "SVU Policies" "*SVUPolicies*wor_id*" UserWorksheet
    #     UpdateFields "SVU Policy Opportunities" "*SVUPolicyOpportunities*WorkbookHeader*" UserWorksheet

     
    #     $intRow++
    # } While ($ColumnID)


    Write-Output "end"
    Write-Output  $intRow
}
function UpdateFields ([String]$excelPattern, [String] $textPattern, $worksheet) {
    If ($ColumnID -contains $excelPattern) {
        # $distinct_count = $worksheet.Cells.Item($intRow, 4).Value()
        # $min_value = $worksheet.Cells.Item($intRow, 5).Value()
        # $max_value = $worksheet.Cells.Item($intRow, 6).Value()
        # Write-Output $distinct_count
        # Write-Output $max_value
        # Write-Output $min_value
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
    }
}
main  