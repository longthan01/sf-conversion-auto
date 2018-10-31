﻿Param(
    [string]$task = ''
)

# global variables #
$executionFolder = Split-Path $MyInvocation.MyCommand.Path
#path to msbuild, use for auto build
Set-Alias msbuild "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSbuild.exe"
Set-Alias auditAutomationTool ".\SF-ConversionAuto.exe"

#############################################################################################################################################################################################################
### FILL OUT THE FOLOWWING VARIABLES BEFORE RUN THE SCRIPT ###
#THE LEDGER BEING MIGRATE
$LEDGER = "YourInsurance"

$WINBEAT = "WINBEAT"
$IBAIS = "IBAIS"
$SOURCE_SYSTEM = $IBAIS

$conv_ledgerName = $LEDGER

#real path to ledger folder in conversion machine
$conv_ledgerFolder = "F:\Your Insurance Broker" 

$CONFIG_AZURE_BLOB_STORAGE_ACCOUNT = "FUCKINGTEST"
$CONFIG_AZURE_BLOB_STORAGE_KEY = "FUCKINGTEST"
$CONFIG_AUDIT_LISTING_FILE = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB_USER_SUFFIX = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB_PASSWORD = "FUCKINGTEST"
$conv_SVUListingFile = "AuditPolicyList.YOURINS_HQ.2018-05-25.csv"

#local working folder (in your development machine), if is current path, must be set to ".\"
$local_workingFolder = "d:\conversion_auto\$LEDGER"
 
#TEST PATH on local machine, only use for testing purpose at development time, comment out when run in conversion machine
#$conv_ledgerFolder = $local_workingFolder

#path to backup file, only use if the source system is winbeat
$conv_ledger_db_backup_path = "" #"$conv_ledgerFolder\Raw Data\Melbourne.bak"

#############################################################################################################################################################################################################

#path to your repo folders
$local_conversionRootPath = "D:\sfg-repos\insight_data_conversion\boa-data-conversion"
$local_sunriseAuditRootPath = "D:\sfg-repos\boa-sunrise-audit"
$local_sunriseExportRootPath = "D:\sfg-repos\boa-sunrise-export"
$local_svuAuditRootPath = "D:\sfg-repos\boa-svu-audit"

#conversion machine paths
$conv_appSourceFolder = $conv_ledgerFolder
$conv_automationReportsFolder = "$conv_ledgerFolder\automation_reports"
$conv_automationBackupFolder = $conv_ledgerFolder + "\automation_backups"
$conv_defaultDatabaseBackupFolder = "$conv_automationBackupFolder\database"
$conv_defaultSourceCodeBackupFolder = "$conv_automationBackupFolder\source code"
$conv_changeCollationSqlScriptPath = "$conv_ledgerFolder\Change_Collation.sql"
$conv_copyCustomConfigScriptPath = "$conv_ledgerFolder\Copy_Custom_Config.cmd"

$conv_ledger_db = $conv_ledgerName
$conv_ledger_insight_db = "$conv_ledgerName" + "Insight"

$conv_preUploadReportsFolder = "$conv_ledgerFolder\Run1\PreUploadReports\Results"
$conv_postConversionDataVerificationReportsForConsultantFolder = "$conv_ledgerFolder\Run1\PostConversionDataVerificationsReports\ForConsultant"
$conv_recordCountReportsFolder = "$conv_ledgerFolder\Run1\Record Counts"

$conv_sunriseAuditResultPath = "$conv_ledgerFolder\boa-sunrise-audit\Output"
$conv_svuAuditResultPath = "$conv_ledgerFolder\boa-svu-audit\Output"

#utility variables
$color_info = 'green'
$color_warning = 'yellow'
$color_error = 'red'

function printUsage() {
    wh "Params:"
    wh "-task"
    wh "(runsheet)rs-backupblobsfolder"
    wh
    wh "(conv)step0-pullcode"
    wh "(dev)step1-build"
    wh "(conv)step2-prepare"
    wh "(conv)step3-zip"
    wh "(conv)step4-extract"
    wh "(conv)step5-config"
    wh "(conv)step6-checkconfig"
    wh "(build)step7-applyconfig"
    wh "(conv)step8-recheckconfig"
    wh "(conv)step9-restoreledgerdb"
    wh "(conv)step10-changecollation"
    wh "(conv)step11-copycreateinsightdbscript"
    wh "(conv)step12-createinsightdb"
    wh "(conv)step13-runaudittools"
    wh "(conv)step14-preparereports"
}
function printVariable($name, $value) {
    wh "`t`t`$$name`: " "cyan" 0
    wh $value "magenta" 1
}
function printEnvironmentVariables() {
    wh
    wh "`t`t`t*** ======= ***" "cyan"
    wh
    printVariable "LEDGER" $LEDGER
    printVariable "SOURCE_SYSTEM" $SOURCE_SYSTEM
    printVariable "conv_ledgerFolder" $conv_ledgerFolder
    printVariable "conv_SVUListingFile" $conv_SVUListingFile
    printVariable "conv_ledger_db_backup_path" $conv_ledger_db_backup_path
    wh
    wh "`t`t`t*** ======= ***" "cyan"
    wh
}
function main() {
    $conv_ledgerFolder = replaceIfCurrentPath $conv_ledgerFolder
    $conv_appSourceFolder = replaceIfCurrentPath $conv_appSourceFolder
    $conv_automationBackupFolder = replaceIfCurrentPath $conv_automationBackupFolder
    $conv_automationReportsFolder = replaceIfCurrentPath $conv_automationReportsFolder
    $conv_defaultDatabaseBackupFolder = replaceIfCurrentPath $conv_defaultDatabaseBackupFolder
    $conv_changeCollationSqlScriptPath = replaceIfCurrentPath $conv_changeCollationSqlScriptPath
    $conv_copyCustomConfigScriptPath = replaceIfCurrentPath $conv_copyCustomConfigScriptPath
    $conv_ledger_db_backup_path = replaceIfCurrentPath $conv_ledger_db_backup_path

    printEnvironmentVariables

    if (($task -eq 'h') -Or ([string]::IsNullOrEmpty($task))) {
        printUsage
    }
    if ($task -eq 'step0-pullcode') {
        wh "Get latest codes from upstream master"
        pullLatestCode
    }
    if ($task -eq 'step1-build') {
        wh "Build conversion and related tools"
        buildSolutions
    }
    if ($task -eq 'step2-prepare') {
        wh "Prepare the conversion environment: create folder 'Raw data' and 'Admin' if they're not existed, backup folder 'Run1' to 'Run1_[current datetime]' if it's existed"
        prepareConvEnvironment
    }

    if ($task -eq 'step3-zip') {
        wh "Archive the conversion and related tools to $local_workingFolder"
        archive
    }
   
    if ($task -eq 'step4-extract') {
        wh "Extract the conversion and related tools to $conv_ledgerFolder"
        extract
    }
    
    if ($task -eq 'step5-config') {
        wh "Read the CONFIG_ variables, then replace it from the custom config file placeholders"
        config
    }
    
    if ($task -eq 'step6-checkconfig') {
        wh "Open custom config files to check whether your configurations are fucking right"
        checkConfig
    }

    if ($task -eq 'step7-applyconfig') {
        wh "Run the $conv_copyCustomConfigScriptPath to copy custom config to main config"
        applyConfig
    }
    if ($task -eq 'step8-recheckconfig') {
        wh "Open main config files and fucking re-check them by your fucking eyes"
        recheckConfig
    }
    if ($task -eq 'step9-restoreledgerdb') {
        wh "Restore $conv_ledger_db, if it's already existed, backup it to $conv_automationBackupFolder"
        restoreLedgerDb
    }
    if ($task -eq 'step10-changecollation') {
        wh "Check if the $conv_ledger_db has the right collation, if it's not, open the change collation script in ssms"
        wh "then you need to run it by your fucking hands"
        changeCollationLedgerDb
    }
    
    if ($task -eq 'step11-copycreateinsightdbscript') {
        wh "Copy Insight database creation script, this function basically copy out the path to script folder to clipboard"
        wh "You need to open build machine then paste it to windows explorer to open the folder then copy the latest script by your fucking hands"
        copyInsightCreationScript
    }

    if ($task -eq 'step12-createinsightdb') {
        wh "Create $($conv_ledgerName + "Insight") database, if the db is already existed, backup it to $conv_automationBackupFolder"
        createInsightDb
    }
    
    if ($task -eq "step13-runaudittools") {
        wh "Run SVU Audit tool, Sunrise export and Sunrise Audit tool"
        runAuditTools
    }

    if ($task -eq "step14-preparereports") {
        wh "Copy PreUpload, DataVerification, SVU Audit, Sunrise Audit and Record count reports to automation_reports folder"
        prepareReports
    }

    if ($task -eq 'rs-backupblobsfolder') {
        wh "Check whether blobs folder is existing, if it is, rename it to Source data from setup_[current date time]"
        backupBlobsFolder
    }
}

# COMMON FUNCTIONS #
function createFolderIfNotExists($path) {
    if (!(Test-Path -Path $path)) {
        wh "Folder $path does not existed, creating new one..."
        New-Item -ItemType Directory -Force -Path $path | Out-Null
    }
}
function replaceIfCurrentPath($path) {
    if ($path.StartsWith(".\")) {
        return $path.SubString(2)
    }
    return $path
}
function zipFileIntoFolder ($sourcePath, $destinationPath) {
   
    # check whether destination path is existed, if it is, do delete
    if (Test-Path -path $destinationPath) {
        # if root folder is existed, delete all of it's items
        Write-Host
        wh "'$destinationPath' folder is existing, do you FUCKING WANT TO DELETE? [y/n], default is [n]" $color_warning 1
        Write-Host
        $confirm = Read-Host
        if ($confirm -eq 'y') {
            Remove-Item -Path $destinationPath -Force -Recurse
        }
        else {
            wh "You choose to no delete"
            return
        }
    }

    wh "Zipping files in $sourcePath"
    #zip all files in source folder into destination folder
    Compress-Archive -path $sourcePath -DestinationPath $destinationPath
}
function extractFileIntoFolder($sourcePath, $destinationPath) {
    if (!(Test-Path -Path $sourcePath)) {
        wh "$sourcePath not found" $color_warning
        return 
    }
    if (Test-Path $destinationPath) {
        if ((Get-ChildItem $destinationPath | Measure-Object).count -ne 0) {
            Write-Host
            wh "'$destinationPath 'is not empty, do you FUCKING WANT TO DELETE? [y/n], default is [n]" $color_warning 0
            Write-Host
            $confirm = Read-Host
            if ($confirm -eq 'y') {
                Remove-Item -Path $destinationPath -Force -Recurse
            }
            else {
                wh "You choose to no delete, exist script now"
                return
            }
        }
    }
    
    wh "Extracting file $sourcePath into $destinationPath"
    Expand-Archive -LiteralPath $sourcePath -DestinationPath $destinationPath
}
function getSqlDefaultPath($type) {
    $query = "SELECT SERVERPROPERTY('InstanceDefault" + $type + "Path') as [Path]"
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
    $result = $rows[0]["Path"]
    return $result
}
function pullCodeFromUpstream($sourceDir) {
    Set-Location -Path $sourceDir
    wh "Pulling latest code into $sourceDir"
    git checkout master
    git pull upstream master
    Set-Location -Path $executionFolder
}
function pullLatestCode() {
    pullCodeFromUpstream $local_conversionRootPath
    pullCodeFromUpstream $local_sunriseAuditRootPath
    pullCodeFromUpstream $local_sunriseExportRootPath
    pullCodeFromUpstream $local_svuAuditRootPath
}
function restoreNugetPackages($solutionDir) {
    wh "Restoring nuget packages..."
    # Restore nuget package
    Get-ChildItem $solutionDir -Recurse -Include "packages.config" |
        ForEach-Object {   
        # Nuget restore packages used within all projects
        nuget restore $_.FullName -PackagesDirectory "$solutionDir\packages"
    }
    Write-Host "RESTORE NUGET PACKAGES DONE" -foreground green
}
function getBuildProjects($solutionRootDir, [string[]]$excludedProjects = @()) {
    $includeProjects = ""
    Get-ChildItem $solutionRootDir -Recurse -Include "*.csproj" |
        ForEach-Object {  
        $itemName = [System.IO.Path]::GetFileName($_.FullName)
        if (-not ($excludedProjects -contains $itemName)) {
            $includeProjects = $includeProjects + $_.BaseName.Replace(".", "_") + ";"
        }
    }
    return $includeProjects.Substring(0, $includeProjects.Length - 1)
}
function buildSolution($solutionPath, $buildMode, [string[]]$excludeProjects = @()) {	
    $compilingProjects = getBuildProjects $(Split-Path -Path $solutionPath) $excludeProjects
    wh "`tBuilding $compilingProjects of $solutionPath"
    if (!(Test-Path -Path $solutionPath)) {
        wh "Solution file not found" $color_error
        return
    }

    if ($buildMode -eq $()) {
        $buildMode = "Release"
    }

    wh "Cleaning $solutionPath"
    # Clean up
    msbuild "$solutionPath" /t:clean /p:configuration=$buildMode
		
    if ($LastExitCode -ne 0) {  
        Throw "Clean " + $solutionPath + " failed"  
    }

    # Build all
    msbuild "$solutionPath" /t:$compilingProjects /p:Configuration=$buildMode | Out-Default 
          
    if ($LastExitCode -ne 0) {
        Throw "Build " + $solutionPath + " failed"  
    }
    wh "Build $solutionPath successfully"
}
function buildSolutions() {
    buildSolution  "$local_sunriseAuditRootPath\boa-sunrise-audit.sln" "Debug"
    buildSolution  "$local_sunriseExportRootPath\SunriseExport.sln" "Debug"
    buildSolution  "$local_svuAuditRootPath\SvuAudit.sln" "Debug"
    buildSolution  "$local_conversionRootPath\DatabaseConversion.sln" "Debug" @("PolicyNotesConversion.csproj", "DocumentDownloader.csproj", "DocumentMultipac.csproj")
}
function getConfigValue($appconfigFilePath, $xpath, $attribute) {
    $appConfig = New-Object Xml
    $appConfig.Load($appConfigFilePath)
    foreach ($config in $appConfig.SelectNodes($xpath)) {
        if ($config.Attributes) {
            $att = $config.Attributes[$attribute]
            if ($att) {
                return $att.Value
            }
        }
    }
    return ""
}

#backup database 
function backupDb($dbName, $folder) {
    if (!(checkDbExist $dbName)) {
        wh "Database $dbName does not existed, skip backup"
    }
    else {
        createFolderIfNotExists $folder
        $backupName = $dbName + "_" + $(now) + ".bak"
        wh "Database $dbName is existing, backing it up into $($folder + '\' + $backupName) with query"
        $backupQuery = @"
    BACKUP DATABASE [$dbName] TO  DISK = N'$($folder + '\' + $backupName)' WITH 
    NOFORMAT, 
    INIT,  
    NAME = N'$dbName-Full Database Backup', 
    SKIP, 
    NOREWIND, 
    NOUNLOAD,  
    STATS = 10        
"@
        Write-Host $backupQuery
        Invoke-Sqlcmd  -ServerInstance '.' -Query $backupQuery -QueryTimeout 900
    }
}

# retrieve database logical names
function retrieveDatabaseLogicalNames($backupFile) {
    $sql = "restore filelistonly from disk = N'$backupFile' with file = 1"
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $sql)
    $dataLogicalName = ''
    $logLogicalName = ''
    if ($rows[0]["Type"] -eq "D") {
        $dataLogicalName = $rows[0]["LogicalName"]
        $logLogicalName = $rows[1]["LogicalName"]
    }
    else {
        $logLogicalName = $rows[0]["LogicalName"]
        $dataLogicalName = $rows[1]["LogicalName"]
    }
    [hashtable]$result = @{}
    $result.data = $dataLogicalName
    $result.log = $logLogicalName
    return $result
}

function wh($value = "", $color = $color_info, $newLine = 1) {
    if ([string]::IsNullOrEmpty($value)) {
        Write-Host
    }

    #add invocation function info
    $callStacks = @(Get-PSCallStack)
    $spaces = ""
    for ($i = 0; $i -lt $callStacks.Length; $i++) {
        $spaces = $spaces + " "
    }

    if ($color -eq $color_warning) {
        $value = " $spaces[!] " + $value
    }
    if ($color -eq $color_error) {
        $value = " $spaces[x] " + $value
    }
    if ($color -eq $color_info) {
        $value = " " + $value
    }
    if ($newLine -eq 1) {
        Write-Host $value -ForegroundColor $color
    }
    else {
        Write-Host $value -ForegroundColor $color -NoNewline
    }
}
function now() {
    $currentDateTime = Get-Date
    $str = "$($currentDateTime.Year)_$($currentDateTime.Month)_$($currentDateTime.Day)_$($currentDateTime.Hour)_$($currentDateTime.Minute)_$($currentDateTime.Second)"
    return $str
}
function checkDbExist($dbName) {
    $checkDbExistQuery = @"
    SELECT name 
    FROM master.dbo.sysdatabases 
    WHERE '[' + name + ']' ='$dbName' 
    OR name = '$dbName'
"@
    wh "Checking database existing with query"
    Write-Host "$checkDbExistQuery"
    $dbExistRows = @(Invoke-Sqlcmd  -ServerInstance '.' -Query $checkDbExistQuery -QueryTimeout 900)
    if ($dbExistRows.Count -eq 0) {
        wh "$dbName does not existed"
        return $false
    }
    wh "$dbName exsited"
    return $true
}

### RUN CONVERSION APP STEPS ###
# prepare environment like create all needing folder, copy audit files, backup Run1, backup old source codes, etc...
function prepareConvEnvironment() {
    # create all needing folders
    createFolderIfNotExists "$conv_ledgerFolder\Raw Data"
    createFolderIfNotExists "$conv_ledgerFolder\Admin"
    createFolderIfNotExists "$conv_defaultDatabaseBackupFolder"

    #backup old source code if this is the n run (n > 1)
    backupOldSourceCode
    #backup Run1 if this is the n run (n > 1)
    $run1 = "$conv_ledgerFolder\Run1"
    if (Test-Path -Path $run1) {
        $newName = "Run1_$(now)"
        Rename-Item -Path $run1 -NewName $newName
        wh "$run1 was renamed to $newName"
    }
}
function backupOldSourceCode() {
    #create current_date_time folder in $conv_defaultSourceCodeBackupFolder
    $backupFolder = "$conv_defaultSourceCodeBackupFolder\$(now)"
    createFolderIfNotExists $backupFolder
    wh "Backing up old source codes into $backupFolder"
    # backup console app
    backupFolder "$conv_ledgerFolder\DatabaseConversion.ConsoleApp" $backupFolder
    # backup import blob app
    backupFolder  "$conv_ledgerFolder\DatabaseConversion.AzureImportBlob" $backupFolder
    # backup data count checker app
    backupFolder  "$conv_ledgerFolder\DataCountChecker" $backupFolder
    # backup import schedule app
    backupFolder  "$conv_ledgerFolder\ImportedSchedule" $backupFolder
    # backup sunrise audit app
    backupFolder "$conv_ledgerFolder\boa-sunrise-audit" $backupFolder
    # backup sunrise export app
    backupFolder  "$conv_ledgerFolder\boa-sunrise-export" $backupFolder
    # backup svu audit app
    backupFolder "$conv_ledgerFolder\boa-svu-audit" $backupFolder
    wh
    wh "Do you fucking want to backup CONFIGURATION FOLDERS? [y/n], choose [y] if you want to overrite the old ones" $color_warning
    wh "or [n] if you want to re-use? default is [n]" $color_warning
    $backupConfigurationConfirm = Read-Host
    if($backupConfigurationConfirm -eq "y")
    {
        #backup old configuration folders
        $configurationFolders = @((Get-ChildItem -Path "$conv_ledgerFolder\*.Config").FullName)
        foreach ($f in $configurationFolders)
        {
            if(Test-Path -Path $f)
            {
                backupFolder "$f" $backupFolder
            }
        }
    }
}
function backupFolder($source, $dest) {
    if (!(Test-Path $source)) {
        wh "$source folder does not existed, skip backup" $color_warning
        return
    }
    Move-Item "$source" $dest
}
#archive conversion app and related tools, this step should be ran in development machine
function archive() {
    createFolderIfNotExists $local_workingFolder

    ### ZIP CONVERSION APP ###
    zipFileIntoFolder $local_conversionRootPath'\DatabaseConversion.ConsoleApp\bin\Debug\*' $local_workingFolder'\DatabaseConversion.ConsoleApp.zip'

    ### ZIP IMPORT BLOB APP ###
    zipFileIntoFolder $local_conversionRootPath'\DatabaseConversion.AzureImportBlob\bin\Debug\*' $local_workingFolder'\DatabaseConversion.AzureImportBlob.zip'

    ### ZIP DATA COUNT CHECKER APP ###
    zipFileIntoFolder $local_conversionRootPath'\DataCountChecker\bin\Debug\*' $local_workingFolder'\DataCountChecker.zip'

    ### ZIP IMPORT SCHEDULE APP ###
    zipFileIntoFolder $local_conversionRootPath'\ImportedSchedule\bin\Debug\*' $local_workingFolder'\ImportedSchedule.zip'

    ### ZIP AUDIT APPS ###
    zipFileIntoFolder $local_sunriseAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-audit.zip'
    zipFileIntoFolder $local_sunriseExportRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-export.zip'
    zipFileIntoFolder $local_svuAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-svu-audit.zip'
}

#extract conversion console app and related tools, this step should be ran in development machine
function extract() {
    # extract console app
    extractFileIntoFolder "$conv_appSourceFolder\DatabaseConversion.ConsoleApp.zip" "$conv_ledgerFolder\DatabaseConversion.ConsoleApp"
    # extract import blob app
    extractFileIntoFolder "$conv_appSourceFolder\DatabaseConversion.AzureImportBlob.zip" "$conv_ledgerFolder\DatabaseConversion.AzureImportBlob"
    # extract data count checker app
    extractFileIntoFolder "$conv_appSourceFolder\DataCountChecker.zip" "$conv_ledgerFolder\DataCountChecker"
    # extract import schedule app
    extractFileIntoFolder "$conv_appSourceFolder\ImportedSchedule.zip" "$conv_ledgerFolder\ImportedSchedule"
    # extract sunrise audit app
    extractFileIntoFolder "$conv_appSourceFolder\boa-sunrise-audit.zip" "$conv_ledgerFolder\boa-sunrise-audit"
    # extract sunrise export app
    extractFileIntoFolder "$conv_appSourceFolder\boa-sunrise-export.zip" "$conv_ledgerFolder\boa-sunrise-export"
    # extract svu audit app
    extractFileIntoFolder "$conv_appSourceFolder\boa-svu-audit.zip" "$conv_ledgerFolder\boa-svu-audit"
    #extract the config templates
    $templateFiles = "$conv_appSourceFolder\appconfig_template.zip"
    if (!(Test-Path -Path $templateFiles)) {
        wh "$templateFiles not found" $color_warning
        return 
    }
    $currentConfigFolders = (Get-ChildItem -Path "$conv_ledgerFolder\*.Config" | Measure-Object).Count
    if($currentConfigFolders -ne 0)
    {
        wh "Configuration folders are existed, check and extract by your fucking hands, the tool does not support extract programmatically."
        return
    }
    wh "Extracting file $templateFiles into $conv_ledgerFolder"
    Expand-Archive -LiteralPath "$conv_appSourceFolder\appconfig_template.zip" -DestinationPath $conv_ledgerFolder -Force
}

#config for each *.config file for a fucking bunch of projects
function setConfigs($appConfigFilePath, $configHashArray) {
    $appConfig = New-Object Xml
    $appConfig.Load($appConfigFilePath)
    $appSettings = $appConfig.GetElementsByTagName("appSettings")
    if ($appSettings) {
        foreach ($add in $appConfig.appSettings.add) {
            if (![string]::IsNullOrEmpty($add.value)) {
                foreach ($config in $configHashArray) {
                    $contains = $add.value.Contains($config.key)
                    $add.value = $add.value.Replace($config.key, $config.value)
                    if($contain)
                    {
                        Write-Host $add.value
                    }
                }
            }
        }
    }
   
    $connStrings = $appConfig.GetElementsByTagName("connectionStrings")
    if ($connStrings) {
        foreach ($add in $appConfig.connectionStrings.add) {
            if (![string]::IsNullOrEmpty($add.connectionString)) {
                foreach ($config in $configHashArray) {
                    $contains = $add.connectionString.Contains($config.key)
                    $add.connectionString = $add.connectionString.Replace($config.key, $config.value)
                    if($contain)
                    {
                        Write-Host $add.connectionString
                    }
                }
            }
        }
    }
    
    $appConfig.Save($appConfigFilePath)
}
function config() {
    wh "We will config app and connection string settings for a bunch of apps" $color_warning
    wh "if this is the n-th run (n > 1), this step shouldn't be performed, fucking confirm to continue (y/n)? default is [n]" $color_warning
    $confirm = Read-Host
    if ($confirm -eq "y") {
        $hashConfigurations = @(
            @{key = "[AZURE_BLOB_STORAGE_ACCOUNT]"; value = $CONFIG_AZURE_BLOB_STORAGE_ACCOUNT},
            @{key = "[AZURE_BLOB_STORAGE_KEY]"; value = $CONFIG_AZURE_BLOB_STORAGE_KEY},
            @{key = "[LEDGER_FOLDER]"; value = $conv_ledgerName},
            @{key = "[AUDIT_LISTING_FILE]"; value = $CONFIG_AUDIT_LISTING_FILE},
            @{key = "[CONV_LEDGER_INSIGHT_DB]"; value = $conv_ledger_insight_db},
            @{key = "[CONV_LEDGER_DB]"; value = $conv_ledger_db},
            @{key = "[AZURE_INSIGHT_DB]"; value = $CONFIG_AZURE_INSIGHT_DB},
            @{key = "[AZURE_INSIGHT_DB_USER_SUFFIX]"; value = $CONFIG_AZURE_INSIGHT_DB_USER_SUFFIX},
            @{key = "[AZURE_INSIGHT_DB_PASSWORD]"; value = $CONFIG_AZURE_INSIGHT_DB_PASSWORD}
        )
        Set-Location $conv_ledgerFolder
        $configFiles = @((Get-ChildItem -Path '*.Config' -Include "*.config").FullName)
        Set-Location $executionFolder
        foreach ($f in $configFiles) {
            wh "Configuring app settings for $f"
            setConfigs $f $hashConfigurations
        }
    }
}
#check connection string and app settings in notepad++ by the fucking eyes
#this step will be ran in conversion machine
function checkConfig() {
    Set-Location $conv_ledgerFolder
    $configFiles = @((Get-ChildItem -Path '*.Config' -Include "*.config").FullName)
    foreach ($f in $configFiles) {
        Write-Host "$f"
    }
    wh "There are $($configFiles.Count) config files, now open all of it to check each settings"
    foreach ($f in $configFiles) {
        Start-Process notepad++ $f
    }
    Set-Location $executionFolder
}

#apply config of connection strings and app settings for a fucking lot of apps
#basically, just run a batch file to copy config
#this step will be ran in conversion machine
function applyConfig() {
    Set-Location $conv_ledgerFolder
    & $conv_copyCustomConfigScriptPath
    Set-Location $executionFolder
}

#now check config files in main app again to see whether step4 is performed correctly
function recheckConfig() {
    Set-Location $conv_ledgerFolder
    $folders = @((Get-ChildItem -Exclude "*.Config" | Where-Object {$_.PSIsContainer}).FullName)
    foreach ($d in $folders) {
        Write-Host "$d"
        $configFiles = @((Get-ChildItem -Path $d -File -Filter "Custom*.config").FullName)
        if ($configFiles.Length -ne 0) {
            wh "$d contains $($configFiles.Length) files, now open all of it"
            foreach ($f in $configFiles) {
                if (![string]::IsNullOrEmpty($f)) {
                    Write-Host "$f"
                    Start-Process notepad++ $f
                }
                else {
                    wh "$f is emtpy, skipped" $color_warning
                }
            }
        }
    }
    Set-Location $executionFolder
}

#backup and restore ledger database
#firstly, check whether ledger database is existing, if it is, back it up into $conv_ledger_db_backup_path folder
#next, restore ledger database to conversion machine with name the same as ledger's name
#for example, if ledger is Melbourne, now restore to database Melbourne
function restoreDb($backupFile, $dbName) {
    $bakFile = $(Split-Path $backupFile -Leaf).Replace(".bak", "")
    if ($bakFile -ne $dbName) {
        wh "The backup file $backupFile IS NOT THE SAME WITH database name $dbName, you you fucking want to continue? [y/n], default is [n]" $color_warning
        $confirm = Read-Host
        if ($confirm -eq "y") {
            $dataDefaultPath = getSqlDefaultPath 'Data'
            $logDefaultPath = getSqlDefaultPath 'Log'
            $logicalNames = retrieveDatabaseLogicalNames $backupFile
            $restoreQuery = @"
            USE [master]
            RESTORE DATABASE [$conv_ledger_db] FROM  DISK = N'$backupFile' WITH  FILE = 1,  
            MOVE N'$($logicalNames.data)' TO N'$($dataDefaultPath + $conv_ledger_db + ".mdf")',  
            MOVE N'$($logicalNames.log)' TO N'$($logDefaultPath + $conv_ledger_db + "_Log.ldf")',  
            NOUNLOAD,
            REPLACE,
            STATS = 5
"@
            wh "Restoring database with query"
            Write-Host $restoreQuery
            Invoke-Sqlcmd -ServerInstance '.' -Query $restoreQuery -QueryTimeout 900
        }
    }
}

#this step is to restore ledger into conversion machine
function restoreLedgerDb() {
    if ($SOURCE_SYSTEM -eq $WINBEAT) {
        restoreLedgerDbWinbeat
    }
    else {
        if ($SOURCE_SYSTEM -eq $IBAIS) {
            restoreLedgerDbIbais
        }
    }
}
function restoreLedgerDbWinbeat() {
    if (!(Test-Path -Path $conv_ledger_db_backup_path)) {
        wh "$conv_ledger_db_backup_path not found"
        exit
    }
    wh "Restore $conv_ledger_db_backup_path into $conv_ledger_db  database"
    Write-Host
    wh "Restore process is starting now, DO YOU FUCKING SURE? [y/n], default is [n]" $color_warning 0 
    Write-Host
    $confirm = Read-Host 
    if ($confirm -eq 'y') {
        backupDb $conv_ledger_db $conv_defaultDatabaseBackupFolder
        restoreDb $conv_ledger_db_backup_path $conv_ledger_db
    }
}

function restoreLedgerDbIbais() {
    wh "Restore $conv_ledger_db_backup_path into $conv_ledger_db  database"
    Write-Host
    wh "Restore process is starting now, DO YOU FUCKING SURE? [y/n], default is [n]" $color_warning 0 
    Write-Host
    $confirm = Read-Host 
    if ($confirm -eq 'y') {
        backupDb $conv_ledger_db $conv_defaultDatabaseBackupFolder

        #delete the old one
        if ((checkDbExist $conv_ledger_db)) {
            wh "Deleting $conv_ledger_db"
            $delQuery = @"
        USE MASTER 
        GO 
        DROP DATABASE $conv_ledger_db
"@
            Invoke-Sqlcmd -ServerInstance '.' -Query $delQuery
        }
    }
    #create the new one
    if (!(checkDbExist $conv_ledger_db)) {
        wh "Creating $conv_ledger_db"
        $createDbQuery = @"
        USE MASTER 
        GO
        CREATE DATABASE $conv_ledger_db
"@
        wh "$conv_ledger_db does not exist, create it now"
        Invoke-Sqlcmd -ServerInstance '.' -Query $createDbQuery -QueryTimeout 900
    }
}
# this step is optional, run it if the ledger database's collation is not SQL_Latin1_General_CP1_CI_AS
function changeCollationLedgerDb() {
    $viewCollationQuery = @"
    SELECT collation_name FROM sys.databases
    WHERE [name] = '$conv_ledger_db'
"@
    $collRows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $viewCollationQuery)
    if ($($collRows).Count -ne 0) {
        $collation = $collRows[0]["collation_name"]
        wh $collation
        if (!$collation.Equals("SQL_Latin1_General_CP1_CI_AS")) {
            if (!(Test-Path -Path $conv_changeCollationSqlScriptPath)) {
                wh "$conv_changeCollationSqlScriptPath not found " $color_error
            }
            else {
                ssms.exe  $conv_changeCollationSqlScriptPath -d $conv_ledger_db -E -S '.'
            }
        }
    }
}

#copy insight create db script
#this step should be performed in fucking build machine
#since only in this machine contain build version for insight boa
#the path to get script is configured in F:\Build\boa\publish\$ver\SQL\CreateScripts\BOALedger
#which $ver is the fucking current version of boa's build
#where to find it? fucking go to runsheet, somewhere in StageBoaLedger command
function copyInsightCreationScript() {
    wh "Which insight version do u fucking want to get sql script?" $color_warning
    $ver = Read-Host
    $path = "F:\Build\boa\publish\$ver\SQL\CreateScripts\BOALedger"

    # since we dont fucking want to copy this script and run it in build machine so 
    # we just put the path to deploy folder here, fucking use your eyes and hands to copy it to conversion machine
    Set-Clipboard "$path"
    wh "$path was copied to clipboard"

    # if (!(Test-Path -Path $path)) {
    #     wh "$path not found" $color_error
    #     exit 0
    # }
    # $filePaths = @((Get-ChildItem -LiteralPath $path | sort {$_.BaseName}).FullName)
    # foreach ($f in $filePaths) {
    #     Write-Host "$f"
    # }
    # if ($filePaths.Count -eq 0) {
    #     wh "There is no files in $path"
    #     exit 0
    # }
    # $lastestFile = $filePaths[$filePaths.Count - 1]
    # Set-Clipboard -LiteralPath $lastestFile
    # wh "Copied $lastestFile to clipboard, now fucking paste it to [conversion machine]$conv_ledgerFolder\Admin"
}

#create LedgerInsight database, run create schema script from step7 by start ssms and run by fucking hand
#this step should be performed in conversion machine
#firstly, check whether [LedgerName]Insight database is existing, if it's not, create one
#next, open the create script copied from step7 (this script should be put in $conv_ledgerFolder\Admin)
#now fucking run it by your fucking hand
function createInsightDb() {
    if ((checkDbExist $conv_ledger_insight_db)) {
        wh "$conv_ledger_insight_db already exist, backing it up and delete the old one"
        backupDb $conv_ledger_insight_db $conv_defaultDatabaseBackupFolder
        #delete the old one
        $delQuery = @"
        USE MASTER 
        GO 
        DROP DATABASE $conv_ledger_insight_db
"@
        Invoke-Sqlcmd -ServerInstance '.' -Query $delQuery
    }

    if (!(checkDbExist $conv_ledger_insight_db)) {
        
        $createDbQuery = @"
        USE MASTER 
        GO
        CREATE DATABASE $conv_ledger_insight_db
"@
        wh "$conv_ledger_insight_db does not exist, create it now"
        Invoke-Sqlcmd -ServerInstance '.' -Query $createDbQuery -QueryTimeout 900
    }
    $createdbScriptFolder = "$conv_ledgerFolder\Admin"
    if (!(Test-Path $createdbScriptFolder)) {
        wh "Folder $createdbScriptFolder does not exsited" $color_error
        return
    }
    $scripts = @((Get-ChildItem -LiteralPath $createdbScriptFolder -Filter "*BOALedgerCreate*").FullName)
    if ($scripts.Count -eq 0) {
        wh "There is no sql create script in $createdbScriptFolder"
    }
    $createdbScript = $scripts[0]
    wh "Opening $createdbScript to run"
    SQLCMD.EXE -i $createdbScript -d $conv_ledger_insight_db -E -S '.'

    #check policies table to make sure there's no rows in here
    $query = "SELECT TOP 1 * FROM policies"
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query -Database $conv_ledger_insight_db)
    if ($rows.Count -eq 0) {
        wh "Created $conv_ledger_insight_db database successfully"
    }
    else {
        wh "There was an issue occured: policies table already contains values" $color_error
    }
}

#run audit tools
function isTableExist($dbName, $tableName)
{
    $query = @"
    SELECT TOP 1 * 
    FROM $dbName.INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = '$tableName'
"@
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
    if($rows.Count -eq 0)
    {
        return $false
    }
    return $true
}
function runSunriseExport() {
    $query = ""
    $sunriseCredentials = @()
    $un = ""
    $pw = ""
    if(isTableExist $conv_ledger_db "SunriseServer")
    {
        $query = "select top 1 * from $conv_ledger_db..[SunriseServer] where code = 'INSNET'"
        $sunriseCredentials = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
        if ($($sunriseCredentials.Count) -ne 0) {
            $un = $sunriseCredentials[0]["Login"]
            $pw = $sunriseCredentials[0]["Password"]
        }
    }
    else {
        $query = "select top 1 * from $conv_ledger_insight_db..[sunrise_server_codes] where sunserco_name = 'INSNET'"
        $sunriseCredentials = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
        if ($($sunriseCredentials.Count) -ne 0) {
            $un = $sunriseCredentials[0]["sunserco_login"]
            $pw = $sunriseCredentials[0]["sunserco_password"]
        }
    }
    
    $sunriseExportToolFolder = "$conv_ledgerFolder\boa-sunrise-export"
    Set-Location $sunriseExportToolFolder
    start "$sunriseExportToolFolder\SunriseExport.exe"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 500
    Set-Location -Path $executionFolder
    auditAutomationTool -procName "SunriseExport" -controlId cbxSunriseURL -controlValue "Production Proxy" 
    auditAutomationTool -procName "SunriseExport" -controlId cbxVersion -controlValue "Latest version only" 
    auditAutomationTool -procName "SunriseExport" -controlId txtSunriseUsername -controlValue "$un" 
    auditAutomationTool -procName "SunriseExport" -controlId txtSunrisePassword -controlValue "$pw" 
}

#run 3 audit tools in order: 
#   1. svu audit
#   2. sunrise export 
#   3. sunrise audit
#wait a few seconds for each run
function runAuditTools() {
    $azureInsightDbConnString = getConfigValue "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\CustomConnectionStrings.config" 'connectionStrings/add[@name="DestinationDatabase"]' "connectionString"
    $svuAuditToolFoler = "$conv_ledgerFolder\boa-svu-audit"
    Set-Location -Path $svuAuditToolFoler
    start "$svuAuditToolFoler\SvuAudit.exe"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 500
    Set-Location -Path $executionFolder
    $listingFile = getConfigValue "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\CustomAppSettings.config" 'appSettings/add[@key="SvuCSVFilePath"]' "value"
    if(!$listingFile)
    {
        $listingFile = "$conv_ledgerFolder\Raw data\$conv_SVUListingFile"
        if(!(Test-Path -Path $listingFile))
        {
            wh "$listingFile does not existed" $color_warning
        }
    }
    auditAutomationTool -procName "SvuAudit" -controlId txtOpportunityFile -controlValue $listingFile
    auditAutomationTool -procName "SvuAudit" -controlId txtConnection -controlValue "$azureInsightDbConnString" 
    countDown 60 "Wait 60s to run Sunrise export"

    runSunriseExport
    countDown 120 "Wait 120s to run Sunrise audit"

    $sunriseAuditToolFolder = "$conv_ledgerFolder\boa-sunrise-audit"
    Set-Location $sunriseAuditToolFolder
    start "$sunriseAuditToolFolder\boa-sunrise-audit.exe"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 500
    Set-Location -Path $executionFolder
    $outputFileFromSunriseExport = Get-ChildItem "$conv_ledgerFolder\boa-sunrise-export\Output" | Sort {$_.LastWriteTime} | select -last 1
    auditAutomationTool -procName "boa-sunrise-audit" -controlId txtPolicyFile -controlValue "$($outputFileFromSunriseExport.FullName)" 
    auditAutomationTool -procName "boa-sunrise-audit" -controlId txtConnection -controlValue "$azureInsightDbConnString" 
}
function countDown($seconds, $message)
{
    foreach($count in (1..$seconds))
    {
        Write-Progress -Id 1 -Activity "$message" -Status "$($seconds - $count) left" -PercentComplete (($count / $seconds) * 100)
        Start-Sleep -Seconds 1
    }
}

#check if folder is empty or not
function isEmptyFolder($folder) {
    if (!(Test-Path -Path $folder)) {
        wh "$folder does not exist" $color_warning
        return $true
    }
    $directoryInfo = Get-ChildItem $folder | Measure-Object
    if ($($directoryInfo.Count) -eq 0) {
        wh "$folder is empty" $color_warning
        return $true
    }
    return $false
}
#copy all files from source folder to destination folder
function copyReportsFromFolder($sourceFolder, $destFolder) {
    if (!(isEmptyFolder $sourceFolder)) {
        createFolderIfNotExists $destFolder
        $itemsCount = $(Get-ChildItem -Path "$sourceFolder\*" | Measure-Object).Count
        wh "Coping $itemsCount items from $sourceFolder to $destFolder"
        Copy-Item -Path "$sourceFolder\*" -Destination "$destFolder" -Force
    }
}

#copy 1 latest file from source folder to destination folder
function copyReportFile($sourceFolder, $destFolder) {
    if (!(isEmptyFolder $sourceFolder)) {
        createFolderIfNotExists $destFolder
        $file = Get-ChildItem "$sourceFolder" | Sort {$_.LastWriteTime} | select -last 1
        wh "Coping $($file.FullName) from $sourceFolder to $destFolder"
        Copy-Item -Path "$($file.FullName)" -Destination "$destFolder" -Force
    }
}
function prepareReports() {
    createFolderIfNotExists $conv_automationReportsFolder

    #copy pre upload reports to automation reports folder 
    copyReportsFromFolder $conv_preUploadReportsFolder "$conv_automationReportsFolder\PreUploadReports"

    #copy post conversion data verification reports to automation reports folder
    copyReportsFromFolder $conv_postConversionDataVerificationReportsForConsultantFolder "$conv_automationReportsFolder\PostConversionDataVerificationsReports"
    
    #copy svu audit result to automation reports folder 
    copyReportFile $conv_svuAuditResultPath $conv_automationReportsFolder

    #copy sunrise audit result to automation reports folder
    copyReportFile $conv_sunriseAuditResultPath $conv_automationReportsFolder

    #copy records count reports to automation reports folder
    $recordCountFolder = "$conv_automationReportsFolder\Records Count"
    copyReportsFromFolder "$conv_recordCountReportsFolder" "$recordCountFolder"
    copyReportFile "$conv_ledgerFolder\Run1\Conversion Record Counts.xlsx" $recordCountFolder
    copyReportFile "$conv_ledgerFolder\Run1\PreConversionRecordCount.txt" $recordCountFolder
    copyReportFile "$conv_ledgerFolder\Run1\PostConversionRecordCount.txt" $recordCountFolder
}

### RUNSHEET FUNCTIONS ###
# Check whether blobs folder is existing, if it is, rename it to ..._[current date time]
function backupBlobsFolder() {
    $blobFolderConvMachine = "Source data from setup"
    $newName = $blobFolderConvMachine + "_$(now)"
    $path = "$conv_ledgerFolder\$blobFolderConvMachine"
    if (!(Test-Path -Path $path)) {
        wh "'$path' does not existed, skip backup blobs folder" $color_warning
        return
    }
    Rename-Item -Path  -NewName $newName
    wh "$conv_ledgerFolder\$blobFolderConvMachine was renamed to $newName"
}

main