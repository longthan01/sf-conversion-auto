Param(
    [string]$task = 'step3-config'
)

# global variables #
$executionFolder = Split-Path $MyInvocation.MyCommand.Path
#path to msbuild, use for auto build
Set-Alias msbuild "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSbuild.exe"
<# Set-Alias nuget "" #>

#THE LEDGER BEING MIGRATE
$LEDGER = "Osman"

$CONFIG_AZURE_BLOB_STORAGE_ACCOUNT = "FUCKINGTEST"
$CONFIG_AZURE_BLOB_STORAGE_KEY = "FUCKINGTEST"
$CONFIG_AUDIT_LISTING_FILE = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB_USER_SUFFIX = "FUCKINGTEST"
$CONFIG_AZURE_INSIGHT_DB_PASSWORD = "FUCKINGTEST"

#local working folder (in your development machine), if is current path, must be set to ".\"
$local_workingFolder = "d:\conversion_auto\$LEDGER"

#path to your repo folders
$local_conversionRootPath = 'D:\sfg-repos\insight_data_conversion\boa-data-conversion'
$local_sunriseAuditRootPath = 'D:\sfg-repos\boa-sunrise-audit'
$local_sunriseExportRootPath = 'D:\sfg-repos\boa-sunrise-export'
$local_svuAuditRootPath = 'D:\sfg-repos\boa-svu-audit'

#conversion machine paths
$conv_ledgerName = $LEDGER
$conv_ledgerFolder = $local_workingFolder #test path on local machine
#$conv_ledgerFolder = "F:\$conv_ledgerName" #real path in conversion machine
$conv_appSourceFolder = $conv_ledgerFolder
$conv_automationBackupFolder = $conv_ledgerFolder + "\automation_backups"
$conv_defaultDatabaseBackupFolder = "$conv_automationBackupFolder\database"
$conv_changeCollationSqlScriptPath = "$conv_ledgerFolder\Change_Collation.sql"
$conv_copyCustomConfigScriptPath = "$conv_ledgerFolder\Copy_Custom_Config.cmd"
$conv_ledger_db_backup = "$conv_ledgerFolder\Raw Data\$conv_ledgerName.bak" #"F:\DKM.Services\Database scripts\2792018.bak"
$conv_ledger_db = $conv_ledgerName
$conv_ledger_insight_db = "$conv_ledgerName" + "Insight"

#utility variables
$color_info = 'green'
$color_warning = 'yellow'
$color_error = 'red'

function printUsage() {
    wh "Params:"
    wh "-task"
    wh "`t(runsheet)rs-backupblobsfolder"
    wh
    wh "`t(conv)step0-pullcode"
    wh "`t(dev)step1-build"
    wh "`t(conv)step2-prepare"
    wh "`t(conv)step3-zip"
    wh "`t(conv)step4-extract"
    wh "`t(conv)step5-config"
    wh "`t(conv)step6-checkconfig"
    wh "`t(build)step7-applyconfig"
    wh "`t(conv)step8-recheckconfig"
    wh "`t(conv)step9-restoreledgerdb"
    wh "`t(conv)step10-changecollation"
    wh "`t(conv)step11-copycreateinsightdbscript"
    wh "`t(conv)step12-createinsightdb"
}
function main() {
    $conv_ledgerFolder = replaceIfCurrentPath $conv_ledgerFolder
    $conv_appSourceFolder = replaceIfCurrentPath $conv_appSourceFolder
    $conv_defaultDatabaseBackupFolder = replaceIfCurrentPath $conv_defaultDatabaseBackupFolder
    $conv_changeCollationSqlScriptPath = replaceIfCurrentPath $conv_changeCollationSqlScriptPath
    $conv_copyCustomConfigScriptPath = replaceIfCurrentPath $conv_copyCustomConfigScriptPath
    $conv_ledger_db_backup = replaceIfCurrentPath $conv_ledger_db_backup

    if (($task -eq 'h') -Or ([string]::IsNullOrEmpty($task))) {
        printUsage
    }
    if ($task -eq 'step0-pullcode') {
        pullLatestCode
    }
    if ($task -eq 'step1-build') {
        buildSolutions
    }
    if ($task -eq 'step2-prepare') {
        prepareConvEnvironment
    }

    if ($task -eq 'step3-zip') {
        archive
    }
   
    if ($task -eq 'step4-extract') {
        extract
    }
    
    if ($task -eq 'step5-config') {
        config
    }
    
    if ($task -eq 'step6-checkconfig') {
        checkConfig
    }

    if ($task -eq 'step7-applyconfig') {
        applyConfig
    }
    if ($task -eq 'step8-recheckconfig') {
        recheckConfig
    }
    if ($task -eq 'step9-restoreledgerdb') {
        restoreLedgerDb
    }
    if ($task -eq 'step10-changecollation') {
        changeCollationLedgerDb
    }
    
    if ($task -eq 'step11-copycreateinsightdbscript') {
        copyInsightCreationScript
    }
    if ($task -eq 'step12-createinsightdb') {
        createInsightDb
    }
    
    if ($task -eq 'rs-backupblobsfolder') {
        backupBlobsFolder
    }
}

# COMMON FUNCTIONS #
function createFolderIfNotExists($path) {
    if (!(Test-Path -Path $path)) {
        wh "Folder $path does not existed, creating new one..."
        New-Item -ItemType Directory -Force -Path $path
    }
}
function replaceIfCurrentPath($path) {
    if ($path.StartsWith(".\")) {
        return $path.SubString(2)
    }
    return $path
}
function zipFileIntoFolder ($folder, $sourcePath, $destinationPath) {
    #create root folder for automation, if it's not existed
    if (!(Test-Path -Path $folder)) {
        wh "$folder folder is not existed, creating the new one"
        New-Item -ItemType directory -Path $folder
    }
    else {
        # check whether destination path is existed, if it is, do delete
        if (Test-Path -path $destinationPath) {
            # if root folder is existed, delete all of it's items
            Write-Host
            wh "`t[!] '$destinationPath' folder is existing, do you FUCKING WANT TO DELETE? [y/n], default is [n]" $color_warning 1
            Write-Host
            $confirm = Read-Host
            if ($confirm -eq 'y') {
                Remove-Item -Path $destinationPath -Force -Recurse
            }
            else {
                wh "You choose to no delete, exist script now"
                exit 0
            }
        }
    }

    wh "Zipping files in $sourcePath"
    #zip all files in source folder into destination folder
    Compress-Archive -path $sourcePath -DestinationPath $destinationPath
}
function extractFileIntoFolder($sourcePath, $destinationPath) {
    if (!(Test-Path -Path $sourcePath)) {
        Write-Host $sourcePath ' is not found'
        exit 0 
    }
    if (Test-Path $destinationPath) {
        if ((Get-ChildItem $destinationPath | Measure-Object).count -ne 0) {
            Write-Host
            wh "`t[!] '$destinationPath 'is not empty, do you FUCKING WANT TO DELETE? [y/n], default is [n]" $color_warning 0
            Write-Host
            $confirm = Read-Host
            if ($confirm -eq 'y') {
                Remove-Item -Path $destinationPath -Force -Recurse
            }
            else {
                wh "You choose to no delete, exist script now"
                exit 0
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
function buildSolution($solutionPath, $buildMode) {	
    wh $solutionPath
    if (!(Test-Path -Path $solutionPath)) {
        wh "Solution file not found" $color_error
        return
    }

    if ($buildMode -eq $()) {
        $buildMode = "Release"
    }

    # Setup log file
    #$solutionFileName = (Get-Item $solutionPath).Name
    #$logFileName = "logs\" + $solutionFileName + "_log.txt";
    <#  if (Test-Path($logFileName)) {
        Remove-Item $logFileName
    } #>

    wh "Cleaning $solutionPath"
    # Clean up
    msbuild "$solutionPath" /t:clean /p:configuration=$buildMode
		
    if ($LastExitCode -ne 0) {  
        Throw "Clean " + $solutionPath + " failed"  
    }

    # Build all
    msbuild "$solutionPath" /p:Configuration=$buildMode /v:m | Out-Default 
          
    if ($LastExitCode -ne 0) {
        Throw "Build " + $solutionPath + " failed"  
    }
    wh "Build $solutionPath successfully"
}

function buildSolutions()
{
    buildSolution  "$local_conversionRootPath\DatabaseConversion.sln" "Debug"
    buildSolution  "$local_sunriseAuditRootPath\boa-sunrise-audit.sln" "Debug"
    buildSolution  "$local_sunriseExportRootPath\SunriseExport.sln" "Debug"
    buildSolution  "$local_svuAuditRootPath\SvuAudit.sln" "Debug"
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
    Write-Host "`t$checkDbExistQuery"
    $dbExistRows = @(Invoke-Sqlcmd  -ServerInstance '.' -Query $checkDbExistQuery -QueryTimeout 900)
    if ($dbExistRows.Count -eq 0) {
        return $false
    }
    return $true
}

### RUN CONVERSION APP STEPS ###
# prepare environment like create all needing folder, copy audit files, backup Run1, etc...
function prepareConvEnvironment() {
    # create all needing folders
    createFolderIfNotExists "$conv_ledgerFolder\Raw Data"
    createFolderIfNotExists "$conv_ledgerFolder\Admin"

    # backup Run1 if this is the n run (n > 1)
    $run1 = "$conv_ledgerFolder\Run1"
    if (Test-Path -Path $run1) {
        $newName = "Run1_$(now)"
        Rename-Item -Path $run1 -NewName $newName
        wh "$run1 was renamed to $newName"
    }
}

#archive conversion app and related tools, this step should be ran in development machine
function archive() {
    ### ZIP CONVERSION APP ###
    zipFileIntoFolder $local_workingFolder $local_conversionRootPath'\DatabaseConversion.ConsoleApp\bin\Debug\*' $local_workingFolder'\DatabaseConversion.ConsoleApp.zip'

    ### ZIP IMPORT BLOB APP ###
    zipFileIntoFolder $local_workingFolder $local_conversionRootPath'\DatabaseConversion.AzureImportBlob\bin\Debug\*' $local_workingFolder'\DatabaseConversion.AzureImportBlob.zip'

    ### ZIP DATA COUNT CHECKER APP ###
    zipFileIntoFolder $local_workingFolder $local_conversionRootPath'\DataCountChecker\bin\Debug\*' $local_workingFolder'\DataCountChecker.zip'

    ### ZIP IMPORT SCHEDULE APP ###
    zipFileIntoFolder $local_workingFolder $local_conversionRootPath'\ImportedSchedule\bin\Debug\*' $local_workingFolder'\ImportedSchedule.zip'

    ### ZIP AUDIT APPS ###
    zipFileIntoFolder $local_workingFolder $local_sunriseAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-audit.zip'
    zipFileIntoFolder $local_workingFolder $local_sunriseExportRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-export.zip'
    zipFileIntoFolder $local_workingFolder $local_svuAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-svu-audit.zip'
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
                    $add.value = $add.value.Replace($config.key, $config.value)
                    Write-Host $add.value
                }
            }
        }
    }
   
    $connStrings = $appConfig.GetElementsByTagName("connectionStrings")
    if ($connStrings) {
        foreach ($add in $appConfig.connectionStrings.add) {
            if (![string]::IsNullOrEmpty($add.connectionString)) {
                foreach ($config in $configHashArray) {
                    $add.connectionString = $add.connectionString.Replace($config.key, $config.value)
                    Write-Host $add.connectionString
                }
            }
        }
    }
    
    $appConfig.Save($appConfigFilePath)
}
function config() {
    wh "`t[!] We will config app and connection string settings for a bunch of apps" $color_warning
    wh "`t[!] if this is the n-th run (n > 1), this step shouldn't be performed, fucking confirm to continue (y/n)? default is [n]" $color_warning
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
        Write-Host "`t$f"
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
        Write-Host "`t$d"
        $configFiles = @((Get-ChildItem -Path $d -File -Filter "Custom*.config").FullName)
        if ($configFiles.Length -ne 0) {
            wh "`t$d contains $($configFiles.Length) files, now open all of it"
            foreach ($f in $configFiles) {
                if (![string]::IsNullOrEmpty($f)) {
                    Write-Host "`t`t$f"
                    Start-Process notepad++ $f
                }
                else {
                    wh "[!] $f is emtpy, skipped" $color_warning
                }
            }
        }
    }
    Set-Location $executionFolder
}

#backup and restore ledger database
#firstly, check whether ledger database is existing, if it is, back it up into $conv_ledger_db_backup folder
#next, restore ledger database to conversion machine with name the same as ledger's name
#for example, if ledger is Melbourne, now restore to database Melbourne
function restoreDb($backupFile, $dbName) {
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
    wh 'Restoring database with query'
    Write-Host $restoreQuery
    Invoke-Sqlcmd -ServerInstance '.' -Query $restoreQuery -QueryTimeout 900
}

#this step is to restore ledger into conversion machine
function restoreLedgerDb() {
    if (!(Test-Path -Path $conv_ledger_db_backup)) {
        wh $conv_ledger_db_backup ' is not found'
        exit
    }
    wh "Restore $conv_ledger_db_backup into $conv_ledger_db  database"
    Write-Host
    wh "`t[!] Restore process is starting now, DO YOU FUCKING SURE? [y/n], default is [n]" $color_warning 0 
    Write-Host
    $confirm = Read-Host 
    if ($confirm -eq 'y') {
        backupDb $conv_ledger_db $conv_defaultDatabaseBackupFolder
        restoreDb $conv_ledger_db_backup $conv_ledger_db
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
    wh "[!] Which insight version do u fucking want to get sql script?" $color_warning
    $ver = Read-Host
    $path = "F:\Build\boa\publish\$ver\SQL\CreateScripts\BOALedger"

    # since we dont fucking want to copy this script and run it in build machine so 
    # we just put the path to deploy folder here, fucking use your eyes and hands to copy it to conversion machine

    wh $path

    # if (!(Test-Path -Path $path)) {
    #     wh "$path not found" $color_error
    #     exit 0
    # }
    # $filePaths = @((Get-ChildItem -LiteralPath $path | sort {$_.BaseName}).FullName)
    # foreach ($f in $filePaths) {
    #     Write-Host "`t$f"
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
        exit 0
    }
    $scripts = @((Get-ChildItem -LiteralPath $createdbScriptFolder).FullName)
    if ($scripts.Count -eq 0) {
        wh "There is no sql script in $createdbScriptFolder"
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

### RUNSHEET FUNCTIONS ###
function backupBlobsFolder() {
    $blobFolderConvMachine = "Source data from setup"
    $newName = $blobFolderConvMachine + "_$(now)"
    $path = "$conv_ledgerFolder\$blobFolderConvMachine"
    if(!(Test-Path -Path $path))
    {
        wh "'$path' does not existed, skip backup blobs folder" $color_warning
        return
    }
    Rename-Item -Path  -NewName $newName
    wh "$conv_ledgerFolder\$blobFolderConvMachine was renamed to $newName"
}

main