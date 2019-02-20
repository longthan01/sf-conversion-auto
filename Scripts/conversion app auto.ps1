
Param(
    [string]$task = ''
)

###################################################
### Read Name=Value config file into hash table
###################################################
function readConfigFile($filePath) {
    $found = Test-Path($filePath)
    if ($found -eq $false) {
        $msg = "ERROR: Configuration file ", $filePath, " not found"
        Write-Host $msg
        Throw $msg
    }
    
    $nvp = @{}

    $lines = Get-Content $filePath

    foreach ($line in $lines) {
        if ($line.length -eq 0)
        { continue }
        
        if ($line.StartsWith("#"))
        { continue }

        $nv = $line.split("=", 2)
        $nv[0] = $nv[0].Trim()
        $nv[1] = $nv[1].Trim()
        
        if ($nvp.ContainsKey($nv[0])) {
            $nvp.set_item($nv[0], $nv[1])
        }
        else {
            $nvp.add($nv[0], $nv[1])
        }
    }
    
    return $nvp
}

###################################################
# Retrieves a config field value and checks it
###################################################
function getConfigFieldValue($config, $fieldName) {
    if (-not $config.ContainsKey($fieldName))
    { return "" }

    $value = $config[$fieldName]
    return $value
}

#path to msbuild, use for auto build
$MSBUILDS = @( 
    @{name = "Enterprise"; path = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSbuild.exe"},
    @{name = "Community"; path = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSbuild.exe"}
)
foreach($b in $MSBUILDS) {
    if(Test-Path -Path $b.path)
    {
        Set-Alias msbuild "$($b.path)"
        break
    }
}

Set-Alias auditAutomationTool ".\SF-ConversionAuto.exe"

# enums #
$WINBEAT = "WINBEAT"
$IBAIS = "IBAIS"

# Obtains configs from configuration file #
$conv_ledgerConfigFile = "ledger_configs.txt"
$ledgerConfigurations = readConfigFile $conv_ledgerConfigFile


$LEDGER = getConfigFieldValue $ledgerConfigurations "CONFIG_LEDGER"
$SOURCE_SYSTEM = getConfigFieldValue $ledgerConfigurations "CONFIG_SYSTEM"
if ($SOURCE_SYSTEM) {
    $SOURCE_SYSTEM = $SOURCE_SYSTEM.ToUpper()
}
$conv_ledgerName = $LEDGER

#real path to ledger folder in conversion machine
$conv_ledgerFolder = getConfigFieldValue $ledgerConfigurations "CONFIG_LEDGER_FOLDER"
# if there's no config for ledger folder, set default as f:\[ledger name]
if (!$conv_ledgerFolder) {
    $conv_ledgerFolder = "F:\$conv_ledgerName"
}
 
#TEST PATH on local machine, only use for testing purpose at development time, comment out when run in conversion machine
#$conv_ledgerFolder = "D:\conversion_auto\$LEDGER"

# global variables #
$executionFolder = Split-Path $MyInvocation.MyCommand.Path
#local working folder (in your development machine), if is current path, must be set to ".\"
$local_workingFolder = $executionFolder

#path to backup file, only use if the source system is winbeat
#relative path to backup file, in raw data folder 
$conv_ledger_db_backup_file_path = getConfigFieldValue $ledgerConfigurations "CONFIG_LEDGER_DB_BACKUP_PATH"
$conv_ledger_db_backup_file_path = "$conv_ledgerFolder\Raw Data\$conv_ledger_db_backup_file_path"

#################################

#path to your repo folders
function getFirstExistedPath($paths)
{
    foreach($p in $paths)
    {
        if (Test-Path -Path $p)
        {
            return $p
        }
    }
    return ""
}
$LOCAL_SOLUTION_PATHS = @(
    @{name = "consoleApp"; paths = @("D:\sfg-repos\insight_data_conversion\boa-data-conversion");},  
    @{name = "sunriseAudit"; paths = @("D:\sfg-repos\boa-sunrise-audit");},  
    @{name = "sunriseExport"; paths = @("D:\sfg-repos\boa-sunrise-export");},  
    @{name = "svuAudit"; paths = @("D:\sfg-repos\boa-svu-audit");}
)

$local_conversionRootPath = getFirstExistedPath $($LOCAL_SOLUTION_PATHS | Where-Object {$_.name -eq "consoleApp"}).paths
if(!$local_conversionRootPath -or !(Test-Path -Path $local_conversionRootPath))
{
    wh "Console app path did not existed" $color_error
}
$local_sunriseAuditRootPath = getFirstExistedPath $($LOCAL_SOLUTION_PATHS | Where-Object {$_.name -eq "sunriseAudit"}).paths
if(!$local_sunriseAuditRootPath -or !(Test-Path -Path $local_sunriseAuditRootPath))
{
    wh "Sunrise audit path did not existed" $color_error
}
$local_sunriseExportRootPath = getFirstExistedPath $($LOCAL_SOLUTION_PATHS | Where-Object {$_.name -eq "sunriseExport"}).paths
if(!$local_sunriseExportRootPath -or !(Test-Path -Path $local_sunriseExportRootPath))
{
    wh "Sunrise export path did not existed" $color_error
}
$local_svuAuditRootPath = getFirstExistedPath $($LOCAL_SOLUTION_PATHS | Where-Object {$_.name -eq "svuAudit"}).paths
if(!$local_svuAuditRootPath -or !(Test-Path -Path $local_svuAuditRootPath))
{
    wh "SVU audit path did not existed" $color_error
}

#conversion machine paths
$CONV_MACHINES = @(
    @{name = "conv03"; ip = "13.77.2.124:50301"; username = "namph.st76389@stfsazure.onmicrosoft.com"; password = ""}
)

$conv_appSourceFolder = $conv_ledgerFolder
$conv_automationReportsFolder = "$conv_ledgerFolder\automation_reports"
$conv_automationBackupFolder = $conv_ledgerFolder + "\automation_backups"
$conv_defaultDatabaseBackupFolder = "$conv_automationBackupFolder\database"
$conv_defaultSourceCodeBackupFolder = "$conv_automationBackupFolder\source code"
$conv_changeCollationSqlScriptPath = "$conv_ledgerFolder\Change_Collation.sql"
$conv_copyCustomConfigScriptPath = "$conv_ledgerFolder\Copy_Custom_Config.cmd"
$conv_siteSpecificScriptsFolder = "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\SQLScripts\SiteSpecific\$conv_ledgerName"
$recordCountFolder = "$conv_automationReportsFolder\Records Count"
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
$color_important = 'magenta'
function printUsage() {
    wh "Params:"
    wh "-task"
    foreach ($t in $TASKS) {
        wh "`t$($t.name)"
        wh "`t`t$($t.desc)" "cyan"
    }
}
function printVariable($name, $value) {
    wh
    wh "`t`t`$$name`: " "cyan" 0
    wh $value "magenta" 0
}
function printEnvironmentVariables() {
    printVariable "LEDGER" $LEDGER
    printVariable "SOURCE_SYSTEM" $SOURCE_SYSTEM
    printVariable "conv_ledgerFolder" $conv_ledgerFolder
    printVariable "conv_ledger_db_backup_file_path" $conv_ledger_db_backup_file_path
    wh
}

$TASKS = @(
    @{name = "help"; handler = "printUsage"; desc = "Get help from super fucking intelligent AI"},
    @{name = "RunFullConversion"; handler = "RunFullConversion"; desc = "Run setup and start conversion"},
    @{name = "step00-PullCode"; handler = "pullLatestCode"; desc = "Get latest codes from upstream master"},
    @{name = "step01-Build"; handler = "buildSolutions"; desc = "Build converson and related tools"},
    @{name = "step02-Prepare"; handler = "prepareConvEnvironment"; desc = "Prepare the conversion environment: create folder 'Raw data' and 'Admin' if they're not existed, backup folder 'Run1' to 'Run1_[current datetime]' if it's existed"},
    @{name = "step03-Zip"; handler = "archive"; desc = "Archive the conversion and related tools to $local_workingFolder"},
    @{name = "step04-Extract"; handler = "extract"; desc = "Extract the conversion and related tools to $conv_ledgerFolder"},
    @{name = "step05-Config"; handler = "config"; desc = "Read the CONFIG_ variables, then replace it from the custom config file placeholders"},
    @{name = "step06-CheckConfig"; handler = "checkConfig"; desc = "Open custom config files to check ether your configurations are fucking right"},
    @{name = "step07-ApplyConfig"; handler = "applyConfig"; desc = "Run the $conv_copyCustomConfigScriptPath to copy custom config to main config"},
    @{name = "step08-RecheckConfig"; handler = "recheckConfig"; desc = "Open main config files and fucking re-check them by your fucking eyes"},
    @{name = "step09-RestoreLedgerDb"; handler = "restoreLedgerDb"; desc = "Restore $conv_ledger_db, if it's already existed, backup it to $conv_automationBackupFolder"},
    @{name = "step10-ChangeCollation"; handler = "changeCollationLedgerDb"; desc = "Check if the $conv_ledger_db has the right collation, if it's not, open the change collation script in ssms then you need to run it by your fucking hands"},
    @{name = "step11-CopyCreateInsightDbScript"; handler = "copyInsightCreationScript"; desc = "Copy Insight database creation script, this function basically copy out the path to script folder to clipboard. You need to open build machine then paste it to windows explorer to open the folder then copy the latest script by your fucking hands"},
    @{name = "step12-CreateInsightDB"; handler = "createInsightDb"; desc = "Create $($conv_ledgerName + "Insight") database, if the db is already existed, backup it to $conv_automationBackupFolder"},
    @{name = "step13-CheckPreconversionScripts"; handler = "checkPreConversionScripts"; desc = "Open $conv_siteSpecificScriptsFolder\Preconversion folder to see if there's any script need to run before the console app"},
    @{name = "step14-CheckPostConversionScripts"; handler = "checkPostConversionScripts"; desc = "Open $conv_siteSpecificScriptsFolder\Postconversion folder to see if there's any script need to run after the console app"},
    @{name = "step15-RunaAditTools"; handler = "runAuditTools"; desc = "Run SVU Audit tool, Sunrise export and Sunrise Audit tool"},
    @{name = "step16-PrepareReports"; handler = "prepareReports"; desc = "Copy PreUpload, DataVerification, SVU Audit, Sunrise Audit and Record count (Data count checker) reports to automation_reports folder"},
    @{name = "rs-BackupBlobsFolder"; handler = "backupBlobsFolder"; desc = "Check whether blobs folder is existing, if it is, rename it to Source data from setup_[current date time]"},
    @{name = "util-OpenAzureDatabase"; handler = "openAzureDb"; desc = "(conv) Open azure database in ssms"},
    @{name = "util-searchLedgerInRc"; handler = "searchLedgerInRc"; desc = "(conv) Search ledger info in rc environment"},
    @{name = "util-openDatabaseInSSMS"; handler = "openDatabaseInSSMS"; desc = "(conv) Open ledger's database in rc environment"},
    @{name = "util-prepareForRerun"; handler = "prepareForRerun"; desc = "(conv) Rerun conversion in case of the previous failed, this function will do: 1.Rename run1 2.Delete and restore source database 3.Delete and create destination database"},
    @{name = "util-collectLogsAfterRun"; handler = "collectLogs"; desc = "(conv) Collect log files after run conversion for fucking checking purpose"},
    @{name = "util-RecordCountAutoFill"; handler = "fillOutRecordCount"; desc = "AI to fill out record count spreadsheet"},
    @{name = "moddy"; handler = "moddy"; desc = "Temp tool for moddy"}

)
function UngDungTuDongChuyenDoi {
    [CmdletBinding()]
    param()
    DynamicParam {
        $ParameterName = 'task'

        $RunTimeDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary

        $AttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

        $ParamAttribute = New-Object System.Management.Automation.ParameterAttribute
        $ParamAttribute.Mandatory = $false
        $ParamAttribute.Position = 0

        $AttributeCollection.Add($ParamAttribute)

        $ValidateItems = $TASKS.name
        
        $ValidateSetAttribute = New-Object System.Management.Automation.ValidateSetAttribute($ValidateItems)

        $AttributeCollection.Add($ValidateSetAttribute)

        $RunTimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter($ParameterName, [string], $AttributeCollection)
        $RunTimeDictionary.Add($ParameterName, $RunTimeParam)

        Return $RunTimeDictionary
   }
   begin {
    #    if (!$PSBoundParameters[$ParameterName])
    #    {
    #        $PSBoundParameters[$ParameterName] = "help"
    #    }
       $task = $PSBoundParameters[$ParameterName]
       if (!$task)
       {
           $task = "help"
       }
   }
   process{
        $conv_ledgerFolder = replaceIfCurrentPath $conv_ledgerFolder
        $conv_appSourceFolder = replaceIfCurrentPath $conv_appSourceFolder
        $conv_automationBackupFolder = replaceIfCurrentPath $conv_automationBackupFolder
        $conv_automationReportsFolder = replaceIfCurrentPath $conv_automationReportsFolder
        $conv_defaultDatabaseBackupFolder = replaceIfCurrentPath $conv_defaultDatabaseBackupFolder
        $conv_changeCollationSqlScriptPath = replaceIfCurrentPath $conv_changeCollationSqlScriptPath
        $conv_copyCustomConfigScriptPath = replaceIfCurrentPath $conv_copyCustomConfigScriptPath
        $conv_ledger_db_backup_file_path = replaceIfCurrentPath $conv_ledger_db_backup_file_path

        printEnvironmentVariables

        # check if azure context is initialized, if not, call add-azurermaccount
        #prepareAzureContext
        # if (($task -eq 'h') -Or ([string]::IsNullOrEmpty($task))) {
        #     printUsage
        #     return
        # }

        $task = $task.ToLower().Trim()
        foreach ($t in $TASKS) {
            if ($t.name -eq $task) {
                wh "[$($t.name)] USAGE: $($t.desc)"
                wh
                &$t.handler
                return
            }
        }
        # default task if there's no parameter provided 
        printUsage
   }
}

# COMMON FUNCTIONS #
function yesNo($default)
{
    $confirm = (Read-Host).Trim().ToLower()
    if(!$confirm)
    {
        $confirm = $default
    }
    else {
        if(($confirm -ne "y") -or ($confirm -ne "n"))
        {
            $confirm = $default
        }
    }
    return $confirm.ToLower()
}

function prepareAzureContext() {
    $azureContext = Get-AzureRmContext
    if (!$($azureContext.Account)) {
        wh "Azure context did not initialized, initialize now..."
        Add-AzureRmAccount
    }
}
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
function zipFile ($sourcePath, $destinationPath) {
   
    # check whether destination path is existed, if it is, do delete
    if (Test-Path -path $destinationPath) {
        # if root folder is existed, delete all of it's items
        Write-Host
        wh "'$destinationPath' folder is existing, do you FUCKING WANT TO DELETE? [y/n], default is [n]" $color_warning 1
        Write-Host
        $confirm = yesNo "n"
        if ($confirm -eq "y") {
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
            $confirm = (Read-Host).Trim()
            if ($confirm -eq "y") {
                Remove-Item -Path $destinationPath -Force -Recurse
            }
            else {
                wh "You choose to no delete, extraction will overrite the current ones"
            }
        }
    }
    
    wh "Extracting file $sourcePath into $destinationPath"
    Expand-Archive -LiteralPath $sourcePath -DestinationPath $destinationPath -Force
}
function getSqlDefaultPath($type) {
    $query = "SELECT SERVERPROPERTY('InstanceDefault" + $type + "Path') as [Path]"
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
    $result = $rows[0]["Path"]
    return $result
}
function pullCodeFromUpstream($sourceDir) {
    if(!(Test-Path -path $sourceDir))
    {
        wh "FOLDER $sourceDir DID NOT EXISTS" $color_error
        return
    }
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
    if (!(Test-Path -Path $solutionPath)) {
        wh "Solution $solutionPath not found" $color_error
        return
    }
    
    $compilingProjects = getBuildProjects $(Split-Path -Path $solutionPath) $excludeProjects
    wh "`tBuilding $compilingProjects of $solutionPath"

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
    if (!(Test-Path -Path $appConfigFilePath)) {
        wh "$appConfigFilePath not found" $color_error
        return $null
    }
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
        return
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
function doWorkRemotely($remoteWorkingFolder, $serverName, $userName, $password, $callback) {
    $psDriveName = ""
    $psDriveCode = 90 #start from Z
    for ($i = $psDriveCode; $i -ge 65; $i--) {
        $byte = @($i)
        $psDriveName = [System.Text.Encoding]::ASCII.GetString($byte)
        $serializedRemotePath = $remoteWorkingFolder.Replace(":", "$")
        if (!(Test-Path -Path "$psDriveName`:")) {
            $pwd = ConvertTo-SecureString -String "$password" -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PsCredential("$userName", $pwd)
            wh "\\$serverName\$serializedRemotePath"
            New-PSDrive -Name "$psDriveName" -PSProvider filesystem -Root "\\$serverName\$serializedRemotePath" -Credential $cred -Scope global
            &$callback $psDriveName
            Remove-PSDrive "$psDriveName"
            return
        }
    }
    throw "Drive $psDriveName existed"
}
# END COMMON FUNCTIONS #

### RUN CONVERSION APP STEPS ###
# prepare environment like create all needing folder, copy audit files, backup Run1, backup old source codes, etc...
function prepareConvEnvironment() {
    # create all needing folders
    createFolderIfNotExists "$conv_ledgerFolder\Raw Data"
    createFolderIfNotExists "$conv_ledgerFolder\Admin"
    createFolderIfNotExists "$conv_defaultDatabaseBackupFolder"

    #backup old source code if this is the n run (n > 1)
    backupOldSourceCode
    backupRun1
}
function backupRun1() {
    #backup Run1 if this is the n run (n > 1)
    $run1 = "$conv_ledgerFolder\Run1"
    if (Test-Path -Path $run1) {
        $newName = "Run1_$(now)"
        Rename-Item -Path $run1 -NewName $newName
        wh "$run1 was renamed to $newName"
    }
}
function backupOldAdminScripts() {
    
}
function backupOldSourceCode() {
    #create current_date_time folder in $conv_defaultSourceCodeBackupFolder
    $backupFolder = "$conv_defaultSourceCodeBackupFolder\$(now)"
    createFolderIfNotExists $backupFolder
    $adminBackupFolder = "$backupFolder\Admin"
    createFolderIfNotExists $adminBackupFolder
    $confirm = ""
    $adminFolder = "$conv_ledgerFolder\Admin"

    if (!(Test-Path -Path $adminFolder)) {
        wh "There's no admin folder in $conv_ledgerFolder"
    }
    else {
        wh "Do you fucking want to backup admin script (BOALedgerCreate.sql)? [y/n], default is [n]" $color_warning
        $confirm = (Read-Host).Trim()
        if ($confirm -eq "y") {
            wh "Moving old BOALedgerCreate.sql to $adminBackupFolder"
            Move-Item -Path "$adminFolder\*BOALedgerCreate.sql" -Destination $adminBackupFolder
        }
    }
   
    wh "Do you fucking want to backup old source codes? [y/n], choose yes if you want to use the new source code, or no if YOU WANT TO RE-USE THE CURRENT SOURCE CODES, default is [n]" $color_warning
    $confirm = (Read-Host).Trim()
    if ($confirm -eq "y") {
        
        wh "Backing up old source codes into $backupFolder"
        # backup console app
        backupFolder "$conv_ledgerFolder\DatabaseConversion.ConsoleApp" $backupFolder @("Logs")
        # backup import blob app
        backupFolder  "$conv_ledgerFolder\DatabaseConversion.AzureImportBlob" $backupFolder @("Logs")
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
    }
       
    wh
    wh "Do you fucking want to backup CONFIGURATION FOLDERS? [y/n], choose [y] if you want to overrite the old ones" $color_warning
    wh "or [n] if YOU WANT TO RE-USE? default is [n]" $color_warning
    $backupConfigurationConfirm = Read-Host
    if ($backupConfigurationConfirm -eq "y") {
        #backup old configuration folders
        $configurationFolders = @((Get-ChildItem -Path "$conv_ledgerFolder\*.Config").FullName)
        foreach ($f in $configurationFolders) {
            if (Test-Path -Path $f) {
                backupFolder "$f" $backupFolder
            }
        }
    }
}
function backupFolder($source, $dest, [string[]]$exclude = @()) {
    if (!(Test-Path $source)) {
        wh "$source folder does not existed, skip backup" $color_warning
        return
    }
    wh "Moving all items from $source to $dest"
    if (!$exclude) {
        Move-Item "$source" "$dest"
    }
    else {
        $directoryName = Split-Path "$source" -Leaf
        Get-ChildItem -Path "$source" -Exclude $exclude | Move-Item -Destination "$dest\$directoryName"
    }
}
#archive conversion app and related tools, this step should be ran in development machine
function archive() {
    createFolderIfNotExists $local_workingFolder

    ### ZIP CONVERSION APP ###
    zipFile $local_conversionRootPath'\DatabaseConversion.ConsoleApp\bin\Debug\*' $local_workingFolder'\DatabaseConversion.ConsoleApp.zip'

    ### ZIP IMPORT BLOB APP ###
    zipFile $local_conversionRootPath'\DatabaseConversion.AzureImportBlob\bin\Debug\*' $local_workingFolder'\DatabaseConversion.AzureImportBlob.zip'

    ### ZIP DATA COUNT CHECKER APP ###
    zipFile $local_conversionRootPath'\DataCountChecker\bin\Debug\*' $local_workingFolder'\DataCountChecker.zip'

    ### ZIP IMPORT SCHEDULE APP ###
    zipFile $local_conversionRootPath'\ImportedSchedule\bin\Debug\*' $local_workingFolder'\ImportedSchedule.zip'

    ### ZIP AUDIT APPS ###
    zipFile $local_sunriseAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-audit.zip'
    zipFile $local_sunriseExportRootPath'\bin\debug\*' $local_workingFolder'\boa-sunrise-export.zip'
    zipFile $local_svuAuditRootPath'\bin\debug\*' $local_workingFolder'\boa-svu-audit.zip'

    # $templateFiles = ""
    # if ($SOURCE_SYSTEM -eq $WINBEAT) {
    #     $templateFiles = "$executionFolder\appconfig_template_winbeat.zip"
    # }
    # else {
    #     if ($SOURCE_SYSTEM -eq $IBAIS) {
    #         $templateFiles = "$executionFolder\appconfig_template_ibais.zip"
    #     }
    # }
    # wh $templateFiles
    # if (!($templateFiles) -or !(Test-Path -Path $templateFiles)) {
    #     wh "There's no configuration template files in $executionFolder" $color_warning
    #     return 
    # }
    # else {
    #     wh "Copying $templateFiles into $local_workingFolder" 
    #     Copy-Item -Path $templateFiles -Destination "$local_workingFolder" -Force
    # }
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
    $templateFiles = ""
    if ($SOURCE_SYSTEM -eq $WINBEAT) {
        $templateFiles = "$conv_appSourceFolder\appconfig_template_winbeat.zip"
    }
    else {
        if ($SOURCE_SYSTEM -eq $IBAIS) {
            $templateFiles = "$conv_appSourceFolder\appconfig_template_ibais.zip"
        }
    }
    
    if (!($templateFiles) -or !(Test-Path -Path $templateFiles)) {
        wh "Template files: $templateFiles not found" $color_warning
        return 
    }
    $currentConfigFolders = (Get-ChildItem -Path "$conv_ledgerFolder\*.Config" | Measure-Object).Count
    if ($currentConfigFolders -ne 0) {
        wh "Configuration folders are existed, check and extract by your fucking hands, the tool does not support extract programmatically." $color_warning
        return
    }
    wh "Extracting file $templateFiles into $conv_ledgerFolder"
    Expand-Archive -LiteralPath "$templateFiles" -DestinationPath $conv_ledgerFolder -Force
}

#config for each *.config file for a fucking bunch of projects
function setConfigs($appConfigFilePath, $configHashArray) {
    wh $appConfigFilePath "white"
    $appConfig = New-Object Xml
    $appConfig.Load($appConfigFilePath)
    $appSettings = $appConfig.GetElementsByTagName("appSettings")
    if ($appSettings) {
        foreach ($add in $appConfig.appSettings.add) {
            if (![string]::IsNullOrEmpty($add.value)) {
                foreach ($config in $configHashArray) {
                    $contains = $add.value.Contains($config.key)
                    $add.value = $add.value.Replace($config.key, $config.value)
                    if ($contain) {
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
                    if ($contain) {
                        Write-Host $add.connectionString
                    }
                }
            }
        }
    }
    
    $appConfig.Save($appConfigFilePath)
}
function config() {
    if (!(Test-Path -Path $conv_ledgerConfigFile)) {
        wh "$conv_ledgerConfigFile did not exist" $color_error
        return
    }
    $credentialConfigs = readConfigFile $conv_ledgerConfigFile

    $azureBlobStorageAccountName = getConfigFieldValue $credentialConfigs "CONFIG_AZURE_BLOB_STORAGE_ACCOUNT"
    $azureBlobStorageAccountKey = getConfigFieldValue $credentialConfigs "CONFIG_AZURE_BLOB_STORAGE_KEY"
    $auditListingFileName = getConfigFieldValue $credentialConfigs "CONFIG_AUDIT_LISTING_FILE"
    $azureInsightDatabase = getConfigFieldValue $credentialConfigs "CONFIG_AZURE_INSIGHT_DB"
    $azureInsightDatabaseUser = getConfigFieldValue $credentialConfigs "CONFIG_AZURE_INSIGHT_DB_USER"
    $azureInsightDatabasePassword = getConfigFieldValue $credentialConfigs "CONFIG_AZURE_INSIGHT_DB_PASSWORD"

    wh "We will config app and connection string settings for a bunch of apps" $color_warning
    wh "if this is the n-th run (n > 1), this step shouldn't be performed, fucking confirm to continue (y/n)? default is [n]" $color_warning
    $confirm = (Read-Host).Trim()
    if ($confirm -eq "y") {
        $hashConfigurations = @(
            @{key = "[AZURE_BLOB_STORAGE_ACCOUNT]"; value = $azureBlobStorageAccountName},
            @{key = "[AZURE_BLOB_STORAGE_KEY]"; value = $azureBlobStorageAccountKey},
            @{key = "[LEDGER_NAME]"; value = $conv_ledgerName},
            @{key = "[LEDGER_FOLDER]"; value = $conv_ledgerFolder},
            @{key = "[LEDGER_RAW_DATA]"; value = $conv_ledger_db_backup_file_path},
            @{key = "[AUDIT_LISTING_FILE]"; value = $auditListingFileName},
            @{key = "[CONV_LEDGER_INSIGHT_DB]"; value = $conv_ledger_insight_db},
            @{key = "[CONV_LEDGER_DB]"; value = $conv_ledger_db},
            @{key = "[AZURE_INSIGHT_DB]"; value = $azureInsightDatabase},
            @{key = "[AZURE_INSIGHT_DB_USER]"; value = $azureInsightDatabaseUser},
            @{key = "[AZURE_INSIGHT_DB_PASSWORD]"; value = $azureInsightDatabasePassword}
        )
        Set-Location $conv_ledgerFolder
        $configFiles = @((Get-ChildItem -Path '*.Config' -Include "*.config").FullName)
        if($($configFiles).Count -eq 0)
        {
            wh "There's no config files in folder $conv_ledgerFolder" $color_error
            return;
        }
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
#firstly, check whether ledger database is existing, if it is, back it up into $conv_ledger_db_backup_file_path folder
#next, restore ledger database to conversion machine with name the same as ledger's name
#for example, if ledger is Melbourne, now restore to database Melbourne
function restoreDb($backupFile, $dbName) {
    $bakFile = $(Split-Path $backupFile -Leaf).Replace(".bak", "")
    if ($bakFile -ne $dbName) {
        wh "The backup file $backupFile IS NOT THE SAME WITH database name $dbName, you you fucking WANT TO CONTINUE? [y/n], default is [y]" $color_warning
        $confirm = (Read-Host).Trim()
        if ($confirm -eq "n") {
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
    wh "Do you want to backup db first? [y/n], default is n" $color_warning
    $confirm = yesNo "n"
    $backup = $false    
    if ($confirm -eq "y") {
        $backup = $true
    }

    if ($SOURCE_SYSTEM -eq $WINBEAT) {
        restoreLedgerDbWinbeat $backup
    }
    else {
        if ($SOURCE_SYSTEM -eq $IBAIS) {
            restoreLedgerDbIbais $backup
        }
    }
}
function restoreLedgerDbWinbeat($backup) {

    if (!(Test-Path -Path $conv_ledger_db_backup_file_path)) {
        wh "$conv_ledger_db_backup_file_path not found"
        exit
    }
    wh "Restore $conv_ledger_db_backup_file_path into $conv_ledger_db  database"
    Write-Host
    
    if ($backup) {
        backupDb $conv_ledger_db $conv_defaultDatabaseBackupFolder
    }
    restoreDb $conv_ledger_db_backup_file_path $conv_ledger_db
}

function restoreLedgerDbIbais($backup) {
    wh "Restore $conv_ledger_db_backup_file_path into $conv_ledger_db  database"
    Write-Host
    if ($backup) {
        backupDb $conv_ledger_db $conv_defaultDatabaseBackupFolder
    }

    #delete the old one
    if ((checkDbExist $conv_ledger_db)) {
        wh "Deleting $conv_ledger_db"
        $delQuery = @"
        USE MASTER 
        ALTER DATABASE [$conv_ledger_db] SET SINGLE_USER WITH ROLLBACK IMMEDIATE
        GO 
        DROP DATABASE $conv_ledger_db
"@
        Invoke-Sqlcmd -ServerInstance '.' -Query $delQuery
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
        else {
            wh "Collation is good, don't need to change."
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
        wh "Do you FUCKING want to backup insight database first? [y/n], default is no [n]" $color_warning
        $confirm = yesNo "n"
        if ($confirm -eq "y") {
            wh "$conv_ledger_insight_db already exist, backing it up and delete the old one"
            backupDb $conv_ledger_insight_db $conv_defaultDatabaseBackupFolder
        }
        #delete the old one
        $delQuery = @"
        USE MASTER
        GO
        ALTER DATABASE [$conv_ledger_insight_db] SET SINGLE_USER WITH ROLLBACK IMMEDIATE
        GO 
        DROP DATABASE $conv_ledger_insight_db
"@
        wh "Deleting database $conv_ledger_insight_db with query"
        Write-Host $delQuery
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
    $scripts = @((Get-ChildItem -LiteralPath $createdbScriptFolder -Filter "*BOALedgerCreate*" | Sort-Object -Descending).FullName)
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
function isTableExist($dbName, $tableName) {
    $query = @"
    SELECT TOP 1 * 
    FROM $dbName.INFORMATION_SCHEMA.TABLES
    WHERE TABLE_NAME = '$tableName'
"@
    $rows = @(Invoke-Sqlcmd -ServerInstance '.' -Query $query)
    if ($rows.Count -eq 0) {
        return $false
    }
    return $true
}
function runSunriseExport() {
    $sunriseExportTool = "$conv_ledgerFolder\boa-sunrise-export\SunriseExport.exe"
    runProgramOnlyIfAProcessIsExited $sunriseExportTool "SvuAudit"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 1000

    $query = ""
    $sunriseCredentials = @()
    $un = ""
    $pw = ""
    if (isTableExist $conv_ledger_db "SunriseServer") {
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
    
    Set-Location -Path $executionFolder
    auditAutomationTool -procName "SunriseExport" -controlId cbxSunriseURL -controlValue "Production Proxy" 
    auditAutomationTool -procName "SunriseExport" -controlId cbxVersion -controlValue "Latest version only" 
    auditAutomationTool -procName "SunriseExport" -controlId txtSunriseUsername -controlValue "$un" 
    auditAutomationTool -procName "SunriseExport" -controlId txtSunrisePassword -controlValue "$pw" 
    auditAutomationTool -procName "SunriseExport" -controlId btnRun 
}

function checkPreConversionScripts() {
    if (!(Test-Path -Path $conv_siteSpecificScriptsFolder)) {
        wh "Site specific scripts for $conv_ledgerName `: $conv_siteSpecificScriptsFolder not found" $color_warning
        return
    }
    $preConsoleAppFolder = "$conv_siteSpecificScriptsFolder\preconsoleapp"
    if (!(Test-Path -Path $preConsoleAppFolder)) {
        wh "Pre console app folder not found, open and fucking check it by your fucking eyes" $color_warning
        start $conv_siteSpecificScriptsFolder
    }
    else {
        start $preConsoleAppFolder
    }
}

function checkPostConversionScripts() {
    if (!(Test-Path -Path $conv_siteSpecificScriptsFolder)) {
        wh "Site specific scripts for $conv_ledgerName `: $conv_siteSpecificScriptsFolder not found" $color_warning
        return
    }
    $preConsoleAppFolder = "$conv_siteSpecificScriptsFolder\postconsoleapp"
    if (!(Test-Path -Path $preConsoleAppFolder)) {
        wh "Pre console app folder not found, open and fucking check it by your fucking eyes" $color_warning
        start $conv_siteSpecificScriptsFolder
    }
    else {
        start $preConsoleAppFolder
    }
}

#run 3 audit tools in order: 
#   1. svu audit
#   2. sunrise export 
#   3. sunrise audit
#wait a few seconds for each run
function runAuditTools() {
    $azureInsightDbConnString = getConfigValue "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\CustomConnectionStrings.config" 'connectionStrings/add[@name="DestinationDatabase"]' "connectionString"
    runProgramOnlyIfAProcessIsExited "$conv_ledgerFolder\boa-svu-audit\SvuAudit.exe"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 1000
    Set-Location -Path $executionFolder
    $listingFile = getConfigValue "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\CustomAppSettings.config" 'appSettings/add[@key="SvuCSVFilePath"]' "value"
    if (!$listingFile -Or !(Test-Path -Path $listingFile)) {
        wh "Listing file $listingFile does not existed" $color_warning
    }
    auditAutomationTool -procName "SvuAudit" -controlId txtOpportunityFile -controlValue $listingFile
    auditAutomationTool -procName "SvuAudit" -controlId txtConnection -controlValue "$azureInsightDbConnString" 
    auditAutomationTool -procName "SvuAudit" -controlId btnRun 
    
    runSunriseExport
    $sunriseAuditTool = "$conv_ledgerFolder\boa-sunrise-audit\boa-sunrise-audit.exe"
    runProgramOnlyIfAProcessIsExited $sunriseAuditTool "SunriseExport"
    #sleep a few milliseconds to ensure the tool completely started before run the automation tool
    Start-Sleep -Milliseconds 1000
    Set-Location -Path $executionFolder
    $sunriseExportOutput = "$conv_ledgerFolder\boa-sunrise-export\Output"
    if (!(Test-Path -Path $sunriseExportOutput))
    {
        wh "Sunrise export output did not exist" $color_error
        return 
    }
    $outputFileFromSunriseExport = Get-ChildItem $sunriseExportOutput | Sort {$_.LastWriteTime} | select -last 1
    auditAutomationTool -procName "boa-sunrise-audit" -controlId txtPolicyFile -controlValue "$($outputFileFromSunriseExport.FullName)" 
    auditAutomationTool -procName "boa-sunrise-audit" -controlId txtConnection -controlValue "$azureInsightDbConnString" 
    auditAutomationTool -procName "boa-sunrise-audit" -controlId btnRun
}
function runProgramOnlyIfAProcessIsExited($sourceProgram, $otherProcessName)
{
    if (!(Test-Path -Path $sourceProgram)) {
        wh "$sourceProgram does not exist" $color_error
        return
    }
    wh "Awaiting $otherProcessName exit to run $sourceProgram"
    $sourceProgramDir = Split-Path $sourceProgram
    Set-Location $sourceProgramDir
    $otherProcessIsExited = $false
    if (!$otherProcessName)
    {
        $otherProcessIsExited = $true
    }
    $count = 0
    $sourceProgramName = Split-Path -Leaf $sourceProgram
    while (!$otherProcessIsExited)
    {
        $otherProcess = Get-Process $otherProcessName -ErrorAction SilentlyContinue
        if ($otherProcess)
        {
            #wait 1s if other process still running
            $count = $count + 1
            wh "$count try, $otherProcessName still running, could not start $sourceProgramName" $color_warning
            Start-Sleep -Seconds 1 
        }
        else {
            wh "$otherProcess"
            #other process exited, turn off the flag
            $otherProcessIsExited = $true
        }
    }
    wh "Starting $sourceProgram"
    start $sourceProgram
}
function countDown($seconds, $message = "Couting down") {
    foreach ($count in (1..$seconds)) {
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

function openAzureDb() {
    $azureInsightDbConnString = getConfigValue "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\CustomConnectionStrings.config" 'connectionStrings/add[@name="DestinationDatabase"]' "connectionString"
    if ([string]::IsNullOrEmpty($azureInsightDbConnString)) {
        return
    }
    $tokens = $azureInsightDbConnString.Split(";")
    $server = "" 
    $un = ""
    $pw = ""
    $db = ""
    
    foreach ($token in $tokens) {
        if ($token.ToLower().Contains("server=")) {
            $server = $token.Split("=")[1]
        }
        if ($token.ToLower().Contains("user id=")) {
            $un = $token.Split("=")[1]
        }
        if ($token.ToLower().Contains("password=")) {
            $pw = $token.Split("=")[1]
        }
        if ($token.ToLower().Contains("database=")) {
            $db = $token.Split("=")[1]
        }
    }
    wh "Opening database $db in $server with command:"
    wh "ssms -S `"$server`" -d `"$db`" -U `"$un`" -P `"$pw`""
    ssms -S "$server" -d "$db" -U "$un" -P "$pw"
}

function getPasswordFromKeyVault($shortVaultName, $valueKeyName) {
    $vaultName = "sfts-boa-$shortVaultName-keyv-01"
    
    $attempts = 0
    do {
        try {
            $s = Get-AzureKeyVaultSecret -VaultName $vaultName -Name $valueKeyName
            return $s.SecretValueText
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($errorMessage.Contains('The remote name could not be resolved')) {
                Write-Host "    Error $errorMessage. Sleeping for 3 seconds and trying again"
                $attempts = $attempts + 1
                Start-Sleep -Seconds 3
            }
            else {
                Throw $errorMessage
            }        
        }
    } until (($attempts -gt 3))

    if ($attempts -gt 3) {
        Throw "Cannot access key vault $vaultName after 3 attempts"
    }
}
function getControlPanelDbConfig($environmentName, $controlPanelDrActivated = $false) {
    $controlPanelDbConfig = @{}
    $controlPanelDbConfig.ServerName = "boa-control-panel-db-01".Replace('-', '')
    if ($controlPanelDrActivated -eq $true) {
        $controlPanelDbConfig.ServerName = "bdr-control-panel-db-01".Replace('-', '')
    }
    $controlPanelDbConfig.Name = "boa-$environmentName-control-panel"
    $controlPanelDbConfig.Username = "$environmentName-control-panel".Replace('-', '_')
    $controlPanelDbConfig.Password = getPasswordFromKeyVault $environmentName "$environmentName-control-panel-db-password"

    return $controlPanelDbConfig
}


function searchLedgerInRc() {
    wh "Enter ledger display name keyword to search"
    $kwDisplayName = Read-Host

    $query = @"
    SELECT lc.Value AS ConfigValue,
    LC.Name AS ConfigName,
    a.AccountId AS AccountId,
    a.[Name] AS AccountName,
    l.Name AS LedgerName,
    l.DisplayName AS LedgerDisplayName
    FROM LedgerConfigs lc
    JOIN Ledgers l ON lc.LedgerId = l.ID
    CROSS APPLY (
		SELECT ACCOUNTID, [Name]
		FROM ACCOUNTS A 
		WHERE A.DisplayName LIKE '%$kwDisplayName%'
	) a
    WHERE l.DisplayName LIKE '%$kwDisplayName%'
    ORDER BY L.NAME
"@
    $rcConfig = getControlPanelDbConfig "rc"
    $ssmsCommand = "ssms -S $($rcConfig.ServerName).database.windows.net -d $($rcConfig.Name) -U $($rcConfig.Username) -P $($rcConfig.Password)"
    Set-Clipboard $ssmsCommand
    wh "$ssmsCommand is copied to clipboard"
    $rows = @(Invoke-Sqlcmd -ServerInstance "$($rcConfig.ServerName).database.windows.net" -Query $query -Username "$($rcConfig.Username)" -Password "$($rcConfig.Password)" -Database "$($rcConfig.Name)")
    wh "ConfigName`tConfigValue`tAccountId`tAccountName`tLedgerName`tDisplayName" "white"
    $cname = ""
    foreach ($r in $rows) {
        $configName = $r["ConfigName"]
        $configValue = $r["ConfigValue"]
        $accountId = $r["AccountId"]
        if (!$accountId) {$accountId = " "}
        $accountName = $r["AccountName"]
        $ledgerName = $r["LedgerName"]
        $displayName = $r["LedgerDisplayName"]

        wh "$configName`t" "white" 0
        wh "$configValue`t" "white" 0
        wh "$accountId`t" "white" 0
        wh "$accountName`t" "white" 0
        wh "$ledgerName`t" "white" 0
        wh "$displayName`t" "white" 1
        if ($cname -and !$cname.Equals($ledgerName)) {
            wh
        }
        $cname = $ledgerName
    }
}

function openDatabaseInSSMS() {
    wh "Enter ledger name:"
    $ledgerName = Read-Host
    $query = @"
    SELECT lc.Value AS ConfigValue,
    LC.Name AS ConfigName,
    l.Name AS LedgerName,
    a.AccountId,
    l.DisplayName AS LedgerDisplayName,
    i.Name AS InfName
    FROM LedgerConfigs lc
    JOIN Ledgers l ON lc.LedgerId = l.ID
    CROSS APPLY (
		SELECT AccountId
		FROM ACCOUNTS A 
		WHERE A.Id = L.AccountId
    ) A
    CROSS APPLY (
		SELECT I.NAME 
		FROM Infrastructures I 
		JOIN ACCOUNTS Acc ON I.Id = Acc.InfrastructureId
		WHERE A.AccountId = Acc.Id
	) I
    WHERE l.Name = '$ledgerName'
	ORDER BY L.Name
"@
    $rcConfig = getControlPanelDbConfig "rc"
    $rows = @(Invoke-Sqlcmd -ServerInstance "$($rcConfig.ServerName).database.windows.net" -Query $query -Username "$($rcConfig.Username)" -Password "$($rcConfig.Password)" -Database "$($rcConfig.Name)")
    if ($($rows.Count) -ne 0) {
        $row = $null
        foreach ($r in $rows) {
            if ($r["ConfigName"] -eq "LedgerDbPassword") {
                $row = $r
            }
        }
        if (!$row) {
            wh "Not found information for $ledgerName" $color_error
            return
        }
        $ledgerDbPassword = $row["ConfigValue"]
        $accountId = $row["AccountId"]
        if (!$accountId) {$accountId = " "}
        $infName = $row["InfName"]
        $environmentName = "rc"
        $ledgerDbServer = "boa$($environmentName)$($infName)db01.database.windows.net"
        $ledgerDb = "boa$accountId`db$ledgerName"
        $ledgerDbUsername = "boa_$ledgerName"

        $ssmsCommand = "ssms -S `"$ledgerDbServer`" -d `"$ledgerDb`" -U `"$ledgerDbUsername`" -P `"$ledgerDbPassword`""
        Set-Clipboard $ssmsCommand
        wh "$ssmsCommand is copied to clipboard"
        Invoke-Expression $ssmsCommand
    }
}

function prepareForRerun() {
    Write-Host
    wh "SURE [y/n]? default[n]" $color_warning 1
    Write-Host
    $confirm = yesNo "n"
    if ($confirm -eq "y") {
        wh "Backing up run1"   
        backupRun1
        wh "Delete and restoring source db"
        restoreLedgerDb
        changeCollationLedgerDb
        createInsightDb
    }
}

function collectLogs()
{
    $automationLogsFolder = "automation_logs\" + $(now)
    createFolderIfNotExists "$automationLogsFolder"
    $run1Log = "Run1"
    $consoleAppLog = "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\Logs"
    $azureImportBlobLog = "$conv_ledgerFolder\DatabaseConversion.AzureImportBlob\Logs"
    collectLog $run1Log $automationLogsFolder
    collectLog $consoleAppLog $automationLogsFolder $true "consoleapp.log"
    collectLog $azureImportBlobLog $automationLogsFolder $true  "azureimportblob.log"
}
function collectLog($sourceFolder, $destinationFolder, $only1LatestFile, $destinationFileName)
{
    if (!(Test-Path -Path $sourceFolder))
    {
        wh "$sourceFolder did not exist" $color_error
        return
    }
    $logFiles = @((Get-ChildItem -Path "$sourceFolder" -Filter "*.log" | Sort-Object LastWriteTime -Descending).FullName)
    if($logFiles.Count -eq 0)
    {
        wh "There is no log files in $sourceFolder" $color_warning
    } else {
        if (!$only1LatestFile)
        {
            wh "Coping log files from $sourceFolder to $destinationFolder"
            foreach ($f in $logFiles) {
                Copy-Item -Path "$f" -Destination "$destinationFolder"
            }
        }
        else {
            if($destinationFileName)
            {
                Copy-Item -Path "$($logFiles[0])" -Destination "$destinationFolder\$destinationFileName"
            }
            else {
                Copy-Item -Path "$($logFiles[0])" -Destination "$destinationFolder"
            }
        }
    }
}
function distributeAutomationScript($psNetworkDrive) {
    Copy-Item -Path "$psNetworkDrive`:\osman\db_bak.txt" -Destination ".\osman_db_bak.txt"
}

function test() {
    $conv03 = ($CONV_MACHINES | Where-Object {$_.name -eq "conv03"})
    wh "$($conv03.ip)"
    doWorkRemotely "f:" "$($conv03.ip)" "$($conv03.username)" "$($conv03.password)" "distributeAutomationScript"
}


function fillOutRecordCount () {
    Write-Output "begin";
    Write-Output $recordCountFolder
    # test local folder
    # $PreConvRecordCountContent = Get-Content -Path "D:\Migration\Test\PreConversionRecordCount.txt"
    # $PostConvRecordCountContent = Get-Content -Path "D:\Migration\Test\PostConversionRecordCount.txt"


    $PreConvRecordCountContent = Get-Content -Path "$recordCountFolder\PreConversionRecordCount.txt"
    $PostConvRecordCountContent = Get-Content -Path "$recordCountFolder\PostConversionRecordCount.txt"
    $targetFilePath = "$recordCountFolder\Conversion Record Counts.xlsx"
    
    # Open record count spreadsheet and write value
    $objExcel = New-Object -com Excel.Application
    $objExcel.Visible = $True
    # $targetFilePath = "D:\Migration\Test\Conversion Record Counts.xlsx"
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

function applyCustomScriptsPreConsoleApp($preConsoleAppFolder)
{
    wh "Coping scripts from $preConsoleAppFolder to a fucking lot of folders"
    $destFolders = @(
        @{path = "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\SQLScripts\SiteSpecific\PostConversion"}
        )
    if($SOURCE_SYSTEM -eq $WINBEAT)
    {
        $destFolders += @{path = "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\SQLScripts\Winbeat\FullExtract"}
    }
    if($SOURCE_SYSTEM -eq $IBAIS)
    {
        $destFolders += @{path = "$conv_ledgerFolder\DatabaseConversion.ConsoleApp\SQLScripts\iBais\NonStandard"}
    }

    $scripts = @((Get-ChildItem -Path "$preConsoleAppFolder" -Filter "*.sql" | Sort-Object LastWriteTime -Descending).FullName)
    wh "Dest folders including: "
    foreach ($p in $destFolders)
    {
        wh "`t$($p.path)" "white"
    }
    wh "Pre-console app scripts including: "
    foreach ($s in $scripts)
    {
        wh "`t$s" "white"
    }
    foreach ($p in $destFolders)
    {
        foreach($s in $scripts)
        {
            $fileName = (Split-Path -Leaf $s)
            $destFile = "$($p.path)\$fileName"
            if(Test-Path -Path $destFile)
            {
                wh
                wh "Coping $s into $($p.path)"
                Copy-Item -Path $s -Destination "$($p.path)\$fineName"
            }
        }
    }
}
function RunFullConversion()
{
    wh "*** STARTING CONVERSION WIZARD ***"
    wh
    wh "Step 1: Setting prepare environment" $color_important
    prepareConvEnvironment
    countDown 3
    wh "Step 2: Extract source codes" $color_important
    extract
    countDown 3
    wh "Step 3: Configure app settings / connection strings" $color_important
    config
    countDown 3
    wh "Step 4: Apply configuration to main config files" $color_important
    applyConfig
    countDown 3
    wh "Step 5: Restore ledger database" $color_important
    restoreLedgerDb
    countDown 3
    wh "Step 6: Create insight database" $color_important
    createInsightDb
    countDown 3
    wh "Step 7: Apply pre-console app scripts" $color_important
    applyCustomScriptsPreConsoleApp "$conv_siteSpecificScriptsFolder\preconsoleapp"
    countDown 3
    wh "Step 8: Start console app" $color_important
    Set-Location "$conv_ledgerFolder\DatabaseConversion.ConsoleApp"
    Start-Process -FilePath ".\DatabaseConversion.ConsoleApp.exe"

    Set-Location $executionFolder
    wh "*** DONE ***"
}

function moddy()
{
    createInsightDb
    $scripts = @((Get-ChildItem -Path "D:\sfg-repos\insight_data_conversion\boa-data-conversion\DatabaseConversion.ConsoleApp\SQLScripts\BrokerReady" -Filter "*.sql" | Sort-Object).FullName)
    foreach ($s in $scripts)
    {
        wh "Executing script $s"
        SQLCMD.EXE -i $s -d "ModdyInsight" -E -S '.'
    }
}