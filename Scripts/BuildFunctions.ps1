#
# Powershell script to build
#
# $HeadURL$
# $LastChangedBy$
# $LastChangedDate$
# $LastChangedRevision$
#

Set-Alias svn "C:\Program Files\CollabNet\Subversion Client\svn.exe"

###################################################
### Read Name=Value config file into hash table
###################################################
function ReadConfigFile($filePath)
{
    $found = Test-Path($filePath)
    if ($found -eq $false)
    {
        $msg = "ERROR: Configuration file ", $filePath, " not found"
        Write-Host $msg
        Throw $msg
    }
    
    $nvp = @{}

    $lines = Get-Content $filePath

    foreach ($line in $lines)
    {
        if ($line.length -eq 0)
        { continue }
        
        if ($line.StartsWith("#"))
        { continue }

        $nv = $line.split("=", 2)
        $nv[0] = $nv[0].Trim()
        $nv[1] = $nv[1].Trim()
        
        if ($nvp.ContainsKey($nv[0]))
        {
            $nvp.set_item($nv[0], $nv[1])
        }
        else
        {
            $nvp.add($nv[0], $nv[1])
        }
    }
    
    return $nvp
}

###################################################
# Retrieves a config field value and checks it
###################################################
function GetConfigFieldValue($config, $fieldName)
{
    if (-not $config.ContainsKey($fieldName))
    { Throw ("ERROR: " + $fieldName + " not set in config file.") }

    $value = $config[$fieldName]
   
    if ($value.length -eq 0)
    { Throw ("ERROR: " + $fieldName + " not set in config file.") }
    
    return $value
}

function GetNotRequiredConfigFieldValue($config, $fieldName)
{
    if (-not $config.ContainsKey($fieldName))
    { Throw ("ERROR: " + $fieldName + " not set in config file.") }

    $value = $config[$fieldName]
       
    return $value
}

$buildSettings = ReadConfigFile ".\Config\Build.settings.txt"

$nuget = GetConfigFieldValue $buildSettings "NugetPath" 
$msBuildPath = GetConfigFieldValue $buildSettings "MsBuildPath" 
Set-Alias nuget $nuget
Set-Alias msbuild $msBuildPath
###################################################
### Pause screen
###################################################
function Pause()
{
   Write-Host "Press any key to continue ..."
   $x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
###################################################
### Get code from Bitbucket
###################################################
function GetCodeBitbucket($sourceDir, $gitPath)
{
    Write-Host "GETTING SOURCE CODE INTO $sourceDir" -foreground green

    if (Test-Path $sourceDir) 
    { 
        Remove-Item -Force -Recurse -Path "$sourceDir"
    }
	New-Item -Path "$sourceDir" -Type directory | Out-Null
	git config --global credential.helper wincred
	git clone --depth 1 --branch "master" --single-branch $gitPath $sourceDir

    if ($LastExitCode -ne 0) {  
		Throw "GIT CLONE FAILED" 
	}
}

###################################################
### Compile a solution
###################################################
function CompileSolution($solutionPath, $solutionConfig, $includeProjects = "", $dotNetVersion = "4")
{	
    Write-Host "BUILDING PROJECTS: " $includeProjects -foreground green
    if ($solutionPath -eq $())
    {
        $msg = "ERROR: solution path not specified"
        Write-Host $msg
        PrintUsage
        Throw $msg
    }

    $found = Test-Path($solutionPath)
    if ($found -eq $false)
    {
        $msg = "ERROR: solution path " + $solutionPath + " not found"
        Write-Host $msg
        PrintUsage
        Throw $msg
    }
    if ($solutionConfig -eq $())
    {
        $solutionConfig = "Release"
    }

    # Setup log file
    $solutionFileName = (Get-Item $solutionPath).Name
    $logFileName = "logs\" + $solutionFileName + "_log.txt";
    if (Test-Path($logFileName))
    {
        Remove-Item $logFileName
    }

    Write-Host "COMPILING " $solutionFileName -foreground green
    # Clean up
	switch ($dotNetVersion) 
	{
		default { msbuild "$solutionPath" /t:clean /p:configuration=Release }
		
	}
    if ($LastExitCode -ne 0)
    {  
        Throw "CLEAN " + $solutionFileName + " FAILED"  
    }

	switch ($dotNetVersion) 
	{
		default { msbuild "$solutionPath" /t:clean /p:configuration=Debug }
	}    
    if ($LastExitCode -ne 0)
    {  
        Throw "CLEAN " + $solutionFileName + " FAILED"  
    }
	
    # Build all
	switch ($dotNetVersion) 
	{
		default { 
            if ([string]::IsNullOrEmpty($includeProjects))
            {
                msbuild "$solutionPath" /p:DeployOnBuild=true /p:Configuration="Release" /v:m | Out-Default 
            }
            else {
                msbuild "$solutionPath" /t:$includeProjects /p:DeployOnBuild=true /p:Configuration="Release" /v:m | Out-Default 
            }
		}
	}     
    if ($LastExitCode -ne 0)
    {  
        Throw "Building " + $solutionFileName + " FAILED"  
    }
    Write-Host "Compile Successful" -foreground green
    Write-Host ""
}

function GetBuildProjects($solutionRootDir, [string[]]$excludedProjects)
{
    $includeProjects = ""
    Get-ChildItem $solutionRootDir -Recurse -Include "*.csproj" |
      ForEach-Object {  
          $itemName =[System.IO.Path]::GetFileName($_.FullName)
          if(-not ($excludedProjects -contains $itemName))
          {
            $includeProjects = $includeProjects + $_.BaseName.Replace(".", "_") + ";"
          }
    }
    return $includeProjects.Substring(0,$includeProjects.Length-1)
}

function RestoreNugetPackages($solutionDir)
{
    Write-Host "BEGIN RESTORE NUGET PACKAGES" -foreground green
    # Restore nuget package
    Get-ChildItem $solutionDir -Recurse -Include "packages.config" |
    ForEach-Object {   
    # Nuget restore packages used within all projects
         nuget restore $_.FullName -PackagesDirectory "$solutionDir\packages"
    }
    Write-Host "RESTORE NUGET PACKAGES DONE" -foreground green
}

###################################################
### Tag Project Bitbucket
###################################################
function TagBitbucketProject($sourceDir, $tag, $buildVersion, $buildDate)
{
    if ([string]::IsNullOrEmpty($tag) -or ($tag -eq "p")) 
    {
        do 
        {
            $tag = Read-Host "Tag this build as V$buildVersion-$buildDate ?`n [Y]es [N]o (default is 'Y')"
            if ([string]::IsNullOrEmpty($tag))
            { $tag = "y" }
            $tag = $tag.ToLower()
        } until (($tag -eq "y") -or ($tag -eq "n"))
    }
    
    if ($tag -eq "y") 
    {
        Push-Location -Path $sourceDir
        #Create local tag
        $msg = "'TAG VERSION $buildVersion $buildDate'"
        git tag -m"$msg" "V$buildVersion-$buildDate"
        if ($LastExitCode -ne 0) {  
            Pop-Location
            Throw "TAG FAILED"  
        }
     
        # Push tag to origin
        git push --tags
        if ($LastExitCode -ne 0) {  
            Pop-Location
            Throw "TAG FAILED"  
        }
    
        Pop-Location
    }
}
###################################################
### Web publish a solution
###################################################
function WebPublishSolution($solutionPath, $solutionConfig, $includeProjects , $dotNetVersion = "4")
{
    if ($solutionPath -eq $())
    {
        $msg = "ERROR: solution path not specified"
        Write-Host $msg
        PrintUsage
        Throw $msg
    }

    $found = Test-Path($solutionPath)
    if ($found -eq $false)
    {
        $msg = "ERROR: solution path " + $solutionPath + " not found"
        Write-Host $msg
        PrintUsage
        Throw $msg
    }
    if ($solutionConfig -eq $())
    {
        $solutionConfig = "Release"
    }

    # Setup log file
    $solutionFileName = (Get-Item $solutionPath).Name
    $logFileName = "logs\" + $solutionFileName + "_log.txt";
    if (Test-Path($logFileName))
    {
        Remove-Item $logFileName
    }

    Write-Host "PUBLISHING  $solutionFileName" -foreground green

    # Clean up
    Write-Host "CLEANING RELEASE MODE" -foreground green
	switch ($dotNetVersion) 
	{
		default { msbuild "$solutionPath" /t:clean /p:configuration=Release }
	}
    if ($LastExitCode -ne 0)
    {  
        Throw "CLEAN " + $solutionFileName + " FAILED"  
    }
    Write-Host "CLEANING DEBUG MODE" -foreground green
	switch ($dotNetVersion) 
	{
		default { msbuild "$solutionPath" /t:clean /p:configuration=Release }
	}    
    if ($LastExitCode -ne 0)
    {  
        Throw "CLEAN " + $solutionFileName + " FAILED"  
    }
	Write-Host "PUBLISHING WEB APP TO $buildPublishDir" -foreground green

    # Build all
    switch ($dotNetVersion) 
	{
		default { 
            if ([string]::IsNullOrEmpty($includeProjects))
            {
                msbuild "$solutionPath" /p:DeployOnBuild=true /p:Configuration="$solutionConfig" /v:m | Out-Default 
            }
            else {
                msbuild "$solutionPath" /t:$includeProjects /p:DeployOnBuild=true /p:Configuration="$solutionConfig" /v:m | Out-Default 
            }
		}
	}     
    Write-Host "PUBLISH SUCCESSFULLY" -foreground green
    Write-Host ""
}



