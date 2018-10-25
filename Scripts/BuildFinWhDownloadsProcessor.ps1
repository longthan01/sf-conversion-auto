#
# Powershell script to build the Finance Warehouse Progress Updater
#
# $HeadURL$
# $LastChangedBy$
# $LastChangedDate$
# $LastChangedRevision$
#

###################################################
# Globals
###################################################
$ErrorActionPreference = 'Stop'

$Invocation = (Get-Variable MyInvocation -Scope 0).Value
$scriptPath = Split-Path $Invocation.MyCommand.Path
Set-Location $scriptPath

$solutionDir = "$scriptPath\..\source\"
$rootDistDir = "$scriptPath\..\dist\"

$solutionFilePath = $solutionDir + "\finwh-downloads-processor.sln"
$solutionConfig = "Release"
$versionFilePath = $solutionDir + "\version.txt"
$buildFilePath = $solutionDir + "\build.txt"
$exeDir = $solutionDir + "\app\bin\" + $solutionConfig
$sqlDir = $solutionDir + "\sql"

###################################################
# Functions
###################################################

# 

# Reading our common functions
. .\BuildFunctions.ps1

function PrintUsage()
{
    Write-Host "USAGE"
    Write-Host ""
    Write-Host "BuildFinWhDownloadsProcessor [/branch:{name}] [/tag:{Y|N|P}]"
    Write-Host ""
    Write-Host "[/branch] = optional git branch. defaults to 'master'."
    Write-Host "[/tag] = optional flag to tag build in Git. Enter (Y)es, (N)o or (P)rompt. Defaults to 'P'."
    Write-Host ""
    Write-Host ""
}

###################################################
# Main
###################################################
$branch = "master"
$doTag = "p"

foreach ($arg in $args)
{
	if ([string]::IsNullOrEmpty($arg))
	{ continue }
	elseif ($arg -eq "-?" -or $arg -eq "/?")
	{
		PrintUsage
		exit
	}
	elseif ($arg.StartsWith("/branch:"))
	{	    
		$branch = "branches/" + $arg.SubString(8)
	}
	elseif ($arg.StartsWith("/tag:"))
	{
		$doTag = $arg.ToLower().SubString(5)
		if ([string]::IsNullOrEmpty($doTag) -or (($doTag -ne "y") -and ($doTag -ne "n") -and ($doTag -ne "p")))
		{
			Throw "Invalid tag option: $doTag. Enter Y|N|P."  
		}
	}
	else
	{
		Throw "Unknown argument: $arg"  
	}
}

$gitPath = "git@bitbucket.org:steadfasttech/sfg-finwh-downloads-processor.git"


#
# Compile
#
GetCodeBitbucket $solutionDir $gitPath
RestoreNugetPackages $solutionDir
$compileProjects = GetBuildProjects $solutionDir @("test.csproj")
CompileSolution $solutionFilePath $solutionConfig $compileProjects

#
# Get version
#
$versionConfig = ReadConfigFile($versionFilePath)
$versionNumber = GetConfigFieldValue $versionConfig "Version"

Write-Host "VERSION NUMBER: $versionNumber" -foreground green

#
# Get build details
#
$gitUrlTokens = $gitPath.split('/')
$repoName = $gitUrlTokens[$gitUrlTokens.length-1].replace('.git','')
$buildDate =  Get-Date -format "yyyyMMdd-HHmmss"

$newBuildFileLines = "BuildVersion=" +$versionNumber + "`r`nBuildDate=" + $buildDate
$newBuildFileLines | Out-File $buildFilePath -force

#
# Copy files to distribution area
#
$distDir = $rootDistDir + "\" + $repoName + '-' + $versionNumber + "-" + $buildDate
if (Test-Path($distDir))
{ 
    Write-Host "DELETING $distDir" -foreground yellow
    Remove-Item $distDir -recurse -force 
}
Copy-Item $exeDir -destination $distDir -recurse
if (Test-Path $sqlDir)
{
	Copy-Item $sqlDir -destination $distDir -recurse
}

Copy-Item $versionFilePath -destination $distDir
Copy-Item $buildFilePath -destination $distDir

$appConfigDownloads = $solutionDir + "\app\app.config.template"
Copy-Item $appConfigDownloads -destination $distDir

#
# Delete app.config because we will merge with app.config.template on install
#
$appConfig = $distDir + "\finwh-downloads-processor.exe.config"
if (Test-Path($appConfig))
{ 
	Remove-Item $appConfig -force
}

Write-Host "DISTRIBUTED FILES TO $distDir" -foreground green

#
# Tag it in Git
#

TagBitbucketProject $solutionDir $doTag $versionNumber $buildDate
#
# Write out command line to run for this build
#
$installCommandLineFile = $rootDistDir + "\..\CommandLine.txt"
$installCommandLine = ".\InstallFinWhDownloadsProcessor.ps1 " + $repoName + '-' + $versionNumber + "-" + $buildDate + "`r`n`r`n"
$installCommandLine | Out-File $installCommandLineFile -append


#
# Finished
#
Write-Host ""
Write-Host "...FINISHED..." -foreground green
Write-Host ""
