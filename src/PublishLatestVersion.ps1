################################################################################
# MIT License
#
# Copyright (c) 2021 Microsoft and Contributors
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
# Filename: PublishLatestVersion.ps1
# Version: 1.0.0.0
# Description: Script to publish new MSI installers to the MsiOverride share
# Owner: Christopher Tart <chtart@microsoft.com>
#################################################################################

Param(
    [Parameter(Mandatory=$false)]
    [Switch] $Help = $false,
    [Parameter(Mandatory=$true)]
    [string] $BaseShare = "",
    [Parameter(Mandatory=$false)]
    [ValidateSet('Preview', 'General', 'GCCGeneral', 'GCCHGeneral', 'DODGeneral')]
    [string] $Ring = "General",
    [Parameter(Mandatory=$false)]
    [string] $OverrideVersion = ""
)

$RingNames = @{
    "Preview" = "ring3_6"; 
    "General" = "general";
    "GCCGeneral" = "general_gcc";
    "GCCHGeneral" = "gcchigh-general";
    "DODGeneral" = "dod-general";
}

$RingFQDNs = @{
    "Preview" = "teams.microsoft.com"; 
    "General" = "teams.microsoft.com";
    "GCCGeneral" = "teams.microsoft.com";
    "GCCHGeneral" = "gov.teams.microsoft.us";
    "DODGeneral" = "dod.teams.microsoft.us";
}

$ScriptName  = "Microsoft Teams MsiOverride Publisher"
$Version     = "1.0.0.0"

function ShowHelp
{
    Write-Output ""
    Write-Output "$ScriptName"
    Write-Output "Version $Version"
    Write-Output ""
    Write-Output "Use this script check for and publish a new Teams MSI installer for use with the"
    Write-Output "MsiOverride update process"
    Write-Output ""
    Write-Output "PublishLatestVersion.ps1 [-Help]"
    Write-Output ""
    Write-Output "PublishLatestVersion.ps1 [-BaseShare <path>] [-PreviewRing] [-OverrideVersion <version>]"
    Write-Output ""
    Write-Output "  -Help       : Displays this help message."
    Write-Output ""
    Write-Output "  -BaseShare: Provides the share location where the MSIs are published to."
    Write-Output ""
    Write-Output "  -PreviewRing: Retrieves the downloads for preview ring (Ring3), instead of general ring."
    Write-Output ""
    Write-Output "  -OverrideVersion: Downloads the designated version and sets it as the target version"
    Write-Output "                    Must be given in a format which is similar to 1.2.00.34567"
    Write-Output ""
    Exit
}

function GetDownloadUrl($version, $bitness, $fileName)
{
    $url = $updateCheckUrl -f $version,$bitness,$RingNames[$Ring],$RingFQDNs[$Ring]

    Write-Host "Sending request to $url"
    $updateCheckResponse = Invoke-WebRequest -Uri $url -UseBasicParsing
    $updateCheckJson = $updateCheckResponse | ConvertFrom-Json

    if($updateCheckJson.isUpdateAvailable)
    {
        $downloadPath = $updateCheckJson.releasesPath.Replace("RELEASES", $fileName)
        Write-Host "Returning $downloadPath"
        return $downloadPath
    }

    Write-Host "Returning null"
    return ""
}

function CreateFolder($path)
{
    New-Item -ItemType Directory -Path $path -ErrorAction Continue | Out-Null
    if(-Not (Test-Path $path))
    {
        Write-Host "Unable to create $path" -ForegroundColor Red
        Exit -1
    }
}

function DeleteFolder($path)
{
    if(Test-Path $path)
    {
        Remove-Item -Path $path -Recurse -Force -ErrorAction Continue | Out-Null
        if(Test-Path $path)
        {
            Write-Host "Unable to delete $path" -ForegroundColor Red
            Exit -1
        }
    }
}

function DeleteFile($path)
{
    if(Test-Path $path)
    {
        Remove-Item -Path $path -Force -ErrorAction Continue | Out-Null
        if(Test-Path $path)
        {
            Write-Host "Unable to delete $path" -ForegroundColor Red
            Exit -1
        }
    }
}

# Constants
$updateCheckUrl = "https://{3}/desktopclient/update/{0}/windows/{1}?ring={2}"
$downloadFormat32 = "https://statics.teams.microsoft.com/production-windows/{0}/Teams_windows.msi"
$downloadFormat64 = "https://statics.teams.microsoft.com/production-windows-x64/{0}/Teams_windows_x64.msi"
$versionRegex = "(?<version>\d+\.\d+\.\d+\.\d+)"
$fileName32 = "Teams_windows.msi"
$fileName64 = "Teams_windows_x64.msi"

# Main Script

if($Help) { ShowHelp }

if(-Not (Test-Path $BaseShare))
{
    Write-Host "Specified BaseShare path does not exist.  Please create it first and ensure it has the proper permissions." -ForegroundColor Red
    Exit -1
}

if($OverrideVersion -ne "" -and !($OverrideVersion -match $versionRegex))
{
    Write-Host "Invalid version format provided.  Please ensure it follows a format similar to 1.2.00.34567" -ForegroundColor Red
    Exit -1
}

# Add TLS 1.2 for older OSs
if (([Net.ServicePointManager]::SecurityProtocol -ne 'SystemDefault') -and 
    !(([Net.ServicePointManager]::SecurityProtocol -band 'Tls12') -eq 'Tls12'))
{
    Write-Host "Adding TLS 1.2 protocol"
    [Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
}

Write-Host "Checking for a new Teams version in ring $Ring ..." -ForegroundColor Green

# Get the current version in use
$currentVersion = "1.3.00.0000"
$versionFile = Join-Path $BaseShare "Version.txt"

if($OverrideVersion -eq "")
{
    $fileVersion = Get-Content $versionFile -ErrorAction SilentlyContinue
    if($fileVersion -match $versionRegex)
    {
        $currentVersion = $Matches.version
    }
    Write-Host "Current version $currentVersion"

    # Try to get new download URLs for both 32 and 64 bit MSIs, returns null if none is available
    $downloadPath32 = GetDownloadUrl $currentVersion "x32" $fileName32
    $downloadPath64 = GetDownloadUrl $currentVersion "x64" $fileName64
}
else
{
    $downloadPath32 = $downloadFormat32 -f $OverrideVersion
    $downloadPath64 = $downloadFormat64 -f $OverrideVersion
}


# Extract new version number from URL
$newVersion = ""
if($downloadPath32 -match $versionRegex)
{
    $newVersion = $Matches.version
    Write-Host "New version $newVersion"
}

# If we have a new version number and both download paths, proceed to get them
if($newVersion -ne "" -and $downloadPath32 -ne "" -and $downloadPath64 -ne "")
{
    # Create a new version folder on the share
    $newFolder = Join-Path $BaseShare $newVersion
    Write-Host "Creating folder $newFolder"
    DeleteFolder($newFolder)
    CreateFolder($newFolder)

    $oldProgressPreference = $ProgressPreference
    $ProgressPreference = 'SilentlyContinue'

    # Download both MSIs to the new version folder
    Write-Host "Downloading $downloadPath32"
    $localPath32 = Join-Path $env:TEMP $fileName32
    DeleteFile $localPath32
    Invoke-WebRequest -Uri $downloadPath32 -OutFile $localPath32
    Write-Host "Download complete.  Moving it to $newFolder"
    Move-Item -Path $localPath32 -Destination $newFolder
    
    Write-Host "Downloading $downloadPath64"
    $localPath64 = Join-Path $env:TEMP $fileName64
    DeleteFile $localPath64
    Invoke-WebRequest -Uri $downloadPath64 -OutFile $localPath64
    Write-Host "Download complete.  Moving it to $newFolder"
    Move-Item -Path $localPath64 -Destination $newFolder

    $ProgressPreference = $oldProgressPreference

    $remotePath32 = Join-Path $newFolder $fileName32
    $remotePath64 = Join-Path $newFolder $fileName64
    if((Test-Path $remotePath32) -and (Test-Path $remotePath64))
    {
        # Update the current version
        Write-Host "Updating Version.txt to $newVersion"
        $newVersion | Out-File $versionFile
        Write-Host "New version published!" -ForegroundColor Green
    }
    else
    {
        Write-Host "One or both MSIs were not found on the base share, something went wrong!" -ForegroundColor Red
    }
}
else
{
    Write-Host "New version is not available!" -ForegroundColor Green
}
