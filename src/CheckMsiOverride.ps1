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
# Filename: CheckMsiOverride.ps1
# Version: 1.0.0.0
# Description: Script to check for and applies Teams msiOverride updates
# Owner: Christopher Tart <chtart@microsoft.com>
#################################################################################

#Requires -RunAsAdministrator

Param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('Share','CDN','Package')]
    [string] $Type = "Share",
    [Parameter(Mandatory=$false)]
    [ValidateSet('Preview', 'General', 'GCCGeneral', 'GCCHGeneral', 'DODGeneral')]
    [string] $Ring = "General",
    [Parameter(Mandatory=$false)]
    [string] $BaseShare = "",
    [Parameter(Mandatory=$false)]
    [string] $OverrideVersion = "",
    [Parameter(Mandatory=$false)]
    [string] $MsiFileName = "",
    [Parameter(Mandatory=$false)]
    [Switch] $AllowInstallOvertopExisting = $false,
    [Parameter(Mandatory=$false)]
    [Switch] $OverwritePolicyKey = $false,
    [Parameter(Mandatory=$false)]
    [Switch] $FixRunKey = $false,
    [Parameter(Mandatory=$false)]
    [Switch] $Uninstall32Bit = $false,
    [Parameter(Mandatory=$false)]
    [Switch] $UninstallAll = $false,
    [Parameter(Mandatory=$false)]
    [Switch] $UninstallAllIfBothPresent = $false
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

$ScriptName  = "Microsoft Teams MsiOverride Checker"
$Version     = "1.0.0.0"

# Trace functions
function InitTracing([string]$traceName, [string]$tracePath = $env:TEMP)
{
    $script:TracePath = Join-Path $tracePath $traceName
    WriteTrace("")
    WriteTrace("Start Trace $(Get-Date)")
}

function WriteTrace([string]$line, [string]$function = "")
{
    $output = $line
    if($function -ne "")
    {
        $output = "[$function] " + $output
    }
    Write-Verbose $output
    $output | Out-File $script:TracePath -Append
}

function WriteInfo([string]$line, [string]$function = "")
{
    $output = $line
    if($function -ne "")
    {
        $output = "[$function] " + $output
    }
    Write-Host $output
    $output | Out-File $script:TracePath -Append
}

function WriteWarning([string]$line)
{
    Write-Host $line -ForegroundColor DarkYellow
    $line | Out-File $script:TracePath -Append
}

function WriteError([string]$line)
{
    Write-Host $line  -ForegroundColor Red
    $line | Out-File $script:TracePath -Append
}

function WriteSuccess([string]$line)
{
    Write-Host $line  -ForegroundColor Green
    $line | Out-File $script:TracePath -Append
}

# Removes temp folder
function Cleanup
{
    WriteTrace "Removing temp folder $TempPath"
    Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

# Runs cleanup and exits
function CleanExit($code = 0)
{
    Cleanup
    WriteTrace("End Trace $(Get-Date)")
    Exit $code
}

function ErrorExit($line, $code)
{
    WriteError($line)
    Write-EventLog -LogName Application -Source $EventLogSource -Category 0 -EntryType Error -EventId ([Math]::Abs($code)) -Message $line
    CleanExit($code)
}

function IsRunningUnderSystem
{
    if(($env:COMPUTERNAME + "$") -eq $env:USERNAME)
    {
        return $true
    }
    return $false
}

function GetFileVersionString($Path)
{
    if (Test-Path $Path)
    {
        $item = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
        if ($item)
        {
            return $item.FileVersion
        }
    }
    return ""
}

function HasReg($Path, $Name)
{
    if (Test-Path $Path)
    {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        if ($item -ne $null)
        {
            return $true
        }
    }
    return $false
}

function GetReg($Path, $Name, $DefaultValue)
{
    if (HasReg -Path $Path -Name $Name)
    {
        $item = Get-ItemProperty -Path $Path -Name $Name
        return $item.$Name
    }
    return $DefaultValue
}

function SetDwordReg($Path, $Name, $Value)
{
    if (!(Test-Path $Path))
    {
        New-Item -Path $Path -Force | Out-Null
    }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWORD
}

function SetExpandStringReg($Path, $Name, $Value)
{
    if (!(Test-Path $Path))
    {
        New-Item -Path $Path -Force | Out-Null
    }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type ExpandString
}

function GetInstallerVersion
{
    return (GetFileVersionString -Path (GetInstallerPath))
}

function GetInstallerPath
{
    if($([Environment]::Is64BitOperatingSystem))
    {
        return (${env:ProgramFiles(x86)} + "\Teams Installer\Teams.exe")
    }
    else
    {
        return ($env:ProgramFiles + "\Teams Installer\Teams.exe")
    }
}

function GetTargetVersion
{
    $versionFile = Join-Path $BaseShare "Version.txt"
    $fileVersion = Get-Content $versionFile -ErrorAction SilentlyContinue
    return (VerifyVersion($fileVersion))
}

function VerifyVersion($Version)
{
    if($Version -match $versionRegex)
    {
        return $Matches.version
    }
    return $null
}

function GetUninstallKey
{
    $UninstallReg1 = Get-ChildItem -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue  | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' }
    $UninstallReg2 = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' }

    WriteTrace("UninstallReg1: $($UninstallReg1.PSChildName)")
    WriteTrace("UninstallReg2: $($UninstallReg2.PSChildName)")

    if($UninstallReg1) { return $UninstallReg1 }
    elseif($UninstallReg2) { return $UninstallReg2 }
    return $null
}

function GetProductsKey
{
    $ProductsRegLM = Get-ChildItem -Path HKLM:\SOFTWARE\Classes\Installer\Products -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' } # ALLUSERS Install
    $ProductsRegCU = Get-ChildItem -Path HKCU:\SOFTWARE\Microsoft\Installer\Products -ErrorAction SilentlyContinue | Get-ItemProperty | Where-Object { $_ -match 'Teams Machine-Wide Installer' } # Local User Install

    WriteTrace("ProductsRegLM: $($ProductsRegLM.PSChildName)")
    WriteTrace("ProductsRegCU: $($ProductsRegCU.PSChildName)")

    $result = @();
    if($ProductsRegLM) { $result += $ProductsRegLM }
    if($ProductsRegCU) { $result += $ProductsRegCU }
    return $result
}

function Has32BitProductKey($productKeys)
{
    foreach($key in $productKeys)
    {
        if($key.PSChildName -eq $MsiProduct32Guid)
        {
            return $true
        }
    }
    return $false
}

function Has64BitProductKey($productKeys)
{
    foreach($key in $productKeys)
    {
        if($key.PSChildName -eq $MsiProduct64Guid)
        {
            return $true
        }
    }
    return $false
}

function GetPackageKey()
{
    [array]$msiKeys = GetProductsKey
    if($msiKeys.count -eq 1)
    {
        $msiKey = $msiKeys[0]
        $msiPkgReg = (Get-ChildItem -Path $msiKey.PSPath -Recurse | Get-ItemProperty | Where-Object { $_ -match 'PackageName' })

        if ($msiPkgReg.PackageName)
        {
            WriteTrace("PackageName: $($msiPkgReg.PackageName)")
            return $msiPkgReg
        }
    }
    return $null
}

function GetInstallBitnessFromUninstall()
{
    $uninstallReg = GetUninstallKey
    if($uninstallReg)
    {
        if ($uninstallReg.PSPath | Select-String -Pattern $MsiPkg64Guid)
        {
            return "x64"
        }
        elseif ($uninstallReg.PSPath | Select-String -Pattern $MsiPkg32Guid)
        {
            return "x86"
        }
    }
    return $null
}

function GetInstallBitnessFromSource()
{
    $msiPkgReg = GetPackageKey
    if($msiPkgReg)
    {
        WriteTrace("LastUsedSource: $($msiPkgReg.LastUsedSource)")
        if ($msiPkgReg.LastUsedSource | Select-String -Pattern ${env:ProgramFiles(x86)})
        {
            return "x86"
        }
        elseif ($msiPkgReg.LastUsedSource | Select-String -Pattern $env:ProgramFiles)
        {
            if($([Environment]::Is64BitOperatingSystem))
            {
                return "x64"
            }
            else
            {
                return "x86"
            }
        }
    }
    return $null
}

function GetInstallBitnessForOS()
{
    if($([Environment]::Is64BitOperatingSystem))
    {
        return "x64"
    }
    else
    {
        return "x86"
    }
}

function GetInstallBitness([ref]$outMode, [ref]$outFileName)
{
    $installBitness = GetInstallBitnessFromUninstall
    $packageKey = GetPackageKey
    # Determine the install bitness and mode
    if($installBitness)
    {
        # Uninstall key existed and we matched to known GUID
        if($packageKey)
        {
            # Update Scenario, Package key existed (meaning MSI was installed by this user, or as ALLUSERS).
            $mode = "update"
        }
        else
        {
            # Install Scenario, Package key did not exist (meaning MSI is installed, but not by this user and not as ALLUSERS).
            $mode = "installovertop"
        }
    }
    else
    {
        # Uninstall key did not exist or we did not match a known GUID
        if($packageKey)
        {
            # Update Scenario, we do have a package key, so we must not have matched a known GUID, so try to read LastUsedSource path (Office installation scenario).
            $mode = "update"
            $installBitness = GetInstallBitnessFromSource
            if(-not $installBitness)
            {
                # Fall back to OS bitness as a last resort.
                $installBitness = GetInstallBitnessForOS
            }
        }
        else
        {
            # Install Scenario, Neither Uninstall key or Package key existed, so it will be a fresh install
            $mode = "install"
            $installBitness = GetInstallBitnessForOS
        }
    }

    $outMode.Value = $mode
    $outFileName.Value = $packageKey.PackageName

    return $installBitness
}

function DeleteFile($path)
{
    if(Test-Path $path)
    {
        Remove-Item -Path $path -Force | Out-Null
        if(Test-Path $path)
        {
            Write-Host "Unable to delete $path" -ForegroundColor Red
            ErrorExit "Failed to delete existing file $path" -8
        }
    }
}

function SetParametersWithCDN([ref]$outVersion, [ref]$outPath)
{
    WriteInfo "Using CDN to check for an update and aquire the new MSI..."
    $updateCheckUrl = "https://{3}/desktopclient/update/{0}/windows/{1}?ring={2}"
    $downloadFormat32 = "https://statics.teams.microsoft.com/production-windows/{0}/Teams_windows.msi"
    $downloadFormat64 = "https://statics.teams.microsoft.com/production-windows-x64/{0}/Teams_windows_x64.msi"
    $bitness = $installBitness.Replace("x86", "x32")

    # Add TLS 1.2 for older OSs
    if (([Net.ServicePointManager]::SecurityProtocol -ne 'SystemDefault') -and 
        !(([Net.ServicePointManager]::SecurityProtocol -band 'Tls12') -eq 'Tls12'))
    {
        WriteTrace "Adding TLS 1.2 protocol"
        [Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
    }

    $downloadPath = ""
    $fileName = ""
    if($bitness -eq "x32")
    {
        $fileName = $FileName32
    }
    else
    {
        $fileName = $FileName64
    }
    if($outVersion.Value -eq "")
    {
        $url = $updateCheckUrl -f $currentVersion,$bitness,$RingNames[$Ring],$RingFQDNs[$Ring]

        WriteInfo "Sending request to $url"
        $updateCheckResponse = Invoke-WebRequest -Uri $url -UseBasicParsing
        $updateCheckJson = $updateCheckResponse | ConvertFrom-Json

        if($updateCheckJson.isUpdateAvailable)
        {
            $downloadPath = $updateCheckJson.releasesPath.Replace("RELEASES", $fileName)
        }
        else
        {
            $outVersion.Value = $currentVersion
            return
        }
    }
    else
    {
        if($bitness -eq "x32")
        {
            $downloadPath = $downloadFormat32 -f $outVersion.Value
        }
        else
        {
            $downloadPath = $downloadFormat64 -f $outVersion.Value
        }
    }
    WriteInfo "Download path: $downloadPath"

    # Extract new version number from URL
    $newVersion = ""
    if($downloadPath -match $versionRegex)
    {
        $newVersion = $Matches.version
        WriteInfo "New version $newVersion"
    }

    # If we have a new version number and the download path, proceed
    if($newVersion -ne "" -and $downloadPath -ne "")
    {
        $localPath = Join-Path $TempPath "CDN"
        New-Item -ItemType Directory -Path $localPath | Out-Null
        $localPath = Join-Path $localPath $fileName
        WriteInfo "Downloading $downloadPath"
        DeleteFile $localPath
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $downloadPath -OutFile $localPath
        $ProgressPreference = $oldProgressPreference
        WriteInfo "Download complete."
        if(Test-Path $localPath)
        {
            WriteInfo "Successfully downloaded new installer to $localPath"
            $outVersion.Value = $newVersion
            $outPath.Value = $localPath
            return
        }
    }
    ErrorExit "Failed to check for or retrieve the update from the CDN!" -9
}

function SetParametersAsPackage([ref]$outVersion, [ref]$outPath)
{
    WriteInfo "Using working directory to aquire the new MSI..."
    if($outVersion.Value -eq "")
    {
        ErrorExit "Target version should already be provided by OverrideVersion parameter"
    }

    $workingDirectory = Get-Location

    WriteInfo "Working Directory: $workingDirectory"

    # Select MSI based on the bitness
    if ($installBitness -eq "x86") 
    {
        WriteInfo "Using 32-bit MSI from working directory"
        $fromMsi = Join-Path $workingDirectory $FileName32 # x86 MSI
    }
    else
    {
        WriteInfo "Using 64-bit MSI from working directory"
        $fromMsi = Join-Path $workingDirectory $FileName64 # x64 MSI
    }
    $outPath.Value = $fromMsi
}

function SetParametersWithShare([ref]$outVersion, [ref]$outPath)
{
    WriteInfo "Using the BaseShare check for an update and aquire the new MSI..."
    if($outVersion.Value -eq "")
    {
        # Get the target Teams Machine Installer version from the share
        $targetVersion = GetTargetVersion
        $outVersion.Value = $targetVersion
    }

    # Select MSI based on the bitness
    if ($installBitness -eq "x86") 
    {
        WriteInfo "Using 32-bit MSI from BaseShare"
        $fromMsi = "$BaseShare\$targetVersion\$FileName32" # x86 MSI
    }
    else
    {
        WriteInfo "Using 64-bit MSI from BaseShare"
        $fromMsi = "$BaseShare\$targetVersion\$FileName64" # x64 MSI
    }
    $outPath.Value = $fromMsi
}

function GetMsiExecFlags()
{
    $msiExecFlags = ""
    # Set msiExec flags based on our mode
    if ($mode -eq "install")
    {
        WriteInfo "This will be an install"
        $msiExecFlags = "/i" # new install flag
    }
    elseif ($mode -eq "update")
    {
        WriteInfo "This will be an override update"
        $msiExecFlags = "/fav" # override flag
    }
    elseif ($mode -eq "installovertop")
    {
        if($AllowInstallOvertopExisting)
        {
            WriteInfo "This will be an install overtop an existing install"
            $msiExecFlags = "/i" # new install flag
        }
        else
        {
            ErrorExit "ERROR: Existing Teams Machine-Wide Installer is present but it was not installed by the current user or as an ALLUSERS=1 install" -4
        }
    }
    else 
    {
        ErrorExit "UNEXPECTED ERROR! Unknown mode" -5
    }
    return $msiExecFlags
}

function CheckPolicyKey()
{
    # Set AllowMsiOverride key if needed
    $AllowMsiExists = (HasReg -Path $AllowMsiRegPath -Name $AllowMsiRegName)
    if ((-not $AllowMsiExists) -or $OverwritePolicyKey)
    {
        WriteInfo "The policy key AllowMsiOverride is not set, setting $AllowMsiRegPath\$AllowMsiRegName to 1..."
        SetDwordReg -Path $AllowMsiRegPath -Name $AllowMsiRegName -Value 1 | Out-Null
    }
    $AllowMsiValue = !!(GetReg -Path $AllowMsiRegPath -Name $AllowMsiRegName -DefaultValue 0)
    WriteInfo "AllowMsiOverride policy is set to $AllowMsiValue"

    if(-not $AllowMsiValue)
    {
        ErrorExit "ERROR: AllowMsiOverride is not enabled by policy!" -1
    }
}

function CheckParameters()
{
    if( $Type -eq "Share" -and $BaseShare -eq "" )
    {
        ErrorExit "ERROR: BaseShare must be provided"
    }
    if( $Type -ne "Share" -and $BaseShare -ne "" )
    {
        ErrorExit "ERROR: BaseShare should only be provided with Share type"
    }
    if( $Type -eq "Package" -and $OverrideVersion -eq "")
    {
        ErrorExit "ERROR: You must provide an OverrideVersion with Package type"
    }
}

# ----- Constants -----

$versionRegex = "(?<version>\d+\.\d+\.\d+\.\d+)"

$AllowMsiRegPath = "HKLM:\Software\Policies\Microsoft\Office\16.0\Teams"
$AllowMsiRegName = "AllowMsiOverride"

$RunKeyPath32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
$RunKeyPath64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"

$MsiPkg32Guid = "{39AF0813-FA7B-4860-ADBE-93B9B214B914}"
$MsiPkg64Guid = "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}"

$MsiProduct32Guid = "3180FA93B7AF0684DAEB399B2B419B41"
$MsiProduct64Guid = "AAB6F137689A4A549863C7A3AAAA67B0"

$FileName32 = "Teams_windows.msi"
$FileName64 = "Teams_windows_x64.msi"

$TempPath     = Join-Path $env:TEMP "TeamsMsiOverrideCheck"

$EventLogSource = "TeamsMsiOverride"

#----- Main Script -----

# Set the default error action preference
$ErrorActionPreference = "Continue"

InitTracing("TeamsMsiOverrideTrace.txt")

WriteTrace("Script Version $Version")
WriteTrace("Parameters Type: $Type, Ring: $Ring, BaseShare: $BaseShare, OverrideVersion: $OverrideVersion, MsiFileName: $MsiFileName, AllowInstallOvertopExisting: $AllowInstallOvertopExisting, OverwritePolicyKey: $OverwritePolicyKey, FixRunKey: $FixRunKey")
WriteTrace("Environment IsSystemAccount: $(IsRunningUnderSystem), IsOS64Bit: $([Environment]::Is64BitOperatingSystem)")

# Create event log source
try { New-EventLog -LogName Application -Source $EventLogSource -ErrorAction SilentlyContinue }
catch { }

# Delete the temp directory
Cleanup

# Validate parameters passed in
CheckParameters

# Check and set AllowMsiOverride key
CheckPolicyKey

$Remove32 = $false
$Remove64 = $false

[array]$productsKeys = GetProductsKey
if($UninstallAll)
{
    WriteInfo "Removing all existing versions of the machine-wide installer."
    $Remove32 = $true
    $Remove64 = $true
}
else
{
    # Check if we have both 32 bit and 64 bit versions of the MSI installed.  This will be an issue.
    if((Has32BitProductKey $productsKeys) -and (Has64BitProductKey $productsKeys))
    {
        # If switch is passed, uninstall the 32 bit MSI before we perform the upgrade on 64 bit.
        if($Uninstall32Bit)
        {
            WriteInfo "Removing 32-bit version of the machine-wide installer."
            $Remove32 = $true
        }
        elseif($UninstallAllIfBothPresent)
        {
            WriteInfo "Removing all existing versions of the machine-wide installer."
            $Remove32 = $true
            $Remove64 = $true
        }
        else
        {
            ErrorExit "It appears you have both 32 and 64 bit versions of the machine-wide installer present.  Please uninstall them and reinstall the correct one, or use the Uninstall32Bit switch to attempt to uninstall the 32 bit version." -16
        }
    }
}

if($Remove32 -and (Has32BitProductKey $productsKeys))
{
    $msiExecUninstallArgs = "/X $MsiPkg32Guid /quiet /l*v $env:TEMP\msiOverrideCheck_msiexecUninstall32.log"

    WriteInfo "About to uninstall 32-bit MSI using this msiexec command:"
    WriteInfo " msiexec.exe $msiExecUninstallArgs"

    $res = Start-Process "msiexec.exe" -ArgumentList $msiExecUninstallArgs -Wait -PassThru -WindowStyle Hidden
    if ($res.ExitCode -eq 0)
    {
        WriteInfo "MsiExec completed successfully."
    }
    else
    {
        ErrorExit "ERROR: MsiExec failed with exit code $($res.ExitCode)" $res.ExitCode
    }
}

if($Remove64 -and (Has64BitProductKey $productsKeys))
{
    $msiExecUninstallArgs = "/X $MsiPkg64Guid /quiet /l*v $env:TEMP\msiOverrideCheck_msiexecUninstall64.log"

    WriteInfo "About to uninstall 64-bit MSI using this msiexec command:"
    WriteInfo " msiexec.exe $msiExecUninstallArgs"

    $res = Start-Process "msiexec.exe" -ArgumentList $msiExecUninstallArgs -Wait -PassThru -WindowStyle Hidden
    if ($res.ExitCode -eq 0)
    {
        WriteInfo "MsiExec completed successfully."
    }
    else
    {
        ErrorExit "ERROR: MsiExec failed with exit code $($res.ExitCode)" $res.ExitCode
    }
}

# Get the existing Teams Machine Installer version
$currentVersion = GetInstallerVersion
if($currentVersion)
{
    WriteInfo "Current Teams Machine-Wide Installer version is $currentVersion"
}
else
{
    WriteInfo "Teams Machine-Wide Installer was not found."
    $currentVersion = "1.3.00.00000"
}

$fromMsi = ""
$mode = ""
$packageFileName = ""
$installBitness = GetInstallBitness ([ref]$mode) ([ref]$packageFileName)

if($packageFileName -is [array])
{
    ErrorExit "Two or more package file names were found, indicating the machine-wide installer may be installed multiple times! Unable to continue." -17
}

$targetVersion = ""
if($OverrideVersion -ne "")
{
    $targetVersion = VerifyVersion $OverrideVersion
    if($targetVersion -eq $null)
    {
        ErrorExit "Specified OverrideVersion is not the correct format.  Please ensure it follows a format similar to 1.2.00.34567"  -10
    }

    if($currentVersion -eq $targetVersion)
    {
        WriteSuccess "Version specified in OverrideVersion is already installed!"
        CleanExit
    }
}

# Set the parameters either using CDN or file share
if($Type -eq "CDN")
{
    SetParametersWithCDN ([ref]$targetVersion) ([ref]$fromMsi)
}
elseif($Type -eq "Package")
{
    SetParametersAsPackage ([ref]$targetVersion) ([ref]$fromMsi)
}
else
{
    SetParametersWithShare([ref]$targetVersion) ([ref]$fromMsi)
}

# Confirm we have the target version
if($targetVersion)
{
    WriteInfo "Target Teams Machine-Wide Installer version is $targetVersion"
}
else
{
    ErrorExit "ERROR: TargetVersion is invalid!" -2
}

# Confirm we don't already have the target version installed
if($currentVersion -eq $targetVersion)
{
    WriteSuccess "Target version already installed!"
    CleanExit
}

# Get our MSIExec flags
$msiExecFlags = GetMsiExecFlags

# Check that we can reach the MSI file
if (-not (Test-Path $fromMsi))
{
    ErrorExit "ERROR: Unable to access the MSI at $fromMsi" -6
}

# Get the new MSI file name (must match the original for an in place repair operation)
if($MsiFileName -ne "")
{
    $msiName = $MsiFileName
}
else
{
    $msiName = $packageFileName
}

if (-not $msiName)
{
    # If this is a new install, or we don't know the MSI name, use the original MSI name
    $msiName = Split-Path $fromMsi -Leaf
}

# Rename (for CDN based) or copy from the share with the new name (for share based)
if($Type -eq "CDN")
{
    WriteInfo "Renaming $fromMsi to $msiName..."
    $toMsi = (Rename-Item -Path $fromMsi -NewName $msiName -PassThru).FullName
}
else
{
    # Copy MSI to our temp folder
    $toMsi = Join-Path $TempPath $msiName
    WriteInfo "Copying $fromMsi to $toMsi..."
    New-Item -ItemType File -Path $toMsi -Force | Out-Null
    Copy-Item -Path $fromMsi -Destination $toMsi | Out-Null
}

#Construct our full MsiExec arg statement
$msiExecArgs = "$msiExecFlags `"$toMsi`" /quiet ALLUSERS=1 /l*v $env:TEMP\msiOverrideCheck_msiexec.log"

# Output our action
WriteInfo "About to perform deployment using this msiexec command:"
WriteInfo " msiexec.exe $msiExecArgs"

# Do the install or upgrade
$res = Start-Process "msiexec.exe" -ArgumentList $msiExecArgs -Wait -PassThru -WindowStyle Hidden
if ($res.ExitCode -eq 0)
{
    WriteInfo "MsiExec completed successfully."
}
else
{
    ErrorExit "ERROR: MsiExec failed with exit code $($res.ExitCode)" $res.ExitCode
}

# Fixup the HKLM Run key if option is set
if($FixRunKey)
{
    $installer = GetInstallerPath
    $keyValue = "`"$installer`" --checkInstall --source=default"
    WriteInfo "Rewriting the HKLM Run key with $keyValue"
    if($([Environment]::Is64BitOperatingSystem))
    {
        SetExpandStringReg $RunKeyPath64 "TeamsMachineInstaller" $keyValue
    }
    else
    {
        SetExpandStringReg $RunKeyPath32 "TeamsMachineInstaller" $keyValue
    }
}

# Get final confirmation we actually did update the installer
$currentVersion = GetInstallerVersion
if($currentVersion)
{
    WriteInfo "New Teams Machine Installer version is $currentVersion"
}
if($currentVersion -eq $targetVersion)
{
    WriteSuccess "Deployment successful, installer is now at target version!"
    try { Write-EventLog -LogName Application -Source $EventLogSource -Category 0 -EntryType Information -EventId 0 -Message "Successfully updated Teams Machine-Wide Installer to $targetVersion" }
    catch { }
    CleanExit
}
else
{
    ErrorExit "ERROR: Script completed, however the Teams Machine-Wide Installer is still not at the target version!" -7
}
