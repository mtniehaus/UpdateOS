<#PSScriptInfo

.VERSION 1.4

.GUID 07e4ef9f-8341-4dc4-bc73-fc277eb6b4e6

.AUTHOR Michael Niehaus

.COMPANYNAME Microsoft

.COPYRIGHT

.TAGS Windows AutoPilot Update OS

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
Version 1.4:  Fixed reboot logic.
Version 1.3:  Force use of Microsoft Update/WU.
Version 1.2:  Updated to work on ARM64.
Version 1.1:  Cleaned up output.
Version 1.0:  Original published version.

#>

<#
.SYNOPSIS
Installs the latest Windows 10 quality updates.
.DESCRIPTION
This script uses the PSWindowsUpdate module to install the latest cumulative update for Windows 10.
.EXAMPLE
.\UpdateOS.ps1
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$False)] [Switch] $HardReboot = $false
)

Process
{

# If we are running as a 32-bit process on an x64 system, re-launch as a 64-bit process
if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64")
{
    if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe")
    {
        & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath"
        Exit $lastexitcode
    }
}

# Create a tag file just so Intune knows this was installed
if (-not (Test-Path "$($env:ProgramData)\Microsoft\UpdateOS"))
{
    Mkdir "$($env:ProgramData)\Microsoft\UpdateOS"
}
Set-Content -Path "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.ps1.tag" -Value "Installed"

# Start logging
Start-Transcript "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.log"

# Main logic
$needReboot = $false
Write-Host "Installing updates with HardReboot = $HardReboot."

# Load module from PowerShell Gallery
$null = Install-PackageProvider -Name NuGet -Force
$null = Install-Module PSWindowsUpdate -Force
Import-Module PSWindowsUpdate

# Install all available updates
Get-WindowsUpdate -Install -IgnoreUserInput -AcceptAll -WindowsUpdate -IgnoreReboot | Select-Object Title, KB, Result | Format-Table
$needReboot = (Get-WURebootStatus -Silent).RebootRequired

# Specify return code
if ($needReboot)
{
    Write-Host "Windows Update indicated that a reboot is needed."
}
else
{
    Write-Host "Windows Update indicated that no reboot is required."
}

# For whatever reason, the reboot needed flag is not always being properly set.  So we always want to force a reboot.
# If this script (as an app) is being used as a dependent app, then a hard reboot is needed to get the "main" app to
# install.
if ($HardReboot)
{
    Write-Host "Exiting with return code 1641 to indicate a hard reboot is needed."
    Stop-Transcript
    Exit 1641
}
else
{
    Write-Host "Exiting with return code 3010 to indicate a soft reboot is needed."
    Stop-Transcript
    Exit 3010
}

}
