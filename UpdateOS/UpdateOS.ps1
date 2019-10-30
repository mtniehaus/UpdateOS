<#PSScriptInfo

.VERSION 1.0

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

# If we are running as a 32-bit process on a 64-bit system, re-launch as a 64-bit process
if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe")
{
    & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath"
    Exit $lastexitcode
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
Write-Host "Installing updates during white glove technician phase."

# Load module from PowerShell Gallery
Install-PackageProvider -Name NuGet -Force
Install-Module PSWindowsUpdate -Force
Import-Module PSWindowsUpdate

# Install all available updates
Get-WindowsUpdate -Install -IgnoreUserInput -AcceptAll -IgnoreReboot
$needReboot = (Get-WUIsPendingReboot)

# Stop logging
Stop-Transcript

# Specify return code
if ($needReboot)
{
    Write-Host "Reboot is needed."
    Exit 3010
    # Exit 1641
}
else
{
    Exit 0
}
