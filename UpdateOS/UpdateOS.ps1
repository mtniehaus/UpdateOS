<#PSScriptInfo

.VERSION 1.6

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
Version 1.6:  Default to soft reboot.
Version 1.5:  Improved logging, reboot logic.
Version 1.4:  Fixed reboot logic.
Version 1.3:  Force use of Microsoft Update/WU.
Version 1.2:  Updated to work on ARM64.
Version 1.1:  Cleaned up output.
Version 1.0:  Original published version.

#>

<#
.SYNOPSIS
Installs the latest Windows 10/11 quality updates.
.DESCRIPTION
This script uses the Windows Update COM objects to install the latest cumulative updates for Windows 10/11.
.EXAMPLE
.\UpdateOS.ps1
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $False)] [ValidateSet('Soft', 'Hard', 'None', 'Delayed')] [String] $Reboot = 'Soft',
    [Parameter(Mandatory = $False)] [Int32] $RebootTimeout = 120,
    [Parameter(Mandatory = $False)] [switch] $ExcludeDrivers
)

Process {

    # If we are running as a 32-bit process on an x64 system, re-launch as a 64-bit process
    if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64") {
        if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe") {
            & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout
            Exit $lastexitcode
        }
    }

    # Create a tag file just so Intune knows this was installed
    if (-not (Test-Path "$($env:ProgramData)\Microsoft\UpdateOS")) {
        Mkdir "$($env:ProgramData)\Microsoft\UpdateOS"
    }
    Set-Content -Path "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.ps1.tag" -Value "Installed"

    # Start logging
    Start-Transcript "$($env:ProgramData)\Microsoft\UpdateOS\UpdateOS.log"

    # Main logic
    $script:needReboot = $false

    # Opt into Microsoft Update
    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    Write-Output "$ts Opting into Microsoft Update"
    $ServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
    $ServiceID = "7971f918-a847-4430-9279-4a52d1efe18d"
    $ServiceManager.AddService2($ServiceId, 7, "") | Out-Null

    # Install all available updates
    $WUDownloader = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
    $WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
    if ($ExcludeDrivers) {
        $queries = @("IsInstalled=0 and Type='Software'")
    }
    else {
        $queries = @("IsInstalled=0 and Type='Software'", "IsInstalled=0 and Type='Driver'")
    }

    $queries | ForEach-Object {

        $WUUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts Getting $_ updates."        
        ((New-Object -ComObject Microsoft.Update.Session).CreateupdateSearcher().Search($_)).Updates | ForEach-Object {
            if (!$_.EulaAccepted) { $_.EulaAccepted = $true }
            if ($_.Title -notmatch "Preview") { [void]$WUUpdates.Add($_) }
        }

        if ($WUUpdates.Count -ge 1) {
            $WUInstaller.ForceQuiet = $true
            $WUInstaller.Updates = $WUUpdates
            $WUDownloader.Updates = $WUUpdates
            $UpdateCount = $WUDownloader.Updates.count
            if ($UpdateCount -ge 1) {
                $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
                Write-Output "$ts Downloading $UpdateCount Updates"
                foreach ($update in $WUInstaller.Updates) { Write-Output "$($update.Title)" }
                $Download = $WUDownloader.Download()
            }
            $InstallUpdateCount = $WUInstaller.Updates.count
            if ($InstallUpdateCount -ge 1) {
                $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
                Write-Output "$ts Installing $InstallUpdateCount Updates"
                $Install = $WUInstaller.Install()
                $ResultMeaning = ($Results | Where-Object { $_.ResultCode -eq $Install.ResultCode }).Meaning
                Write-Output $ResultMeaning
                $script:needReboot = $Install.RebootRequired
            } 
        }
        else {
            Write-Output "No Updates Found"
        } 
    }

    # Specify return code
    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($script:needReboot) {
        Write-Host "$ts Windows Update indicated that a reboot is needed."
    }
    else {
        Write-Host "$ts Windows Update indicated that no reboot is required."
    }

    # For whatever reason, the reboot needed flag is not always being properly set.  So we always want to force a reboot.
    # If this script (as an app) is being used as a dependent app, then a hard reboot is needed to get the "main" app to
    # install.
    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($Reboot -eq "Hard") {
        Write-Host "$ts Exiting with return code 1641 to indicate a hard reboot is needed."
        Stop-Transcript
        Exit 1641
    }
    elseif ($Reboot -eq "Soft") {
        Write-Host "$ts Exiting with return code 3010 to indicate a soft reboot is needed."
        Stop-Transcript
        Exit 3010
    }
    elseif ($Reboot -eq "Delayed") {
        Write-Host "$ts Rebooting with a $RebootTimeout second delay"
        & shutdown.exe /r /t $RebootTimeout /c "Rebooting to complete the installation of Windows updates."
        Exit 0
    }
    else {
        Write-Host "$ts Skipping reboot based on Reboot parameter (None)"
        Exit 0
    }

}
