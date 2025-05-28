<#PSScriptInfo

.VERSION 2.0

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
Version 2.0:  Restructured download and install logic
Version 1.10: Fixed AcceptEula logic.
Version 1.9:  Added -ExcludeUpdates switch.
Version 1.8:  Added logic to pass the -ExcludeDrivers switch when relaunching as 64-bit.
Version 1.7:  Switched to Windows Update COM objects.
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
    [Parameter(Mandatory = $False)] [switch] $ExcludeDrivers,
    [Parameter(Mandatory = $False)] [switch] $ExcludeUpdates
)

Process {

    # If we are running as a 32-bit process on an x64 system, re-launch as a 64-bit process
    if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64") {
        if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe") {
            if ($ExcludeDrivers) {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout -ExcludeDrivers
            } elseif ($ExcludeUpdates) {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout -ExcludeUpdates
            } else {
                & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath" -Reboot $Reboot -RebootTimeout $RebootTimeout
            }
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
        # Updates only
        $queries = @("IsInstalled=0 and Type='Software'")
    }
    elseif ($ExcludeUpdates) {
        # Drivers only
        $queries = @("IsInstalled=0 and Type='Driver'")
    } else {
        # Both
        $queries = @("IsInstalled=0 and Type='Software'", "IsInstalled=0 and Type='Driver'")
    }

    $WUUpdates = New-Object -ComObject Microsoft.Update.UpdateColl
    $queries | ForEach-Object {
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts Getting $_ updates."        
        try {
            ((New-Object -ComObject Microsoft.Update.Session).CreateupdateSearcher().Search($_)).Updates | ForEach-Object {
                if (!$_.EulaAccepted) { $_.AcceptEula() }
                if ($_.Title -notmatch "Preview") { [void]$WUUpdates.Add($_) }
            }
        } catch {
            # If this script is running during specialize, error 8024004A will happen:
            # 8024004A	Windows Update agent operations are not available while OS setup is running.
            $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
            Write-Warning "$ts Unable to search for updates: $_"
        }
    }

    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($WUUpdates.Count -eq 0) {
        Write-Host "$ts No Updates Found"
        Exit 0
    } else {
        Write-Host "$ts Updates found: $($WUUpdates.count)"
    }
    
    foreach ($update in $WUUpdates) {
    
        $singleUpdate = New-Object -ComObject Microsoft.Update.UpdateColl
        $singleUpdate.Add($update) | Out-Null
    
        $WUDownloader = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateDownloader()
        $WUDownloader.Updates = $singleUpdate
    
        $WUInstaller = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateInstaller()
        $WUInstaller.Updates = $singleUpdate
        $WUInstaller.ForceQuiet = $true
    
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Output "$ts Downloading update: $($update.Title)"
        $Download = $WUDownloader.Download()
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts   Download result: $($Download.ResultCode) ($($Download.HResult))"
    
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts Installing update: $($update.Title)"
        $Results = $WUInstaller.Install()
        $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
        Write-Host "$ts   Install result: $($Results.ResultCode) ($($Results.HResult))"

        # result code 2 = success, see https://learn.microsoft.com/en-us/windows/win32/api/wuapi/ne-wuapi-operationresultcode
        
        if ($Results.RebootRequired) {
            $script:needReboot = $true
        }
    }

    # Specify return code
    $ts = get-date -f "yyyy/MM/dd hh:mm:ss tt"
    if ($script:needReboot) {
        Write-Host "$ts Windows Update indicated that a reboot is needed."

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
    }
    else {
        Write-Host "$ts Windows Update indicated that no reboot is required."
    }
    Exit 0
}
