# UpdateOS
Sample app for installing Windows updates during an Autopilot deployment.
Add the UpdateOS.intunewin app to Intune and specify the following command line:

powershell.exe -noprofile -executionpolicy bypass -file .\UpdateOS.ps1

Optional switches can be added to the command line:

- -Reboot Soft|Hard|None|Delayed: how to handle a required reboot (default Soft, return code 3010).
- -RebootTimeout <seconds>: delay used with -Reboot Delayed (default 120).
- -ExcludeDrivers: only install software (quality) updates.
- -ExcludeUpdates: only install driver updates.
- -IncludeFeatureUpdates: also install feature updates, including enablement packages such as 25H2 on 24H2. These are skipped by default.
- -FirmwareOnly: only install firmware (BIOS/UEFI) updates offered through Windows Update.

The tag file used for detection is written after the updates have been processed. If the install is stopped by the Intune
installation time limit, detection fails and Intune retries the app later. Downloading a cumulative update can take a long
time on slow or shared connections, so consider raising the app's "Installation time required (mins)" from the default of 60.

To "uninstall" the app, the following can be used (for example, to get the app to re-install):

cmd.exe /c del %ProgramData%\Microsoft\UpdateOS\UpdateOS.ps1.tag

Specify the platforms and minimum OS version that you want to support.

For a detection rule, specify the path and file and "File or folder exists" detection method:

%ProgramData%\Microsoft\UpdateOS\UpdateOS.ps1.tag

Deploy the app as a required app to an appropriate set of devices.
