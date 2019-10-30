# UpdateOS
Sample app for installing Windows updates during an Autopilot deployment.
Add the UpdateOS.intunewin app to Intune and specify the following command line:

powershell.exe -noprofile -executionpolicy bypass -file .\UpdateOS.ps1

To "uninstall" the app, the following can be used (for example, to get the app to re-install):

cmd.exe /c del %ProgramData%\Microsoft\UpdateOS\UpdateOS.ps1.tag

Specify the platforms and minimum OS version that you want to support.

For a detection rule, specify the path and file and "File or folder exists" detection method:

%ProgramData%\Microsoft\UpdateOS
UpdateOS.ps1.tag

Deploy the app as a required app to an appropriate set of devices.