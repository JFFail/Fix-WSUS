# Fix-WSUS
Designed to be run locally. Copy .ps1 file and make sure the ExecutionPolicy will allow it to be run. Execute with no parameters as an administrator. Upon completion, you can opt to reboot the server immediately or do it later. Once the script has finished, the next task should be to try installing a single patch and check the results of it.

Include the -BackupFiles switch if you want to clean the metafiles associated with patching. The -ClearTemp switch should be used to clear the contents of the default SYSTEM temp directory (C:\Windows\Temp.)
    
Note that if the Remove-WindowsPackage cmdlet fails, you'll likely need to repair the Windows install via:

	Repair-WindowsImage -Online -RestoreHealth -LimitAccess -NoRestart

I don't do it automatically.
