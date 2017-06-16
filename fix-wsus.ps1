<#
    Author: John Fabry
    Creation Date: 5/25/2017
    Last Update: 6/9/2017
    Point: Blanket fix for WSUS corruption to reset components. Designed for Server 2012 R2.
    
    Usage: Designed to be run locally. Copy .ps1 file and make sure the ExecutionPolicy will
      allow it to be run. Execute with no parameters as an administrator. Upon completion,
      you can opt to reboot the server immediately or do it later. Once the script has finished,
      the next task should be to try installing a single patch and check the results of it.
      Include the -BackupFiles switch if you want to clean the metafiles associated with patching.
      The -ClearTemp switch should be used to clear the contents of the default SYSTEM temp
      directory (C:\Windows\Temp.)
    
      Note that if the Remove-WindowsPackage cmdlet fails, you'll likely need to repair the
      Windows install via:
    
      Repair-WindowsImage -Online -RestoreHealth -LimitAccess -NoRestart
    
      I don't do it automatically.
#>

# Parameter for *if* we want to clear all the files since I many times do not.
param
(
    [Parameter(Mandatory=$false)][switch]$BackupFiles,
    [Parameter(Mandatory=$false)][switch]$ClearTemp
)

# Just because I try-catch so many times and got tired of having these 3 lines instead of 1.
function PromptAndQuit
{
    param([string]$Message)

    Write-Output "$Message"
    Read-Host "Hit Enter to quit" | Out-Null
    exit
}

# Display that we're starting.
Write-Output "Running started: $(Get-Date)"

# Remove contents of the default Windows system temp directory.
if($ClearTemp) {
    Write-Output "Clearing the contents of C:\Windows\Temp...`n`tThis will take a long time."
    Remove-Item -Path "C:\Windows\Temp\*" -Recurse -ErrorAction SilentlyContinue
}

# Do the backups if requested.
if($BackupFiles) {
    # Rename the log folders.
    if((Get-Service -Name wuauserv).Status -ne "Stopped") {
        Write-Output "Windows Update is running. Stopping it..."
        try {
            Stop-Service -Name wuauserv -Force -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not stop Windows Update service!"
        }
    }
    if((Get-Service -Name CryptSvc).Status -ne "Stopped") {
        Write-Output "Crypto is running. Stopping it..."
        try {
            Stop-Service -Name CryptSvc -Force -ErrorAction Stop 
        } catch {
            PromptAndQuit -Message "WARNING: Could not stop Crypto service!"
        }
    }
    if((Get-Service -Name BITS).Status -ne "Stopped") {
        Write-Output "Background transfer is running. Stopping it..."
        try {
            Stop-Service -Name BITS -Force -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not stop Background transfer service!"
        }
    }
    if((Get-Service -Name msiserver).Status -ne "Stopped") {
        Write-Output "MSI is running. Stopping it..."
        try {
            Stop-Service -Name msiserver -Force -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not stop MSI service!"
        }
    }

    # Check if SoftwareDistribution and catroot2 have already got bad copies to remove first.
    Write-Output "Beginning rename and backup process..."
    if(Test-Path -LiteralPath C:\Windows\SoftwareDistribution.old) {
        Write-Output "`tRemoving previous SoftwareDistribution.old folder..."
        try {
            Remove-Item -LiteralPath C:\Windows\SoftwareDistribution.old -Recurse -Force -Confirm:$false -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not remove SoftwareDistribution.old!"
        }
    }
    if(Test-Path -LiteralPath C:\Windows\SoftwareDistribution) {
        Write-Output "`tRenaming current SoftwareDistribution folder to .old..."
        try {
            Move-Item -LiteralPath C:\Windows\SoftwareDistribution -Destination C:\Windows\SoftwareDistribution.old -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not move SoftwareDistribution!"
        }
    }
    if(Test-Path -LiteralPath C:\Windows\System32\catroot2.old) {
        Write-Output "`tRemoving previous catroot2.old folder..."
        try {
            Remove-Item -LiteralPath C:\Windows\System32\catroot2.old -Recurse -Force -Confirm:$false -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not remove catroot2.old!"
        }
    }
    if(Test-Path -LiteralPath C:\Windows\System32\catroot2) {
        Write-Output "`tRenaming current catroot2 folder to .old..."
        try {
            Move-Item -LiteralPath C:\Windows\System32\catroot2 -Destination C:\Windows\System32\catroot2.old -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not move catroot2!"
        }
    }

    # Check if WindowsUpdate.log already has a bad copy to remove first.
    if(Test-Path -LiteralPath C:\Windows\WindowsUpdate.log.old) {
        Write-Output "`tRemoving previous WindowsUpdate.log.old file..."
        try {
            Remove-Item -LiteralPath C:\windows\WindowsUpdate.log.old -Force -Confirm:$false -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING: Could not remove WindowsUpdate.log.old!"
        }
    }
    if(Test-Path -LiteralPath C:\Windows\WindowsUpdate.log) {
        Write-Output "`tRenaming current WindowsUpdate.log to .old..."
        try {
            Move-Item -LiteralPath C:\Windows\WindowsUpdate.log -Destination C:\Windows\WindowsUpdate.log.old -Force -ErrorAction Stop -Confirm:$false
        } catch {
            PromptAndQuit -Message "WARNING: Could not rename WindowsUpdate.log!"
        }
    }

    # Restart the services.
    Write-Output "Restarting the previously stopped services..."
    try {
        Start-Service -Name wuauserv -ErrorAction Stop
    } catch {
        PromptAndQuit -Message "WARNING: Could not restart Windows Update service!"
    }
    try {
        Start-Service -Name CryptSvc -ErrorAction Stop
    } catch {
        PromptAndQuit -Message "WARNING: Could not restart Crypto service!"
    }
    try {
        Start-Service -Name BITS -ErrorAction Stop
    } catch {
        PromptAndQuit -Message "WARNING: Could not restart Background transfer service!"
    }
    try {
        Start-Service -Name msiserver -ErrorAction Stop
    } catch {
        PromptAndQuit -Message "WARNING: Could not restart MSI service!"
    }
}

# Get all of the packages which are not "Installed".
Write-Output "Checking for any packages not fully installed..."
try {
    $packages = Get-WindowsPackage -Online | Where-Object { $_.PackageState -ne "Installed" }
} catch {
    PromptAndQuit -Message "WARNING: Could not get package list!"
}
$totalPackages = $packages.Count
Write-Output "`n`tFound $totalPackages packages to remove...`n"

# Process each package and remove it.
if($totalPackages -gt 0) {
    Write-Output "Starting package removal..."
    $currentNumber = 0
    foreach($package in $packages) {
        $currentNumber++
        Write-Output "Processing removal on package $currentNumber of $totalPackages."
        Write-Output "`tRemoving $($package.PackageName)..."
        try {
            Remove-WindowsPackage -Online -NoRestart -PackageName $package.PackageName -ErrorAction Stop
        } catch {
            PromptAndQuit -Message "WARNING! Could not remove $($package.PackageName)!"
        }
    }

    # Show when finished.
    Write-Output "Running finished: $(Get-Date)"

    # See if a reboot is in the cards. Reboots are suppressed via Remove-WindowsPackage, but you should pay attention to
    # whether or not one is needed. You'll receive a suppression warning (in yellow) if one is required.
    Write-Output "Do you want to reboot?"
    Write-Output "`n`tIf you see a yellow restart suppression warning above, it's recommended!"
    while($true) {
        $needReboot = Read-Host "Y or N"

        if($needReboot.ToLower() -eq "n") {
            Write-Output "Done!"
            break
        } elseif ($needReboot.ToLower() -eq "y") {
            Write-Output "Rebooting in 30 seconds..."
            Restart-Computer
            break
        } else {
            Write-Output "Enter 'y' or 'n'!"
        }
    }
} else {
    Write-Output "No bad packages found!"
    Write-Output "Running finished: $(Get-Date)"
}