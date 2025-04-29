<#
.SYNOPSIS
    Resets Microsoft 365 (Office) and Teams apps, clears caches, and removes associated account information.
    Alternatively, offers Microsoft Support and Recovery Assistant (SARA) Enterprise scenarios.

.DESCRIPTION
    This script performs one of two major repair actions based on user input:
    1. Clears all Office and Teams caches, resets Microsoft 365 apps, and removes cached credentials.
    2. Runs Microsoft SARA Enterprise support scenarios including Outlook Scan, Office Uninstall, Office Activation, Reset Office Activation, and Shared Computer Activation.

    Script logs all actions to "C:\LOGS\SARA_TOOLKIT.log".

.VERSION
    1.1.0

.AUTHOR
    JS

.DATE
    02/12/2025

.REQUIREMENTS
    - Windows 10 or Windows 11
    - Administrative privileges (required for some operations)
    - Microsoft 365 (Office) and Teams installed (or SARA files available)

.NOTES
    - Ensure you have a backup of critical data before running this script.
    - Logs are saved to "C:\LOGS\SARA_TOOLKIT.log".
    - Requires running this script as Administrator

.CHANGELOG
    1.0.0 - Initial release
    1.1.0 - Added support for Microsoft SARA Enterprise scenarios
#>

param(
    [switch]$Unattended
)

#requires -Version 5.1


$ErrorActionPreference = 'Stop'
$logPath = "C:\LOGS\SARA_TOOLKIT.log"
$logDir = "C:\LOGS"
$saraSourcePath = "\\EXTRACT ENTERPRISE SARA FILES HERE"  #Update With Your Path
$saraLocalPath = "C:\TEMP\SARA"   #SARA Working Folder

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Output $logMessage | Out-File -FilePath $logPath -Append
    Write-Host $logMessage
}

try {
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        Write-Log "Created log directory: $logDir"
    }
} catch {
    Write-Log "Failed to create log directory: $_" -Level "ERROR"
    exit 1
}

Write-Log "Starting Office/Teams Reset or Microsoft SARA Script (Version 1.1.0)"

# Check if Office or Teams are installed
$office64BitPath = "$env:ProgramFiles\Microsoft Office"
$office32BitPath = "$env:ProgramFiles(x86)\Microsoft Office"
$officeInstalled = (Test-Path $office64BitPath) -or (Test-Path $office32BitPath)
$teamsInstalled = Test-Path "$env:LOCALAPPDATA\Microsoft\Teams"

if (-not ($officeInstalled -or $teamsInstalled)) {
    Write-Log "Neither Microsoft 365 nor Teams is installed. Exiting script." -Level "ERROR"
    exit 1
}
else {
    Write-Log "Office is Installed, continuing with Tool."
}

if (-not $Unattended) {
    Write-Host "Select operation mode:`n1. Scripted Reset (Office/Teams) [Default]`n2. Microsoft SARA Enterprise Scenarios"
    $mode = Read-Host "Enter choice (1 or 2)"
} else {
    $mode = "1"
}

if ($mode -eq "2") {
    Write-Log "User selected Microsoft SARA Enterprise Scenarios."
    try {
        if (-not (Test-Path $saraLocalPath)) {
            New-Item -ItemType Directory -Path $saraLocalPath -Force | Out-Null
            Write-Log "Created local SARA path: $saraLocalPath"
        }
        Copy-Item -Path "$saraSourcePath\*" -Destination $saraLocalPath -Recurse -Force
        Write-Log "Copied SARA files from $saraSourcePath to $saraLocalPath"

        $saraScenarios = @(
    @{ Name = 'Outlook Scan';               Arguments = '-S ExpertExperienceAdminTask -AcceptEula -OfflineScan -Logfolder c:\temp' },
    @{ Name = 'Office Uninstall';           Arguments = '-S OfficeScrubScenario -AcceptEula -OfficeVersion All' },
    @{ Name = 'Office Activation';          Arguments = '-S OfficeActivationScenario -AcceptEula -CloseOffice' },
    @{ Name = 'Reset Office Activation';    Arguments = '-S ResetOfficeActivation -AcceptEula -CloseOffice' },
    @{ Name = 'Shared Computer Activation'; Arguments = '-S OfficeSharedComputerScenario -AcceptEula -CloseOffice' }
)

$exe = Join-Path $saraLocalPath 'SaRAcmd.exe'

foreach ($scenario in $saraScenarios) {
    if (-not $Unattended) {
        Write-Host "Run $($scenario.Name)? (Y/N)"
        $confirm = Read-Host "Your choice"
    } else {
        $confirm = "Y"
    }

    if ($confirm -match '^[Yy]$') {
        Write-Log "Launching SARA scenario: $($scenario.Name)"
        try {
            if (-not (Test-Path $exe)) {
                throw "SaRA executable not found at $exe"
            }

            Start-Process -FilePath $exe -ArgumentList $scenario.Arguments -Wait -WindowStyle Hidden
            Write-Log "Completed SARA scenario: $($scenario.Name)"
        } catch {
            Write-Log "Error running SARA scenario [$($scenario.Name)]: $_" -Level "ERROR"
        }
    } else {
        Write-Log "Skipped SARA scenario: $($scenario.Name)"
    }
}
    } catch {
        Write-Log "SARA scenario failure: $_" -Level "ERROR"
    }
} else {
    Write-Log "Proceeding with Scripted Office and Teams Reset..."
}

# Step 1: Clear Office and Teams caches
Write-Host "`nStep 1: Clear Office and Teams caches"
Write-Host "This will delete cached files for Office and Teams, which may resolve performance issues."
$response = Read-Host "Do you want to proceed? [Y/N]"
if ($Unattended -or ($response -match "^[Yy]$")) {
    try {
        Write-Log "Clearing Office caches..."
        $officeCachePaths = @(
            "$env:LOCALAPPDATA\Microsoft\Office\16.0\Cache",
            "$env:LOCALAPPDATA\Microsoft\Office\15.0\Cache",
            "$env:APPDATA\Microsoft\Office\Recent",
            "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
            "$env:LOCALAPPDATA\Microsoft\Office\15.0\Wef"
        )
        foreach ($path in $officeCachePaths) {
            if (Test-Path -Path $path) {
                Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleared cache at: $path"
            }
        }

        Write-Log "Clearing Teams caches..."
        $teamsCachePaths = @(
            "$env:APPDATA\Microsoft\Teams\Cache",
            "$env:APPDATA\Microsoft\Teams\blob_storage",
            "$env:APPDATA\Microsoft\Teams\databases",
            "$env:APPDATA\Microsoft\Teams\GPUcache",
            "$env:APPDATA\Microsoft\Teams\IndexedDB",
            "$env:APPDATA\Microsoft\Teams\Local Storage",
            "$env:APPDATA\Microsoft\Teams\tmp"
        )
        foreach ($path in $teamsCachePaths) {
            if (Test-Path -Path $path) {
                Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleared cache at: $path"
            }
        }
        Write-Log "Office and Teams caches cleared successfully."
    }
    catch {
        Write-Log "Failed to clear caches: $_" -Level "ERROR"
    }
}
else {
    Write-Log "User skipped clearing caches."
}

# Step 2: Reset Office 365 and Teams apps
Write-Host "`nStep 2: Reset Office 365 and Teams apps"
Write-Host "This will reset Office 365 and Teams to their default settings, which may resolve app issues."
$response = Read-Host "Do you want to proceed? [Y/N]"
if ($Unattended -or ($response -match "^[Yy]$")) {
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Write-Log "Error: This script must be run as Administrator to reset packages effectively." -Level "ERROR"
            throw "Administrator privileges required."
        }

        Write-Log "Resetting Office 365 apps..."
        $officeApps = Get-AppxPackage -Name "*Microsoft*.*Office*" -AllUsers
        if ($officeApps) {
            foreach ($app in $officeApps) {
                if (Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue) {
                    Reset-AppxPackage -Package $app.PackageFullName
                    Write-Log "Reset Office app: $($app.Name)"
                } else {
                    Write-Log "Reset-AppxPackage not available. Skipped Office app: $($app.Name)" -Level "WARNING"
                }
            }
        }

        Write-Log "Resetting Teams app..."
        $teamsApps = Get-AppxPackage -Name "*Teams*" -AllUsers
        if ($teamsApps) {
            foreach ($app in $teamsApps) {
                if (Get-Command Reset-AppxPackage -ErrorAction SilentlyContinue) {
                    Reset-AppxPackage -Package $app.PackageFullName
                    Write-Log "Reset Teams app: $($app.Name)"
                } else {
                    Write-Log "Reset-AppxPackage not available. Skipped Teams app: $($app.Name)" -Level "WARNING"
                }
            }
        }
        Write-Log "Office 365 and Teams apps reset completed."
    }
    catch {
        Write-Log "Failed to reset apps: $_" -Level "ERROR"
    }
}
else {
    Write-Log "User skipped resetting apps."
}

# Step 3: Reset/remove Microsoft 365 accounts and cached information
Write-Host "`nStep 3: Reset/remove Microsoft 365 accounts and cached information"
Write-Host "This will remove cached credentials and account information for Microsoft 365."
$response = Read-Host "Do you want to proceed? [Y/N]"
if ($Unattended -or ($response -match "^[Yy]$")) {
    try {
        Write-Log "Removing Microsoft 365 cached credentials..."
        $credentialPaths = @(
            "$env:LOCALAPPDATA\Microsoft\Credentials",
            "$env:APPDATA\Microsoft\Credentials"
        )
        foreach ($path in $credentialPaths) {
            if (Test-Path -Path $path) {
                $credentialFiles = Get-ChildItem -Path $path -Recurse | Where-Object {
                    $_.Name -match "(?i)(microsoft|teams|365)"
                }
                foreach ($file in $credentialFiles) {
                    Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Removed credential file: $($file.FullName)"
                }
            }
        }

        Write-Log "Clearing Office identity cache..."
        foreach ($ver in 14..17) {
            Remove-Item -Path "HKCU:\Software\Microsoft\Office\$ver.0\Common\Identity" -Recurse -Force -ErrorAction SilentlyContinue
        }
        Write-Log "Office identity cache cleared."

        Write-Log "Clearing Teams identity cache..."
        Remove-Item -Path "HKCU:\Software\Microsoft\Teams" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Teams identity cache cleared."
    }
    catch {
        Write-Log "Failed to remove accounts and cached information: $_" -Level "ERROR"
    }
}
else {
    Write-Log "User skipped removing accounts and cached information."
}

# Reboot prompt (shared by both modes)
if (-not $Unattended) {
    Write-Host "\nStep: System reboot"
    Write-Host "A reboot is recommended to complete the process."
    $confirm = Read-Host "Do you want to reboot now? (Y/N)"
} else {
    $confirm = "N"
}

if ($confirm -match "^[Yy]$") {
    Write-Log "User chose to reboot the system."
    try {
        Restart-Computer -Force
    } catch {
        Write-Log "Failed to initiate reboot: $_" -Level "ERROR"
    }
} else {
    Write-Log "User skipped system reboot."
}

Write-Log "Script completed. Logs are saved to: $logPath"
Write-Host "\nScript completed. Logs are saved to: $logPath"
