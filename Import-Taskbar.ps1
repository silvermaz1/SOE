<#
.SYNOPSIS
This script imports a configured taskbar and copies any required .lnk files to the AppData folder.

.DESCRIPTION
The script works by performing the following steps:
1. Creates RunOnce Key in first deployed user, this key will call the VBS script that runs the PS script hidden.
2. Creates a Scheduled-Task to check if the Taskbar have been applied, this runs as a "logon-script" and will run for all users that logs on to the device.
3. Creates a Installation result file in Registry that can be used for checking if application was successfully installed. (Please see $StoreResults)
4. Log is written to Windows\Temp, please review this if you have any issues.
#>

# If running in a 64-bit process, relaunch as 32-bit
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH
    }
    Catch {
        Throw "Failed to start $PSCOMMANDPATH"
    }
    Exit
}

# Required variables
$Destination = "$env:SystemDrive\ProgramData\AutoPilotConfig\Taskbar"
$scriptPath = "$Destination\RunHidden.vbs"
$Source = $PSScriptRoot

# Log File Info
$Now = Get-Date -Format "yyyyMMdd-HHmmss"
$LogPath = "$ENV:WINDIR\Temp\ImportTaskbar_$Now.log"

# Function to write log entries
function Write-LogEntry {
    param(
        [string]$Message,
        [string]$Username,
        [string]$Error
    )
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
    
    if ($Username) {
        $logMessage += " - User: $Username"
    }
    
    if ($Error) {
        $logMessage += " - Error: $Error"
    }
    
    $logMessage | Out-File -FilePath $LogPath -Append
}

# CleanUpAndExit
Function CleanUpAndExit() {
    Param(
        [Parameter()][int]$ErrorLevel = 0
    )
	
    # Write results to the registry for Intune Detection
    $StoreResults = "\Count\Taskbar" # Change this to something that fits your organization.
    $Key = "HKEY_LOCAL_MACHINE\Software$StoreResults"
    $NOW = Get-Date -Format "yyyyMMdd-HHmmss"

    If ($ErrorLevel -eq 0) {
        [Microsoft.Win32.Registry]::SetValue($Key, "Success", $NOW)
    } else {
        [Microsoft.Win32.Registry]::SetValue($Key, "Failure", $NOW)
        [Microsoft.Win32.Registry]::SetValue($Key, "Error Code", $Errorlevel)
    }
    
    # Exit Script with the specified ErrorLevel
    EXIT $ErrorLevel
}

# Copy required files
if (!(Test-Path -Path $Destination)) {
    New-Item -ItemType Directory -Path $Destination
}

# Copy all files and folders from $PSScriptRoot to $Destination
Copy-Item -Path "$Source\*" -Destination $Destination -Recurse -Force

# Set RunKey for all users during deployment
function Set-RunKey {
    $Success = $true

    foreach ($userPath in (Get-ChildItem "Registry::HKEY_USERS\" | Where-Object { $_.Name -notmatch '_Classes|S-1-5-18|S-1-5-19|S-1-5-20|\.DEFAULT' })) {
        $username = $userPath.PSChildName

        try {
            $RunPath = "HKEY_USERS\$username\Software\Microsoft\Windows\CurrentVersion\RunOnce"

            # Set the registry values for RunOnce using [Microsoft.Win32.Registry]::SetValue
            [Microsoft.Win32.Registry]::SetValue($RunPath, "TaskbarImport", "wscript.exe `"$scriptPath`"", [Microsoft.Win32.RegistryValueKind]::String)

            # Add a log entry
            Write-LogEntry "Registry values are set correctly for user $username" -Username $username
        }
        catch {
            $Success = $false

            # Add a log entry with the username and error message
            Write-LogEntry "An error occurred for user $username" -Username $username -Error $($_.Exception.Message)
        }
    }

    return $Success
}

# Create a Scheduled Task to run for all users that logon to the machine
function New-ScheduledTaskAllUsers {

    Write-LogEntry "Creating Scheduled Task..."

    $Success = $true

    # Specify the command and argument
    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$Destination\RunHidden.vbs`""

    # Define the principal for all users
    $principal = New-ScheduledTaskPrincipal -GroupId "S-1-5-32-545" -RunLevel Highest

    # Set the trigger to be at any user logon
    $trigger = New-ScheduledTaskTrigger -AtLogOn

    try {
        # Create the scheduled task with the principal for all users
        Register-ScheduledTask -Action $action -Trigger $trigger -TaskName "SetTaskbar" -Description "Set Taskbar at LogOn" -Principal $principal

        Write-LogEntry "Scheduled task created successfully."
    }
    catch {
        $ErrorMessage = "An error occurred while creating the scheduled task: $($_.Exception.Message)"
        Write-LogEntry $ErrorMessage
        Write-Host $ErrorMessage -ForegroundColor Red
        $Success = $false  # Set success to false if an error occurs

        # Add a log entry with the error message
        Write-LogEntry "An error occurred while creating the scheduled task" -Error $($_.Exception.Message)
    }

    return $Success
}

$importRunKeyResult = Set-RunKey
$taskCreationResult = New-ScheduledTaskAllUsers

# If all functions ran successfully, exit with error code 0; otherwise, use error code 101
if ($importRunKeyResult -and $taskCreationResult) {
    Write-LogEntry "All functions completed successfully. Cleaning up and exiting..." -Username $username
    CleanUpAndExit -ErrorLevel 0
} else {
    Write-LogEntry "One or more functions encountered errors. Cleaning up and exiting..." -Username $username
    CleanUpAndExit -ErrorLevel 101
}

Write-LogEntry "Script execution completed."
$LogPath = $null