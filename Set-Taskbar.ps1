<#
.SYNOPSIS
This script imports a configured taskbar and copies any required .lnk files to the AppData folder.

.DESCRIPTION
The script works by performing the following steps:
1. Imports shortcut files for new taskbar.
2. Imports registry values for taskbar icons.
3. Restarts Explorer.exe to make sure Taskbar settings is instantly applied.
4. Creates a Registry key to check if the taskbar have been applied to the user already (TaskbarCheck is placed in HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband)
#>

$RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
$RegistryValueFavoritesResolve = "FavoritesResolve"
$RegistryValueFavorites = "Favorites"
$RegistryValueRun = "TaskbarCheck"
$RegFilePath = "$PSScriptRoot\Taskbar.reg"

# Restart Explorer
function Restart-Explorer {
    try {
        taskkill /f /im explorer.exe
        Start-Process explorer.exe
        Write-Host "Explorer.exe restarted successfully."
    }
    catch {
        $errorMessage = "Failed to restart explorer.exe. Error: $($_.Exception.Message)"
        Write-Host $errorMessage
    }
}

# Copy-TaskbarFiles
function Copy-TaskbarFilesForCurrentUser {
    param(
        [string]$sourceDir = "$PSScriptRoot\AppData"
    )

    try {
        $userProfilePath = $env:USERPROFILE
        $targetDir = Join-Path -Path $userProfilePath -ChildPath "AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

        # Ensure the target directory exists
        if (!(Test-Path -Path $targetDir)) {
            New-Item -ItemType Directory -Force -Path $targetDir
        }

        # Copy files from source to target
        Copy-Item -Path "$sourceDir\*" -Destination $targetDir -Recurse -Force

        Write-Host "Files are copied successfully for the current user."
    }
    catch {
        $ErrorMessage = "An error occurred while copying files: $($_.Exception.Message)"
        Write-Host $ErrorMessage -ForegroundColor Red
    }
}

# Modify-RegistryValueFromRegFile
function Modify-RegistryValueFromRegFile {
# Read the content of the .reg file
$RegContent = Get-Content -Path $RegFilePath -Raw

# Use regex to extract the hex data from the .reg content
$HexData = [regex]::Match($RegContent, '"Favorites"=hex:([^"]+)').Groups[1].Value

# Remove line continuation characters, spaces, and carriage returns
$HexData = $HexData -replace '[\r\n\\ ]', ''

# Convert the hexadecimal data to a byte array
$ByteArray = $HexData -split ',' | ForEach-Object { [byte]([convert]::ToInt32($_, 16)) }

# Modify the desired registry value with the new byte array
Set-ItemProperty -Path $RegistryPath -Name $RegistryValueFavorites -Value $ByteArray -Force

# Set RegistryValueFavoritesResolve for taskbar consitency
Set-ItemProperty -Path $userRegistryPath -Name $RegistryValueFavoritesResolve -Value $ByteArray -Force

# Set TaskbarCheck to make sure it only run once per user
Set-ItemProperty -Path $RegistryPath -Name $RegistryValueRun -Value "true" -Force

Write-Host "Modified taskbar registry values."
}

# Check if the registry value indicating taskbar changes have been made
$taskbarChangesMade = Get-ItemProperty -Path $RegistryPath -Name $RegistryValueRun -ErrorAction SilentlyContinue

if ($taskbarChangesMade -eq $null) {
    # The registry value indicating taskbar changes has not been set, so run Taskbar functions
    Write-Host "Setting Taskbar"
    Copy-TaskbarFilesForCurrentUser
    Modify-RegistryValueFromRegFile
    Restart-Explorer
} else {
    Write-Host "The Taskbar changes have already been set. Nothing to do."
}