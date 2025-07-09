' Set the working directory to the script's location
Set objShell = CreateObject("WScript.Shell")
workingDirectory = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = workingDirectory

' Run the PowerShell script with the correct working directory
objShell.Run "powershell.exe -ExecutionPolicy Bypass -File """ & workingDirectory & "\Set-Taskbar.ps1""", 0, True