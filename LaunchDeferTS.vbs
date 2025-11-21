' SCCM Hidden Console Launcher for Task Sequence Deferral Tool
' This VBScript launches PowerShell with a completely hidden console window
' Perfect for SCCM deployments where no console should be visible to end users

' Get the directory where this VBS script is located
Dim objFSO, objShell, strScriptPath, strScriptDir, strPSScript, strCommand

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Get script directory
strScriptPath = WScript.ScriptFullName
strScriptDir = objFSO.GetParentFolderName(strScriptPath)

' Build path to PowerShell script
strPSScript = objFSO.BuildPath(strScriptDir, "deferTS.ps1")

' Build PowerShell command
' -NoProfile: Don't load PowerShell profile
' -NonInteractive: Disable interactive prompts
' -ExecutionPolicy Bypass: Allow script execution
' -WindowStyle Hidden: Hide PowerShell window (redundant but explicit)
strCommand = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strPSScript & """"

' Run PowerShell with hidden window (0 = Hidden, True = Wait for completion)
' Return the exit code from PowerShell
WScript.Quit objShell.Run(strCommand, 0, True)
