Option Explicit

Dim strProgramDataFolder, strHome, strUser, WshShell, SysRoot
Dim intRunError, objShell, objFSO

Set WshShell = CreateObject("Wscript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

SysRoot = WshShell.ExpandEnvironmentStrings("%SystemDrive%")
strProgramDataFolder = SysRoot & "\ProgramData\AJG"

If objFSO.FolderExists(strProgramDataFolder) Then

	'Wscript.Echo "Folder " & strProgramDataFolder & " exists. Attempting to update permissions..."
    ' Assign user permission to home folder
    intRunError = WshShell.Run("%COMSPEC% /c Echo Y | cacls " & strProgramDataFolder & " /t /e /c /g Users:C ", 2, True)
    
    
    If intRunError <> 0 Then
        'Wscript.Echo "Error assigning permissions for user " & strUser & " to home folder " & strProgramDataFolder
    End If
    
End If

WScript.Quit
