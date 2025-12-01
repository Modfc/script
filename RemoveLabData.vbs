Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Define the base folder path
baseFolder = "C:\\Users\\ISLab_User\\temp"

' Define the folders to be cleared
directories = Array( _
    "Z:\\3D Objects", _
    "Z:\\My Desktop", _
    "Z:\\My Documents", _
    "Z:\\My Download", _
    "Z:\\My Pictures", _
    "Z:\\My Music", _
    "Z:\\My Videos" _
)

' Delete contents of specified directories
For Each folder In directories
    If objFSO.FolderExists(folder) Then
        On Error Resume Next ' Ignore errors if a file or folder is locked
        For Each file In objFSO.GetFolder(folder).Files
            file.Delete True
        Next
        For Each subFolder In objFSO.GetFolder(folder).SubFolders
            subFolder.Delete True
        Next
        On Error GoTo 0
    End If
Next

' Delete the contents of "C:\Users\ISLab_User\temp\My Downloads"
destinationFolder = "C:\\Users\\ISLab_User\\temp\\My Downloads"
If objFSO.FolderExists(destinationFolder) Then
    On Error Resume Next
    For Each file In objFSO.GetFolder(destinationFolder).Files
        file.Delete True
    Next
    For Each subFolder In objFSO.GetFolder(destinationFolder).SubFolders
        subFolder.Delete True
    Next
    On Error GoTo 0
End If

' Empty the contents of "Z:\My Download" into "C:\Users\ISLab_User\temp\My Downloads"
sourceFolder = "Z:\\My Download"
If objFSO.FolderExists(sourceFolder) Then
    If Not objFSO.FolderExists(destinationFolder) Then
        objFSO.CreateFolder(destinationFolder)
    End If

    On Error Resume Next
    ' Move files from source to destination
    For Each file In objFSO.GetFolder(sourceFolder).Files
        objFSO.MoveFile file.Path, destinationFolder & "\\"
    Next

    ' Move subfolders from source to destination
    For Each subFolder In objFSO.GetFolder(sourceFolder).SubFolders
        objFSO.MoveFolder subFolder.Path, destinationFolder & "\\"
    Next
    On Error GoTo 0
End If

' Empty Recycle Bin
objShell.Run "cmd.exe /c echo Y| PowerShell.exe -NoProfile -Command Clear-RecycleBin", 0, True

' Set the custom Startup folder path
startupFolder = "C:\\Users\\ISLab_User\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
scriptPath = WScript.ScriptFullName

' Check if the script is already in the custom Startup folder
If Not objFSO.FileExists(startupFolder & "\\RemoveLabData.vbs") Then
    objFSO.CopyFile scriptPath, startupFolder & "\\RemoveLabData.vbs", True
End If
