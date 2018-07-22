Attribute VB_Name = "common_FileSystem"
Option Explicit

Public Function FileSystem_StripExtension(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, ".")
    
    FileSystem_StripExtension = IIf(position = 0, path, Left$(path, position - 1))
End Function

Public Function FileSystem_StripPath(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, "\")
    
    FileSystem_StripPath = IIf(position = 0, path, Right$(path, Len(path) - position))
End Function

Public Function FileSystem_FolderExists(ByVal path As String) As Boolean
    On Error GoTo Err:
    
    FileSystem_FolderExists = (GetAttr(path) And vbDirectory) = vbDirectory
    Exit Function
    
Err:
    FileSystem_FolderExists = False
End Function

Public Function FileSystem_FileExists(ByVal path As String) As Boolean
    On Error GoTo Err:
    
    FileSystem_FileExists = (GetAttr(path) And vbDirectory) <> vbDirectory
    Exit Function
    
Err:
    FileSystem_FileExists = False
End Function
