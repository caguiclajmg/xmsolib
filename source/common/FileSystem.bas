Attribute VB_Name = "common_FileSystem"
Option Explicit

Public Function FileSystem_StripExtension(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, ".")
    
    FileSystem_StripExtension = IIf(position = 0, path, Left$(path, position - 1))
End Function

Public Function FileSystem_EnumerateFiles(ByVal path As String, Optional ByVal match As String = "*", Optional ByVal flags As VbFileAttribute = vbNormal) As String()
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    Dim count As Long, fileName As String
    
    fileName = Dir$(path & match, flags)
    While fileName <> vbNullString
        If (fileName <> ".") And (fileName <> "..") Then count = count + 1
        
        fileName = Dir$()
    Wend
    
    If count = 0 Then Exit Function
    
    Dim index As Long: index = 1
    ReDim result(1 To count) As String
    
    fileName = Dir$(path & match, flags)
    While fileName <> vbNullString
        If (fileName <> ".") And (fileName <> "..") Then
            result(index) = fileName
            index = index + 1
        End If
        
        fileName = Dir$()
    Wend
    
    FileSystem_EnumerateFiles = result
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
