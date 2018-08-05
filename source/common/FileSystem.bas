Attribute VB_Name = "common_FileSystem"
Option Explicit
#If Mac Then
Public Const DIRECTORY_SEPARATOR As String = "/"
#Else
Public Const DIRECTORY_SEPARATOR As String = "\"
#End If

Public Function FileSystem_StripExtension(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, ".")
    
    FileSystem_StripExtension = IIf(position = 0, path, Left$(path, position - 1))
End Function

Public Function FileSystem_EnumerateFiles(ByVal path As String, Optional ByVal match As String = "*", Optional ByVal flags As VbFileAttribute = vbNormal) As String()
    If Right$(path, 1) <> "\" Then path = path & DIRECTORY_SEPARATOR
    
    Dim count As Long, filename As String
    
    filename = Dir$(path & match, flags)
    While filename <> vbNullString
        If (filename <> ".") And (filename <> "..") Then count = count + 1
        
        filename = Dir$()
    Wend
    
    If count = 0 Then Exit Function
    
    Dim index As Long: index = 1
    ReDim result(1 To count) As String
    
    filename = Dir$(path & match, flags)
    While filename <> vbNullString
        If (filename <> ".") And (filename <> "..") Then
            result(index) = filename
            index = index + 1
        End If
        
        filename = Dir$()
    Wend
    
    FileSystem_EnumerateFiles = result
End Function

Public Function FileSystem_StripPath(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, "\")
    
    FileSystem_StripPath = IIf(position = 0, path, Right$(path, Len(path) - position))
End Function

Public Function FileSystem_FolderExists(ByVal path As String) As Boolean
    On Error GoTo Error:
    
    FileSystem_FolderExists = (GetAttr(path) And vbDirectory) = vbDirectory
    Exit Function
    
Error:
    FileSystem_FolderExists = False
End Function

Public Function FileSystem_FileExists(ByVal path As String) As Boolean
    On Error GoTo Error:
    
    FileSystem_FileExists = (GetAttr(path) And vbDirectory) <> vbDirectory
    Exit Function
    
Error:
    FileSystem_FileExists = False
End Function
