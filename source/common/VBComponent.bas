Attribute VB_Name = "common_VBComponent"
Option Explicit

Public Function VBComponent_Import(ByRef document As Object, ByVal path As String) As VBComponent
    Set VBComponent_Import = document.VBProject.VBComponents.Import(path)
End Function

Public Sub VBComponent_Export(ByRef document As Object, ByVal name As String, ByVal path As String, Optional ByVal filename As String = vbNullString)
    Dim component As VBComponent: Set component = document.VBProject.VBComponents(name)
    
    Dim extension As String
    Select Case component.Type
        Case vbext_ct_ClassModule, vbext_ct_Document
            extension = "cls"
            
        Case vbext_ct_MSForm
            extension = "frm"
            
        Case vbext_ct_StdModule
            extension = "bas"
            
        Case Else
            extension = vbNullString
    End Select
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    document.VBProject.VBComponents(name).export path & IIf(filename = vbNullString, component.name, filename) & IIf(extension = vbNullString, vbNullString, "." & extension)
End Sub

Public Function VBComponent_ComponentExists(ByRef document As Object, ByVal name As String) As Boolean
    On Error GoTo Err:
    
    Dim component As VBComponent: Set component = document.VBProject.VBComponents(name)
    
    VBComponent_ComponentExists = True
    Exit Function
    
Err:
    VBComponent_ComponentExists = False
End Function
