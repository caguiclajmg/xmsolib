Attribute VB_Name = "common_VBComponent"
Option Explicit

Public Function VBComponent_GetCode(ByVal component As VBComponent) As String
    VBComponent_GetCode = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
End Function

Public Function VBComponent_FromString(ByVal project As VBProject, ByVal ctype As vbext_ComponentType, ByVal name As String, ByVal code As String) As VBComponent
    Dim component As VBComponent: Set component = project.VBComponents.Add(ctype)
    component.CodeModule.AddFromString code
    
    Set VBComponent_FromString = component
End Function

Public Function VBComponent_Import(ByVal project As VBProject, ByVal path As String) As VBComponent
    Set VBComponent_Import = project.VBComponents.Import(path)
End Function

Public Sub VBComponent_Export(ByVal project As VBProject, ByVal name As String, ByVal path As String, Optional ByVal filename As String = vbNullString)
    Dim component As VBComponent: Set component = project.VBComponents(name)
    
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
    project.VBComponents(name).export path & IIf(filename = vbNullString, component.name, filename) & IIf(extension = vbNullString, vbNullString, "." & extension)
End Sub

Public Function VBComponent_Exists(ByVal project As VBProject, ByVal name As String) As Boolean
    On Error GoTo Err:
    
    Dim component As VBComponent: Set component = project.VBComponents(name)
    
    VBComponent_Exists = True
    Exit Function
    
Err:
    VBComponent_Exists = False
End Function
