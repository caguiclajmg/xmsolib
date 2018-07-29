Attribute VB_Name = "common_VBComponent"
Option Explicit

Public Sub VBComponent_Export(ByRef document As Object, ByVal name As String, ByVal path As String)
    document.VBProject.VBComponents(name).export path
End Sub

Public Function VBComponent_ComponentExists(ByRef document As Object, ByVal name As String) As Boolean
    On Error GoTo Err:
    
    Dim component As VBComponent: Set component = document.VBProject.VBComponents(name)
    
    VBComponent_ComponentExists = True
    Exit Function
    
Err:
    VBComponent_ComponentExists = False
End Function
