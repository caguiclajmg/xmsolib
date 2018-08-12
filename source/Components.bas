Attribute VB_Name = "modComponents"
Option Explicit
#If Mac Then
Private Const DIRECTORY_SEPARATOR As String = "/"
#Else
Private Const DIRECTORY_SEPARATOR As String = "\"
#End If

Private Function FolderExists(ByVal path As String) As Boolean
    On Error GoTo Err:
    
    FolderExists = (GetAttr(path) And vbDirectory) = vbDirectory
    Exit Function
    
Err:
    FolderExists = False
End Function

Private Function FileExists(ByVal path As String) As Boolean
    On Error GoTo Err:
    
    FileExists = (GetAttr(path) And vbDirectory) <> vbDirectory
    Exit Function
    
Err:
    FileExists = False
End Function

Private Sub CreateDirectory(ByVal path As String)
    MkDir path
End Sub

Public Sub Components_Save(ByVal id As String, ByVal document As Object, ByVal path As String)
    Dim components As VBComponents: Set components = document.VBProject.VBComponents
    
    If Right$(path, 1) <> DIRECTORY_SEPARATOR Then path = path & DIRECTORY_SEPARATOR
    
    Dim component As VBComponent
    For Each component In components
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
        
        Dim component_path As String, component_name As String
        If component.Type = vbext_ct_ClassModule Then
            component_path = path & "common" & DIRECTORY_SEPARATOR
            component_name = component.name
        ElseIf component.Type = vbext_ct_StdModule Then
            If Not ((component.name Like "common_*") Or (component.name Like id & "_*")) Then GoTo NextComponent
            
            Dim tokens() As String: tokens = Split(component.name, "_")
            
            component_path = path & tokens(0) & DIRECTORY_SEPARATOR
            component_name = tokens(1)
        Else
            GoTo NextComponent
        End If
        
        component_path = component_path & component_name & IIf(extension = vbNullString, vbNullString, "." & extension)
        
        If FileExists(component_path) Then Kill (component_path)
        
        component.Export component_path
        
NextComponent:
    Next
    
    components("modComponents").Export path & "Components.bas"
End Sub

Public Sub Components_Clear(ByVal document As Object)
    On Error Resume Next
    
    Dim components As VBComponents: Set components = document.VBProject.VBComponents
    
    Dim regexp As Object: Set regexp = CreateObject("vbscript.regexp")
    regexp.pattern = "(excel|common)\_(.+)"
    
    Dim component As VBComponent
    For Each component In components
        If regexp.test(component.name) Then components.Remove component
    Next
End Sub

Public Sub Components_Load(ByVal id As String, ByVal document As Object)
    Dim components As VBComponents: Set components = document.VBProject.VBComponents
    Dim path_root As String: path_root = document.path & DIRECTORY_SEPARATOR
    
    Dim path_base, file As String
    
    path_base = path_root & "common" & DIRECTORY_SEPARATOR
    file = Dir$(path_base, vbNormal)
    While file <> vbNullString
        components.Import path_base & file
        file = Dir$()
    Wend

    path_base = path_root & id & DIRECTORY_SEPARATOR
    file = Dir$(path_base, vbNormal)
    While file <> vbNullString
        components.Import path_base & file
        file = Dir$()
    Wend
End Sub

Public Sub Components_Reload(ByVal id As String, ByVal document As Object)
    Components_Clear document
    Components_Load id, document
End Sub

Public Sub Components_Compile(ByVal id As String, ByVal document As Object)
    Dim components As VBComponents: Set components = document.VBProject.VBComponents
    
    Dim resultCode As String
    
    Dim component As VBComponent
    For Each component In components
        If (component.Type = vbext_ct_StdModule) And ((component.name Like id & "_*") Or (component.name Like "common_*")) Then
            Dim componentCode As String: componentCode = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
            componentCode = replace(componentCode, "Option Explicit", vbNullString)
            resultCode = resultCode & componentCode
        End If
    Next
    
    Dim resultComponent As VBComponent: Set resultComponent = document.VBProject.VBComponents.Add(vbext_ct_StdModule)
    resultComponent.CodeModule.AddFromString resultCode
    resultComponent.name = "xmsolib"
    
    resultComponent.Export document.path & DIRECTORY_SEPARATOR & ".." & DIRECTORY_SEPARATOR & "build" & DIRECTORY_SEPARATOR & "xmsolib.bas"
    components.Remove resultComponent
End Sub
