Attribute VB_Name = "common_Regex"
Option Explicit

Public Function RegEx_Test(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As Boolean
#If Mac Then
    ' TODO: RegEx implementation for macOS
    RegEx_Test = True
#Else
    Dim regexp As Object: Set regexp = CreateObject("VBScript.RegExp")
    regexp.pattern = pattern
    regexp.Global = flagGlobal
    regexp.ignoreCase = flagIgnoreCase
    
    RegEx_Test = regexp.test(test)
#End If
End Function

Public Function RegEx_Replace(ByVal test As String, ByVal pattern As String, ByVal replace As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As String
#If Mac Then
    ' TODO: RegEx implementation for macOS
    RegEx_Replace = test
#Else
    Dim regexp As Object: Set regexp = CreateObject("VBScript.RegExp")
    regexp.pattern = pattern
    regexp.Global = flagGlobal
    regexp.ignoreCase = flagIgnoreCase
    
    RegEx_Replace = regexp.replace(test, replace)
#End If
End Function
