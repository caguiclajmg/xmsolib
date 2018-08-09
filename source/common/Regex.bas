Attribute VB_Name = "common_Regex"
Option Explicit
#If Not Mac Then
Private regexpObject_ As Object
#End If

Public Type RegexMatch
    Index As Long
    submatches() As String
    Value As String
End Type

Public Function RegEx_Execute(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As RegexMatch()
    Dim result() As RegexMatch
    
#If Mac Then
    ' TODO: RegEx implementation for macOS
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")
    
    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    
    Dim matches As Object: Set matches = regexpObject_.Execute(test)
    ReDim result(1 To matches.count) As RegexMatch
    
    Dim i As Long
    For i = LBound(result) To UBound(result)
        Dim match As Object: Set match = matches.Item(i - 1)
        
        result(i).Index = match.FirstIndex
        
        Dim submatches As Object: Set submatches = match.submatches
        
        If submatches.count > 0 Then
            ReDim result(i).submatches(1 To submatches.count)
            
            Dim j As Long
            For j = LBound(result(i).submatches) To UBound(result(i).submatches)
                result(i).submatches(j) = submatches.Item(j - 1)
            Next
        End If
        
        result(i).Value = match.Value
    Next
#End If

    RegEx_Execute = result
End Function

Public Function RegEx_Test(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As Boolean
#If Mac Then
    ' TODO: RegEx implementation for macOS
    RegEx_Test = True
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")
    
    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    
    RegEx_Test = regexpObject_.test(test)
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
