Attribute VB_Name = "common_RegEx"
Option Explicit
#If Not Mac Then
Private regexpObject_ As Object
#End If

Public Function RegEx_Execute(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False, Optional ByVal flagMultiLine As Boolean = True) As RegexMatchCollection
    Dim result As RegexMatchCollection: Set result = New RegexMatchCollection

#If Mac Then
    ' TODO: RegEx implementation for macOS
    Err.Raise 5, Description:="RegEx_Execute not yet implemented on macOS."
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")

    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    regexpObject_.multiLine = flagMultiLine

    Dim matches As Object: Set matches = regexpObject_.Execute(test)

    Dim match As match
    For Each match In matches
        Dim resultMatch As RegexMatch: Set resultMatch = New RegexMatch
        resultMatch.Index = match.FirstIndex
        resultMatch.Length = match.Length
        resultMatch.value = match.value

        Dim submatch As Variant
        For Each submatch In match.SubMatches
            resultMatch.SubMatchCollection.Add submatch
        Next

        result.MatchCollection.Add resultMatch
    Next
#End If

    Set RegEx_Execute = result
End Function

Public Function RegEx_Test(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False, Optional ByVal flagMultiLine As Boolean = True) As Boolean
#If Mac Then
    ' TODO: RegEx implementation for macOS
    Err.Raise 5, Description:="RegEx_Execute not yet implemented on macOS."
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")

    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    regexpObject_.multiLine = flagMultiLine

    RegEx_Test = regexpObject_.test(test)
#End If
End Function

Public Function RegEx_Replace(ByVal test As String, ByVal pattern As String, ByVal replace As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False, Optional ByVal flagMultiLine As Boolean = True) As String
#If Mac Then
    ' TODO: RegEx implementation for macOS
    Err.Raise 5, Description:="RegEx_Execute not yet implemented on macOS."
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")

    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    regexpObject_.multiLine = flagMultiLine

    RegEx_Replace = regexpObject_.replace(test, replace)
#End If
End Function
