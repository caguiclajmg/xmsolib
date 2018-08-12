Attribute VB_Name = "common_String"
Option Explicit

Public Function String_IsNullOrWhitespace(ByVal value As String) As Boolean
    String_IsNullOrWhitespace = String_IsNull(String_Trim(value))
End Function

Public Function String_IsNull(ByVal value As String) As Boolean
    String_IsNull = (value = vbNullString)
End Function

Public Function String_IsNumber(ByVal value As String) As Boolean
    On Error GoTo Error:
    
    Dim number As Double: number = CDbl(value)
    
    String_IsNumber = True
    Exit Function
    
Error:
    String_IsNumber = False
End Function

Public Function String_Contains(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    String_Contains = (InStr(1, value, match, compareMethod) <> 0)
End Function

Public Function String_StartsWith(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbTextCompare) As Boolean
    If String_IsNull(value) Then
        String_StartsWith = False
        Exit Function
    End If
    
    String_StartsWith = (InStr(1, value, match, compareMethod) = 1)
End Function

Public Function String_EndsWith(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If String_IsNull(value) Then
        String_EndsWith = False
        Exit Function
    End If
    
    String_EndsWith = ((InStrRev(value, match, -1, compareMethod) + Len(match) - 1) = Len(value))
End Function

Public Function String_Insert(ByVal value As String, ByVal other As String, ByVal position As Long) As String
    String_Insert = Left$(value, position - 1) & other & Right$(value, Len(value) - position + 1)
End Function

Public Function String_TrimStart(ByVal value As String, Optional ByVal match As String = " ", Optional compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_StartsWith(value, match, compareMethod)
        value = Right$(value, Len(value) - match_length)
    Wend
    
    String_TrimStart = value
End Function

Public Function String_TrimEnd(ByVal value As String, Optional ByVal match As String = " ", Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_EndsWith(value, match, compareMethod)
        value = Left$(value, Len(value) - match_length)
    Wend
    
    String_TrimEnd = value
End Function

Public Function String_Trim(ByVal value As String, Optional ByVal match As String = " ", Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    String_Trim = String_TrimStart(String_TrimEnd(value, match, compareMethod), match, compareMethod)
End Function

Public Function String_Format(ByVal format As String, ParamArray parameters() As Variant) As String
    Dim result As String: result = format
    
    Dim matches As RegexMatchCollection: Set matches = RegEx_Execute(format, "\{(\d+)\}")
    
    Dim match As RegexMatch, offset As Long
    For Each match In matches
        Dim parameterIndex As Long: parameterIndex = CLng(match.SubMatches(1))
        result = Left$(result, match.Index + offset) & parameters(parameterIndex) & Mid$(result, match.Index + offset + match.Length + 1)
        offset = offset + Len(CStr(parameters(parameterIndex))) - match.Length
    Next
    
    String_Format = result
End Function
