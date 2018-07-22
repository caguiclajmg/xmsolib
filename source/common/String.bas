Attribute VB_Name = "common_String"
Option Explicit

Public Function String_Contains(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    String_Contains = (InStr(1, value, match, compareMethod) <> 0)
End Function

Public Function String_StartsWith(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbTextCompare) As Boolean
    String_StartsWith = (InStr(1, value, match, compareMethod) = 1)
End Function

Public Function String_EndsWith(ByVal value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    String_EndsWith = ((InStrRev(value, match, -1, compareMethod) + Len(match) - 1) = Len(value))
End Function

Public Function String_TrimStart(ByVal value As String, Optional ByVal match As String = " ", Optional compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_StartsWith(value, match, compareMethod)
        value = Right$(value, Len(value) - match_length)
    Wend
    
    String_TrimStart = value
End Function

Public Function String_Insert(ByVal value As String, ByVal other As String, ByVal position As Long) As String
    String_Insert = Left$(value, position - 1) & other & Right$(value, Len(value) - position + 1)
End Function

Public Function String_TrimEnd(ByVal value As String, Optional ByVal match As String = " ", Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_EndsWith(value, match, compareMethod)
        value = Left$(value, Len(value) - match_length)
    Wend
    
    String_TrimEnd = value
End Function
