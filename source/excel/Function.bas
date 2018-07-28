Attribute VB_Name = "excel_Function"
Option Explicit

Public Function Function_IFS(ParamArray pairs() As Variant) As Variant
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        If CBool(pairs(i)) Then
            Function_IFS = pairs(i + 1)
            Exit Function
        End If
    Next
End Function
