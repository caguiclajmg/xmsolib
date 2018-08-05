Attribute VB_Name = "excel_UDF"
Option Explicit

Public Function UDF_Ifs(ParamArray pairs() As Variant) As Variant
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        If CBool(pairs(i)) Then
            Function_IFS = pairs(i + 1)
            Exit Function
        End If
    Next
End Function
