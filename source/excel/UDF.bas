Attribute VB_Name = "excel_UDF"
Option Explicit

Public Function UDF_Ifs(ByVal default As Variant, ParamArray pairs() As Variant) As Variant
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        If CBool(pairs(i)) Then
            UDF_Ifs = pairs(i + 1)
            Exit Function
        End If
    Next
    
    UDF_Ifs = default
End Function

Public Function UDF_Match(ByVal range As range, ByVal value As Variant) As Variant
    Dim Index As Long: Index = Range_Match(range, value)
    
    If Index = -1 Then
        UDF_Match = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Match = Index
End Function

Public Function UDF_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    Dim value As Variant: value = Range_Lookup(lookupRange, lookupValue, returnRange)
    
    If IsNull(value) Then
        UDF_Lookup = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Lookup = value
End Function
