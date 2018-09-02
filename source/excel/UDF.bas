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

Public Function UDF_Match(ByVal range As range, ByVal Value As Variant) As Variant
    Dim index As Long: index = Range_Match(range, Value)
    
    If index = -1 Then
        UDF_Match = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Match = index
End Function

Public Function UDF_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    Dim Value As Variant: Value = Range_Lookup(lookupRange, lookupValue, returnRange)
    
    If IsNull(Value) Then
        UDF_Lookup = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Lookup = Value
End Function
