Attribute VB_Name = "excel_UDF"
Option Explicit

Public Function UDF_Ifs(ParamArray pairs() As Variant) As Variant
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        If CBool(pairs(i)) Then
            UDF_Ifs = pairs(i + 1)
            Exit Function
        End If
    Next
    
    UDF_Ifs = CVErr(xlValue)
End Function

Public Function UDF_Match(ByVal range As range, ByVal value As Variant) As Variant
    Dim index As Long: index = Range_Match(range, value)
    
    If index = -1 Then
        UDF_Match = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Match = index
End Function

Public Function UDF_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    Dim value As Variant: value = Range_Lookup(lookupRange, lookupValue, returnRange)
    
    If IsNull(value) Then
        UDF_Lookup = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Lookup = value
End Function
