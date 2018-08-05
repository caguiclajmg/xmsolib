Attribute VB_Name = "excel_Range"
Option Explicit

Public Function Range_Lookup(ByVal lookupColumn As range, ByVal lookupValue As Variant, Optional ByRef returnColumn As range = Nothing) As Variant
    If returnColumn Is Nothing Then Set returnColumn = lookupColumn
    Range_Lookup = returnColumn(rowIndex:=Range_FindInColumn(lookupColumn, lookupValue)).value
End Function

Public Function Range_FindInColumn(ByVal range As range, ByVal value As Variant) As Long
    Dim i As Long
    For i = 1 To range.Rows.count
        If range(rowIndex:=i).value = value Then
            Range_FindInColumn = i
            Exit Function
        End If
    Next
    
    Range_FindInColumn = 0
End Function

Public Function Range_Count(ByRef range As range) As Long
    Range_Count = range.count
End Function

Public Function Range_CountNumber(ByRef range As range) As Long
    Dim count As Long
    
    Dim cell As range
    For Each cell In range
        If String_IsNumber(cell.value) Then count = count + 1
    Next
    
    Range_CountNumber = count
End Function

Public Function Range_CountBlank(ByRef range As range) As Long
    Dim count As Long
    
    Dim cell As range
    For Each cell In range
        If String_IsNull(cell.value) Then count = count + 1
    Next
    
    Range_CountBlank = count
End Function

Public Function Range_Sum(ByRef range As range) As Double
    Dim result As Double
    
    Dim cell As range
    For Each cell In range
        If String_IsNumber(cell) Then result = result + CDbl(cell)
    Next
    
    Range_Sum = result
End Function

Public Function Range_Average(ByRef range As range) As Double
    Dim result As Double, count As Long
    
    Dim cell As range
    For Each cell In range
        If String_IsNumber(cell) Then
            result = result + CDbl(cell)
            count = count + 1
        End If
    Next
    
    Range_Average = result / count
End Function
