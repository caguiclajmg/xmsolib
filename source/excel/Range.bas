Attribute VB_Name = "excel_Range"
Option Explicit

Public Function Range_Lookup(ByVal lookupColumn As range, ByVal lookupValue As Variant, Optional ByVal returnColumn As range = Nothing) As Variant
    If returnColumn Is Nothing Then Set returnColumn = lookupColumn
    Range_Lookup = returnColumn(rowIndex:=Range_FindInColumn(lookupColumn, lookupValue)).value
End Function

Private Sub Test()
    Debug.Print Range_FindInColumn(shtDefault.range("J8:J12"), "QQQ")
End Sub

Public Function Range_FindInColumn(ByVal range As range, ByVal value As Variant) As Long
    On Error GoTo Error:
    
    Range_FindInColumn = CLng(range.Application.WorksheetFunction.match(value, range, 0))
    Exit Function
    
Error:
    Range_FindInColumn = 0
End Function

Public Function Range_Count(ByVal range As range) As Long
    Range_Count = range.count
End Function

Public Function Range_CountNumber(ByVal range As range) As Long
    Dim count As Long
    
    Dim cell As range
    For Each cell In range
        If String_IsNumber(cell.value) Then count = count + 1
    Next
    
    Range_CountNumber = count
End Function

Public Function Range_CountBlank(ByVal range As range) As Long
    Dim count As Long
    
    Dim cell As range
    For Each cell In range
        If String_IsNull(cell.value) Then count = count + 1
    Next
    
    Range_CountBlank = count
End Function

Public Function Range_Sum(ByVal range As range) As Double
    Dim result As Double
    
    Dim cell As range
    For Each cell In range
        If String_IsNumber(cell) Then result = result + CDbl(cell)
    Next
    
    Range_Sum = result
End Function

Public Function Range_Average(ByVal range As range) As Double
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
