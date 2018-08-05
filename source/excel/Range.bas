Attribute VB_Name = "excel_Range"
Option Explicit

Public Function Range_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    Range_Lookup = returnRange(lookupRange.Application.WorksheetFunction.match(lookupValue, lookupRange, True))
End Function

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
    Range_CountNumber = CDbl(range.Application.WorksheetFunction.count(range))
End Function

Public Function Range_CountBlank(ByVal range As range) As Long
    Range_CountBlank = CLng(range.Application.WorksheetFunction.CountBlank(range))
End Function

Public Function Range_Sum(ByVal range As range) As Double
    Range_Sum = range.Application.WorksheetFunction.Sum(range)
End Function

Public Function Range_Average(ByVal range As range) As Double
    Range_Average = range.Application.WorksheetFunction.Average(range)
End Function
