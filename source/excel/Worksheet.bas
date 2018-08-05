Attribute VB_Name = "excel_Worksheet"
Option Explicit

Public Function Worksheet_TableExists(ByRef sheet As Worksheet, ByVal index As Variant) As Boolean
    On Error GoTo Err:
    
    Dim table As listObject: Set table = sheet.ListObjects(index)
    
    Worksheet_TableExists = True
    Exit Function
    
Err:
    Worksheet_TableExists = False
End Function

Public Function Worksheet_ChartExists(ByRef sheet As Worksheet, ByVal index As Variant) As Boolean
    On Error GoTo Err:
    
    Dim chart As ChartObject: Set chart = sheet.ChartObjects(index)
    
    Worksheet_ChartExists = True
    Exit Function
    
Err:
    Worksheet_ChartExists = False
End Function

Public Function Worksheet_FindTable(ByRef sheet As Worksheet, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As listObject
    Dim table As listObject
    For Each table In sheet.ListObjects
        If String_StartsWith(table.name, name, compareMethod) Then
            Set Worksheet_FindTable = table
            Exit Function
        End If
    Next
    
    Set Worksheet_FindTable = Nothing
End Function

Public Function Worksheet_FindChart(ByRef sheet As Worksheet, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As ChartObject
    Dim chart As ChartObject
    For Each chart In sheet.ChartObjects
        If String_StartsWith(chart.name, name, compareMethod) Then
            Set Worksheet_FindChart = chart
            Exit Function
        End If
    Next
End Function
