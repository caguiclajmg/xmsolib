Attribute VB_Name = "excel_Worksheet"
Option Explicit

Public Function Worksheet_TableExists(ByVal sheet As Worksheet, ByVal index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim table As listObject: Set table = sheet.ListObjects(index)
    
    Worksheet_TableExists = True
    Exit Function
    
Error:
    Worksheet_TableExists = False
End Function

Public Function Worksheet_ChartExists(ByVal sheet As Worksheet, ByVal index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim Chart As ChartObject: Set Chart = sheet.ChartObjects(index)
    
    Worksheet_ChartExists = True
    Exit Function
    
Error:
    Worksheet_ChartExists = False
End Function

Public Function Worksheet_FindTable(ByVal sheet As Worksheet, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As listObject
    Dim table As listObject
    For Each table In sheet.ListObjects
        If String_StartsWith(table.name, name, compareMethod) Then
            Set Worksheet_FindTable = table
            Exit Function
        End If
    Next
    
    Set Worksheet_FindTable = Nothing
End Function

Public Function Worksheet_FindChart(ByVal sheet As Worksheet, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As ChartObject
    Dim Chart As ChartObject
    For Each Chart In sheet.ChartObjects
        If String_StartsWith(Chart.name, name, compareMethod) Then
            Set Worksheet_FindChart = Chart
            Exit Function
        End If
    Next
    
    Set Worksheet_FindChart = Nothing
End Function
