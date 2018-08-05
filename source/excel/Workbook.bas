Attribute VB_Name = "excel_Workbook"
Option Explicit

Public Function Workbook_WorksheetExists(ByVal book As Workbook, ByVal index As Variant) As Boolean
    On Error GoTo Err:
    
    Dim sheet As Worksheet: Set sheet = book.Worksheets(index)
    
    Workbook_WorksheetExists = True
    Exit Function
    
Err:
    Workbook_WorksheetExists = False
End Function

Public Function Workbook_FindWorksheet(ByVal book As Workbook, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        If String_StartsWith(sheet.name, name, compareMethod) Then
            Workbook_FindWorksheet = sheet
            Exit Function
        End If
    Next
    
    Set Workbook_FindWorksheet = Nothing
End Function

Public Function Workbook_FindTable(ByVal book As Workbook, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As listObject
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        Dim list As listObject: Set list = Worksheet_FindTable(sheet, name, compareMethod)
        If Not list Is Nothing Then
            Set Workbook_FindTable = list
            Exit Function
        End If
    Next
    
    Set Workbook_FindTable = Nothing
End Function

