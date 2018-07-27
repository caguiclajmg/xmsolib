Attribute VB_Name = "excel_ListObject"
Option Explicit

Private Sub Test()
    Dim table As listObject: Set table = shtDefault.ListObjects(1)
    Dim row As ListRow
    
    Set row = table.ListRows.Add()
    row.range(rowIndex:=1, columnIndex:=1) = "Template"
    
    Set row = table.ListRows.Add()
    row.range(rowIndex:=1, columnIndex:=1) = "Test 1"
    
    Set row = table.ListRows.Add()
    row.range(rowIndex:=1, columnIndex:=1) = "Test 2"
    
    Set row = table.ListRows.Add()
    row.range(rowIndex:=1, columnIndex:=1) = "Test 3"
    
    ListObject_ClearData table, True
End Sub

Public Function ListObject_InsertColumn(ByRef listObject As listObject, ByVal name As String, Optional ByVal position = 0) As ListColumn
    If position = 0 Then position = listObject.ListColumns.count + 1
    
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns.Add(position)
    columnObject.name = name
    
    Set ListObject_InsertColumn = columnObject
End Function

Public Function ListObject_FillColumn(ByRef column As ListColumn, ParamArray values() As Variant) As ListColumn
    Dim listObject As listObject: Set listObject = column.Parent
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long, rowIndex As Long
    For i = LBound(values) To UBound(values)
        column.range(rowIndex:=rowIndex + rowOffset) = values(i)
        rowIndex = rowIndex + 1
    Next
    
    Set ListObject_FillColumn = column
End Function

Public Function ListObject_FillRow(ByRef row As ListRow, ParamArray values() As Variant) As ListRow
    Dim i As Long, columnIndex As Long
    For i = LBound(values) To UBound(values)
        row.range(columnIndex:=columnIndex) = values(i)
        columnIndex = columnIndex + 1
    Next
    
    Set ListObject_FillRow = row
End Function

Public Function ListObject_FillRowAssociative(ByRef row As ListRow, ParamArray values() As Variant) As ListRow
    Dim listObject As listObject: Set listObject = row.Parent
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim column As String: column = values(i)(0)
        Dim value As String: value = values(i)(1)
        
        listObject.ListColumns(column).range(rowIndex:=row.index + rowOffset) = value
    Next
    
    Set ListObject_FillRowAssociative = listObject.DataBodyRange(rowIndex:=row.index)
End Function

Public Sub ListObject_ClearData(ByRef listObject As listObject, Optional ByVal preserveTemplateRow As Boolean = False)
    With listObject.DataBodyRange
        If preserveTemplateRow Then
            .offset(1).Resize(.Rows.count - 1, .Columns.count).Delete
        Else
            .Delete
        End If
    End With
End Sub

Public Function ListObject_ColumnExists(ByRef listObject As listObject, ByVal index As Variant) As Boolean
    On Error GoTo Err:
    
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns(index)
    
    ListObject_ColumnExists = True
    Exit Function
    
Err:
    ListObject_ColumnExists = False
End Function

Public Function ListObject_FindColumn(ByRef listObject As listObject, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As ListColumn
    Dim columnObject As ListColumn
    For Each columnObject In listObject.ListColumns
        If String_StartsWith(columnObject.name, name, compareMethod) Then
            Set ListObject_FindColumn = columnObject
            Exit Function
        End If
    Next
    
    Set ListObject_FindColumn = Nothing
End Function

Public Function ListObject_FindInColumn(ByRef column As ListColumn, ByVal value As Variant) As Long
    Dim i As Long
    For i = 1 To column.range.count
        If column.range(rowIndex:=i) = value Then
            ListObject_FindInColumn = i
            Exit Function
        End If
    Next
    
    ListObject_FindInColumn = 0
End Function

Public Function ListObject_FindRow(ByRef listObject As listObject, ParamArray match() As Variant) As ListRow
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long
    For i = 1 To listObject.ListRows.count
        Dim found As Boolean: found = True
        Dim j As Long
        For j = LBound(match) To UBound(match)
            Dim columnName As String: columnName = match(j)(0)
            Dim columnValue As Variant: columnValue = match(j)(1)
            
            Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns(columnName)
            Dim cellValue As Variant: cellValue = columnObject.range(rowIndex:=i + rowOffset)
            
            If cellValue <> columnValue Then
                found = False
                Exit For
            End If
        Next
        
        If found Then
            Set ListObject_FindRow = listObject.ListRows(i)
            Exit Function
        End If
    Next
    
    Set ListObject_FindRow = Nothing
    Exit Function
End Function
