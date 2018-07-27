Attribute VB_Name = "excel_ListObject"
Option Explicit

Private Sub Test()
    Dim table As listObject: Set table = shtDefault.ListObjects(1)
    Dim row As ListRow
    
    Set row = table.ListRows.Add()
    row.Range(rowIndex:=1, columnIndex:=1) = "Template"
    
    Set row = table.ListRows.Add()
    row.Range(rowIndex:=1, columnIndex:=1) = "Test 1"
    
    Set row = table.ListRows.Add()
    row.Range(rowIndex:=1, columnIndex:=1) = "Test 2"
    
    Set row = table.ListRows.Add()
    row.Range(rowIndex:=1, columnIndex:=1) = "Test 3"
    
    ListObject_ClearData table, True
End Sub

Public Function ListObject_InsertColumn(ByRef listObject As listObject, ByVal name As String, Optional ByVal position = 0) As ListColumn
    If position = 0 Then position = listObject.ListColumns.count + 1
    
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns.Add(position)
    columnObject.name = name
    
    Set ListObject_InsertColumn = columnObject
End Function

Public Function ListObject_FillColumn(ByRef listObject As listObject, ByVal index As Variant, ParamArray values() As Variant) As ListColumn
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns(index)
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long, rowIndex As Long
    For i = LBound(values) To UBound(values)
        columnObject.Range(rowIndex:=rowIndex + rowOffset) = values(i)
        rowIndex = rowIndex + 1
    Next
    
    Set ListObject_FillColumn = columnObject
End Function

Public Function ListObject_FillRow(ByRef listObject As listObject, ByVal index As Long, ParamArray values() As Variant) As ListRow
    Dim rowObject As ListRow: Set rowObject = listObject.ListRows(index)
    
    Dim i As Long, columnIndex As Long
    For i = LBound(values) To UBound(values)
        rowObject.Range(columnIndex:=columnIndex) = values(i)
        columnIndex = columnIndex + 1
    Next
    
    Set ListObject_FillRow = rowObject
End Function

Public Function ListObject_FillRowAssociative(ByRef listObject As listObject, ByVal index As Long, ParamArray values() As Variant) As ListRow
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim column As String: column = values(i)(0)
        Dim value As String: value = values(i)(1)
        
        listObject.ListColumns(column).Range(rowIndex:=index + rowOffset) = value
    Next
    
    Set ListObject_FillRowAssociative = listObject.DataBodyRange(rowIndex:=index)
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
            Dim cellValue As Variant: cellValue = columnObject.Range(rowIndex:=i + rowOffset)
            
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
