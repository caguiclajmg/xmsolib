Attribute VB_Name = "excel_ListObject"
Option Explicit

Public Function ListObject_ColumnExists(ByRef listObject As listObject, ByVal index As Variant) As Boolean
    On Error GoTo Err:
    
    Dim column As ListColumn: Set column = listObject.ListColumns(index)
    
    ListObject_ColumnExists = True
    Exit Function
    
Err:
    ListObject_ColumnExists = False
End Function

Public Function ListObject_FindColumn(ByRef listObject As listObject, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As ListColumn
    Dim column As ListColumn
    For i = 1 To listObject.ListColumns.Count
        If String_StartsWith(column.name, name, compareMethod) Then
            Set ListObject_FindColumn = column
            Exit Function
        End If
    Next
    
    Set ListObject_FindColumn = Nothing
End Function

Public Function ListObject_FindRow(ByRef listObject As listObject, ParamArray match() As Variant) As ListRow
    Dim offset As Long: offset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long
    For i = 1 To listObject.ListRows.Count
        Dim found As Boolean: found = True
        Dim j As Long
        For j = LBound(match) To UBound(match)
            Dim name As String: name = match(j)(0)
            Dim value As Variant: value = match(j)(1)
            
            Dim column As ListColumn: Set column = listObject.ListColumns(name)
            Dim cell As Variant: cell = column.Range(RowIndex:=i + offset)
            
            If value <> cell Then
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
