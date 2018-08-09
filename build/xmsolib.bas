Attribute VB_Name = "xmsolib"
Option Explicit



Public Function Array_Count(ByVal arr As Variant) As Long
    Array_Count = UBound(arr) - LBound(arr) + 1
End Function

Public Function Array_Equals(ByVal arr As Variant, ByVal other As Variant) As Boolean
    If Array_Count(arr) <> Array_Count(other) Then
        Array_Equals = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> other(i) Then
            Array_Equals = False
            Exit Function
        End If
    Next
    
    Array_Equals = True
End Function

Public Function Array_Contains(ByVal arr As Variant, ByVal match As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = match Then
            Array_Contains = True
            Exit Function
        End If
    Next
    
    Array_Contains = False
End Function
#If Mac Then
Public Const DIRECTORY_SEPARATOR As String = "/"
#Else
Public Const DIRECTORY_SEPARATOR As String = "\"
#End If

Public Function FileSystem_StripExtension(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, ".")
    
    FileSystem_StripExtension = IIf(position = 0, path, Left$(path, position - 1))
End Function

Public Function FileSystem_EnumerateFiles(ByVal path As String, Optional ByVal match As String = "*", Optional ByVal flags As VbFileAttribute = vbNormal) As String()
    If Right$(path, 1) <> "\" Then path = path & DIRECTORY_SEPARATOR
    
    Dim count As Long, filename As String
    
    filename = Dir$(path & match, flags)
    While filename <> vbNullString
        If (filename <> ".") And (filename <> "..") Then count = count + 1
        
        filename = Dir$()
    Wend
    
    If count = 0 Then Exit Function
    
    Dim Index As Long: Index = 1
    ReDim result(1 To count) As String
    
    filename = Dir$(path & match, flags)
    While filename <> vbNullString
        If (filename <> ".") And (filename <> "..") Then
            result(Index) = filename
            Index = Index + 1
        End If
        
        filename = Dir$()
    Wend
    
    FileSystem_EnumerateFiles = result
End Function

Public Function FileSystem_StripPath(ByVal path As String) As String
    Dim position As Long: position = InStrRev(path, "\")
    
    FileSystem_StripPath = IIf(position = 0, path, Right$(path, Len(path) - position))
End Function

Public Function FileSystem_FolderExists(ByVal path As String) As Boolean
    On Error GoTo Error:
    
    FileSystem_FolderExists = (GetAttr(path) And vbDirectory) = vbDirectory
    Exit Function
    
Error:
    FileSystem_FolderExists = False
End Function

Public Function FileSystem_FileExists(ByVal path As String) As Boolean
    On Error GoTo Error:
    
    FileSystem_FileExists = (GetAttr(path) And vbDirectory) <> vbDirectory
    Exit Function
    
Error:
    FileSystem_FileExists = False
End Function

Public Function Math_Fibonacci(ByVal n As Long) As Double
    If n = 0 Then
        Math_Fibonacci = 0#
        Exit Function
    End If
    
    If n = 1 Then
        Math_Fibonacci = 1#
        Exit Function
    End If
    
    Dim previous As Double: previous = 0#
    Dim current As Double: current = 1#
    
    Dim i As Long
    For i = 2 To n
        Dim newCurrent As Double: newCurrent = previous + current
        previous = current
        current = newCurrent
    Next
    
    Math_Fibonacci = current
End Function

Public Function String_IsNullOrWhitespace(ByVal Value As String) As Boolean
    String_IsNullOrWhitespace = String_IsNull(String_Trim(Value))
End Function

Public Function String_IsNull(ByVal Value As String) As Boolean
    String_IsNull = (Value = vbNullString)
End Function

Public Function String_IsNumber(ByVal Value As String) As Boolean
    On Error GoTo Error:
    
    Dim number As Double: number = CDbl(Value)
    
    String_IsNumber = True
    Exit Function
    
Error:
    String_IsNumber = False
End Function

Public Function String_Contains(ByVal Value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    String_Contains = (InStr(1, Value, match, compareMethod) <> 0)
End Function

Public Function String_StartsWith(ByVal Value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbTextCompare) As Boolean
    If String_IsNull(Value) Then
        String_StartsWith = False
        Exit Function
    End If
    
    String_StartsWith = (InStr(1, Value, match, compareMethod) = 1)
End Function

Public Function String_EndsWith(ByVal Value As String, ByVal match As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Boolean
    If String_IsNull(Value) Then
        String_EndsWith = False
        Exit Function
    End If
    
    String_EndsWith = ((InStrRev(Value, match, -1, compareMethod) + Len(match) - 1) = Len(Value))
End Function

Public Function String_Insert(ByVal Value As String, ByVal other As String, ByVal position As Long) As String
    String_Insert = Left$(Value, position - 1) & other & Right$(Value, Len(Value) - position + 1)
End Function

Public Function String_TrimStart(ByVal Value As String, Optional ByVal match As String = " ", Optional compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_StartsWith(Value, match, compareMethod)
        Value = Right$(Value, Len(Value) - match_length)
    Wend
    
    String_TrimStart = Value
End Function

Public Function String_TrimEnd(ByVal Value As String, Optional ByVal match As String = " ", Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim match_length As Long: match_length = Len(match)
    
    While String_EndsWith(Value, match, compareMethod)
        Value = Left$(Value, Len(Value) - match_length)
    Wend
    
    String_TrimEnd = Value
End Function

Public Function String_Trim(ByVal Value As String, Optional ByVal match As String = " ", Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As String
    String_Trim = String_TrimStart(String_TrimEnd(Value, match, compareMethod), match, compareMethod)
End Function

Public Function String_Format(ByVal format As String, ParamArray parameters() As Variant) As String
    Dim result As String: result = format
    
    Dim matches() As RegexMatch: matches = RegEx_Execute(result, "\{\d+\}")
    
    Dim i As Long
    For i = LBound(matches) To UBound(matches)
        Dim Index As Long: Index = CLng(Mid$(matches(i).Value, 2, Len(matches(i).Value) - 1))
        matches(i).Value = parameters(i)
    Next
    
    Dim current As Long: current = LBound(matches)
    While result Like RegEx_Test(result, "\{\d+\}")
        result = RegEx_Replace(result, "\{\d+\}", matches(current).Value, flagGlobal:=False)
        current = current + 1
    Wend
    
    String_Format = result
End Function

Public Const BYTE_MIN As Byte = 0
Public Const BYTE_MAX As Byte = 255

Public Const INT_MIN As Integer = -32768
Public Const INT_MAX As Integer = 32767

Public Const LONG_MIN As Long = -2147483648#
Public Const LONG_MAX As Long = 2147483647

Public Const SINGLE_MIN As Single = -3.4028235E+38
Public Const SINGLE_MAX As Single = 3.4028235E+38

Public Const DOUBLE_MIN As Double = -1.79769313486231E+308
Public Const DOUBLE_MAX As Double = 1.79769313486231E+308

Public Function VBComponent_GetCode(ByVal component As VBComponent) As String
    VBComponent_GetCode = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
End Function

Public Function VBComponent_FromString(ByVal project As VBProject, ByVal ctype As vbext_ComponentType, ByVal name As String, ByVal code As String) As VBComponent
    Dim component As VBComponent: Set component = project.VBComponents.Add(ctype)
    component.CodeModule.AddFromString code
    
    Set VBComponent_FromString = component
End Function

Public Function VBComponent_Import(ByVal project As VBProject, ByVal path As String) As VBComponent
    Set VBComponent_Import = project.VBComponents.Import(path)
End Function

Public Sub VBComponent_Export(ByVal project As VBProject, ByVal name As String, ByVal path As String, Optional ByVal filename As String = vbNullString)
    Dim component As VBComponent: Set component = project.VBComponents(name)
    
    Dim extension As String
    Select Case component.Type
        Case vbext_ct_ClassModule, vbext_ct_Document
            extension = "cls"
            
        Case vbext_ct_MSForm
            extension = "frm"
            
        Case vbext_ct_StdModule
            extension = "bas"
            
        Case Else
            extension = vbNullString
    End Select
    
    If Right$(path, 1) <> "\" Then path = path & "\"
    project.VBComponents(name).Export path & IIf(filename = vbNullString, component.name, filename) & IIf(extension = vbNullString, vbNullString, "." & extension)
End Sub

Public Function VBComponent_Exists(ByVal project As VBProject, ByVal name As String) As Boolean
    On Error GoTo Error:
    
    Dim component As VBComponent: Set component = project.VBComponents(name)
    
    VBComponent_Exists = True
    Exit Function
    
Error:
    VBComponent_Exists = False
End Function

Public Function ListObject_InsertColumn(ByVal listObject As listObject, ByVal name As String, Optional ByVal position = 0) As ListColumn
    If position = 0 Then position = listObject.ListColumns.count + 1
    
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns.Add(position)
    columnObject.name = name
    
    Set ListObject_InsertColumn = columnObject
End Function

Public Function ListObject_FillColumn(ByVal column As ListColumn, ParamArray values() As Variant) As ListColumn
    Dim listObject As listObject: Set listObject = column.Parent
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long, rowIndex As Long
    For i = LBound(values) To UBound(values)
        column.range(rowIndex:=rowIndex + rowOffset) = values(i)
        rowIndex = rowIndex + 1
    Next
    
    Set ListObject_FillColumn = column
End Function

Public Function ListObject_FillRow(ByVal row As ListRow, ParamArray values() As Variant) As ListRow
    Dim i As Long, columnIndex As Long
    For i = LBound(values) To UBound(values)
        row.range(columnIndex:=columnIndex) = values(i)
        columnIndex = columnIndex + 1
    Next
    
    Set ListObject_FillRow = row
End Function

Public Function ListObject_FillRowAssociative(ByVal row As ListRow, ParamArray values() As Variant) As ListRow
    Dim listObject As listObject: Set listObject = row.Parent
    Dim rowOffset As Long: rowOffset = IIf(listObject.HeaderRowRange Is Nothing, 0, 1)
    
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim column As String: column = values(i)(0)
        Dim Value As String: Value = values(i)(1)
        
        listObject.ListColumns(column).range(rowIndex:=row.Index + rowOffset) = Value
    Next
    
    Set ListObject_FillRowAssociative = listObject.DataBodyRange(rowIndex:=row.Index)
End Function

Public Sub ListObject_ClearData(ByVal listObject As listObject, Optional ByVal preserveTemplateRow As Boolean = False)
    With listObject.DataBodyRange
        If preserveTemplateRow Then
            .offset(1).Resize(.Rows.count - 1, .Columns.count).Delete
        Else
            .Delete
        End If
    End With
End Sub

Public Function ListObject_ColumnExists(ByVal listObject As listObject, ByVal Index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim columnObject As ListColumn: Set columnObject = listObject.ListColumns(Index)
    
    ListObject_ColumnExists = True
    Exit Function
    
Error:
    ListObject_ColumnExists = False
End Function

Public Function ListObject_FindColumn(ByVal listObject As listObject, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As ListColumn
    Dim columnObject As ListColumn
    For Each columnObject In listObject.ListColumns
        If String_StartsWith(columnObject.name, name, compareMethod) Then
            Set ListObject_FindColumn = columnObject
            Exit Function
        End If
    Next
    
    Set ListObject_FindColumn = Nothing
End Function

Public Function ListObject_FindInColumn(ByVal column As ListColumn, ByVal Value As Variant) As Long
    Dim i As Long
    For i = 1 To column.range.count
        If column.range(rowIndex:=i) = Value Then
            ListObject_FindInColumn = i
            Exit Function
        End If
    Next
    
    ListObject_FindInColumn = -1
End Function

Public Function ListObject_FindRow(ByVal listObject As listObject, ParamArray match() As Variant) As ListRow
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
End Function

Public Function Range_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    On Error GoTo Error:
    
    Dim Index As Long: Index = Range_Match(lookupRange, lookupValue)
    If Index = -1 Then Err.Raise xlReference
    
    Range_Lookup = returnRange(Index).Value
    Exit Function
    
Error:
    Range_Lookup = Null
End Function

Public Function Range_Match(ByVal range As range, ByVal Value As Variant) As Long
    On Error GoTo Error:
    
    Range_Match = CLng(range.Application.WorksheetFunction.match(Value, range, 0))
    Exit Function
    
Error:
    Range_Match = -1
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

Public Function UDF_Ifs(ByVal default As Variant, ParamArray pairs() As Variant) As Variant
    Dim i As Long
    For i = LBound(pairs) To UBound(pairs) Step 2
        If CBool(pairs(i)) Then
            UDF_Ifs = pairs(i + 1)
            Exit Function
        End If
    Next
    
    UDF_Ifs = default
End Function

Public Function UDF_Match(ByVal range As range, ByVal Value As Variant) As Variant
    Dim Index As Long: Index = Range_Match(range, Value)
    
    If Index = -1 Then
        UDF_Match = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Match = Index
End Function

Public Function UDF_Lookup(ByVal lookupRange As range, ByVal lookupValue As Variant, ByVal returnRange As range) As Variant
    Dim Value As Variant: Value = Range_Lookup(lookupRange, lookupValue, returnRange)
    
    If IsNull(Value) Then
        UDF_Lookup = CVErr(xlValue)
        Exit Function
    End If
    
    UDF_Lookup = Value
End Function

Public Function Workbook_WorksheetExists(ByVal book As Workbook, ByVal Index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim sheet As Worksheet: Set sheet = book.Worksheets(Index)
    
    Workbook_WorksheetExists = True
    Exit Function
    
Error:
    Workbook_WorksheetExists = False
End Function

Public Function Workbook_FindWorksheet(ByVal book As Workbook, ByVal name As String, Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) As Worksheet
    Dim sheet As Worksheet
    For Each sheet In book.Worksheets
        If String_StartsWith(sheet.name, name, compareMethod) Then
            Set Workbook_FindWorksheet = sheet
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


Public Function Worksheet_TableExists(ByVal sheet As Worksheet, ByVal Index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim table As listObject: Set table = sheet.ListObjects(Index)
    
    Worksheet_TableExists = True
    Exit Function
    
Error:
    Worksheet_TableExists = False
End Function

Public Function Worksheet_ChartExists(ByVal sheet As Worksheet, ByVal Index As Variant) As Boolean
    On Error GoTo Error:
    
    Dim chart As ChartObject: Set chart = sheet.ChartObjects(Index)
    
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
    Dim chart As ChartObject
    For Each chart In sheet.ChartObjects
        If String_StartsWith(chart.name, name, compareMethod) Then
            Set Worksheet_FindChart = chart
            Exit Function
        End If
    Next
    
    Set Worksheet_FindChart = Nothing
End Function
#If Not Mac Then
Private regexpObject_ As Object
#End If

Public Type RegexMatch
    Index As Long
    submatches() As String
    Value As String
End Type

Public Function RegEx_Execute(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As RegexMatch()
    Dim result() As RegexMatch
    
#If Mac Then
    ' TODO: RegEx implementation for macOS
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")
    
    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    
    Dim matches As Object: Set matches = regexpObject_.Execute(test)
    ReDim result(1 To matches.count) As RegexMatch
    
    Dim i As Long
    For i = LBound(result) To UBound(result)
        Dim match As Object: Set match = matches.Item(i - 1)
        
        result(i).Index = match.FirstIndex
        
        Dim submatches As Object: Set submatches = match.submatches
        
        If submatches.count > 0 Then
            ReDim result(i).submatches(1 To submatches.count)
            
            Dim j As Long
            For j = LBound(result(i).submatches) To UBound(result(i).submatches)
                result(i).submatches(j) = submatches.Item(j - 1)
            Next
        End If
        
        result(i).Value = match.Value
    Next
#End If

    RegEx_Execute = result
End Function

Public Function RegEx_Test(ByVal test As String, ByVal pattern As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As Boolean
#If Mac Then
    ' TODO: RegEx implementation for macOS
    RegEx_Test = True
#Else
    If regexpObject_ Is Nothing Then Set regexpObject_ = CreateObject("VBScript.RegExp")
    
    regexpObject_.pattern = pattern
    regexpObject_.Global = flagGlobal
    regexpObject_.ignoreCase = flagIgnoreCase
    
    RegEx_Test = regexpObject_.test(test)
#End If
End Function

Public Function RegEx_Replace(ByVal test As String, ByVal pattern As String, ByVal replace As String, Optional ByVal flagGlobal As Boolean = True, Optional ByVal flagIgnoreCase As Boolean = False) As String
#If Mac Then
    ' TODO: RegEx implementation for macOS
    RegEx_Replace = test
#Else
    Dim regexp As Object: Set regexp = CreateObject("VBScript.RegExp")
    regexp.pattern = pattern
    regexp.Global = flagGlobal
    regexp.ignoreCase = flagIgnoreCase
    
    RegEx_Replace = regexp.replace(test, replace)
#End If
End Function
