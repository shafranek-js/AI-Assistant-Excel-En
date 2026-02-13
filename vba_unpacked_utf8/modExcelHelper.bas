Attribute VB_Name = "modExcelHelper"
'========================================
' Excel helper module
'========================================

Option Explicit

'----------------------------------------
' Get Excel version label
'----------------------------------------
Private Function GetExcelVersionName() As String
    Dim ver As String
    ver = Application.Version
    
    Select Case ver
        Case "12.0"
            GetExcelVersionName = "Excel 2007 - Use only features available in Excel 2007"
        Case "14.0"
            GetExcelVersionName = "Excel 2010 - use only features available in Excel 2010"
        Case "15.0"
            GetExcelVersionName = "Excel 2013 - use only features available in Excel 2013"
        Case "16.0"
            GetExcelVersionName = "Excel 2016/2019/365"
        Case Else
            GetExcelVersionName = "Excel " & ver
    End Select
End Function

'----------------------------------------
' Get current workbook context
'----------------------------------------
Public Function GetWorkbookContext() As String
    On Error Resume Next
    
    Dim context As String
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim headers As String
    Dim sampleData As String
    Dim i As Long, j As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    ' Adding an Excel version
    context = "Excel Version: " & GetExcelVersionName() & vbCrLf
    context = context & "Active sheet: " & ActiveSheet.Name & vbCrLf
    context = context & "Workbook sheets: "
    
    For Each ws In ActiveWorkbook.Worksheets
        context = context & ws.Name & ", "
    Next ws
    context = Left(context, Len(context) - 2) & vbCrLf
    
    ' Range used
    Set usedRng = ActiveSheet.UsedRange
    If Not usedRng Is Nothing Then
        ' Get exact coordinates
        firstRow = usedRng.Row
        firstCol = usedRng.Column
        lastRow = firstRow + usedRng.Rows.Count - 1
        lastCol = firstCol + usedRng.columns.Count - 1
        
        context = context & "Used range: " & usedRng.address & vbCrLf
        context = context & "Starts at row " & firstRow & ", column " & ColLetter(firstCol) & vbCrLf
        context = context & "Ends at row " & lastRow & ", column " & ColLetter(lastCol) & vbCrLf
        context = context & "Total rows: " & usedRng.Rows.Count & ", columns: " & usedRng.columns.Count & vbCrLf
        
        ' Headers (first line of data with their addresses)
        If usedRng.Rows.Count > 0 Then
            context = context & vbCrLf & "Data structure (header row):" & vbCrLf
            For j = 1 To Application.Min(usedRng.columns.Count, 10)
                Dim cellAddr As String
                cellAddr = ColLetter(firstCol + j - 1) & firstRow
                context = context & "  " & cellAddr & ": " & usedRng.Cells(1, j).value & vbCrLf
            Next j
            
            ' Adding example data (second line)
            If usedRng.Rows.Count > 1 Then
                context = context & vbCrLf & "Sample data (row " & (firstRow + 1) & "):" & vbCrLf
                For j = 1 To Application.Min(usedRng.columns.Count, 10)
                    cellAddr = ColLetter(firstCol + j - 1) & (firstRow + 1)
                    context = context & "  " & cellAddr & ": " & usedRng.Cells(2, j).value & vbCrLf
                Next j
            End If
        End If
    End If
    
    GetWorkbookContext = context
End Function

'----------------------------------------
' Convert column number to letter
'----------------------------------------
Private Function ColLetter(colNum As Long) As String
    Dim result As String
    Dim n As Long
    
    n = colNum
    result = ""
    
    Do While n > 0
        result = Chr(((n - 1) Mod 26) + 65) & result
        n = (n - 1) \ 26
    Loop
    
    ColLetter = result
End Function

'----------------------------------------
' Get selected data
'----------------------------------------
Public Function GetSelectedData() As String
    On Error Resume Next
    
    Dim sel As Range
    Dim result As String
    Dim i As Long, j As Long
    Dim maxRows As Long
    Dim maxCols As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set sel = Selection
    
    If sel Is Nothing Then
        GetSelectedData = ""
        Exit Function
    End If
    
    If TypeName(sel) <> "Range" Then
        GetSelectedData = ""
        Exit Function
    End If
    
    ' Get exact coordinates
    firstRow = sel.Row
    firstCol = sel.Column
    lastRow = firstRow + sel.Rows.Count - 1
    lastCol = firstCol + sel.columns.Count - 1
    
    result = "=== SELECTED DATA ===" & vbCrLf
    result = result & "Range: " & sel.address & vbCrLf
    result = result & "First cell: " & ColLetter(firstCol) & firstRow & vbCrLf
    result = result & "Last cell: " & ColLetter(lastCol) & lastRow & vbCrLf
    result = result & "Size: " & sel.Rows.Count & " rows x " & sel.columns.Count & " columns" & vbCrLf & vbCrLf
    
    ' Limit output size
    maxRows = Application.Min(sel.Rows.Count, 30)
    maxCols = Application.Min(sel.columns.Count, 10)
    
    ' Output data with row addresses
    result = result & "Data (with cell addresses):" & vbCrLf
    
    ' Table header with column letters
    result = result & "Row" & vbTab
    For j = 1 To maxCols
        result = result & ColLetter(firstCol + j - 1) & vbTab
    Next j
    result = result & vbCrLf
    
    ' Data
    For i = 1 To maxRows
        result = result & (firstRow + i - 1) & vbTab
        For j = 1 To maxCols
            result = result & sel.Cells(i, j).value
            If j < maxCols Then result = result & vbTab
        Next j
        result = result & vbCrLf
    Next i
    
    If sel.Rows.Count > maxRows Then
        result = result & "... and " & (sel.Rows.Count - maxRows) & " more rows" & vbCrLf
    End If
    
    ' Hint for AI
    result = result & vbCrLf & "IMPORTANT: Use the EXACT cell addresses from the data above!" & vbCrLf
    result = result & "First data row: " & firstRow & ", last: " & lastRow & vbCrLf
    result = result & "First column: " & ColLetter(firstCol) & ", last: " & ColLetter(lastCol) & vbCrLf
    
    GetSelectedData = result
End Function

'----------------------------------------
' Setting a value to a cell
'----------------------------------------
Public Sub SetCellValue(address As String, value As Variant)
    On Error Resume Next
    ActiveSheet.Range(address).value = value
End Sub

'----------------------------------------
' Setting a formula to a cell
'----------------------------------------
Public Sub SetCellFormula(address As String, formula As String)
    On Error Resume Next
    ActiveSheet.Range(address).formula = formula
End Sub

'----------------------------------------
' Set values in range
'----------------------------------------
Public Sub SetRangeValues(address As String, values As Variant)
    On Error Resume Next
    ActiveSheet.Range(address).value = values
End Sub

'----------------------------------------
' Format range
'----------------------------------------
Public Sub FormatRange(address As String, formatType As String, formatValue As String)
    On Error Resume Next
    
    Dim rng As Range
    Set rng = ActiveSheet.Range(address)
    
    Select Case formatType
        Case "numberformat"
            rng.NumberFormat = formatValue
        Case "bold"
            rng.Font.Bold = (formatValue = "true")
        Case "italic"
            rng.Font.Italic = (formatValue = "true")
        Case "fontcolor"
            rng.Font.Color = CLng(formatValue)
        Case "fillcolor"
            rng.Interior.Color = CLng(formatValue)
        Case "fontsize"
            rng.Font.Size = CInt(formatValue)
        Case "align"
            Select Case formatValue
                Case "left": rng.HorizontalAlignment = xlLeft
                Case "center": rng.HorizontalAlignment = xlCenter
                Case "right": rng.HorizontalAlignment = xlRight
            End Select
    End Select
End Sub

'----------------------------------------
' Auto-fit column widths
'----------------------------------------
Public Sub AutoFitColumns(Optional address As String = "")
    On Error Resume Next
    
    If Len(address) > 0 Then
        ActiveSheet.Range(address).columns.AutoFit
    Else
        ActiveSheet.UsedRange.columns.AutoFit
    End If
End Sub

'----------------------------------------
' Sort range
'----------------------------------------
Public Sub SortRange(address As String, columnIndex As Long, ascending As Boolean)
    On Error Resume Next
    
    Dim rng As Range
    Set rng = ActiveSheet.Range(address)
    
    rng.Sort Key1:=rng.columns(columnIndex), _
             Order1:=IIf(ascending, xlAscending, xlDescending), _
             Header:=xlYes
End Sub

'----------------------------------------
' Create table
'----------------------------------------
Public Sub CreateTable(address As String, tableName As String)
    On Error Resume Next
    
    Dim rng As Range
    Dim tbl As ListObject
    
    Set rng = ActiveSheet.Range(address)
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    
    If Len(tableName) > 0 Then
        tbl.Name = tableName
    End If
    
    tbl.TableStyle = "TableStyleMedium2"
End Sub

'----------------------------------------
' Remove duplicates
'----------------------------------------
Public Sub RemoveDuplicates(address As String, Optional columns As String = "")
    On Error Resume Next
    
    Dim rng As Range
    Set rng = ActiveSheet.Range(address)
    
    If Len(columns) = 0 Then
        rng.RemoveDuplicates columns:=1, Header:=xlYes
    Else
        ' Parsing columns from the string "1,2,3"
        Dim colArr() As String
        Dim colNums() As Long
        Dim i As Long
        
        colArr = Split(columns, ",")
        ReDim colNums(UBound(colArr))
        
        For i = 0 To UBound(colArr)
            colNums(i) = CLng(Trim(colArr(i)))
        Next i
        
        rng.RemoveDuplicates columns:=colNums, Header:=xlYes
    End If
End Sub

'----------------------------------------
' Find and replace
'----------------------------------------
Public Sub FindAndReplace(findText As String, replaceText As String, Optional address As String = "")
    On Error Resume Next
    
    Dim rng As Range
    
    If Len(address) > 0 Then
        Set rng = ActiveSheet.Range(address)
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    
    rng.Replace What:=findText, Replacement:=replaceText, LookAt:=xlPart
End Sub

