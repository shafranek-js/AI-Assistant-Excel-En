Attribute VB_Name = "CommandParser"
Option Explicit

Private Function ValidateArgCount(action As String, actualArgs As Long, minArgs As Long, maxArgs As Long, ByRef reason As String) As Boolean
    If actualArgs < minArgs Then
        reason = action & " expects at least " & minArgs & " args, got " & actualArgs
        ValidateArgCount = False
        Exit Function
    End If
    If maxArgs >= 0 And actualArgs > maxArgs Then
        reason = action & " expects at most " & maxArgs & " args, got " & actualArgs
        ValidateArgCount = False
        Exit Function
    End If
    ValidateArgCount = True
End Function

Private Function IsValidRangeRef(ws As Worksheet, rangeRef As String) As Boolean
    On Error GoTo InvalidRef
    Dim r As Range
    If Len(Trim$(rangeRef)) = 0 Then Exit Function
    Set r = ws.Range(Trim$(rangeRef))
    IsValidRangeRef = Not r Is Nothing
    Exit Function
InvalidRef:
    IsValidRangeRef = False
End Function

Private Function IsValidApplicationRangeRef(rangeRef As String, requireSingleCell As Boolean) As Boolean
    On Error GoTo InvalidRef
    Dim r As Range
    If Len(Trim$(rangeRef)) = 0 Then Exit Function
    Set r = Application.Range(Trim$(rangeRef))
    If r Is Nothing Then Exit Function
    If requireSingleCell Then
        IsValidApplicationRangeRef = (r.Cells.CountLarge = 1)
    Else
        IsValidApplicationRangeRef = True
    End If
    Exit Function
InvalidRef:
    IsValidApplicationRangeRef = False
End Function

Private Function IsValidSingleCellRef(ws As Worksheet, cellRef As String) As Boolean
    On Error GoTo InvalidRef
    Dim r As Range
    If Len(Trim$(cellRef)) = 0 Then Exit Function
    Set r = ws.Range(Trim$(cellRef))
    If r Is Nothing Then Exit Function
    IsValidSingleCellRef = (r.Cells.CountLarge = 1)
    Exit Function
InvalidRef:
    IsValidSingleCellRef = False
End Function

Private Function IsValidMultiRangeRef(ws As Worksheet, rangeList As String) As Boolean
    Dim items() As String
    Dim i As Long
    
    items = Split(rangeList, ",")
    For i = LBound(items) To UBound(items)
        If Not IsValidRangeRef(ws, Trim$(items(i))) Then
            IsValidMultiRangeRef = False
            Exit Function
        End If
    Next i
    IsValidMultiRangeRef = True
End Function

Private Function IsValidRowRef(ws As Worksheet, rowText As String) As Boolean
    Dim n As Long
    If Not TryParseLong(rowText, n) Then Exit Function
    IsValidRowRef = (n >= 1 And n <= ws.Rows.Count)
End Function

Private Function IsValidColumnRef(ws As Worksheet, colText As String) As Boolean
    Dim idx As Long
    IsValidColumnRef = TryGetColumnIndex(ws, colText, idx)
End Function

Private Function IsValidColumnSelector(ws As Worksheet, selector As String) As Boolean
    Dim n As Long
    If TryParseLong(selector, n) Then
        IsValidColumnSelector = (n >= 1 And n <= ws.Columns.Count)
    Else
        IsValidColumnSelector = IsValidColumnRef(ws, selector)
    End If
End Function

Private Function TryGetColumnIndex(ws As Worksheet, colText As String, ByRef outIdx As Long) As Boolean
    On Error GoTo InvalidCol
    Dim c As String
    c = Trim$(colText)
    If Len(c) = 0 Then Exit Function
    outIdx = ws.Columns(c).Column
    TryGetColumnIndex = (outIdx >= 1 And outIdx <= ws.Columns.Count)
    Exit Function
InvalidCol:
    TryGetColumnIndex = False
End Function

Private Function TryParseLong(valueText As String, ByRef outVal As Long) As Boolean
    On Error GoTo ParseErr
    If Len(Trim$(valueText)) = 0 Then Exit Function
    outVal = CLng(Trim$(valueText))
    TryParseLong = True
    Exit Function
ParseErr:
    TryParseLong = False
End Function

Private Function TryParseDouble(valueText As String, ByRef outVal As Double) As Boolean
    On Error GoTo ParseErr
    If Len(Trim$(valueText)) = 0 Then Exit Function
    outVal = CDbl(Trim$(valueText))
    TryParseDouble = True
    Exit Function
ParseErr:
    TryParseDouble = False
End Function

Private Function IsValidColorSpec(colorText As String) As Boolean
    Dim c As String
    Dim rgbParts() As String
    Dim i As Long
    Dim n As Long
    
    c = UCase$(Trim$(colorText))
    If Len(c) = 0 Then Exit Function
    
    If Left$(c, 4) = "RGB:" Then
        rgbParts = Split(Mid$(c, 5), ",")
        If UBound(rgbParts) <> 2 Then Exit Function
        For i = 0 To 2
            If Not TryParseLong(Trim$(rgbParts(i)), n) Then Exit Function
            If n < 0 Or n > 255 Then Exit Function
        Next i
        IsValidColorSpec = True
        Exit Function
    End If
    
    Select Case c
        Case "RED", "GREEN", "BLUE", "YELLOW", "ORANGE", "PURPLE", "PINK", "CYAN", _
             "WHITE", "BLACK", "GRAY", "GREY", "LIGHTGRAY", "LIGHTGREY", "DARKGRAY", _
             "DARKGREY", "BROWN", "LIME", "NAVY", "TEAL", "MAROON", "OLIVE", "GOLD", "SILVER"
            IsValidColorSpec = True
        Case Else
            IsValidColorSpec = False
    End Select
End Function

Private Function IsValidSortOrder(orderText As String) As Boolean
    Select Case UCase$(Trim$(orderText))
        Case "ASC", "DESC"
            IsValidSortOrder = True
        Case Else
            IsValidSortOrder = False
    End Select
End Function

Private Function IsValidChartTypeName(chartType As String) As Boolean
    Select Case UCase$(Trim$(chartType))
        Case "LINE", "BAR", "COLUMN", "PIE", "AREA", "SCATTER", "XY", "DOUGHNUT", _
             "RADAR", "SURFACE", "BUBBLE", "STOCK", "CYLINDER", "CONE", "PYRAMID", _
             "LINE_MARKERS", "AREA_STACKED", "BAR_STACKED", "COLUMN_STACKED"
            IsValidChartTypeName = True
        Case Else
            IsValidChartTypeName = False
    End Select
End Function

Private Function IsValidChartIndexToken(indexText As String) As Boolean
    Dim n As Long
    Dim s As String
    s = UCase$(Trim$(indexText))
    If s = "" Or s = "0" Or s = "LAST" Or s = "NEW" Then
        IsValidChartIndexToken = True
        Exit Function
    End If
    If TryParseLong(s, n) Then
        IsValidChartIndexToken = (n > 0)
    End If
End Function

Private Function IsValidBooleanToken(valueText As String) As Boolean
    Select Case UCase$(Trim$(valueText))
        Case "TRUE", "FALSE", "1", "0", "YES", "NO"
            IsValidBooleanToken = True
        Case Else
            IsValidBooleanToken = False
    End Select
End Function

Public Function ExtractCommands(response As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim commands As String
    
    startPos = InStr(response, "```commands")
    If startPos = 0 Then
        ExtractCommands = ""
        Exit Function
    End If
    
    startPos = startPos + Len("```commands") + 1
    endPos = InStr(startPos, response, "```")
    
    If endPos = 0 Then
        ExtractCommands = ""
        Exit Function
    End If
    
    commands = Trim(Mid(response, startPos, endPos - startPos))
    ExtractCommands = commands
End Function

Public Function ValidateCommandStrict(cmd As String, ByRef reason As String) As Boolean
    Dim parts() As String
    Dim action As String
    Dim argCount As Long
    Dim ws As Worksheet
    Dim n As Long
    Dim d As Double
    
    reason = ""
    ValidateCommandStrict = False
    
    parts = Split(cmd, "|")
    On Error Resume Next
    argCount = UBound(parts)
    On Error GoTo 0
    
    If argCount < 0 Then
        reason = "Malformed command"
        Exit Function
    End If
    
    action = UCase$(Trim$(parts(0)))
    If Len(action) = 0 Then
        reason = "Missing action"
        Exit Function
    End If
    
    Set ws = ActiveSheet
    If ws Is Nothing Then
        reason = "No active worksheet"
        Exit Function
    End If
    
    Select Case action
        Case "SHOW_ALL_COMMENTS", "HIDE_ALL_COMMENTS", "FREEZE_TOP_ROW", "FREEZE_FIRST_COLUMN", _
             "UNFREEZE_PANES", "CLEAR_PRINT_AREA", "DELETE_PICTURES", "DELETE_SHAPES", _
             "CALCULATE", "CALCULATE_SHEET", "REMOVE_SUBTOTALS", "CHART_DELETE_ALL", _
             "PIVOT_REFRESH_ALL", "REMOVE_AUTOFILTER", "MOVE_CHART"
            If Not ValidateArgCount(action, argCount, 0, 0, reason) Then Exit Function
            
        Case "CLEAR_CONTENTS", "CLEAR_FORMAT", "CLEAR_ALL", "BOLD", "ITALIC", "UNDERLINE", _
             "STRIKETHROUGH", "WRAP_TEXT", "MERGE", "UNMERGE", "FORMAT_PERCENT", "AUTOFIT", _
             "AUTOFIT_ROWS", "AUTOFILTER", "CLEAR_FILTER", "CLEAR_COND_FORMAT", "CLEAR_VALIDATION", _
             "LOCK_CELLS", "UNLOCK_CELLS", "SET_PRINT_AREA", "REMOVE_SPACES", "UPPER_CASE", _
             "LOWER_CASE", "PROPER_CASE", "FLASH_FILL", "SELECT", "BORDER_THICK"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "DELETE_COMMENT", "SHOW_COMMENT", "HIDE_COMMENT", "REMOVE_HYPERLINK", "FREEZE_PANES", "GOTO"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not IsValidSingleCellRef(ws, parts(1)) Then reason = "Invalid cell": Exit Function
            
        Case "SET_VALUE", "SET_FORMULA", "COPY", "CUT", "PASTE_VALUES", "TRANSPOSE"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid first range": Exit Function
            If action = "COPY" Or action = "CUT" Or action = "PASTE_VALUES" Or action = "TRANSPOSE" Then
                If Not IsValidRangeRef(ws, parts(2)) Then reason = "Invalid second range": Exit Function
            ElseIf action = "SET_FORMULA" And Len(Trim$(parts(2))) = 0 Then
                reason = "Formula is empty"
                Exit Function
            End If
            
        Case "FILL_DOWN", "FILL_RIGHT"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidSingleCellRef(ws, parts(1)) Or Not IsValidSingleCellRef(ws, parts(2)) Then
                reason = "Invalid fill cells"
                Exit Function
            End If
            
        Case "FILL_SERIES"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            If Not TryParseDouble(parts(2), d) Then reason = "Invalid numeric step": Exit Function
            
        Case "FONT_NAME", "FORMAT_NUMBER", "FORMAT_DATE", "VALIDATION_LIST", "VALIDATION_CUSTOM"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "FONT_SIZE"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            If Not TryParseLong(parts(2), n) Or n <= 0 Then reason = "Invalid size": Exit Function
            
        Case "FONT_COLOR", "FILL_COLOR"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            If Not IsValidColorSpec(parts(2)) Then reason = "Invalid color": Exit Function
            
        Case "BORDER"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "ALIGN_H", "ALIGN_V"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "FORMAT_CURRENCY"
            If Not ValidateArgCount(action, argCount, 1, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "COLUMN_WIDTH"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidColumnRef(ws, parts(1)) Then reason = "Invalid column": Exit Function
            If Not TryParseDouble(parts(2), d) Or d <= 0 Then reason = "Invalid width": Exit Function
            
        Case "ROW_HEIGHT"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRowRef(ws, parts(1)) Then reason = "Invalid row": Exit Function
            If Not TryParseDouble(parts(2), d) Or d <= 0 Then reason = "Invalid height": Exit Function
            
        Case "INSERT_ROW", "DELETE_ROW", "HIDE_ROW", "SHOW_ROW"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not IsValidRowRef(ws, parts(1)) Then reason = "Invalid row": Exit Function
            
        Case "INSERT_ROWS"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRowRef(ws, parts(1)) Then reason = "Invalid start row": Exit Function
            If Not TryParseLong(parts(2), n) Or n <= 0 Then reason = "Invalid row count": Exit Function
            
        Case "DELETE_ROWS", "HIDE_ROWS", "SHOW_ROWS", "GROUP_ROWS", "UNGROUP_ROWS", "PRINT_TITLES_ROWS"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRowRef(ws, parts(1)) Or Not IsValidRowRef(ws, parts(2)) Then reason = "Invalid row range": Exit Function
            
        Case "INSERT_COLUMN", "DELETE_COLUMN", "HIDE_COLUMN", "SHOW_COLUMN"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not IsValidColumnRef(ws, parts(1)) Then reason = "Invalid column": Exit Function
            
        Case "INSERT_COLUMNS"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidColumnRef(ws, parts(1)) Then reason = "Invalid column": Exit Function
            If Not TryParseLong(parts(2), n) Or n <= 0 Then reason = "Invalid column count": Exit Function
            
        Case "DELETE_COLUMNS", "GROUP_COLUMNS", "UNGROUP_COLUMNS", "PRINT_TITLES_COLS"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidColumnRef(ws, parts(1)) Or Not IsValidColumnRef(ws, parts(2)) Then reason = "Invalid column range": Exit Function
            
        Case "SORT"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid sort range": Exit Function
            If Not IsValidColumnSelector(ws, parts(2)) Then reason = "Invalid sort column": Exit Function
            If Not IsValidSortOrder(parts(3)) Then reason = "Invalid sort order": Exit Function
            
        Case "SORT_MULTI"
            If Not ValidateArgCount(action, argCount, 6, 6, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid sort range": Exit Function
            If Not IsValidColumnSelector(ws, parts(2)) Or Not IsValidColumnSelector(ws, parts(4)) Then reason = "Invalid sort column": Exit Function
            If Not IsValidSortOrder(parts(3)) Or Not IsValidSortOrder(parts(5)) Then reason = "Invalid sort order": Exit Function
            
        Case "FILTER", "FILTER_TOP"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid filter range": Exit Function
            If Not IsValidColumnSelector(ws, parts(2)) Then reason = "Invalid filter column": Exit Function
            If action = "FILTER_TOP" Then
                If Not TryParseLong(parts(3), n) Or n <= 0 Then reason = "Invalid top count": Exit Function
            End If
            
        Case "REMOVE_DUPLICATES"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function

        Case "CREATE_TABLE"
            If Not ValidateArgCount(action, argCount, 1, 2, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
             
        Case "FIND_REPLACE", "PIVOT_REFRESH"
            If Not ValidateArgCount(action, argCount, 1, 3, reason) Then Exit Function
            
        Case "FIND_REPLACE_RANGE"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "CREATE_CHART"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidMultiRangeRef(ws, parts(1)) Then reason = "Invalid chart range": Exit Function
            If Not IsValidChartTypeName(parts(2)) Then reason = "Invalid chart type": Exit Function
            
        Case "CREATE_CHART_POS", "CREATE_CHART_AT"
            If Not ValidateArgCount(action, argCount, 4, 7, reason) Then Exit Function
            If Not IsValidMultiRangeRef(ws, parts(1)) Then reason = "Invalid chart range": Exit Function
            If Not IsValidChartTypeName(parts(2)) Then reason = "Invalid chart type": Exit Function
            
        Case "CHART_TITLE", "CHART_LEGEND", "CHART_TYPE", "CHART_MOVE", "CHART_DELETE", "CHART_AXIS_TITLE", "CHART_RESIZE"
            If Not ValidateArgCount(action, argCount, 1, 3, reason) Then Exit Function
            If Not IsValidChartIndexToken(parts(1)) Then reason = "Invalid chart index": Exit Function
            
        Case "CREATE_PIVOT"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidApplicationRangeRef(parts(1), False) Then reason = "Invalid source range": Exit Function
            If Not IsValidApplicationRangeRef(parts(2), True) Then reason = "Invalid destination cell": Exit Function
            If Len(Trim$(parts(3))) = 0 Then reason = "Missing pivot name": Exit Function
            
        Case "PIVOT_ADD_ROW", "PIVOT_ADD_COLUMN", "PIVOT_ADD_FILTER", "PIVOT_ADD_VALUE"
            If Not ValidateArgCount(action, argCount, 2, 3, reason) Then Exit Function
            
        Case "ADD_SHEET", "DELETE_SHEET", "HIDE_SHEET", "SHOW_SHEET", "ACTIVATE_SHEET", "DELETE_NAME"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Len(Trim$(parts(1))) = 0 Then reason = "Missing name": Exit Function
            
        Case "ADD_SHEET_AFTER", "RENAME_SHEET", "COPY_SHEET", "MOVE_SHEET", "TAB_COLOR", "PROTECT_SHEET", "UNPROTECT_SHEET", "CREATE_NAME"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Len(Trim$(parts(1))) = 0 Or Len(Trim$(parts(2))) = 0 Then reason = "Missing arguments": Exit Function
            If action = "TAB_COLOR" Then
                If Not IsValidColorSpec(parts(2)) Then reason = "Invalid color": Exit Function
            End If
            
        Case "COND_HIGHLIGHT", "COND_TOP", "COND_BOTTOM", "COND_DUPLICATE", "COND_UNIQUE", "COND_TEXT", "COND_BLANK", "COND_NOT_BLANK", _
             "DATA_BARS", "COLOR_SCALE", "COLOR_SCALE3", "ICON_SET"
            If Not ValidateArgCount(action, argCount, 2, 4, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "VALIDATION_NUMBER", "VALIDATION_DATE", "VALIDATION_TEXT_LENGTH"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case "ADD_COMMENT", "EDIT_COMMENT", "ADD_CHECKBOX", "ADD_DROPDOWN"
            If Not ValidateArgCount(action, argCount, 2, 2, reason) Then Exit Function
            If Not IsValidSingleCellRef(ws, parts(1)) Then reason = "Invalid cell": Exit Function
            
        Case "ADD_HYPERLINK", "ADD_HYPERLINK_CELL"
            If Not ValidateArgCount(action, argCount, 3, 3, reason) Then Exit Function
            If Not IsValidSingleCellRef(ws, parts(1)) Then reason = "Invalid cell": Exit Function
            
        Case "ZOOM"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not TryParseLong(parts(1), n) Or n < 10 Or n > 400 Then reason = "Zoom must be 10..400": Exit Function
            
        Case "PAGE_ORIENTATION"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            
        Case "PAGE_MARGINS", "FIT_TO_PAGE"
            If Not ValidateArgCount(action, argCount, 2, 4, reason) Then Exit Function
            
        Case "PRINT_GRIDLINES"
            If Not ValidateArgCount(action, argCount, 1, 1, reason) Then Exit Function
            If Not IsValidBooleanToken(parts(1)) Then reason = "Expected TRUE/FALSE": Exit Function
            
        Case "INSERT_PICTURE"
            If Not ValidateArgCount(action, argCount, 5, 5, reason) Then Exit Function
            If Len(Dir(Trim$(parts(1)))) = 0 Then reason = "Image file not found": Exit Function
            
        Case "ADD_BUTTON"
            If Not ValidateArgCount(action, argCount, 5, 5, reason) Then Exit Function
            
        Case "TEXT_TO_COLUMNS", "SUBTOTAL"
            If Not ValidateArgCount(action, argCount, 2, 3, reason) Then Exit Function
            If Not IsValidRangeRef(ws, parts(1)) Then reason = "Invalid range": Exit Function
            
        Case Else
            reason = "Unknown command"
            Exit Function
    End Select
    
    ValidateCommandStrict = True
End Function


