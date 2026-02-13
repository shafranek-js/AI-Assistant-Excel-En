Attribute VB_Name = "modAIHelper"
'========================================
' AI API integration module
' DeepSeek and OpenRouter (Claude)
' FULL EXCEL COMMAND SET
'========================================

Option Explicit

' API constants
Private Const DEEPSEEK_URL As String = "https://api.deepseek.com/chat/completions"
Private Const OPENROUTER_URL As String = "https://openrouter.ai/api/v1/chat/completions"
Private Const OPENAI_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const DEEPSEEK_MODEL As String = "deepseek-chat"
Private Const CLAUDE_MODEL As String = "anthropic/claude-4.5-sonnet-20250929"
Private Const GPT_MODEL As String = "openai/gpt-5.2"
Private Const GPT_DIRECT_MODEL As String = "gpt-5.2"
Private Const GPT_CODEX_DIRECT_MODEL As String = "gpt-5.2-codex"
Private Const GEMINI_MODEL As String = "google/gemini-3-pro-preview"
Private Const GEMINI_FLASH_MODEL As String = "google/gemini-3-flash-preview-20251217"

' LM Studio default settings
Private Const LMSTUDIO_DEFAULT_IP As String = "127.0.0.1"
Private Const LMSTUDIO_DEFAULT_PORT As String = "1234"
Private Const HTTP_MAX_RETRIES As Long = 3
Private Const HTTP_BACKOFF_BASE_SECONDS As Long = 1
Private Const HTTP_BACKOFF_MAX_SECONDS As Long = 8
Private Const MAX_TOKENS_DEFAULT As Long = 4096
Private Const MAX_TOKENS_OPENAI_DIRECT As Long = 1400
Private Const MAX_CONTEXT_CHARS_DEFAULT As Long = 24000
Private Const MAX_CONTEXT_CHARS_OPENAI_DIRECT As Long = 12000
Private Const CODEX_CLI_TIMEOUT_SECONDS As Long = 90
Private Const CODEX_CLI_TEMP_SUBDIR As String = "ExcelAIAssistantCodex"

' API key storage (Registry)
Private Const REG_PATH As String = "HKEY_CURRENT_USER\Software\ExcelAIAssistant\"

'----------------------------------------
' Get API key from Registry
'----------------------------------------
Public Function GetApiKey(keyName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetApiKey = wsh.RegRead(REG_PATH & keyName)
    Set wsh = Nothing
End Function

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

'----------------------------------------
' Save API key to Registry
'----------------------------------------
Public Sub SaveApiKey(keyName As String, keyValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & keyName, keyValue, "REG_SZ"
    Set wsh = Nothing
End Sub

'----------------------------------------
' Check API key availability
'----------------------------------------
Public Function HasApiKey(model As String) As Boolean
    Select Case LCase$(Trim$(model))
        Case "deepseek"
            HasApiKey = Len(GetApiKey("DeepSeekKey")) > 0
        Case "openai", "gpt-direct"
            HasApiKey = Len(GetApiKey("OpenAIKey")) > 0
        Case Else
            ' Claude/GPT/Gemini via OpenRouter
            HasApiKey = Len(GetApiKey("OpenRouterKey")) > 0
    End Select
End Function

'----------------------------------------
' Send request to AI
'----------------------------------------
Public Function SendToAI(userMessage As String, model As String, Optional excelContext As String = "", Optional imagePath As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim apiUrl As String
    Dim apiKey As String
    Dim modelName As String
    Dim requestBody As String
    Dim response As String
    Dim imageBase64 As String
    Dim maxTokens As Long
    Dim effectiveContext As String
    Dim requiresApiKey As Boolean
    
    ' Validate image
    imageBase64 = ""
    If Len(imagePath) > 0 Then
        ' DeepSeek does not support images
        If model = "deepseek" Then
            SendToAI = "ERROR: DeepSeek does not support images. Select another model (Claude, GPT, Gemini)."
            Exit Function
        End If
        If model <> "codex-cli" Then
            imageBase64 = ImageToBase64(imagePath)
            If Left(imageBase64, 6) = "ERROR:" Then
                SendToAI = "ERROR loading image: " & Mid(imageBase64, 7)
                Exit Function
            End If
        End If
    End If
    
    ' Select API endpoint
    requiresApiKey = True
    If model = "deepseek" Then
        apiUrl = DEEPSEEK_URL
        apiKey = GetApiKey("DeepSeekKey")
        modelName = DEEPSEEK_MODEL
    ElseIf model = "gpt-direct" Then
        apiUrl = OPENAI_URL
        apiKey = GetApiKey("OpenAIKey")
        modelName = GPT_DIRECT_MODEL
    ElseIf model = "gpt-codex-direct" Then
        apiUrl = OPENAI_URL
        apiKey = GetApiKey("OpenAIKey")
        modelName = GPT_CODEX_DIRECT_MODEL
    ElseIf model = "codex-cli" Then
        requiresApiKey = False
    Else
        ' All other models via OpenRouter
        apiUrl = OPENROUTER_URL
        apiKey = GetApiKey("OpenRouterKey")
        
        Select Case model
            Case "claude"
                modelName = CLAUDE_MODEL
            Case "gpt"
                modelName = GPT_MODEL
            Case "gemini"
                modelName = GEMINI_MODEL
            Case "gemini-flash"
                modelName = GEMINI_FLASH_MODEL
            Case Else
                modelName = CLAUDE_MODEL
        End Select
    End If
    
    If requiresApiKey And Len(apiKey) = 0 Then
        SendToAI = "ERROR: API key is not configured. Open Settings."
        Exit Function
    End If
    
    maxTokens = GetMaxTokensForModel(model)
    effectiveContext = ClampContextForModel(excelContext, model)
    
    If model = "codex-cli" Then
        SendToAI = SendToCodexCLI(userMessage, effectiveContext, imagePath)
        Exit Function
    End If

    ' Build system prompt
    Dim systemPrompt As String
    systemPrompt = BuildSystemPrompt(effectiveContext)
    
    ' Build JSON request (with or without image)
    If Len(imageBase64) > 0 Then
        requestBody = BuildRequestJSONWithImage(systemPrompt, userMessage, modelName, imageBase64, imagePath, maxTokens)
    Else
        requestBody = BuildRequestJSON(systemPrompt, userMessage, modelName, maxTokens)
    End If
    
    ' Send request
    response = SendHTTPRequest(apiUrl, apiKey, requestBody, model)
    
    ' Parse response
    SendToAI = ParseResponse(response)
    Exit Function
    
ErrorHandler:
    SendToAI = "ERROR: " & Err.Description
End Function

'----------------------------------------
' Convert image to Base64
'----------------------------------------
Private Function ImageToBase64(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileData() As Byte
    Dim fileLen As Long
    
    ' Check file exists
    If Dir(filePath) = "" Then
        ImageToBase64 = "ERROR:File not found"
        Exit Function
    End If
    
    ' Read file
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileLen = LOF(fileNum)
    
    If fileLen = 0 Then
        Close #fileNum
        ImageToBase64 = "ERROR:File is empty"
        Exit Function
    End If
    
    If fileLen > 20000000 Then ' 20MB limit
        Close #fileNum
        ImageToBase64 = "ERROR:File too large (max 20MB)"
        Exit Function
    End If
    
    ReDim fileData(fileLen - 1)
    Get #fileNum, , fileData
    Close #fileNum
    
    ' Convert to Base64
    ImageToBase64 = EncodeBase64(fileData)
    Exit Function
    
ErrorHandler:
    ImageToBase64 = "ERROR:" & Err.Description
End Function

'----------------------------------------
' Encode byte array to Base64
'----------------------------------------
Private Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As Object
    Dim objNode As Object
    
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

'----------------------------------------
' Get MIME type by file extension
'----------------------------------------
Private Function GetMimeType(filePath As String) As String
    Dim ext As String
    ext = LCase(Right(filePath, 4))
    
    Select Case ext
        Case ".png"
            GetMimeType = "image/png"
        Case ".jpg", "jpeg"
            GetMimeType = "image/jpeg"
        Case ".gif"
            GetMimeType = "image/gif"
        Case "webp"
            GetMimeType = "image/webp"
        Case Else
            GetMimeType = "image/png"
    End Select
End Function

'----------------------------------------
' Build JSON with image (Vision API)
'----------------------------------------
Private Function BuildRequestJSONWithImage(systemPrompt As String, userMessage As String, modelName As String, imageBase64 As String, imagePath As String, Optional maxTokens As Long = MAX_TOKENS_DEFAULT) As String
    Dim json As String
    Dim mimeType As String
    If maxTokens <= 0 Then maxTokens = MAX_TOKENS_DEFAULT
    
    mimeType = GetMimeType(imagePath)
    
    json = "{" & vbCrLf
    json = json & "  ""model"": """ & modelName & """," & vbCrLf
    json = json & "  ""messages"": [" & vbCrLf
    json = json & "    {""role"": ""system"", ""content"": """ & EscapeJSON(systemPrompt) & """}," & vbCrLf
    json = json & "    {""role"": ""user"", ""content"": [" & vbCrLf
    json = json & "      {""type"": ""image_url"", ""image_url"": {""url"": ""data:" & mimeType & ";base64," & imageBase64 & """}}," & vbCrLf
    json = json & "      {""type"": ""text"", ""text"": """ & EscapeJSON(userMessage) & """}" & vbCrLf
    json = json & "    ]}" & vbCrLf
    json = json & "  ]," & vbCrLf
    json = json & "  ""max_tokens"": " & CStr(maxTokens) & vbCrLf
    json = json & "}"
    
    BuildRequestJSONWithImage = json
End Function

'----------------------------------------
' Building a system prompt
'----------------------------------------
Private Function BuildSystemPrompt(excelContext As String) As String
    Dim prompt As String
    
    prompt = "You are an AI assistant for Microsoft Excel. You perform actions with data AUTOMATICALLY." & vbCrLf & vbCrLf
    prompt = prompt & "CRITICAL:" & vbCrLf
    prompt = prompt & "1. Always return executable commands." & vbCrLf
    prompt = prompt & "2. Do not give instructions to the user; perform actions via commands." & vbCrLf
    prompt = prompt & "3. Always use addresses from context; headers may not be on row 1." & vbCrLf
    prompt = prompt & "4. Use the EXACT addresses from the 'Excel Context' section below!" & vbCrLf
    prompt = prompt & "5. FORMULAS: use ENGLISH names (SUM, IF, VLOOKUP...) with comma separators. The system will localize them automatically." & vbCrLf & vbCrLf
    prompt = prompt & "Response format:" & vbCrLf
    prompt = prompt & "1. Brief description (1-2 sentences)" & vbCrLf
    prompt = prompt & "2. MANDATORY block of commands:" & vbCrLf & vbCrLf
    prompt = prompt & "```commands" & vbCrLf
    prompt = prompt & "SET_VALUE|<address>|value" & vbCrLf
    prompt = prompt & "SET_FORMULA|<address>|formula" & vbCrLf
    prompt = prompt & "```" & vbCrLf & vbCrLf
    prompt = prompt & "ADDRESS RULE: If data starts on row N, then:" & vbCrLf
    prompt = prompt & "- Add a new header in row N (next to existing headers)" & vbCrLf
    prompt = prompt & "- Start formulas in row N+1 (first data row)" & vbCrLf
    prompt = prompt & "Example: if headers are on row 2 and data is on rows 3-5:" & vbCrLf
    prompt = prompt & "  SET_VALUE|E2|NewHeader (row 2, header row)" & vbCrLf
    prompt = prompt & "  SET_FORMULA|E3|=B3*C3 (row 3, first data row)" & vbCrLf
    prompt = prompt & "  FILL_DOWN|E3|E5 (fill to last data row)" & vbCrLf & vbCrLf
    
    ' === FULL COMMAND LIST ===
    prompt = prompt & "=== AVAILABLE COMMANDS ===" & vbCrLf & vbCrLf
    
    ' WORKING WITH CELLS
    prompt = prompt & "--- CELLS ---" & vbCrLf
    prompt = prompt & "SET_VALUE|address|value - write value" & vbCrLf
    prompt = prompt & "SET_FORMULA|address|formula - write the formula" & vbCrLf
    prompt = prompt & "FILL_DOWN|start|end - fill down" & vbCrLf
    prompt = prompt & "FILL_RIGHT|start|end - fill right" & vbCrLf
    prompt = prompt & "FILL_SERIES|range|step - fill the sequence" & vbCrLf
    prompt = prompt & "CLEAR_CONTENTS|range - clear contents" & vbCrLf
    prompt = prompt & "CLEAR_FORMAT|range - clear formatting" & vbCrLf
    prompt = prompt & "CLEAR_ALL|range - clear all" & vbCrLf
    prompt = prompt & "COPY|from|where - copy" & vbCrLf
    prompt = prompt & "CUT|from|where - cut" & vbCrLf
    prompt = prompt & "PASTE_VALUES|from|where - paste values" & vbCrLf
    prompt = prompt & "TRANSPOSE|from|to - transpose" & vbCrLf & vbCrLf
    
    ' FORMATTING
    prompt = prompt & "--- FORMATTING ---" & vbCrLf
    prompt = prompt & "BOLD|range - bold" & vbCrLf
    prompt = prompt & "ITALIC|range - italics" & vbCrLf
    prompt = prompt & "UNDERLINE|range - underline" & vbCrLf
    prompt = prompt & "STRIKETHROUGH|range - strikethrough" & vbCrLf
    prompt = prompt & "FONT_NAME|range|font_name - font" & vbCrLf
    prompt = prompt & "FONT_SIZE|range|size - font size" & vbCrLf
    prompt = prompt & "FONT_COLOR|range|color - text color (RED,GREEN,BLUE,BLACK,WHITE,YELLOW,ORANGE,PURPLE,GRAY or RGB:255,0,0)" & vbCrLf
    prompt = prompt & "FILL_COLOR|range|color - fill" & vbCrLf
    prompt = prompt & "BORDER|range|style - borders (ALL,TOP,BOTTOM,LEFT,RIGHT,NONE)" & vbCrLf
    prompt = prompt & "BORDER_THICK|range - thick borders" & vbCrLf
    prompt = prompt & "ALIGN_H|range|alignment - horizontal (LEFT,CENTER,RIGHT)" & vbCrLf
    prompt = prompt & "ALIGN_V|range|alignment - vertical (TOP,CENTER,BOTTOM)" & vbCrLf
    prompt = prompt & "WRAP_TEXT|range - text wrapping" & vbCrLf
    prompt = prompt & "MERGE|range - merge cells" & vbCrLf
    prompt = prompt & "UNMERGE|range - unmerge cells" & vbCrLf
    prompt = prompt & "FORMAT_NUMBER|range|format - number format (#,##0.00)" & vbCrLf
    prompt = prompt & "FORMAT_DATE|range|format - date format (DD.MM.YYYY)" & vbCrLf
    prompt = prompt & "FORMAT_PERCENT|range - percentage format" & vbCrLf
    prompt = prompt & "FORMAT_CURRENCY|range|symbol - currency format" & vbCrLf
    prompt = prompt & "AUTOFIT|range - auto-fit columns" & vbCrLf
    prompt = prompt & "AUTOFIT_ROWS|range - auto-fit rows" & vbCrLf
    prompt = prompt & "COLUMN_WIDTH|column|width - column width" & vbCrLf
    prompt = prompt & "ROW_HEIGHT|row|height - row height" & vbCrLf & vbCrLf
    
    ' ROWS AND COLUMNS
    prompt = prompt & "--- ROWS AND COLUMNS ---" & vbCrLf
    prompt = prompt & "INSERT_ROW|number - insert a row" & vbCrLf
    prompt = prompt & "INSERT_ROWS|number|quantity - insert multiple rows" & vbCrLf
    prompt = prompt & "INSERT_COLUMN|letter - insert a column" & vbCrLf
    prompt = prompt & "INSERT_COLUMNS|letter|quantity - insert multiple columns" & vbCrLf
    prompt = prompt & "DELETE_ROW|number - delete a row" & vbCrLf
    prompt = prompt & "DELETE_ROWS|start|end - delete rows" & vbCrLf
    prompt = prompt & "DELETE_COLUMN|letter - delete a column" & vbCrLf
    prompt = prompt & "DELETE_COLUMNS|start|end - delete columns" & vbCrLf
    prompt = prompt & "HIDE_ROW|number - hide row" & vbCrLf
    prompt = prompt & "HIDE_ROWS|start|end - hide rows" & vbCrLf
    prompt = prompt & "SHOW_ROW|number - show row" & vbCrLf
    prompt = prompt & "SHOW_ROWS|start|end - show rows" & vbCrLf
    prompt = prompt & "HIDE_COLUMN|letter - hide column" & vbCrLf
    prompt = prompt & "SHOW_COLUMN|letter - show column" & vbCrLf
    prompt = prompt & "GROUP_ROWS|start|end - group rows" & vbCrLf
    prompt = prompt & "UNGROUP_ROWS|start|end - ungroup rows" & vbCrLf
    prompt = prompt & "GROUP_COLUMNS|start|end - group columns" & vbCrLf
    prompt = prompt & "UNGROUP_COLUMNS|start|end - ungroup columns" & vbCrLf & vbCrLf
    
    ' SORTING AND FILTERING
    prompt = prompt & "--- SORTING AND FILTERING ---" & vbCrLf
    prompt = prompt & "SORT|range|column|ASC/DESC - sorting" & vbCrLf
    prompt = prompt & "SORT_MULTI|range|number1|order1|number2|order2 - multi-level sorting" & vbCrLf
    prompt = prompt & "AUTOFILTER|range - enable autofilter" & vbCrLf
    prompt = prompt & "FILTER|range|column|value - filter" & vbCrLf
    prompt = prompt & "FILTER_TOP|range|column|count - top N values" & vbCrLf
    prompt = prompt & "CLEAR_FILTER|range - clear the filter" & vbCrLf
    prompt = prompt & "REMOVE_AUTOFILTER - remove autofilter" & vbCrLf
    prompt = prompt & "REMOVE_DUPLICATES|range|columns - remove duplicates" & vbCrLf
    prompt = prompt & "FIND_REPLACE|what|to_what - find and replace" & vbCrLf
    prompt = prompt & "FIND_REPLACE_RANGE|range|what|to_what - replacement in a range" & vbCrLf & vbCrLf
    
    ' GRAPHICS
    prompt = prompt & "--- CHARTS ---" & vbCrLf
    prompt = prompt & "CREATE_CHART|range|type|name - create a chart" & vbCrLf
    prompt = prompt & "  Types: LINE, BAR, COLUMN, PIE, AREA, SCATTER, DOUGHNUT" & vbCrLf
    prompt = prompt & "  IMPORTANT: Select ONLY the columns you need! Use non-adjacent ranges separated by commas." & vbCrLf
    prompt = prompt & "  Examples:" & vbCrLf
    prompt = prompt & "    CREATE_CHART|A2:A5,B2:B5|LINE|Sum by dates - axis X=dates, Y=sums" & vbCrLf
    prompt = prompt & "    CREATE_CHART|A2:B5|COLUMN|Sales - two columns (categories + values)" & vbCrLf
    prompt = prompt & "    CREATE_CHART|B2:B5|PIE|Distribution - one column for pie" & vbCrLf
    prompt = prompt & "CREATE_CHART_AT|range|type|name|cell - chart in the specified cell" & vbCrLf
    prompt = prompt & "CHART_TITLE|LAST|text - title" & vbCrLf
    prompt = prompt & "CHART_LEGEND|LAST|position - legend (TOP,BOTTOM,LEFT,RIGHT,NONE)" & vbCrLf
    prompt = prompt & "CHART_AXIS_TITLE|LAST|X|text - X axis label" & vbCrLf
    prompt = prompt & "CHART_AXIS_TITLE|LAST|Y|text - Y axis label" & vbCrLf
    prompt = prompt & "CHART_TYPE|LAST|type - change type" & vbCrLf
    prompt = prompt & "CHART_MOVE|LAST|cell - move" & vbCrLf
    prompt = prompt & "CHART_RESIZE|LAST|width|height - size" & vbCrLf
    prompt = prompt & "CHART_DELETE|LAST - delete the last one" & vbCrLf
    prompt = prompt & "CHART_DELETE_ALL - delete all" & vbCrLf
    prompt = prompt & "  Index: LAST=last, 1=first, 2=second..." & vbCrLf & vbCrLf
    
    ' PIVOT TABLES
    prompt = prompt & "--- PIVOT TABLES ---" & vbCrLf
    prompt = prompt & "CREATE_PIVOT|source|destination|name - create a pivot" & vbCrLf
    prompt = prompt & "PIVOT_ADD_ROW|name|field - add a field to rows" & vbCrLf
    prompt = prompt & "PIVOT_ADD_COLUMN|name|field - add a field to the columns" & vbCrLf
    prompt = prompt & "PIVOT_ADD_VALUE|name|field|function - add value (SUM,COUNT,AVERAGE,MAX,MIN)" & vbCrLf
    prompt = prompt & "PIVOT_ADD_FILTER|name|field - add filter" & vbCrLf
    prompt = prompt & "PIVOT_REFRESH|name - update the summary" & vbCrLf
    prompt = prompt & "PIVOT_REFRESH_ALL - update all summary" & vbCrLf & vbCrLf
    
    ' SHEETS
    prompt = prompt & "--- SHEETS ---" & vbCrLf
    prompt = prompt & "ADD_SHEET|name - add sheet" & vbCrLf
    prompt = prompt & "ADD_SHEET_AFTER|name|after - add sheet after" & vbCrLf
    prompt = prompt & "DELETE_SHEET|name - delete sheet" & vbCrLf
    prompt = prompt & "RENAME_SHEET|old|new - rename" & vbCrLf
    prompt = prompt & "COPY_SHEET|name|new_name - copy sheet" & vbCrLf
    prompt = prompt & "MOVE_SHEET|name|position - move sheet" & vbCrLf
    prompt = prompt & "HIDE_SHEET|name - hide sheet" & vbCrLf
    prompt = prompt & "SHOW_SHEET|name - show sheet" & vbCrLf
    prompt = prompt & "ACTIVATE_SHEET|name - activate sheet" & vbCrLf
    prompt = prompt & "TAB_COLOR|name|color - label color" & vbCrLf
    prompt = prompt & "PROTECT_SHEET|name|password - protect the sheet" & vbCrLf
    prompt = prompt & "UNPROTECT_SHEET|name|password - remove protection" & vbCrLf & vbCrLf
    
    ' NAMED RANGES
    prompt = prompt & "--- NAMED RANGES ---" & vbCrLf
    prompt = prompt & "CREATE_NAME|name|range - create a name" & vbCrLf
    prompt = prompt & "DELETE_NAME|name - delete name" & vbCrLf & vbCrLf
    
    ' CONDITIONAL FORMATTING
    prompt = prompt & "--- CONDITIONAL FORMATTING ---" & vbCrLf
    prompt = prompt & "COND_HIGHLIGHT|range|operator|value|color - highlight (operator: >,<,=,>=,<=,<>,BETWEEN)" & vbCrLf
    prompt = prompt & "COND_TOP|range|quantity|color - top N" & vbCrLf
    prompt = prompt & "COND_BOTTOM|range|quantity|color - last N" & vbCrLf
    prompt = prompt & "COND_DUPLICATE|range|color - duplicates" & vbCrLf
    prompt = prompt & "COND_UNIQUE|range|color - unique" & vbCrLf
    prompt = prompt & "COND_TEXT|range|text|color - contains text" & vbCrLf
    prompt = prompt & "COND_BLANK|range|color - empty cells" & vbCrLf
    prompt = prompt & "COND_NOT_BLANK|range|color - non-blank cells" & vbCrLf
    prompt = prompt & "DATA_BARS|range|color - histograms" & vbCrLf
    prompt = prompt & "COLOR_SCALE|range|color1|color2 - color scale" & vbCrLf
    prompt = prompt & "COLOR_SCALE3|range|color1|color2|color3 - 3-color scale" & vbCrLf
    prompt = prompt & "ICON_SET|range|set - icons (ARROWS,FLAGS,STARS,BARS)" & vbCrLf
    prompt = prompt & "CLEAR_COND_FORMAT|range - clear conditional formatting" & vbCrLf & vbCrLf
    
    ' DATA CHECK
    prompt = prompt & "--- DATA CHECK ---" & vbCrLf
    prompt = prompt & "VALIDATION_LIST|range|values ​​- drop-down list (values ​​separated by ;)" & vbCrLf
    prompt = prompt & "VALIDATION_NUMBER|range|min|max - numbers in the range" & vbCrLf
    prompt = prompt & "VALIDATION_DATE|range|start|end - dates in the range" & vbCrLf
    prompt = prompt & "VALIDATION_TEXT_LENGTH|range|min|max - text length" & vbCrLf
    prompt = prompt & "VALIDATION_CUSTOM|range|formula - custom formula" & vbCrLf
    prompt = prompt & "CLEAR_VALIDATION|range - clear validation" & vbCrLf & vbCrLf
    
    ' COMMENTS AND NOTES
    prompt = prompt & "--- COMMENTS ---" & vbCrLf
    prompt = prompt & "ADD_COMMENT|address|text - add a comment" & vbCrLf
    prompt = prompt & "EDIT_COMMENT|address|text - change comment" & vbCrLf
    prompt = prompt & "DELETE_COMMENT|address - delete a comment" & vbCrLf
    prompt = prompt & "SHOW_COMMENT|address - show comment" & vbCrLf
    prompt = prompt & "HIDE_COMMENT|address - hide comment" & vbCrLf
    prompt = prompt & "SHOW_ALL_COMMENTS - show all" & vbCrLf
    prompt = prompt & "HIDE_ALL_COMMENTS - hide all" & vbCrLf & vbCrLf
    
    ' HYPERLINKS
    prompt = prompt & "--- HYPERLINKS ---" & vbCrLf
    prompt = prompt & "ADD_HYPERLINK|address|url|text - add a link" & vbCrLf
    prompt = prompt & "ADD_HYPERLINK_CELL|address|link_to_cell|text - link to cell" & vbCrLf
    prompt = prompt & "REMOVE_HYPERLINK|address - remove link" & vbCrLf & vbCrLf
    
    ' PROTECTION
    prompt = prompt & "--- PROTECTION ---" & vbCrLf
    prompt = prompt & "LOCK_CELLS|range - lock cells" & vbCrLf
    prompt = prompt & "UNLOCK_CELLS|range - unlock cells" & vbCrLf & vbCrLf
    
    ' VIEWING AREA
    prompt = prompt & "--- VIEWING AREA ---" & vbCrLf
    prompt = prompt & "FREEZE_PANES|address - freeze areas" & vbCrLf
    prompt = prompt & "FREEZE_TOP_ROW - freeze the top row" & vbCrLf
    prompt = prompt & "FREEZE_FIRST_COLUMN - freeze the first column" & vbCrLf
    prompt = prompt & "UNFREEZE_PANES - unfasten" & vbCrLf
    prompt = prompt & "ZOOM|percentage - scale" & vbCrLf
    prompt = prompt & "GOTO|address - go to cell" & vbCrLf
    prompt = prompt & "SELECT|range - select a range" & vbCrLf & vbCrLf
    
    ' SEAL
    prompt = prompt & "--- SEAL ---" & vbCrLf
    prompt = prompt & "SET_PRINT_AREA|range - print area" & vbCrLf
    prompt = prompt & "CLEAR_PRINT_AREA - clear the print area" & vbCrLf
    prompt = prompt & "PAGE_ORIENTATION|PORTRAIT/LANDSCAPE - orientation" & vbCrLf
    prompt = prompt & "PAGE_MARGINS|left|right|top|bottom - margins (in cm)" & vbCrLf
    prompt = prompt & "PRINT_TITLES_ROWS|start|end - print title rows" & vbCrLf
    prompt = prompt & "PRINT_TITLES_COLS|start|end - print title columns" & vbCrLf
    prompt = prompt & "PRINT_GRIDLINES|TRUE/FALSE - print a grid" & vbCrLf
    prompt = prompt & "FIT_TO_PAGE|width|height - fit to pages" & vbCrLf & vbCrLf
    
    ' IMAGES
    prompt = prompt & "--- IMAGES ---" & vbCrLf
    prompt = prompt & "INSERT_PICTURE|path|left|top|width|height - insert an image" & vbCrLf
    prompt = prompt & "DELETE_PICTURES - delete all images" & vbCrLf & vbCrLf
    
    ' FORMS
    prompt = prompt & "--- FORMS ---" & vbCrLf
    prompt = prompt & "ADD_BUTTON|left|top|width|height|text - add a button" & vbCrLf
    prompt = prompt & "ADD_CHECKBOX|address|text - add a checkbox" & vbCrLf
    prompt = prompt & "ADD_DROPDOWN|address|values ​​- add a drop-down list" & vbCrLf
    prompt = prompt & "DELETE_SHAPES - delete all shapes" & vbCrLf & vbCrLf
    
    ' SPECIAL
    prompt = prompt & "--- SPECIAL ---" & vbCrLf
    prompt = prompt & "CALCULATE - recalculate" & vbCrLf
    prompt = prompt & "CALCULATE_SHEET - recalculate the sheet" & vbCrLf
    prompt = prompt & "TEXT_TO_COLUMNS|range|separator - text by columns" & vbCrLf
    prompt = prompt & "REMOVE_SPACES|range - remove extra spaces" & vbCrLf
    prompt = prompt & "UPPER_CASE|range - to UPPER CASE" & vbCrLf
    prompt = prompt & "LOWER_CASE|range - to lower case" & vbCrLf
    prompt = prompt & "PROPER_CASE|range - Each Word Capitalized" & vbCrLf
    prompt = prompt & "FLASH_FILL|range - instant filling" & vbCrLf
    prompt = prompt & "SUBTOTAL|range|function|column - subtotals (SUM,COUNT,AVERAGE)" & vbCrLf
    prompt = prompt & "REMOVE_SUBTOTALS - remove subtotals" & vbCrLf & vbCrLf
    
    prompt = prompt & "=== END OF COMMAND LIST ===" & vbCrLf & vbCrLf
    prompt = prompt & "Answer in " & GetResponseLanguageForPrompt() & ". ALWAYS include the ```commands``` block with commands!" & vbCrLf
    prompt = prompt & "If the task is complex, use several commands sequentially." & vbCrLf & vbCrLf
    
    If Len(excelContext) > 0 Then
        prompt = prompt & "Excel context:" & vbCrLf & excelContext
    End If
    
    BuildSystemPrompt = prompt
End Function

Private Function GetResponseLanguageForPrompt() As String
    Dim lang As String
    lang = Trim(GetLMStudioSetting("ResponseLanguage"))

    Select Case LCase$(lang)
        Case "russian"
            GetResponseLanguageForPrompt = "Russian"
        Case "ukrainian"
            GetResponseLanguageForPrompt = "Ukrainian"
        Case "czech"
            GetResponseLanguageForPrompt = "Czech"
        Case "spanish"
            GetResponseLanguageForPrompt = "Spanish"
        Case "german"
            GetResponseLanguageForPrompt = "German"
        Case Else
            GetResponseLanguageForPrompt = "English"
    End Select
End Function

Private Function GetMaxTokensForModel(model As String) As Long
    Select Case LCase$(Trim$(model))
        Case "gpt-direct", "gpt-codex-direct"
            GetMaxTokensForModel = MAX_TOKENS_OPENAI_DIRECT
        Case Else
            GetMaxTokensForModel = MAX_TOKENS_DEFAULT
    End Select
End Function

Private Function GetMaxContextCharsForModel(model As String) As Long
    Select Case LCase$(Trim$(model))
        Case "gpt-direct", "gpt-codex-direct"
            GetMaxContextCharsForModel = MAX_CONTEXT_CHARS_OPENAI_DIRECT
        Case Else
            GetMaxContextCharsForModel = MAX_CONTEXT_CHARS_DEFAULT
    End Select
End Function

Private Function ClampContextForModel(excelContext As String, model As String) As String
    Dim maxChars As Long
    Dim marker As String
    Dim keepHead As Long
    Dim keepTail As Long
    Dim ctx As String
    
    maxChars = GetMaxContextCharsForModel(model)
    ctx = excelContext
    
    If maxChars <= 0 Or Len(ctx) <= maxChars Then
        ClampContextForModel = ctx
        Exit Function
    End If
    
    marker = vbCrLf & "...[context truncated to reduce token usage]..." & vbCrLf
    
    keepHead = CLng(maxChars * 0.6)
    keepTail = maxChars - keepHead - Len(marker)
    If keepTail < 0 Then keepTail = 0
    
    ClampContextForModel = Left$(ctx, keepHead) & marker & Right$(ctx, keepTail)
End Function

'----------------------------------------
' Building a JSON request
'----------------------------------------
Private Function BuildRequestJSON(systemPrompt As String, userMessage As String, modelName As String, Optional maxTokens As Long = MAX_TOKENS_DEFAULT) As String
    Dim json As String
    If maxTokens <= 0 Then maxTokens = MAX_TOKENS_DEFAULT
    
    ' Escaping special characters
    systemPrompt = EscapeJSON(systemPrompt)
    userMessage = EscapeJSON(userMessage)
    
    json = "{"
    json = json & """model"": """ & modelName & ""","
    json = json & """messages"": ["
    json = json & "{""role"": ""system"", ""content"": """ & systemPrompt & """},"
    json = json & "{""role"": ""user"", ""content"": """ & userMessage & """}"
    json = json & "],"
    json = json & """temperature"": 0.1,"
    json = json & """max_tokens"": " & CStr(maxTokens)
    json = json & "}"
    
    BuildRequestJSON = json
End Function

'----------------------------------------
' JSON escaping
'----------------------------------------
Private Function EscapeJSON(text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    EscapeJSON = result
End Function

'----------------------------------------
' HTTP request
'----------------------------------------
Private Function SendHTTPRequest(url As String, apiKey As String, body As String, model As String) As String
    ' Checking for the presence of a key
    If Len(Trim(apiKey)) = 0 Then
        SendHTTPRequest = "{""error"": ""API key is not configured. Open Settings and enter the key.""}"
        Exit Function
    End If
    
    SendHTTPRequest = SendHttpPostWithRetry(url, body, apiKey, model, False)
End Function

Private Function SendHttpPostWithRetry(url As String, body As String, apiKey As String, model As String, isLocal As Boolean) As String
    Dim attempt As Long
    Dim http As Object
    Dim statusCode As Long
    Dim statusText As String
    Dim responseText As String
    Dim retryAfter As Long
    Dim errMsg As String
    Dim canRetry As Boolean
    
    SendHttpPostWithRetry = "{""error"": ""Unknown network error.""}"
    
    For attempt = 0 To HTTP_MAX_RETRIES
        On Error GoTo SendError
        
        Set http = CreateHttpClient()
        If http Is Nothing Then
            SendHttpPostWithRetry = "{""error"": ""Could not create HTTP object.""}"
            Exit Function
        End If
        
        On Error Resume Next
        If isLocal Then
            http.setTimeouts 5000, 10000, 120000, 300000
        Else
            http.setTimeouts 5000, 10000, 60000, 120000
        End If
        Err.Clear
        http.Open "POST", url, False
        http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "Authorization", "Bearer " & apiKey
        
        If Not isLocal Then
            If model = "claude" Then
                http.setRequestHeader "HTTP-Referer", "https://excel-ai-assistant.local"
                http.setRequestHeader "X-Title", "Excel AI Assistant VBA"
            End If
        End If
        
        http.send body
        
        statusCode = CLng(http.Status)
        statusText = CStr(http.statusText)
        responseText = CStr(http.responseText)
        
        If statusCode = 200 Then
            SendHttpPostWithRetry = responseText
            Set http = Nothing
            Exit Function
        End If
        
        If isLocal Then
            SendHttpPostWithRetry = BuildLocalHttpError(statusCode, statusText, responseText)
        Else
            SendHttpPostWithRetry = BuildCloudHttpError(statusCode, statusText, responseText)
        End If
        
        canRetry = IsRetryableStatus(statusCode)
        If canRetry And attempt < HTTP_MAX_RETRIES Then
            retryAfter = GetRetryAfterSeconds(http)
            WaitWithBackoff attempt, retryAfter
            Set http = Nothing
            GoTo NextAttempt
        Else
            Set http = Nothing
            Exit Function
        End If
        
NextAttempt:
        On Error GoTo 0
    Next attempt
    
    Exit Function
    
SendError:
    errMsg = "Error " & Err.Number & ": " & Err.Description
    SendHttpPostWithRetry = "{""error"": """ & EscapeJSON(errMsg) & """}"
    
    canRetry = IsRetryableNetworkError(Err.Number, Err.Description)
    If canRetry And attempt < HTTP_MAX_RETRIES Then
        WaitWithBackoff attempt, 0
        Err.Clear
        Set http = Nothing
        Resume NextAttempt
    End If
    
    On Error GoTo 0
    Set http = Nothing
End Function

Private Function CreateHttpClient() As Object
    On Error Resume Next
    Set CreateHttpClient = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If CreateHttpClient Is Nothing Then
        Set CreateHttpClient = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If CreateHttpClient Is Nothing Then
        Set CreateHttpClient = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
    If CreateHttpClient Is Nothing Then
        Set CreateHttpClient = CreateObject("MSXML2.XMLHTTP.6.0")
    End If
    If CreateHttpClient Is Nothing Then
        Set CreateHttpClient = CreateObject("MSXML2.XMLHTTP")
    End If
End Function

Private Function BuildCloudHttpError(statusCode As Long, statusText As String, responseBody As String) As String
    Dim bodyLower As String
    bodyLower = LCase$(responseBody)
    
    ' Preserve provider error payload when possible, ParseResponse will extract message/code.
    If InStr(responseBody, """error""") > 0 Then
        BuildCloudHttpError = responseBody
        Exit Function
    End If
    
    If statusCode = 401 Then
        BuildCloudHttpError = "{""error"": ""Authorization error (401). Please check your API key.""}"
    ElseIf statusCode = 403 Then
        BuildCloudHttpError = "{""error"": ""Access denied (403). Check your API key and limits.""}"
    ElseIf statusCode = 429 Then
        If InStr(bodyLower, "insufficient_quota") > 0 Then
            BuildCloudHttpError = "{""error"": ""Quota/billing limit reached (429 insufficient_quota). Check OpenAI billing and project limits.""}"
        Else
            BuildCloudHttpError = "{""error"": ""Request limit exceeded (429). Please wait.""}"
        End If
    ElseIf statusCode >= 500 Then
        BuildCloudHttpError = "{""error"": ""Server error (" & statusCode & ").""}"
    Else
        BuildCloudHttpError = "{""error"": ""HTTP " & statusCode & ": " & statusText & """}"
    End If
End Function

Private Function BuildLocalHttpError(statusCode As Long, statusText As String, responseBody As String) As String
    If InStr(responseBody, """error""") > 0 Then
        BuildLocalHttpError = responseBody
    Else
        BuildLocalHttpError = "{""error"": ""HTTP " & statusCode & ": " & statusText & """}"
    End If
End Function

Private Function IsRetryableStatus(statusCode As Long) As Boolean
    IsRetryableStatus = (statusCode = 429 Or statusCode >= 500)
End Function

Private Function IsRetryableNetworkError(errNumber As Long, errDescription As String) As Boolean
    Dim d As String
    d = UCase$(errDescription)
    
    Select Case errNumber
        Case -2147012894, -2147012867, -2147012866, -2147012746
            IsRetryableNetworkError = True
            Exit Function
    End Select
    
    If InStr(d, "TIMEOUT") > 0 Then
        IsRetryableNetworkError = True
    ElseIf InStr(d, "TIMED OUT") > 0 Then
        IsRetryableNetworkError = True
    ElseIf InStr(d, "TEMPORARY") > 0 Then
        IsRetryableNetworkError = True
    ElseIf InStr(d, "CONNECTION") > 0 Then
        IsRetryableNetworkError = True
    ElseIf InStr(d, "UNAVAILABLE") > 0 Then
        IsRetryableNetworkError = True
    Else
        IsRetryableNetworkError = False
    End If
End Function

Private Function GetRetryAfterSeconds(http As Object) As Long
    On Error Resume Next
    Dim v As String
    v = Trim$(CStr(http.getResponseHeader("Retry-After")))
    
    If Len(v) = 0 Then
        GetRetryAfterSeconds = 0
    ElseIf IsNumeric(v) Then
        GetRetryAfterSeconds = CLng(v)
        If GetRetryAfterSeconds < 0 Then GetRetryAfterSeconds = 0
    Else
        GetRetryAfterSeconds = 0
    End If
End Function

Private Sub WaitWithBackoff(attempt As Long, retryAfterSeconds As Long)
    Dim waitSeconds As Long
    Dim i As Long
    
    If retryAfterSeconds > 0 Then
        waitSeconds = retryAfterSeconds
    Else
        waitSeconds = HTTP_BACKOFF_BASE_SECONDS * (2 ^ attempt)
    End If
    
    If waitSeconds < 1 Then waitSeconds = 1
    If waitSeconds > HTTP_BACKOFF_MAX_SECONDS Then waitSeconds = HTTP_BACKOFF_MAX_SECONDS
    
    For i = 1 To waitSeconds
        DoEvents
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
End Sub

'----------------------------------------
' Parsing JSON response
'----------------------------------------
Private Function ParseResponse(jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    Dim content As String
    Dim startPos As Long
    Dim endPos As Long
    Dim errMsg As String
    Dim lowerJson As String
    
    ' Checking for errors
    If InStr(jsonResponse, """error""") > 0 Then
        lowerJson = LCase$(jsonResponse)
        
        If InStr(lowerJson, "insufficient_quota") > 0 Then
            ParseResponse = "API ERROR: Quota/billing limit reached (insufficient_quota). Check OpenAI billing and project limits."
            Exit Function
        End If
        
        errMsg = ExtractJsonStringByKey(jsonResponse, "message")
        If Len(errMsg) = 0 Then errMsg = ExtractJsonStringByKey(jsonResponse, "error")
        If Len(errMsg) = 0 Then errMsg = ExtractJsonStringByKey(jsonResponse, "detail")
        
        If Len(errMsg) = 0 Then
            If InStr(lowerJson, "rate_limit") > 0 Then
                errMsg = "Request limit exceeded. Please wait and retry with less data."
            Else
                errMsg = "Unknown error"
            End If
        End If
        
        ParseResponse = "API ERROR: " & errMsg
        Exit Function
    End If
    
    ' Looking for content in the response
    startPos = InStr(jsonResponse, """content"":")
    If startPos = 0 Then
        ParseResponse = "ERROR: Could not parse response"
        Exit Function
    End If
    
    ' Finding the beginning of the value
    startPos = InStr(startPos, jsonResponse, ":") + 1
    
    ' Skip spaces
    Do While Mid(jsonResponse, startPos, 1) = " "
        startPos = startPos + 1
    Loop
    
    ' Checking if it starts with a quote
    If Mid(jsonResponse, startPos, 1) = """" Then
        startPos = startPos + 1
        ' Looking for the closing quote (not escaped)
        endPos = startPos
        Do
            endPos = InStr(endPos, jsonResponse, """")
            If endPos = 0 Then Exit Do
            ' Checking whether it is shielded
            If Mid(jsonResponse, endPos - 1, 1) <> "\" Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        
        If endPos > startPos Then
            content = Mid(jsonResponse, startPos, endPos - startPos)
        End If
    End If
    
    ' Removing shielding
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\t", vbTab)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")
    
    ParseResponse = content
    Exit Function
    
ErrorHandler:
    ParseResponse = "Parsing ERROR: " & Err.Description
End Function

Private Function ExtractJsonStringByKey(jsonText As String, keyName As String) As String
    On Error GoTo ErrorHandler
    
    Dim keyPos As Long
    Dim colonPos As Long
    Dim q1 As Long
    Dim q2 As Long
    Dim result As String
    
    keyPos = InStr(1, jsonText, """" & keyName & """", vbTextCompare)
    If keyPos = 0 Then Exit Function
    
    colonPos = InStr(keyPos, jsonText, ":")
    If colonPos = 0 Then Exit Function
    
    q1 = InStr(colonPos + 1, jsonText, """")
    If q1 = 0 Then Exit Function
    
    q2 = q1 + 1
    Do
        q2 = InStr(q2, jsonText, """")
        If q2 = 0 Then Exit Do
        If Mid$(jsonText, q2 - 1, 1) <> "\" Then Exit Do
        q2 = q2 + 1
    Loop
    
    If q2 > q1 Then
        result = Mid$(jsonText, q1 + 1, q2 - q1 - 1)
        result = Replace(result, "\n", " ")
        result = Replace(result, "\t", " ")
        result = Replace(result, "\""", """")
        result = Replace(result, "\\", "\")
        ExtractJsonStringByKey = result
    End If
    Exit Function
    
ErrorHandler:
    ExtractJsonStringByKey = ""
End Function

'----------------------------------------
' Extracting commands from response
'----------------------------------------
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

'----------------------------------------
' Executing commands
'----------------------------------------
Public Function ExecuteCommands(commands As String) As String
    On Error GoTo ErrorHandler
    
    Dim lines() As String
    Dim i As Long
    Dim cmd As String
    Dim executedCount As Long
    Dim rejectedCount As Long
    Dim failedCount As Long
    Dim result As String
    Dim validationError As String
    Dim details As String
    Dim detailCount As Long
    Dim maxDetails As Long
    
    If Len(commands) = 0 Then
        ExecuteCommands = ""
        Exit Function
    End If
    
    lines = Split(commands, vbLf)
    executedCount = 0
    rejectedCount = 0
    failedCount = 0
    details = ""
    detailCount = 0
    maxDetails = 5
    
    For i = 0 To UBound(lines)
        cmd = Trim(Replace(lines(i), vbCr, ""))
        If Len(cmd) > 0 Then
            validationError = ""
            If ValidateCommandStrict(cmd, validationError) Then
                If ExecuteSingleCommand(cmd) Then
                    executedCount = executedCount + 1
                Else
                    failedCount = failedCount + 1
                    If detailCount < maxDetails Then
                        details = details & vbCrLf & "- Runtime failure: " & cmd
                        detailCount = detailCount + 1
                    End If
                End If
            Else
                rejectedCount = rejectedCount + 1
                If detailCount < maxDetails Then
                    details = details & vbCrLf & "- Rejected (" & validationError & "): " & cmd
                    detailCount = detailCount + 1
                End If
            End If
        End If
    Next i
    
    result = "[Commands executed: " & executedCount & ", rejected: " & rejectedCount & ", failed: " & failedCount & "]"
    If Len(details) > 0 Then
        result = result & vbCrLf & "Details:" & details
    End If
    ExecuteCommands = result
    Exit Function
    
ErrorHandler:
    ExecuteCommands = "Runtime error: " & Err.Description
End Function

'----------------------------------------
' Strict command validation before execution
'----------------------------------------
Private Function ValidateCommandStrict(cmd As String, ByRef reason As String) As Boolean
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
             "LOWER_CASE", "PROPER_CASE", "FLASH_FILL", "SELECT"
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


'----------------------------------------
' Color parsing
'----------------------------------------
Private Function ParseColor(colorStr As String) As Long
    On Error Resume Next
    
    Dim c As String
    Dim rgbParts() As String
    
    c = UCase(Trim(colorStr))
    Debug.Print "ParseColor: input=[" & colorStr & "] upper=[" & c & "]"
    
    ' Checking the RGB format
    If Len(c) >= 4 And Left(c, 4) = "RGB:" Then
        rgbParts = Split(Mid(c, 5), ",")
        If UBound(rgbParts) >= 2 Then
            ParseColor = RGB(CLng(Trim(rgbParts(0))), CLng(Trim(rgbParts(1))), CLng(Trim(rgbParts(2))))
            Debug.Print "ParseColor: RGB result=" & ParseColor
            Exit Function
        End If
    End If
    
    ' Predefined colors
    Select Case c
        Case "RED": ParseColor = RGB(255, 0, 0)
        Case "GREEN": ParseColor = RGB(0, 128, 0)
        Case "BLUE": ParseColor = RGB(0, 0, 255)
        Case "YELLOW": ParseColor = RGB(255, 255, 0)
        Case "ORANGE": ParseColor = RGB(255, 165, 0)
        Case "PURPLE": ParseColor = RGB(128, 0, 128)
        Case "PINK": ParseColor = RGB(255, 192, 203)
        Case "CYAN": ParseColor = RGB(0, 255, 255)
        Case "WHITE": ParseColor = RGB(255, 255, 255)
        Case "BLACK": ParseColor = RGB(0, 0, 0)
        Case "GRAY", "GREY": ParseColor = RGB(128, 128, 128)
        Case "LIGHTGRAY", "LIGHTGREY": ParseColor = RGB(192, 192, 192)
        Case "DARKGRAY", "DARKGREY": ParseColor = RGB(64, 64, 64)
        Case "BROWN": ParseColor = RGB(139, 69, 19)
        Case "LIME": ParseColor = RGB(0, 255, 0)
        Case "NAVY": ParseColor = RGB(0, 0, 128)
        Case "TEAL": ParseColor = RGB(0, 128, 128)
        Case "MAROON": ParseColor = RGB(128, 0, 0)
        Case "OLIVE": ParseColor = RGB(128, 128, 0)
        Case "GOLD": ParseColor = RGB(255, 215, 0)
        Case "SILVER": ParseColor = RGB(192, 192, 192)
        Case Else: ParseColor = RGB(0, 0, 0) ' Default black
    End Select
    
    Debug.Print "ParseColor: result=" & ParseColor
    If Err.Number <> 0 Then
        Debug.Print "ParseColor: ERROR " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Function

'----------------------------------------
' Formula localization (English functions -> Russian)
' COMPLETE LIST OF ALL EXCEL FUNCTIONS
'----------------------------------------
Private Function LocalizeFormula(formula As String) As String
    Dim result As String
    result = formula
    
    ' ===MATHEMATICAL ===
    result = ReplaceFunc(result, "ABS", "ABS")
    result = ReplaceFunc(result, "ACOS", "ACOS")
    result = ReplaceFunc(result, "ACOSH", "ACOSH")
    result = ReplaceFunc(result, "ACOT", "ACOT")
    result = ReplaceFunc(result, "ACOTH", "ACOTH")
    result = ReplaceFunc(result, "AGGREGATE", "UNIT")
    result = ReplaceFunc(result, "ARABIC", "ARAB")
    result = ReplaceFunc(result, "ASIN", "ASIN")
    result = ReplaceFunc(result, "ASINH", "ASINH")
    result = ReplaceFunc(result, "ATAN", "ATAN")
    result = ReplaceFunc(result, "ATAN2", "ATAN2")
    result = ReplaceFunc(result, "ATANH", "ATANH")
    result = ReplaceFunc(result, "BASE", "BASE")
    result = ReplaceFunc(result, "CEILING.MATH", "OVERTOP.MAT")
    result = ReplaceFunc(result, "CEILING.PRECISE", "OKRUP.PRECISION")
    result = ReplaceFunc(result, "CEILING", "OKRVVERH")
    result = ReplaceFunc(result, "COMBIN", "NUMBERCOMB")
    result = ReplaceFunc(result, "COMBINA", "NUMBERCOMBA")
    result = ReplaceFunc(result, "COS", "COS")
    result = ReplaceFunc(result, "COSH", "COSH")
    result = ReplaceFunc(result, "COT", "COT")
    result = ReplaceFunc(result, "COTH", "COTH")
    result = ReplaceFunc(result, "CSC", "CSC")
    result = ReplaceFunc(result, "CSCH", "CSCH")
    result = ReplaceFunc(result, "DECIMAL", "DES")
    result = ReplaceFunc(result, "DEGREES", "DEGREES")
    result = ReplaceFunc(result, "EVEN", "EVEN")
    result = ReplaceFunc(result, "EXP", "EXP")
    result = ReplaceFunc(result, "FACT", "FACT")
    result = ReplaceFunc(result, "FACTDOUBLE", "DVFACTR")
    result = ReplaceFunc(result, "FLOOR.MATH", "OKRVNIZ.MAT")
    result = ReplaceFunc(result, "FLOOR.PRECISE", "OKRV.PRECISION")
    result = ReplaceFunc(result, "FLOOR", "OKRVNIZ")
    result = ReplaceFunc(result, "GCD", "GCD")
    result = ReplaceFunc(result, "INT", "WHOLE")
    result = ReplaceFunc(result, "ISO.CEILING", "ISO.OVERUP")
    result = ReplaceFunc(result, "LCM", "NOC")
    result = ReplaceFunc(result, "LN", "LN")
    result = ReplaceFunc(result, "LOG10", "LOG10")
    result = ReplaceFunc(result, "LOG", "LOG")
    result = ReplaceFunc(result, "MDETERM", "MOPRED")
    result = ReplaceFunc(result, "MINVERSE", "MOBR")
    result = ReplaceFunc(result, "MMULT", "MUMNIFE")
    result = ReplaceFunc(result, "MOD", "OSTAT")
    result = ReplaceFunc(result, "MROUND", "ROUND")
    result = ReplaceFunc(result, "MULTINOMIAL", "MULTINOM")
    result = ReplaceFunc(result, "MUNIT", "MEDIN")
    result = ReplaceFunc(result, "ODD", "ODD")
    result = ReplaceFunc(result, "PI", "PI")
    result = ReplaceFunc(result, "POWER", "DEGREE")
    result = ReplaceFunc(result, "PRODUCT", "PRODUCT")
    result = ReplaceFunc(result, "QUOTIENT", "PRIVATE")
    result = ReplaceFunc(result, "RADIANS", "RADIANS")
    result = ReplaceFunc(result, "RANDBETWEEN", "CASE BETWEEN")
    result = ReplaceFunc(result, "RAND", "RAND")
    result = ReplaceFunc(result, "ROMAN", "ROMAN")
    result = ReplaceFunc(result, "ROUNDDOWN", "ROUND BOTTOM")
    result = ReplaceFunc(result, "ROUNDUP", "ROUNDUP")
    result = ReplaceFunc(result, "ROUND", "ROUND")
    result = ReplaceFunc(result, "SEC", "SEC")
    result = ReplaceFunc(result, "SECH", "SECH")
    result = ReplaceFunc(result, "SERIESSUM", "SERIES.SUM")
    result = ReplaceFunc(result, "SIGN", "SIGN")
    result = ReplaceFunc(result, "SIN", "SIN")
    result = ReplaceFunc(result, "SINH", "SINH")
    result = ReplaceFunc(result, "SQRT", "ROOT")
    result = ReplaceFunc(result, "SQRTPI", "KORE|NY.PI")
    result = ReplaceFunc(result, "SUBTOTAL", "INTERMEDIATE.RESULTS")
    result = ReplaceFunc(result, "SUMIFS", "SUMIFS")
    result = ReplaceFunc(result, "SUMIF", "SUMIF")
    result = ReplaceFunc(result, "SUMPRODUCT", "SUMPRODUCT")
    result = ReplaceFunc(result, "SUMSQ", "SUMMKV")
    result = ReplaceFunc(result, "SUMX2MY2", "SUMMDISC")
    result = ReplaceFunc(result, "SUMX2PY2", "SUMMSUMMKV")
    result = ReplaceFunc(result, "SUMXMY2", "SUM DIFFERENCE")
    result = ReplaceFunc(result, "SUM", "SUM")
    result = ReplaceFunc(result, "TAN", "TAN")
    result = ReplaceFunc(result, "TANH", "TANH")
    result = ReplaceFunc(result, "TRUNC", "OTBR")
    
    ' === LOGICAL ===
    result = ReplaceFunc(result, "AND", "AND")
    result = ReplaceFunc(result, "FALSE", "LIE")
    result = ReplaceFunc(result, "IFERROR", "IFERROR")
    result = ReplaceFunc(result, "IFNA", "ESND")
    result = ReplaceFunc(result, "IFS", "CONDITIONS")
    result = ReplaceFunc(result, "IF", "IF")
    result = ReplaceFunc(result, "NOT", "NOT")
    result = ReplaceFunc(result, "OR", "OR")
    result = ReplaceFunc(result, "SWITCH", "SWITCH")
    result = ReplaceFunc(result, "TRUE", "TRUE")
    result = ReplaceFunc(result, "XOR", "EXCLUDED")
    
    ' === TEXT ===
    result = ReplaceFunc(result, "ASC", "ASC")
    result = ReplaceFunc(result, "BAHTTEXT", "BATT.TEXT")
    result = ReplaceFunc(result, "CHAR", "SYMBOL")
    result = ReplaceFunc(result, "CLEAN", "PECHSIMV")
    result = ReplaceFunc(result, "CODE", "CODSIM")
    result = ReplaceFunc(result, "CONCATENATE", "CONNECT")
    result = ReplaceFunc(result, "CONCAT", "SCENER")
    result = ReplaceFunc(result, "DOLLAR", "RUBLE")
    result = ReplaceFunc(result, "EXACT", "COINCIDENCE")
    result = ReplaceFunc(result, "FIND", "FIND")
    result = ReplaceFunc(result, "FIXED", "FIXED")
    result = ReplaceFunc(result, "LEFT", "LEVSIMV")
    result = ReplaceFunc(result, "LEN", "DLST")
    result = ReplaceFunc(result, "LOWER", "LOWER")
    result = ReplaceFunc(result, "MID", "PSTR")
    result = ReplaceFunc(result, "NUMBERVALUE", "VALUE")
    result = ReplaceFunc(result, "PHONETIC", "PHONETIC")
    result = ReplaceFunc(result, "PROPER", "PROPNACH")
    result = ReplaceFunc(result, "REPLACE", "REPLACE")
    result = ReplaceFunc(result, "REPT", "REPEAT")
    result = ReplaceFunc(result, "RIGHT", "RIGHT")
    result = ReplaceFunc(result, "SEARCH", "SEARCH")
    result = ReplaceFunc(result, "SUBSTITUTE", "SUBSTITUTE")
    result = ReplaceFunc(result, "TEXTJOIN", "COMBINE")
    result = ReplaceFunc(result, "TEXT", "TEXT")
    result = ReplaceFunc(result, "TRIM", "SPACE")
    result = ReplaceFunc(result, "UNICHAR", "UNISIM")
    result = ReplaceFunc(result, "UNICODE", "UNICODE")
    result = ReplaceFunc(result, "UPPER", "CAPITAL")
    result = ReplaceFunc(result, "VALUE", "SIGNIFICANT")
    
    ' === DATE AND TIME ===
    result = ReplaceFunc(result, "DATE", "DATE")
    result = ReplaceFunc(result, "DATEDIF", "RAZNDAT")
    result = ReplaceFunc(result, "DATEVALUE", "DATEVALUE")
    result = ReplaceFunc(result, "DAY", "DAY")
    result = ReplaceFunc(result, "DAYS360", "DAYS360")
    result = ReplaceFunc(result, "DAYS", "DAYS")
    result = ReplaceFunc(result, "EDATE", "DATAMES")
    result = ReplaceFunc(result, "EOMONTH", "EON-MONTH")
    result = ReplaceFunc(result, "HOUR", "HOUR")
    result = ReplaceFunc(result, "ISOWEEKNUM", "WEEK NUMBER.ISO")
    result = ReplaceFunc(result, "MINUTE", "MINUTES")
    result = ReplaceFunc(result, "MONTH", "MONTH")
    result = ReplaceFunc(result, "NETWORKDAYS.INTL", "NETWORKDAYS.INTL")
    result = ReplaceFunc(result, "NETWORKDAYS", "NETWORKDAYS")
    result = ReplaceFunc(result, "NOW", "TDATE")
    result = ReplaceFunc(result, "SECOND", "SECONDS")
    result = ReplaceFunc(result, "TIMEVALUE", "TIMEVALUE")
    result = ReplaceFunc(result, "TIME", "TIME")
    result = ReplaceFunc(result, "TODAY", "TODAY")
    result = ReplaceFunc(result, "WEEKDAY", "WEEKDAY")
    result = ReplaceFunc(result, "WEEKNUM", "WEEK NUMBER")
    result = ReplaceFunc(result, "WORKDAY.INTL", "WORKDAY INTERNATIONAL")
    result = ReplaceFunc(result, "WORKDAY", "WORKDAY")
    result = ReplaceFunc(result, "YEARFRAC", "PERCENTAGE OF THE YEAR")
    result = ReplaceFunc(result, "YEAR", "YEAR")
    
    ' === LINKS AND SEARCH ===
    result = ReplaceFunc(result, "ADDRESS", "ADDRESS")
    result = ReplaceFunc(result, "AREAS", "AREAS")
    result = ReplaceFunc(result, "CHOOSE", "CHOICE")
    result = ReplaceFunc(result, "COLUMNS", "NUMBERCOLUMN")
    result = ReplaceFunc(result, "COLUMN", "COLUMN")
    result = ReplaceFunc(result, "FORMULATEXT", "F.TEXT")
    result = ReplaceFunc(result, "GETPIVOTDATA", "GET.PICTTABLE.DATA")
    result = ReplaceFunc(result, "HLOOKUP", "GPR")
    result = ReplaceFunc(result, "HYPERLINK", "HYPERLINK")
    result = ReplaceFunc(result, "INDEX", "INDEX")
    result = ReplaceFunc(result, "INDIRECT", "INDIRECT")
    result = ReplaceFunc(result, "LOOKUP", "VIEW")
    result = ReplaceFunc(result, "MATCH", "SEARCH")
    result = ReplaceFunc(result, "OFFSET", "OFFSET")
    result = ReplaceFunc(result, "ROWS", "LINE")
    result = ReplaceFunc(result, "ROW", "LINE")
    result = ReplaceFunc(result, "RTD", "DRV")
    result = ReplaceFunc(result, "TRANSPOSE", "TRANSSP")
    result = ReplaceFunc(result, "VLOOKUP", "VLOOKUP")
    result = ReplaceFunc(result, "XLOOKUP", "VIEWX")
    result = ReplaceFunc(result, "XMATCH", "MATCHX")
    
    ' === STATISTICAL ===
    result = ReplaceFunc(result, "AVEDEV", "SROTCL")
    result = ReplaceFunc(result, "AVERAGEIFS", "AVERAGEIFS")
    result = ReplaceFunc(result, "AVERAGEIF", "AVERAGEIF")
    result = ReplaceFunc(result, "AVERAGEA", "AVERAGE")
    result = ReplaceFunc(result, "AVERAGE", "AVERAGE")
    result = ReplaceFunc(result, "BETA.DIST", "BETA.DIST")
    result = ReplaceFunc(result, "BETA.INV", "BETA.OBR")
    result = ReplaceFunc(result, "BINOM.DIST.RANGE", "BINOM.DIST.RANGE")
    result = ReplaceFunc(result, "BINOM.DIST", "BINOM.DIST")
    result = ReplaceFunc(result, "BINOM.INV", "BINOM.OBR")
    result = ReplaceFunc(result, "CHISQ.DIST.RT", "CH2.DIST.PH")
    result = ReplaceFunc(result, "CHISQ.DIST", "CH2.DIST")
    result = ReplaceFunc(result, "CHISQ.INV.RT", "CH2.OBR.PH")
    result = ReplaceFunc(result, "CHISQ.INV", "CH2.OBR")
    result = ReplaceFunc(result, "CHISQ.TEST", "CHI2.TEST")
    result = ReplaceFunc(result, "CONFIDENCE.NORM", "TRUST.NORM")
    result = ReplaceFunc(result, "CONFIDENCE.T", "TRUSTEE.STUDENT")
    result = ReplaceFunc(result, "CORREL", "CORREL")
    result = ReplaceFunc(result, "COUNTA", "COUNTING")
    result = ReplaceFunc(result, "COUNTBLANK", "COUNT VOIDS")
    result = ReplaceFunc(result, "COUNTIFS", "COUNTIFS")
    result = ReplaceFunc(result, "COUNTIF", "COUNTIF")
    result = ReplaceFunc(result, "COUNT", "CHECK")
    result = ReplaceFunc(result, "COVARIANCE.P", "COVARIANCE.G")
    result = ReplaceFunc(result, "COVARIANCE.S", "COVARIANCE.B")
    result = ReplaceFunc(result, "DEVSQ", "QUADROTCL")
    result = ReplaceFunc(result, "EXPON.DIST", "EXP.DIST.")
    result = ReplaceFunc(result, "F.DIST.RT", "F.DIST.PH")
    result = ReplaceFunc(result, "F.DIST", "F.DIST")
    result = ReplaceFunc(result, "F.INV.RT", "F.REV.PH")
    result = ReplaceFunc(result, "F.INV", "F.OBR")
    result = ReplaceFunc(result, "FISHER", "FISCHER")
    result = ReplaceFunc(result, "FISHERINV", "FISHEROBR")
    result = ReplaceFunc(result, "FORECAST.ETS.CONFINT", "FORECAST.ETS.DEVINTERVAL")
    result = ReplaceFunc(result, "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.SEASONALITY")
    result = ReplaceFunc(result, "FORECAST.ETS.STAT", "FORECAST.ETS.STAT")
    result = ReplaceFunc(result, "FORECAST.ETS", "FORECAST.ETS")
    result = ReplaceFunc(result, "FORECAST.LINEAR", "PREDICTION")
    result = ReplaceFunc(result, "FORECAST", "PREDICTION")
    result = ReplaceFunc(result, "FREQUENCY", "FREQUENCY")
    result = ReplaceFunc(result, "F.TEST", "F.TEST")
    result = ReplaceFunc(result, "GAMMA.DIST", "GAMMA.DIST")
    result = ReplaceFunc(result, "GAMMA.INV", "GAMMA.OBR")
    result = ReplaceFunc(result, "GAMMALN.PRECISE", "GAMMAL ACCURACY")
    result = ReplaceFunc(result, "GAMMALN", "GAMMALN")
    result = ReplaceFunc(result, "GAMMA", "GAMMA")
    result = ReplaceFunc(result, "GAUSS", "GAUSS")
    result = ReplaceFunc(result, "GEOMEAN", "SRGEOM")
    result = ReplaceFunc(result, "GROWTH", "HEIGHT")
    result = ReplaceFunc(result, "HARMEAN", "SRGARM")
    result = ReplaceFunc(result, "HYPGEOM.DIST", "HYPERGEOM.DIST")
    result = ReplaceFunc(result, "INTERCEPT", "CUT")
    result = ReplaceFunc(result, "KURT", "EXCESS")
    result = ReplaceFunc(result, "LARGE", "BIGGEST")
    result = ReplaceFunc(result, "LINEST", "LINEST")
    result = ReplaceFunc(result, "LOGEST", "LGRFPRIBL")
    result = ReplaceFunc(result, "LOGNORM.DIST", "LOGNORM.DIST")
    result = ReplaceFunc(result, "LOGNORM.INV", "LOGNORM.REV")
    result = ReplaceFunc(result, "MAXA", "MAXA")
    result = ReplaceFunc(result, "MAXIFS", "MAXESLIMN")
    result = ReplaceFunc(result, "MAX", "MAX")
    result = ReplaceFunc(result, "MEDIAN", "MEDIAN")
    result = ReplaceFunc(result, "MINA", "MINE")
    result = ReplaceFunc(result, "MINIFS", "MINESLIMN")
    result = ReplaceFunc(result, "MIN", "MIN")
    result = ReplaceFunc(result, "MODE.MULT", "MODA.NSK")
    result = ReplaceFunc(result, "MODE.SNGL", "FASHION.ONE")
    result = ReplaceFunc(result, "MODE", "FASHION")
    result = ReplaceFunc(result, "NEGBINOM.DIST", "OTRBINOM.DIST.")
    result = ReplaceFunc(result, "NORM.DIST", "NORMAL DIST.")
    result = ReplaceFunc(result, "NORM.INV", "NORM.REV")
    result = ReplaceFunc(result, "NORM.S.DIST", "NORM.ST.DIST.")
    result = ReplaceFunc(result, "NORM.S.INV", "NORM.ST.REV")
    result = ReplaceFunc(result, "PEARSON", "PEARSON")
    result = ReplaceFunc(result, "PERCENTILE.EXC", "PERCENTILE.EXC.")
    result = ReplaceFunc(result, "PERCENTILE.INC", "PERCENTILE.ON")
    result = ReplaceFunc(result, "PERCENTILE", "PERCENTILE")
    result = ReplaceFunc(result, "PERCENTRANK.EXC", "PERCENTRANK.EXCL.")
    result = ReplaceFunc(result, "PERCENTRANK.INC", "PERCENTRANK.ON")
    result = ReplaceFunc(result, "PERCENTRANK", "PERCENTRANK")
    result = ReplaceFunc(result, "PERMUT", "STOP")
    result = ReplaceFunc(result, "PERMUTATIONA", "STOP")
    result = ReplaceFunc(result, "PHI", "FI")
    result = ReplaceFunc(result, "POISSON.DIST", "POISSON.DIST.")
    result = ReplaceFunc(result, "PROB", "PROBABILITY")
    result = ReplaceFunc(result, "QUARTILE.EXC", "QUARTILE.EXCL.")
    result = ReplaceFunc(result, "QUARTILE.INC", "QUARTILE.ON")
    result = ReplaceFunc(result, "QUARTILE", "QUARTILE")
    result = ReplaceFunc(result, "RANK.AVG", "RANK.SR")
    result = ReplaceFunc(result, "RANK.EQ", "RANK.RV")
    result = ReplaceFunc(result, "RANK", "RANK")
    result = ReplaceFunc(result, "RSQ", "KVPIERSON")
    result = ReplaceFunc(result, "SKEW.P", "SKOS.G")
    result = ReplaceFunc(result, "SKEW", "SKOS")
    result = ReplaceFunc(result, "SLOPE", "INCLINE")
    result = ReplaceFunc(result, "SMALL", "LEAST")
    result = ReplaceFunc(result, "STANDARDIZE", "NORMALIZATION")
    result = ReplaceFunc(result, "STDEV.P", "STDEV.G")
    result = ReplaceFunc(result, "STDEV.S", "STDEV.V")
    result = ReplaceFunc(result, "STDEVA", "STDEV")
    result = ReplaceFunc(result, "STDEVPA", "STDEV")
    result = ReplaceFunc(result, "STDEVP", "STDEV")
    result = ReplaceFunc(result, "STDEV", "STANDARD DEVIATION")
    result = ReplaceFunc(result, "STEYX", "STOSHYX")
    result = ReplaceFunc(result, "T.DIST.2T", "STUDENT.DIST.2X")
    result = ReplaceFunc(result, "T.DIST.RT", "STUDENT.DIST.PH")
    result = ReplaceFunc(result, "T.DIST", "STUDENT.DIST")
    result = ReplaceFunc(result, "TREND", "TREND")
    result = ReplaceFunc(result, "TRIMMEAN", "CURRENT AVERAGE")
    result = ReplaceFunc(result, "T.INV.2T", "STUDENT.OBR.2X")
    result = ReplaceFunc(result, "T.INV", "STUDENT.OBR")
    result = ReplaceFunc(result, "T.TEST", "STUDENT TEST")
    result = ReplaceFunc(result, "VAR.P", "DISP.G")
    result = ReplaceFunc(result, "VAR.S", "DISP.B")
    result = ReplaceFunc(result, "VARA", "DISPA")
    result = ReplaceFunc(result, "VARPA", "DISPRA")
    result = ReplaceFunc(result, "VARP", "DISPR")
    result = ReplaceFunc(result, "VAR", "DISP")
    result = ReplaceFunc(result, "WEIBULL.DIST", "WEIBULL.DIST")
    result = ReplaceFunc(result, "Z.TEST", "Z.TEST")
    
    ' === INFORMATIONAL ===
    result = ReplaceFunc(result, "CELL", "CELL")
    result = ReplaceFunc(result, "ERROR.TYPE", "ERROR TYPE")
    result = ReplaceFunc(result, "INFO", "INFORM")
    result = ReplaceFunc(result, "ISBLANK", "EMPTY")
    result = ReplaceFunc(result, "ISERR", "EOS")
    result = ReplaceFunc(result, "ISERROR", "ERROR")
    result = ReplaceFunc(result, "ISEVEN", "EVEN")
    result = ReplaceFunc(result, "ISFORMULA", "FORMULA")
    result = ReplaceFunc(result, "ISLOGICAL", "ELOGIC")
    result = ReplaceFunc(result, "ISNA", "UNM")
    result = ReplaceFunc(result, "ISNONTEXT", "ENETEXT")
    result = ReplaceFunc(result, "ISNUMBER", "ISNUMBER")
    result = ReplaceFunc(result, "ISODD", "NUTS")
    result = ReplaceFunc(result, "ISREF", "LINK")
    result = ReplaceFunc(result, "ISTEXT", "ETEXT")
    result = ReplaceFunc(result, "NA", "ND")
    result = ReplaceFunc(result, "SHEET", "SHEET")
    result = ReplaceFunc(result, "SHEETS", "SHEETS")
    result = ReplaceFunc(result, "TYPE", "TYPE")
    
    ' === FINANCIAL ===
    result = ReplaceFunc(result, "ACCRINT", "ACCUMULATED INCOME")
    result = ReplaceFunc(result, "ACCRINTM", "ACCUMULATED INCOME REDEMPTION")
    result = ReplaceFunc(result, "AMORDEGRC", "AMORUM")
    result = ReplaceFunc(result, "AMORLINC", "AMORUV")
    result = ReplaceFunc(result, "COUPDAYBS", "DAYSKUPONDO")
    result = ReplaceFunc(result, "COUPDAYS", "DAYSCOUPON")
    result = ReplaceFunc(result, "COUPDAYSNC", "DAYSCOUPONAFTER")
    result = ReplaceFunc(result, "COUPNCD", "DATECOUPONAFTER")
    result = ReplaceFunc(result, "COUPNUM", "COUPON NUMBER")
    result = ReplaceFunc(result, "COUPPCD", "DATACOUPONDO")
    result = ReplaceFunc(result, "CUMIPMT", "GENERAL PAYMENT")
    result = ReplaceFunc(result, "CUMPRINC", "TOTAL INCOME")
    result = ReplaceFunc(result, "DB", "FOO")
    result = ReplaceFunc(result, "DDB", "DDOB")
    result = ReplaceFunc(result, "DISC", "DISCOUNT")
    result = ReplaceFunc(result, "DOLLARDE", "RUBLE.DES")
    result = ReplaceFunc(result, "DOLLARFR", "RUBLE.FRACTION")
    result = ReplaceFunc(result, "DURATION", "DURATION")
    result = ReplaceFunc(result, "EFFECT", "EFFECT")
    result = ReplaceFunc(result, "FV", "BS")
    result = ReplaceFunc(result, "FVSCHEDULE", "BZSCHEDULE")
    result = ReplaceFunc(result, "INTRATE", "INORMA")
    result = ReplaceFunc(result, "IPMT", "PRPLT")
    result = ReplaceFunc(result, "IRR", "VSD")
    result = ReplaceFunc(result, "ISPMT", "PROCESS PAYMENT")
    result = ReplaceFunc(result, "MDURATION", "MDLIT")
    result = ReplaceFunc(result, "MIRR", "MVSD")
    result = ReplaceFunc(result, "NOMINAL", "RATING")
    result = ReplaceFunc(result, "NPER", "NPER")
    result = ReplaceFunc(result, "NPV", "NPV")
    result = ReplaceFunc(result, "ODDFPRICE", "PRICE UPON REGULAR")
    result = ReplaceFunc(result, "ODDFYIELD", "INCOMEPERVERNERG")
    result = ReplaceFunc(result, "ODDLPRICE", "PRICE REGULAR")
    result = ReplaceFunc(result, "ODDLYIELD", "INCOME AFTER REGULAR")
    result = ReplaceFunc(result, "PDURATION", "PDLIT")
    result = ReplaceFunc(result, "PMT", "PLT")
    result = ReplaceFunc(result, "PPMT", "OSPLT")
    result = ReplaceFunc(result, "PRICEDISC", "PRICE DISCOUNT")
    result = ReplaceFunc(result, "PRICEMAT", "PRICE CASH")
    result = ReplaceFunc(result, "PRICE", "PRICE")
    result = ReplaceFunc(result, "PV", "PS")
    result = ReplaceFunc(result, "RATE", "BID")
    result = ReplaceFunc(result, "RECEIVED", "RECEIVED")
    result = ReplaceFunc(result, "RRI", "EQ.RATE")
    result = ReplaceFunc(result, "SLN", "nuclear submarine")
    result = ReplaceFunc(result, "SYD", "ASCH")
    result = ReplaceFunc(result, "TBILLEQ", "RAVNOCHEK")
    result = ReplaceFunc(result, "TBILLPRICE", "PRICECHECK")
    result = ReplaceFunc(result, "TBILLYIELD", "INCOMECHECK")
    result = ReplaceFunc(result, "VDB", "POO")
    result = ReplaceFunc(result, "XIRR", "CHISTVNDOH")
    result = ReplaceFunc(result, "XNPV", "CHISTNZ")
    result = ReplaceFunc(result, "YIELDDISC", "INCOME DISCOUNT")
    result = ReplaceFunc(result, "YIELDMAT", "INCOME REDEMPTION")
    result = ReplaceFunc(result, "YIELD", "INCOME")
    
    ' === ENGINEERING ===
    result = ReplaceFunc(result, "BESSELI", "BESSEL.I")
    result = ReplaceFunc(result, "BESSELJ", "BESSEL.J")
    result = ReplaceFunc(result, "BESSELK", "BESSEL.K")
    result = ReplaceFunc(result, "BESSELY", "BESSEL.Y")
    result = ReplaceFunc(result, "BIN2DEC", "DV.V.DES")
    result = ReplaceFunc(result, "BIN2HEX", "DV.H.HEX")
    result = ReplaceFunc(result, "BIN2OCT", "DV.V.EIGHT")
    result = ReplaceFunc(result, "BITAND", "BIT.I")
    result = ReplaceFunc(result, "BITLSHIFT", "BIT.SHIFT")
    result = ReplaceFunc(result, "BITOR", "BIT.OR")
    result = ReplaceFunc(result, "BITRSHIFT", "BIT.SHIFT")
    result = ReplaceFunc(result, "BITXOR", "BIT.ORPED")
    result = ReplaceFunc(result, "COMPLEX", "COMPLEX")
    result = ReplaceFunc(result, "CONVERT", "CONVERT")
    result = ReplaceFunc(result, "DEC2BIN", "DES.V.DV")
    result = ReplaceFunc(result, "DEC2HEX", "DES.HEX")
    result = ReplaceFunc(result, "DEC2OCT", "DES.V.EIGHT")
    result = ReplaceFunc(result, "DELTA", "DELTA")
    result = ReplaceFunc(result, "ERF.PRECISE", "FOS.EXACT")
    result = ReplaceFunc(result, "ERFC.PRECISE", "DFOSH.PRECISION")
    result = ReplaceFunc(result, "ERFC", "DFOSH")
    result = ReplaceFunc(result, "ERF", "FOS")
    result = ReplaceFunc(result, "GESTEP", "THRESHOLD")
    result = ReplaceFunc(result, "HEX2BIN", "HEX H. DW")
    result = ReplaceFunc(result, "HEX2DEC", "HEX.V.DES")
    result = ReplaceFunc(result, "HEX2OCT", "HEX.EIGHT")
    result = ReplaceFunc(result, "IMABS", "IMAG.ABS")
    result = ReplaceFunc(result, "IMAGINARY", "IMAGINARY PART")
    result = ReplaceFunc(result, "IMARGUMENT", "IMAGINAL ARGUMENT")
    result = ReplaceFunc(result, "IMCONJUGATE", "IMAGINAL MATE")
    result = ReplaceFunc(result, "IMCOS", "IMIM.COS")
    result = ReplaceFunc(result, "IMCOSH", "INSIM.COSH")
    result = ReplaceFunc(result, "IMCOT", "INSIM.COT")
    result = ReplaceFunc(result, "IMCSC", "IMIM.CSC")
    result = ReplaceFunc(result, "IMCSCH", "IMIM.CSCH")
    result = ReplaceFunc(result, "IMDIV", "IMAGINAL CASE")
    result = ReplaceFunc(result, "IMEXP", "IMAG.EXP")
    result = ReplaceFunc(result, "IMLN", "IMAG.LN")
    result = ReplaceFunc(result, "IMLOG10", "IMAG.LOG10")
    result = ReplaceFunc(result, "IMLOG2", "IMAG.LOG2")
    result = ReplaceFunc(result, "IMPOWER", "IMAGINARY DEGREE")
    result = ReplaceFunc(result, "IMPRODUCT", "IMAGINAL PRODUCT")
    result = ReplaceFunc(result, "IMREAL", "IMAGINAL THINGS")
    result = ReplaceFunc(result, "IMSEC", "IMAG.SEC")
    result = ReplaceFunc(result, "IMSECH", "INSIM.SECH")
    result = ReplaceFunc(result, "IMSIN", "IMAG.SIN")
    result = ReplaceFunc(result, "IMSINH", "IMAG.SINH")
    result = ReplaceFunc(result, "IMSQRT", "IMAGINAL ROOT")
    result = ReplaceFunc(result, "IMSUB", "IMAGINAL DIFFERENCE")
    result = ReplaceFunc(result, "IMSUM", "IMAG.SUM")
    result = ReplaceFunc(result, "IMTAN", "IMAG.TAN")
    result = ReplaceFunc(result, "OCT2BIN", "EIGHT.H.DV")
    result = ReplaceFunc(result, "OCT2DEC", "EIGHT.V.DES")
    result = ReplaceFunc(result, "OCT2HEX", "EIGHT.HEX")
    
    ' === DATABASES ===
    result = ReplaceFunc(result, "DAVERAGE", "DSRVALUE")
    result = ReplaceFunc(result, "DCOUNT", "COUNT")
    result = ReplaceFunc(result, "DCOUNTA", "ACCOUNTS")
    result = ReplaceFunc(result, "DGET", "BIZVLECH")
    result = ReplaceFunc(result, "DMAX", "DMAX")
    result = ReplaceFunc(result, "DMIN", "DMIN")
    result = ReplaceFunc(result, "DPRODUCT", "BDPRODUCT")
    result = ReplaceFunc(result, "DSTDEVP", "DSTDEV")
    result = ReplaceFunc(result, "DSTDEV", "DSTANDOFF")
    result = ReplaceFunc(result, "DSUM", "BDSUMM")
    result = ReplaceFunc(result, "DVARP", "BDDISPP")
    result = ReplaceFunc(result, "DVAR", "BDDISP")
    
    ' === WEB ===
    result = ReplaceFunc(result, "ENCODEURL", "ENCODINGURL")
    result = ReplaceFunc(result, "FILTERXML", "FILTER.XML")
    result = ReplaceFunc(result, "WEBSERVICE", "WEB SERVICE")
    
    ' === DYNAMIC ARRAYS (Excel 365) ===
    result = ReplaceFunc(result, "FILTER", "FILTER")
    result = ReplaceFunc(result, "RANDARRAY", "RANDARY")
    result = ReplaceFunc(result, "SEQUENCE", "AFTERNOON")
    result = ReplaceFunc(result, "SORTBY", "SORTPO")
    result = ReplaceFunc(result, "SORT", "GRADE")
    result = ReplaceFunc(result, "UNIQUE", "UNIQ")
    
    ' === OTHER ===
    result = ReplaceFunc(result, "EUROCONVERT", "EURO")
    result = ReplaceFunc(result, "N", "H")
    result = ReplaceFunc(result, "T", "T")
    
    ' Replacing TRUE/FALSE constants
    result = ReplaceConstant(result, "TRUE", "TRUE")
    result = ReplaceConstant(result, "FALSE", "LIE")
    
    ' Replace the argument separator (comma -> semicolon for Russian locale)
    result = Replace(result, ",", Application.International(xlListSeparator))
    
    LocalizeFormula = result
End Function

'----------------------------------------
' Convert a column letter or number to a column number relative to a range
' colRef - can be a number (1, 2, 3) or a letter (A, B, C)
' rng - range relative to which the number is calculated
' Returns the column number in a range (1-based)
'----------------------------------------
Private Function GetColumnNumber(colRef As String, rng As Range, ws As Worksheet) As Long
    Dim col As String
    col = Trim(colRef)
    
    If IsNumeric(col) Then
        ' Already a number - return it as is
        GetColumnNumber = CLng(col)
    Else
        ' Column letter - calculate the number relative to the beginning of the range
        On Error Resume Next
        Dim absColNum As Long
        absColNum = ws.columns(col).Column ' Absolute column number (A=1, B=2...)
        
        If absColNum > 0 Then
            ' Calculate the relative number in a range
            GetColumnNumber = absColNum - rng.Column + 1
            If GetColumnNumber < 1 Then GetColumnNumber = 1
            If GetColumnNumber > rng.columns.Count Then GetColumnNumber = rng.columns.Count
        Else
            GetColumnNumber = 1 ' Default first column
        End If
        On Error GoTo 0
    End If
End Function

'----------------------------------------
' Replacing the function name (taking into account that this is a function, not part of the text)
'----------------------------------------
Private Function ReplaceFunc(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String
    
    result = formula
    pos = InStr(1, UCase(result), UCase(engName) & "(")
    
    Do While pos > 0
        ' We check that there is no letter before the function name (so as not to replace part of another function)
        If pos = 1 Then
            result = rusName & Mid(result, pos + Len(engName))
        Else
            before = Mid(result, pos - 1, 1)
            If Not (before >= "A" And before <= "Z") And Not (before >= "a" And before <= "z") And Not (before >= "A" And before <= "I") Then
                result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            End If
        End If
        pos = InStr(pos + Len(rusName), UCase(result), UCase(engName) & "(")
    Loop
    
    ReplaceFunc = result
End Function

'----------------------------------------
' Constant replacement (TRUE, FALSE) - without requiring a parenthesis after the name
'----------------------------------------
Private Function ReplaceConstant(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String
    Dim after As String
    Dim isWordBoundary As Boolean
    
    result = formula
    pos = InStr(1, UCase(result), UCase(engName))
    
    Do While pos > 0
        isWordBoundary = True
        
        ' Checking the symbol before
        If pos > 1 Then
            before = Mid(result, pos - 1, 1)
            If (before >= "A" And before <= "Z") Or (before >= "a" And before <= "z") Or (before >= "A" And before <= "I") Or (before >= "0" And before <= "9") Then
                isWordBoundary = False
            End If
        End If
        
        ' Checking the symbol after
        If pos + Len(engName) <= Len(result) Then
            after = Mid(result, pos + Len(engName), 1)
            If (after >= "A" And after <= "Z") Or (after >= "a" And after <= "z") Or (after >= "A" And after <= "I") Or (after >= "0" And after <= "9") Then
                isWordBoundary = False
            End If
        End If
        
        If isWordBoundary Then
            result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            pos = InStr(pos + Len(rusName), UCase(result), UCase(engName))
        Else
            pos = InStr(pos + 1, UCase(result), UCase(engName))
        End If
    Loop
    
    ReplaceConstant = result
End Function

'----------------------------------------
' Getting the graph type
'----------------------------------------
Private Function GetChartType(chartTypeName As String) As Long
    Select Case UCase(Trim(chartTypeName))
        Case "LINE": GetChartType = 4 ' xlLine
        Case "BAR": GetChartType = 57 ' xlBarClustered
        Case "COLUMN": GetChartType = 51 ' xlColumnClustered
        Case "PIE": GetChartType = 5 ' xlPie
        Case "AREA": GetChartType = 1 ' xlArea
        Case "SCATTER", "XY": GetChartType = -4169 ' xlXYScatter
        Case "DOUGHNUT": GetChartType = -4120 ' xlDoughnut
        Case "RADAR": GetChartType = -4151 ' xlRadar
        Case "SURFACE": GetChartType = 83 ' xlSurface
        Case "BUBBLE": GetChartType = 15 ' xlBubble
        Case "STOCK": GetChartType = 88 ' xlStockHLC
        Case "CYLINDER": GetChartType = 95 ' xlCylinderCol
        Case "CONE": GetChartType = 99 ' xlConeCol
        Case "PYRAMID": GetChartType = 103 ' xlPyramidCol
        Case "LINE_MARKERS": GetChartType = 65 ' xlLineMarkers
        Case "AREA_STACKED": GetChartType = 76 ' xlAreaStacked
        Case "BAR_STACKED": GetChartType = 58 ' xlBarStacked
        Case "COLUMN_STACKED": GetChartType = 52 ' xlColumnStacked
        Case Else: GetChartType = 4 ' xlLine by default
    End Select
End Function

'----------------------------------------
' Find a pivot table by name in all sheets
'----------------------------------------
Private Function FindPivotTable(pivotName As String) As pivotTable
    Dim wsSearch As Worksheet
    Dim ptSearch As pivotTable
    Dim searchName As String
    
    searchName = Trim(pivotName)
    Set FindPivotTable = Nothing
    
    On Error Resume Next
    ' Search all workbook sheets
    For Each wsSearch In ActiveWorkbook.Worksheets
        For Each ptSearch In wsSearch.PivotTables
            If ptSearch.Name = searchName Then
                Set FindPivotTable = ptSearch
                Exit Function
            End If
        Next ptSearch
    Next wsSearch
    On Error GoTo 0
End Function

'----------------------------------------
' Get chart index
' Supports: 0, LAST, NEW = last chart; 1,2,3... = specific index
'----------------------------------------
Private Function GetChartIndex(ws As Worksheet, indexStr As String) As Long
    Dim idx As Long
    Dim s As String
    
    s = UCase(Trim(indexStr))
    
    ' If 0, LAST, NEW, or empty: return last chart
    If s = "0" Or s = "LAST" Or s = "NEW" Or s = "" Then
        GetChartIndex = ws.ChartObjects.Count
        Exit Function
    End If
    
    ' Let's try to convert it to a number
    On Error Resume Next
    idx = CLng(s)
    On Error GoTo 0
    
    If idx > 0 Then
        GetChartIndex = idx
    Else
        ' Default - last
        GetChartIndex = ws.ChartObjects.Count
    End If
End Function

'----------------------------------------
' Execute one command (FULL VERSION)
'----------------------------------------
Private Function ExecuteSingleCommand(cmd As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim parts() As String
    Dim action As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim i As Long
    
    parts = Split(cmd, "|")
    If UBound(parts) < 0 Then
        ExecuteSingleCommand = False
        Exit Function
    End If
    
    action = UCase(Trim(parts(0)))
    Set ws = ActiveSheet
    
    ' Debugging: showing the command and number of parts
    Debug.Print "CMD: " & action & " | Parts: " & (UBound(parts) + 1) & " | Full: " & cmd
    
    Select Case action
    
        ' ========== WORKING WITH CELLS ==========
        
        Case "SET_VALUE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).value = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "SET_FORMULA"
            If UBound(parts) >= 2 Then
                Dim localFormula As String
                localFormula = LocalizeFormula(parts(2))
                Debug.Print "SET_FORMULA: Original=" & parts(2)
                Debug.Print "SET_FORMULA: Localized=" & localFormula
                ' Using FormulaLocal for Russian formulas
                ws.Range(parts(1)).FormulaLocal = localFormula
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_DOWN"
            If UBound(parts) >= 2 Then
                Dim srcCell As Range, destRng As Range
                Set srcCell = ws.Range(parts(1))
                Set destRng = ws.Range(parts(1) & ":" & parts(2))
                srcCell.Copy
                destRng.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_RIGHT"
            If UBound(parts) >= 2 Then
                Dim srcCellR As Range, destRngR As Range
                Set srcCellR = ws.Range(parts(1))
                Set destRngR = ws.Range(parts(1) & ":" & parts(2))
                srcCellR.Copy
                destRngR.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_SERIES"
            If UBound(parts) >= 2 Then
                Dim stepVal As Double
                stepVal = 1
                If UBound(parts) >= 2 Then stepVal = CDbl(parts(2))
                ws.Range(parts(1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=stepVal
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_CONTENTS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearContents
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearFormats
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_ALL"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Clear
                ExecuteSingleCommand = True
            End If
            
        Case "COPY"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "CUT"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Cut Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "PASTE_VALUES"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "TRANSPOSE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteAll, Transpose:=True
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
        
        ' ========== FORMATTING ==========
            
        Case "BOLD"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Bold = True
                ExecuteSingleCommand = True
            End If
            
        Case "ITALIC"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Italic = True
                ExecuteSingleCommand = True
            End If
            
        Case "UNDERLINE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Underline = xlUnderlineStyleSingle
                ExecuteSingleCommand = True
            End If
            
        Case "STRIKETHROUGH"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Strikethrough = True
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_NAME"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_SIZE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Size = CLng(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "BORDER"
            If UBound(parts) >= 2 Then
                Dim borderStyle As String
                borderStyle = UCase(Trim(parts(2)))
                With ws.Range(parts(1))
                    Select Case borderStyle
                        Case "ALL"
                            .Borders.LineStyle = xlContinuous
                        Case "TOP"
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                        Case "BOTTOM"
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        Case "LEFT"
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Case "RIGHT"
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                        Case "NONE"
                            .Borders.LineStyle = xlNone
                    End Select
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "BORDER_THICK"
            If UBound(parts) >= 1 Then
                With ws.Range(parts(1)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "ALIGN_H"
            If UBound(parts) >= 2 Then
                Dim hAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "LEFT": hAlign = xlLeft
                    Case "CENTER": hAlign = xlCenter
                    Case "RIGHT": hAlign = xlRight
                    Case "JUSTIFY": hAlign = xlJustify
                    Case Else: hAlign = xlGeneral
                End Select
                ws.Range(parts(1)).HorizontalAlignment = hAlign
                ExecuteSingleCommand = True
            End If
            
        Case "ALIGN_V"
            If UBound(parts) >= 2 Then
                Dim vAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "TOP": vAlign = xlTop
                    Case "CENTER": vAlign = xlCenter
                    Case "BOTTOM": vAlign = xlBottom
                    Case Else: vAlign = xlCenter
                End Select
                ws.Range(parts(1)).VerticalAlignment = vAlign
                ExecuteSingleCommand = True
            End If
            
        Case "WRAP_TEXT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).WrapText = True
                ExecuteSingleCommand = True
            End If
            
        Case "MERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Merge
                ExecuteSingleCommand = True
            End If
            
        Case "UNMERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).UnMerge
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_NUMBER"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_DATE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_PERCENT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).NumberFormat = "0.00%"
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_CURRENCY"
            If UBound(parts) >= 1 Then
                Dim currSymbol As String
                currSymbol = "?"
                If UBound(parts) >= 2 Then currSymbol = parts(2)
                ws.Range(parts(1)).NumberFormat = "#,##0.00 " & currSymbol
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFIT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).columns.AutoFit
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFIT_ROWS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Rows.AutoFit
                ExecuteSingleCommand = True
            End If
            
        Case "COLUMN_WIDTH"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1)).ColumnWidth = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "ROW_HEIGHT"
            If UBound(parts) >= 2 Then
                ws.Rows(CLng(parts(1))).RowHeight = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If
        
        ' ========== ROWS AND COLUMN ==========
            
        Case "INSERT_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_ROWS"
            If UBound(parts) >= 2 Then
                Dim rowNum As Long, rowCount As Long
                rowNum = CLng(parts(1))
                rowCount = CLng(parts(2))
                ws.Rows(rowNum & ":" & (rowNum + rowCount - 1)).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim colNum As Long, colCount As Long
                colCount = CLng(parts(2))
                For i = 1 To colCount
                    ws.columns(parts(1)).Insert
                Next i
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "GROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If
            
        Case "UNGROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If
            
        Case "GROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If
            
        Case "UNGROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If
        
        ' ========== SORTING AND FILTERING ==========
            
        Case "SORT"
            If UBound(parts) >= 3 Then
                Dim SortRange As Range, sortCol As Long, sortOrder As Long
                Dim sortColStr As String, sortKeyRange As Range
                Set SortRange = ws.Range(parts(1))
                sortColStr = Trim(parts(2))
                
                ' Defining a column for sorting
                If IsNumeric(sortColStr) Then
                    ' Number - column number in the range (1, 2, 3...)
                    sortCol = CLng(sortColStr)
                    Set sortKeyRange = SortRange.columns(sortCol)
                Else
                    ' Column letter (A, B, C...) - use intersection with range
                    Set sortKeyRange = Intersect(SortRange, ws.columns(sortColStr))
                    If sortKeyRange Is Nothing Then
                        ' If the letter is out of range, take the first column
                        Set sortKeyRange = SortRange.columns(1)
                    End If
                End If
                
                sortOrder = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                SortRange.Sort Key1:=sortKeyRange, Order1:=sortOrder, Header:=xlGuess
                ExecuteSingleCommand = True
            End If
            
        Case "SORT_MULTI"
            If UBound(parts) >= 5 Then
                Dim sRng As Range
                Dim sCol1 As Long, sCol2 As Long
                Set sRng = ws.Range(parts(1))
                Dim o1 As Long, o2 As Long
                sCol1 = GetColumnNumber(parts(2), sRng, ws)
                sCol2 = GetColumnNumber(parts(4), sRng, ws)
                o1 = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                o2 = IIf(UCase(parts(5)) = "ASC", xlAscending, xlDescending)
                sRng.Sort Key1:=sRng.columns(sCol1), Order1:=o1, _
                          Key2:=sRng.columns(sCol2), Order2:=o2, Header:=xlGuess
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                ws.Range(parts(1)).AutoFilter
                ExecuteSingleCommand = True
            End If
            
        Case "FILTER"
            If UBound(parts) >= 3 Then
                Dim fRng As Range
                Dim fCol As Long
                Set fRng = ws.Range(parts(1))
                fCol = GetColumnNumber(parts(2), fRng, ws)
                If Not ws.AutoFilterMode Then fRng.AutoFilter
                fRng.AutoFilter Field:=fCol, Criteria1:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "FILTER_TOP"
            If UBound(parts) >= 3 Then
                Dim ftRng As Range
                Dim ftCol As Long
                Set ftRng = ws.Range(parts(1))
                ftCol = GetColumnNumber(parts(2), ftRng, ws)
                If Not ws.AutoFilterMode Then ftRng.AutoFilter
                ftRng.AutoFilter Field:=ftCol, Criteria1:=CLng(parts(3)), Operator:=xlTop10Items
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_FILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then
                    ws.Range(parts(1)).AutoFilter
                    ws.Range(parts(1)).AutoFilter
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_AUTOFILTER"
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ExecuteSingleCommand = True
            
        Case "REMOVE_DUPLICATES"
            If UBound(parts) >= 2 Then
                Dim dupRng As Range
                Dim colsArr() As Long
                Dim colsList() As String
                Set dupRng = ws.Range(parts(1))
                colsList = Split(parts(2), ",")
                ReDim colsArr(UBound(colsList))
                For i = 0 To UBound(colsList)
                    colsArr(i) = CLng(Trim(colsList(i)))
                Next i
                dupRng.RemoveDuplicates columns:=colsArr, Header:=xlYes
                ExecuteSingleCommand = True
            End If
            
        Case "FIND_REPLACE"
            If UBound(parts) >= 2 Then
                Dim replaceWith As String
                replaceWith = ""
                If UBound(parts) >= 2 Then replaceWith = parts(2)
                ws.UsedRange.Replace What:=parts(1), Replacement:=replaceWith, LookAt:=xlPart
                ExecuteSingleCommand = True
            End If
            
        Case "FIND_REPLACE_RANGE"
            If UBound(parts) >= 3 Then
                ws.Range(parts(1)).Replace What:=parts(2), Replacement:=parts(3), LookAt:=xlPart
                ExecuteSingleCommand = True
            End If
        
        ' ========== CHARTS ==========
            
        Case "CREATE_CHART"
            ' CREATE_CHART|range|type|name
            ' Supports non-adjacent ranges: A2:A5,B2:B5
            If UBound(parts) >= 2 Then
                Dim chartObj As ChartObject
                Dim chartType As Long
                Dim dataRange As Range
                Dim chartLeft As Double, chartTop As Double
                Dim rangeStr As String
                Dim rangeParts() As String
                Dim rngPart As Variant
                
                rangeStr = Trim(parts(1))
                
                ' Checking for non-adjacent ranges
                On Error Resume Next
                If InStr(rangeStr, ",") > 0 Then
                    ' Non-contiguous ranges
                    rangeParts = Split(rangeStr, ",")
                    Set dataRange = ws.Range(Trim(rangeParts(0)))
                    For i = 1 To UBound(rangeParts)
                        Set dataRange = Union(dataRange, ws.Range(Trim(rangeParts(i))))
                    Next i
                Else
                    Set dataRange = ws.Range(rangeStr)
                End If
                On Error GoTo ErrorHandler
                
                If dataRange Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If
                
                chartType = GetChartType(parts(2))
                
                ' Position the graph to the right of the data
                chartLeft = dataRange.Areas(1).Cells(1, dataRange.Areas(1).columns.Count).Offset(0, 2).Left
                chartTop = dataRange.Areas(1).Top
                
                Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=400, Height:=250)
                chartObj.Chart.SetSourceData Source:=dataRange
                chartObj.Chart.chartType = chartType
                
                If UBound(parts) >= 3 Then
                    If Len(Trim(parts(3))) > 0 Then
                        chartObj.Chart.HasTitle = True
                        chartObj.Chart.ChartTitle.text = parts(3)
                    End If
                End If
                
                ExecuteSingleCommand = True
            End If
            
        Case "CREATE_CHART_POS", "CREATE_CHART_AT"
            ' CREATE_CHART_POS|range|type|title|cell_or_left|top|width|height
            ' CREATE_CHART_AT|range|type|name|cell - simplified version
            ' Supports non-adjacent ranges: A2:A5,B2:B5
            If UBound(parts) >= 3 Then
                Dim chartObj2 As ChartObject
                Dim chartType2 As Long
                Dim dataRange2 As Range
                Dim posLeft As Double, posTop As Double
                Dim chWidth As Double, chHeight As Double
                Dim rangeStr2 As String
                Dim rangeParts2() As String
                
                rangeStr2 = Trim(parts(1))
                
                ' Checking for non-adjacent ranges
                On Error Resume Next
                If InStr(rangeStr2, ",") > 0 Then
                    rangeParts2 = Split(rangeStr2, ",")
                    Set dataRange2 = ws.Range(Trim(rangeParts2(0)))
                    For i = 1 To UBound(rangeParts2)
                        Set dataRange2 = Union(dataRange2, ws.Range(Trim(rangeParts2(i))))
                    Next i
                Else
                    Set dataRange2 = ws.Range(rangeStr2)
                End If
                On Error GoTo ErrorHandler
                
                If dataRange2 Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If
                
                chartType2 = GetChartType(parts(2))
                
                ' Default values
                chWidth = 400
                chHeight = 250
                posLeft = 300
                posTop = 50
                
                ' Determining the position
                If UBound(parts) >= 4 Then
                    ' Check if this is a cell address or a number
                    On Error Resume Next
                    Dim posCell As Range
                    Set posCell = ws.Range(parts(4))
                    If Not posCell Is Nothing Then
                        ' This is the cell address
                        posLeft = posCell.Left
                        posTop = posCell.Top
                    Else
                        ' This number
                        posLeft = CDbl(parts(4))
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 5 Then
                    On Error Resume Next
                    posTop = CDbl(parts(5))
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 6 Then
                    On Error Resume Next
                    chWidth = CDbl(parts(6))
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 7 Then
                    On Error Resume Next
                    chHeight = CDbl(parts(7))
                    On Error GoTo ErrorHandler
                End If
                
                Set chartObj2 = ws.ChartObjects.Add(Left:=posLeft, Top:=posTop, Width:=chWidth, Height:=chHeight)
                chartObj2.Chart.SetSourceData Source:=dataRange2
                chartObj2.Chart.chartType = chartType2
                
                If Len(Trim(parts(3))) > 0 Then
                    chartObj2.Chart.HasTitle = True
                    chartObj2.Chart.ChartTitle.text = parts(3)
                End If
                
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_TITLE"
            ' CHART_TITLE|index|text (index: 1 = first chart, 0 or LAST = last)
            If UBound(parts) >= 2 Then
                Dim chIdx As Long
                chIdx = GetChartIndex(ws, parts(1))
                If chIdx > 0 And chIdx <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx).Chart.HasTitle = True
                    ws.ChartObjects(chIdx).Chart.ChartTitle.text = parts(2)
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_LEGEND"
            If UBound(parts) >= 2 Then
                Dim chIdx2 As Long
                chIdx2 = GetChartIndex(ws, parts(1))
                If chIdx2 > 0 And chIdx2 <= ws.ChartObjects.Count Then
                    Dim legPos As Long
                    Select Case UCase(Trim(parts(2)))
                        Case "TOP": legPos = xlLegendPositionTop
                        Case "BOTTOM": legPos = xlLegendPositionBottom
                        Case "LEFT": legPos = xlLegendPositionLeft
                        Case "RIGHT": legPos = xlLegendPositionRight
                        Case "NONE"
                            ws.ChartObjects(chIdx2).Chart.HasLegend = False
                            ExecuteSingleCommand = True
                            Exit Function
                        Case Else: legPos = xlLegendPositionBottom
                    End Select
                    ws.ChartObjects(chIdx2).Chart.HasLegend = True
                    ws.ChartObjects(chIdx2).Chart.Legend.Position = legPos
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_AXIS_TITLE"
            If UBound(parts) >= 3 Then
                Dim chIdx3 As Long
                chIdx3 = GetChartIndex(ws, parts(1))
                If chIdx3 > 0 And chIdx3 <= ws.ChartObjects.Count Then
                    On Error Resume Next
                    Dim ax As Object
                    If UCase(Trim(parts(2))) = "X" Then
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlCategory)
                    Else
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlValue)
                    End If
                    If Not ax Is Nothing Then
                        ax.HasTitle = True
                        ax.AxisTitle.text = parts(3)
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_TYPE"
            If UBound(parts) >= 2 Then
                Dim chIdx4 As Long
                chIdx4 = GetChartIndex(ws, parts(1))
                If chIdx4 > 0 And chIdx4 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx4).Chart.chartType = GetChartType(parts(2))
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_MOVE"
            If UBound(parts) >= 2 Then
                Dim chIdx5 As Long
                chIdx5 = GetChartIndex(ws, parts(1))
                If chIdx5 > 0 And chIdx5 <= ws.ChartObjects.Count Then
                    ' Checking the cell address or coordinates
                    On Error Resume Next
                    Dim moveCell As Range
                    Set moveCell = ws.Range(parts(2))
                    If Not moveCell Is Nothing Then
                        ws.ChartObjects(chIdx5).Left = moveCell.Left
                        ws.ChartObjects(chIdx5).Top = moveCell.Top
                    ElseIf UBound(parts) >= 3 Then
                        ws.ChartObjects(chIdx5).Left = CLng(parts(2))
                        ws.ChartObjects(chIdx5).Top = CLng(parts(3))
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_RESIZE"
            If UBound(parts) >= 3 Then
                Dim chIdx6 As Long
                chIdx6 = GetChartIndex(ws, parts(1))
                If chIdx6 > 0 And chIdx6 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx6).Width = CLng(parts(2))
                    ws.ChartObjects(chIdx6).Height = CLng(parts(3))
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_DELETE"
            If UBound(parts) >= 1 Then
                Dim chIdx7 As Long
                chIdx7 = GetChartIndex(ws, parts(1))
                If chIdx7 > 0 And chIdx7 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx7).Delete
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_DELETE_ALL"
            Dim co As ChartObject
            For Each co In ws.ChartObjects
                co.Delete
            Next co
            ExecuteSingleCommand = True
            
        Case "MOVE_CHART"
            ' For compatibility - skip
            ExecuteSingleCommand = True
        
        ' ========== PIVOT TABLES ==========
            
        Case "CREATE_PIVOT"
            ' CREATE_PIVOT|source|destination|name
            ' Example: CREATE_PIVOT|A1:D10|F1|MyPivot
            ' Or: CREATE_PIVOT|Sheet1!A1:D10|Sheet2!A1|MyPivot
            If UBound(parts) >= 3 Then
                Dim pivotCache As pivotCache
                Dim pivotTable As pivotTable
                Dim srcRng As Range
                Dim destCell As Range
                Dim destSheet As Worksheet
                Dim pivotName As String
                
                ' Using Application.Range to support sheet name links
                On Error Resume Next
                Set srcRng = Application.Range(parts(1))
                If srcRng Is Nothing Then
                    ' Let's try without the sheet name
                    Set srcRng = ws.Range(parts(1))
                End If
                
                ' For assignment - check whether a new sheet needs to be created
                Set destCell = Application.Range(parts(2))
                If destCell Is Nothing Then
                    ' If the sheet does not exist, create it
                    Dim destParts() As String
                    If InStr(parts(2), "!") > 0 Then
                        destParts = Split(parts(2), "!")
                        Dim sheetName As String
                        sheetName = Replace(Replace(destParts(0), "'", ""), "!", "")
                        ' Checking the existence of a sheet
                        Dim sheetExists As Boolean
                        sheetExists = False
                        Dim wsCheck As Worksheet
                        For Each wsCheck In ActiveWorkbook.Worksheets
                            If wsCheck.Name = sheetName Then
                                sheetExists = True
                                Set destSheet = wsCheck
                                Exit For
                            End If
                        Next wsCheck
                        If Not sheetExists Then
                            Set destSheet = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
                            destSheet.Name = sheetName
                        End If
                        Set destCell = destSheet.Range(destParts(1))
                    Else
                        Set destCell = ws.Range(parts(2))
                    End If
                End If
                On Error GoTo ErrorHandler
                
                If Not srcRng Is Nothing And Not destCell Is Nothing Then
                    pivotName = Trim(parts(3))
                    
                    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=srcRng)
                        
                    Set pivotTable = pivotCache.CreatePivotTable( _
                        TableDestination:=destCell, _
                        tableName:=pivotName)
                        
                    ExecuteSingleCommand = True
                End If
            End If
            
        Case "PIVOT_ADD_ROW"
            If UBound(parts) >= 2 Then
                Dim pt As pivotTable
                Set pt = FindPivotTable(parts(1))
                If Not pt Is Nothing Then
                    On Error Resume Next
                    pt.PivotFields(parts(2)).Orientation = xlRowField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_COLUMN"
            If UBound(parts) >= 2 Then
                Dim pt2 As pivotTable
                Set pt2 = FindPivotTable(parts(1))
                If Not pt2 Is Nothing Then
                    On Error Resume Next
                    pt2.PivotFields(parts(2)).Orientation = xlColumnField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_VALUE"
            If UBound(parts) >= 3 Then
                Dim pt3 As pivotTable
                Dim pfFunc As Long
                Set pt3 = FindPivotTable(parts(1))
                If Not pt3 Is Nothing Then
                    Select Case UCase(Trim(parts(3)))
                        Case "SUM": pfFunc = xlSum
                        Case "COUNT": pfFunc = xlCount
                        Case "AVERAGE": pfFunc = xlAverage
                        Case "MAX": pfFunc = xlMax
                        Case "MIN": pfFunc = xlMin
                        Case Else: pfFunc = xlSum
                    End Select
                    On Error Resume Next
                    pt3.AddDataField pt3.PivotFields(parts(2)), , pfFunc
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_FILTER"
            If UBound(parts) >= 2 Then
                Dim pt4 As pivotTable
                Set pt4 = FindPivotTable(parts(1))
                If Not pt4 Is Nothing Then
                    On Error Resume Next
                    pt4.PivotFields(parts(2)).Orientation = xlPageField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_REFRESH"
            If UBound(parts) >= 1 Then
                Dim pt5 As pivotTable
                Set pt5 = FindPivotTable(parts(1))
                If Not pt5 Is Nothing Then
                    pt5.RefreshTable
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_REFRESH_ALL"
            ActiveWorkbook.RefreshAll
            ExecuteSingleCommand = True
        
        ' ========== SHEETS ==========
            
        Case "ADD_SHEET"
            If UBound(parts) >= 1 Then
                Dim newSheet As Worksheet
                Set newSheet = ActiveWorkbook.Worksheets.Add
                newSheet.Name = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_SHEET_AFTER"
            If UBound(parts) >= 2 Then
                Dim newSheet2 As Worksheet
                Set newSheet2 = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(parts(2)))
                newSheet2.Name = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_SHEET"
            If UBound(parts) >= 1 Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Worksheets(parts(1)).Delete
                Application.DisplayAlerts = True
                ExecuteSingleCommand = True
            End If
            
        Case "RENAME_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "COPY_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Copy after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
                ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "MOVE_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Move before:=ActiveWorkbook.Worksheets(CLng(parts(2)))
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetHidden
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetVisible
                ExecuteSingleCommand = True
            End If
            
        Case "ACTIVATE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Activate
                ExecuteSingleCommand = True
            End If
            
        Case "TAB_COLOR"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Tab.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "PROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd As String
                pwd = ""
                If UBound(parts) >= 2 Then pwd = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Protect Password:=pwd
                ExecuteSingleCommand = True
            End If
            
        Case "UNPROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd2 As String
                pwd2 = ""
                If UBound(parts) >= 2 Then pwd2 = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Unprotect Password:=pwd2
                ExecuteSingleCommand = True
            End If
        
        ' ========== NAMED RANGES ==========
            
        Case "CREATE_NAME"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Names.Add Name:=parts(1), RefersTo:=ws.Range(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_NAME"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ActiveWorkbook.Names(parts(1)).Delete
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If
        
        ' ========== CONDITIONAL FORMATTING ==========
            
        Case "COND_HIGHLIGHT"
            ' COND_HIGHLIGHT|range|formula|color (4 parts)
            ' COND_HIGHLIGHT|range|operator|value|color (5 parts)
            If UBound(parts) >= 3 Then
                Dim hlRng As Range
                Dim hlFormula As String
                Dim hlFirstCell As String
                Dim hlColor As String
                Dim hlOp As String
                Dim hlFC As Object
                
                Debug.Print "COND_HIGHLIGHT: Step 1 - parsing range: " & parts(1)
                Set hlRng = ws.Range(parts(1))
                Debug.Print "COND_HIGHLIGHT: Step 2 - range set OK: " & hlRng.address
                hlFirstCell = hlRng.Cells(1, 1).address(False, False)
                Debug.Print "COND_HIGHLIGHT: Step 3 - firstCell: " & hlFirstCell
                
                If UBound(parts) = 3 Then
                    ' 4 parts: range|formula|color
                    Debug.Print "COND_HIGHLIGHT: Step 4a - 4 parts mode"
                    hlFormula = Trim(parts(2))
                    hlColor = Trim(parts(3))
                Else
                    ' 5 parts: range|operator|value|color
                    Debug.Print "COND_HIGHLIGHT: Step 4b - 5 parts mode"
                    hlOp = Trim(parts(2))
                    hlColor = Trim(parts(4))
                    Debug.Print "COND_HIGHLIGHT: Step 5 - op=" & hlOp & " color=" & hlColor
                    
                    Select Case hlOp
                        Case ">", "<", ">=", "<=", "<>"
                            hlFormula = hlFirstCell & hlOp & Trim(parts(3))
                        Case "="
                            If IsNumeric(Trim(parts(3))) Then
                                hlFormula = hlFirstCell & "=" & Trim(parts(3))
                            Else
                                hlFormula = Trim(parts(3))
                            End If
                        Case Else
                            hlFormula = Trim(parts(3))
                    End Select
                End If
                
                Debug.Print "COND_HIGHLIGHT: Step 6 - formula before =: " & hlFormula
                
                ' Add = to the beginning if not
                If Left(hlFormula, 1) <> "=" Then hlFormula = "=" & hlFormula
                
                Debug.Print "COND_HIGHLIGHT: Step 7 - formula after =: " & hlFormula
                
                ' Localize the formula (MOD -> REMAIN, etc., comma -> semicolon)
                hlFormula = LocalizeFormula(hlFormula)
                Debug.Print "COND_HIGHLIGHT: Step 8 - formula localized: " & hlFormula
                Debug.Print "COND_HIGHLIGHT: Step 9 - adding FormatCondition..."
                
                ' IMPORTANT: Select the first cell of the range so that Excel correctly
                ' interpreted relative references in formula
                hlRng.Cells(1, 1).Select
                
                ' Adding conditional formatting
                Set hlFC = hlRng.FormatConditions.Add(Type:=xlExpression, Formula1:=hlFormula)
                
                Debug.Print "COND_HIGHLIGHT: Step 10 - hlColor=[" & hlColor & "]"
                
                Dim hlColorValue As Long
                On Error Resume Next
                hlColorValue = ParseColor(hlColor)
                If Err.Number <> 0 Then
                    Debug.Print "COND_HIGHLIGHT: ParseColor ERROR: " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                Debug.Print "COND_HIGHLIGHT: Step 10a - colorValue=" & hlColorValue
                Debug.Print "COND_HIGHLIGHT: Step 10b - hlFC type=" & TypeName(hlFC)
                
                hlFC.Interior.Color = hlColorValue
                
                Debug.Print "COND_HIGHLIGHT: Step 11 - DONE"
                ExecuteSingleCommand = True
            End If
            
        Case "COND_TOP"
            If UBound(parts) >= 3 Then
                Dim cfRng2 As Range
                Set cfRng2 = ws.Range(parts(1))
                cfRng2.FormatConditions.AddTop10
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).TopBottom = xlTop10Top
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_BOTTOM"
            If UBound(parts) >= 3 Then
                Dim cfRng3 As Range
                Set cfRng3 = ws.Range(parts(1))
                cfRng3.FormatConditions.AddTop10
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).TopBottom = xlTop10Bottom
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_DUPLICATE"
            If UBound(parts) >= 2 Then
                Dim cfRng4 As Range
                Set cfRng4 = ws.Range(parts(1))
                cfRng4.FormatConditions.AddUniqueValues
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).DupeUnique = xlDuplicate
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_UNIQUE"
            If UBound(parts) >= 2 Then
                Dim cfRng5 As Range
                Set cfRng5 = ws.Range(parts(1))
                cfRng5.FormatConditions.AddUniqueValues
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).DupeUnique = xlUnique
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_TEXT"
            If UBound(parts) >= 3 Then
                Dim cfRng6 As Range
                Set cfRng6 = ws.Range(parts(1))
                cfRng6.FormatConditions.Add Type:=xlTextString, String:=parts(2), TextOperator:=xlContains
                cfRng6.FormatConditions(cfRng6.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng7 As Range
                Set cfRng7 = ws.Range(parts(1))
                cfRng7.FormatConditions.Add Type:=xlBlanksCondition
                cfRng7.FormatConditions(cfRng7.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_NOT_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng8 As Range
                Set cfRng8 = ws.Range(parts(1))
                cfRng8.FormatConditions.Add Type:=xlNoBlanksCondition
                cfRng8.FormatConditions(cfRng8.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "DATA_BARS"
            If UBound(parts) >= 2 Then
                Dim cfRng9 As Range
                Set cfRng9 = ws.Range(parts(1))
                cfRng9.FormatConditions.AddDatabar
                cfRng9.FormatConditions(cfRng9.FormatConditions.Count).BarColor.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COLOR_SCALE"
            If UBound(parts) >= 3 Then
                Dim cfRng10 As Range
                Set cfRng10 = ws.Range(parts(1))
                cfRng10.FormatConditions.AddColorScale ColorScaleType:=2
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COLOR_SCALE3"
            If UBound(parts) >= 4 Then
                Dim cfRng11 As Range
                Set cfRng11 = ws.Range(parts(1))
                cfRng11.FormatConditions.AddColorScale ColorScaleType:=3
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(3).FormatColor.Color = ParseColor(parts(4))
                ExecuteSingleCommand = True
            End If
            
        Case "ICON_SET"
            If UBound(parts) >= 2 Then
                Dim cfRng12 As Range
                Dim iconSetType As Long
                Set cfRng12 = ws.Range(parts(1))
                
                Select Case UCase(Trim(parts(2)))
                    Case "ARROWS": iconSetType = 1 ' xl3Arrows
                    Case "FLAGS": iconSetType = 7 ' xl3Flags
                    Case "STARS": iconSetType = 13 ' xl3Stars
                    Case "BARS": iconSetType = 14 ' xl4RedToBlack
                    Case Else: iconSetType = 1
                End Select
                
                cfRng12.FormatConditions.AddIconSetCondition
                cfRng12.FormatConditions(cfRng12.FormatConditions.Count).IconSet = ActiveWorkbook.IconSets(iconSetType)
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_COND_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).FormatConditions.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== DATA CHECK ==========
            
        Case "VALIDATION_LIST"
            If UBound(parts) >= 2 Then
                Dim valRng As Range
                Set valRng = ws.Range(parts(1))
                valRng.Validation.Delete
                valRng.Validation.Add Type:=xlValidateList, Formula1:=Replace(parts(2), ";", ",")
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_NUMBER"
            If UBound(parts) >= 3 Then
                Dim valRng2 As Range
                Set valRng2 = ws.Range(parts(1))
                valRng2.Validation.Delete
                valRng2.Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_DATE"
            If UBound(parts) >= 3 Then
                Dim valRng3 As Range
                Set valRng3 = ws.Range(parts(1))
                valRng3.Validation.Delete
                valRng3.Validation.Add Type:=xlValidateDate, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_TEXT_LENGTH"
            If UBound(parts) >= 3 Then
                Dim valRng4 As Range
                Set valRng4 = ws.Range(parts(1))
                valRng4.Validation.Delete
                valRng4.Validation.Add Type:=xlValidateTextLength, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_CUSTOM"
            If UBound(parts) >= 2 Then
                Dim valRng5 As Range
                Dim valFormula As String
                Set valRng5 = ws.Range(parts(1))
                valFormula = LocalizeFormula(parts(2))
                valRng5.Validation.Delete
                valRng5.Validation.Add Type:=xlValidateCustom, Formula1:=valFormula
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_VALIDATION"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Validation.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== COMMENTS ==========
            
        Case "ADD_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell As Range
                Set cmtCell = ws.Range(parts(1))
                If Not cmtCell.Comment Is Nothing Then cmtCell.Comment.Delete
                cmtCell.AddComment parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "EDIT_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell2 As Range
                Set cmtCell2 = ws.Range(parts(1))
                If Not cmtCell2.Comment Is Nothing Then
                    cmtCell2.Comment.text text:=parts(2)
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell3 As Range
                Set cmtCell3 = ws.Range(parts(1))
                If Not cmtCell3.Comment Is Nothing Then cmtCell3.Comment.Delete
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell4 As Range
                Set cmtCell4 = ws.Range(parts(1))
                If Not cmtCell4.Comment Is Nothing Then cmtCell4.Comment.Visible = True
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell5 As Range
                Set cmtCell5 = ws.Range(parts(1))
                If Not cmtCell5.Comment Is Nothing Then cmtCell5.Comment.Visible = False
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ALL_COMMENTS"
            Dim cmt As Comment
            For Each cmt In ws.Comments
                cmt.Visible = True
            Next cmt
            ExecuteSingleCommand = True
            
        Case "HIDE_ALL_COMMENTS"
            Dim cmt2 As Comment
            For Each cmt2 In ws.Comments
                cmt2.Visible = False
            Next cmt2
            ExecuteSingleCommand = True
        
        ' ========== HYPERLINKS ==========
            
        Case "ADD_HYPERLINK"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_HYPERLINK_CELL"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:="", SubAddress:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_HYPERLINK"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Hyperlinks.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== PROTECTION ==========
            
        Case "LOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = True
                ExecuteSingleCommand = True
            End If
            
        Case "UNLOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = False
                ExecuteSingleCommand = True
            End If
        
        ' ========== VIEWING AREA ==========
            
        Case "FREEZE_PANES"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ActiveWindow.FreezePanes = True
                ExecuteSingleCommand = True
            End If
            
        Case "FREEZE_TOP_ROW"
            ws.Range("A2").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True
            
        Case "FREEZE_FIRST_COLUMN"
            ws.Range("B1").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True
            
        Case "UNFREEZE_PANES"
            ActiveWindow.FreezePanes = False
            ExecuteSingleCommand = True
            
        Case "ZOOM"
            If UBound(parts) >= 1 Then
                ActiveWindow.Zoom = CLng(parts(1))
                ExecuteSingleCommand = True
            End If
            
        Case "GOTO"
            If UBound(parts) >= 1 Then
                Application.Goto Reference:=ws.Range(parts(1)), Scroll:=True
                ExecuteSingleCommand = True
            End If
            
        Case "SELECT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ExecuteSingleCommand = True
            End If
        
        ' ========== PRINT ==========
            
        Case "SET_PRINT_AREA"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintArea = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_PRINT_AREA"
            ws.PageSetup.PrintArea = ""
            ExecuteSingleCommand = True
            
        Case "PAGE_ORIENTATION"
            If UBound(parts) >= 1 Then
                If UCase(Trim(parts(1))) = "LANDSCAPE" Then
                    ws.PageSetup.Orientation = xlLandscape
                Else
                    ws.PageSetup.Orientation = xlPortrait
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PAGE_MARGINS"
            If UBound(parts) >= 4 Then
                With ws.PageSetup
                    .LeftMargin = Application.CentimetersToPoints(CDbl(parts(1)))
                    .RightMargin = Application.CentimetersToPoints(CDbl(parts(2)))
                    .TopMargin = Application.CentimetersToPoints(CDbl(parts(3)))
                    .BottomMargin = Application.CentimetersToPoints(CDbl(parts(4)))
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_TITLES_ROWS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleRows = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_TITLES_COLS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleColumns = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_GRIDLINES"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintGridlines = (UCase(Trim(parts(1))) = "TRUE")
                ExecuteSingleCommand = True
            End If
            
        Case "FIT_TO_PAGE"
            If UBound(parts) >= 2 Then
                With ws.PageSetup
                    .Zoom = False
                    .FitToPagesWide = CLng(parts(1))
                    .FitToPagesTall = CLng(parts(2))
                End With
                ExecuteSingleCommand = True
            End If
        
        ' ========== IMAGES ==========
            
        Case "INSERT_PICTURE"
            If UBound(parts) >= 5 Then
                Dim pic As Object
                Set pic = ws.Shapes.AddPicture(parts(1), msoFalse, msoTrue, _
                    CLng(parts(2)), CLng(parts(3)), CLng(parts(4)), CLng(parts(5)))
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_PICTURES"
            Dim shp As Shape
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then shp.Delete
            Next shp
            ExecuteSingleCommand = True
        
        ' ========== FORMS ==========
            
        Case "ADD_BUTTON"
            If UBound(parts) >= 5 Then
                Dim btn As Object
                Set btn = ws.Buttons.Add(CLng(parts(1)), CLng(parts(2)), CLng(parts(3)), CLng(parts(4)))
                btn.Caption = parts(5)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_CHECKBOX"
            If UBound(parts) >= 2 Then
                Dim chk As Object
                Dim chkCell As Range
                Set chkCell = ws.Range(parts(1))
                Set chk = ws.CheckBoxes.Add(chkCell.Left, chkCell.Top, 100, 15)
                chk.Caption = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_DROPDOWN"
            If UBound(parts) >= 2 Then
                Dim dd As Object
                Dim ddCell As Range
                Set ddCell = ws.Range(parts(1))
                Set dd = ws.DropDowns.Add(ddCell.Left, ddCell.Top, 100, 15)
                dd.List = Split(parts(2), ";")
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_SHAPES"
            Dim shp2 As Shape
            For Each shp2 In ws.Shapes
                shp2.Delete
            Next shp2
            ExecuteSingleCommand = True
        
        ' ========== SPECIAL ==========
            
        Case "CALCULATE"
            Application.Calculate
            ExecuteSingleCommand = True
            
        Case "CALCULATE_SHEET"
            ws.Calculate
            ExecuteSingleCommand = True
            
        Case "TEXT_TO_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim ttcRng As Range
                Dim delim As String
                Set ttcRng = ws.Range(parts(1))
                delim = parts(2)
                
                Dim delimTab As Boolean, delimSemi As Boolean, delimComma As Boolean, delimSpace As Boolean, delimOther As Boolean
                Dim otherChar As String
                
                Select Case UCase(delim)
                    Case "TAB": delimTab = True
                    Case "SEMICOLON", ";": delimSemi = True
                    Case "COMMA", ",": delimComma = True
                    Case "SPACE", " ": delimSpace = True
                    Case Else
                        delimOther = True
                        otherChar = delim
                End Select
                
                ttcRng.TextToColumns Destination:=ttcRng, DataType:=xlDelimited, _
                    Tab:=delimTab, Semicolon:=delimSemi, Comma:=delimComma, _
                    Space:=delimSpace, Other:=delimOther, otherChar:=otherChar
                    
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_SPACES"
            If UBound(parts) >= 1 Then
                Dim spRng As Range, spCell As Range
                Set spRng = ws.Range(parts(1))
                For Each spCell In spRng
                    If Not IsEmpty(spCell.value) Then
                        spCell.value = Application.WorksheetFunction.Trim(spCell.value)
                    End If
                Next spCell
                ExecuteSingleCommand = True
            End If
            
        Case "UPPER_CASE"
            If UBound(parts) >= 1 Then
                Dim ucRng As Range, ucCell As Range
                Set ucRng = ws.Range(parts(1))
                For Each ucCell In ucRng
                    If Not IsEmpty(ucCell.value) Then
                        ucCell.value = UCase(ucCell.value)
                    End If
                Next ucCell
                ExecuteSingleCommand = True
            End If
            
        Case "LOWER_CASE"
            If UBound(parts) >= 1 Then
                Dim lcRng As Range, lcCell As Range
                Set lcRng = ws.Range(parts(1))
                For Each lcCell In lcRng
                    If Not IsEmpty(lcCell.value) Then
                        lcCell.value = LCase(lcCell.value)
                    End If
                Next lcCell
                ExecuteSingleCommand = True
            End If
            
        Case "PROPER_CASE"
            If UBound(parts) >= 1 Then
                Dim pcRng As Range, pcCell As Range
                Set pcRng = ws.Range(parts(1))
                For Each pcCell In pcRng
                    If Not IsEmpty(pcCell.value) Then
                        pcCell.value = Application.WorksheetFunction.Proper(pcCell.value)
                    End If
                Next pcCell
                ExecuteSingleCommand = True
            End If
            
        Case "FLASH_FILL"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ws.Range(parts(1)).FlashFill
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If
            
        Case "SUBTOTAL"
            If UBound(parts) >= 3 Then
                Dim stRng As Range
                Dim stFunc As Long
                Dim stCol As Long
                Set stRng = ws.Range(parts(1))
                
                Select Case UCase(Trim(parts(2)))
                    Case "SUM": stFunc = xlSum
                    Case "COUNT": stFunc = xlCount
                    Case "AVERAGE": stFunc = xlAverage
                    Case "MAX": stFunc = xlMax
                    Case "MIN": stFunc = xlMin
                    Case Else: stFunc = xlSum
                End Select
                
                stCol = GetColumnNumber(parts(3), stRng, ws)
                stRng.Subtotal GroupBy:=1, Function:=stFunc, TotalList:=Array(stCol)
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_SUBTOTALS"
            ws.UsedRange.RemoveSubtotal
            ExecuteSingleCommand = True
            
        Case Else
            ExecuteSingleCommand = False
    End Select
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ExecuteSingleCommand Error: " & Err.Number & " - " & Err.Description & " | CMD: " & cmd
    ExecuteSingleCommand = False
End Function

'========================================
' LM STUDIO FUNCTIONS
'========================================

'----------------------------------------
' Get LM Studio setting from Registry
'----------------------------------------
Public Function GetLMStudioSetting(settingName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Dim result As String
    Set wsh = CreateObject("WScript.Shell")
    result = wsh.RegRead(REG_PATH & "LMStudio_" & settingName)
    Set wsh = Nothing
    
    ' Default values
    If Len(result) = 0 Then
        Select Case settingName
            Case "IP": result = LMSTUDIO_DEFAULT_IP
            Case "Port": result = LMSTUDIO_DEFAULT_PORT
            Case "Model": result = ""
            Case "Enabled": result = "0"
            Case "PreviewCommands": result = "0"
            Case "ResponseLanguage": result = "English"
        End Select
    End If
    
    GetLMStudioSetting = result
End Function

'----------------------------------------
' Save LM Studio setting to Registry
'----------------------------------------
Public Sub SaveLMStudioSetting(settingName As String, settingValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & "LMStudio_" & settingName, settingValue, "REG_SZ"
    Set wsh = Nothing
End Sub

'----------------------------------------
' Check LM Studio availability
'----------------------------------------
Public Function IsLMStudioAvailable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim ip As String
    Dim port As String
    
    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    url = "http://" & ip & ":" & port & "/v1/models"
    
    Set http = CreateHttpClient()
    If http Is Nothing Then
        IsLMStudioAvailable = False
        Exit Function
    End If
    
    http.setTimeouts 2000, 2000, 2000, 2000
    http.Open "GET", url, False
    http.send
    
    IsLMStudioAvailable = (http.Status = 200)
    Set http = Nothing
    Exit Function
    
ErrorHandler:
    IsLMStudioAvailable = False
End Function

'----------------------------------------
' Get list of models from LM Studio
'----------------------------------------
Public Function GetLMStudioModels() As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim ip As String
    Dim port As String
    Dim response As String
    Dim models As String
    Dim pos As Long
    Dim endPos As Long
    Dim modelId As String
    
    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    url = "http://" & ip & ":" & port & "/v1/models"
    
    Set http = CreateHttpClient()
    If http Is Nothing Then
        GetLMStudioModels = "ERROR:Could not create HTTP object."
        Exit Function
    End If
    
    http.setTimeouts 5000, 5000, 5000, 5000
    http.Open "GET", url, False
    http.send
    
    If http.Status <> 200 Then
        GetLMStudioModels = "ERROR:HTTP " & http.Status
        Set http = Nothing
        Exit Function
    End If
    
    response = http.responseText
    Set http = Nothing
    
    ' Parse JSON to extract model ids
    models = ""
    pos = 1
    Do
        pos = InStr(pos, response, """id""")
        If pos = 0 Then Exit Do
        
        pos = InStr(pos, response, ":")
        If pos = 0 Then Exit Do
        
        pos = InStr(pos, response, """")
        If pos = 0 Then Exit Do
        pos = pos + 1
        
        endPos = InStr(pos, response, """")
        If endPos = 0 Then Exit Do
        
        modelId = Mid(response, pos, endPos - pos)
        
        If Len(models) > 0 Then models = models & "|"
        models = models & modelId
        
        pos = endPos + 1
    Loop
    
    GetLMStudioModels = models
    Exit Function
    
ErrorHandler:
    GetLMStudioModels = "ERROR:" & Err.Description
End Function

'----------------------------------------
' Checking if the local model is enabled
'----------------------------------------
Public Function IsLocalModelEnabled() As Boolean
    IsLocalModelEnabled = (GetLMStudioSetting("Enabled") = "1")
End Function

'----------------------------------------
' Checking if the local model is configured
'----------------------------------------
Public Function HasLocalModel() As Boolean
    ' Local mode should be considered available only when LM Studio endpoint is reachable.
    HasLocalModel = IsLMStudioAvailable()
End Function

'----------------------------------------
' Send request to local LM Studio model
'----------------------------------------
Public Function SendToLocalAI(userMessage As String, Optional excelContext As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim ip As String
    Dim port As String
    Dim modelName As String
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim systemPrompt As String
    
    ip = GetLMStudioSetting("IP")
    port = GetLMStudioSetting("Port")
    modelName = GetLMStudioSetting("Model")
    
    If Len(ip) = 0 Or Len(port) = 0 Then
        SendToLocalAI = "ERROR: LM Studio settings are not configured. Open Settings."
        Exit Function
    End If
    
    url = "http://" & ip & ":" & port & "/v1/chat/completions"
    
    ' If model is not specified, use first available model
    If Len(modelName) = 0 Then
        Dim models As String
        models = GetLMStudioModels()
        If Left(models, 6) = "ERROR:" Then
            SendToLocalAI = "ERROR: Failed to get list of models: " & Mid(models, 7)
            Exit Function
        End If
        If Len(models) = 0 Then
            SendToLocalAI = "ERROR: There are no models loaded in LM Studio."
            Exit Function
        End If
        ' Use first model
        If InStr(models, "|") > 0 Then
            modelName = Left(models, InStr(models, "|") - 1)
        Else
            modelName = models
        End If
    End If
    
    ' Build system prompt
    systemPrompt = BuildSystemPrompt(excelContext)
    
    ' Build JSON request
    requestBody = BuildRequestJSON(systemPrompt, userMessage, modelName)
    
    ' Send request
    response = SendLocalHTTPRequest(url, requestBody)
    
    ' Parse response
    SendToLocalAI = ParseResponse(response)
    Exit Function
    
ErrorHandler:
    SendToLocalAI = "ERROR: " & Err.Description
End Function

'----------------------------------------
' HTTP request for local model (no API key required)
'----------------------------------------
Private Function SendLocalHTTPRequest(url As String, body As String) As String
    SendLocalHTTPRequest = SendHttpPostWithRetry(url, body, "lm-studio", "", True)
End Function

'----------------------------------------
' Check Codex CLI availability
'----------------------------------------
Public Function IsCodexCliAvailable() As Boolean
    On Error GoTo ErrorHandler
    
    Dim outText As String
    Dim errText As String
    Dim exitCode As Long
    
    If Not RunCommandCapture("codex --version", 10, outText, errText, exitCode) Then
        IsCodexCliAvailable = False
        Exit Function
    End If
    
    IsCodexCliAvailable = (exitCode = 0)
    Exit Function
    
ErrorHandler:
    IsCodexCliAvailable = False
End Function

'----------------------------------------
' Send request via local Codex CLI (ChatGPT Plus)
'----------------------------------------
Public Function SendToCodexCLI(userMessage As String, Optional excelContext As String = "", Optional imagePath As String = "") As String
    On Error GoTo ErrorHandler
    
    Dim workDir As String
    Dim tempRoot As String
    Dim tempDir As String
    Dim suffix As String
    Dim promptPath As String
    Dim outPath As String
    Dim systemPrompt As String
    Dim fullPrompt As String
    Dim cmd As String
    Dim outText As String
    Dim errText As String
    Dim exitCode As Long
    Dim runOk As Boolean
    
    If Not IsCodexCliAvailable() Then
        SendToCodexCLI = "ERROR: Codex CLI is not available. Install Codex CLI and ensure 'codex' is in PATH."
        Exit Function
    End If
    
    workDir = GetCodexWorkDir()
    
    tempRoot = Environ$("TEMP")
    If Len(tempRoot) = 0 Then tempRoot = CurDir$
    If Right$(tempRoot, 1) <> "\" Then tempRoot = tempRoot & "\"
    tempDir = tempRoot & CODEX_CLI_TEMP_SUBDIR
    
    If Not EnsureFolderExists(tempDir) Then
        SendToCodexCLI = "ERROR: Cannot create temp directory for Codex CLI: " & tempDir
        Exit Function
    End If
    
    Randomize
    suffix = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 9000) + 1000)
    promptPath = tempDir & "\prompt_" & suffix & ".txt"
    outPath = tempDir & "\response_" & suffix & ".txt"
    
    systemPrompt = BuildSystemPrompt(excelContext)
    fullPrompt = systemPrompt & vbCrLf & vbCrLf & "User task:" & vbCrLf & userMessage
    
    If Not WriteTextFileUtf8(promptPath, fullPrompt) Then
        SendToCodexCLI = "ERROR: Failed to prepare request for Codex CLI."
        GoTo Cleanup
    End If
    
    cmd = "type " & QuoteForCmd(promptPath) & " | codex exec --skip-git-repo-check --color never --output-last-message " & QuoteForCmd(outPath) & " -C " & QuoteForCmd(workDir)
    If Len(Trim$(imagePath)) > 0 Then
        cmd = cmd & " --image " & QuoteForCmd(imagePath)
    End If
    cmd = cmd & " -"
    
    runOk = RunCommandCapture(cmd, CODEX_CLI_TIMEOUT_SECONDS, outText, errText, exitCode)
    If Not runOk Then
        SendToCodexCLI = "ERROR: Codex CLI request timed out. Try a shorter request or less data."
        GoTo Cleanup
    End If
    
    If exitCode <> 0 Then
        If exitCode = -1073741510 Then
            SendToCodexCLI = "ERROR: Codex CLI process was interrupted (exit -1073741510). Do not close the console window during execution."
            GoTo Cleanup
        End If
        
        If InStr(LCase$(errText), "not logged in") > 0 Then
            SendToCodexCLI = "ERROR: Codex CLI is not logged in. Run 'codex login' in terminal and try again."
        Else
            SendToCodexCLI = "ERROR: Codex CLI failed (exit " & exitCode & "). " & CompactErrorText(errText)
        End If
        GoTo Cleanup
    End If
    
    If Len(Dir$(outPath)) = 0 Then
        SendToCodexCLI = "ERROR: Codex CLI returned no output."
        GoTo Cleanup
    End If
    
    SendToCodexCLI = Trim$(ReadTextFileUtf8(outPath))
    If Len(SendToCodexCLI) = 0 Then
        If Len(Trim$(outText)) > 0 Then
            SendToCodexCLI = Trim$(outText)
        Else
            SendToCodexCLI = "ERROR: Empty response from Codex CLI."
        End If
    End If
    
Cleanup:
    On Error Resume Next
    If Len(Dir$(promptPath)) > 0 Then Kill promptPath
    If Len(Dir$(outPath)) > 0 Then Kill outPath
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    SendToCodexCLI = "ERROR: " & Err.Description
End Function

Private Function GetCodexWorkDir() As String
    Dim p As String
    
    On Error Resume Next
    p = ""
    If Not ActiveWorkbook Is Nothing Then
        p = Trim$(ActiveWorkbook.Path)
    End If
    On Error GoTo 0
    
    If Len(p) = 0 Then p = Trim$(CurDir$)
    If Len(p) = 0 Then p = Trim$(Application.DefaultFilePath)
    If Len(p) = 0 Then p = Trim$(Environ$("USERPROFILE"))
    If Len(p) = 0 Then p = Trim$(ThisWorkbook.Path)
    If Len(p) = 0 Then p = "."
    
    GetCodexWorkDir = p
End Function

Private Function EnsureFolderExists(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
    EnsureFolderExists = True
    Exit Function
ErrorHandler:
    EnsureFolderExists = False
End Function

Private Function QuoteForCmd(value As String) As String
    QuoteForCmd = """" & Replace(value, """", """""") & """"
End Function

Private Function WriteTextFileUtf8(filePath As String, content As String) As Boolean
    On Error GoTo ErrorHandler
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2
    stm.Close
    Set stm = Nothing
    WriteTextFileUtf8 = True
    Exit Function
ErrorHandler:
    WriteTextFileUtf8 = WriteTextFileAnsi(filePath, content)
End Function

Private Function ReadTextFileUtf8(filePath As String) As String
    On Error GoTo ErrorHandler
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUtf8 = stm.ReadText(-1)
    stm.Close
    Set stm = Nothing
    Exit Function
ErrorHandler:
    ReadTextFileUtf8 = ReadTextFileAnsi(filePath)
End Function

Private Function RunCommandCapture(commandLine As String, timeoutSeconds As Long, ByRef stdoutText As String, ByRef stderrText As String, ByRef exitCode As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim wsh As Object
    Dim execObj As Object
    Dim started As Single
    Dim launchCmd As String
    Dim tempRoot As String
    Dim tempDir As String
    Dim suffix As String
    Dim outPath As String
    Dim errPath As String
    
    tempRoot = Environ$("TEMP")
    If Len(tempRoot) = 0 Then tempRoot = CurDir$
    If Right$(tempRoot, 1) <> "\" Then tempRoot = tempRoot & "\"
    tempDir = tempRoot & CODEX_CLI_TEMP_SUBDIR
    If Not EnsureFolderExists(tempDir) Then
        stdoutText = ""
        stderrText = "Cannot create temp directory for command execution."
        exitCode = -1
        RunCommandCapture = False
        Exit Function
    End If
    
    Randomize
    suffix = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 9000) + 1000)
    outPath = tempDir & "\exec_out_" & suffix & ".txt"
    errPath = tempDir & "\exec_err_" & suffix & ".txt"
    
    launchCmd = "cmd /d /c (" & commandLine & ") > " & QuoteForCmd(outPath) & " 2> " & QuoteForCmd(errPath)
    Set wsh = CreateObject("WScript.Shell")
    Set execObj = wsh.Exec(launchCmd)
    started = Timer
    
    Do While execObj.Status = 0
        DoEvents
        If GetElapsedSeconds(started) >= timeoutSeconds Then
            On Error Resume Next
            execObj.Terminate
            On Error GoTo 0
            stdoutText = ReadTextFileUtf8(outPath)
            stderrText = ReadTextFileUtf8(errPath)
            stderrText = stderrText & IIf(Len(Trim$(stderrText)) > 0, " ", "") & "Timeout after " & timeoutSeconds & " seconds."
            exitCode = -1
            RunCommandCapture = False
            GoTo Cleanup
        End If
        SpinWaitSeconds 0.2
    Loop
    
    stdoutText = ReadTextFileUtf8(outPath)
    stderrText = ReadTextFileUtf8(errPath)
    exitCode = CLng(execObj.ExitCode)
    RunCommandCapture = True
    GoTo Cleanup
    
Cleanup:
    On Error Resume Next
    If Len(outPath) > 0 Then
        If Len(Dir$(outPath)) > 0 Then Kill outPath
    End If
    If Len(errPath) > 0 Then
        If Len(Dir$(errPath)) > 0 Then Kill errPath
    End If
    Set execObj = Nothing
    Set wsh = Nothing
    On Error GoTo 0
    Exit Function
    
ErrorHandler:
    If InStr(1, Err.Description, "ActiveX component can't create object", vbTextCompare) > 0 Then
        Err.Clear
        RunCommandCapture = RunCommandCaptureShellFallback(commandLine, timeoutSeconds, stdoutText, stderrText, exitCode)
        Exit Function
    End If
    
    stdoutText = ""
    stderrText = Err.Description
    exitCode = -1
    RunCommandCapture = False
End Function

Private Function GetElapsedSeconds(started As Single) As Double
    Dim currentTime As Single
    currentTime = Timer
    If currentTime >= started Then
        GetElapsedSeconds = currentTime - started
    Else
        GetElapsedSeconds = (86400# - started) + currentTime
    End If
End Function

Private Sub SpinWaitSeconds(waitSeconds As Double)
    Dim started As Single
    started = Timer
    Do While GetElapsedSeconds(started) < waitSeconds
        DoEvents
    Loop
End Sub

Private Function CompactErrorText(rawText As String) As String
    Dim text As String
    text = Trim$(rawText)
    text = Replace(text, vbCr, " ")
    text = Replace(text, vbLf, " ")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    
    If Len(text) = 0 Then
        CompactErrorText = "No error details."
    ElseIf Len(text) > 220 Then
        CompactErrorText = Left$(text, 220) & "..."
    Else
        CompactErrorText = text
    End If
End Function

Private Function WriteTextFileAnsi(filePath As String, content As String) As Boolean
    On Error GoTo ErrorHandler
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, content;
    Close #ff
    WriteTextFileAnsi = True
    Exit Function
ErrorHandler:
    On Error Resume Next
    If ff > 0 Then Close #ff
    WriteTextFileAnsi = False
End Function

Private Function ReadTextFileAnsi(filePath As String) As String
    On Error GoTo ErrorHandler
    Dim ff As Integer
    Dim lineText As String
    Dim result As String
    
    ff = FreeFile
    Open filePath For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, lineText
        If Len(result) > 0 Then result = result & vbCrLf
        result = result & lineText
    Loop
    Close #ff
    
    ReadTextFileAnsi = result
    Exit Function
ErrorHandler:
    On Error Resume Next
    If ff > 0 Then Close #ff
    ReadTextFileAnsi = ""
End Function

Private Function RunCommandCaptureShellFallback(commandLine As String, timeoutSeconds As Long, ByRef stdoutText As String, ByRef stderrText As String, ByRef exitCode As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tempRoot As String
    Dim tempDir As String
    Dim suffix As String
    Dim outPath As String
    Dim errPath As String
    Dim codePath As String
    Dim wrapped As String
    Dim started As Single
    Dim codeText As String
    
    tempRoot = Environ$("TEMP")
    If Len(tempRoot) = 0 Then tempRoot = CurDir$
    If Right$(tempRoot, 1) <> "\" Then tempRoot = tempRoot & "\"
    tempDir = tempRoot & CODEX_CLI_TEMP_SUBDIR
    If Not EnsureFolderExists(tempDir) Then
        stderrText = "Cannot create temp directory for shell fallback."
        stdoutText = ""
        exitCode = -1
        RunCommandCaptureShellFallback = False
        Exit Function
    End If
    
    Randomize
    suffix = Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Rnd() * 9000) + 1000)
    outPath = tempDir & "\fallback_out_" & suffix & ".txt"
    errPath = tempDir & "\fallback_err_" & suffix & ".txt"
    codePath = tempDir & "\fallback_code_" & suffix & ".txt"
    
    wrapped = "cmd /d /c (" & commandLine & ") > " & QuoteForCmd(outPath) & " 2> " & QuoteForCmd(errPath) & " & echo %errorlevel% > " & QuoteForCmd(codePath)
    Shell wrapped, vbHide
    
    started = Timer
    Do While Len(Dir$(codePath)) = 0
        DoEvents
        If GetElapsedSeconds(started) >= timeoutSeconds Then
            stdoutText = ""
            stderrText = "Timeout after " & timeoutSeconds & " seconds."
            exitCode = -1
            RunCommandCaptureShellFallback = False
            GoTo Cleanup
        End If
        SpinWaitSeconds 0.2
    Loop
    
    stdoutText = ReadTextFileUtf8(outPath)
    stderrText = ReadTextFileUtf8(errPath)
    codeText = Trim$(ReadTextFileAnsi(codePath))
    If IsNumeric(codeText) Then
        exitCode = CLng(codeText)
    Else
        exitCode = -1
    End If
    
    RunCommandCaptureShellFallback = True
    GoTo Cleanup
    
ErrorHandler:
    stdoutText = ""
    stderrText = Err.Description
    exitCode = -1
    RunCommandCaptureShellFallback = False
    
Cleanup:
    On Error Resume Next
    If Len(outPath) > 0 Then
        If Len(Dir$(outPath)) > 0 Then Kill outPath
    End If
    If Len(errPath) > 0 Then
        If Len(Dir$(errPath)) > 0 Then Kill errPath
    End If
    If Len(codePath) > 0 Then
        If Len(Dir$(codePath)) > 0 Then Kill codePath
    End If
    On Error GoTo 0
End Function
