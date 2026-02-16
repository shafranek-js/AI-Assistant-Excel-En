Attribute VB_Name = "PromptBuilder"
Option Explicit

Private Const MAX_TOKENS_DEFAULT As Long = 4096
Private Const MAX_TOKENS_OPENAI_DIRECT As Long = 1400
Private Const MAX_CONTEXT_CHARS_DEFAULT As Long = 24000
Private Const MAX_CONTEXT_CHARS_OPENAI_DIRECT As Long = 12000

Public Function ImageToBase64(filePath As String) As String
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

Public Function BuildRequestJSONWithImage(systemPrompt As String, userMessage As String, modelName As String, imageBase64 As String, imagePath As String, Optional maxTokens As Long = MAX_TOKENS_DEFAULT) As String
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

Public Function BuildSystemPrompt(excelContext As String) As String
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
    prompt = prompt & "CREATE_TABLE|range|name - convert range to Smart Table (name is optional)" & vbCrLf
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

Public Function GetMaxTokensForModel(model As String) As Long
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

Public Function ClampContextForModel(excelContext As String, model As String) As String
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

Public Function BuildRequestJSON(systemPrompt As String, userMessage As String, modelName As String, Optional maxTokens As Long = MAX_TOKENS_DEFAULT) As String
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

Public Function EscapeJSON(text As String) As String
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



