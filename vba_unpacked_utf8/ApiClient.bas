Attribute VB_Name = "ApiClient"
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
Private Const CODEX_CLI_TIMEOUT_SECONDS As Long = 90
Private Const CODEX_CLI_TEMP_SUBDIR As String = "ExcelAIAssistantCodex"

' API key storage (Registry)
Private Const REG_PATH As String = "HKEY_CURRENT_USER\Software\ExcelAIAssistant\"

Public Function GetApiKey(keyName As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetApiKey = wsh.RegRead(REG_PATH & keyName)
    Set wsh = Nothing
End Function

Public Sub SaveApiKey(keyName As String, keyValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & keyName, keyValue, "REG_SZ"
    Set wsh = Nothing
End Sub

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

Public Sub SaveLMStudioSetting(settingName As String, settingValue As String)
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite REG_PATH & "LMStudio_" & settingName, settingValue, "REG_SZ"
    Set wsh = Nothing
End Sub

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

Public Function IsLocalModelEnabled() As Boolean
    IsLocalModelEnabled = (GetLMStudioSetting("Enabled") = "1")
End Function

Public Function HasLocalModel() As Boolean
    ' Local mode should be considered available only when LM Studio endpoint is reachable.
    HasLocalModel = IsLMStudioAvailable()
End Function

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

Private Function SendLocalHTTPRequest(url As String, body As String) As String
    SendLocalHTTPRequest = SendHttpPostWithRetry(url, body, "lm-studio", "", True)
End Function

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


