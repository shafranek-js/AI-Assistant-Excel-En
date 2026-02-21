Attribute VB_Name = "frmChat"
Attribute VB_Base = "0{5DC78F22-B3FF-4F43-A3CB-922D56074F42}{90BEFFAC-02B5-4FA8-A512-F43F7344A1E9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private chatHistory As String
Private attachedImagePath As String
Private planCommands() As String
Private planStatuses() As String
Private planCount As Long
Private lastUserMessage As String
Private lastModel As String
Private lastContext As String
Private lastUseLocal As Boolean
Private lastImagePath As String
Private hasRetryPayload As Boolean

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "AI Assistant for Excel"
    optCloud.Caption = "Cloud"
    optLocal.Caption = "Local"
    chkIncludeData.Caption = "Include Data"
    chkPreviewCommands.Caption = "Preview Commands"
    lblSafetyMode.Caption = "Safety Mode"
    lblActionPlan.Caption = "Action Plan"
    lblSystemLog.Caption = "System Log"
    btnClearPlan.Caption = "Clear Plan"
    btnRetry.Caption = "Retry"
    btnSend.Caption = "Send"
    btnClear.Caption = "Clear"
    btnAttach.Caption = "Attach"
    btnSettings.Caption = "Settings"
    btnClose.Caption = "Close"
    On Error GoTo 0

    attachedImagePath = ""
    lblAttachment.Caption = ""
    cmbModel.Clear
    cmbModel.AddItem "Gemini 3 Flash"
    cmbModel.AddItem "GPT-5.2 (OpenRouter)"
    cmbModel.AddItem "GPT-5.2 (Direct OpenAI)"
    cmbModel.AddItem "GPT-5.2 Codex (Direct)"
    cmbModel.AddItem "Codex CLI (ChatGPT Plus)"
    cmbModel.AddItem "Gemini CLI (Google)"
    cmbModel.AddItem "Gemini 3 Pro"
    cmbModel.AddItem "Claude Sonnet 4.5"
    cmbModel.AddItem "DeepSeek"
    cmbModel.ListWidth = 280

    cmbSafetyMode.Clear
    cmbSafetyMode.AddItem "Preview"
    cmbSafetyMode.AddItem "Auto-apply"
    cmbSafetyMode.AddItem "Read-only"

    If HasApiKey("openrouter") Then
        cmbModel.Value = "Gemini 3 Flash"
    ElseIf HasApiKey("openai") Then
        cmbModel.Value = "GPT-5.2 (Direct OpenAI)"
    ElseIf HasApiKey("deepseek") Then
        cmbModel.Value = "DeepSeek"
    ElseIf IsCodexCliAvailable() Then
        cmbModel.Value = "Codex CLI (ChatGPT Plus)"
    ElseIf IsGeminiCliAvailable() Then
        cmbModel.Value = "Gemini CLI (Google)"
    Else
        cmbModel.ListIndex = 0
    End If

    chkIncludeData.Value = True
    chkPreviewCommands.Value = (GetLMStudioSetting("PreviewCommands") = "1")
    cmbSafetyMode.Value = NormalizeSafetyMode(GetLMStudioSetting("SafetyMode"))
    If Len(cmbSafetyMode.Value) = 0 Then cmbSafetyMode.Value = "Preview"
    UpdateSafetyModeUI

    chatHistory = "AI: Hello! I will help with data analysis, formulas and formatting." & vbCrLf & _
                  "Select the data and describe the task." & vbCrLf & _
                  "For feature requests, contact: t.me/jsx_xsj" & vbCrLf & vbCrLf
    txtChat.Value = chatHistory
    lblStatus.Caption = "Ready"

    ResetActionPlan
    txtSystemLog.Value = ""
    txtSystemLog.Locked = True
    btnRetry.Enabled = False
    hasRetryPayload = False

    If IsLocalModelEnabled() And HasLocalModel() Then
        optLocal.Value = True
    Else
        optCloud.Value = True
    End If
    UpdateModelMode
End Sub

Private Function NormalizeSafetyMode(ByVal modeText As String) As String
    Select Case LCase$(Trim$(modeText))
        Case "preview"
            NormalizeSafetyMode = "Preview"
        Case "auto-apply", "auto apply", "auto"
            NormalizeSafetyMode = "Auto-apply"
        Case "read-only", "read only", "readonly"
            NormalizeSafetyMode = "Read-only"
        Case Else
            NormalizeSafetyMode = "Preview"
    End Select
End Function

Private Sub UpdateSafetyModeUI()
    Dim safetyMode As String
    safetyMode = NormalizeSafetyMode(cmbSafetyMode.Value)

    Select Case safetyMode
        Case "Auto-apply"
            chkPreviewCommands.Value = False
            chkPreviewCommands.Enabled = False
            chkPreviewCommands.Caption = "Preview Commands (off)"
        Case "Read-only"
            chkPreviewCommands.Value = False
            chkPreviewCommands.Enabled = False
            chkPreviewCommands.Caption = "Preview Commands (N/A)"
        Case Else
            chkPreviewCommands.Enabled = True
            chkPreviewCommands.Caption = "Preview Commands"
            chkPreviewCommands.Value = True
    End Select

    SaveLMStudioSetting "PreviewCommands", IIf(chkPreviewCommands.Value, "1", "0")
End Sub

Private Sub ResetActionPlan()
    planCount = 0
    Erase planCommands
    Erase planStatuses
    RenderActionPlan
    lblActionSummary.Caption = "No actions yet."
End Sub

Private Sub SetActionSummary(ByVal summaryText As String)
    lblActionSummary.Caption = summaryText
End Sub

Private Sub RenderActionPlan()
    Dim i As Long
    lstActionPlan.Clear
    For i = 1 To planCount
        lstActionPlan.AddItem "[" & planStatuses(i) & "] " & planCommands(i)
    Next i
End Sub

Private Sub SetActionPlanFromCommands(ByVal commands As String, ByVal defaultStatus As String)
    Dim lines() As String
    Dim i As Long
    Dim cmd As String

    ResetActionPlan
    If Len(commands) = 0 Then Exit Sub

    lines = Split(commands, vbLf)
    planCount = 0
    For i = 0 To UBound(lines)
        cmd = Trim$(Replace(lines(i), vbCr, ""))
        If Len(cmd) > 0 Then
            planCount = planCount + 1
            ReDim Preserve planCommands(1 To planCount)
            ReDim Preserve planStatuses(1 To planCount)
            planCommands(planCount) = cmd
            planStatuses(planCount) = defaultStatus
        End If
    Next i

    RenderActionPlan
End Sub

Private Sub SetAllActionStatuses(ByVal statusText As String)
    Dim i As Long
    For i = 1 To planCount
        planStatuses(i) = statusText
    Next i
    RenderActionPlan
End Sub

Private Sub MarkFirstActionByCommand(ByVal commandText As String, ByVal statusText As String)
    Dim i As Long
    For i = 1 To planCount
        If planCommands(i) = commandText Then
            If planStatuses(i) = "Executed" Or planStatuses(i) = "Planned" Then
                planStatuses(i) = statusText
                Exit For
            End If
        End If
    Next i
End Sub

Private Function FirstLine(ByVal textValue As String) As String
    Dim p As Long
    p = InStr(textValue, vbCrLf)
    If p > 0 Then
        FirstLine = Left$(textValue, p - 1)
    Else
        FirstLine = textValue
    End If
End Function

Private Sub ApplyExecutionDetails(ByVal execResult As String)
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim cmdText As String
    Dim posClose As Long

    lines = Split(execResult, vbCrLf)
    For i = 0 To UBound(lines)
        lineText = Trim$(lines(i))
        If Left$(lineText, 19) = "- Runtime failure: " Then
            cmdText = Mid$(lineText, 20)
            MarkFirstActionByCommand cmdText, "Failed"
        ElseIf Left$(lineText, 11) = "- Rejected (" Then
            posClose = InStr(lineText, "): ")
            If posClose > 0 Then
                cmdText = Mid$(lineText, posClose + 3)
                MarkFirstActionByCommand cmdText, "Rejected"
            End If
        End If
    Next i

    RenderActionPlan
End Sub

Private Sub AppendSystemLog(ByVal messageText As String)
    If Len(Trim$(messageText)) = 0 Then Exit Sub
    If Len(txtSystemLog.Value) > 0 Then
        txtSystemLog.Value = txtSystemLog.Value & vbCrLf
    End If
    txtSystemLog.Value = txtSystemLog.Value & Format$(Now, "hh:nn:ss") & " | " & messageText
    txtSystemLog.SelStart = Len(txtSystemLog.Value)
End Sub

Private Function IsTransientError(ByVal responseText As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(responseText))

    If Len(t) = 0 Then Exit Function
    If InStr(t, "timed out") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "timeout") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "rate limit") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "request limit") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "server error") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "temporar") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "network") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "connection") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "http 429") > 0 Then IsTransientError = True: Exit Function
    If InStr(t, "http 5") > 0 Then IsTransientError = True: Exit Function
End Function

Private Function IsSystemErrorResponse(ByVal responseText As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(responseText))
    IsSystemErrorResponse = (Left$(t, 6) = "error:" Or Left$(t, 10) = "api error:")
End Function

Private Function BuildPreflightNotes(ByVal useLocal As Boolean, ByVal model As String, ByVal imagePath As String) As String
    Dim notes As String
    Dim hasSelectionRange As Boolean

    hasSelectionRange = (TypeName(Selection) = "Range")

    If useLocal Then
        If Not HasLocalModel() Then
            notes = notes & "- LM Studio endpoint is not reachable." & vbCrLf
        End If
    Else
        If model = "codex-cli" And Not IsCodexCliAvailable() Then
            notes = notes & "- Codex CLI is not available in PATH." & vbCrLf
        ElseIf model = "gemini-cli" And Not IsGeminiCliAvailable() Then
            notes = notes & "- Gemini CLI is not available in PATH." & vbCrLf
        End If
    End If

    If chkIncludeData.Value And Not hasSelectionRange Then
        notes = notes & "- Include Data is enabled, but no range is selected." & vbCrLf
    End If

    If Not chkIncludeData.Value Then
        notes = notes & "- Include Data is disabled: only workbook metadata will be sent." & vbCrLf
    End If

    If Len(imagePath) > 0 Then
        notes = notes & "- Attached image: " & GetFileName(imagePath) & vbCrLf
    End If

    BuildPreflightNotes = notes
End Function

Private Sub UpdateAttachmentChip()
    If Len(attachedImagePath) = 0 Then
        lblAttachment.Caption = ""
    Else
        lblAttachment.Caption = "[Image] " & GetFileName(attachedImagePath) & " (double-click to remove)"
    End If
End Sub

Private Sub optCloud_Click()
    UpdateModelMode
    SaveLMStudioSetting "Enabled", "0"
End Sub

Private Sub optLocal_Click()
    UpdateModelMode
    SaveLMStudioSetting "Enabled", "1"
End Sub

Private Sub UpdateModelMode()
    If optLocal.Value Then
        cmbModel.Visible = False
        lblLocalModel.Visible = True
        Dim modelName As String
        modelName = GetLMStudioSetting("Model")
        If Len(modelName) = 0 Then
            lblLocalModel.Caption = "LM Studio (auto)"
        ElseIf Len(modelName) > 25 Then
            lblLocalModel.Caption = Left(modelName, 22) & "..."
        Else
            lblLocalModel.Caption = modelName
        End If
    Else
        cmbModel.Visible = True
        lblLocalModel.Visible = False
    End If
End Sub

Private Sub ExecuteAssistantRequest(ByVal userMessage As String, ByVal useLocal As Boolean, ByVal model As String, ByVal context As String, ByVal imagePath As String, ByVal isRetry As Boolean)
    Dim aiResponse As String
    Dim commands As String
    Dim execResult As String
    Dim shouldExecuteCommands As Boolean
    Dim safetyMode As String
    Dim preflightNotes As String
    Dim chatAiText As String
    Dim previewText As String

    If isRetry Then
        chatHistory = chatHistory & "You (retry): " & userMessage & vbCrLf & vbCrLf
    Else
        chatHistory = chatHistory & "You: " & userMessage & vbCrLf & vbCrLf
    End If
    txtChat.Value = chatHistory
    txtInput.Value = ""

    preflightNotes = BuildPreflightNotes(useLocal, model, imagePath)
    If Len(preflightNotes) > 0 Then
        If MsgBox("Preflight check:" & vbCrLf & vbCrLf & preflightNotes & vbCrLf & "Continue?", vbQuestion + vbYesNo, "Preflight Check") <> vbYes Then
            Exit Sub
        End If
    End If

    If useLocal Then
        lblStatus.Caption = "LM Studio..."
    ElseIf model = "codex-cli" Then
        lblStatus.Caption = "Codex CLI..."
    ElseIf model = "gemini-cli" Then
        lblStatus.Caption = "Gemini CLI..."
    Else
        lblStatus.Caption = "Sending..."
    End If
    Me.Repaint

    If useLocal Then
        aiResponse = SendToLocalAI(userMessage, context)
    Else
        aiResponse = SendToAI(userMessage, model, context, imagePath)
    End If

    If IsTransientError(aiResponse) Then
        btnRetry.Enabled = True
        AppendSystemLog "Transient error detected. Retry is available."
    Else
        btnRetry.Enabled = False
    End If

    attachedImagePath = ""
    UpdateAttachmentChip

    commands = ExtractCommands(aiResponse)
    If Len(commands) > 0 Then
        SetActionPlanFromCommands commands, "Planned"
        safetyMode = NormalizeSafetyMode(cmbSafetyMode.Value)

        Select Case safetyMode
            Case "Read-only"
                SetAllActionStatuses "Skipped"
                SetActionSummary "Read-only: commands were prepared but not executed."
                AppendSystemLog "Read-only mode skipped command execution."
                chatAiText = aiResponse & vbCrLf & vbCrLf & "[Read-only mode: commands were not executed]"
            Case "Auto-apply"
                lblStatus.Caption = "Executing..."
                Me.Repaint
                execResult = ExecuteCommands(commands)
                SetAllActionStatuses "Executed"
                ApplyExecutionDetails execResult
                SetActionSummary FirstLine(execResult)
                AppendSystemLog FirstLine(execResult)
                chatAiText = aiResponse & vbCrLf & vbCrLf & "[" & execResult & "]"
            Case Else
                shouldExecuteCommands = True
                If chkPreviewCommands.Value Then
                    previewText = "The assistant prepared these commands:" & vbCrLf & vbCrLf & commands & vbCrLf & vbCrLf & "Execute now?"
                    shouldExecuteCommands = (MsgBox(previewText, vbQuestion + vbYesNo, "Preview Commands") = vbYes)
                End If

                If shouldExecuteCommands Then
                    lblStatus.Caption = "Executing..."
                    Me.Repaint
                    execResult = ExecuteCommands(commands)
                    SetAllActionStatuses "Executed"
                    ApplyExecutionDetails execResult
                    SetActionSummary FirstLine(execResult)
                    AppendSystemLog FirstLine(execResult)
                    chatAiText = aiResponse & vbCrLf & vbCrLf & "[" & execResult & "]"
                Else
                    SetAllActionStatuses "Canceled"
                    SetActionSummary "Execution canceled by user."
                    AppendSystemLog "Execution canceled by user."
                    chatAiText = aiResponse & vbCrLf & vbCrLf & "[Command execution canceled by user]"
                End If
        End Select
    Else
        SetActionSummary "No commands block in response."
        chatAiText = aiResponse
    End If

    If IsSystemErrorResponse(aiResponse) Then
        AppendSystemLog aiResponse
        chatAiText = "[System message logged. See System Log panel.]"
    End If

    chatHistory = chatHistory & "AI: " & chatAiText & vbCrLf & vbCrLf
    txtChat.Value = chatHistory
    txtChat.SelStart = Len(txtChat.Value)
    lblStatus.Caption = "Ready"
End Sub

Private Sub btnSend_Click()
    Dim userMessage As String
    Dim context As String
    Dim model As String
    Dim useLocal As Boolean
    Dim keyType As String

    userMessage = Trim(txtInput.Value)
    If Len(userMessage) = 0 Then Exit Sub

    ResetActionPlan

    useLocal = optLocal.Value

    If useLocal Then
        If Not HasLocalModel() Then
            MsgBox "LM Studio is not configured. Open Settings.", vbExclamation
            Exit Sub
        End If
    Else
        Select Case cmbModel.Value
            Case "DeepSeek": model = "deepseek"
            Case "Claude Sonnet 4.5": model = "claude"
            Case "GPT-5.2 (OpenRouter)": model = "gpt"
            Case "GPT-5.2 (Direct OpenAI)": model = "gpt-direct"
            Case "GPT-5.2 Codex (Direct)": model = "gpt-codex-direct"
            Case "Codex CLI (ChatGPT Plus)": model = "codex-cli"
            Case "Gemini CLI (Google)": model = "gemini-cli"
            Case "Gemini 3 Pro": model = "gemini"
            Case "Gemini 3 Flash": model = "gemini-flash"
            Case Else: model = "deepseek"
        End Select

        If model = "deepseek" Then
            keyType = "deepseek"
        ElseIf model = "codex-cli" Or model = "gemini-cli" Then
            keyType = ""
        ElseIf model = "gpt-direct" Or model = "gpt-codex-direct" Then
            keyType = "openai"
        Else
            keyType = "openrouter"
        End If

        If model <> "codex-cli" And model <> "gemini-cli" And Not HasApiKey(keyType) Then
            MsgBox "API key is not configured. Open Settings.", vbExclamation
            Exit Sub
        End If
    End If

    context = GetWorkbookContext()
    If chkIncludeData.Value Then
        context = context & vbCrLf & GetSelectedData()
    End If

    lastUserMessage = userMessage
    lastModel = model
    lastContext = context
    lastUseLocal = useLocal
    lastImagePath = attachedImagePath
    hasRetryPayload = True

    ExecuteAssistantRequest userMessage, useLocal, model, context, attachedImagePath, False
End Sub

Private Sub chkPreviewCommands_Click()
    SaveLMStudioSetting "PreviewCommands", IIf(chkPreviewCommands.Value, "1", "0")
End Sub

Private Sub cmbSafetyMode_Change()
    cmbSafetyMode.Value = NormalizeSafetyMode(cmbSafetyMode.Value)
    SaveLMStudioSetting "SafetyMode", cmbSafetyMode.Value
    UpdateSafetyModeUI
End Sub

Private Sub btnClear_Click()
    chatHistory = "AI: Chat history cleared." & vbCrLf & vbCrLf
    txtChat.Value = chatHistory
    attachedImagePath = ""
    UpdateAttachmentChip
End Sub

Private Sub btnClearPlan_Click()
    ResetActionPlan
End Sub

Private Sub btnRetry_Click()
    If Not hasRetryPayload Then Exit Sub
    ExecuteAssistantRequest lastUserMessage, lastUseLocal, lastModel, lastContext, lastImagePath, True
End Sub

Private Sub btnAttach_Click()
    Dim fd As Object
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select an image"
        .Filters.Clear
        .Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.gif;*.webp"
        .AllowMultiSelect = False

        If .Show = -1 Then
            attachedImagePath = .SelectedItems(1)
            UpdateAttachmentChip
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub lblAttachment_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(attachedImagePath) > 0 Then
        attachedImagePath = ""
        UpdateAttachmentChip
    End If
End Sub

Private Function GetFileName(fullPath As String) As String
    Dim parts() As String
    parts = Split(fullPath, "\")
    GetFileName = parts(UBound(parts))
End Function

Private Sub btnSettings_Click()
    frmSettings.Show vbModal
    UpdateModelMode
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub txtInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        btnSend_Click
        KeyCode = 0
    End If
End Sub
