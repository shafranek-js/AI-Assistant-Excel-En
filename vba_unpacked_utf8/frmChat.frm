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

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "AI Assistant for Excel"
    optCloud.Caption = "Cloud"
    optLocal.Caption = "Local"
    chkIncludeData.Caption = "Include Data"
    chkPreviewCommands.Caption = "Preview Commands"
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
    cmbModel.AddItem "Gemini 3 Pro"
    cmbModel.AddItem "Claude Sonnet 4.5"
    cmbModel.AddItem "DeepSeek"
    cmbModel.ListWidth = 280
    
    If HasApiKey("openrouter") Then
        cmbModel.value = "Gemini 3 Flash"
    ElseIf HasApiKey("openai") Then
        cmbModel.value = "GPT-5.2 (Direct OpenAI)"
    ElseIf HasApiKey("deepseek") Then
        cmbModel.value = "DeepSeek"
    ElseIf IsCodexCliAvailable() Then
        cmbModel.value = "Codex CLI (ChatGPT Plus)"
    Else
        cmbModel.ListIndex = 0
    End If
    
    chkIncludeData.value = True
    chkPreviewCommands.value = (GetLMStudioSetting("PreviewCommands") = "1")
    chatHistory = "AI: Hello! I will help with data analysis, formulas and formatting." & vbCrLf & "Select the data and describe the task." & vbCrLf & "For feature requests, contact: t.me/jsx_xsj" & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    lblStatus.Caption = "Ready"
    
    ' Set model mode
    If IsLocalModelEnabled() And HasLocalModel() Then
        optLocal.value = True
    Else
        optCloud.value = True
    End If
    UpdateModelMode
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
    If optLocal.value Then
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

Private Sub btnSend_Click()
    Dim userMessage As String
    Dim aiResponse As String
    Dim context As String
    Dim model As String
    Dim useLocal As Boolean
    
    userMessage = Trim(txtInput.value)
    If Len(userMessage) = 0 Then Exit Sub
    
    useLocal = optLocal.value
    
    If useLocal Then
        If Not HasLocalModel() Then
            MsgBox "LM Studio is not configured. Open Settings.", vbExclamation
            Exit Sub
        End If
    Else
        Select Case cmbModel.value
            Case "DeepSeek"
                model = "deepseek"
            Case "Claude Sonnet 4.5"
                model = "claude"
            Case "GPT-5.2 (OpenRouter)"
                model = "gpt"
            Case "GPT-5.2 (Direct OpenAI)"
                model = "gpt-direct"
            Case "GPT-5.2 Codex (Direct)"
                model = "gpt-codex-direct"
            Case "Codex CLI (ChatGPT Plus)"
                model = "codex-cli"
            Case "Gemini 3 Pro"
                model = "gemini"
            Case "Gemini 3 Flash"
                model = "gemini-flash"
            Case Else
                model = "deepseek"
        End Select
        
        Dim keyType As String
        If model = "deepseek" Then
            keyType = "deepseek"
        ElseIf model = "codex-cli" Then
            keyType = ""
        ElseIf model = "gpt-direct" Or model = "gpt-codex-direct" Then
            keyType = "openai"
        Else
            keyType = "openrouter"
        End If
        
        If model <> "codex-cli" And Not HasApiKey(keyType) Then
            MsgBox "API key is not configured. Open Settings.", vbExclamation
            Exit Sub
        End If
    End If
    
    chatHistory = chatHistory & "You: " & userMessage & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    txtInput.value = ""
    
    context = GetWorkbookContext()
    If chkIncludeData.value Then
        context = context & vbCrLf & GetSelectedData()
    End If
    
    If useLocal Then
        lblStatus.Caption = "LM Studio..."
    ElseIf model = "codex-cli" Then
        lblStatus.Caption = "Codex CLI..."
    Else
        lblStatus.Caption = "Sending..."
    End If
    Me.Repaint
    
    If useLocal Then
        aiResponse = SendToLocalAI(userMessage, context)
    Else
        aiResponse = SendToAI(userMessage, model, context, attachedImagePath)
    End If
    
    attachedImagePath = ""
    lblAttachment.Caption = ""
    
    Dim commands As String
    Dim execResult As String
    Dim shouldExecuteCommands As Boolean
    commands = ExtractCommands(aiResponse)
    
    If Len(commands) > 0 Then
        shouldExecuteCommands = True

        If chkPreviewCommands.value Then
            Dim previewText As String
            previewText = "The assistant prepared these commands:" & vbCrLf & vbCrLf & commands & vbCrLf & vbCrLf & "Execute now?"
            shouldExecuteCommands = (MsgBox(previewText, vbQuestion + vbYesNo, "Preview Commands") = vbYes)
        End If

        If shouldExecuteCommands Then
            lblStatus.Caption = "Executing..."
            Me.Repaint
            execResult = ExecuteCommands(commands)
            aiResponse = aiResponse & vbCrLf & vbCrLf & "[" & execResult & "]"
        Else
            aiResponse = aiResponse & vbCrLf & vbCrLf & "[Command execution canceled by user]"
        End If
    End If
    
    chatHistory = chatHistory & "AI: " & aiResponse & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    txtChat.SelStart = Len(txtChat.value)
    lblStatus.Caption = "Ready"
End Sub

Private Sub chkPreviewCommands_Click()
    SaveLMStudioSetting "PreviewCommands", IIf(chkPreviewCommands.value, "1", "0")
End Sub

Private Sub btnClear_Click()
    chatHistory = "AI: Chat history cleared." & vbCrLf & vbCrLf
    txtChat.value = chatHistory
    attachedImagePath = ""
    lblAttachment.Caption = ""
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
            lblAttachment.Caption = "Attached: " & GetFileName(attachedImagePath)
        End If
    End With
    Set fd = Nothing
End Sub

Private Sub lblAttachment_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(attachedImagePath) > 0 Then
        attachedImagePath = ""
        lblAttachment.Caption = ""
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
