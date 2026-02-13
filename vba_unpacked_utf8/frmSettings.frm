Attribute VB_Name = "frmSettings"
Attribute VB_Base = "0{4F1B3E88-ABBD-4DA5-AB90-63A02455D148}{AAB48C23-C249-4258-B910-A352D18C8EB9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = "AI Assistant Settings"
    Me.Width = 500
    Me.Height = 340
    btnRefreshModels.Caption = "Refresh Models"
    btnSave.Caption = "Save"
    btnCancel.Caption = "Cancel"
    lblOpenAI.Caption = "OpenAI API key"
    lblResponseLanguage.Caption = "Response language"
    On Error GoTo 0

    txtDeepSeekKey.value = GetApiKey("DeepSeekKey")
    txtOpenRouterKey.value = GetApiKey("OpenRouterKey")
    txtOpenAIKey.value = GetApiKey("OpenAIKey")
    
    txtLMStudioIP.value = GetLMStudioSetting("IP")
    txtLMStudioPort.value = GetLMStudioSetting("Port")
    
    LoadLMStudioModels
    LoadResponseLanguages
    
    Dim savedModel As String
    savedModel = GetLMStudioSetting("Model")
    If Len(savedModel) > 0 Then
        On Error Resume Next
        cmbLMStudioModel.value = savedModel
        On Error GoTo 0
    End If
    
    Dim savedLanguage As String
    savedLanguage = GetLMStudioSetting("ResponseLanguage")
    If Len(savedLanguage) = 0 Then savedLanguage = "English"
    SetResponseLanguageValue savedLanguage
End Sub

Private Sub LoadLMStudioModels()
    On Error Resume Next
    cmbLMStudioModel.Clear
    cmbLMStudioModel.AddItem "(auto)"
    
    Dim models As String
    Dim modelArr() As String
    Dim i As Long
    
    models = GetLMStudioModels()
    
    If Left(models, 6) <> "ERROR:" And Len(models) > 0 Then
        modelArr = Split(models, "|")
        For i = LBound(modelArr) To UBound(modelArr)
            If Len(Trim(modelArr(i))) > 0 Then
                cmbLMStudioModel.AddItem modelArr(i)
            End If
        Next i
    End If
End Sub

Private Sub LoadResponseLanguages()
    On Error Resume Next
    cmbResponseLanguage.Clear
    cmbResponseLanguage.AddItem "English"
    cmbResponseLanguage.AddItem "Russian"
    cmbResponseLanguage.AddItem "Ukrainian"
    cmbResponseLanguage.AddItem "Czech"
    cmbResponseLanguage.AddItem "Spanish"
    cmbResponseLanguage.AddItem "German"
    On Error GoTo 0
End Sub

Private Sub SetResponseLanguageValue(ByVal lang As String)
    Dim normalized As String
    normalized = NormalizeResponseLanguage(lang)

    On Error Resume Next
    cmbResponseLanguage.value = normalized
    If Len(cmbResponseLanguage.value) = 0 Then
        cmbResponseLanguage.ListIndex = 0
    End If
    On Error GoTo 0
End Sub

Private Function NormalizeResponseLanguage(ByVal lang As String) As String
    Select Case LCase(Trim(lang))
        Case "english": NormalizeResponseLanguage = "English"
        Case "russian": NormalizeResponseLanguage = "Russian"
        Case "ukrainian": NormalizeResponseLanguage = "Ukrainian"
        Case "czech": NormalizeResponseLanguage = "Czech"
        Case "spanish": NormalizeResponseLanguage = "Spanish"
        Case "german": NormalizeResponseLanguage = "German"
        Case Else: NormalizeResponseLanguage = "English"
    End Select
End Function

Private Sub btnRefreshModels_Click()
    lblLMStatus.Caption = "Loading models..."
    Me.Repaint
    
    If IsLMStudioAvailable() Then
        LoadLMStudioModels
        lblLMStatus.Caption = "LM Studio is available"
        lblLMStatus.ForeColor = &H8000&
    Else
        lblLMStatus.Caption = "Unavailable"
        lblLMStatus.ForeColor = &HFF&
    End If
End Sub

Private Sub btnSave_Click()
    SaveApiKey "DeepSeekKey", Trim(txtDeepSeekKey.value)
    SaveApiKey "OpenRouterKey", Trim(txtOpenRouterKey.value)
    SaveApiKey "OpenAIKey", Trim(txtOpenAIKey.value)
    
    SaveLMStudioSetting "IP", Trim(txtLMStudioIP.value)
    SaveLMStudioSetting "Port", Trim(txtLMStudioPort.value)
    
    Dim selectedModel As String
    selectedModel = cmbLMStudioModel.value
    If InStr(selectedModel, "(auto") > 0 Then selectedModel = ""
    SaveLMStudioSetting "Model", selectedModel
    
    Dim responseLanguage As String
    responseLanguage = NormalizeResponseLanguage(cmbResponseLanguage.value)
    SaveLMStudioSetting "ResponseLanguage", responseLanguage
    
    MsgBox "Settings saved!", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
