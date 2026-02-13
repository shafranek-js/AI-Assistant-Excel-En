Attribute VB_Name = "modMain"
'========================================
' Main module: entry points
'========================================

Option Explicit

'----------------------------------------
' Show chat window
'----------------------------------------
Public Sub ShowAIAssistant()
    frmChat.Show vbModeless
End Sub

'----------------------------------------
' Show settings
'----------------------------------------
Public Sub ShowSettings()
    frmSettings.Show vbModal
End Sub

'----------------------------------------
' Create menu when the add-in loads
'----------------------------------------
Public Sub Auto_Open()
    CreateMenu
End Sub

'----------------------------------------
' Remove menu when the add-in unloads
'----------------------------------------
Public Sub Auto_Close()
    DeleteMenu
End Sub

'----------------------------------------
' Create menu button
'----------------------------------------
Private Sub CreateMenu()
    On Error Resume Next
    
    Dim cmdBar As CommandBar
    Dim cmdBtn As CommandBarButton
    
    ' Remove old button if it exists
    DeleteMenu
    
    ' For Excel 2007+, add button to worksheet menu bar
    Set cmdBar = Application.CommandBars("Worksheet Menu Bar")
    
    ' Add button
    Set cmdBtn = cmdBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    
    With cmdBtn
        .Caption = "AI Assistant"
        .Style = msoButtonCaption
        .OnAction = "ShowAIAssistant"
        .Tag = "AIAssistantButton"
    End With
End Sub

'----------------------------------------
' Remove menu
'----------------------------------------
Private Sub DeleteMenu()
    On Error Resume Next
    
    Dim ctrl As CommandBarControl
    
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Tag = "AIAssistantButton" Then
            ctrl.Delete
        End If
    Next ctrl
End Sub

'----------------------------------------
' Quick analysis of selected data
'----------------------------------------
Public Sub QuickAnalyze()
    Dim selectedData As String
    Dim context As String
    Dim response As String
    Dim model As String
    
    selectedData = GetSelectedData()
    
    If Len(selectedData) = 0 Then
        MsgBox "Select data to analyze first.", vbExclamation, "AI Assistant"
        Exit Sub
    End If
    
    context = GetWorkbookContext() & vbCrLf & selectedData
    
    ' Using DeepSeek by default
    model = "deepseek"
    If Not HasApiKey(model) Then
        model = "claude"
    End If
    
    If Not HasApiKey(model) Then
        MsgBox "API keys are not configured. Open Settings.", vbExclamation, "AI Assistant"
        ShowSettings
        Exit Sub
    End If
    
    Application.StatusBar = "AI is analyzing data..."
    
    response = SendToAI("Analyze this data and look for formatting or data quality issues:", model, context)
    
    Application.StatusBar = False
    
    MsgBox response, vbInformation, "Analysis Result"
End Sub

