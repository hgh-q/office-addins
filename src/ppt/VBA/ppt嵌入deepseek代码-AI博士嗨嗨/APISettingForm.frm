VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} APISettingForm 
   Caption         =   "API 设置"
   ClientHeight    =   8640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9980
   OleObjectBlob   =   "APISettingForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "APISettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' 删除未使用的类型定义
'Private Type LocalSettings
'    APIType As MoudleAPISettings.APIType
'End Type

Private Sub UserForm_Initialize()
    ' 设置窗体样式
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "API 设置"
    
    ' 初始化所有控件
    InitializeControls
    
    ' 加载保存的设置
    LoadSavedSettings
    
    ' 设置 API Keys (保持现有功能，但显示为掩码)
    TextBoxDeepSeekKey.Text = CurrentSettings.DeepSeekKey
    TextBoxBailianKey.Text = CurrentSettings.BailianKey
    TextBoxSiliconKey.Text = CurrentSettings.SiliconKey
    TextBoxHuoshanKey.Text = CurrentSettings.HuoshanKey
    
    ' 确保所有 API Key 输入框初始状态为隐藏
    TextBoxDeepSeekKey.PasswordChar = "*"
    TextBoxBailianKey.PasswordChar = "*"
    TextBoxSiliconKey.PasswordChar = "*"
    TextBoxHuoshanKey.PasswordChar = "*"
    
    ' 更新 UI 状态
    UpdateUIState
End Sub

Private Sub InitializeControls()
    ' 设置 Frame
    With Me.FrameAIModel
        .Caption = "AI模型"
        .Left = 10
        .Top = 10
        .Width = 480
        .Height = 380
    End With
    
    ' DeepSeek 组
    With Me.OptionDeepSeekV3
        .Caption = "DeepSeekV3"
        .Left = 20
        .Top = 20
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionDeepSeekR1
        .Caption = "DeepSeekR1"
        .Left = 180
        .Top = 20
        .Width = 120
        .Height = 20
    End With
    
    With Me.LabelDeepSeekKey
        .Caption = "API KEY:"
        .Left = 20
        .Top = 50
        .Width = 40
        .Height = 20
    End With
    
    With Me.TextBoxDeepSeekKey
        .Left = 65
        .Top = 48
        .Width = 395
        .Height = 20
        .PasswordChar = "*"
    End With
    
    ' 阿里云组
    With Me.OptionQwenMax
        .Caption = "通义千问 Max"
        .Left = 20
        .Top = 90
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionBailianV3
        .Caption = "阿里云百炼 DeepSeekV3"
        .Left = 180
        .Top = 90
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionBailianR1
        .Caption = "阿里云百炼 DeepSeekR1"
        .Left = 340
        .Top = 90
        .Width = 120
        .Height = 20
    End With
    
    With Me.LabelBailianKey
        .Caption = "API KEY:"
        .Left = 20
        .Top = 120
        .Width = 40
        .Height = 20
    End With
    
    With Me.TextBoxBailianKey
        .Left = 65
        .Top = 118
        .Width = 395
        .Height = 20
        .PasswordChar = "*"
    End With
    
    ' 硅基流动组
    With Me.OptionSiliconFlowV3
        .Caption = "硅基流动 DeepSeekV3"
        .Left = 20
        .Top = 160
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionSiliconFlowR1
        .Caption = "硅基流动 DeepSeekR1"
        .Left = 180
        .Top = 160
        .Width = 120
        .Height = 20
    End With
    
    With Me.LabelSiliconKey
        .Caption = "API KEY:"
        .Left = 20
        .Top = 190
        .Width = 40
        .Height = 20
    End With
    
    With Me.TextBoxSiliconKey
        .Left = 65
        .Top = 188
        .Width = 395
        .Height = 20
        .PasswordChar = "*"
    End With
    
    ' 火山方舟组
    With Me.OptionHuoshanV3
        .Caption = "火山方舟 DeepSeekV3"
        .Left = 20
        .Top = 230
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionHuoshanR1
        .Caption = "火山方舟 DeepSeekR1"
        .Left = 180
        .Top = 230
        .Width = 120
        .Height = 20
    End With
    
    With Me.LabelHuoshanKey
        .Caption = "API KEY:"
        .Left = 20
        .Top = 260
        .Width = 40
        .Height = 20
    End With
    
    With Me.TextBoxHuoshanKey
        .Left = 65
        .Top = 258
        .Width = 395
        .Height = 20
        .PasswordChar = "*"
    End With
    
    ' 本地部署组
    With Me.OptionLocalR1_14B
        .Caption = "本地 DeepSeekR1 14B"
        .Left = 20
        .Top = 300
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionLocalR1_32B
        .Caption = "本地 DeepSeekR1 32B"
        .Left = 180
        .Top = 300
        .Width = 120
        .Height = 20
    End With
    
    With Me.LabelLocalKey
        .Caption = "API KEY:"
        .Left = 20
        .Top = 330
        .Width = 40
        .Height = 20
    End With
    
    With Me.TextBoxLocalKey
        .Text = "ollama"
        .Left = 65
        .Top = 328
        .Width = 395
        .Height = 20
        .Enabled = False
    End With
    
    ' 确定按钮
    With Me.ConfirmButton
        .Caption = "确定"
        .Left = 205
        .Top = 400
        .Width = 80
        .Height = 25
    End With
End Sub

Private Sub LoadSavedSettings()
    ' 获取保存的设置
    Dim settings As Long
    settings = GetCurrentAPISettings()
    
    ' 先清除所有选项
    ClearAllOptions
    
    ' 根据保存的设置选择对应的选项
    Select Case settings
        Case 0  ' QwenMax
            Me.OptionQwenMax.Value = True
        Case 1  ' DeepSeekV3
            Me.OptionDeepSeekV3.Value = True
        Case 2  ' DeepSeekR1
            Me.OptionDeepSeekR1.Value = True
        Case 3  ' BailianV3
            Me.OptionBailianV3.Value = True
        Case 4  ' BailianR1
            Me.OptionBailianR1.Value = True
        Case 5  ' HuoshanV3
            Me.OptionHuoshanV3.Value = True
        Case 6  ' HuoshanR1
            Me.OptionHuoshanR1.Value = True
        Case 7  ' LocalR1_14B
            Me.OptionLocalR1_14B.Value = True
        Case 8  ' LocalR1_32B
            Me.OptionLocalR1_32B.Value = True
        Case 9  ' SiliconFlowV3
            Me.OptionSiliconFlowV3.Value = True
        Case 10 ' SiliconFlowR1
            Me.OptionSiliconFlowR1.Value = True
    End Select
    
    ' 更新 UI 状态
    UpdateUIState
End Sub

' 统一的 OptionButton Click 事件处理函数
Private Sub OptionButton_Click(ByVal keyType As String)
    EnableAPIKeyInputs keyType
End Sub

' 所有 OptionButton 的 Click 事件
Private Sub OptionDeepSeekV3_Click()
    OptionButton_Click "DeepSeek"
End Sub

Private Sub OptionDeepSeekR1_Click()
    OptionButton_Click "DeepSeek"
End Sub

Private Sub OptionQwenMax_Click()
    OptionButton_Click "Bailian"
End Sub

Private Sub OptionBailianV3_Click()
    OptionButton_Click "Bailian"
End Sub

Private Sub OptionBailianR1_Click()
    OptionButton_Click "Bailian"
End Sub

Private Sub OptionSiliconFlowV3_Click()
    OptionButton_Click "Silicon"
End Sub

Private Sub OptionSiliconFlowR1_Click()
    OptionButton_Click "Silicon"
End Sub

Private Sub OptionHuoshanV3_Click()
    OptionButton_Click "Huoshan"
End Sub

Private Sub OptionHuoshanR1_Click()
    OptionButton_Click "Huoshan"
End Sub

' 本地部署选项的 Click 事件
Private Sub OptionLocalR1_14B_Click()
    OptionButton_Click "Local"
End Sub

Private Sub OptionLocalR1_32B_Click()
    OptionButton_Click "Local"
End Sub

'/**
' * 统一处理文本框变更事件
' * @param txtBox - 发生变更的文本框对象
' * @param keyType - API Key 的类型（如 "DeepSeek", "Bailian" 等）
' */
Private Sub HandleTextBoxChange(ByVal txtBox As MSForms.TextBox, ByVal keyType As String)
    ' 如果当前无文本，则退出
    If Len(txtBox.Text) = 0 Then Exit Sub
    
    ' 保存输入的 API Key
    SaveAPIKey keyType, txtBox.Text
End Sub

' 确定按钮点击事件
Private Sub ConfirmButton_Click()
    SaveSettings
End Sub

Private Sub ClearAllOptions()
    ' 只清除所有选项的选中状态
    Me.OptionDeepSeekV3.Value = False
    Me.OptionDeepSeekR1.Value = False
    Me.OptionQwenMax.Value = False
    Me.OptionBailianV3.Value = False
    Me.OptionBailianR1.Value = False
    Me.OptionSiliconFlowV3.Value = False
    Me.OptionSiliconFlowR1.Value = False
    Me.OptionHuoshanV3.Value = False
    Me.OptionHuoshanR1.Value = False
    Me.OptionLocalR1_14B.Value = False
    Me.OptionLocalR1_32B.Value = False
    
    ' 保持本地部署选项的设置
    Me.TextBoxLocalKey.Text = "ollama"
    Me.TextBoxLocalKey.Enabled = False
End Sub

' 修改保存设置函数，只需要保存选中的模型类型
Private Sub SaveSettings()
    Dim APIType As Long
    
    APIType = GetSelectedModelType()
    If APIType = -1 Then
        MsgBox "请选择一个 AI 模型", vbExclamation
        Exit Sub
    End If
    
    ' 验证 API Key
    If Not MoudleAPISettings.IsLocalModel(APIType) Then
        Dim currentKey As String
        currentKey = GetCurrentAPIKey()
        If currentKey = "" Then
            MsgBox "请先设置所选模型的 API Key", vbExclamation
            Exit Sub
        End If
    End If
    
    ' 更新设置
    CurrentSettings.SelectedType = APIType
    
    ' 立即保存到注册表
    SaveSettingsToRegistry
    
    Me.Hide
End Sub

' 修改 GetSelectedModelType 函数的返回值类型声明
Private Function GetSelectedModelType() As Long
    On Error GoTo ErrorHandler
    
    Dim result As Long
    
    If Me.OptionQwenMax.Value Then
        result = 0  ' QwenMax
    ElseIf Me.OptionDeepSeekV3.Value Then
        result = 1  ' DeepSeekV3
    ElseIf Me.OptionDeepSeekR1.Value Then
        result = 2  ' DeepSeekR1
    ElseIf Me.OptionBailianV3.Value Then
        result = 3  ' BailianV3
    ElseIf Me.OptionBailianR1.Value Then
        result = 4  ' BailianR1
    ElseIf Me.OptionHuoshanV3.Value Then
        result = 5  ' HuoshanV3
    ElseIf Me.OptionHuoshanR1.Value Then
        result = 6  ' HuoshanR1
    ElseIf Me.OptionLocalR1_14B.Value Then
        result = 7  ' LocalR1_14B
    ElseIf Me.OptionLocalR1_32B.Value Then
        result = 8  ' LocalR1_32B
    ElseIf Me.OptionSiliconFlowV3.Value Then
        result = 9  ' SiliconFlowV3
    ElseIf Me.OptionSiliconFlowR1.Value Then
        result = 10 ' SiliconFlowR1
    Else
        result = -1
    End If
    
    GetSelectedModelType = result
    Exit Function
    
ErrorHandler:
    Debug.Print "错误发生在 GetSelectedModelType: " & Err.Description
    Debug.Print "错误号: " & Err.Number
    GetSelectedModelType = -1
End Function

' 修改 EnableAPIKeyInputs 函数，确保启用输入框时保持密码掩码
Private Sub EnableAPIKeyInputs(ByVal keyType As String)
    DisableAllAPIKeyInputs
    
    Select Case keyType
        Case "DeepSeek"
            Me.TextBoxDeepSeekKey.Enabled = True
        Case "Bailian"
            Me.TextBoxBailianKey.Enabled = True
        Case "Silicon"
            Me.TextBoxSiliconKey.Enabled = True
        Case "Huoshan"
            Me.TextBoxHuoshanKey.Enabled = True
    End Select
End Sub

' 修改 DisableAllAPIKeyInputs 函数，确保禁用时保持密码掩码
Private Sub DisableAllAPIKeyInputs()
    Me.TextBoxDeepSeekKey.Enabled = False
    Me.TextBoxBailianKey.Enabled = False
    Me.TextBoxSiliconKey.Enabled = False
    Me.TextBoxHuoshanKey.Enabled = False
    Me.TextBoxLocalKey.Enabled = False
    
    ' 确保禁用时显示掩码
    If Not Me.TextBoxDeepSeekKey.Text = "" Then Me.TextBoxDeepSeekKey.PasswordChar = "*"
    If Not Me.TextBoxBailianKey.Text = "" Then Me.TextBoxBailianKey.PasswordChar = "*"
    If Not Me.TextBoxSiliconKey.Text = "" Then Me.TextBoxSiliconKey.PasswordChar = "*"
    If Not Me.TextBoxHuoshanKey.Text = "" Then Me.TextBoxHuoshanKey.PasswordChar = "*"
End Sub

' 更新 UI 状态
Private Sub UpdateUIState()
    If Me.OptionDeepSeekV3.Value Or Me.OptionDeepSeekR1.Value Then
        EnableAPIKeyInputs "DeepSeek"
    ElseIf Me.OptionQwenMax.Value Or Me.OptionBailianV3.Value Or Me.OptionBailianR1.Value Then
        EnableAPIKeyInputs "Bailian"
    ElseIf Me.OptionSiliconFlowV3.Value Or Me.OptionSiliconFlowR1.Value Then
        EnableAPIKeyInputs "Silicon"
    ElseIf Me.OptionHuoshanV3.Value Or Me.OptionHuoshanR1.Value Then
        EnableAPIKeyInputs "Huoshan"
    ElseIf Me.OptionLocalR1_14B.Value Or Me.OptionLocalR1_32B.Value Then
        DisableAllAPIKeyInputs
    End If
End Sub

Private Sub TextBoxDeepSeekKey_Change()
    HandleTextBoxChange Me.TextBoxDeepSeekKey, "DeepSeek"
End Sub

Private Sub TextBoxBailianKey_Change()
    HandleTextBoxChange Me.TextBoxBailianKey, "Bailian"
End Sub

Private Sub TextBoxSiliconKey_Change()
    HandleTextBoxChange Me.TextBoxSiliconKey, "Silicon"
End Sub

Private Sub TextBoxHuoshanKey_Change()
    HandleTextBoxChange Me.TextBoxHuoshanKey, "Huoshan"
End Sub

'/**
' * 保存 API Key 到对应的设置中
' * @param keyType - API Key 的类型（如 "DeepSeek", "Bailian" 等）
' * @param apiKey - 要保存的 API Key 值
' */
Private Sub SaveAPIKey(ByVal keyType As String, ByVal apiKey As String)
    Debug.Print "=== SaveAPIKey 开始 ==="
    Debug.Print "保存 " & keyType & " API Key, 长度: " & Len(apiKey)
    
    Select Case keyType
        Case "DeepSeek"
            CurrentSettings.DeepSeekKey = apiKey
            Debug.Print "已保存到 DeepSeekKey"
        Case "Bailian"
            CurrentSettings.BailianKey = apiKey
            Debug.Print "已保存到 BailianKey"
        Case "Silicon"
            CurrentSettings.SiliconKey = apiKey
            Debug.Print "已保存到 SiliconKey"
        Case "Huoshan"
            CurrentSettings.HuoshanKey = apiKey
            Debug.Print "已保存到 HuoshanKey"
    End Select
    
    ' 立即保存到注册表以确保设置持久化
    SaveSettingsToRegistry
    
    Debug.Print "=== SaveAPIKey 结束 ==="
End Sub

' 在文本框获得焦点时自动显示内容
Private Sub TextBoxDeepSeekKey_GotFocus()
    TextBoxDeepSeekKey.PasswordChar = ""
End Sub

Private Sub TextBoxBailianKey_GotFocus()
    TextBoxBailianKey.PasswordChar = ""
End Sub

Private Sub TextBoxSiliconKey_GotFocus()
    TextBoxSiliconKey.PasswordChar = ""
End Sub

Private Sub TextBoxHuoshanKey_GotFocus()
    TextBoxHuoshanKey.PasswordChar = ""
End Sub

' 在文本框失去焦点时自动隐藏内容
Private Sub TextBoxDeepSeekKey_LostFocus()
    TextBoxDeepSeekKey.PasswordChar = "*"
End Sub

Private Sub TextBoxBailianKey_LostFocus()
    TextBoxBailianKey.PasswordChar = "*"
End Sub

Private Sub TextBoxSiliconKey_LostFocus()
    TextBoxSiliconKey.PasswordChar = "*"
End Sub

Private Sub TextBoxHuoshanKey_LostFocus()
    TextBoxHuoshanKey.PasswordChar = "*"
End Sub
