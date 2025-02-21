VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} APISettingForm 
   Caption         =   "API ����"
   ClientHeight    =   8640
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9980
   OleObjectBlob   =   "APISettingForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "APISettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ɾ��δʹ�õ����Ͷ���
'Private Type LocalSettings
'    APIType As MoudleAPISettings.APIType
'End Type

Private Sub UserForm_Initialize()
    ' ���ô�����ʽ
    Me.BackColor = RGB(240, 240, 240)
    Me.Caption = "API ����"
    
    ' ��ʼ�����пؼ�
    InitializeControls
    
    ' ���ر��������
    LoadSavedSettings
    
    ' ���� API Keys (�������й��ܣ�����ʾΪ����)
    TextBoxDeepSeekKey.Text = CurrentSettings.DeepSeekKey
    TextBoxBailianKey.Text = CurrentSettings.BailianKey
    TextBoxSiliconKey.Text = CurrentSettings.SiliconKey
    TextBoxHuoshanKey.Text = CurrentSettings.HuoshanKey
    
    ' ȷ������ API Key ������ʼ״̬Ϊ����
    TextBoxDeepSeekKey.PasswordChar = "*"
    TextBoxBailianKey.PasswordChar = "*"
    TextBoxSiliconKey.PasswordChar = "*"
    TextBoxHuoshanKey.PasswordChar = "*"
    
    ' ���� UI ״̬
    UpdateUIState
End Sub

Private Sub InitializeControls()
    ' ���� Frame
    With Me.FrameAIModel
        .Caption = "AIģ��"
        .Left = 10
        .Top = 10
        .Width = 480
        .Height = 380
    End With
    
    ' DeepSeek ��
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
    
    ' ��������
    With Me.OptionQwenMax
        .Caption = "ͨ��ǧ�� Max"
        .Left = 20
        .Top = 90
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionBailianV3
        .Caption = "�����ư��� DeepSeekV3"
        .Left = 180
        .Top = 90
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionBailianR1
        .Caption = "�����ư��� DeepSeekR1"
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
    
    ' ���������
    With Me.OptionSiliconFlowV3
        .Caption = "������� DeepSeekV3"
        .Left = 20
        .Top = 160
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionSiliconFlowR1
        .Caption = "������� DeepSeekR1"
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
    
    ' ��ɽ������
    With Me.OptionHuoshanV3
        .Caption = "��ɽ���� DeepSeekV3"
        .Left = 20
        .Top = 230
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionHuoshanR1
        .Caption = "��ɽ���� DeepSeekR1"
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
    
    ' ���ز�����
    With Me.OptionLocalR1_14B
        .Caption = "���� DeepSeekR1 14B"
        .Left = 20
        .Top = 300
        .Width = 120
        .Height = 20
    End With
    
    With Me.OptionLocalR1_32B
        .Caption = "���� DeepSeekR1 32B"
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
    
    ' ȷ����ť
    With Me.ConfirmButton
        .Caption = "ȷ��"
        .Left = 205
        .Top = 400
        .Width = 80
        .Height = 25
    End With
End Sub

Private Sub LoadSavedSettings()
    ' ��ȡ���������
    Dim settings As Long
    settings = GetCurrentAPISettings()
    
    ' ���������ѡ��
    ClearAllOptions
    
    ' ���ݱ��������ѡ���Ӧ��ѡ��
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
    
    ' ���� UI ״̬
    UpdateUIState
End Sub

' ͳһ�� OptionButton Click �¼�������
Private Sub OptionButton_Click(ByVal keyType As String)
    EnableAPIKeyInputs keyType
End Sub

' ���� OptionButton �� Click �¼�
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

' ���ز���ѡ��� Click �¼�
Private Sub OptionLocalR1_14B_Click()
    OptionButton_Click "Local"
End Sub

Private Sub OptionLocalR1_32B_Click()
    OptionButton_Click "Local"
End Sub

'/**
' * ͳһ�����ı������¼�
' * @param txtBox - ����������ı������
' * @param keyType - API Key �����ͣ��� "DeepSeek", "Bailian" �ȣ�
' */
Private Sub HandleTextBoxChange(ByVal txtBox As MSForms.TextBox, ByVal keyType As String)
    ' �����ǰ���ı������˳�
    If Len(txtBox.Text) = 0 Then Exit Sub
    
    ' ��������� API Key
    SaveAPIKey keyType, txtBox.Text
End Sub

' ȷ����ť����¼�
Private Sub ConfirmButton_Click()
    SaveSettings
End Sub

Private Sub ClearAllOptions()
    ' ֻ�������ѡ���ѡ��״̬
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
    
    ' ���ֱ��ز���ѡ�������
    Me.TextBoxLocalKey.Text = "ollama"
    Me.TextBoxLocalKey.Enabled = False
End Sub

' �޸ı������ú�����ֻ��Ҫ����ѡ�е�ģ������
Private Sub SaveSettings()
    Dim APIType As Long
    
    APIType = GetSelectedModelType()
    If APIType = -1 Then
        MsgBox "��ѡ��һ�� AI ģ��", vbExclamation
        Exit Sub
    End If
    
    ' ��֤ API Key
    If Not MoudleAPISettings.IsLocalModel(APIType) Then
        Dim currentKey As String
        currentKey = GetCurrentAPIKey()
        If currentKey = "" Then
            MsgBox "����������ѡģ�͵� API Key", vbExclamation
            Exit Sub
        End If
    End If
    
    ' ��������
    CurrentSettings.SelectedType = APIType
    
    ' �������浽ע���
    SaveSettingsToRegistry
    
    Me.Hide
End Sub

' �޸� GetSelectedModelType �����ķ���ֵ��������
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
    Debug.Print "�������� GetSelectedModelType: " & Err.Description
    Debug.Print "�����: " & Err.Number
    GetSelectedModelType = -1
End Function

' �޸� EnableAPIKeyInputs ������ȷ�����������ʱ������������
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

' �޸� DisableAllAPIKeyInputs ������ȷ������ʱ������������
Private Sub DisableAllAPIKeyInputs()
    Me.TextBoxDeepSeekKey.Enabled = False
    Me.TextBoxBailianKey.Enabled = False
    Me.TextBoxSiliconKey.Enabled = False
    Me.TextBoxHuoshanKey.Enabled = False
    Me.TextBoxLocalKey.Enabled = False
    
    ' ȷ������ʱ��ʾ����
    If Not Me.TextBoxDeepSeekKey.Text = "" Then Me.TextBoxDeepSeekKey.PasswordChar = "*"
    If Not Me.TextBoxBailianKey.Text = "" Then Me.TextBoxBailianKey.PasswordChar = "*"
    If Not Me.TextBoxSiliconKey.Text = "" Then Me.TextBoxSiliconKey.PasswordChar = "*"
    If Not Me.TextBoxHuoshanKey.Text = "" Then Me.TextBoxHuoshanKey.PasswordChar = "*"
End Sub

' ���� UI ״̬
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
' * ���� API Key ����Ӧ��������
' * @param keyType - API Key �����ͣ��� "DeepSeek", "Bailian" �ȣ�
' * @param apiKey - Ҫ����� API Key ֵ
' */
Private Sub SaveAPIKey(ByVal keyType As String, ByVal apiKey As String)
    Debug.Print "=== SaveAPIKey ��ʼ ==="
    Debug.Print "���� " & keyType & " API Key, ����: " & Len(apiKey)
    
    Select Case keyType
        Case "DeepSeek"
            CurrentSettings.DeepSeekKey = apiKey
            Debug.Print "�ѱ��浽 DeepSeekKey"
        Case "Bailian"
            CurrentSettings.BailianKey = apiKey
            Debug.Print "�ѱ��浽 BailianKey"
        Case "Silicon"
            CurrentSettings.SiliconKey = apiKey
            Debug.Print "�ѱ��浽 SiliconKey"
        Case "Huoshan"
            CurrentSettings.HuoshanKey = apiKey
            Debug.Print "�ѱ��浽 HuoshanKey"
    End Select
    
    ' �������浽ע�����ȷ�����ó־û�
    SaveSettingsToRegistry
    
    Debug.Print "=== SaveAPIKey ���� ==="
End Sub

' ���ı����ý���ʱ�Զ���ʾ����
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

' ���ı���ʧȥ����ʱ�Զ���������
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
