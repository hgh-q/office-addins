VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeepSeekTool 
   Caption         =   "SAIC-PPTAI����"
   ClientHeight    =   7440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780
   OleObjectBlob   =   "DeepSeekTool.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "DeepSeekTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' ����Ҫ�����ؼ�����Ϊ�����Ѿ��ڴ���������д���
Private watermark As MSForms.Label

' ���ģ�鼶�����洢�Ի���ʷ
Private chatHistory As Collection

' ���ģ�鼶����
Private Const DEFAULT_INPUT_TEXT As String = "�����������ı�..."

Private Sub UserForm_Initialize()
    ' ��ʼ������
    Me.Caption = "SAIC-PPTAI����"
    Me.Width = 500
    Me.Height = 400
    Me.BackColor = RGB(240, 240, 240)  ' ǳ��ɫ����
    
    ' ���ñ�ǩ
    With Label1
        .Caption = "�������ݣ�PPT�Ű�����/AI�Ի�����"
        .Left = 12
        .Top = 12
        .Width = 150
        .AutoSize = True
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    With Label2
        .Caption = "��������"
        .Left = 12
        .Top = 165
        .AutoSize = True
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    ' �����ı���
    With TextBox1
        .Width = 370
        .Height = 132
        .Left = 12
        .Top = 24
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Text = DEFAULT_INPUT_TEXT  ' ʹ�ó���
        .Font.Name = "����"
        .Font.Size = 9
        .ForeColor = RGB(128, 128, 128)  ' Ĭ���ı���ʾΪ��ɫ
    End With
    
    With TextBox2
        .Width = 462
        .Height = 180
        .Left = 12
        .Top = 180
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = &H8000000F  ' ϵͳ��ɫ
        .Text = "����������ʾ������..."
        .Font.Name = "����"
        .Font.Size = 9
        
        ' ���ˮӡ��ǩ
        Set watermark = Me.Controls.Add("Forms.Label.1", "WatermarkLabel")
        With watermark
            .Caption = "AI��ʿ����"
            .Font.Size = 24
            .Font.Name = "΢���ź�"
            .ForeColor = RGB(192, 192, 192)
            .BackStyle = fmBackStyleTransparent
            .TextAlign = fmTextAlignCenter
            .Width = 200
            .Height = 40
            .Left = TextBox2.Left + (TextBox2.Width - .Width) / 2
            .Top = TextBox2.Top + (TextBox2.Height - .Height) / 2
            .ZOrder (0)
        End With
    End With
    
    ' �����Ҳఴť��
    With ProcessButton
        .Caption = "�����Ű�"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 24
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    With ChatButton
        .Caption = "AI�Ի�"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 60
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    With ClearButton
        .Caption = "���"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 96
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    With SettingsButton
        .Caption = "ģ������"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 132
        .Font.Name = "����"
        .Font.Size = 9
    End With
    
    ' ��ʼ���Ի���ʷ
    Set chatHistory = New Collection
End Sub

' ���ı����ý���ʱ
Private Sub TextBox1_Enter()
    If TextBox1.Text = DEFAULT_INPUT_TEXT Then
        TextBox1.Text = ""  ' ���Ĭ���ı�
    End If
End Sub

' ���ı���ʧȥ����ʱ
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(TextBox1.Text) = "" Then
        TextBox1.Text = DEFAULT_INPUT_TEXT  ' �ָ�Ĭ���ı�
    End If
End Sub

' �޸� TextBox1_Change
Private Sub TextBox1_Change()
    ' �����Ƿ���Ĭ���ı������ı���ɫ
    If TextBox1.Text = DEFAULT_INPUT_TEXT Then
        TextBox1.ForeColor = RGB(128, 128, 128)  ' ��ɫ
    Else
        TextBox1.ForeColor = RGB(0, 0, 0)  ' ��ɫ
    End If
    
    TextBox2.Text = "����������ʾ������..."
    watermark.Visible = True
End Sub

Private Sub ProcessButton_Click()
    TextBox2.Text = "�������ɴ���..."
    watermark.Visible = False  ' ���ɴ���ʱ����ˮӡ
    ProcessText
End Sub

Private Sub ChatButton_Click()
    Debug.Print "��ʼ AI �Ի�..."
    Debug.Print "��ǰѡ���ģ������: " & CurrentSettings.SelectedType
    
    ' ��ȡ�����ı�
    Dim inputText As String
    inputText = TextBox1.Text
    Debug.Print "�����ı�: " & inputText
    
    ' ���� API
    Dim response As String
    response = GetChatResponse(inputText)  ' ʹ�� MoudleAI �еĺ���
    Debug.Print "API ������Ӧ: " & response
    
    If Left(response, 5) <> "Error" Then
        ' ��ʾ���
        If Len(response) > 0 Then
            TextBox2.Text = response
            watermark.Visible = False
            
            ' ѯ���Ƿ���뵽PPT
            If MsgBox("�Ƿ񽫻ظ����ݲ��뵽PPT�У�", vbQuestion + vbYesNo) = vbYes Then
                ' ѯ�ʲ��뷽ʽ
                Dim insertType As String
                If MsgBox("�Ƿ����Ϊ�ı��򣿣��������Ϊ��ע��", vbQuestion + vbYesNo) = vbYes Then
                    insertType = "textbox"
                Else
                    insertType = "notes"
                End If
                
                ' ��������
                If InsertTextToPPT(response, insertType) Then
                    MsgBox "�����Ѳ���", vbInformation
                Else
                    MsgBox "����ʧ��", vbCritical
                End If
            End If
        Else
            Debug.Print "�������Ϊ��"
            MsgBox "��ȡ�ظ�ʧ�ܣ�������", vbExclamation
            TextBox2.Text = "��ȡ�ظ�ʧ�ܣ�������"
            watermark.Visible = True
        End If
    Else
        Debug.Print "API ���ô���: " & response
        MsgBox "����ʧ�ܣ�" & response, vbExclamation
        TextBox2.Text = "����ʧ�ܣ�" & response
        watermark.Visible = True
    End If
End Sub

Private Sub ClearButton_Click()
    TextBox1.Text = ""
    TextBox2.Text = "����������ʾ������..."
    watermark.Visible = True
    ' ��նԻ���ʷ
    Set chatHistory = New Collection
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = 2 Then  ' Ctrl + Enter
        ProcessButton_Click
    ElseIf KeyCode = vbKeyEscape Then  ' ESC
        Me.Hide
    End If
End Sub

Private Sub ProcessText()
    On Error GoTo ErrorHandler
    
    ' If True Then
    If False Then
        Dim pptApp As Object
        Set pptApp = Application
        pptApp.Run "TempModule_20250220093411.CenterCurrentSlideContent"
        Exit Sub
    End If

    Dim noteStr As String
    noteStr = TextBox1.Text
    
    If Len(noteStr) = 0 Then
        TextBox2.Text = "������Ҫ���������"
        watermark.Visible = True
        Exit Sub
    End If
    
    If ActiveWindow Is Nothing Or ActiveWindow.Selection.SlideRange.Count = 0 Then
        TextBox2.Text = "����ѡ��һ���õ�Ƭ"
        Exit Sub
    End If
    
    Dim codeStr As String
    Dim ret As Boolean
    
    codeStr = GetCodeStringByRequest(noteStr)  ' ���� MoudleAI �еĺ���
    
    If Len(codeStr) > 0 Then
        Debug.Print "׼��ִ�д��룺" & vbCrLf & codeStr
        ret = RunDynamicCode(codeStr)
        If ret Then
            Call SetSlideNotesText(codeStr)
            TextBox2.Text = "�������"
            watermark.Visible = False
        Else
            TextBox2.Text = "����ִ��ʧ�ܣ���鿴��ʱ�����˽���ϸ��Ϣ"
            watermark.Visible = False
        End If
    Else
        TextBox2.Text = "API ����ʧ�ܣ������������Ӻ� API Key"
        watermark.Visible = True
    End If
    
    Exit Sub

ErrorHandler:
    TextBox2.Text = "�������: " & Err.Description
    watermark.Visible = True
    Debug.Print "��������: " & Err.Source & ", �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Sub

' ������ð�ť����¼�
Private Sub SettingsButton_Click()
    ShowAPISettings
End Sub


