VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeepSeekTool 
   Caption         =   "SAIC-PPTAI助手"
   ClientHeight    =   7440
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   9780
   OleObjectBlob   =   "DeepSeekTool.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "DeepSeekTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' 不需要声明控件，因为它们已经在窗体设计器中创建
Private watermark As MSForms.Label

' 添加模块级变量存储对话历史
Private chatHistory As Collection

' 添加模块级变量
Private Const DEFAULT_INPUT_TEXT As String = "在这里输入文本..."

Private Sub UserForm_Initialize()
    ' 初始化窗体
    Me.Caption = "SAIC-PPTAI助手"
    Me.Width = 500
    Me.Height = 400
    Me.BackColor = RGB(240, 240, 240)  ' 浅灰色背景
    
    ' 设置标签
    With Label1
        .Caption = "输入内容（PPT排版需求/AI对话）："
        .Left = 12
        .Top = 12
        .Width = 150
        .AutoSize = True
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    With Label2
        .Caption = "处理结果："
        .Left = 12
        .Top = 165
        .AutoSize = True
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    ' 设置文本框
    With TextBox1
        .Width = 370
        .Height = 132
        .Left = 12
        .Top = 24
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Text = DEFAULT_INPUT_TEXT  ' 使用常量
        .Font.Name = "宋体"
        .Font.Size = 9
        .ForeColor = RGB(128, 128, 128)  ' 默认文本显示为灰色
    End With
    
    With TextBox2
        .Width = 462
        .Height = 180
        .Left = 12
        .Top = 180
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .BackColor = &H8000000F  ' 系统灰色
        .Text = "处理结果将显示在这里..."
        .Font.Name = "宋体"
        .Font.Size = 9
        
        ' 添加水印标签
        Set watermark = Me.Controls.Add("Forms.Label.1", "WatermarkLabel")
        With watermark
            .Caption = "AI博士嗨嗨"
            .Font.Size = 24
            .Font.Name = "微软雅黑"
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
    
    ' 设置右侧按钮组
    With ProcessButton
        .Caption = "智能排版"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 24
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    With ChatButton
        .Caption = "AI对话"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 60
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    With ClearButton
        .Caption = "清空"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 96
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    With SettingsButton
        .Caption = "模型设置"
        .Width = 80
        .Height = 25
        .Left = 395
        .Top = 132
        .Font.Name = "宋体"
        .Font.Size = 9
    End With
    
    ' 初始化对话历史
    Set chatHistory = New Collection
End Sub

' 当文本框获得焦点时
Private Sub TextBox1_Enter()
    If TextBox1.Text = DEFAULT_INPUT_TEXT Then
        TextBox1.Text = ""  ' 清除默认文本
    End If
End Sub

' 当文本框失去焦点时
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim(TextBox1.Text) = "" Then
        TextBox1.Text = DEFAULT_INPUT_TEXT  ' 恢复默认文本
    End If
End Sub

' 修改 TextBox1_Change
Private Sub TextBox1_Change()
    ' 根据是否是默认文本设置文本颜色
    If TextBox1.Text = DEFAULT_INPUT_TEXT Then
        TextBox1.ForeColor = RGB(128, 128, 128)  ' 灰色
    Else
        TextBox1.ForeColor = RGB(0, 0, 0)  ' 黑色
    End If
    
    TextBox2.Text = "处理结果将显示在这里..."
    watermark.Visible = True
End Sub

Private Sub ProcessButton_Click()
    TextBox2.Text = "正在生成代码..."
    watermark.Visible = False  ' 生成代码时隐藏水印
    ProcessText
End Sub

Private Sub ChatButton_Click()
    Debug.Print "开始 AI 对话..."
    Debug.Print "当前选择的模型类型: " & CurrentSettings.SelectedType
    
    ' 获取输入文本
    Dim inputText As String
    inputText = TextBox1.Text
    Debug.Print "输入文本: " & inputText
    
    ' 调用 API
    Dim response As String
    response = GetChatResponse(inputText)  ' 使用 MoudleAI 中的函数
    Debug.Print "API 返回响应: " & response
    
    If Left(response, 5) <> "Error" Then
        ' 显示结果
        If Len(response) > 0 Then
            TextBox2.Text = response
            watermark.Visible = False
            
            ' 询问是否插入到PPT
            If MsgBox("是否将回复内容插入到PPT中？", vbQuestion + vbYesNo) = vbYes Then
                ' 询问插入方式
                Dim insertType As String
                If MsgBox("是否插入为文本框？（否则插入为备注）", vbQuestion + vbYesNo) = vbYes Then
                    insertType = "textbox"
                Else
                    insertType = "notes"
                End If
                
                ' 插入内容
                If InsertTextToPPT(response, insertType) Then
                    MsgBox "内容已插入", vbInformation
                Else
                    MsgBox "插入失败", vbCritical
                End If
            End If
        Else
            Debug.Print "解析结果为空"
            MsgBox "获取回复失败，请重试", vbExclamation
            TextBox2.Text = "获取回复失败，请重试"
            watermark.Visible = True
        End If
    Else
        Debug.Print "API 调用错误: " & response
        MsgBox "调用失败：" & response, vbExclamation
        TextBox2.Text = "调用失败：" & response
        watermark.Visible = True
    End If
End Sub

Private Sub ClearButton_Click()
    TextBox1.Text = ""
    TextBox2.Text = "处理结果将显示在这里..."
    watermark.Visible = True
    ' 清空对话历史
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
        TextBox2.Text = "请输入要处理的内容"
        watermark.Visible = True
        Exit Sub
    End If
    
    If ActiveWindow Is Nothing Or ActiveWindow.Selection.SlideRange.Count = 0 Then
        TextBox2.Text = "请先选择一个幻灯片"
        Exit Sub
    End If
    
    Dim codeStr As String
    Dim ret As Boolean
    
    codeStr = GetCodeStringByRequest(noteStr)  ' 调用 MoudleAI 中的函数
    
    If Len(codeStr) > 0 Then
        Debug.Print "准备执行代码：" & vbCrLf & codeStr
        ret = RunDynamicCode(codeStr)
        If ret Then
            Call SetSlideNotesText(codeStr)
            TextBox2.Text = "处理完成"
            watermark.Visible = False
        Else
            TextBox2.Text = "代码执行失败，请查看即时窗口了解详细信息"
            watermark.Visible = False
        End If
    Else
        TextBox2.Text = "API 调用失败，请检查网络连接和 API Key"
        watermark.Visible = True
    End If
    
    Exit Sub

ErrorHandler:
    TextBox2.Text = "处理出错: " & Err.Description
    watermark.Visible = True
    Debug.Print "错误发生在: " & Err.Source & ", 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Sub

' 添加设置按钮点击事件
Private Sub SettingsButton_Click()
    ShowAPISettings
End Sub


