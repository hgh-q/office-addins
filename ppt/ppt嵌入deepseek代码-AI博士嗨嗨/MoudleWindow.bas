Attribute VB_Name = "MoudleWindow"
Option Explicit

' 打开工具窗口的函数
Public Sub OpenTool()
    On Error GoTo ErrorHandler
    
    ' 显示窗体
    DeepSeekTool.Show
    Exit Sub
    
ErrorHandler:
    MsgBox "打开工具窗口失败: " & Err.Description, vbCritical
    Debug.Print "错误发生在: " & Err.Source & ", 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Sub

' 处理窗体输入的函数
Public Function ProcessToolInput(inputText As String) As String
    On Error GoTo ErrorHandler
    
    ' 获取代码
    Dim codeStr As String
    codeStr = GetCodeStringByRequest(inputText)
    
    If Len(codeStr) > 0 Then
        Dim ret As Boolean
        ret = RunDynamicCode(codeStr)
        If ret Then
            ProcessToolInput = "处理完成"
            ' 隐藏水印
            DeepSeekTool.Controls("WatermarkLabel").Visible = False
        Else
            ProcessToolInput = "代码执行失败"
            ' 显示水印
            DeepSeekTool.Controls("WatermarkLabel").Visible = True
        End If
    Else
        ProcessToolInput = "未能获取到有效代码"
        ' 显示水印
        DeepSeekTool.Controls("WatermarkLabel").Visible = True
    End If
    
    Exit Function
    
ErrorHandler:
    ProcessToolInput = "处理出错: " & Err.Description
    ' 显示水印
    On Error Resume Next
    DeepSeekTool.Controls("WatermarkLabel").Visible = True
    Debug.Print "错误发生在: " & Err.Source & ", 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Function

' 清空窗体输入的函数
Public Sub ClearToolInput()
    On Error Resume Next
    With DeepSeekTool
        .TextBox1.Text = "在这里输入文本..."  ' 使用默认文本
        .TextBox1.ForeColor = RGB(128, 128, 128)  ' 设置为灰色
        .TextBox2.Text = "处理结果将显示在这里..."
        .Controls("WatermarkLabel").Visible = True  ' 显示水印
    End With
End Sub
