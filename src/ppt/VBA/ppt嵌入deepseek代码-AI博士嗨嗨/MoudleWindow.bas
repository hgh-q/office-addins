Attribute VB_Name = "MoudleWindow"
Option Explicit

' �򿪹��ߴ��ڵĺ���
Public Sub OpenTool()
    On Error GoTo ErrorHandler
    
    ' ��ʾ����
    DeepSeekTool.Show
    Exit Sub
    
ErrorHandler:
    MsgBox "�򿪹��ߴ���ʧ��: " & Err.Description, vbCritical
    Debug.Print "��������: " & Err.Source & ", �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Sub

' ����������ĺ���
Public Function ProcessToolInput(inputText As String) As String
    On Error GoTo ErrorHandler
    
    ' ��ȡ����
    Dim codeStr As String
    codeStr = GetCodeStringByRequest(inputText)
    
    If Len(codeStr) > 0 Then
        Dim ret As Boolean
        ret = RunDynamicCode(codeStr)
        If ret Then
            ProcessToolInput = "�������"
            ' ����ˮӡ
            DeepSeekTool.Controls("WatermarkLabel").Visible = False
        Else
            ProcessToolInput = "����ִ��ʧ��"
            ' ��ʾˮӡ
            DeepSeekTool.Controls("WatermarkLabel").Visible = True
        End If
    Else
        ProcessToolInput = "δ�ܻ�ȡ����Ч����"
        ' ��ʾˮӡ
        DeepSeekTool.Controls("WatermarkLabel").Visible = True
    End If
    
    Exit Function
    
ErrorHandler:
    ProcessToolInput = "�������: " & Err.Description
    ' ��ʾˮӡ
    On Error Resume Next
    DeepSeekTool.Controls("WatermarkLabel").Visible = True
    Debug.Print "��������: " & Err.Source & ", �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Function

' ��մ�������ĺ���
Public Sub ClearToolInput()
    On Error Resume Next
    With DeepSeekTool
        .TextBox1.Text = "�����������ı�..."  ' ʹ��Ĭ���ı�
        .TextBox1.ForeColor = RGB(128, 128, 128)  ' ����Ϊ��ɫ
        .TextBox2.Text = "����������ʾ������..."
        .Controls("WatermarkLabel").Visible = True  ' ��ʾˮӡ
    End With
End Sub
