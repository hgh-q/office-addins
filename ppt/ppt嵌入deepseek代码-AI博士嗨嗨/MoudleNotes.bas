Attribute VB_Name = "MoudleNotes"
' 备注排版功能
Public Sub NoteFormat()
    On Error GoTo ErrorHandler
    
    ' 检查是否有选中的幻灯片
    If ActiveWindow Is Nothing Then
        MsgBox "请先选择一个幻灯片"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.SlideRange.Count = 0 Then
        MsgBox "请先选择一个幻灯片"
        Exit Sub
    End If
    
    ' 获取当前页面备注内容
    Dim userInput As String
    userInput = GetSlideNotesText()
    
    If Len(Trim(userInput)) = 0 Then
        MsgBox "当前页面备注为空，请先添加内容", vbInformation
        Exit Sub
    End If
    
    ' 获取代码
    Dim codeStr As String
    codeStr = GetCodeStringByRequest(userInput)
    
    If Len(codeStr) > 0 Then
        Dim ret As Boolean
        ret = RunDynamicCode(codeStr)
        If ret Then
            If Not SetSlideNotesText(codeStr) Then
                MsgBox "设置备注失败", vbCritical
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "处理出错: " & Err.Description, vbCritical
    Debug.Print "错误发生在: " & Err.Source & ", 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Sub

' 备注对话功能
Public Sub NoteChat()
    On Error GoTo ErrorHandler
    
    ' 检查是否有选中的幻灯片
    If ActiveWindow Is Nothing Then
        MsgBox "请先选择一个幻灯片"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.SlideRange.Count = 0 Then
        MsgBox "请先选择一个幻灯片"
        Exit Sub
    End If
    
    ' 获取当前页面备注内容
    Dim userInput As String
    userInput = GetSlideNotesText()
    
    If Len(Trim(userInput)) = 0 Then
        MsgBox "当前页面备注为空，请先添加问题内容", vbInformation
        Exit Sub
    End If
    
    ' 获取AI回复
    Dim response As String
    response = GetChatResponse(userInput)
    
    ' 直接将回复覆盖到备注
    If Len(response) > 0 Then
        If Not SetSlideNotesText(response) Then
            MsgBox "设置备注失败", vbCritical
        End If
    Else
        MsgBox "获取回复失败", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "处理出错: " & Err.Description, vbCritical
    Debug.Print "错误发生在: " & Err.Source & ", 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Sub

' 获取备注文本
Public Function GetSlideNotesText() As String
    On Error GoTo ErrorHandler
    
    Dim slide As slide
    Dim notesText As String
    Dim slideIndex As String
    Dim shp As Shape
    Dim foundNotes As Boolean
    
    ' 获取选中的幻灯片
    Set slide = ActiveWindow.Selection.SlideRange(1)
    slideIndex = ActiveWindow.View.Slide.SlideIndex
    ' 开始查找备注页上的形状
    foundNotes = False
        
    ' 遍历备注页上的形状
    For Each shp In slide.NotesPage.Shapes
        ' 检查形状是否包含文本框且包含文本
        If shp.HasTextFrame And shp.TextFrame.HasText Then
            notesText = shp.TextFrame.TextRange.Text
            foundNotes = True
            If notesText = slideIndex Then
                foundNotes = False
            End If
            Exit For
        End If
    Next shp
        
    ' 如果没有找到备注文本
    If foundNotes Then
        GetSlideNotesText = notesText
    Else
        GetSlideNotesText = ""
    End If

    Exit Function
    
ErrorHandler:
    Debug.Print "GetSlideNotesText 错误: " & Err.Description
    GetSlideNotesText = ""
End Function

' 设置备注文本
Public Function SetSlideNotesText(noteStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim slide As slide
    Dim shp As Shape
    Dim foundNotes As Boolean
    
    ' 获取选中的幻灯片
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' 开始查找备注页上的形状
    foundNotes = False
        
    ' 遍历备注页上的形状
    For Each shp In slide.NotesPage.Shapes
        ' 检查形状是否包含文本框
        If shp.HasTextFrame Then
            ' 设置备注文本
            shp.TextFrame.TextRange.Text = noteStr
            foundNotes = True
            Exit For
        End If
    Next shp
        
    ' 如果没有找到备注文本，则添加一个新的备注文本
    If Not foundNotes Then
        Set shp = slide.NotesPage.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 500, 200)
        shp.TextFrame.TextRange.Text = noteStr
    End If
    
    SetSlideNotesText = True
    Exit Function
    
ErrorHandler:
    Debug.Print "SetSlideNotesText 错误: " & Err.Description
    SetSlideNotesText = False
End Function
