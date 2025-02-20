Attribute VB_Name = "MoudleNotes"
' ��ע�Ű湦��
Public Sub NoteFormat()
    On Error GoTo ErrorHandler
    
    ' ����Ƿ���ѡ�еĻõ�Ƭ
    If ActiveWindow Is Nothing Then
        MsgBox "����ѡ��һ���õ�Ƭ"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.SlideRange.Count = 0 Then
        MsgBox "����ѡ��һ���õ�Ƭ"
        Exit Sub
    End If
    
    ' ��ȡ��ǰҳ�汸ע����
    Dim userInput As String
    userInput = GetSlideNotesText()
    
    If Len(Trim(userInput)) = 0 Then
        MsgBox "��ǰҳ�汸עΪ�գ������������", vbInformation
        Exit Sub
    End If
    
    ' ��ȡ����
    Dim codeStr As String
    codeStr = GetCodeStringByRequest(userInput)
    
    If Len(codeStr) > 0 Then
        Dim ret As Boolean
        ret = RunDynamicCode(codeStr)
        If ret Then
            If Not SetSlideNotesText(codeStr) Then
                MsgBox "���ñ�עʧ��", vbCritical
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "�������: " & Err.Description, vbCritical
    Debug.Print "��������: " & Err.Source & ", �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Sub

' ��ע�Ի�����
Public Sub NoteChat()
    On Error GoTo ErrorHandler
    
    ' ����Ƿ���ѡ�еĻõ�Ƭ
    If ActiveWindow Is Nothing Then
        MsgBox "����ѡ��һ���õ�Ƭ"
        Exit Sub
    End If
    
    If ActiveWindow.Selection.SlideRange.Count = 0 Then
        MsgBox "����ѡ��һ���õ�Ƭ"
        Exit Sub
    End If
    
    ' ��ȡ��ǰҳ�汸ע����
    Dim userInput As String
    userInput = GetSlideNotesText()
    
    If Len(Trim(userInput)) = 0 Then
        MsgBox "��ǰҳ�汸עΪ�գ����������������", vbInformation
        Exit Sub
    End If
    
    ' ��ȡAI�ظ�
    Dim response As String
    response = GetChatResponse(userInput)
    
    ' ֱ�ӽ��ظ����ǵ���ע
    If Len(response) > 0 Then
        If Not SetSlideNotesText(response) Then
            MsgBox "���ñ�עʧ��", vbCritical
        End If
    Else
        MsgBox "��ȡ�ظ�ʧ��", vbCritical
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "�������: " & Err.Description, vbCritical
    Debug.Print "��������: " & Err.Source & ", �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Sub

' ��ȡ��ע�ı�
Public Function GetSlideNotesText() As String
    On Error GoTo ErrorHandler
    
    Dim slide As slide
    Dim notesText As String
    Dim slideIndex As String
    Dim shp As Shape
    Dim foundNotes As Boolean
    
    ' ��ȡѡ�еĻõ�Ƭ
    Set slide = ActiveWindow.Selection.SlideRange(1)
    slideIndex = ActiveWindow.View.Slide.SlideIndex
    ' ��ʼ���ұ�עҳ�ϵ���״
    foundNotes = False
        
    ' ������עҳ�ϵ���״
    For Each shp In slide.NotesPage.Shapes
        ' �����״�Ƿ�����ı����Ұ����ı�
        If shp.HasTextFrame And shp.TextFrame.HasText Then
            notesText = shp.TextFrame.TextRange.Text
            foundNotes = True
            If notesText = slideIndex Then
                foundNotes = False
            End If
            Exit For
        End If
    Next shp
        
    ' ���û���ҵ���ע�ı�
    If foundNotes Then
        GetSlideNotesText = notesText
    Else
        GetSlideNotesText = ""
    End If

    Exit Function
    
ErrorHandler:
    Debug.Print "GetSlideNotesText ����: " & Err.Description
    GetSlideNotesText = ""
End Function

' ���ñ�ע�ı�
Public Function SetSlideNotesText(noteStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim slide As slide
    Dim shp As Shape
    Dim foundNotes As Boolean
    
    ' ��ȡѡ�еĻõ�Ƭ
    Set slide = ActiveWindow.Selection.SlideRange(1)
        
    ' ��ʼ���ұ�עҳ�ϵ���״
    foundNotes = False
        
    ' ������עҳ�ϵ���״
    For Each shp In slide.NotesPage.Shapes
        ' �����״�Ƿ�����ı���
        If shp.HasTextFrame Then
            ' ���ñ�ע�ı�
            shp.TextFrame.TextRange.Text = noteStr
            foundNotes = True
            Exit For
        End If
    Next shp
        
    ' ���û���ҵ���ע�ı��������һ���µı�ע�ı�
    If Not foundNotes Then
        Set shp = slide.NotesPage.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 500, 200)
        shp.TextFrame.TextRange.Text = noteStr
    End If
    
    SetSlideNotesText = True
    Exit Function
    
ErrorHandler:
    Debug.Print "SetSlideNotesText ����: " & Err.Description
    SetSlideNotesText = False
End Function
