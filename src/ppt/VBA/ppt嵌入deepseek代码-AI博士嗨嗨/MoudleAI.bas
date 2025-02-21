Attribute VB_Name = "MoudleAI"
' ��ȡ����
Public Function GetCodeStringByRequest(inputStr As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' ��ʼ�� HTTP ����
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' ���ݵ�ǰѡ���ģ������ URL �͹���������
    Select Case CurrentSettings.SelectedType
        Case APIType.QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputStr, True)  ' True ��ʾ����ģʽ
            
        Case APIType.DeepSeekV3, APIType.DeepSeekR1
            url = "https://api.deepseek.com/v1/chat/completions"
            requestBody = BuildDeepSeekRequest(inputStr, True)
            
        Case APIType.BailianV3, APIType.BailianR1
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildBailianRequest(inputStr, True)
            
        Case APIType.HuoshanV3, APIType.HuoshanR1
            url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
            requestBody = BuildHuoshanRequest(inputStr, True)
            
        Case APIType.LocalR1_14B, APIType.LocalR1_32B
            url = "https://api.siliconflow.cn/v1/chat/completions"
            requestBody = BuildLocalRequest(inputStr, True)
            
        Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
            url = "https://api.siliconflow.cn/v1/chat/completions"
            requestBody = BuildSiliconRequest(inputStr, True)
            
        Case Else
            ' Ĭ��ʹ�� QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputStr, True)
    End Select
    
    ' ��ȡ��ǰģ�͵� API Key
    Dim apiKey As String
    apiKey = GetCurrentAPIKey()
    
    ' ��������
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' �����Ӧ״̬
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "API��Ӧ��" & response
        GetCodeStringByRequest = ParseAPIResponse(response, CurrentSettings.SelectedType, False)
    Else
        Debug.Print "API����ʧ�ܣ�" & http.Status & " - " & http.statusText
        Debug.Print "�����壺" & requestBody
        Debug.Print "��Ӧ���ݣ�" & http.responseText
        GetCodeStringByRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "GetCodeStringByRequest ����: " & Err.Description
    GetCodeStringByRequest = ""
End Function

' ����ͨ��ǧ������
Public Function BuildQwenRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    BuildQwenRequest = "{""model"":""qwen-max-0125""," & _
                      """messages"":[" & _
                      "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                      "{""role"":""user"",""content"":""" & inputStr & """}" & _
                      "]}"
End Function

' ���� DeepSeek ����
Public Function BuildDeepSeekRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' ���ݵ�ǰѡ���ģ������ѡ��ͬ��ģ������
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.DeepSeekV3
            modelName = "deepseek-chat"
        Case APIType.DeepSeekR1
            modelName = "deepseek-reasoner"
        Case Else
            modelName = "deepseek-chat"  ' Ĭ��ʹ�� V3 �汾
    End Select
    
    ' ���� DeepSeek API ������
    BuildDeepSeekRequest = "{""model"":""" & modelName & """," & _
                          """messages"":[" & _
                          "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                          "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                          """stream"":false" & _
                          "}"
End Function

' ������������
Public Function BuildBailianRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' ���ݵ�ǰѡ���ģ������ѡ��ͬ��ģ������
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.BailianV3
            modelName = "deepseek-v3"
        Case APIType.BailianR1
            modelName = "deepseek-r1"
        Case Else
            modelName = "deepseek-v3"  ' Ĭ��ʹ�� v3 �汾
    End Select
    
    ' �������� API ������
    BuildBailianRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}" & _
                         "]}"
End Function

' ������ɽ��������
Public Function BuildHuoshanRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' ���ݵ�ǰѡ���ģ������ѡ��ͬ��ģ������
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.HuoshanV3
            modelName = "ep-20250212151644-m2nfh"
        Case APIType.HuoshanR1
            modelName = "ep-20250211090924-r9hdx"
        Case Else
            modelName = "ep-20250212151644-m2nfh"  ' Ĭ��ʹ�� V3 �汾
    End Select
    
    ' ������ɽ���� API ������
    BuildHuoshanRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                         """stream"":false" & _
                         "}"
End Function

' ������������
Public Function BuildLocalRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim modelName As String
    
    Select Case CurrentSettings.SelectedType
        Case APIType.LocalR1_14B
            modelName = "mydeepseek-r1:14b"
        Case APIType.LocalR1_32B
            modelName = "deepseek-ai/DeepSeek-R1-Distill-Llama-70B"
        Case Else
            modelName = "mydeepseek-r1:14b"  ' Ĭ��ʹ�� 14B �汾
    End Select
    
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' �������� API ������
    BuildLocalRequest = "{""model"":""" & modelName & """," & _
                       """messages"":[" & _
                       "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                       "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                       """stream"":false" & _
                       "}"
End Function

' ���������������
Public Function BuildSiliconRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "����һ�� PowerPoint VBA ���ר�ҡ���ֱ�ӷ��������� VBA ���룬��Ҫ�����κ�˵�����֣���Ҫʹ��markdown��ʽ�����漰��������ʱ����ʹ�� .Font.Name ������ .Font.NameFarEast ���������壬ȷ����Ӣ��ʹ����ͬ���塣�����ı���ȶ�����������ã���Ҫͬʱ���� .Font.Name �� .Font.NameFarEast ��ȷ������������ȷ��ʾ��"
    Else
        systemPrompt = "����һ���������˵�AI���֡�"
    End If
    
    ' ���������ַ����е������ַ�
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' ���ݵ�ǰѡ���ģ������ѡ��ͬ��ģ������
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.SiliconFlowV3
            modelName = "deepseek-ai/DeepSeek-V3"
        Case APIType.SiliconFlowR1
            modelName = "deepseek-ai/DeepSeek-R1"
        Case Else
            modelName = "deepseek-ai/DeepSeek-V3"  ' Ĭ��ʹ�� V3 �汾
    End Select
    
    ' ����������� API ������
    BuildSiliconRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                         """stream"":false," & _
                         """max_tokens"":4096" & _
                         "}"
End Function

' ���д���
Public Function RunDynamicCode(incodeStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' ��������ַ���
    Dim codeStr As String
    codeStr = Replace(incodeStr, "\n", vbCrLf)
    codeStr = Replace(codeStr, "\r", "")
    codeStr = Replace(codeStr, "```vba", "")
    codeStr = Replace(codeStr, "```", "")
    codeStr = Trim(codeStr)
    
    ' ��ȡ������
    Dim procName As String
    procName = ExtractProcedureName(codeStr)
    
    If procName = "" Then
        Debug.Print "�޷���ȡ������"
        RunDynamicCode = False
        Exit Function
    End If
    
    Debug.Print "׼��ִ�й���: " & procName
    Debug.Print "�������ݣ�" & vbCrLf & codeStr
    
    ' ��ȡ VBProject
    Dim pptApp As Object
    Set pptApp = Application
    
    Dim vbProj As Object
    Set vbProj = pptApp.ActivePresentation.VBProject
    
    ' ������ʱģ��
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    vbComp.Name = "TempModule_" & Format(Now, "yyyymmddhhnnss")
    
    ' ��Ӵ���
    vbComp.CodeModule.AddFromString codeStr
    
    ' ִ�д���
    On Error Resume Next
    Err.Clear
    
    ' ʹ�� Run ����ִ��
    Dim moduleAndProc As String
    moduleAndProc = vbComp.Name & "." & procName
    Debug.Print "ִ��: " & moduleAndProc
    
    pptApp.Run moduleAndProc
    
    If Err.Number <> 0 Then
        Debug.Print "����ʱ����: " & Err.Description & " (�����: " & Err.Number & ")"
        RunDynamicCode = False
    Else
        RunDynamicCode = True
    End If
    
    On Error GoTo ErrorHandler
    
    ' ����
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "RunDynamicCode ����: " & Err.Description
    Debug.Print "����λ��: " & Err.Source
    Debug.Print "�����: " & Err.Number
    
    If Not vbComp Is Nothing Then
        On Error Resume Next
        vbProj.VBComponents.Remove vbComp
    End If
    RunDynamicCode = False
End Function

Function ExtractProcedureName(codeStr As String) As String
    ' ʹ��������ʽ��ȡ������
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' ƥ�� Sub ��Ĺ�����
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = "Sub\s+([a-zA-Z_][a-zA-Z0-9_]*)"
    
    Dim matches As Object
    Set matches = regex.Execute(codeStr)
    
    If matches.Count > 0 Then
        ExtractProcedureName = matches(0).SubMatches(0)
    Else
        ExtractProcedureName = ""
    End If
End Function

' json��ȡֵ
Function GetJsonParsing(JsonString As String) As String
    On Error GoTo ErrorHandler
    
    Dim jsonDict As Object
    Dim content As String
    
    ' ��������� JSON �ַ���
    JsonString = Trim(JsonString)
    
    ' ��� JSON �ַ����Ƿ�Ϊ��
    If Len(JsonString) = 0 Then
        Debug.Print "JSON �ַ���Ϊ��"
        GetJsonParsing = ""
        Exit Function
    End If
    
    ' ���� JSON �ַ���
    Set jsonDict = ParseJson(JsonString)
    
    ' ����Ҫ���ֶ��Ƿ����
    If jsonDict Is Nothing Then
        Debug.Print "JSON ����ʧ�ܣ�����Ϊ��"
        GetJsonParsing = ""
        Exit Function
    End If
    
    If Not jsonDict.Exists("choices") Then
        Debug.Print "JSON ȱ�� choices �ֶ�"
        GetJsonParsing = ""
        Exit Function
    End If
    
    ' ��ȡ��������
    content = jsonDict("choices")(1)("message")("content")  ' �޸������������ 1 ��ʼ
    
    ' ����������ݣ��Ƴ� markdown ��ǣ�
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' �������������ַ�
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")
    
    ' �������
    Debug.Print "�����ɹ�"
    Debug.Print "��ȡ�Ĵ��룺" & vbCrLf & content
    
    GetJsonParsing = content
    Exit Function
    
ErrorHandler:
    Debug.Print "GetJsonParsing ����: " & Err.Description
    Debug.Print "����λ��: " & Err.Source
    Debug.Print "�����: " & Err.Number
    Debug.Print "ԭʼ JSON: " & JsonString
    GetJsonParsing = ""
End Function

' ��ȡ�Ի��ظ�
Public Function GetChatResponse(ByVal inputText As String) As String
    ' ʹ��ͳһ�Ĵ�����
    GetChatResponse = MoudleAPISettings.CallSelectedAPI(inputText, False)  ' False ��ʾ�Ǵ���ģʽ
End Function

' ���ı����뵽PPT
Public Function InsertTextToPPT(textContent As String, insertType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Debug.Print "����Ƿ���ѡ�еĻõ�Ƭ"
    ' ����Ƿ���ѡ�еĻõ�Ƭ
    If ActiveWindow Is Nothing Then
        MsgBox "����ѡ��һ���õ�Ƭ", vbInformation
        InsertTextToPPT = False
        Exit Function
    End If
    
    Debug.Print "���ȡ��ǰѡ�еĻõ�Ƭ"
    ' ��ȡ��ǰѡ�еĻõ�Ƭ
    Dim sld As slide
    ' If ActiveWindow.Selection.Type = ppSelectionSlides Then
    Set sld = ActiveWindow.Selection.SlideRange(1)
    ' Else
    '     MsgBox "����ѡ��һ���õ�Ƭ", vbInformation
    '     InsertTextToPPT = False
    '     Exit Function
    ' End If
    
    Debug.Print "д��õ�Ƭ"
    Select Case LCase(insertType)
        Case "textbox"  ' �����ı���
            With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 400, 300)
                With .TextFrame.TextRange
                    .Text = textContent
                    .Font.Size = 24
                    .Font.Name = "΢���ź�"
                    .Font.NameFarEast = "΢���ź�"
                End With
            End With
            
        Case "notes"   ' ���뱸ע
            With sld.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
                .Text = .Text & vbNewLine & textContent
            End With
            
        Case Else
            InsertTextToPPT = False
            Exit Function
    End Select
    
    InsertTextToPPT = True
    Exit Function
    
ErrorHandler:
    InsertTextToPPT = False
    Debug.Print "��������: InsertTextToPPT, �����: " & Err.Number
    Debug.Print "��������: " & Err.Description
End Function

Function GetChatResponseWithHistory(ByRef history As Collection) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim messagesJson As String
    
    ' ��ʼ�� HTTP ����
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' ��������� URL �� API Key
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    Dim apiKey As String
    apiKey = "sk-9a69f144f3fd4356b6741c56374e6eab"
    
    ' ������Ϣ��ʷJSON
    messagesJson = "{""role"":""system"",""content"":""����һ���������˵�AI���֡�""}"
    
    ' �����ʷ�Ի�
    Dim i As Long
    For i = 1 To history.Count
        Dim item As Variant
        item = history.item(i)
        messagesJson = messagesJson & ",{""role"":""" & item(0) & """,""content"":""" & _
                      Replace(Replace(Replace(item(1), vbCrLf, " "), vbCr, " "), """", "\""") & """}"
    Next i
    
    ' ����������������
    requestBody = "{""model"":""qwen-max-0125"",""messages"":[" & messagesJson & "]}"
    
    ' ��������
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' �����Ӧ״̬
    If http.Status = 200 Then
        response = http.responseText
        GetChatResponseWithHistory = GetJsonParsing(response)
    Else
        GetChatResponseWithHistory = "�Ի�����ʧ�ܣ�" & http.Status & " - " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    GetChatResponseWithHistory = "��������" & Err.Description
End Function

' ���� ParseAPIResponse ����
Public Function ParseAPIResponse(response As String, modelType As APIType, Optional isChat As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    Dim jsonDict As Object
    Dim content As String
    
    ' ��������� JSON �ַ���
    response = Trim(response)
    
    ' ��� JSON �ַ����Ƿ�Ϊ��
    If Len(response) = 0 Then
        Debug.Print "JSON �ַ���Ϊ��"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' ���� JSON �ַ���
    Set jsonDict = ParseJson(response)
    
    ' ����Ҫ���ֶ��Ƿ����
    If jsonDict Is Nothing Then
        Debug.Print "JSON ����ʧ�ܣ�����Ϊ��"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    If Not jsonDict.Exists("choices") Then
        Debug.Print "JSON ȱ�� choices �ֶ�"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' ��ȡ��������
    content = jsonDict("choices")(1)("message")("content")  ' �޸������������ 1 ��ʼ
    
    ' ����������ݣ��Ƴ� markdown ��ǣ�
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' �������������ַ�
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")
    
    ' �������
    Debug.Print "�����ɹ�"
    Debug.Print "��ȡ�Ĵ��룺" & vbCrLf & content
    
    ' ����ģ�����ʹ�����Ӧ
    ' Select Case modelType
    '     Case APIType.LocalR1_14B, APIType.LocalR1_32B
    '         ' ������ģ�͵���Ӧ\
    '         If Not jsonDict.Exists("choices") Then
    '             Debug.Print "JSON ȱ�� message �ֶ�"
    '             ParseAPIResponse = ""
    '             Exit Function
    '         End If
    '         ' content = jsonDict("message")("content")
    '         content = jsonDict("choices")(1)("message")("content")
    ' End Select
    
    ParseAPIResponse = content
    Exit Function
    
ErrorHandler:
    Debug.Print "ParseAPIResponse ����: " & Err.Description
    Debug.Print "����λ��: " & Err.Source
    Debug.Print "�����: " & Err.Number
    Debug.Print "ԭʼ��Ӧ: " & response
    ParseAPIResponse = ""
End Function


