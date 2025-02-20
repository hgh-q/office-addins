Attribute VB_Name = "MoudleAI"
' 获取代码
Public Function GetCodeStringByRequest(inputStr As String) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' 初始化 HTTP 请求
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 根据当前选择的模型设置 URL 和构建请求体
    Select Case CurrentSettings.SelectedType
        Case APIType.QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputStr, True)  ' True 表示代码模式
            
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
            ' 默认使用 QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputStr, True)
    End Select
    
    ' 获取当前模型的 API Key
    Dim apiKey As String
    apiKey = GetCurrentAPIKey()
    
    ' 发送请求
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' 检查响应状态
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "API响应：" & response
        GetCodeStringByRequest = ParseAPIResponse(response, CurrentSettings.SelectedType, False)
    Else
        Debug.Print "API调用失败：" & http.Status & " - " & http.statusText
        Debug.Print "请求体：" & requestBody
        Debug.Print "响应内容：" & http.responseText
        GetCodeStringByRequest = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "GetCodeStringByRequest 错误: " & Err.Description
    GetCodeStringByRequest = ""
End Function

' 构建通义千问请求
Public Function BuildQwenRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    BuildQwenRequest = "{""model"":""qwen-max-0125""," & _
                      """messages"":[" & _
                      "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                      "{""role"":""user"",""content"":""" & inputStr & """}" & _
                      "]}"
End Function

' 构建 DeepSeek 请求
Public Function BuildDeepSeekRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' 根据当前选择的模型类型选择不同的模型名称
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.DeepSeekV3
            modelName = "deepseek-chat"
        Case APIType.DeepSeekR1
            modelName = "deepseek-reasoner"
        Case Else
            modelName = "deepseek-chat"  ' 默认使用 V3 版本
    End Select
    
    ' 构建 DeepSeek API 请求体
    BuildDeepSeekRequest = "{""model"":""" & modelName & """," & _
                          """messages"":[" & _
                          "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                          "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                          """stream"":false" & _
                          "}"
End Function

' 构建百炼请求
Public Function BuildBailianRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' 根据当前选择的模型类型选择不同的模型名称
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.BailianV3
            modelName = "deepseek-v3"
        Case APIType.BailianR1
            modelName = "deepseek-r1"
        Case Else
            modelName = "deepseek-v3"  ' 默认使用 v3 版本
    End Select
    
    ' 构建百炼 API 请求体
    BuildBailianRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}" & _
                         "]}"
End Function

' 构建火山方舟请求
Public Function BuildHuoshanRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' 根据当前选择的模型类型选择不同的模型名称
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.HuoshanV3
            modelName = "ep-20250212151644-m2nfh"
        Case APIType.HuoshanR1
            modelName = "ep-20250211090924-r9hdx"
        Case Else
            modelName = "ep-20250212151644-m2nfh"  ' 默认使用 V3 版本
    End Select
    
    ' 构建火山方舟 API 请求体
    BuildHuoshanRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                         """stream"":false" & _
                         "}"
End Function

' 构建本地请求
Public Function BuildLocalRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim modelName As String
    
    Select Case CurrentSettings.SelectedType
        Case APIType.LocalR1_14B
            modelName = "mydeepseek-r1:14b"
        Case APIType.LocalR1_32B
            modelName = "deepseek-ai/DeepSeek-R1-Distill-Llama-70B"
        Case Else
            modelName = "mydeepseek-r1:14b"  ' 默认使用 14B 版本
    End Select
    
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' 构建本地 API 请求体
    BuildLocalRequest = "{""model"":""" & modelName & """," & _
                       """messages"":[" & _
                       "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                       "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                       """stream"":false" & _
                       "}"
End Function

' 构建硅基流动请求
Public Function BuildSiliconRequest(inputStr As String, isCodeMode As Boolean) As String
    Dim systemPrompt As String
    If isCodeMode Then
        systemPrompt = "你是一个 PowerPoint VBA 编程专家。请直接返回完整的 VBA 代码，不要包含任何说明文字，不要使用markdown格式。当涉及字体设置时，请使用 .Font.Name 而不是 .Font.NameFarEast 来设置字体，确保中英文使用相同字体。对于文本框等对象的字体设置，需要同时设置 .Font.Name 和 .Font.NameFarEast 以确保中文字体正确显示。"
    Else
        systemPrompt = "你是一个乐于助人的AI助手。"
    End If
    
    ' 处理输入字符串中的特殊字符
    inputStr = Replace(Replace(Replace(Replace(inputStr, vbCrLf, " "), vbCr, " "), vbLf, " "), """", "\""")
    
    ' 根据当前选择的模型类型选择不同的模型名称
    Dim modelName As String
    Select Case CurrentSettings.SelectedType
        Case APIType.SiliconFlowV3
            modelName = "deepseek-ai/DeepSeek-V3"
        Case APIType.SiliconFlowR1
            modelName = "deepseek-ai/DeepSeek-R1"
        Case Else
            modelName = "deepseek-ai/DeepSeek-V3"  ' 默认使用 V3 版本
    End Select
    
    ' 构建硅基流动 API 请求体
    BuildSiliconRequest = "{""model"":""" & modelName & """," & _
                         """messages"":[" & _
                         "{""role"":""system"",""content"":""" & systemPrompt & """}," & _
                         "{""role"":""user"",""content"":""" & inputStr & """}]," & _
                         """stream"":false," & _
                         """max_tokens"":4096" & _
                         "}"
End Function

' 运行代码
Public Function RunDynamicCode(incodeStr As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 清理代码字符串
    Dim codeStr As String
    codeStr = Replace(incodeStr, "\n", vbCrLf)
    codeStr = Replace(codeStr, "\r", "")
    codeStr = Replace(codeStr, "```vba", "")
    codeStr = Replace(codeStr, "```", "")
    codeStr = Trim(codeStr)
    
    ' 提取过程名
    Dim procName As String
    procName = ExtractProcedureName(codeStr)
    
    If procName = "" Then
        Debug.Print "无法提取过程名"
        RunDynamicCode = False
        Exit Function
    End If
    
    Debug.Print "准备执行过程: " & procName
    Debug.Print "代码内容：" & vbCrLf & codeStr
    
    ' 获取 VBProject
    Dim pptApp As Object
    Set pptApp = Application
    
    Dim vbProj As Object
    Set vbProj = pptApp.ActivePresentation.VBProject
    
    ' 创建临时模块
    Dim vbComp As Object
    Set vbComp = vbProj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
    vbComp.Name = "TempModule_" & Format(Now, "yyyymmddhhnnss")
    
    ' 添加代码
    vbComp.CodeModule.AddFromString codeStr
    
    ' 执行代码
    On Error Resume Next
    Err.Clear
    
    ' 使用 Run 方法执行
    Dim moduleAndProc As String
    moduleAndProc = vbComp.Name & "." & procName
    Debug.Print "执行: " & moduleAndProc
    
    pptApp.Run moduleAndProc
    
    If Err.Number <> 0 Then
        Debug.Print "运行时错误: " & Err.Description & " (错误号: " & Err.Number & ")"
        RunDynamicCode = False
    Else
        RunDynamicCode = True
    End If
    
    On Error GoTo ErrorHandler
    
    ' 清理
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "RunDynamicCode 错误: " & Err.Description
    Debug.Print "错误位置: " & Err.Source
    Debug.Print "错误号: " & Err.Number
    
    If Not vbComp Is Nothing Then
        On Error Resume Next
        vbProj.VBComponents.Remove vbComp
    End If
    RunDynamicCode = False
End Function

Function ExtractProcedureName(codeStr As String) As String
    ' 使用正则表达式提取过程名
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' 匹配 Sub 类的过程名
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

' json获取值
Function GetJsonParsing(JsonString As String) As String
    On Error GoTo ErrorHandler
    
    Dim jsonDict As Object
    Dim content As String
    
    ' 清理输入的 JSON 字符串
    JsonString = Trim(JsonString)
    
    ' 检查 JSON 字符串是否为空
    If Len(JsonString) = 0 Then
        Debug.Print "JSON 字符串为空"
        GetJsonParsing = ""
        Exit Function
    End If
    
    ' 解析 JSON 字符串
    Set jsonDict = ParseJson(JsonString)
    
    ' 检查必要的字段是否存在
    If jsonDict Is Nothing Then
        Debug.Print "JSON 解析失败：返回为空"
        GetJsonParsing = ""
        Exit Function
    End If
    
    If Not jsonDict.Exists("choices") Then
        Debug.Print "JSON 缺少 choices 字段"
        GetJsonParsing = ""
        Exit Function
    End If
    
    ' 获取代码内容
    content = jsonDict("choices")(1)("message")("content")  ' 修改这里的索引从 1 开始
    
    ' 清理代码内容（移除 markdown 标记）
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' 清理其他特殊字符
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")
    
    ' 调试输出
    Debug.Print "解析成功"
    Debug.Print "获取的代码：" & vbCrLf & content
    
    GetJsonParsing = content
    Exit Function
    
ErrorHandler:
    Debug.Print "GetJsonParsing 错误: " & Err.Description
    Debug.Print "错误位置: " & Err.Source
    Debug.Print "错误号: " & Err.Number
    Debug.Print "原始 JSON: " & JsonString
    GetJsonParsing = ""
End Function

' 获取对话回复
Public Function GetChatResponse(ByVal inputText As String) As String
    ' 使用统一的处理函数
    GetChatResponse = MoudleAPISettings.CallSelectedAPI(inputText, False)  ' False 表示非代码模式
End Function

' 将文本插入到PPT
Public Function InsertTextToPPT(textContent As String, insertType As String) As Boolean
    On Error GoTo ErrorHandler
    
    Debug.Print "检查是否有选中的幻灯片"
    ' 检查是否有选中的幻灯片
    If ActiveWindow Is Nothing Then
        MsgBox "请先选择一个幻灯片", vbInformation
        InsertTextToPPT = False
        Exit Function
    End If
    
    Debug.Print "检获取当前选中的幻灯片"
    ' 获取当前选中的幻灯片
    Dim sld As slide
    ' If ActiveWindow.Selection.Type = ppSelectionSlides Then
    Set sld = ActiveWindow.Selection.SlideRange(1)
    ' Else
    '     MsgBox "请先选择一个幻灯片", vbInformation
    '     InsertTextToPPT = False
    '     Exit Function
    ' End If
    
    Debug.Print "写入幻灯片"
    Select Case LCase(insertType)
        Case "textbox"  ' 插入文本框
            With sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 100, 100, 400, 300)
                With .TextFrame.TextRange
                    .Text = textContent
                    .Font.Size = 24
                    .Font.Name = "微软雅黑"
                    .Font.NameFarEast = "微软雅黑"
                End With
            End With
            
        Case "notes"   ' 插入备注
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
    Debug.Print "错误发生在: InsertTextToPPT, 错误号: " & Err.Number
    Debug.Print "错误描述: " & Err.Description
End Function

Function GetChatResponseWithHistory(ByRef history As Collection) As String
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    Dim messagesJson As String
    
    ' 初始化 HTTP 请求
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 设置请求的 URL 和 API Key
    url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    Dim apiKey As String
    apiKey = "sk-9a69f144f3fd4356b6741c56374e6eab"
    
    ' 构建消息历史JSON
    messagesJson = "{""role"":""system"",""content"":""你是一个乐于助人的AI助手。""}"
    
    ' 添加历史对话
    Dim i As Long
    For i = 1 To history.Count
        Dim item As Variant
        item = history.item(i)
        messagesJson = messagesJson & ",{""role"":""" & item(0) & """,""content"":""" & _
                      Replace(Replace(Replace(item(1), vbCrLf, " "), vbCr, " "), """", "\""") & """}"
    Next i
    
    ' 构建完整的请求体
    requestBody = "{""model"":""qwen-max-0125"",""messages"":[" & messagesJson & "]}"
    
    ' 发送请求
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' 检查响应状态
    If http.Status = 200 Then
        response = http.responseText
        GetChatResponseWithHistory = GetJsonParsing(response)
    Else
        GetChatResponseWithHistory = "对话请求失败：" & http.Status & " - " & http.statusText
    End If
    
    Exit Function
    
ErrorHandler:
    GetChatResponseWithHistory = "发生错误：" & Err.Description
End Function

' 更新 ParseAPIResponse 函数
Public Function ParseAPIResponse(response As String, modelType As APIType, Optional isChat As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    Dim jsonDict As Object
    Dim content As String
    
    ' 清理输入的 JSON 字符串
    response = Trim(response)
    
    ' 检查 JSON 字符串是否为空
    If Len(response) = 0 Then
        Debug.Print "JSON 字符串为空"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' 解析 JSON 字符串
    Set jsonDict = ParseJson(response)
    
    ' 检查必要的字段是否存在
    If jsonDict Is Nothing Then
        Debug.Print "JSON 解析失败：返回为空"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    If Not jsonDict.Exists("choices") Then
        Debug.Print "JSON 缺少 choices 字段"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' 获取代码内容
    content = jsonDict("choices")(1)("message")("content")  ' 修改这里的索引从 1 开始
    
    ' 清理代码内容（移除 markdown 标记）
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' 清理其他特殊字符
    content = Replace(content, "\n", vbCrLf)
    content = Replace(content, "\""", """")
    content = Replace(content, "\\", "\")
    
    ' 调试输出
    Debug.Print "解析成功"
    Debug.Print "获取的代码：" & vbCrLf & content
    
    ' 根据模型类型处理响应
    ' Select Case modelType
    '     Case APIType.LocalR1_14B, APIType.LocalR1_32B
    '         ' 处理本地模型的响应\
    '         If Not jsonDict.Exists("choices") Then
    '             Debug.Print "JSON 缺少 message 字段"
    '             ParseAPIResponse = ""
    '             Exit Function
    '         End If
    '         ' content = jsonDict("message")("content")
    '         content = jsonDict("choices")(1)("message")("content")
    ' End Select
    
    ParseAPIResponse = content
    Exit Function
    
ErrorHandler:
    Debug.Print "ParseAPIResponse 错误: " & Err.Description
    Debug.Print "错误位置: " & Err.Source
    Debug.Print "错误号: " & Err.Number
    Debug.Print "原始响应: " & response
    ParseAPIResponse = ""
End Function


