Attribute VB_Name = "MoudleAPISettings"
Option Explicit

' API 类型枚举
Public Enum APIType
    QwenMax = 0         ' 通义千问
    DeepSeekV3 = 1      ' DeepSeek
    DeepSeekR1 = 2
    BailianV3 = 3       ' 百炼
    BailianR1 = 4
    HuoshanV3 = 5       ' 火山方舟
    HuoshanR1 = 6
    LocalR1_14B = 7     ' 本地部署
    LocalR1_32B = 8
    SiliconFlowV3 = 9   ' 硅基流动
    SiliconFlowR1 = 10
End Enum

' 当前 API 设置
Public Type APISettings
    SelectedType As APIType
    DeepSeekKey As String
    BailianKey As String
    SiliconKey As String
    HuoshanKey As String
End Type

' 当前设置实例
Public CurrentSettings As APISettings

' 在模块开始处添加
Private Sub InitializeSettings()
    If CurrentSettings.SelectedType = 0 Then  ' 如果未初始化
        CurrentSettings.SelectedType = APIType.QwenMax  ' 设置默认值
    End If
End Sub

' 显示设置窗体
Public Sub ShowAPISettings()
    ' 先加载保存的设置
    LoadSettingsFromRegistry
    
    ' 确保设置已初始化
    InitializeSettings
    
    Dim frmSettings As APISettingForm
    Set frmSettings = New APISettingForm
    frmSettings.Show vbModal
    
    ' 保存设置到注册表
    SaveSettingsToRegistry
End Sub

' 获取当前 API 设置
Public Function GetCurrentAPISettings() As APIType
    GetCurrentAPISettings = CurrentSettings.SelectedType
End Function

' 设置当前 API
Public Sub SetCurrentAPI(ByVal modelType As APIType, ByVal key As String)
    CurrentSettings.SelectedType = modelType
    Select Case GetModelGroup(modelType)
        Case "DeepSeek"
            CurrentSettings.DeepSeekKey = key
        Case "Bailian"
            CurrentSettings.BailianKey = key
        Case "Silicon"
            CurrentSettings.SiliconKey = key
        Case "Huoshan"
            CurrentSettings.HuoshanKey = key
    End Select
End Sub

' 获取当前 API Key
Public Function GetCurrentAPIKey() As String
    Debug.Print "=== GetCurrentAPIKey 开始 ==="
    Debug.Print "当前模型类型: " & CurrentSettings.SelectedType
    
    Dim apiKey As String
    Select Case CurrentSettings.SelectedType
        Case APIType.DeepSeekV3, APIType.DeepSeekR1
            apiKey = CurrentSettings.DeepSeekKey
            Debug.Print "使用 DeepSeek Key, 长度: " & Len(apiKey)
        Case APIType.QwenMax, APIType.BailianV3, APIType.BailianR1
            apiKey = CurrentSettings.BailianKey
            Debug.Print "使用 Bailian Key, 长度: " & Len(apiKey)
        Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
            apiKey = CurrentSettings.SiliconKey
            Debug.Print "使用 Silicon Key, 长度: " & Len(apiKey)
        Case APIType.HuoshanV3, APIType.HuoshanR1
            apiKey = CurrentSettings.HuoshanKey
            Debug.Print "使用 Huoshan Key, 长度: " & Len(apiKey)
        Case APIType.LocalR1_14B, APIType.LocalR1_32B
            apiKey = "sk-ooshywirgmrcdismctrllimnudbctvhhzybuzbqipervbrjy"
            Debug.Print "使用 Local Key: ollama"
    End Select
    
    Debug.Print "=== GetCurrentAPIKey 结束 ==="
    GetCurrentAPIKey = apiKey
End Function

' 根据当前设置获取 API 调用函数
Public Function GetAPIFunction() As String
    Select Case CurrentSettings.SelectedType
        Case APIType.DeepSeekV3
            GetAPIFunction = "DeepSeekV3"
        Case APIType.DeepSeekR1
            GetAPIFunction = "DeepSeekR1"
        Case APIType.SiliconFlowV3
            GetAPIFunction = "DeepSeekR1_SiliconFlow_V3"
        Case APIType.SiliconFlowR1
            GetAPIFunction = "DeepSeekR1_SiliconFlow"
        Case APIType.QwenMax
            GetAPIFunction = "QwenMax"
        Case APIType.BailianV3
            GetAPIFunction = "DeepSeekR1_Bailian_V3"
        Case APIType.BailianR1
            GetAPIFunction = "DeepSeekR1_Bailian"
        Case APIType.HuoshanV3
            GetAPIFunction = "DeepSeekR1_HuoshanArk_V3"
        Case APIType.HuoshanR1
            GetAPIFunction = "DeepSeekR1_HuoshanArk"
        Case APIType.LocalR1_14B
            GetAPIFunction = "LocalR1_14B"
        Case APIType.LocalR1_32B
            GetAPIFunction = "LocalR1_32B"
        Case Else
            GetAPIFunction = "QwenMax"
    End Select
End Function

' 添加新的函数
Public Function CallSelectedAPI(inputText As String, Optional isCodeMode As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CallSelectedAPI 开始 ==="
    Debug.Print "输入文本: " & inputText
    Debug.Print "代码模式: " & isCodeMode
    Debug.Print "当前模型: " & CurrentSettings.SelectedType
    
    Dim response As String
    response = CallAPI(CurrentSettings.SelectedType, inputText, isCodeMode)
    
    Debug.Print "API 返回响应: " & Left(response, 1000) & "..."  ' 只打印前100个字符
    
    If Left(response, 5) <> "Error" Then
        Debug.Print "开始解析响应..."
        CallSelectedAPI = ParseAPIResponse(response, CurrentSettings.SelectedType, isCodeMode)
    Else
        Debug.Print "API 调用出错"
        CallSelectedAPI = response
    End If
    
    Debug.Print "=== CallSelectedAPI 结束 ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "CallSelectedAPI 错误: " & Err.Description
    CallSelectedAPI = "Error: " & Err.Description
End Function

'/**
' * 解析不同 AI 模型的 API 响应，提取内容并处理思考过程
' *
' * @description
' * 该函数负责处理各种 AI 模型的 JSON 响应，包括：
' * 1. 基础响应内容的提取
' * 2. R1 系列模型的思考过程处理
' * 3. 特殊字符的清理和格式化
' * 4. 错误处理和日志记录
' *
' * @param response {String} - API 返回的原始 JSON 字符串
' * @param modelType {APIType} - 当前使用的模型类型（枚举值）
' * @param isCodeMode {Boolean} - 可选参数，默认为 False
' *                              True: 智能编辑模式，只返回最终结果
' *                              False: AI 对话模式，包含思考过程
' *
' * @return {String} 处理后的响应内容
' *                  对话模式下包含"思考过程"和"最终答案"
' *                  代码模式下仅包含最终结果
' *
' * @throws 可能抛出的错误：
' *         - JSON 解析错误
' *         - 响应格式错误
' *         - 字段缺失错误
' */
Public Function ParseAPIResponse(response As String, modelType As APIType, Optional isCodeMode As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    ' === 初始化与日志 ===
    Debug.Print "=== ParseAPIResponse 开始 ==="
    Debug.Print "模型类型: " & modelType
    Debug.Print "是否代码模式: " & isCodeMode
    Debug.Print "响应长度: " & Len(response)
    
    ' === 变量声明 ===
    Dim jsonDict As Object
    Dim content As String
    Dim reasoningContent As String
    
    ' === JSON 解析与验证 ===
    Set jsonDict = ParseJson(response)
    
    If jsonDict Is Nothing Then
        Debug.Print "JSON 解析失败：返回为空"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' === 本地模型特殊处理 ===
    If IsLocalModel(modelType) Then
        If Not jsonDict.Exists("choices") Then
            Debug.Print "JSON 缺少 choices 字段"
            ParseAPIResponse = ""
            Exit Function
        End If
        ' content = jsonDict("message")("content")

        If jsonDict("choices").Count > 0 Then
            content = jsonDict("choices")(1)("message")("content")
            ' reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")

            Debug.Print "Content: " & content
            Debug.Print "Reasoning Content: " & reasoningContent
        Else
            Debug.Print "Choices 数组为空"
            content = "Choices 数组为空"
        End If
        
        ' 根据模式处理内容
        If isCodeMode Then
            Debug.Print "不保留思考过程处理"
            ' 智能排版模式：提取代码块内容
            ' Dim codeStart As Long, codeEnd As Long
            
            ' ' 跳过思考部分，直接找代码块
            ' codeStart = InStr(1, content, "```vba")
            ' If codeStart = 0 Then
            '     codeStart = InStr(1, content, "```")
            ' End If
            
            ' If codeStart > 0 Then
            '     ' 跳过语言标识
            '     codeStart = InStr(codeStart + 3, content, vbLf)
            '     If codeStart > 0 Then
            '         codeStart = codeStart + 1
                    
            '         ' 查找结束标记
            '         codeEnd = InStr(codeStart, content, "```")
            '         If codeEnd > 0 Then
            '             content = Mid(content, codeStart, codeEnd - codeStart)
            '         End If
            '     End If
            ' End If
        Else
            ' AI 对话模式：保留思考过程
            Debug.Print "保留思考过程处理"
            ' content = Replace(content, "<think>", "思考过程：")
            ' content = Replace(content, "</think>" & vbCrLf, vbCrLf & "最终回答：")
        End If
        
        ' 清理内容
        content = Replace(content, "\n", vbCrLf)  ' 处理换行符
        content = Trim(content)
    Else
        ' === 模型特定响应处理 ===
        Select Case modelType
            Case APIType.QwenMax
                ' 通义千问系列的响应格式
                Debug.Print "处理通义千问响应"
                
                ' 验证响应结构
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON 缺少 choices 字段"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                
                ' 提取基础响应内容
                content = jsonDict("choices")(1)("message")("content")
                
                ' 通义千问不支持思考过程，直接返回内容
                Debug.Print "通义千问响应内容: " & content
                
            Case APIType.BailianV3, APIType.BailianR1
                ' 百炼系列模型处理
                Debug.Print "处理百炼系列响应"
                
                ' 验证响应结构
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON 缺少 choices 字段"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                
                ' 提取基础响应内容
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 版本特殊处理：在非代码模式下处理思考过程
                If modelType = APIType.BailianR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        ' 组合思考过程和最终答案
                        content = "思考过程：" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "最终答案：" & vbCrLf & content
                    End If
                End If
                
            Case APIType.DeepSeekV3, APIType.DeepSeekR1
                ' DeepSeek 系列的响应格式
                Debug.Print "处理 DeepSeek 响应"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON 缺少 choices 字段"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 版本的思考过程处理
                If modelType = APIType.DeepSeekR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "思考过程：" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "最终答案：" & vbCrLf & content
                    End If
                End If
                
            Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
                ' 硅基流动系列的响应格式
                Debug.Print "处理硅基流动响应"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON 缺少 choices 字段"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 版本的思考过程处理
                If modelType = APIType.SiliconFlowR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "思考过程：" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "最终答案：" & vbCrLf & content
                    End If
                End If
                
            Case APIType.HuoshanV3, APIType.HuoshanR1
                ' 火山方舟系列的响应格式
                Debug.Print "处理火山方舟响应"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON 缺少 choices 字段"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 版本的思考过程处理
                If modelType = APIType.HuoshanR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "思考过程：" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "最终答案：" & vbCrLf & content
                    End If
                End If
                
            Case Else
                Debug.Print "未知的模型类型"
                ParseAPIResponse = ""
                Exit Function
        End Select
    End If
    
    ' === 内容清理与格式化 ===
    ' 移除 markdown 标记
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' 处理特殊字符
    content = Replace(content, "\n", vbCrLf)  ' 换行符标准化
    content = Replace(content, "\""", """")   ' 引号转义处理
    content = Replace(content, "\\", "\")     ' 反斜杠转义处理
    
    ' === 返回结果 ===
    Debug.Print "最终返回内容: " & content
    ParseAPIResponse = content
    Exit Function
    
    ' === 错误处理 ===
ErrorHandler:
    Debug.Print "ParseAPIResponse 错误: " & Err.Description
    Debug.Print "错误位置: " & Err.Source
    Debug.Print "错误号: " & Err.Number
    Debug.Print "原始 JSON: " & response
    ParseAPIResponse = ""
End Function

' 判断是否为本地模型
Public Function IsLocalModel(ByVal modelType As APIType) As Boolean
    Select Case modelType
        Case LocalR1_14B, LocalR1_32B
            IsLocalModel = True
        Case Else
            IsLocalModel = False
    End Select
End Function

' 获取模型组
Public Function GetModelGroup(ByVal modelType As APIType) As String
    Select Case modelType
        Case DeepSeekV3, DeepSeekR1
            GetModelGroup = "DeepSeek"
        Case QwenMax, BailianV3, BailianR1
            GetModelGroup = "Bailian"
        Case SiliconFlowV3, SiliconFlowR1
            GetModelGroup = "Silicon"
        Case HuoshanV3, HuoshanR1
            GetModelGroup = "Huoshan"
        Case LocalR1_14B, LocalR1_32B
            GetModelGroup = "Local"
    End Select
End Function

' 在 CallSelectedAPI 函数前添加
Public Function CallAPI(ByVal modelType As APIType, ByVal inputText As String, Optional ByVal isCodeMode As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CallAPI 开始 ==="
    Debug.Print "模型类型: " & modelType
    Debug.Print "输入文本: " & inputText
    Debug.Print "代码模式: " & isCodeMode
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' 初始化 HTTP 请求
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' 获取当前模型的 API Key
    Dim apiKey As String
    apiKey = GetCurrentAPIKey()
    
    ' 验证 API Key
    If Not IsLocalModel(modelType) And Len(apiKey) = 0 Then
        Debug.Print "API Key 为空"
        CallAPI = "Error: API Key 未设置"
        Exit Function
    End If
    
    Debug.Print "使用的 API Key 长度: " & Len(apiKey)
    
    ' 根据当前选择的模型设置 URL 和构建请求体
    Select Case modelType
        Case APIType.QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputText, isCodeMode)
            
        Case APIType.DeepSeekV3, APIType.DeepSeekR1
            url = "https://api.deepseek.com/v1/chat/completions"
            requestBody = BuildDeepSeekRequest(inputText, isCodeMode)
            
        Case APIType.BailianV3, APIType.BailianR1
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildBailianRequest(inputText, isCodeMode)
            
        Case APIType.HuoshanV3, APIType.HuoshanR1
            url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
            requestBody = BuildHuoshanRequest(inputText, isCodeMode)
            
        Case APIType.LocalR1_14B, APIType.LocalR1_32B
            url = "https://api.siliconflow.cn/v1/chat/completions"
            requestBody = BuildLocalRequest(inputText, isCodeMode)
            
        Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
            url = "https://api.siliconflow.cn/v1/chat/completions"
            requestBody = BuildSiliconRequest(inputText, isCodeMode)
            
        Case Else
            ' 默认使用 QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputText, isCodeMode)
    End Select
    
    ' 发送请求前添加详细的请求信息调试
    Debug.Print String(50, "-")
    Debug.Print "完整请求信息:"
    Debug.Print "URL: " & url
    Debug.Print "Content-Type: application/json; charset=utf-8"
    Debug.Print "Authorization: Bearer " & apiKey  ' 输出完整的 API Key 用于调试
    Debug.Print "请求体: " & requestBody
    Debug.Print String(50, "-")
    
    ' 发送请求
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' 响应处理时添加更多信息
    Debug.Print String(50, "-")
    Debug.Print "响应详情:"
    Debug.Print "状态码: " & http.Status
    Debug.Print "状态文本: " & http.statusText
    If http.Status <> 200 Then
        Debug.Print "错误响应: " & http.responseText
    End If
    Debug.Print String(50, "-")
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "API调用成功，响应长度: " & Len(response)
        CallAPI = response
    ElseIf http.Status = 401 Then
        Debug.Print "API认证失败：" & http.Status & " - " & http.statusText
        Debug.Print "响应内容：" & http.responseText
        CallAPI = "Error: API认证失败 - 请检查 API Key 是否正确"
    Else
        Debug.Print "API调用失败：" & http.Status & " - " & http.statusText
        Debug.Print "请求体：" & requestBody
        Debug.Print "响应内容：" & http.responseText
        CallAPI = "Error: API调用失败 - " & http.Status & " " & http.statusText
    End If
    
    Debug.Print "=== CallAPI 结束 ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "CallAPI 错误: " & Err.Description
    CallAPI = "Error: " & Err.Description
End Function

' 保存设置到注册表
Public Sub SaveSettingsToRegistry()
    Debug.Print "=== SaveSettingsToRegistry 开始 ==="
    Debug.Print "保存前的设置:"
    Debug.Print "SelectedType: " & CurrentSettings.SelectedType
    Debug.Print "DeepSeekKey: " & CurrentSettings.DeepSeekKey
    Debug.Print "BailianKey: " & CurrentSettings.BailianKey
    Debug.Print "SiliconKey: " & CurrentSettings.SiliconKey
    Debug.Print "HuoshanKey: " & CurrentSettings.HuoshanKey
    
    ' 保存设置
    SaveSetting "DeepSeekPPT", "Settings", "SelectedType", CStr(CurrentSettings.SelectedType)
    SaveSetting "DeepSeekPPT", "Settings", "DeepSeekKey", CurrentSettings.DeepSeekKey
    SaveSetting "DeepSeekPPT", "Settings", "BailianKey", CurrentSettings.BailianKey
    SaveSetting "DeepSeekPPT", "Settings", "SiliconKey", CurrentSettings.SiliconKey
    SaveSetting "DeepSeekPPT", "Settings", "HuoshanKey", CurrentSettings.HuoshanKey
    
    Debug.Print "=== SaveSettingsToRegistry 结束 ==="
End Sub

' 从注册表加载设置
Public Sub LoadSettingsFromRegistry()
    Debug.Print "=== LoadSettingsFromRegistry 开始 ==="
    
    ' 读取设置前的状态
    Debug.Print "读取前的设置:"
    Debug.Print "当前类型: " & CurrentSettings.SelectedType
    
    ' 读取设置
    CurrentSettings.SelectedType = CLng(GetSetting("DeepSeekPPT", "Settings", "SelectedType", "0"))
    CurrentSettings.DeepSeekKey = GetSetting("DeepSeekPPT", "Settings", "DeepSeekKey", "")
    CurrentSettings.BailianKey = GetSetting("DeepSeekPPT", "Settings", "BailianKey", "")
    CurrentSettings.SiliconKey = GetSetting("DeepSeekPPT", "Settings", "SiliconKey", "")
    CurrentSettings.HuoshanKey = GetSetting("DeepSeekPPT", "Settings", "HuoshanKey", "")
    
    ' 输出调试信息
    Debug.Print "读取后的设置:"
    Debug.Print "SelectedType: " & CurrentSettings.SelectedType
    Debug.Print "DeepSeekKey 长度: " & Len(CurrentSettings.DeepSeekKey)
    If Len(CurrentSettings.DeepSeekKey) > 0 Then
        Debug.Print "DeepSeekKey 前5位: " & Left(CurrentSettings.DeepSeekKey, 5)
    End If
    Debug.Print "BailianKey 长度: " & Len(CurrentSettings.BailianKey)
    Debug.Print "SiliconKey 长度: " & Len(CurrentSettings.SiliconKey)
    Debug.Print "HuoshanKey 长度: " & Len(CurrentSettings.HuoshanKey)
    
    ' 验证当前模型是否有对应的 API Key
    Dim currentKey As String
    currentKey = GetCurrentAPIKey()
    Debug.Print "当前模型的 API Key 长度: " & Len(currentKey)
    
    Debug.Print "=== LoadSettingsFromRegistry 结束 ==="
End Sub

