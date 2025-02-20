Attribute VB_Name = "MoudleAPISettings"
Option Explicit

' API ����ö��
Public Enum APIType
    QwenMax = 0         ' ͨ��ǧ��
    DeepSeekV3 = 1      ' DeepSeek
    DeepSeekR1 = 2
    BailianV3 = 3       ' ����
    BailianR1 = 4
    HuoshanV3 = 5       ' ��ɽ����
    HuoshanR1 = 6
    LocalR1_14B = 7     ' ���ز���
    LocalR1_32B = 8
    SiliconFlowV3 = 9   ' �������
    SiliconFlowR1 = 10
End Enum

' ��ǰ API ����
Public Type APISettings
    SelectedType As APIType
    DeepSeekKey As String
    BailianKey As String
    SiliconKey As String
    HuoshanKey As String
End Type

' ��ǰ����ʵ��
Public CurrentSettings As APISettings

' ��ģ�鿪ʼ�����
Private Sub InitializeSettings()
    If CurrentSettings.SelectedType = 0 Then  ' ���δ��ʼ��
        CurrentSettings.SelectedType = APIType.QwenMax  ' ����Ĭ��ֵ
    End If
End Sub

' ��ʾ���ô���
Public Sub ShowAPISettings()
    ' �ȼ��ر��������
    LoadSettingsFromRegistry
    
    ' ȷ�������ѳ�ʼ��
    InitializeSettings
    
    Dim frmSettings As APISettingForm
    Set frmSettings = New APISettingForm
    frmSettings.Show vbModal
    
    ' �������õ�ע���
    SaveSettingsToRegistry
End Sub

' ��ȡ��ǰ API ����
Public Function GetCurrentAPISettings() As APIType
    GetCurrentAPISettings = CurrentSettings.SelectedType
End Function

' ���õ�ǰ API
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

' ��ȡ��ǰ API Key
Public Function GetCurrentAPIKey() As String
    Debug.Print "=== GetCurrentAPIKey ��ʼ ==="
    Debug.Print "��ǰģ������: " & CurrentSettings.SelectedType
    
    Dim apiKey As String
    Select Case CurrentSettings.SelectedType
        Case APIType.DeepSeekV3, APIType.DeepSeekR1
            apiKey = CurrentSettings.DeepSeekKey
            Debug.Print "ʹ�� DeepSeek Key, ����: " & Len(apiKey)
        Case APIType.QwenMax, APIType.BailianV3, APIType.BailianR1
            apiKey = CurrentSettings.BailianKey
            Debug.Print "ʹ�� Bailian Key, ����: " & Len(apiKey)
        Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
            apiKey = CurrentSettings.SiliconKey
            Debug.Print "ʹ�� Silicon Key, ����: " & Len(apiKey)
        Case APIType.HuoshanV3, APIType.HuoshanR1
            apiKey = CurrentSettings.HuoshanKey
            Debug.Print "ʹ�� Huoshan Key, ����: " & Len(apiKey)
        Case APIType.LocalR1_14B, APIType.LocalR1_32B
            apiKey = "sk-ooshywirgmrcdismctrllimnudbctvhhzybuzbqipervbrjy"
            Debug.Print "ʹ�� Local Key: ollama"
    End Select
    
    Debug.Print "=== GetCurrentAPIKey ���� ==="
    GetCurrentAPIKey = apiKey
End Function

' ���ݵ�ǰ���û�ȡ API ���ú���
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

' ����µĺ���
Public Function CallSelectedAPI(inputText As String, Optional isCodeMode As Boolean = True) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CallSelectedAPI ��ʼ ==="
    Debug.Print "�����ı�: " & inputText
    Debug.Print "����ģʽ: " & isCodeMode
    Debug.Print "��ǰģ��: " & CurrentSettings.SelectedType
    
    Dim response As String
    response = CallAPI(CurrentSettings.SelectedType, inputText, isCodeMode)
    
    Debug.Print "API ������Ӧ: " & Left(response, 1000) & "..."  ' ֻ��ӡǰ100���ַ�
    
    If Left(response, 5) <> "Error" Then
        Debug.Print "��ʼ������Ӧ..."
        CallSelectedAPI = ParseAPIResponse(response, CurrentSettings.SelectedType, isCodeMode)
    Else
        Debug.Print "API ���ó���"
        CallSelectedAPI = response
    End If
    
    Debug.Print "=== CallSelectedAPI ���� ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "CallSelectedAPI ����: " & Err.Description
    CallSelectedAPI = "Error: " & Err.Description
End Function

'/**
' * ������ͬ AI ģ�͵� API ��Ӧ����ȡ���ݲ�����˼������
' *
' * @description
' * �ú������������ AI ģ�͵� JSON ��Ӧ��������
' * 1. ������Ӧ���ݵ���ȡ
' * 2. R1 ϵ��ģ�͵�˼�����̴���
' * 3. �����ַ�������͸�ʽ��
' * 4. ���������־��¼
' *
' * @param response {String} - API ���ص�ԭʼ JSON �ַ���
' * @param modelType {APIType} - ��ǰʹ�õ�ģ�����ͣ�ö��ֵ��
' * @param isCodeMode {Boolean} - ��ѡ������Ĭ��Ϊ False
' *                              True: ���ܱ༭ģʽ��ֻ�������ս��
' *                              False: AI �Ի�ģʽ������˼������
' *
' * @return {String} ��������Ӧ����
' *                  �Ի�ģʽ�°���"˼������"��"���մ�"
' *                  ����ģʽ�½��������ս��
' *
' * @throws �����׳��Ĵ���
' *         - JSON ��������
' *         - ��Ӧ��ʽ����
' *         - �ֶ�ȱʧ����
' */
Public Function ParseAPIResponse(response As String, modelType As APIType, Optional isCodeMode As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    ' === ��ʼ������־ ===
    Debug.Print "=== ParseAPIResponse ��ʼ ==="
    Debug.Print "ģ������: " & modelType
    Debug.Print "�Ƿ����ģʽ: " & isCodeMode
    Debug.Print "��Ӧ����: " & Len(response)
    
    ' === �������� ===
    Dim jsonDict As Object
    Dim content As String
    Dim reasoningContent As String
    
    ' === JSON ��������֤ ===
    Set jsonDict = ParseJson(response)
    
    If jsonDict Is Nothing Then
        Debug.Print "JSON ����ʧ�ܣ�����Ϊ��"
        ParseAPIResponse = ""
        Exit Function
    End If
    
    ' === ����ģ�����⴦�� ===
    If IsLocalModel(modelType) Then
        If Not jsonDict.Exists("choices") Then
            Debug.Print "JSON ȱ�� choices �ֶ�"
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
            Debug.Print "Choices ����Ϊ��"
            content = "Choices ����Ϊ��"
        End If
        
        ' ����ģʽ��������
        If isCodeMode Then
            Debug.Print "������˼�����̴���"
            ' �����Ű�ģʽ����ȡ���������
            ' Dim codeStart As Long, codeEnd As Long
            
            ' ' ����˼�����֣�ֱ���Ҵ����
            ' codeStart = InStr(1, content, "```vba")
            ' If codeStart = 0 Then
            '     codeStart = InStr(1, content, "```")
            ' End If
            
            ' If codeStart > 0 Then
            '     ' �������Ա�ʶ
            '     codeStart = InStr(codeStart + 3, content, vbLf)
            '     If codeStart > 0 Then
            '         codeStart = codeStart + 1
                    
            '         ' ���ҽ������
            '         codeEnd = InStr(codeStart, content, "```")
            '         If codeEnd > 0 Then
            '             content = Mid(content, codeStart, codeEnd - codeStart)
            '         End If
            '     End If
            ' End If
        Else
            ' AI �Ի�ģʽ������˼������
            Debug.Print "����˼�����̴���"
            ' content = Replace(content, "<think>", "˼�����̣�")
            ' content = Replace(content, "</think>" & vbCrLf, vbCrLf & "���ջش�")
        End If
        
        ' ��������
        content = Replace(content, "\n", vbCrLf)  ' �����з�
        content = Trim(content)
    Else
        ' === ģ���ض���Ӧ���� ===
        Select Case modelType
            Case APIType.QwenMax
                ' ͨ��ǧ��ϵ�е���Ӧ��ʽ
                Debug.Print "����ͨ��ǧ����Ӧ"
                
                ' ��֤��Ӧ�ṹ
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON ȱ�� choices �ֶ�"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                
                ' ��ȡ������Ӧ����
                content = jsonDict("choices")(1)("message")("content")
                
                ' ͨ��ǧ�ʲ�֧��˼�����̣�ֱ�ӷ�������
                Debug.Print "ͨ��ǧ����Ӧ����: " & content
                
            Case APIType.BailianV3, APIType.BailianR1
                ' ����ϵ��ģ�ʹ���
                Debug.Print "�������ϵ����Ӧ"
                
                ' ��֤��Ӧ�ṹ
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON ȱ�� choices �ֶ�"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                
                ' ��ȡ������Ӧ����
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 �汾���⴦���ڷǴ���ģʽ�´���˼������
                If modelType = APIType.BailianR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        ' ���˼�����̺����մ�
                        content = "˼�����̣�" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "���մ𰸣�" & vbCrLf & content
                    End If
                End If
                
            Case APIType.DeepSeekV3, APIType.DeepSeekR1
                ' DeepSeek ϵ�е���Ӧ��ʽ
                Debug.Print "���� DeepSeek ��Ӧ"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON ȱ�� choices �ֶ�"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 �汾��˼�����̴���
                If modelType = APIType.DeepSeekR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "˼�����̣�" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "���մ𰸣�" & vbCrLf & content
                    End If
                End If
                
            Case APIType.SiliconFlowV3, APIType.SiliconFlowR1
                ' �������ϵ�е���Ӧ��ʽ
                Debug.Print "������������Ӧ"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON ȱ�� choices �ֶ�"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 �汾��˼�����̴���
                If modelType = APIType.SiliconFlowR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "˼�����̣�" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "���մ𰸣�" & vbCrLf & content
                    End If
                End If
                
            Case APIType.HuoshanV3, APIType.HuoshanR1
                ' ��ɽ����ϵ�е���Ӧ��ʽ
                Debug.Print "�����ɽ������Ӧ"
                If Not jsonDict.Exists("choices") Then
                    Debug.Print "JSON ȱ�� choices �ֶ�"
                    ParseAPIResponse = ""
                    Exit Function
                End If
                content = jsonDict("choices")(1)("message")("content")
                
                ' R1 �汾��˼�����̴���
                If modelType = APIType.HuoshanR1 And _
                   Not isCodeMode And jsonDict("choices")(1)("message").Exists("reasoning_content") Then
                    reasoningContent = jsonDict("choices")(1)("message")("reasoning_content")
                    If Not IsEmpty(reasoningContent) Then
                        content = "˼�����̣�" & vbCrLf & _
                                 reasoningContent & vbCrLf & vbCrLf & _
                                 "���մ𰸣�" & vbCrLf & content
                    End If
                End If
                
            Case Else
                Debug.Print "δ֪��ģ������"
                ParseAPIResponse = ""
                Exit Function
        End Select
    End If
    
    ' === �����������ʽ�� ===
    ' �Ƴ� markdown ���
    content = Replace(content, "```vba", "")
    content = Replace(content, "```", "")
    content = Trim(content)
    
    ' ���������ַ�
    content = Replace(content, "\n", vbCrLf)  ' ���з���׼��
    content = Replace(content, "\""", """")   ' ����ת�崦��
    content = Replace(content, "\\", "\")     ' ��б��ת�崦��
    
    ' === ���ؽ�� ===
    Debug.Print "���շ�������: " & content
    ParseAPIResponse = content
    Exit Function
    
    ' === ������ ===
ErrorHandler:
    Debug.Print "ParseAPIResponse ����: " & Err.Description
    Debug.Print "����λ��: " & Err.Source
    Debug.Print "�����: " & Err.Number
    Debug.Print "ԭʼ JSON: " & response
    ParseAPIResponse = ""
End Function

' �ж��Ƿ�Ϊ����ģ��
Public Function IsLocalModel(ByVal modelType As APIType) As Boolean
    Select Case modelType
        Case LocalR1_14B, LocalR1_32B
            IsLocalModel = True
        Case Else
            IsLocalModel = False
    End Select
End Function

' ��ȡģ����
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

' �� CallSelectedAPI ����ǰ���
Public Function CallAPI(ByVal modelType As APIType, ByVal inputText As String, Optional ByVal isCodeMode As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    Debug.Print "=== CallAPI ��ʼ ==="
    Debug.Print "ģ������: " & modelType
    Debug.Print "�����ı�: " & inputText
    Debug.Print "����ģʽ: " & isCodeMode
    
    Dim http As Object
    Dim url As String
    Dim requestBody As String
    Dim response As String
    
    ' ��ʼ�� HTTP ����
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' ��ȡ��ǰģ�͵� API Key
    Dim apiKey As String
    apiKey = GetCurrentAPIKey()
    
    ' ��֤ API Key
    If Not IsLocalModel(modelType) And Len(apiKey) = 0 Then
        Debug.Print "API Key Ϊ��"
        CallAPI = "Error: API Key δ����"
        Exit Function
    End If
    
    Debug.Print "ʹ�õ� API Key ����: " & Len(apiKey)
    
    ' ���ݵ�ǰѡ���ģ������ URL �͹���������
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
            ' Ĭ��ʹ�� QwenMax
            url = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
            requestBody = BuildQwenRequest(inputText, isCodeMode)
    End Select
    
    ' ��������ǰ�����ϸ��������Ϣ����
    Debug.Print String(50, "-")
    Debug.Print "����������Ϣ:"
    Debug.Print "URL: " & url
    Debug.Print "Content-Type: application/json; charset=utf-8"
    Debug.Print "Authorization: Bearer " & apiKey  ' ��������� API Key ���ڵ���
    Debug.Print "������: " & requestBody
    Debug.Print String(50, "-")
    
    ' ��������
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json; charset=utf-8"
        .setRequestHeader "Authorization", "Bearer " & apiKey
        .send requestBody
    End With
    
    ' ��Ӧ����ʱ��Ӹ�����Ϣ
    Debug.Print String(50, "-")
    Debug.Print "��Ӧ����:"
    Debug.Print "״̬��: " & http.Status
    Debug.Print "״̬�ı�: " & http.statusText
    If http.Status <> 200 Then
        Debug.Print "������Ӧ: " & http.responseText
    End If
    Debug.Print String(50, "-")
    
    If http.Status = 200 Then
        response = http.responseText
        Debug.Print "API���óɹ�����Ӧ����: " & Len(response)
        CallAPI = response
    ElseIf http.Status = 401 Then
        Debug.Print "API��֤ʧ�ܣ�" & http.Status & " - " & http.statusText
        Debug.Print "��Ӧ���ݣ�" & http.responseText
        CallAPI = "Error: API��֤ʧ�� - ���� API Key �Ƿ���ȷ"
    Else
        Debug.Print "API����ʧ�ܣ�" & http.Status & " - " & http.statusText
        Debug.Print "�����壺" & requestBody
        Debug.Print "��Ӧ���ݣ�" & http.responseText
        CallAPI = "Error: API����ʧ�� - " & http.Status & " " & http.statusText
    End If
    
    Debug.Print "=== CallAPI ���� ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "CallAPI ����: " & Err.Description
    CallAPI = "Error: " & Err.Description
End Function

' �������õ�ע���
Public Sub SaveSettingsToRegistry()
    Debug.Print "=== SaveSettingsToRegistry ��ʼ ==="
    Debug.Print "����ǰ������:"
    Debug.Print "SelectedType: " & CurrentSettings.SelectedType
    Debug.Print "DeepSeekKey: " & CurrentSettings.DeepSeekKey
    Debug.Print "BailianKey: " & CurrentSettings.BailianKey
    Debug.Print "SiliconKey: " & CurrentSettings.SiliconKey
    Debug.Print "HuoshanKey: " & CurrentSettings.HuoshanKey
    
    ' ��������
    SaveSetting "DeepSeekPPT", "Settings", "SelectedType", CStr(CurrentSettings.SelectedType)
    SaveSetting "DeepSeekPPT", "Settings", "DeepSeekKey", CurrentSettings.DeepSeekKey
    SaveSetting "DeepSeekPPT", "Settings", "BailianKey", CurrentSettings.BailianKey
    SaveSetting "DeepSeekPPT", "Settings", "SiliconKey", CurrentSettings.SiliconKey
    SaveSetting "DeepSeekPPT", "Settings", "HuoshanKey", CurrentSettings.HuoshanKey
    
    Debug.Print "=== SaveSettingsToRegistry ���� ==="
End Sub

' ��ע����������
Public Sub LoadSettingsFromRegistry()
    Debug.Print "=== LoadSettingsFromRegistry ��ʼ ==="
    
    ' ��ȡ����ǰ��״̬
    Debug.Print "��ȡǰ������:"
    Debug.Print "��ǰ����: " & CurrentSettings.SelectedType
    
    ' ��ȡ����
    CurrentSettings.SelectedType = CLng(GetSetting("DeepSeekPPT", "Settings", "SelectedType", "0"))
    CurrentSettings.DeepSeekKey = GetSetting("DeepSeekPPT", "Settings", "DeepSeekKey", "")
    CurrentSettings.BailianKey = GetSetting("DeepSeekPPT", "Settings", "BailianKey", "")
    CurrentSettings.SiliconKey = GetSetting("DeepSeekPPT", "Settings", "SiliconKey", "")
    CurrentSettings.HuoshanKey = GetSetting("DeepSeekPPT", "Settings", "HuoshanKey", "")
    
    ' ���������Ϣ
    Debug.Print "��ȡ�������:"
    Debug.Print "SelectedType: " & CurrentSettings.SelectedType
    Debug.Print "DeepSeekKey ����: " & Len(CurrentSettings.DeepSeekKey)
    If Len(CurrentSettings.DeepSeekKey) > 0 Then
        Debug.Print "DeepSeekKey ǰ5λ: " & Left(CurrentSettings.DeepSeekKey, 5)
    End If
    Debug.Print "BailianKey ����: " & Len(CurrentSettings.BailianKey)
    Debug.Print "SiliconKey ����: " & Len(CurrentSettings.SiliconKey)
    Debug.Print "HuoshanKey ����: " & Len(CurrentSettings.HuoshanKey)
    
    ' ��֤��ǰģ���Ƿ��ж�Ӧ�� API Key
    Dim currentKey As String
    currentKey = GetCurrentAPIKey()
    Debug.Print "��ǰģ�͵� API Key ����: " & Len(currentKey)
    
    Debug.Print "=== LoadSettingsFromRegistry ���� ==="
End Sub

