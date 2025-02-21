Sub CallDeepSeekAPI()
    Dim question As String
    Dim responseData As String
    Dim url As String
    Dim apiKey As String
    Dim http As Object
    Dim content As String
    Dim json As Object
    Dim requestBody As String
    Dim logSheet As Worksheet

    ' 获取问题
    question = ThisWorkbook.Sheets(1).Range("A1").Value
    url = "https://api.siliconflow.cn/v1/chat/completions"
    apiKey = "sk-ooshywirgmrcdismctrllimnudbctvhhzybuzbqipervbrjy"  ' API Key
    
    ' 获取日志工作表，如果没有就创建一个
    On Error Resume Next
    Set logSheet = ThisWorkbook.Sheets("Log")
    On Error GoTo 0
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        logSheet.Name = "Log"
        logSheet.Cells(1, 1).Value = "Time"
        logSheet.Cells(1, 2).Value = "Message"
        logSheet.Cells(1, 3).Value = "Status"
    End If

    ' 记录日志：发送请求前
    Call LogToSheet(logSheet, "Starting API request for question: " & question)

    ' 创建 HTTP 对象
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey

    ' 准备请求体
    requestBody = "{""model"":""deepseek-ai/DeepSeek-R1-Distill-Llama-70B"",""messages"":[{""role"":""user"",""content"":""" & question & """}]}"

    ' 发送请求
    http.send requestBody

    ' 记录日志：请求完成
    Call LogToSheet(logSheet, "API request sent. Waiting for response...")

    ' 检查状态
    If http.Status = 200 Then
        responseData = http.responseText
        ' 记录日志：响应成功
        Call LogToSheet(logSheet, "Response received: Success (Status 200)")
        Call LogToSheet(logSheet, "Response content: " & responseData)

                
        ' 打印返回的响应，查看是否为有效的 JSON
        Debug.Print responseData  ' 查看响应内容

        ' 解析 JSON 响应 需要开启Microsoft Script Control 1.0
        ' On Error GoTo JsonParseError
        ' Set json = scriptControl.Eval("JSON.parse('" & responseData & "')")
        ' Set json = JsonConverter.ParseJson(responseData)
        ' On Error GoTo 0
        
        startPos = InStr(responseData, """content"":""") + Len("""content"":""")
        endPos = InStr(startPos, responseData, """")
        content = Mid(responseData, startPos, endPos - startPos)
        
        ' 提取内容
        ' content = json("choices")(1)("message")("content")
        
        ' 将内容写入 A2 单元格
        ThisWorkbook.Sheets(1).Range("A2").Value = content

        ' 记录日志：成功完成
        Call LogToSheet(logSheet, "Content extracted successfully.")
    Else
        ' 记录日志：错误
        Call LogToSheet(logSheet, "Error: " & http.Status & " - " & http.statusText)

        ' 处理错误
        ThisWorkbook.Sheets(1).Range("A2").Value = "Error: " & http.Status & " - " & http.statusText
    End If
    Exit Sub

JsonParseError:
    ' 处理 JSON 解析错误
    Call LogToSheet(logSheet, "JSON parsing error: " & Err.Description)
    ThisWorkbook.Sheets(1).Range("A2").Value = "Error parsing JSON"
End Sub


' 日志写入函数
Sub LogToSheet(logSheet As Worksheet, message As String)
    Dim lastRow As Long
    lastRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    logSheet.Cells(lastRow, 1).Value = Now
    logSheet.Cells(lastRow, 2).Value = message
    logSheet.Cells(lastRow, 3).Value = "Logged"
End Sub

