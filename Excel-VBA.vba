Sub CallDeepSeekAPI()
    Dim question As String
    Dim response As String
    Dim url As String
    Dim apiKey As String
    Dim http As Object
    Dim content As String
    Dim json As Object
    Dim requestBody As String

    question = ThisWorkbook.Sheets(1).Range("A1").Value
    url = "https://api.siliconflow.cn/v1/chat/completions"
    apiKey = "sk-ooshywirgmrcdismctrllimnudbctvhhzybuzbqipervbrjy"  ' Consider storing securely!
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey

    ' Prepare request body
    requestBody = "{""model"":""deepseek-ai/DeepSeek-R1-Distill-Llama-70B"",""messages"":[{""role"":""user"",""content"":""" & question & """}]}"

    ' Send request
    http.send requestBody

    ' Check status
    If http.Status = 200 Then
        ' Parse JSON response
        Set json = JsonConverter.ParseJson(http.responseText)
        
        ' Extract content from response
        content = json("choices")(1)("message")("content")
        
        ' Write response content to cell
        ThisWorkbook.Sheets(1).Range("A2").Value = content
    Else
        ' Handle errors and provide more details
        ThisWorkbook.Sheets(1).Range("A2").Value = "Error: " & http.Status & " - " & http.statusText
    End If
End Sub
