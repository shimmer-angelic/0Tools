' 全局变量用于存储对话历史
Dim conversationHistory As Collection

' 初始化对话历史
Sub InitializeConversation()
    Set conversationHistory = New Collection
    ' 添加系统角色的初始消息
    conversationHistory.Add Array("system", "You are a Word assistant")
End Sub

Function CallDeepSeekAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String
    Dim messages As String
    Dim i As Integer

    ' 如果对话历史为空，初始化对话历史
    If conversationHistory Is Nothing Then
        InitializeConversation
    End If

    ' 构建消息数组
    messages = "["
    For i = 1 To conversationHistory.Count
        Dim role As String
        Dim content As String
        role = conversationHistory(i)(0)
        content = conversationHistory(i)(1)
        messages = messages & "{""role"":""" & role & """,""content"":""" & content & """},"
    Next i

    ' 添加当前用户输入
    messages = messages & "{""role"":""user"",""content"":""" & inputText & """}]"

    API = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
    SendTxt = "{""model"": ""qwen-max-latest"", ""messages"": " & messages & ", ""stream"": false}"

    Set Http = CreateObject("MSXML2.XMLHTTP")
    With Http
        .Open "POST", API, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & api_key
        .send SendTxt
        status_code = .Status
        response = .responseText
    End With

    ' 弹出窗口显示 API 响应（调试用）
    ' MsgBox "API Response: " & response, vbInformation, "Debug Info"

    If status_code = 200 Then
        CallDeepSeekAPI = response
    Else
        CallDeepSeekAPI = "Error: " & status_code & " - " & response
    End If

    Set Http = Nothing
End Function

Sub DeepSeekV3()
    Dim api_key As String
    Dim inputText As String
    Dim response As String
    Dim regex As Object
    Dim matches As Object
    Dim originalSelection As Object

    api_key = "sk-7ba685cbecf44c059c9a9e1d17ef37e6"
    If api_key = "" Then
        MsgBox "Please enter the API key."
        Exit Sub
    ElseIf Selection.Type <> wdSelectionNormal Then
        MsgBox "Please select text."
        Exit Sub
    End If

    ' 保存原始选中的文本
    Set originalSelection = Selection.Range.Duplicate

    inputText = Replace(Replace(Replace(Replace(Replace(Selection.Text, "\", "\\"), vbCrLf, ""), vbCr, ""), vbLf, ""), Chr(34), "\""")
    response = CallDeepSeekAPI(api_key, inputText)

    If Left(response, 5) <> "Error" Then
        Set regex = CreateObject("VBScript.RegExp")
        With regex
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = """content"":""(.*?)"""
        End With
        Set matches = regex.Execute(response)
        If matches.Count > 0 Then
            response = matches(0).SubMatches(0)
            response = Replace(Replace(response, """", Chr(34)), """", Chr(34))

            ' 将用户的输入和模型的回复添加到对话历史
            conversationHistory.Add Array("user", inputText)
            conversationHistory.Add Array("assistant", response)

            ' 限制对话历史长度为最多 10 次对话（不包括系统消息）
            LimitConversationHistory 10

            ' 取消选中原始文本
            Selection.Collapse Direction:=wdCollapseEnd

            ' 将内容插入到选中文字的下一行
            Selection.TypeParagraph ' 插入新行
            Selection.TypeText Text:=response

            ' 将光标移回原来选中文本的末尾
            originalSelection.Select
        Else
            MsgBox "Failed to parse API response.", vbExclamation
        End If
    Else
        MsgBox response, vbCritical
    End If
End Sub

' 限制对话历史长度
Sub LimitConversationHistory(maxLength As Integer)
    Dim systemMessage As Variant
    Dim tempCollection As Collection
    Dim i As Integer

    ' 保存系统消息
    systemMessage = conversationHistory(1)

    ' 创建临时集合以存储最新的对话
    Set tempCollection = New Collection

    ' 添加系统消息
    tempCollection.Add systemMessage

    ' 从对话历史中保留最新的 maxLength 条记录（不包括系统消息）
    Dim startIndex As Integer
    startIndex = conversationHistory.Count - (maxLength * 2) + 1
    If startIndex < 2 Then startIndex = 2 ' 确保不会删除系统消息

    For i = startIndex To conversationHistory.Count
        tempCollection.Add conversationHistory(i)
    Next i

    ' 更新对话历史
    Set conversationHistory = tempCollection
End Sub
