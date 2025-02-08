Attribute VB_Name = "大模型调用"
'作者: 程序员库森
'链接：https://zhuanlan.zhihu.com/p/21641716763
'来源: 知乎
'著作权归作者所有。商业转载请联系作者获得授权，非商业转载请注明出处。

Function CallDeepSeekAPI(api_key As String, inputText As String) As String
    Dim API As String
    Dim SendTxt As String
    Dim Http As Object
    Dim status_code As Integer
    Dim response As String

    API = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions" ' "https://api.deepseek.com/chat/completions"
    SendTxt = "{""model"": ""qwen-max-latest"", ""messages"": [{""role"":""system"", _
        ""content"":""You are a Word assistant""}, {""role"":""user"", ""content"":""" & inputText & """}], ""stream"": false}"

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

    inputText = Replace(Replace(Replace(Replace(Replace(Selection.Text, "\", "\\"), vbCrLf, ""), _ 
        vbCr, ""), vbLf, ""), Chr(34), "\""")
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
