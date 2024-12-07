Sub ExportHandles_InfoToExcel()
    '在CAD中
    '将所选对的句柄值存在活动单元格的备注中，
    '将文本或数字或对象的长度值存放到活动单元格中。

    Dim AcadObj As AcadObject
    Dim HandleValue As String
    Dim HandleString As String
    Dim TextString As String
    Dim NumberString As String
    Dim LengthString As String
    Dim ExcelApp As Object
    Dim WkBook As Object
    Dim WkSheet As Object
    Dim Ents As AcadSelectionSet
    Dim Ent As AcadEntity
    
    ' 创建一个新的选择集
    On Error Resume Next
    Set Ents = ThisDrawing.SelectionSets.Item("MySelectionSet")
    If Ents Is Nothing Then
        Set Ents = ThisDrawing.SelectionSets.Add("MySelectionSet")
    Else
        Ents.Clear
    End If
    On Error GoTo 0
    
    ' 选择多个对象
    Ents.SelectOnScreen
    
    ' 初始化字符串
    HandleString = ""
    TextString = ""
    NumberString = ""
    LengthString = ""
    
    
    ' 获取每个选定对象的Handle并连接成一个字符串
    For Each Ent In Ents
        ' 拼接Handle字符串
        If HandleString = "" Then
            HandleString = Ent.Handle
        Else
            HandleString = HandleString & "|" & Ent.Handle
        End If
        
        ' 判断对象类型
        Select Case TypeName(Ent)
            Case "IAcadText", "IAcadMText"
                ' 非数字文本
                If IsNumeric(Ent.TextString) Then
                    ' 数字文本
                    If NumberString = "" Then
                        NumberString = Ent.TextString
                    Else
                        NumberString = NumberString & "+" & Ent.TextString
                    End If
                Else
                    ' 非数字文本
                    TextString = TextString & Ent.TextString
                End If
            Case "IAcadLine", "IAcadArc", "IAcadLWPolyline"
                ' 长度属性值
                If LengthString = "" Then
                    LengthString = Format(Ent.Length, "0.000")
                Else
                    LengthString = LengthString & "+" & Format(Ent.Length, "0.000")
                End If
        End Select
    Next Ent
    
    ' 连接到Excel
    On Error Resume Next
    Set ExcelApp = GetObject(, "Excel.Application")
    If ExcelApp Is Nothing Then
        Set ExcelApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    
    ' 获取活动工作簿和工作表
    Set WkBook = ExcelApp.ActiveWorkbook
    Set WkSheet = ExcelApp.ActiveSheet
    
    ' 检查并删除现有的备注
    With WkSheet.Cells(ExcelApp.ActiveCell.Row, ExcelApp.ActiveCell.Column)
        If Not .Comment Is Nothing Then
            .Comment.Delete
        End If
        ' 添加Handle字符串备注
        .AddComment
        .Comment.Text Text:=HandleString
    End With
    
    ' 将结果填写到活动单元格
    WkSheet.Cells(ExcelApp.ActiveCell.Row, ExcelApp.ActiveCell.Column).Value = TextString & LengthString & NumberString
    ' 将数字字符串和长度字符串也填写到活动单元格的备注里

    ' 显示Excel应用程序
    WkSheet.Cells(ExcelApp.ActiveCell.Row + 1, ExcelApp.ActiveCell.Column).Select
    ExcelApp.Visible = True
    
End Sub

