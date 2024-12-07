#If VBA7 Then
    Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
'前面的申明 需放在在模块的最前面,如果模块中已有则不能重复
Sub ExportHandles_InfoToExcel()
    '在excel中启动程序后，切换到CAD中，将所选对的句柄值存在活动单元格的备注中，将文本或数字或对象的长度值存放到活动单元格中。
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
    Dim cadApp As Object
    Dim cadHwnd As LongPtr
    Dim excelHwnd As LongPtr
    
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    FilterType(0) = 0
    FilterData(0) = "Text"
    
    ' 启动或连接到AutoCAD应用程序
    On Error Resume Next
    Set cadApp = GetObject(, "AutoCAD.Application")
    If cadApp Is Nothing Then
        Set cadApp = CreateObject("AutoCAD.Application")
        cadApp.Visible = True
    End If
    On Error GoTo 0

    ' 获取AutoCAD窗口句柄
    cadHwnd = FindWindow(vbNullString, cadApp.Caption)

    ' 创建一个新的选择集
    ' 创建选择集
    On Error Resume Next
    Set Ents = cadApp.ActiveDocument.SelectionSets.Add("ObjSelection")
    If Err.Number <> 0 Then
        Set Ents = cadApp.ActiveDocument.SelectionSets.Item("ObjSelection")
        Ents.Clear
    End If
    On Error GoTo 0
    
    ' 切换到AutoCAD窗口
    SetForegroundWindow cadHwnd

    ' 选择多个对象
'    On Error Resume Next ' 忽略选择时的错误
    Ents.SelectOnScreen ' FilterType, FilterData
'    If Err.Number <> 0 Then
'        MsgBox "没有选择到对象，请重试。"
'        Exit Sub
'    End If
'    On Error GoTo 0

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
    
    
    ' 获取活动工作簿和工作表
    Set WkBook = ActiveWorkbook
    Set WkSheet = ActiveSheet
    
    ' 检查并删除现有的备注
    With WkSheet.Cells(ActiveCell.Row, ActiveCell.Column)
        If Not .Comment Is Nothing Then
            .Comment.Delete
        End If
        ' 添加Handle字符串备注
        .AddComment
        .Comment.Text Text:=HandleString
    End With
    
    ' 将结果填写到活动单元格
    WkSheet.Cells(ActiveCell.Row, ActiveCell.Column).Value = TextString & LengthString & NumberString
    
    ' 显示Excel应用程序
    WkSheet.Cells(ActiveCell.Row + 1, ActiveCell.Column).Select
  
    ' 切换回Excel窗口
    excelHwnd = FindWindow("XLMAIN", vbNullString) ' 获取Excel主窗口句柄
    If excelHwnd <> 0 Then
        ShowWindow excelHwnd, 5 ' 显示窗口
        SetForegroundWindow excelHwnd ' 切换到Excel窗口
    End If
End Sub

