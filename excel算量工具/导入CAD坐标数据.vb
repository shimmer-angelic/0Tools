#If VBA7 Then
    Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long

'前面的申明 需放在在模块的最前面
Public Sub ExportTextToExcel()
    '在excel中启动程序后，切换到CAD中，将选择到的文本内容及其坐标点输出到excel表中
    Dim cadApp As Object
    Dim sstext As Object
    Dim textObj As Object
    Dim ExcelWorkbook As Object
    Dim ExcelSheet As Object
    Dim dataArray() As Variant
    Dim i As Long
    Dim cadHwnd As LongPtr
    
    ' 启动或连接到AutoCAD应用程序
    On Error Resume Next
    Set cadApp = GetObject(, "AutoCAD.Application")
    If Err.Number <> 0 Then
        Set cadApp = CreateObject("AutoCAD.Application")
        cadApp.Visible = True
    End If
    On Error GoTo 0

    ' 获取AutoCAD窗口句柄
    cadHwnd = FindWindow(vbNullString, cadApp.Caption)
    
    ' 设置筛选器类型和数据
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    
    FilterType(0) = 0
    FilterData(0) = "Text"
    
    ' 创建选择集
    On Error Resume Next
    Set sstext = cadApp.ActiveDocument.SelectionSets.Add("textSelection")
    If Err.Number <> 0 Then
        Set sstext = cadApp.ActiveDocument.SelectionSets.Item("textSelection")
        sstext.Clear
    End If
    On Error GoTo 0
    
    ' 提示用户进行选择
    MsgBox "请在AutoCAD中选择文本对象，然后按确定继续。"
    
    ' 切换到AutoCAD窗口
    SetForegroundWindow cadHwnd
    Dim insertionPoint As Variant

    
    ' 选择文本对象
    sstext.SelectOnScreen FilterType, FilterData
    
    If sstext.Count > 0 Then
        ' 使用当前的Excel应用程序
        Set ExcelWorkbook = ThisWorkbook
        Set ExcelSheet = ExcelWorkbook.ActiveSheet
        
        ' 初始化数据数组
        ReDim dataArray(1 To sstext.Count, 1 To 3)
        i = 0
        
        ' 填充数据数组
        For Each textObj In sstext
            i = i + 1

            insertionPoint = textObj.insertionPoint
            dataArray(i, 1) = Round(insertionPoint(0), 3)
            dataArray(i, 2) = Round(insertionPoint(1), 3)
            dataArray(i, 3) = textObj.TextString
        Next textObj
        
        ' 将数据数组一次性赋值给工作表
        ExcelSheet.Range(ActiveCell.Address).Resize(sstext.Count, 3).Value = dataArray
    End If
    
    ' 确保处理所有选择集后再删除
    On Error Resume Next
    Dim ssobj As Object
    For Each ssobj In cadApp.ActiveDocument.SelectionSets
        ssobj.Clear
    Next ssobj
    cadApp.ActiveDocument.SelectionSets.Item("textSelection").Delete
    On Error GoTo 0
    
    ' 切换回Excel窗口
    Dim excelHwnd As LongPtr
    excelHwnd = FindWindow("XLMAIN", vbNullString) ' 获取Excel主窗口句柄
    If excelHwnd <> 0 Then
        ShowWindow excelHwnd, 5 ' 5 = SW_SHOW
        SetForegroundWindow excelHwnd ' 切换到Excel窗口
    End If

End Sub


