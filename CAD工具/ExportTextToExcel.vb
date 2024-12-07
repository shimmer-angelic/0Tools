Public Sub ExportTextToExcel()
'将文本内容及其坐标点输出到excel表中
    Dim sstext As AcadSelectionSet
    Dim textObj As AcadText
    Dim Excel As Object
    Dim ExcelWorkbook As Object
    Dim ExcelSheet As Object
    Dim dataArray() As Variant
    Dim i As Long
    
    ' 设置筛选器类型和数据
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    
    FilterType(0) = 0
    FilterData(0) = "Text"
    
    ' 创建选择集
    On Error Resume Next
    Set sstext = ThisDrawing.SelectionSets.Add("textSelection")
    If Err.Number <> 0 Then
        Set sstext = ThisDrawing.SelectionSets.Item("textSelection")
        sstext.Clear
    End If
    On Error GoTo 0
    
    ' 选择文本对象
    sstext.SelectOnScreen FilterType, FilterData
    
    If sstext.Count > 0 Then
        ' 创建Excel应用程序对象
        Set Excel = CreateObject("Excel.Application")
        Excel.Visible = True
        Set ExcelWorkbook = Excel.Workbooks.Add
        Set ExcelSheet = Excel.ActiveSheet
        
        ' 初始化数据数组
        ReDim dataArray(1 To sstext.Count, 1 To 3)
        i = 0
        
        ' 填充数据数组
        For Each textObj In sstext
            i = i + 1
            dataArray(i, 1) = Round(textObj.InsertionPoint(1), 3)
            dataArray(i, 2) = Round(textObj.InsertionPoint(0), 3)
            dataArray(i, 3) = textObj.TextString
        Next textObj
        
        ' 将数据数组一次性赋值给工作表
        ExcelSheet.Range("A1").Resize(sstext.Count, 3).Value = dataArray
        
        ' 保存Excel文件（可选）
        ' ExcelWorkbook.SaveAs "C:\Users\yongl\Desktop\OutputFile.xlsx" ' 修改为你想要保存的路径
        
        ' 关闭Excel（可选）
        ' Excel.Quit
    End If
    
    ' 确保处理所有选择集后再删除
    On Error Resume Next
    Dim ssobj As AcadSelectionSet
    For Each ssobj In ThisDrawing.SelectionSets
        ssobj.Clear
    Next ssobj
    ThisDrawing.SelectionSets.Item("textSelection").Delete
    On Error GoTo 0
    
End Sub
