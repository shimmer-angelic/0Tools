Sub SelectCADObjectsFromHandles()
    '通过单元备注栏内的图元句柄信息，筛选活动CAD窗口中的图元对象，并高亮显示
    Dim ExcelApp As Object
    Dim WkBook As Object
    Dim WkSheet As Object
    Dim HandleString As String
    Dim Handles() As String
    Dim i As Integer
    Dim AcadApp As Object
    Dim AcadDoc As Object
    Dim Entity As Object
    Dim SSet As AcadSelectionSet
    Dim ssobjs() As AcadEntity
    Dim count As Integer
    
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
    
    ' 从活动单元格的备注中获取Handle字符串
    HandleString = WkSheet.Cells(ExcelApp.ActiveCell.Row, ExcelApp.ActiveCell.Column).Comment.Text
    
    ' 分割Handle字符串
    Handles = Split(HandleString, "|")
    
    ' 连接到AutoCAD
    On Error Resume Next
    Set AcadApp = GetObject(, "AutoCAD.Application")
    If AcadApp Is Nothing Then
        Set AcadApp = CreateObject("AutoCAD.Application")
    End If
    On Error GoTo 0
    
    ' 激活AutoCAD窗口
    AcadApp.Visible = True
    Set AcadDoc = AcadApp.ActiveDocument
    
    ' 检查并删除现有的选择集
    On Error Resume Next
    Set SSet = AcadDoc.SelectionSets.Item("SS1")
    If Not SSet Is Nothing Then
        SSet.Delete
    End If
    On Error GoTo 0
    
    ' 创建新的选择集
    Set SSet = AcadDoc.SelectionSets.Add("SS1")
    
    ' 初始化数组以存储对象
    ReDim ssobjs(0 To UBound(Handles)) As AcadEntity
    count = 0
    
    ' 获取每个Handle对应的图元对象并添加到数组
    For i = LBound(Handles) To UBound(Handles)
        Set Entity = AcadDoc.HandleToObject(Handles(i))
        Set ssobjs(count) = Entity
        count = count + 1
    Next i
    
    ' 将对象数组添加到选择集中
    SSet.AddItems ssobjs
    SSet.Highlight True
End Sub
