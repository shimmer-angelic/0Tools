Sub CalculateArea()
    '输入多个点，计算封闭区域面积
    ' 定义变量
    Dim Pline As AcadLWPolyline
    Dim obj As Object
    Dim region As AcadRegion
    Dim hatch As AcadHatch
    Dim area As Double
    Dim centroid As Variant
    Dim text As AcadText
    Dim hatchScale As Double
    Dim textHeight As Double
    
    ' 让用户输入图案填充的比例和文字高度
    userinput = ThisDrawing.Utility.GetString(False, "请输入图案填充的比例（默认值为10）: ")
    If userinput = "" Then
        hatchScale = 10#
    Else
        hatchScale = CDbl(userinput)
    End If
    
    userinput = ThisDrawing.Utility.GetString(False, "请输入文字的高度<25>: ")
    If userinput = "" Then
        textHeight = 25#
    Else
        textHeight = CDbl(userinput)
    End If
  
    '绘制多段线
    
    Set Pline = CreatePolyline()

    
    ' 检查并封闭多段线
    If Not Pline.Closed Then
        Pline.Closed = True
    End If
    
    ' 将多段线转换为区域
    Dim objArray(0) As Object
    Dim regionArray As Variant
    Set objArray(0) = Pline
    regionArray = ThisDrawing.ModelSpace.AddRegion(objArray)
    Set region = regionArray(0)
    
    Dim drawHatch As String
    Dim drawPline As String
    drawHatch = ThisDrawing.Utility.GetString(False, "是否需要绘制图案?回车不绘。 (Y/N): ")
    drawPline = ThisDrawing.Utility.GetString(False, "是否需要删除边框?回车不删除。 (Y/N): ")
    drawPline = UCase(drawPline)
    drawHatch = UCase(drawHatch)
    ' 创建图案填充并设置比例
    If drawHatch = "Y" Then
        Set hatch = ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "ANSI31", True)
        hatch.AppendOuterLoop (regionArray)
        hatch.Evaluate
        hatch.PatternScale = hatchScale
    End If
    
    If drawPline = "Y" Then
        Pline.Delete
    End If
    ' 计算面积
    area = region.area
    
    ' 计算重心（质心）
    centroid = region.centroid
    Dim centroidPoint(0 To 2) As Double
    centroidPoint(0) = centroid(0)
    centroidPoint(1) = centroid(1)
    centroidPoint(2) = 0
    
    ' 在重心处添加面积标注并设置文字高度
    Set text = ThisDrawing.ModelSpace.AddText(Format(area, "0.00"), centroidPoint, textHeight)
    text.Alignment = acAlignmentCenter
    text.TextAlignmentPoint = centroidPoint
    region.Delete
    
    
   ' MsgBox "面积计算完成，已在图形中标注。"
End Sub



Function CreatePolyline() As AcadLWPolyline
    '根据屏幕上的选择的点绘制多段线
    ' 定义变量
    Dim points() As Double
    Dim point As Variant
    Dim i As Integer
    Dim Pline As AcadLWPolyline
    
    ' 初始化点数组
    i = 0
    ReDim points(1 To 2)
    
    ' 让用户输入点
    On Error Resume Next
    Do
        point = ThisDrawing.Utility.GetPoint(, "请在屏幕上选择点 (回车键结束): ")
        If Err.Number <> 0 Then
            Exit Do
        End If
        i = i + 1
        ReDim Preserve points(1 To i * 2)
        points(i * 2 - 1) = point(0)
        points(i * 2) = point(1)
    Loop While Err.Number = 0
    On Error GoTo 0
    
    ' 检查是否至少有三个点
    If i < 3 Then
        MsgBox "至少需要三个点才能创建多段线。"
        Exit Function
    End If
    
    ' 创建多段线并赋值给 Pline 对象
    Set Pline = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
    Pline.Closed = True
    
    ' 返回多段线对象
    Set CreatePolyline = Pline
End Function


