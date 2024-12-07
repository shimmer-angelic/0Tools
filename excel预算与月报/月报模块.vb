Sub AddNewMonthReport()
    '从加载项中复制模板到新项目生成月报
    Dim templateWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim i As Integer
    Dim templateSheet As Worksheet
    Dim ws As Worksheet
    Dim tempSheet As Worksheet
    Dim replaceName As String
    ' 设置加载项工作簿（假设加载项工作簿是ThisWorkbook）
    Set templateWorkbook = ThisWorkbook
    
    ' 设置当前工作簿
    Set currentWorkbook = ActiveWorkbook
    replaceName = "[" & templateWorkbook.name & "]"
    
    ' 添加一张临时工作表以避免错误
    Set tempSheet = currentWorkbook.Worksheets.Add
    tempSheet.name = "TempSheet"
    ' 关闭警告
    Application.DisplayAlerts = False
    
    ' 删除当前工作簿中的所有原有工作表
    For Each ws In currentWorkbook.Worksheets
        If ws.name <> "TempSheet" Then
            ws.Delete
        End If
    Next ws
    
    ' 开启警告
    Application.DisplayAlerts = True

    '单独复制数据源表
    Set templateSheet = templateWorkbook.Sheets(4)
    templateSheet.Copy After:=currentWorkbook.Sheets(currentWorkbook.Sheets.count)
    
    
    ' 遍历并复制工作表  1-14表为预算用表，17-28表为月报用表
    For i = 17 To 28
        Set templateSheet = templateWorkbook.Sheets(i)
        templateSheet.Copy After:=currentWorkbook.Sheets(currentWorkbook.Sheets.count)

    Next i

    ' 删除临时工作表
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    
'    '替换公式
'    currentWorkbook.Sheets("数据源00表").Activate
'    Cells.Replace What:=replaceName & "分项对比表(02表)", Replacement:="", LookAt:= _
'    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'    ReplaceFormat:=False
'    , FormulaVersion:=xlReplaceFormula2

    For i = 1 To 13
         '替换公式中引用的模板文件名
        currentWorkbook.Sheets(i).Activate
        Cells.Replace What:=replaceName, Replacement:="", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        ', FormulaVersion:=xlReplaceFormula2
    
    Next i
    
    
    '调用窗体输入新项目信息
    newproject.Show
    FileName = templateWorkbook.name

 
End Sub

Sub 导入预算数据()
    MsgBox "正在开发中................."
    '关闭屏刷新
    '关闭表格自动计算
    'open resource file
    Dim bow As Object
    Dim new_data() As Variant
    tarLastRow = ActiveSheet.Range("D10000").End(xlUp).Row - 1
    StartTime = Timer

  '打开源文件，并需进行匹配的数据存入二维数组
    Dim pasteRange As Range
    Dim resRng As Range
    
'    FilePath = Application.GetOpenFilename("Excel Files (*.xls*),*.xls*", Title:="选择一个Excel文件")
'    If FilePath = False Or IsEmpty(FilePath) Then Exit Sub
    ' 使用文件对话框选择目标文件（要将VBA代码导入的文件）
    Dim targetFileDialog As FileDialog
    Set targetFileDialog = Application.FileDialog(msoFileDialogFilePicker)

    ' 设置文件对话框的标题和过滤器
    targetFileDialog.Title = "选择数据源文件"
    targetFileDialog.Filters.Clear
    targetFileDialog.Filters.Add "Excel 文件", "*.xls*"

    ' 显示文件对话框
    If targetFileDialog.Show = -1 Then ' 用户点击了 "打开"
        ' 获取用户选择的目标文件路径和文件名
      '  Dim FilePath As String
        Dim FileName As String
        filePath = targetFileDialog.SelectedItems(1)
 '       FileName = Mid(TargetFilePath, InStrRev(TargetFilePath, "\") + 1)
    End If
    Set wb = Workbooks.Open(FileName:=filePath, ReadOnly:=True, Notify:=False, AddToMRU:=False)
  
  
  
  '  ThisWorkbook.Worksheets("分项对比表(02表)").Activate
    resLastRow = wb.Worksheets("分项对比表(02表)").Cells(Rows.count, "D").End(xlUp).Row - 2
   ' resArr = wb.Worksheets("分项对比表(02表)").Range("D6:D" & wb.Worksheets("分项对比表(02表)").Cells(Rows.Count, "D").End(xlUp).Row - 7).Value
    Set resRng = wb.Worksheets("分项对比表(02表)").Range("A6:O" & resLastRow)
 '   resArr2 = wb.Worksheets("分项对比表(02表)").Range("C6:D" & wb.Worksheets("分项对比表(02表)").Cells(Rows.Count, "D").End(xlUp).Row - 7).Value
    
   ThisWorkbook.Activate
    '检查目标区域大小是否相同
   If resLastRow > tarLastRow Then
    'insert row
     With Rows(tarLastRow + 1 & ":" & resLastRow)
         .Insert Shift:=xlShiftDown
         .Select
     End With
   ElseIf resLastRow < tarLastRow Then
    'delete row
      Rows(resLastRow + 1 & ":" & tarLastRow).Delete Shift:=xlShiftUp
   End If


'复制公式
   
    resRng.Copy
    
    ThisWorkbook.Activate
    Range("A6").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

'分组显示
    Call 数据分组显示

'建立有效验证,
    Call 设置月报格式
    
    Call 预算汇总公式
    
    Application.DisplayAlerts = False
    wb.Close savechanges:=False
    Application.DisplayAlerts = True
    EndTime = Timer
       '计算程序运行时间
    ElapsedTime = Format((EndTime - StartTime), "#0.000") & " seconds"
    '输出程序运行时间
    MsgBox "程序运行时间：" & ElapsedTime

End Sub

Sub 设置月报格式()
    MsgBox "正在开发中................."
    '设置条件格式
    Cells.Select
    Selection.FormatConditions.Delete   '清除原有格式
    lastRow = Range("C10000").End(xlUp).Row
    Range("A6:AB" & lastRow).Select
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$C6=""单位"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$C6=""分项"""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

   
    
'
    Set rng = ActiveSheet.Range("C4:AB" & lastRow)
    If (Application.CountIf(rng, "人工") + Application.CountIf(rng, "材料") + Application.CountIf(rng, "机械")) > 0 Then
        '设置要素公式
        ActiveSheet.Range("A4:AB" & lastRow).AutoFilter Field:=3, Criteria1:=Array( _
            "人工", "机械", "材料", "分包", "其它"), Operator:=xlFilterValues
         Range("T2:AA2").Select
        Selection.Copy
        Range("T6:AA" & lastRow).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        '设置下拉列表
        For i = 6 To lastRow
            If Range("C" & i).Value = "材料" Or Range("C" & i).Value = "机械" Or Range("C" & i).Value = "人工" Or Range("C" & i).Value = "分包" Then
                fRow = i
                Exit For
            End If
        Next i
        Range("D6:D" & lastRow).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=indirect(C" & fRow & "&""下拉"" )"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
            
        '分项公式
        ActiveSheet.Range("A4:AB" & lastRow).AutoFilter Field:=3, Criteria1:=Array( _
            "分项"), Operator:=xlFilterValues
         Range("P1:AA1").Select
        Selection.Copy
        Range("P6:AA" & lastRow).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
                
        Range("J1:L1").Copy
        Range("J6:L" & lastRow).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        
        
        
        
        '设置费用类别验证
        Range("C6:C" & lastRow).Select
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="单项,单位,分项,人工,材料,机械,分包,其它"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .IMEMode = xlIMEModeNoControl
            .ShowInput = True
            .ShowError = True
        End With
        
        '恢复全选
        ActiveSheet.Range("A4:AB" & lastRow).AutoFilter Field:=3, Criteria1:=Array( _
        "人工", "机械", "材料", "分包", "其它", "单项", "单位", "分项", "="), Operator:=xlFilterValues
    End If

End Sub

Sub UpdateListObject()
'更新细目编码表,该功能用处不大
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim srcRange As Range
    Dim destTable As ListObject
    Dim uniqueValues As Collection
    Dim cell As Range
    Dim i As Integer

    Set srcSheet = ActiveWorkbook.Sheets("02综合台帐（人材机）")
    Set destSheet = ActiveWorkbook.Sheets("数据源00表")
    Set srcRange = srcSheet.Range("K7:K" & srcSheet.Cells(Rows.count, "K").End(xlUp).Row)
    Set destTable = destSheet.ListObjects("细目编码")
    
    ' 清空目的地表格
    destTable.DataBodyRange.ClearContents
    
    ' 获取唯一值
    Set uniqueValues = New Collection
    On Error Resume Next
    For Each cell In srcRange
        If cell.Value <> "" Then
            uniqueValues.Add cell.Value, CStr(cell.Value)
        End If
    Next cell
    On Error GoTo 0

    ' 调整 ListObject 的行数
    If uniqueValues.count > destTable.ListRows.count Then
        For i = destTable.ListRows.count + 1 To uniqueValues.count
            destTable.ListRows.Add
        Next i
    ElseIf uniqueValues.count < destTable.ListRows.count Then
        For i = destTable.ListRows.count To uniqueValues.count + 1 Step -1
            destTable.ListRows(i).Delete
        Next i
    End If

    ' 将唯一值添加到目的地表格
    For i = 1 To uniqueValues.count
        destTable.DataBodyRange.Cells(i, 1).Value = uniqueValues(i)
    Next i
End Sub


Sub 月报汇总公式()
' 宏由 yongl 录制，时间: 2023/05/05


actAdd = ActiveCell.Address
Dim shName As String
shName = Right(ActiveWorkbook.ActiveSheet.name, 5)
Dim formuStr As String
Dim n As Integer
Dim arrBM() As Variant
Dim arrCBGS  As Variant
Dim arrFHGS As Variant
Dim arrTBGS As Variant
   
Dim StartTime As Double
Dim EndTime As Double
Dim ElapsedTime As String
StartTime = Timer

lenBM = Len(Cells(6, "A"))
lastRow = Range("D" & Rows.count).End(xlUp).Row
Set rng = Range("A6:A" & lastRow)
If Application.CountBlank(rng) > 0 Then
    MsgBox "编码列空行，请先运行《编码分组》 后再汇总！", vbCritical
    Exit Sub
End If
'arrBM = Range("A6:A" & lastRow).Value
arrBM = Application.Transpose(Range("A6:A" & lastRow).Value)
If shName Like "*产值对比表*" Then
   
    
    arrCBGS = Application.Transpose(Range("O6:O" & lastRow).formula)
    arrTBGS = Application.Transpose(Range("I6:I" & lastRow).formula)
    arrFHGS = Application.Transpose(Range("L6:L" & lastRow).formula)
    'colNum = Array("I", "L", "O")
    
  '  For j = 0 To UBound(colNum)
        For i = 1 To UBound(arrBM) - 1
            n = Len(arrBM(i))
            If n < lenBM * 3 + 1 Then    '对单项、单位、分项进行汇总
             '   If (n < lenBM * 3 Or (n = lenBM * 3 And colNum(j) = "O")) Then '对单位、单项或分项成本合价汇总
                    K = i + 1
                    While (Left(arrBM(K), n) = arrBM(i))
                        If Len(arrBM(K)) = n + lenBM Then     '仅对下一级汇总
                            formuStr = formuStr & "+" & "O" & K + 5
                        End If
                        If K < UBound(arrBM) Then
                            K = K + 1
                        Else
                            GoTo flag1
                        End If
                    Wend
flag1:
                    If formuStr <> "" Then
                        If n < lenBM * 3 Then
                        ' 对单项、单位同时将公式复制投标合价和复核合价
                            arrCBGS(i) = "=(" & formuStr & ")"
                            arrTBGS(i) = "=(" & Replace(formuStr, "O", "I") & ")"
                            arrFHGS(i) = "=(" & Replace(formuStr, "O", "L") & ")"
                            
                        Else
                           arrCBGS(i) = "=(" & formuStr & ")"
                        End If
                        formuStr = ""
                    Else
                    '若无下级，则原单元格内容不变
                    '    arrCBGS(i) = 0
                    '   arrTBGS(i) = 0
                    '   arrFHGS(i) = 0
                    End If
         '       End If
            End If
        Next i
    Range("O6:O" & lastRow) = Application.Transpose(arrCBGS)
    Range("I6:I" & lastRow) = Application.Transpose(arrTBGS)
    Range("L6:L" & lastRow) = Application.Transpose(arrFHGS)
   ' Next j
    Cells(5, "I").formula = "=SUMPRODUCT((LEN($A6:$A" & lastRow & ")=LEN($A$6))*1,I6:I" & lastRow & ")"
    Cells(5, "L").formula = "=SUMPRODUCT((LEN($A6:$A" & lastRow & ")=LEN($A$6))*1,L6:L" & lastRow & ")"
    Cells(5, "O").formula = "=SUMPRODUCT((LEN($A6:$A" & lastRow & ")=LEN($A$6))*1,O6:O" & lastRow & ")"
    

ElseIf shName Like "*材料*" Or shName Like "*人工*" Or shName Like "*机械*" Or shName Like "*分包*" Then

    arrFHGS = Application.Transpose(Range("L6:L" & lastRow).formula)
    'colNum = Array("I", "L", "O")
    
  '  For j = 0 To UBound(colNum)
        For i = 1 To UBound(arrBM) - 1
            n = Len(arrBM(i))
            If n < lenBM * 2 + 1 Then    '对空值填充公式
                K = i + 1
                While (Left(arrBM(K), n) = arrBM(i))
                    If Len(arrBM(K)) = n + lenBM Then     '仅对下一级汇总
                        formuStr = formuStr & "+" & "L" & K + 5
                    End If
                    If K < UBound(arrBM) Then
                        K = K + 1
                    Else
                        GoTo flag2
                    End If
                Wend
flag2:
                If formuStr <> "" Then
                    
                        arrFHGS(i) = "=" & formuStr
                        
                Else
                   arrFHGS(i) = ""
                End If
                formuStr = ""
            End If
        Next i

    Range("L6:L" & lastRow) = Application.Transpose(arrFHGS)
   ' Next j

Else
    Exit Sub
End If

'取消筛选

Range("A4:P" & lastRow).AutoFilter Field:=3
ActiveWorkbook.Names.Add name:="'分项对比表(02表)'!_FilterDatabase", RefersTo:="='分项对比表(02表)'!$A$4:$P$" & lastRow, Visible:=False

'重新整理要素的合价公式
Dim arrYSMC As Variant
Dim arrYSQH As Variant
arrYSMC = Application.Transpose(Range("C6:C" & lastRow).Value)
arrYSQH = Application.Transpose(Range("O6:O" & lastRow).formula)
For i = 1 To UBound(arrYSMC)
    If InStr("人工材料机械分包", arrYSMC(i)) Then arrYSQH(i) = "=M" & i + 5 & "*" & "N" & i + 5
Next i
Range("O6:O" & lastRow).formula = Application.Transpose(arrYSQH)

'计算程序运行时间
EndTime = Timer

ElapsedTime = Format((EndTime - StartTime), "#0.000") & " seconds"
'输出程序运行时间
MsgBox "程序运行时间：" & ElapsedTime
    Range(actAdd).Select

End Sub


Sub 生成月报对比表()
shName = ActiveSheet.name
If shName Like "*人工*预算对比*" Then
        cateG = "人工"
    ElseIf shName Like "*材料*预算对比*" Then
        cateG = "材料"
    ElseIf shName Like "*机械*预算对比*" Then
        cateG = "机械"
    ElseIf shName Like "*分包*预算对比*" Then
        cateG = "分包"
    Else
        Exit Sub
End If
ActiveWorkbook.Sheets("01工程产值对比表").Activate
Call 费用类别检查

If ActiveSheet.AutoFilterMode Then
      ' 如果存在筛选，则清除筛选
      On Error Resume Next
      ActiveSheet.ShowAllData
      On Error GoTo 0
    End If
Call 通用编码
ActiveWorkbook.Sheets(shName).Activate

'fs = InputBox("请输入汇总方式(1 or 2)：" + vbCrLf + "1 整体汇总" + vbCrLf + "2 按单位工程汇总", , 1)
'If fs = "1" Then
    Call 月报对比表1
'Else
'    Call 对比表2
'End If
'设置公式
lastRow = Range("D20000").End(xlUp).Row
If lastRow < 7 Then Exit Sub
Range("I5").Select
Selection.Copy
Range("I7:I" & lastRow).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

Range("K5:S5").Select
Selection.Copy
Range("K7:S" & lastRow).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

Range("A2").Select
Selection.Copy
Range("A7:A" & lastRow).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False

colNumbers = Array("I", "k", "n", "p", "q", "r", "s")
For Each Key In colNumbers
    Range(Key & "6").formula = "=sum(" & Key & "7:" & Key & "2000)"
Next Key

'设置格式

    '设置边框
    Range("A7:V" & lastRow).Select

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Font
        .Size = 10
    End With




'设置条件格式
    Range("A7").Activate
    Cells.FormatConditions.Delete
    Selection.FormatConditions.Add Type:=xlExpression, Operator:=-1, Formula1:="=$F7=" & """" & "" & """" & "", Formula2:=""
    Selection.FormatConditions(1).Interior.ThemeColor = 8

'设置P列格式
    Range("P6:P" & lastRow).Select
    Selection.NumberFormatLocal = "0.00_);[红色](0.00)"
    
'移动说明文字
    
lastRow = Range("D20000").End(xlUp).Row

ActiveSheet.Shapes("文本框 3").Top = Range("A" & lastRow + 4).Top
ActiveSheet.Shapes("文本框 3").Left = Range("A" & lastRow + 4).Left

End Sub

Sub 月报对比表1()        '不分单位工程汇总
Dim Arr(), Crr(), Drr(), Mrr(), Nrr(), Vrr(), UnitN(), unitName(), pbclName(), pbclVol(), pbclPri(), RMrr()
Dim Drr2(), Jrr2(), Grr2(), Hrr2(), Frr2()
Dim ratioMaterial() As Variant
Set ws = ActiveWorkbook.Sheets("01工程产值对比表")
Set wsmat = ActiveWorkbook.Sheets("数据源00表")

lastRow = ws.Range("D20000").End(xlUp).Row
Arr = Application.Transpose(ws.Range("A7:A" & lastRow))
Crr = Application.Transpose(ws.Range("C7:C" & lastRow))
Drr = Application.Transpose(ws.Range("D7:D" & lastRow))
Mrr = Application.Transpose(ws.Range("M7:M" & lastRow))
Nrr = Application.Transpose(ws.Range("N7:N" & lastRow))
Vrr = Application.Transpose(ws.Range("V7:V" & lastRow))

shName = ActiveSheet.name
If shName Like "*人工*预算对比*" Then
        cateG = "人工"
    ElseIf shName Like "*材料*预算对比*" Then
        cateG = "材料"
    ElseIf shName Like "*机械*预算对比*" Then
        cateG = "机械"
    ElseIf shName Like "*分包*预算对比*" Then
        cateG = "分包"
    Else
        Exit Sub
End If

If cateG = "材料" Then
'导入配比信息
    Dim materialArr As Variant
    Dim materialCrr As Variant
    Dim rng As Range
    Set rng = Range("材料")
    materialCrr = rng.Columns(3).formula
    materialArr = rng.Columns(1).Value
    pbhl = Range("配比").Value
    pbrow = Range("配比").Row
    pbcol = Range("配比").Column
    
    '获取配比材料区域
    RMrr = wsmat.Range(Cells(pbrow - 5, pbcol), Cells(pbrow - 5, pbcol + Range("配比").Columns.count)).Value
    
    '定义配比材料数组
    列维度 = Int((UBound(RMrr, 2) - 6) / 3) '减6去掉多余列
    
    ReDim ratioMaterial(1 To 列维度, 1 To 4)
    Dim i As Long
    K = 0
    For i = 8 To UBound(RMrr, 2) Step 3 '根据配合比表修改超始列
        Key = RMrr(1, i)
        If Key <> "" Then
            K = K + 1
            ratioMaterial(K, 1) = Key
            ratioMaterial(K, 2) = CalculateColumnAverage(pbhl, i) '单价
            ratioMaterial(K, 3) = 0         '预算量
            ratioMaterial(K, 4) = 0         '累计量
        End If
    Next i

End If




If Range("D10000").End(xlUp).Row > 6 Then Rows("7:" & Range("D10000").End(xlUp).Row).Delete

'汇总计算各要素的数量
Dim 要素名称 As String
Dim 预算用量 As Double
Dim 累计用量 As Double
j = 1
K = 1
Set dict预算量 = CreateObject("Scripting.Dictionary")
Set dict累计量 = CreateObject("Scripting.Dictionary")
While (j < lastRow - 6)
    要素名称 = Drr(j)
    预算用量 = Mrr(j)
    累计用量 = Vrr(j)
    If Crr(j) = cateG And Drr(j) <> "" And Mrr(j) > 0 Then
       'result = Application.match(Drr(j), Application.Index(pbhl, 0, 1), 0)
        '判断是否为配合比材料,仅对材料表进行配比料判断
        If shName Like "*材料*" Then
            result = IsRatioMaterial(Drr(j), pbhl, materialArr, materialCrr)
        Else
            result = False
        End If
        If Not result Then
            If Not dict预算量.exists(Drr(j)) Then
            '初始化字典中的键值对
                dict预算量.Add Drr(j), Mrr(j)
                dict累计量.Add Drr(j), Vrr(j)
            Else
                '更新字典中的键值对
                dict预算量(Drr(j)) = dict预算量(Drr(j)) + Mrr(j)
                dict累计量(Drr(j)) = dict累计量(Drr(j)) + Vrr(j)
            End If
        Else
            '更新配比材料用量
            Call CalculateMaterialRatio(要素名称, 预算用量, pbhl, ratioMaterial, 累计用量)
        End If
    End If
    j = j + 1
Wend

If cateG = "材料" Then
'输出配比材料数量
    For n = 1 To UBound(ratioMaterial)
        If ratioMaterial(n, 3) > 0 Then
            ReDim Preserve Drr2(1 To K)
            ReDim Preserve Frr2(1 To K)
            ReDim Preserve Jrr2(1 To K)
            ReDim Preserve Grr2(1 To K)
            ReDim Preserve Hrr2(1 To K)
            
            Drr2(K) = ratioMaterial(n, 1)
            Frr2(K) = "t"
            Grr2(K) = ratioMaterial(n, 3) / 1000
            Hrr2(K) = ratioMaterial(n, 2) * 1000
            Jrr2(K) = ratioMaterial(n, 4) / 1000
            K = K + 1
        End If
    Next n


End If
'继续输出其它要素项目
For Each Key In dict预算量.keys
        ReDim Preserve Drr2(1 To K)
        ReDim Preserve Frr2(1 To K)
        ReDim Preserve Jrr2(1 To K)
        ReDim Preserve Grr2(1 To K)
        ReDim Preserve Hrr2(1 To K)
        
        Drr2(K) = Key
        Frr2(K) = "=IFERROR(VLOOKUP($D" & K + 6 & "," & cateG & ",2,FALSE),"""")"
        Grr2(K) = dict预算量(Key)
        Hrr2(K) = "=IFERROR(VLOOKUP($D" & K + 6 & "," & cateG & ",3,FALSE),"""")"
        Jrr2(K) = dict累计量(Key)
        K = K + 1
Next Key
On Error GoTo Flag
If Drr2(1) <> "" Then   ' Exit Sub   '如果未赋值就中止程序
    '输出
    Range("D7").Resize(K - 1, 1).Value = Application.Transpose(Drr2)
    Range("F7").Resize(K - 1, 1).formula = Application.Transpose(Frr2)
    Range("J7").Resize(K - 1, 1).Value = Application.Transpose(Jrr2)
    Range("G7").Resize(K - 1, 1).formula = Application.Transpose(Grr2)
    Range("H7").Resize(K - 1, 1).formula = Application.Transpose(Hrr2)
End If

Flag:

On Error GoTo 0
End Sub
