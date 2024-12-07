Public Sub addValue()
    ' This subroutine adds a real number to all selected text objects.

    Dim textObj As AcadText
    Dim selSet As AcadSelectionSet
    Dim addValue As Double

    ' Prompt user for a real number to add.
    addValue = ThisDrawing.Utility.GetReal("Enter a value to add to the selected texts: ")
    ' 创建一个新的选择集
    On Error Resume Next
    Set selSet = ThisDrawing.SelectionSets.Item("MySelectionSet")
    If Ents Is Nothing Then
        Set selSet = ThisDrawing.SelectionSets.Add("MySelectionSet")
    Else
        selSet.Clear
    End If
    On Error GoTo 0
    '在屏幕上选择需改变的数字对象
    selSet.SelectOnScreen

    ' Loop through each selected text object and update its value.
    For Each textObj In selSet
        If TypeOf textObj Is AcadText And IsNumeric(textObj.textString) Then
            Dim newValue As String
            newValue = Format(CDbl(textObj.textString) + addValue, "0.00")
            textObj.textString = newValue
            textObj.Update
        End If
    Next textObj

    ' Clean up the temporary selection set.
    selSet.Delete
End Sub