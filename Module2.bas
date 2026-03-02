Attribute VB_Name = "Module2"


Sub UpdateListIndentsBasedOnSelectedLevel()
    Dim oList As List
    Dim oListLevel As ListLevel
    Dim selectedLevel As Integer
    Dim i As Integer
    
    ' Check if the selection is within a list
    If Selection.Range.ListFormat.ListType = wdListNoNumbering Then
        MsgBox "The selection is not part of a list.", vbExclamation
        Exit Sub
    End If
    
    ' Get the current list
    Set oList = Selection.Range.ListFormat.List
    
    ' Get the list level of the selected paragraph
    selectedLevel = Selection.Range.ListFormat.ListLevelNumber
    
    ' Ensure that the list has at least the selected number of levels
    If selectedLevel > oList.ListLevels.count Then
        MsgBox "The selected level does not exist in this list.", vbExclamation
        Exit Sub
    End If
    
    ' Get the ListLevel object corresponding to the selected level
    Set oListLevel = oList.ListLevels(selectedLevel)
    
    ' Now update all list levels to match the selected level's indents
    For i = 1 To oList.ListLevels.count
        ' Update each list level with the indent of the selected list level
        oList.ListLevels(i).LeftIndent = oListLevel.LeftIndent
        oList.ListLevels(i).FirstLineIndent = oListLevel.FirstLineIndent
    Next i
    
    ' Inform the user that the indents have been updated
    MsgBox "List indents have been updated based on the selected list level.", vbInformation
End Sub

