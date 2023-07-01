Attribute VB_Name = "MakeDefinition"

Sub MakeDefinition()
    Dim oRng As Range
    Dim savRng As Range
    
    Set oRng = Selection.Range
    Set savRng = Selection.Range
    If Len(oRng) = 0 Then
        MsgBox "Nothing selected", vbCritical
        GoTo lbl_Exit
    End If
    With oRng
        'avoid inadvertently selected spaces at start and end of the selection
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32)
        'make selection bold
        .Font.Bold = True
        'add parenthesis and quotation marks
        .InsertBefore Chr(40) & Chr(147)
        .InsertAfter Chr(148) & Chr(41)
        'optional remove bold emphasis from parenthesis and quotation marks
        .Characters.First.Bold = False
        .Characters.First.Next.Bold = False
        .Characters.Last.Bold = False
        .Characters.Last.Previous.Bold = False
    End With
    savRng.Select
    Selection.Collapse (wdCollapseEnd)
lbl_Exit:
    Exit Sub
End Sub
