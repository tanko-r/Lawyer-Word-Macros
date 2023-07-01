in file: word/vbaProject.bin - OLE stream: 'VBA/FKSDONote'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub FKSDONoteMacro()
'Puts selected text into brackets, makes uppercase and bold, and highlights in gray, MTK style.
'Doesn't work quite right.  The closing bracket ends up bold and in highlight.

Dim oRng As Range
Set oRng = Selection.Range
    
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
        .Case = wdUpperCase
        .HighlightColorIndex = wdGray25
        'add brackets
        .InsertBefore Chr(91)
        .InsertAfter Chr(93)
        'optional remove bold emphasis from parenthesis and quotation marks
        .Characters.First.Bold = False
        .Characters.Last.Bold = False
        .Characters.Last.Next.Bold = False
        .Characters.Last.HighlightColorIndex = wdNoHighlight
        .Characters.Last.Next.HighlightColorIndex = wdNoHighlight
    End With
lbl_Exit:
    Exit Sub

End Sub


-------------------------------------------------------------------------------
