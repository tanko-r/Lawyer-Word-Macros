Attribute VB_Name = "HilightTerm"

Sub HighlightSelectedTerm()

Dim oRng As Range
Set oRng = Selection.Range

With oRng
    'avoid inadvertently selected spaces at start and end of the selection
    .MoveEndWhile Chr(32), wdBackward
    .MoveStartWhile Chr(32)
End With

Options.DefaultHighlightColorIndex = wdYellow

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.MatchCase = False
    With ActiveDocument.content.Find
        .text = oRng
        .Replacement.text = "^&"
        .Replacement.Highlight = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With


End Sub


Sub UnHighlightSelectedTerm()

Dim oRng As Range
Set oRng = Selection.Range

With oRng
    'avoid inadvertently selected spaces at start and end of the selection
    .MoveEndWhile Chr(32), wdBackward
    .MoveStartWhile Chr(32)
End With

Options.DefaultHighlightColorIndex = wdNoHighlight

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.MatchCase = False
    With ActiveDocument.content.Find
        .text = oRng
        .Replacement.text = "^&"
        .Replacement.Highlight = True
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With

Options.DefaultHighlightColorIndex = wdYellow

End Sub
