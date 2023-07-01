in file: word/vbaProject.bin - OLE stream: 'VBA/HighlightBrackets'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


Sub HighlightBrackets()

Set searchRange = ActiveDocument.content

'Set highlight color
Options.DefaultHighlightColorIndex = wdYellow

'Replace text in quotes with bold
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.Highlight = True
    With ActiveDocument.content.Find
.text = "\[*\]"
.Replacement.text = "^&"
.Forward = True
.Wrap = wdFindContinue
        .Format = True
        .Replacement.Highlight = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.MatchCase = False
    Selection.Find.MatchWildcards = False
    
End Sub
-------------------------------------------------------------------------------
