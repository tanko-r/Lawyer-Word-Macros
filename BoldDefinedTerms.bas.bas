in file: word/vbaProject.bin - OLE stream: 'VBA/BoldDefinedTerms'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Sub BoldDefinedTerms()
'
    'Replace straight quotes with curly quotes
    Options.AutoFormatAsYouTypeReplaceQuotes = True
    With ActiveDocument.Range.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = """"
        .Replacement.text = """"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With


    'Replace text in quotes with bold, limited to capitalized terms
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = True
    With Selection.Find
        .text = "[" & ChrW$(8220) & "][A-Z]*[" & ChrW$(8221) & "]"
        .Replacement.text = "^&"
        .Replacement.Font.Underline = wdUnderlineNone
        .Replacement.Font.Italic = False
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    
    'Replace bolded quotation marks with not bold
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Bold = False
    With Selection.Find
        .text = "[" & ChrW$(8220) & ChrW$(8221) & "]"
        .Replacement.text = "^&"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    Selection.Find.text = ""
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.text = ""
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.MatchCase = False
    Selection.Find.MatchWildcards = False
End Sub
-------------------------------------------------------------------------------
