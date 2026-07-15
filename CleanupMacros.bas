Attribute VB_Name = "CleanupMacros"
Sub ReplaceCurlyWithStraightQuotes()
    Dim originalAutoFormatQuotes As Boolean
    
    ' Save current AutoFormat settings
    originalAutoFormatQuotes = Options.AutoFormatAsYouTypeReplaceQuotes
    
    ' Turn off automatic curly quotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    
    ' Clear existing formatting in Find and Replace
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    ' Replace curly double quotes with straight double quotes
    With Selection.Find
        .text = """"
        .Replacement.text = """"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .matchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Replace curly single quotes/apostrophes with straight single quotes
    With Selection.Find
        .text = "'"
        .Replacement.text = "'"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .matchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    
    ' Restore original AutoFormat settings
    Options.AutoFormatAsYouTypeReplaceQuotes = originalAutoFormatQuotes
    
    MsgBox "Curly quotes have been replaced with straight quotes.", vbInformation, "Success"
End Sub
