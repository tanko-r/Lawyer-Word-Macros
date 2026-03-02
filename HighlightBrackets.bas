Attribute VB_Name = "HighlightBrackets"
Sub HighlightBrackets()
    'Set highlight color
    Options.DefaultHighlightColorIndex = wdYellow

    Dim searchRange As Range
    Set searchRange = ActiveDocument.content
    
    'Clear existing formatting from the find operation
    searchRange.Find.ClearFormatting
    searchRange.Find.Replacement.ClearFormatting
    
    With searchRange.Find
        .text = "\[*\]"
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True
        
        'Loop through all found instances
        Do While .Execute
            'Check if the found text starts with "signature" or "signatures" (case-insensitive)
            If LCase(searchRange.text) Like "[[]signature*" = False Then
                'Highlight the found text if the condition is not met
                searchRange.HighlightColorIndex = wdYellow
            End If
            'Collapse the range to continue searching from the end of the found text
            searchRange.Collapse wdCollapseEnd
        Loop
    End With
    
    'Clean up
    Set searchRange = Nothing
End Sub
