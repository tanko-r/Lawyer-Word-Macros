Attribute VB_Name = "HilightTerm"
Sub HighlightSelectedTerm()
    Dim oRng As Range
    Dim rngCount As Range
    Dim strText As String
    Dim iCount As Long
    
    Set oRng = Selection.Range

    ' 1. Prepare the text: trim spaces from start/end
    With oRng
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32)
        strText = .text ' Store the clean text in a variable
    End With

    ' If nothing is selected (or just spaces), exit to prevent errors
    If Len(strText) = 0 Then Exit Sub

    ' 2. Count Pass: Loop through document to count matches without changing them yet
    Set rngCount = ActiveDocument.content
    iCount = 0
    
    With rngCount.Find
        .ClearFormatting
        .text = strText
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop ' Stop at end of document so we don't loop forever
        .Format = False
        
        ' Loop through matches to increment counter
        Do While .Execute = True
            iCount = iCount + 1
        Loop
    End With

    ' 3. Highlight Pass: Apply the highlights (Original Logic)
    Options.DefaultHighlightColorIndex = wdYellow

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With ActiveDocument.content.Find
        .text = strText
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

    ' 4. Show the Popup Result
    MsgBox iCount & " instances of '" & strText & "' were highlighted.", vbInformation, "Highlight Complete"

End Sub

Sub UnHighlightSelectedTerm()

    Dim oRng As Range
    Dim rngCount As Range
    Dim strText As String
    Dim iCount As Long
    
    Set oRng = Selection.Range
    
    ' 1. Prepare the text: trim spaces
    With oRng
        'avoid inadvertently selected spaces at start and end of the selection
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32)
        strText = .text
    End With
    
    If Len(strText) = 0 Then Exit Sub

    ' 2. Count Pass: Count ONLY instances that are ALREADY highlighted
    Set rngCount = ActiveDocument.content
    iCount = 0
    
    With rngCount.Find
        .ClearFormatting
        .text = strText
        .Highlight = True ' <--- This forces the find to only look for highlighted text
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = True ' Must be True for .Highlight to work
        
        Do While .Execute = True
            iCount = iCount + 1
        Loop
    End With

    ' 3. UnHighlight Pass: Apply NoHighlight ONLY to the highlighted instances
    Options.DefaultHighlightColorIndex = wdNoHighlight
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With ActiveDocument.content.Find
        .ClearFormatting
        .text = strText
        .Highlight = True ' <--- Target only the highlighted ones for replacement too
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
    
    ' Restore default color to Yellow for future use
    Options.DefaultHighlightColorIndex = wdYellow
    
    ' 4. Show the Popup Result
    MsgBox iCount & " instances of '" & strText & "' were unhighlighted.", vbInformation, "Unhighlight Complete"

End Sub
