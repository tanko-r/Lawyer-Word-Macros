Attribute VB_Name = "TEMP_DELETE"
Sub FormatTextBetweenPercentAndDeleteMarkers()
    Dim rng As Range
    Dim rngInner As Range
    Dim strPattern As String

    ' Define the wildcard pattern to find %% followed by any characters, followed by %%
    ' The parentheses () capture the characters between the %%
    strPattern = "<u>(*)</u>"

    ' Set the range to the entire document story
    Set rng = ActiveDocument.content

    ' Set the find options
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = strPattern
        .Replacement.text = "" ' Not replacing text using replacement string, we modify in code
        .Forward = True
        .Wrap = wdFindStop ' Stop at the end of the document - safer with deletions
        .Format = False ' We are finding by text pattern, not formatting
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True ' Essential to use the wildcard pattern
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        ' Execute the find and loop through results
        Do While .Execute
            ' *** IMPORTANT: The rng object now represents the found text "%%...%%" ***

            ' Define a new range for the content *between* the %%
            ' This range starts 2 characters after the start of the found range
            ' and ends 2 characters before the end of the found range.
            ' We use ActiveDocument.Range with explicit Start and End positions
            ' because rng will automatically adjust after deletions.
            Set rngInner = ActiveDocument.Range(Start:=rng.Start + 2, End:=rng.End - 2)

            ' Apply the desired formatting to the inner range
            With rngInner.Font
                .Underline = wdUnderlineSingle ' Apply single underline
                .Color = wdColorBlue          ' Apply blue color
                ' .ColorIndex = wdBlue        ' Alternative using ColorIndex if preferred
            End With

            ' *** Now, delete the %% markers ***
            ' It's generally safest to delete from the end of the found range backwards.

            ' Delete the trailing %% (last 2 characters of the original found range)
            ' Get the original end position *before* deletion
            Dim originalEnd As Long
            originalEnd = rng.End
            ActiveDocument.Range(Start:=originalEnd - 2, End:=originalEnd).Delete

            ' Delete the leading %% (first 2 characters of the original found range)
            ' The start position of the original found range is still valid here
            ActiveDocument.Range(Start:=rng.Start, End:=rng.Start + 2).Delete

            ' *** Continue the search ***
            ' The document content has shifted due to deletions.
            ' The rng object might have automatically adjusted its bounds, but collapsing
            ' the *original* rng (which still holds the initial found bounds) to its
            ' original end position effectively moves the insertion point past
            ' where the deleted text was. This ensures the next search starts correctly.
             rng.Collapse Direction:=wdCollapseEnd

            ' Check if the collapsed range is at or near the end of the document
            ' Exit the loop if we're at the very end to prevent potential infinite loops
            ' if the last item ends at the last character before the final paragraph mark.
             If rng.End >= ActiveDocument.content.End - 1 Then Exit Do

        Loop
    End With

    MsgBox "Text between %% characters formatted with blue underline and markers deleted.", vbInformation
End Sub
