Attribute VB_Name = "RCWReformatter"
Option Explicit ' Enforces variable declaration

Sub RCWReformatter()

    ' --- Configuration ---
    Const FONT_NAME As String = "Aptos"
    Const FONT_SIZE As Single = 10
    
    ' Measurement Conversion
    Const INCH_TO_POINTS As Single = 72
    
    ' Indentation and Tab Values (in Inches)
    Const HANGING_INDENT_INCHES As Single = 0.5 ' Used for hanging indent and tab offset for Levels 1-4
    
    Const LEVEL_1_INDENT_INCHES As Single = 0.5  ' Starts at the left margin
    Const LEVEL_2_INDENT_INCHES As Single = 1
    Const LEVEL_3_INDENT_INCHES As Single = 1.5
    Const LEVEL_4_INDENT_INCHES As Single = 2
    Const DEFAULT_INDENT_INCHES As Single = 1.5 ' Indent for non-level, non-"Sec." paragraphs
    Const SEC_HEADER_INDENT_INCHES As Single = 0.5 ' Indent for "Sec. " paragraphs
    
    ' --- Calculated Point Values ---
    Dim HANGING_INDENT_PTS As Single
    Dim LEVEL_1_INDENT_PTS As Single
    Dim LEVEL_2_INDENT_PTS As Single
    Dim LEVEL_3_INDENT_PTS As Single
    Dim LEVEL_4_INDENT_PTS As Single
    Dim DEFAULT_INDENT_PTS As Single
    Dim SEC_HEADER_INDENT_PTS As Single
    
    HANGING_INDENT_PTS = HANGING_INDENT_INCHES * INCH_TO_POINTS
    LEVEL_1_INDENT_PTS = LEVEL_1_INDENT_INCHES * INCH_TO_POINTS
    LEVEL_2_INDENT_PTS = LEVEL_2_INDENT_INCHES * INCH_TO_POINTS
    LEVEL_3_INDENT_PTS = LEVEL_3_INDENT_INCHES * INCH_TO_POINTS
    LEVEL_4_INDENT_PTS = LEVEL_4_INDENT_INCHES * INCH_TO_POINTS
    DEFAULT_INDENT_PTS = DEFAULT_INDENT_INCHES * INCH_TO_POINTS
    SEC_HEADER_INDENT_PTS = SEC_HEADER_INDENT_INCHES * INCH_TO_POINTS

    ' --- Regular Expressions ---
    Dim regex As Object ' For primary level check
    Dim regex2 As Object ' For secondary level check
    Dim matches As Object
    Dim match As Object
    Dim matches2 As Object ' For secondary check results
    Dim match2 As Object   ' For secondary check result
    Dim para As Paragraph
    Dim paraText As String
    Dim char As Range ' For character-level formatting (color)
    Dim levelFound As Boolean
    Dim formattingLevel As Integer ' The deepest level found (1, 2, 3, or 4)
    Dim numPart1 As String        ' The text of the first level marker, e.g., "(1)"
    Dim numPart2 As String        ' The text of the second level marker, e.g., "(a)"
    Dim textAfterNumPart1 As String
    Dim numPart1EndPos As Long
    
    ' Use Late Binding
    On Error Resume Next
    Set regex = CreateObject("VBScript.RegExp")
    Set regex2 = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        MsgBox "Error creating Regex object(s).", vbCritical, "Regex Error"
        Exit Sub
    End If
    On Error GoTo 0 ' Turn error handling off
    
    regex.Global = False
    regex.MultiLine = False
    regex2.Global = False
    regex2.MultiLine = False ' Secondary check also works on the remaining part of the line start

    ' Define Regex Patterns (Order Matters: more specific/nested first)
    ' Now includes single-char romanettes in L3
    ' Secondary patterns need ^ to match start of the *remaining* string
    Const L4_PATTERN As String = "^\s*\(?\(?(\([A-Z]\))\s*" ' Match part for finding position
    Const L3_PATTERN As String = "^\s*\(?\(?(\((i|v|x|ii|iii|iv|vi|vii|viii|ix|xi|xii|xiii|xiv|xv|xvi|xvii|xviii|xix|xx)\))\s*" ' Includes i,v,x
    Const L2_PATTERN As String = "^\s*\(?\(?(\([a-z][a-z]?\))\s*" ' Excludes those matched by L3 first
    Const L1_PATTERN As String = "^\s*\(?\(?(\([1-9][0-9]?\))\s*"
    
    ' Patterns for secondary check (must match at the VERY beginning of the remaining string)
    Const L4_PATTERN_SEC As String = "^\(?\(?(\([A-Z]\))\s*" ' No initial \s* needed for secondary check
    Const L3_PATTERN_SEC As String = "^\(?\(?(\((i|v|x|ii|iii|iv|vi|vii|viii|ix|xi|xii|xiii|xiv|xv|xvi|xvii|xviii|xix|xx)\))\s*"
    Const L2_PATTERN_SEC As String = "^\(?\(?(\([a-z][a-z]?\))\s*"

    ' --- Processing ---
    Application.ScreenUpdating = False
    Application.StatusBar = "Formatting statutory text (V3)..."

    Dim paraCount As Long
    Dim totalParas As Long
    totalParas = ActiveDocument.Paragraphs.count
    paraCount = 0
    Dim numberPartsToTab As Collection ' To hold the list of number parts needing tabs

    For Each para In ActiveDocument.Paragraphs
        paraCount = paraCount + 1
        If paraCount Mod 10 = 0 Then
             Application.StatusBar = "Formatting statutory text (V3)... Paragraph " & paraCount & " of " & totalParas
        End If

        paraText = para.Range.text
        levelFound = False
        formattingLevel = 0 ' Reset for each paragraph
        numPart1 = ""
        numPart2 = ""
        Set numberPartsToTab = New Collection ' Reset for each paragraph

        ' 1. Apply universal font settings FIRST
        With para.Range.Font
            .Name = FONT_NAME
            .Size = FONT_SIZE
        End With
        
        ' 2. Apply Color based on Underline/Strikethrough
        For Each char In para.Range.Characters
             If char.Font.StrikeThrough Then
                 char.Font.ColorIndex = wdRed
             ElseIf char.Font.Underline <> wdUnderlineNone Then
                 char.Font.ColorIndex = wdBlue
             End If
             ' Ignore characters without U/S to preserve other colors
        Next char

        ' 3. Check for Special Cases ("Sec. " header)
        If LCase(Trim(paraText)) Like "sec. *" Then
            ApplySecHeaderFormatting para, SEC_HEADER_INDENT_PTS
            levelFound = True ' Consider it "handled"
            
        Else ' 4. Check for Numbered/Lettered Levels (potentially two levels)
            
            ' Check Level 4 first
            regex.pattern = L4_PATTERN
            Set matches = regex.Execute(paraText)
            If matches.count > 0 Then
                levelFound = True
                Set match = matches(0)
                numPart1 = match.SubMatches(0) ' e.g., "(A)"
                numPart1EndPos = match.FirstIndex + match.Length ' Position after the matched L4 pattern (incl. trailing space)
                formattingLevel = 4
                
                ' Check for subsequent Level 3 or 2 immediately after
                If numPart1EndPos < Len(paraText) Then
                    textAfterNumPart1 = Mid(paraText, numPart1EndPos + 1)
                    ' Check L3 second
                    regex2.pattern = L3_PATTERN_SEC
                    Set matches2 = regex2.Execute(textAfterNumPart1)
                    If matches2.count > 0 Then
                        Set match2 = matches2(0)
                        numPart2 = match2.SubMatches(0)
                        formattingLevel = 3 ' Deeper level dictates formatting
                    Else
                         ' Check L2 second
                         regex2.pattern = L2_PATTERN_SEC
                         Set matches2 = regex2.Execute(textAfterNumPart1)
                         If matches2.count > 0 Then
                             Set match2 = matches2(0)
                             numPart2 = match2.SubMatches(0)
                             formattingLevel = 2 ' Deeper level dictates formatting
                         End If
                    End If
                End If
            End If

            ' Check Level 3 (only if L4 not found)
            If Not levelFound Then
                regex.pattern = L3_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    numPart1 = match.SubMatches(0) ' e.g., "(i)" or "(ix)"
                    numPart1EndPos = match.FirstIndex + match.Length
                    formattingLevel = 3
                    
                    ' Check for subsequent Level 2 immediately after
                    If numPart1EndPos < Len(paraText) Then
                        textAfterNumPart1 = Mid(paraText, numPart1EndPos + 1)
                        regex2.pattern = L2_PATTERN_SEC
                        Set matches2 = regex2.Execute(textAfterNumPart1)
                        If matches2.count > 0 Then
                            Set match2 = matches2(0)
                            numPart2 = match2.SubMatches(0)
                            formattingLevel = 2 ' Deeper level dictates formatting
                        End If
                    End If
                End If
            End If
            
            ' Check Level 2 (only if L4/L3 not found)
            If Not levelFound Then
                regex.pattern = L2_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    numPart1 = match.SubMatches(0) ' e.g., "(a)" or "(bb)"
                    numPart1EndPos = match.FirstIndex + match.Length
                    formattingLevel = 2
                    ' No deeper levels possible after L2 in this schema (L4/L3 checked first)
                End If
            End If

            ' Check Level 1 (only if L4/L3/L2 not found)
            If Not levelFound Then
                regex.pattern = L1_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    numPart1 = match.SubMatches(0) ' e.g., "(1)" or "(12)"
                    numPart1EndPos = match.FirstIndex + match.Length
                    formattingLevel = 1
                    
                    ' Check for subsequent L4, L3, or L2 immediately after
                    If numPart1EndPos < Len(paraText) Then
                        textAfterNumPart1 = Mid(paraText, numPart1EndPos + 1)
                         ' Check L4 second
                        regex2.pattern = L4_PATTERN_SEC
                        Set matches2 = regex2.Execute(textAfterNumPart1)
                        If matches2.count > 0 Then
                             Set match2 = matches2(0)
                             numPart2 = match2.SubMatches(0)
                             formattingLevel = 4 ' Deeper level dictates formatting
                        Else
                            ' Check L3 second
                            regex2.pattern = L3_PATTERN_SEC
                            Set matches2 = regex2.Execute(textAfterNumPart1)
                             If matches2.count > 0 Then
                                 Set match2 = matches2(0)
                                 numPart2 = match2.SubMatches(0)
                                 formattingLevel = 3 ' Deeper level dictates formatting
                             Else
                                 ' Check L2 second
                                 regex2.pattern = L2_PATTERN_SEC
                                 Set matches2 = regex2.Execute(textAfterNumPart1)
                                 If matches2.count > 0 Then
                                      Set match2 = matches2(0)
                                      numPart2 = match2.SubMatches(0)
                                      formattingLevel = 2 ' Deeper level dictates formatting
                                 End If
                             End If
                        End If
                    End If
                End If
            End If
            
            ' 5. Apply Formatting based on findings
            If levelFound And formattingLevel > 0 Then
                 ' Add found number parts to the collection for tabbing
                 If numPart1 <> "" Then numberPartsToTab.Add numPart1
                 If numPart2 <> "" Then numberPartsToTab.Add numPart2

                 ' Apply indentation based on the deepest level found
                 Select Case formattingLevel
                     Case 4: ApplyLevelFormatting para, LEVEL_4_INDENT_PTS, HANGING_INDENT_PTS
                     Case 3: ApplyLevelFormatting para, LEVEL_3_INDENT_PTS, HANGING_INDENT_PTS
                     Case 2: ApplyLevelFormatting para, LEVEL_2_INDENT_PTS, HANGING_INDENT_PTS
                     Case 1: ApplyLevelFormatting para, LEVEL_1_INDENT_PTS, HANGING_INDENT_PTS
                 End Select
                 
                 ' Ensure tabs after each identified number part
                 If numberPartsToTab.count > 0 Then
                      EnsureTabsAfterMultipleNumbers para, numberPartsToTab
                 End If
                 
            ElseIf Not levelFound Then
                 ' Apply Default Formatting if not "Sec." and no level found
                 ApplyDefaultFormatting para, DEFAULT_INDENT_PTS
            End If
            
        End If ' End If for "Sec. *" vs Level/Default Check

        ' 6. Apply Universal Spacing Settings AFTER specific indentation/alignment
        With para.Format
             .LineSpacingRule = wdLineSpaceSingle ' Set line spacing to 1.0
             .spaceAfter = 6 ' Set spacing after paragraph to 6 points
        End With

    Next para

    ' --- Cleanup ---
    Set regex = Nothing
    Set regex2 = Nothing
    Set matches = Nothing
    Set match = Nothing
    Set matches2 = Nothing
    Set match2 = Nothing
    Set para = Nothing
    Set char = Nothing
    Set numberPartsToTab = Nothing ' Clear collection
    Application.ScreenUpdating = True
    Application.StatusBar = False ' Clear status bar
    MsgBox "Document formatting complete (V3). Processed " & totalParas & " paragraphs.", vbInformation, "Formatting Finished"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplyLevelFormatting
' Purpose   : Applies indentation, hanging indent, alignment, and tab stop for numbered/lettered levels.
'---------------------------------------------------------------------------------------
Private Sub ApplyLevelFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single, ByVal hangingIndentPts As Single)
    On Error Resume Next
    With para.Format
        .LeftIndent = leftIndentPts
        .FirstLineIndent = -hangingIndentPts
        .Alignment = wdAlignParagraphLeft
        .SpaceBefore = 0 ' Reset
        .TabStops.ClearAll
        .TabStops.Add Position:=leftIndentPts + hangingIndentPts, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplyDefaultFormatting
' Purpose   : Applies indentation and justification for paragraphs not matching a level
'             pattern and not starting with "Sec. ".
'---------------------------------------------------------------------------------------
Private Sub ApplyDefaultFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single)
    On Error Resume Next
    If Len(Trim(Replace(Replace(para.Range.text, vbCr, ""), vbTab, ""))) > 0 Then
        With para.Format
            .LeftIndent = leftIndentPts
            .FirstLineIndent = 0
            .Alignment = wdAlignParagraphJustify
            .SpaceBefore = 0 ' Reset
            .TabStops.ClearAll
        End With
    Else ' Blank paragraphs
         With para.Format
            .LeftIndent = 0
            .FirstLineIndent = 0
            .Alignment = wdAlignParagraphLeft
            .SpaceBefore = 0 ' Reset
            .TabStops.ClearAll
         End With
    End If
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplySecHeaderFormatting
' Purpose   : Applies specific formatting for paragraphs starting with "Sec. ".
'---------------------------------------------------------------------------------------
Private Sub ApplySecHeaderFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single)
     On Error Resume Next
     With para.Format
         .LeftIndent = leftIndentPts
         .FirstLineIndent = 0
         .Alignment = wdAlignParagraphLeft
         .SpaceBefore = 0 ' Reset
         .TabStops.ClearAll
     End With
     On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : EnsureTabsAfterMultipleNumbers
' Author    : Gemini
' Date      : 2025-04-26
' Purpose   : Finds each number part specified in the collection sequentially within
'             the paragraph and ensures exactly one tab character follows it.
' Arguments : para - The Paragraph object to modify.
'             numberParts - A Collection object containing the strings of the number parts
'                           in the order they appear (e.g., "(1)", "(a)").
'---------------------------------------------------------------------------------------
Private Sub EnsureTabsAfterMultipleNumbers(ByVal para As Paragraph, ByVal numberParts As Collection)
    Dim paraRange As Range
    Dim findRange As Range
    Dim part As Variant
    Dim currentSearchStart As Long
    Dim found As Boolean
    Dim charAfterRange As Range
    Dim posAfterPart As Long

    If numberParts Is Nothing Or numberParts.count = 0 Then Exit Sub ' Nothing to do

    On Error GoTo ErrorHandler

    Set paraRange = para.Range
    currentSearchStart = paraRange.Start ' Start search from the beginning of the paragraph

    For Each part In numberParts
        Set findRange = ActiveDocument.Range(Start:=currentSearchStart, End:=paraRange.End)

        With findRange.Find
            .ClearFormatting
            .text = part ' The specific string like "(1)" or "(a)"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = True ' Important for (a) vs (A)
            .MatchWholeWord = False ' Number part might be adjacent to text
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            found = .Execute()
        End With

        If found Then
            ' findRange now corresponds to the location of 'part'
            posAfterPart = findRange.End ' Position immediately after the found part

            ' Check character immediately after the part
            If posAfterPart < paraRange.End - 1 Then ' Ensure not at the very end before para mark
                Set charAfterRange = ActiveDocument.Range(Start:=posAfterPart, End:=posAfterPart + 1)

                If charAfterRange.text = " " Then
                    charAfterRange.Delete ' Delete the space
                    ' Re-check character at the same position (now the next original char or tab)
                    Set charAfterRange = ActiveDocument.Range(Start:=posAfterPart, End:=posAfterPart + 1)
                End If
                
                ' If it's not a tab, insert one
                If charAfterRange.text <> vbTab Then
                     ActiveDocument.Range(Start:=posAfterPart, End:=posAfterPart).InsertBefore vbTab
                     ' Update position for next search start to account for inserted tab
                     currentSearchStart = posAfterPart + 1
                Else
                     ' It is a tab, update position for next search start
                     currentSearchStart = posAfterPart + 1
                End If
                
            ElseIf posAfterPart = paraRange.End - 1 Then ' Part is last thing before para mark
                 ' Insert tab before the paragraph mark
                 ActiveDocument.Range(Start:=posAfterPart, End:=posAfterPart).InsertBefore vbTab
                 currentSearchStart = posAfterPart + 1 ' Update position
            Else
                ' Part is at the very end? Unlikely. Just advance search position.
                currentSearchStart = posAfterPart
            End If
            
        Else
            ' Part not found sequentially, log warning and stop processing tabs for this para
            Debug.Print "Warning: Could not find sequential part '" & part & "' in paragraph: " & Left(para.Range.text, 70)
            Exit For ' Stop trying to find subsequent parts in this paragraph
        End If

        Set findRange = Nothing ' Reset range for next loop iteration
        Set charAfterRange = Nothing
    Next part

ErrorHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in EnsureTabsAfterMultipleNumbers for paragraph '" & Left(para.Range.text, 70) & "': " & Err.Description
        Err.Clear
    End If
    Set paraRange = Nothing
    Set findRange = Nothing
    Set charAfterRange = Nothing
End Sub
