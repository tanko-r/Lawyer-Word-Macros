Attribute VB_Name = "BillReformatter"
Option Explicit ' Enforces variable declaration

Sub FormatStatutoryText_V2()

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
    Dim regex As Object ' Requires reference to 'Microsoft VBScript Regular Expressions 5.5' or uses late binding
    Dim matches As Object ' VBScript_RegExp_55.MatchCollection
    Dim match As Object   ' VBScript_RegExp_55.Match
    Dim para As Paragraph
    Dim paraText As String
    Dim matchedNumberPart As String
    Dim levelFound As Boolean
    Dim char As Range ' For character-level formatting (color)
    
    ' Use Late Binding for broader compatibility (avoids needing to set a reference)
    On Error Resume Next ' Basic error handling for CreateObject
    Set regex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        MsgBox "Error creating Regex object." & vbCrLf & _
               "Please ensure VBScript runtime is available and enabled.", vbCritical, "Regex Error"
        Exit Sub
    End If
    On Error GoTo 0 ' Turn error handling off
    
    regex.Global = False ' We only care about the beginning of the paragraph
    regex.MultiLine = False
    
    ' Define Regex Patterns (Order Matters: more specific/nested first)
    ' Pattern breakdown:
    ' ^\s* : Start of the line, followed by zero or more whitespace characters
    ' \(?\(?    : Optional one or two literal opening parentheses (handles ((...) )
    ' (...)     : Capturing group 1: The actual number/letter part like (A), (i), (a), (1)
    '   \(      : Literal opening parenthesis
    '   [...]   : Character class for the specific level (A-Z, roman, a-z, digits)
    '   \)      : Literal closing parenthesis
    ' \s* : Zero or more whitespace characters after the number part
    ' (.*)      : Capturing group 2: The rest of the paragraph text (not strictly needed for formatting but good practice)
    Const L4_PATTERN As String = "^\s*\(?\(?(\([A-Z]\))\s*(.*)" ' Level 4: (A), (B), ...
    Const L3_PATTERN As String = "^\s*\(?\(?(\((i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii|xiii|xiv|xv|xvi|xvii|xviii|xix|xx)\))\s*(.*)" ' Level 3: (i) to (xx)
    Const L2_PATTERN As String = "^\s*\(?\(?(\([a-z][a-z]?\))\s*(.*)" ' Level 2: (a), (b), ..., (z), (aa), (bb), ...
    Const L1_PATTERN As String = "^\s*\(?\(?(\([1-9][0-9]?\))\s*(.*)" ' Level 1: (1) to (99)

    ' --- Processing ---
    Application.ScreenUpdating = False ' Speed up macro execution
    Application.StatusBar = "Formatting statutory text (V2)..." ' Progress indicator

    Dim paraCount As Long
    Dim totalParas As Long
    totalParas = ActiveDocument.Paragraphs.count
    paraCount = 0

    For Each para In ActiveDocument.Paragraphs
        paraCount = paraCount + 1
        If paraCount Mod 10 = 0 Then ' Update status bar periodically
             Application.StatusBar = "Formatting statutory text (V2)... Paragraph " & paraCount & " of " & totalParas
        End If

        paraText = para.Range.text
        levelFound = False
        matchedNumberPart = "" ' Reset for each paragraph

        ' 1. Apply universal font settings FIRST
        With para.Range.Font
            .Name = FONT_NAME
            .Size = FONT_SIZE
            ' Color is handled below based on underline/strikethrough
        End With
        
        ' 2. Apply Color based on Underline/Strikethrough (preserves the U/S)
        ' Iterate character by character to apply color without affecting other formatting
        For Each char In para.Range.Characters
            If char.Font.StrikeThrough Then
                char.Font.ColorIndex = wdRed
            ElseIf char.Font.Underline <> wdUnderlineNone Then ' Check for any type of underline
                char.Font.ColorIndex = wdBlue
            Else
                ' Optional: Reset color for characters without U/S if needed,
                ' but usually desired to keep existing colors unless specified.
                ' char.Font.ColorIndex = wdAuto ' Uncomment to force non-U/S to default color
            End If
        Next char

        ' 3. Check for Special Cases ("Sec. " or "New Section" header)
        If LCase(Trim(paraText)) Like "sec. *" Then
            ApplySecHeaderFormatting para, SEC_HEADER_INDENT_PTS
            ' Skip level checks for Sec headers
        ElseIf LCase(Trim(paraText)) = "NEW SECTION*" Then
            ApplySecHeaderFormatting para, SEC_HEADER_INDENT_PTS
            ' Skip level checks for Sec headers
            
        Else ' 4. Check for Numbered/Lettered Levels
            
            ' Check Level 4: (A), (B), etc.
            regex.pattern = L4_PATTERN
            Set matches = regex.Execute(paraText)
            If matches.count > 0 Then
                levelFound = True
                Set match = matches(0) ' Get the first match
                matchedNumberPart = match.SubMatches(0) ' Get the captured group "(A)"
                ApplyLevelFormatting para, LEVEL_4_INDENT_PTS, HANGING_INDENT_PTS
                EnsureTabAfterNumber para, matchedNumberPart
            End If

            ' Check Level 3: (i), (ii), etc.
            If Not levelFound Then
                regex.pattern = L3_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    matchedNumberPart = match.SubMatches(0) ' Get the captured group "(i)" etc.
                    ApplyLevelFormatting para, LEVEL_3_INDENT_PTS, HANGING_INDENT_PTS
                    EnsureTabAfterNumber para, matchedNumberPart
                End If
            End If
            
            ' Check Level 2: (a), (aa), etc.
            If Not levelFound Then
                regex.pattern = L2_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    matchedNumberPart = match.SubMatches(0) ' Get the captured group "(a)" etc.
                    ApplyLevelFormatting para, LEVEL_2_INDENT_PTS, HANGING_INDENT_PTS
                    EnsureTabAfterNumber para, matchedNumberPart
                End If
            End If

            ' Check Level 1: (1), (10), etc.
            If Not levelFound Then
                regex.pattern = L1_PATTERN
                Set matches = regex.Execute(paraText)
                If matches.count > 0 Then
                    levelFound = True
                    Set match = matches(0)
                    matchedNumberPart = match.SubMatches(0) ' Get the captured group "(1)" etc.
                    ApplyLevelFormatting para, LEVEL_1_INDENT_PTS, HANGING_INDENT_PTS
                    EnsureTabAfterNumber para, matchedNumberPart
                End If
            End If
            
            ' Apply Default Formatting if not "Sec." and no level was found
            If Not levelFound Then
                ApplyDefaultFormatting para, DEFAULT_INDENT_PTS
            End If
            
        End If ' End If for "Sec. *" vs Level/Default Check

        ' 5. Apply Universal Spacing Settings AFTER specific indentation/alignment
        With para.Format
             .LineSpacingRule = wdLineSpaceSingle ' Set line spacing to 1.0
             .spaceAfter = 6 ' Set spacing after paragraph to 6 points
        End With

    Next para

    ' --- Cleanup ---
    Set regex = Nothing
    Set matches = Nothing
    Set match = Nothing
    Set para = Nothing
    Set char = Nothing
    Application.ScreenUpdating = True
    Application.StatusBar = False ' Clear status bar
    MsgBox "Document formatting complete (V2). Processed " & totalParas & " paragraphs.", vbInformation, "Formatting Finished"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplyLevelFormatting
' Author    : Gemini/Revised
' Date      : 2025-04-25
' Purpose   : Applies indentation, hanging indent, alignment, and tab stop for numbered/lettered levels.
'             NOTE: Does NOT set spacing - universal spacing applied after this.
' Arguments : para - The Paragraph object to format.
'             leftIndentPts - The left indentation in points.
'             hangingIndentPts - The hanging indentation in points (also used for tab offset).
'---------------------------------------------------------------------------------------
Private Sub ApplyLevelFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single, ByVal hangingIndentPts As Single)
    On Error Resume Next ' Prevent errors on unusual paragraph objects
    With para.Format
        .LeftIndent = leftIndentPts
        .FirstLineIndent = -hangingIndentPts ' Negative value creates a hanging indent
        .Alignment = wdAlignParagraphLeft    ' Levels are typically left-aligned
        ' Reset spacing here in case paragraph had unusual settings; universal setting applied later
        .SpaceBefore = 0
        .TabStops.ClearAll
        ' Add a single left-aligned tab stop positioned relative to the page margin
        ' The position is where the text after the number should start.
        .TabStops.Add Position:=leftIndentPts + hangingIndentPts, Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplyDefaultFormatting
' Author    : Gemini/Revised
' Date      : 2025-04-25
' Purpose   : Applies indentation and justification for paragraphs not matching a level
'             pattern and not starting with "Sec. ".
'             NOTE: Does NOT set spacing - universal spacing applied after this.
' Arguments : para - The Paragraph object to format.
'             leftIndentPts - The left indentation in points (typically DEFAULT_INDENT_PTS).
'---------------------------------------------------------------------------------------
Private Sub ApplyDefaultFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single)
    On Error Resume Next ' Prevent errors on unusual paragraph objects
    ' Check if paragraph is essentially empty (just paragraph mark, maybe spaces/tabs)
    If Len(Trim(Replace(Replace(para.Range.text, vbCr, ""), vbTab, ""))) > 0 Then
        With para.Format
            .LeftIndent = leftIndentPts
            .FirstLineIndent = 0 ' No hanging indent for default text
            .Alignment = wdAlignParagraphJustify ' Justify default text as requested
             ' Reset spacing here in case paragraph had unusual settings; universal setting applied later
            .SpaceBefore = 0
            .TabStops.ClearAll ' Remove any pre-existing tabs
        End With
    Else
        ' Handle blank paragraphs: Reset formatting completely to avoid inheriting unwanted styles.
         With para.Format
            .LeftIndent = 0
            .FirstLineIndent = 0
            .Alignment = wdAlignParagraphLeft ' Default alignment for blank lines
             ' Reset spacing here in case paragraph had unusual settings; universal setting applied later
            .SpaceBefore = 0
            .TabStops.ClearAll
         End With
    End If
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ApplySecHeaderFormatting
' Author    : User Request/Gemini
' Date      : 2025-04-25
' Purpose   : Applies specific formatting for paragraphs starting with "Sec. ".
'             NOTE: Does NOT set spacing - universal spacing applied after this.
' Arguments : para - The Paragraph object to format.
'             leftIndentPts - The left indentation in points (typically 0).
'---------------------------------------------------------------------------------------
Private Sub ApplySecHeaderFormatting(ByVal para As Paragraph, ByVal leftIndentPts As Single)
     On Error Resume Next ' Prevent errors on unusual paragraph objects
     With para.Format
         .LeftIndent = 0 ' Typically 0
         .FirstLineIndent = -0.5
         .Alignment = wdAlignParagraphLeft ' Sec headers are left-aligned
         ' Reset spacing here in case paragraph had unusual settings; universal setting applied later
         .SpaceBefore = 18
         .TabStops.ClearAll ' Remove any pre-existing tabs
     End With
     On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureTabAfterNumber
' Author    : Gemini
' Date      : 2025-04-25 ' (No functional change needed for V2)
' Purpose   : Ensures exactly one tab character exists immediately after the identified
'             number/letter part (e.g., after "(1)", "(a)", "(i)", "(A)").
'             Removes space if present, inserts tab if missing.
' Arguments : para - The Paragraph object containing the number.
'             numberPart - The exact string of the number part found by regex (e.g., "(1)", "(a)").
'---------------------------------------------------------------------------------------
Private Sub EnsureTabAfterNumber(ByVal para As Paragraph, ByVal numberPart As String)
    Dim paraRange As Range
    Dim startPos As Long
    Dim endPos As Long ' Index of the last character of numberPart within paraRange.Text
    Dim charAfterRange As Range
    Dim targetPos As Long ' Position within the document
    
    On Error GoTo ErrorHandler ' Handle potential errors during range manipulation

    Set paraRange = para.Range ' Includes the paragraph mark (vbCr) at the end
    
    ' Find the start position of the number part within the paragraph's text
    ' Use InStr carefully, consider potential leading chars like (( if regex allows
    ' Regex ensures numberPart is like "(a)", "(1)", etc. We need its position.
    ' A simple InStr should work if numberPart is unique enough at the start.
    startPos = InStr(1, paraRange.text, numberPart, vbBinaryCompare) ' Case-sensitive find is appropriate here

    If startPos > 0 Then
        ' Calculate the index position *immediately after* the numberPart in the Range's text
        endPos = startPos + Len(numberPart) - 1
        
        ' Calculate the absolute document position *after* the numberPart
        targetPos = paraRange.Start + endPos
        
        ' Check if there's a character *after* the numberPart and *before* the paragraph end mark
        If targetPos < paraRange.End - 1 Then ' paraRange.End points *after* the vbCr
            
            ' Get the single character immediately following the numberPart
            Set charAfterRange = ActiveDocument.Range(Start:=targetPos + 1, End:=targetPos + 2)
            
            ' Check if it's a space
            If charAfterRange.text = " " Then
                charAfterRange.Delete ' Delete the space
                ' Re-evaluate the character at the position *after* deletion
                Set charAfterRange = ActiveDocument.Range(Start:=targetPos + 1, End:=targetPos + 2)
            End If
            
             ' Check again if it's NOT a tab (could be start of text, or something else)
            If charAfterRange.text <> vbTab Then
                ' Insert a tab *before* this character (i.e., right after numberPart)
                 ActiveDocument.Range(Start:=targetPos + 1, End:=targetPos + 1).InsertBefore vbTab
            ' Else: It's already a tab, do nothing
            End If
             
        ElseIf targetPos = paraRange.End - 2 Then ' Number part is the last thing before the vbCr
             ' Insert tab before the paragraph mark
             ActiveDocument.Range(Start:=targetPos + 1, End:=targetPos + 1).InsertBefore vbTab
        Else
             ' This case (e.g., empty paragraph matching pattern?) is unlikely but safe to handle
             Debug.Print "Warning: Unexpected position for '" & numberPart & "' in paragraph: " & Left(para.Range.text, 50)
        End If
    Else
        ' Number part wasn't found with InStr, which is unexpected if regex matched.
        Debug.Print "Warning: Could not locate '" & numberPart & "' via InStr in paragraph: " & Left(para.Range.text, 50)
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        Debug.Print "Error in EnsureTabAfterNumber for paragraph '" & Left(para.Range.text, 50) & "': " & Err.Description
        Err.Clear ' Clear the error to allow the macro to continue with the next paragraph
    End If
    ' Cleanup Range objects
    Set paraRange = Nothing
    Set charAfterRange = Nothing
End Sub

