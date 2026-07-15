Attribute VB_Name = "ConvertToCrossRef"
Option Explicit

Private Const MAX_BOOKMARK_NAME As Long = 40
Private Const MAX_CAPTION_SCAN As Long = 120
Private Const FALLBACK_CAPTION_LENGTH As Long = 70
Private Const FALLBACK_CAPTION_WORDS As Long = 8

Private Const SEP_PUNCTUATION As Long = 1
Private Const SEP_CONNECTOR As Long = 2


' ============================================================
' 1. SELECTION-BASED REPLACEMENT FOR THE ORIGINAL MACRO
' ============================================================

Public Sub ConvertToCrossRef()

    Dim doc As Document
    Dim sourceRange As Range
    Dim targetParagraph As Range

    Dim targetMap As Object
    Dim duplicateMap As Object

    Dim lookupNumber As String
    Dim bookmarkName As String
    Dim targetStart As Long

    On Error GoTo FatalError

    Set doc = ActiveDocument
    Set sourceRange = Selection.Range.Duplicate

    If Not CleanSelectedReferenceRange( _
        sourceRange, lookupNumber) Then

        MsgBox "Please select a valid section number.", _
               vbExclamation, "Invalid selection"
        Exit Sub
    End If

    Set targetMap = CreateObject("Scripting.Dictionary")
    Set duplicateMap = CreateObject("Scripting.Dictionary")

    targetMap.CompareMode = vbTextCompare
    duplicateMap.CompareMode = vbTextCompare

    BuildNumberedParagraphMap doc, targetMap, duplicateMap

    If Not targetMap.Exists(lookupNumber) Then

        HighlightFailure sourceRange

        MsgBox "A numbered paragraph for section """ & _
               lookupNumber & """ could not be found." & vbCr & _
               "The selected reference has been highlighted green.", _
               vbInformation, "Cross-reference target not found"

        Exit Sub

    End If

    If duplicateMap.Exists(lookupNumber) Then

        HighlightFailure sourceRange

        MsgBox "More than one numbered paragraph uses section """ & _
               lookupNumber & """." & vbCr & _
               "The selected reference has been highlighted green " & _
               "because the target is ambiguous.", _
               vbInformation, "Ambiguous cross-reference target"

        Exit Sub

    End If

    targetStart = CLng(targetMap(lookupNumber))

    Set targetParagraph = ParagraphRangeAt( _
        doc, targetStart)

    If targetParagraph Is Nothing Then
        GoTo ConversionFailed
    End If

    If Not EnsureCaptionBookmark( _
        doc, targetParagraph, lookupNumber, bookmarkName) Then

        GoTo ConversionFailed
    End If

    If Not ReplaceWithBookmarkNumberCrossReference( _
        sourceRange, bookmarkName) Then

        GoTo ConversionFailed
    End If

    sourceRange.Collapse wdCollapseEnd
    sourceRange.Select

    Exit Sub


ConversionFailed:

    HighlightFailure sourceRange

    MsgBox "The reference to section """ & lookupNumber & _
           """ could not be converted." & vbCr & _
           "The selected reference has been highlighted green.", _
           vbInformation, "Cross-reference conversion failed"

    Exit Sub


FatalError:

    On Error Resume Next
    HighlightFailure sourceRange
    On Error GoTo 0

    MsgBox "The conversion stopped because of error " & _
           Err.Number & ":" & vbCr & Err.Description, _
           vbExclamation, "Cross-reference conversion stopped"

End Sub


' ============================================================
' 2. DOCUMENT-WIDE VERSION
' ============================================================

Public Sub ConvertAllPlainTextSectionRefs()

    Dim doc As Document

    Dim targetMap As Object
    Dim duplicateMap As Object
    Dim bookmarkMap As Object

    Dim rootStory As Range
    Dim story As Range
    Dim nextStory As Range

    Dim converted As Long
    Dim Failed As Long
    Dim skipped As Long

    Dim oldScreenUpdating As Boolean
    Dim undoStarted As Boolean

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo FatalError

    Set doc = ActiveDocument

    Set targetMap = CreateObject("Scripting.Dictionary")
    Set duplicateMap = CreateObject("Scripting.Dictionary")
    Set bookmarkMap = CreateObject("Scripting.Dictionary")

    targetMap.CompareMode = vbTextCompare
    duplicateMap.CompareMode = vbTextCompare
    bookmarkMap.CompareMode = vbTextCompare

    BuildNumberedParagraphMap doc, targetMap, duplicateMap

    If targetMap.count = 0 Then

        MsgBox "There are no usable numbered paragraphs " & _
               "in the main document.", _
               vbExclamation, "No numbered paragraphs"

        Exit Sub

    End If

    oldScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error Resume Next

    Application.UndoRecord.StartCustomRecord _
        "Convert section references to bookmark references"

    undoStarted = (Err.Number = 0)

    Err.Clear
    On Error GoTo FatalError

    For Each rootStory In doc.StoryRanges

        Set story = rootStory.Duplicate

        Do
            Set nextStory = Nothing

            On Error Resume Next
            Set nextStory = story.NextStoryRange
            On Error GoTo FatalError

            ProcessStoryBackward _
                story, _
                targetMap, _
                duplicateMap, _
                bookmarkMap, _
                converted, _
                Failed, _
                skipped

            If nextStory Is Nothing Then Exit Do

            Set story = nextStory
        Loop

    Next rootStory

    On Error Resume Next

    If undoStarted Then
        Application.UndoRecord.EndCustomRecord
    End If

    Application.ScreenUpdating = oldScreenUpdating

    On Error GoTo 0

    MsgBox "Converted: " & converted & vbCr & _
           "Failed and highlighted green: " & Failed & vbCr & _
           "Existing fields skipped: " & skipped, _
           vbInformation, "Cross-reference conversion complete"

    Exit Sub


FatalError:

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    If undoStarted Then
        Application.UndoRecord.EndCustomRecord
    End If

    Application.ScreenUpdating = oldScreenUpdating

    On Error GoTo 0

    MsgBox "The conversion stopped because of error " & _
           errorNumber & ":" & vbCr & errorDescription, _
           vbExclamation, "Cross-reference conversion stopped"

End Sub


' ============================================================
' PROCESS A STORY FROM END TO BEGIN
' ============================================================

Private Sub ProcessStoryBackward( _
    ByVal story As Range, _
    ByVal targetMap As Object, _
    ByVal duplicateMap As Object, _
    ByVal bookmarkMap As Object, _
    ByRef converted As Long, _
    ByRef Failed As Long, _
    ByRef skipped As Long)

    Dim searchEnd As Long
    Dim cueStart As Long
    Dim pluralCue As Boolean

    Dim cue As Range
    Dim sourceRange As Range
    Dim failureRange As Range
    Dim targetParagraph As Range

    Dim candidates As Collection
    Dim tokens As Collection

    Dim token As String
    Dim key As String
    Dim bookmarkName As String

    Dim originalStart As Long
    Dim originalLength As Long

    Dim i As Long

    searchEnd = story.End

    Do While searchEnd > story.Start

        pluralCue = False

        Set cue = FindLastCue( _
            story, searchEnd, pluralCue)

        If cue Is Nothing Then Exit Do

        cueStart = cue.Start

        Set candidates = New Collection
        Set tokens = New Collection

        CollectReferenceRanges _
            story, _
            cue.End, _
            pluralCue, _
            candidates, _
            tokens

        ' Work right to left so insertion of a field does not
        ' invalidate ranges that remain to be processed.
        For i = candidates.count To 1 Step -1

            Set sourceRange = candidates(i)
            token = CStr(tokens(i))
            key = NormalizeNumber(token)

            originalStart = sourceRange.Start
            originalLength = sourceRange.End - sourceRange.Start

            If IsInsideField(sourceRange) Then

                skipped = skipped + 1

            ElseIf NormalizeNumber(sourceRange.text) <> key Then

                HighlightFailure sourceRange
                Failed = Failed + 1

            ElseIf Not targetMap.Exists(key) Then

                HighlightFailure sourceRange
                Failed = Failed + 1

            ElseIf duplicateMap.Exists(key) Then

                HighlightFailure sourceRange
                Failed = Failed + 1

            Else

                bookmarkName = vbNullString

                If bookmarkMap.Exists(key) Then

                    bookmarkName = CStr(bookmarkMap(key))

                Else

                    Set targetParagraph = ParagraphRangeAt( _
                        ActiveDocument, CLng(targetMap(key)))

                    If Not targetParagraph Is Nothing Then

                        If EnsureCaptionBookmark( _
                            ActiveDocument, _
                            targetParagraph, _
                            key, _
                            bookmarkName) Then

                            bookmarkMap.Add key, bookmarkName
                        End If

                    End If

                End If

                If Len(bookmarkName) = 0 Then

                    HighlightFailure sourceRange
                    Failed = Failed + 1

                ElseIf ReplaceWithBookmarkNumberCrossReference( _
                    sourceRange, bookmarkName) Then

                    converted = converted + 1

                Else

                    Set failureRange = story.Duplicate

                    On Error Resume Next

                    failureRange.SetRange _
                        originalStart, _
                        originalStart + originalLength

                    HighlightFailure failureRange

                    On Error GoTo 0

                    Failed = Failed + 1

                End If

            End If

        Next i

        searchEnd = cueStart

    Loop

End Sub


' ============================================================
' BUILD A SECTION-NUMBER-TO-PARAGRAPH MAP
' ============================================================

Private Sub BuildNumberedParagraphMap( _
    ByVal doc As Document, _
    ByVal targetMap As Object, _
    ByVal duplicateMap As Object)

    Dim mainStory As Range
    Dim para As Paragraph

    Dim listType As WdListType
    Dim listText As String
    Dim numberText As String
    Dim key As String

    Set mainStory = doc.StoryRanges(wdMainTextStory)

    For Each para In mainStory.Paragraphs

        listText = vbNullString
        listType = wdListNoNumbering

        On Error Resume Next

        listType = para.Range.ListFormat.listType
        listText = para.Range.ListFormat.ListString

        On Error GoTo 0

        If listType <> wdListNoNumbering And _
           Len(listText) > 0 Then

            numberText = ExtractLeadingSectionNumber(listText)
            key = NormalizeNumber(numberText)

            If Len(key) > 0 Then

                If targetMap.Exists(key) Then

                    duplicateMap(key) = True

                Else

                    targetMap.Add key, CLng(para.Range.Start)

                End If

            End If

        End If

    Next para

End Sub


Private Function ParagraphRangeAt( _
    ByVal doc As Document, _
    ByVal paragraphStart As Long) As Range

    Dim pointRange As Range

    On Error GoTo NotFound

    Set pointRange = doc.Range( _
        Start:=paragraphStart, _
        End:=paragraphStart)

    Set ParagraphRangeAt = _
        pointRange.Paragraphs(1).Range.Duplicate

    Exit Function

NotFound:

    Set ParagraphRangeAt = Nothing

End Function


' ============================================================
' CREATE OR REUSE THE CAPTION-NAMED BOOKMARK
' ============================================================

Private Function EnsureCaptionBookmark( _
    ByVal doc As Document, _
    ByVal targetParagraph As Range, _
    ByVal sectionNumber As String, _
    ByRef bookmarkName As String) As Boolean

    Dim captionRange As Range

    Dim captionText As String
    Dim baseName As String
    Dim candidateName As String
    Dim sectionSuffix As String
    Dim numericSuffix As String

    Dim suffixNumber As Long
    Dim availableLength As Long

    On Error GoTo Failed

    Set captionRange = GetCaptionRange( _
        targetParagraph, captionText)

    If captionRange Is Nothing Then Exit Function

    baseName = MakeBookmarkBaseName( _
        captionText, sectionNumber)

    sectionSuffix = "_" & _
        SanitizeBookmarkText(sectionNumber)

    candidateName = Left$(baseName, MAX_BOOKMARK_NAME)

    suffixNumber = 0

    Do

        If Not doc.Bookmarks.Exists(candidateName) Then

            doc.Bookmarks.Add _
                Name:=candidateName, _
                Range:=captionRange

            bookmarkName = candidateName
            EnsureCaptionBookmark = True
            Exit Function

        End If

        If BookmarkIsInParagraph( _
            doc.Bookmarks(candidateName), _
            targetParagraph) Then

            bookmarkName = candidateName
            EnsureCaptionBookmark = True
            Exit Function

        End If

        suffixNumber = suffixNumber + 1

        If suffixNumber = 1 Then
            numericSuffix = sectionSuffix
        Else
            numericSuffix = sectionSuffix & _
                            "_" & CStr(suffixNumber)
        End If

        availableLength = _
            MAX_BOOKMARK_NAME - Len(numericSuffix)

        If availableLength < 1 Then
            availableLength = 1
        End If

        candidateName = _
            Left$(baseName, availableLength) & _
            numericSuffix

    Loop

Failed:

    bookmarkName = vbNullString
    EnsureCaptionBookmark = False

End Function


Private Function BookmarkIsInParagraph( _
    ByVal targetBookmark As Bookmark, _
    ByVal targetParagraph As Range) As Boolean

    Dim bookmarkRange As Range

    On Error GoTo NotSameParagraph

    Set bookmarkRange = targetBookmark.Range

    BookmarkIsInParagraph = _
        bookmarkRange.Start >= targetParagraph.Start And _
        bookmarkRange.Start < targetParagraph.End

    Exit Function

NotSameParagraph:

    BookmarkIsInParagraph = False

End Function


' ============================================================
' EXTRACT THE CAPTION
' ============================================================

Private Function GetCaptionRange( _
    ByVal targetParagraph As Range, _
    ByRef captionText As String) As Range

    Dim usableRange As Range
    Dim captionRange As Range

    Dim paragraphText As String
    Dim captionEnd As Long
    Dim fallbackEnd As Long

    Set usableRange = targetParagraph.Duplicate

    ' Remove paragraph and table-cell terminators.
    Do While usableRange.End > usableRange.Start

        Select Case AscW(Right$(usableRange.text, 1))

            Case 7, 13
                usableRange.MoveEnd wdCharacter, -1

            Case Else
                Exit Do

        End Select

    Loop

    TrimLeadingWhitespace usableRange
    TrimTrailingWhitespace usableRange

    If usableRange.End <= usableRange.Start Then Exit Function

    paragraphText = usableRange.text

    captionEnd = FindCaptionTerminator(paragraphText)

    If captionEnd = 0 Then

        fallbackEnd = FindFallbackCaptionEnd( _
            paragraphText, _
            FALLBACK_CAPTION_WORDS, _
            FALLBACK_CAPTION_LENGTH)

        captionEnd = fallbackEnd

    End If

    If captionEnd <= 0 Then Exit Function

    Set captionRange = usableRange.Duplicate

    captionRange.SetRange _
        usableRange.Start, _
        usableRange.Start + captionEnd

    captionText = Trim$(captionRange.text)

    If Len(captionText) = 0 Then Exit Function

    Set GetCaptionRange = captionRange

End Function


Private Function FindCaptionTerminator( _
    ByVal paragraphText As String) As Long

    Dim scanLimit As Long
    Dim i As Long

    Dim ch As String
    Dim previousCharacter As String
    Dim nextCharacter As String

    scanLimit = Len(paragraphText)

    If scanLimit > MAX_CAPTION_SCAN Then
        scanLimit = MAX_CAPTION_SCAN
    End If

    For i = 1 To scanLimit

        ch = Mid$(paragraphText, i, 1)

        If ch = "." Or ch = ":" Then

            previousCharacter = vbNullString
            nextCharacter = vbNullString

            If i > 1 Then
                previousCharacter = _
                    Mid$(paragraphText, i - 1, 1)
            End If

            If i < Len(paragraphText) Then
                nextCharacter = _
                    Mid$(paragraphText, i + 1, 1)
            End If

            ' Ignore periods inside decimal references,
            ' such as the periods in 5.4.1.
            If ch = "." And _
               IsDigitCharacter(previousCharacter) And _
               IsDigitCharacter(nextCharacter) Then

                GoTo ContinueLoop
            End If

            ' A caption-ending period or colon should ordinarily
            ' be followed by whitespace or the end of the paragraph.
            If i = Len(paragraphText) Or _
               IsWhitespaceCharacter(nextCharacter) Then

                FindCaptionTerminator = i
                Exit Function

            End If

        End If

ContinueLoop:
    Next i

End Function


Private Function FindFallbackCaptionEnd( _
    ByVal paragraphText As String, _
    ByVal maximumWords As Long, _
    ByVal maximumCharacters As Long) As Long

    Dim limit As Long
    Dim i As Long
    Dim wordCount As Long
    Dim inWord As Boolean

    Dim ch As String

    limit = Len(paragraphText)

    If limit > maximumCharacters Then
        limit = maximumCharacters
    End If

    For i = 1 To limit

        ch = Mid$(paragraphText, i, 1)

        If IsWhitespaceCharacter(ch) Then

            If inWord Then

                wordCount = wordCount + 1
                inWord = False

                If wordCount >= maximumWords Then

                    FindFallbackCaptionEnd = i - 1
                    Exit Function

                End If

            End If

        Else

            inWord = True

        End If

    Next i

    FindFallbackCaptionEnd = limit

End Function


' ============================================================
' MAKE A WORD-VALID BOOKMARK NAME
' ============================================================

Private Function MakeBookmarkBaseName( _
    ByVal captionText As String, _
    ByVal sectionNumber As String) As String

    Dim cleanCaption As String

    cleanCaption = captionText

    Do While Len(cleanCaption) > 0

        Select Case Right$(cleanCaption, 1)

            Case ".", ":", ";", ","
                cleanCaption = _
                    Left$(cleanCaption, Len(cleanCaption) - 1)

            Case Else
                Exit Do

        End Select

    Loop

    cleanCaption = SanitizeBookmarkText(cleanCaption)

    If Len(cleanCaption) = 0 Then
        cleanCaption = "Section_" & _
                       SanitizeBookmarkText(sectionNumber)
    End If

    ' Word bookmark names must begin with a letter.
    MakeBookmarkBaseName = _
        Left$("Sec_" & cleanCaption, MAX_BOOKMARK_NAME)

End Function


Private Function SanitizeBookmarkText( _
    ByVal value As String) As String

    Dim result As String
    Dim ch As String

    Dim i As Long
    Dim previousWasUnderscore As Boolean

    For i = 1 To Len(value)

        ch = Mid$(value, i, 1)

        If IsAlphaNumericCharacter(ch) Then

            result = result & ch
            previousWasUnderscore = False

        ElseIf Not previousWasUnderscore Then

            result = result & "_"
            previousWasUnderscore = True

        End If

    Next i

    Do While Len(result) > 0 And _
             Left$(result, 1) = "_"

        result = Mid$(result, 2)
    Loop

    Do While Len(result) > 0 And _
             Right$(result, 1) = "_"

        result = Left$(result, Len(result) - 1)
    Loop

    SanitizeBookmarkText = result

End Function


' ============================================================
' INSERT A REF FIELD TO THE BOOKMARK
' ============================================================

Private Function ReplaceWithBookmarkNumberCrossReference( _
    ByVal sourceRange As Range, _
    ByVal bookmarkName As String) As Boolean

    Dim insertionRange As Range
    Dim restorationRange As Range
    Dim insertedField As Field

    Dim originalText As String
    Dim originalStart As Long

    On Error GoTo InsertFailed

    originalText = sourceRange.text
    originalStart = sourceRange.Start

    Set insertionRange = sourceRange.Duplicate

    insertionRange.text = vbNullString
    insertionRange.Collapse wdCollapseStart

    ' \w = full-context paragraph number
    ' \h = hyperlink to the bookmark
    Set insertedField = insertionRange.Fields.Add( _
        Range:=insertionRange, _
        Type:=wdFieldRef, _
        text:=bookmarkName & " \w \h", _
        PreserveFormatting:=True)

    insertedField.Update

    If InStr( _
        1, insertedField.result.text, _
        "Error!", vbTextCompare) > 0 Then

        GoTo InsertFailed
    End If

    ReplaceWithBookmarkNumberCrossReference = True
    Exit Function


InsertFailed:

    On Error Resume Next

    If Not insertedField Is Nothing Then
        insertedField.Delete
    End If

    Set restorationRange = sourceRange.Duplicate

    restorationRange.SetRange _
        originalStart, originalStart

    restorationRange.InsertAfter originalText

    On Error GoTo 0

    ReplaceWithBookmarkNumberCrossReference = False

End Function


' ============================================================
' FIND SECTION / SECTIONS / SECTION SIGNS
' ============================================================

Private Function FindLastCue( _
    ByVal story As Range, _
    ByVal searchEnd As Long, _
    ByRef pluralCue As Boolean) As Range

    Dim candidate As Range
    Dim best As Range

    Dim bestStart As Long
    Dim candidatePlural As Boolean

    bestStart = -1
    pluralCue = False

    Set candidate = FindLastText( _
        story, searchEnd, "Sections", True)

    If Not candidate Is Nothing Then

        If candidate.Start > bestStart Then

            Set best = candidate.Duplicate
            bestStart = candidate.Start
            pluralCue = True

        End If

    End If

    Set candidate = FindLastText( _
        story, searchEnd, "Section", True)

    If Not candidate Is Nothing Then

        If candidate.Start > bestStart Then

            Set best = candidate.Duplicate
            bestStart = candidate.Start
            pluralCue = False

        End If

    End If

    Set candidate = FindLastText( _
        story, searchEnd, ChrW(167), False)

    If Not candidate Is Nothing Then

        candidatePlural = False

        If candidate.Start > story.Start Then

            If CharacterAt( _
                story, candidate.Start - 1) = ChrW(167) Then

                candidate.SetRange _
                    candidate.Start - 1, candidate.End

                candidatePlural = True

            End If

        End If

        If Not candidatePlural Then

            If candidate.End < searchEnd Then

                If CharacterAt( _
                    story, candidate.End) = ChrW(167) Then

                    candidate.SetRange _
                        candidate.Start, candidate.End + 1

                    candidatePlural = True

                End If

            End If

        End If

        If candidate.Start > bestStart Then

            Set best = candidate.Duplicate
            bestStart = candidate.Start
            pluralCue = candidatePlural

        End If

    End If

    If Not best Is Nothing Then
        Set FindLastCue = best
    End If

End Function


Private Function FindLastText( _
    ByVal story As Range, _
    ByVal searchEnd As Long, _
    ByVal findText As String, _
    ByVal matchWholeWord As Boolean) As Range

    Dim searchRange As Range

    If searchEnd <= story.Start Then Exit Function

    Set searchRange = story.Duplicate

    searchRange.SetRange _
        story.Start, searchEnd

    With searchRange.Find

        .ClearFormatting
        .text = findText

        .Forward = False
        .Wrap = wdFindStop
        .Format = False

        .MatchCase = False
        .matchWholeWord = matchWholeWord
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        If .Execute Then
            Set FindLastText = searchRange.Duplicate
        End If

    End With

End Function


' ============================================================
' PARSE SECTION NUMBERS FOLLOWING THE CUE
' ============================================================

Private Sub CollectReferenceRanges( _
    ByVal story As Range, _
    ByVal afterCue As Long, _
    ByVal pluralCue As Boolean, _
    ByVal candidates As Collection, _
    ByVal tokens As Collection)

    Dim position As Long
    Dim tokenStart As Long
    Dim tokenEnd As Long
    Dim separatorKind As Long

    position = afterCue

    SkipHorizontalSpaces story, position

    If Not ReadSectionNumberAt( _
        story, position, tokenStart, tokenEnd) Then

        Exit Sub
    End If

    AddReferenceCandidate _
        story, tokenStart, tokenEnd, candidates, tokens

    Do

        position = tokenEnd
        separatorKind = 0

        If Not ConsumeListSeparator( _
            story, position, separatorKind) Then

            Exit Do
        End If

        SkipHorizontalSpaces story, position

        If IsExplicitCueAt(story, position) Then
            Exit Do
        End If

        ' Avoid treating a number after a singular reference and
        ' comma as a second section number:
        '
        '   Section 5.1, 2025 revision
        If Not pluralCue And _
           separatorKind = SEP_PUNCTUATION Then

            Exit Do
        End If

        If Not ReadSectionNumberAt( _
            story, position, tokenStart, tokenEnd) Then

            Exit Do
        End If

        AddReferenceCandidate _
            story, tokenStart, tokenEnd, candidates, tokens

    Loop

End Sub


Private Sub AddReferenceCandidate( _
    ByVal story As Range, _
    ByVal tokenStart As Long, _
    ByVal tokenEnd As Long, _
    ByVal candidates As Collection, _
    ByVal tokens As Collection)

    Dim candidate As Range

    Set candidate = story.Duplicate

    candidate.SetRange tokenStart, tokenEnd

    candidates.Add candidate
    tokens.Add candidate.text

End Sub


Private Function ReadSectionNumberAt( _
    ByVal story As Range, _
    ByVal position As Long, _
    ByRef tokenStart As Long, _
    ByRef tokenEnd As Long) As Boolean

    Dim p As Long
    Dim q As Long

    Dim ch As String

    p = position

    If p >= story.End Then Exit Function

    If Not IsDigitCharacter( _
        CharacterAt(story, p)) Then

        Exit Function
    End If

    tokenStart = p
    p = p + 1

    Do While p < story.End

        ch = CharacterAt(story, p)

        If IsAlphaNumericCharacter(ch) Then

            p = p + 1

        ElseIf ch = "." Then

            If p + 1 < story.End Then

                If IsAlphaNumericCharacter( _
                    CharacterAt(story, p + 1)) Then

                    p = p + 1
                Else
                    Exit Do
                End If

            Else
                Exit Do
            End If

        ElseIf ch = "(" Then

            q = p + 1

            If q >= story.End Then Exit Do

            If Not IsAlphaNumericCharacter( _
                CharacterAt(story, q)) Then

                Exit Do
            End If

            Do While q < story.End And _
                     IsAlphaNumericCharacter( _
                        CharacterAt(story, q))

                q = q + 1
            Loop

            If q < story.End And _
               CharacterAt(story, q) = ")" Then

                p = q + 1
            Else
                Exit Do
            End If

        Else

            Exit Do

        End If

    Loop

    tokenEnd = p

    ReadSectionNumberAt = _
        (tokenEnd > tokenStart)

End Function


' ============================================================
' PARSE LIST SEPARATORS
' ============================================================

Private Function ConsumeListSeparator( _
    ByVal story As Range, _
    ByRef position As Long, _
    ByRef separatorKind As Long) As Boolean

    Dim p As Long
    Dim ch As String

    p = position

    SkipHorizontalSpaces story, p

    If p >= story.End Then Exit Function

    ch = CharacterAt(story, p)

    If ch = "," Or ch = ";" Then

        p = p + 1

        SkipHorizontalSpaces story, p

        If StartsWithWordAt(story, p, "and") Then

            p = p + Len("and")

        ElseIf StartsWithWordAt(story, p, "or") Then

            p = p + Len("or")

        End If

        position = p
        separatorKind = SEP_PUNCTUATION
        ConsumeListSeparator = True

        Exit Function

    End If

    If ch = "-" Or _
       ch = ChrW(8211) Or _
       ch = ChrW(8212) Then

        position = p + 1
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

    If ch = "&" Then

        position = p + 1
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

    If StartsWithWordAt(story, p, "through") Then

        position = p + Len("through")
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

    If StartsWithWordAt(story, p, "and") Then

        position = p + Len("and")
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

    If StartsWithWordAt(story, p, "or") Then

        position = p + Len("or")
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

    If StartsWithWordAt(story, p, "to") Then

        position = p + Len("to")
        separatorKind = SEP_CONNECTOR
        ConsumeListSeparator = True

        Exit Function

    End If

End Function


Private Function IsExplicitCueAt( _
    ByVal story As Range, _
    ByVal position As Long) As Boolean

    If position >= story.End Then Exit Function

    If CharacterAt(story, position) = ChrW(167) Then

        IsExplicitCueAt = True
        Exit Function

    End If

    If StartsWithWordAt(story, position, "Sections") Then

        IsExplicitCueAt = True
        Exit Function

    End If

    If StartsWithWordAt(story, position, "Section") Then
        IsExplicitCueAt = True
    End If

End Function


' ============================================================
' SELECTION CLEANUP
' ============================================================

Private Function CleanSelectedReferenceRange( _
    ByVal sourceRange As Range, _
    ByRef lookupNumber As String) As Boolean

    Dim ch As String

    Do While sourceRange.End > sourceRange.Start

        ch = sourceRange.Characters(1).text

        If IsWhitespaceCharacter(ch) Then
            sourceRange.MoveStart wdCharacter, 1
        Else
            Exit Do
        End If

    Loop

    Do While sourceRange.End > sourceRange.Start

        ch = Right$(sourceRange.text, 1)

        Select Case AscW(ch)

            Case 9, 11, 13, 32, 160
                sourceRange.MoveEnd wdCharacter, -1

            Case Else
                Exit Do

        End Select

    Loop

    ' Leave final sentence punctuation outside the replacement range.
    Do While sourceRange.End > sourceRange.Start

        ch = Right$(sourceRange.text, 1)

        If ch = "." And _
           Not PeriodIsPartOfSectionNumber(sourceRange.text) Then

            sourceRange.MoveEnd wdCharacter, -1
        Else
            Exit Do
        End If

    Loop

    If sourceRange.End <= sourceRange.Start Then
        Exit Function
    End If

    lookupNumber = NormalizeNumber(sourceRange.text)

    If Len(lookupNumber) = 0 Then Exit Function

    If Not IsDigitCharacter(Left$(lookupNumber, 1)) Then
        Exit Function
    End If

    CleanSelectedReferenceRange = True

End Function


Private Function PeriodIsPartOfSectionNumber( _
    ByVal value As String) As Boolean

    Dim lengthValue As Long

    lengthValue = Len(value)

    If lengthValue < 2 Then Exit Function

    ' A final period after a digit is treated as punctuation.
    ' Internal periods remain part of the section number.
    PeriodIsPartOfSectionNumber = False

End Function


' ============================================================
' EXTRACT AND NORMALIZE NUMBERING
' ============================================================

Private Function ExtractLeadingSectionNumber( _
    ByVal value As String) As String

    Dim p As Long
    Dim q As Long
    Dim numberStart As Long

    Dim ch As String

    p = 1

    Do While p <= Len(value)

        If IsDigitCharacter(Mid$(value, p, 1)) Then
            Exit Do
        End If

        p = p + 1
    Loop

    If p > Len(value) Then Exit Function

    numberStart = p
    p = p + 1

    Do While p <= Len(value)

        ch = Mid$(value, p, 1)

        If IsAlphaNumericCharacter(ch) Then

            p = p + 1

        ElseIf ch = "." Then

            If p < Len(value) And _
               IsAlphaNumericCharacter( _
                    Mid$(value, p + 1, 1)) Then

                p = p + 1
            Else
                Exit Do
            End If

        ElseIf ch = "(" Then

            q = p + 1

            If q > Len(value) Then Exit Do

            If Not IsAlphaNumericCharacter( _
                Mid$(value, q, 1)) Then

                Exit Do
            End If

            Do While q <= Len(value) And _
                     IsAlphaNumericCharacter( _
                        Mid$(value, q, 1))

                q = q + 1
            Loop

            If q <= Len(value) And _
               Mid$(value, q, 1) = ")" Then

                p = q + 1
            Else
                Exit Do
            End If

        Else

            Exit Do

        End If

    Loop

    ExtractLeadingSectionNumber = _
        Mid$(value, numberStart, p - numberStart)

End Function


Private Function NormalizeNumber( _
    ByVal value As String) As String

    value = Replace(value, " ", vbNullString)
    value = Replace(value, vbTab, vbNullString)
    value = Replace(value, ChrW(160), vbNullString)
    value = Replace(value, ChrW(8239), vbNullString)

    Do While Len(value) > 0 And _
             Right$(value, 1) = "."

        value = Left$(value, Len(value) - 1)
    Loop

    NormalizeNumber = LCase$(value)

End Function


' ============================================================
' FIELD AND HIGHLIGHT HELPERS
' ============================================================

Private Function IsInsideField( _
    ByVal sourceRange As Range) As Boolean

    Dim probe As Range

    On Error GoTo NotInsideField

    If sourceRange.Fields.count > 0 Then

        IsInsideField = True
        Exit Function

    End If

    Set probe = sourceRange.Duplicate
    probe.Collapse wdCollapseStart

    If probe.Information(wdInFieldCode) Or _
       probe.Information(wdInFieldResult) Then

        IsInsideField = True
        Exit Function

    End If

    If sourceRange.End > sourceRange.Start Then

        Set probe = sourceRange.Duplicate

        probe.SetRange _
            sourceRange.End - 1, _
            sourceRange.End - 1

        If probe.Information(wdInFieldCode) Or _
           probe.Information(wdInFieldResult) Then

            IsInsideField = True
            Exit Function

        End If

    End If

NotInsideField:

End Function


Private Sub HighlightFailure(ByVal sourceRange As Range)

    On Error Resume Next
    sourceRange.HighlightColorIndex = wdBrightGreen
    On Error GoTo 0

End Sub


' ============================================================
' RANGE AND CHARACTER HELPERS
' ============================================================

Private Sub SkipHorizontalSpaces( _
    ByVal story As Range, _
    ByRef position As Long)

    Dim ch As String

    Do While position < story.End

        ch = CharacterAt(story, position)

        If IsHorizontalWhitespaceCharacter(ch) Then
            position = position + 1
        Else
            Exit Do
        End If

    Loop

End Sub


Private Sub TrimLeadingWhitespace(ByVal sourceRange As Range)

    Dim ch As String

    Do While sourceRange.End > sourceRange.Start

        ch = sourceRange.Characters(1).text

        If IsWhitespaceCharacter(ch) Then
            sourceRange.MoveStart wdCharacter, 1
        Else
            Exit Do
        End If

    Loop

End Sub


Private Sub TrimTrailingWhitespace(ByVal sourceRange As Range)

    Dim ch As String

    Do While sourceRange.End > sourceRange.Start

        ch = Right$(sourceRange.text, 1)

        If IsWhitespaceCharacter(ch) Then
            sourceRange.MoveEnd wdCharacter, -1
        Else
            Exit Do
        End If

    Loop

End Sub


Private Function StartsWithWordAt( _
    ByVal story As Range, _
    ByVal position As Long, _
    ByVal wordText As String) As Boolean

    Dim wordEnd As Long
    Dim followingCharacter As String

    wordEnd = position + Len(wordText)

    If position < story.Start Then Exit Function
    If wordEnd > story.End Then Exit Function

    If StrComp( _
        TextBetween(story, position, wordEnd), _
        wordText, vbTextCompare) <> 0 Then

        Exit Function
    End If

    If wordEnd < story.End Then

        followingCharacter = _
            CharacterAt(story, wordEnd)

        If IsAlphaNumericCharacter(followingCharacter) Or _
           followingCharacter = "_" Then

            Exit Function
        End If

    End If

    StartsWithWordAt = True

End Function


Private Function CharacterAt( _
    ByVal story As Range, _
    ByVal position As Long) As String

    Dim characterRange As Range

    If position < story.Start Then Exit Function
    If position >= story.End Then Exit Function

    Set characterRange = story.Duplicate

    characterRange.SetRange _
        position, position + 1

    CharacterAt = characterRange.text

End Function


Private Function TextBetween( _
    ByVal story As Range, _
    ByVal startPosition As Long, _
    ByVal endPosition As Long) As String

    Dim textRange As Range

    If endPosition <= startPosition Then Exit Function

    Set textRange = story.Duplicate

    textRange.SetRange _
        startPosition, endPosition

    TextBetween = textRange.text

End Function


Private Function IsDigitCharacter( _
    ByVal value As String) As Boolean

    Dim characterCode As Long

    If Len(value) = 0 Then Exit Function

    characterCode = AscW(Left$(value, 1))

    IsDigitCharacter = _
        characterCode >= AscW("0") And _
        characterCode <= AscW("9")

End Function


Private Function IsAlphaNumericCharacter( _
    ByVal value As String) As Boolean

    Dim characterCode As Long

    If Len(value) = 0 Then Exit Function

    characterCode = AscW(Left$(value, 1))

    IsAlphaNumericCharacter = _
        (characterCode >= AscW("0") And _
         characterCode <= AscW("9")) Or _
        (characterCode >= AscW("A") And _
         characterCode <= AscW("Z")) Or _
        (characterCode >= AscW("a") And _
         characterCode <= AscW("z"))

End Function


Private Function IsHorizontalWhitespaceCharacter( _
    ByVal value As String) As Boolean

    Dim characterCode As Long

    If Len(value) = 0 Then Exit Function

    characterCode = AscW(Left$(value, 1))

    Select Case characterCode

        Case 9, 32, 160, 8239
            IsHorizontalWhitespaceCharacter = True

    End Select

End Function


Private Function IsWhitespaceCharacter( _
    ByVal value As String) As Boolean

    Dim characterCode As Long

    If Len(value) = 0 Then Exit Function

    characterCode = AscW(Left$(value, 1))

    Select Case characterCode

        Case 9, 10, 11, 13, 32, 160, 8239
            IsWhitespaceCharacter = True

    End Select

End Function

'Option Explicit
'
'Sub ConvertToCrossRef()
'
'    Dim refList As Variant
'    Dim LookUp As String
'    Dim Ref As String
'    Dim s As Integer, t As Integer
'    Dim i As Integer
'    Dim oRng As Range
'    Dim oRngStr As String
'    Dim spaceAfter As Boolean
'    Dim numPart As String ' Variable to hold the extracted number part
'    Dim delimiterPos As Integer ' Position of the first space or tab
'
'    On Error GoTo CleanUp ' Changed error handler name for clarity
'
'    Set oRng = Selection.Range
'    oRngStr = Selection.Range.text
'
'    ' Check if there was a space immediately after the original selection
'    If Len(oRngStr) > 0 Then
'        If oRng.Characters(Len(oRngStr)).text = Chr(32) Then spaceAfter = True
'    End If
'
'    ' --- Clean up the selected text ---
'    With oRng
'        ' Trim leading/trailing spaces from the range itself first
'        .MoveEndWhile Chr(32), wdBackward
'        .MoveStartWhile Chr(32), wdForward ' More reliable way to trim leading spaces
'
'        ' Refine trimming of trailing characters
'        Do While .End > .Start
'            Select Case Asc(Right(.text, 1))
'                Case 13, 11, 32, 46 ' CR, VT, Space, Period
'                    .MoveEnd wdCharacter, -1
'                Case Else
'                    Exit Do ' Stop if it's not one of the trailing chars we want to remove
'            End Select
'        Loop
'
'        ' Check if selection is empty after cleaning
'        If .End <= .Start Then GoTo ErrExitEmptySelection
'
'        LookUp = .text ' Assign the cleaned text to LookUp
'    End With
'    ' --- End Selection Cleaning ---
'
'
'    With ActiveDocument
'        ' Use wdRefTypeNumberedItem to retrieve numbered paragraphs
'        refList = .GetCrossReferenceItems(wdRefTypeNumberedItem)
'
'        If IsEmpty(refList) Then GoTo ErrExitNoItems ' Check if any numbered items exist
'
'        ' --- Loop through potential reference targets ---
'        For i = 1 To UBound(refList) ' Loop forward for clarity, index 'i' is needed later
'            Ref = Trim(refList(i)) ' Get the numbered item string, e.g., "3. Heading Text" or "3.2 Subheading"
'
'            ' Find the position of the first space or tab, which separates the number from the text
'            s = InStr(1, Ref, " ")
'            t = InStr(1, Ref, Chr(9)) ' Chr(9) is Tab
'
'            If s > 0 And t > 0 Then
'                delimiterPos = IIf(s < t, s, t) ' Find the earlier delimiter
'            ElseIf s > 0 Then
'                delimiterPos = s ' Only space found
'            ElseIf t > 0 Then
'                delimiterPos = t ' Only tab found
'            Else
'                delimiterPos = 0 ' No delimiter found (might be just the number)
'            End If
'
'            ' Extract the number part
'            If delimiterPos > 0 Then
'                numPart = Trim(Left(Ref, delimiterPos - 1)) ' Get text before delimiter, trim spaces
'            Else
'                numPart = Ref ' The whole string is potentially the number part
'            End If
'
'            ' *** CORE FIX: Remove trailing period from the extracted number part if it exists ***
'            If Right(numPart, 1) = "." Then
'                numPart = Left(numPart, Len(numPart) - 1)
'            End If
'
'            ' Now compare the cleaned number part from the document (NumPart)
'            ' with the cleaned selected text (LookUp)
'            If StrComp(numPart, LookUp, vbTextCompare) = 0 Then ' Case-insensitive comparison
'                ' Match found, exit the loop
'                Exit For
'            End If
'        Next i
'        ' --- End Loop ---
'
'
'        ' --- Insert Cross Reference if match found ---
'        If i <= UBound(refList) Then ' Check if the loop completed because a match was found (i will be <= UBound)
'            ' A match was found at index 'i'
'            Selection.InsertCrossReference ReferenceType:="Numbered item", _
'                                           ReferenceKind:=wdNumberFullContext, _
'                                           ReferenceItem:=CStr(i), _
'                                           InsertAsHyperlink:=True, _
'                                           IncludePosition:=False, _
'                                           SeparateNumbers:=False, _
'                                           SeparatorString:=" "
'            ' Add back the trailing space if it was originally present
'            If spaceAfter Then Selection.Range.InsertAfter (Chr(32))
'            ' Collapse selection after inserting
'            Selection.Collapse wdCollapseEnd
'        Else
'            ' No match was found after checking all items
'            MsgBox "A cross reference to """ & LookUp & """ couldn't be set." & vbCr & _
'                   "A paragraph starting with that number" & vbCr & _
'                   "couldn't be found in the document.", _
'                   vbInformation, "Cross reference target not found"
'        End If
'        ' --- End Insert ---
'
'    End With
'
'    GoTo CleanUp ' Skip error handlers if successful
'
'ErrExitEmptySelection:
'    MsgBox "Please select a valid paragraph number reference.", _
'           vbExclamation, "Invalid selection"
'    GoTo CleanUp
'
'ErrExitNoItems:
'    MsgBox "There are no numbered items in this document to cross-reference.", _
'            vbExclamation, "No Numbered Items Found"
'    GoTo CleanUp
'
'CleanUp:
'    Set oRng = Nothing
'    Set refList = Nothing
'
'End Sub
'
