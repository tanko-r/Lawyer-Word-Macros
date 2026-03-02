Attribute VB_Name = "ExhibitList"
Option Explicit

Sub GenerateSimpleExhibitList()

    Dim dictExhibits As Object ' Dictionary to store Exhibit Ref -> Title
    Dim rngSearch As Range     ' Range to search within (entire document)
    Dim rngFound As Range      ' Range where a match is found
    Dim para As Paragraph      ' Paragraph object containing the found text
    Dim nextPara As Paragraph  ' The paragraph following 'para'
    Dim rngTitle As Range      ' Range of the paragraph likely containing the title
    Dim refText As String      ' The found reference text (e.g., "Exhibit A")
    Dim titleText As String    ' The extracted title text
    Dim exhibitKeys() As Variant ' Array to hold sorted exhibit keys
    Dim tempKeys As Variant    ' Temporary variant to hold keys result
    Dim i As Long              ' Loop counter
    Dim startOfList As Range   ' To remember where to insert the list

    ' --- Initialization ---
    Set dictExhibits = CreateObject("Scripting.Dictionary")
    dictExhibits.CompareMode = vbTextCompare ' Case-insensitive keys

    Set rngSearch = ActiveDocument.content
    Set rngFound = ActiveDocument.content ' Work with a copy for Find
    Set startOfList = Selection.Range ' Remember current cursor position

    Application.ScreenUpdating = True ' Speed up macro

    ' --- Find Configuration ---
    With rngFound.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = False             ' Don't look for specific formatting initially
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = False          ' Handles "Exhibit" or "EXHIBIT"
        .MatchWholeWord = False
        .MatchWildcards = True      ' MUST be True for the pattern

        ' Pattern: "Exhibit" followed by 1+ spaces, followed by a single letter or number
        .text = "aSCHEDULE"

        ' We will check alignment AFTER finding the text
        '.ParagraphFormat.Alignment = wdAlignParagraphLeft ' Reset alignment search initially
    End With

    ' --- Search Loop ---
    Do While rngFound.Find.Execute

        Set para = rngFound.Paragraphs(1) ' Get the paragraph object of the match

        ' --- VERIFICATION STEP ---
        ' Check 1: Is the paragraph center-aligned?
        If para.Alignment = wdAlignParagraphCenter Then
            ' Check 2: Does the entire paragraph (cleaned) ONLY contain the found text?
            refText = Trim(rngFound.text) ' Get the found text (e.g., "EXHIBIT A")
            If Trim(CleanParagraphText(para.Range.text)) = refText Then

                ' --- VERIFIED HEADING FOUND - Now find the Title ---
                Set nextPara = Nothing ' Reset
                On Error Resume Next   ' In case para.Next fails (e.g., last paragraph)
                Set nextPara = para.Next
                On Error GoTo 0        ' Turn error handling back on

                If Not nextPara Is Nothing Then ' Check if a next paragraph exists
                    Set rngTitle = nextPara.Range ' Get range of the next paragraph

                    ' Skip empty paragraphs immediately following the heading
                    Do While Len(Trim(Replace(rngTitle.text, Chr(13), ""))) = 0
                        Set para = nextPara ' Current 'next' becomes the new starting point
                        Set nextPara = Nothing ' Reset for next iteration
                        On Error Resume Next
                        Set nextPara = para.Next ' Try to get the one after that
                        On Error GoTo 0

                        If nextPara Is Nothing Then ' No more paragraphs after the empty one(s)
                            Set rngTitle = Nothing ' Indicate no title found
                            Exit Do              ' Exit the skipping loop
                        Else
                            Set rngTitle = nextPara.Range ' Update title range to the next non-empty candidate
                        End If
                    Loop

                    If Not rngTitle Is Nothing Then ' Check again after potentially skipping empty paras
                         titleText = CleanParagraphText(rngTitle.text)
                         If Len(titleText) > 0 Then
                             ' Add to dictionary (handles duplicates automatically - keeps first found)
                             If Not dictExhibits.Exists(refText) Then
                                 dictExhibits.Add refText, titleText
                             End If
                         End If ' End If Len(titleText) > 0
                    End If ' End If Not rngTitle Is Nothing (after skipping empty)
                End If ' End If Not nextPara Is Nothing (initial check)

            End If ' End If Text match check
        End If ' End If Alignment check

        ' Prepare for the next find iteration
        rngFound.Collapse wdCollapseEnd
        If rngFound.End < rngSearch.End Then rngFound.MoveStart wdCharacter, 1
        If rngFound.End >= rngSearch.End Then Exit Do

    Loop ' End Find loop

    ' --- Check if anything was found ---
    If dictExhibits.count = 0 Then
        MsgBox "No centered Exhibit headings matching the criteria were found.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' --- Prepare and Sort Exhibit Keys ---
    tempKeys = dictExhibits.Keys
    If IsArray(tempKeys) Then
        exhibitKeys = tempKeys
        BubbleSortStringArray exhibitKeys
    Else
        ReDim exhibitKeys(0 To 0)
        exhibitKeys(0) = tempKeys
    End If

    ' --- Insert the list at the original selection point ---
    startOfList.Select ' Go back to where the cursor was
    Selection.ParagraphFormat.SpaceBefore = 6
    Selection.ParagraphFormat.spaceAfter = 6

    ' Insert Exhibits
    For i = LBound(exhibitKeys) To UBound(exhibitKeys)
        Selection.TypeText exhibitKeys(i) & vbTab & dictExhibits(exhibitKeys(i))
        Selection.TypeParagraph ' New line for next item
    Next i

    Selection.Collapse wdCollapseEnd ' Deselect
    Application.ScreenUpdating = True
    MsgBox "Exhibit list generated.", vbInformation

End Sub

' --- Helper function to clean paragraph text (remove trailing paragraph mark and trim) ---
Private Function CleanParagraphText(pText As String) As String
    Dim cleanText As String
    cleanText = pText
    If Len(cleanText) > 0 Then
        Select Case Right(cleanText, 1)
            Case Chr(13), Chr(11) ' Paragraph mark or manual line break
                cleanText = Left(cleanText, Len(cleanText) - 1)
        End Select
    End If
    CleanParagraphText = Trim(cleanText)
End Function

' --- Helper function for basic string array sorting ---
Private Sub BubbleSortStringArray(arr() As Variant) ' Takes array ByRef implicitly
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim sorted As Boolean

    If Not IsArrayAllocated(arr) Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1
        sorted = True
        For j = LBound(arr) To UBound(arr) - 1 - (i - LBound(arr))
            If StrComp(CStr(arr(j)), CStr(arr(j + 1)), vbTextCompare) = 1 Then
                temp = arr(j + 1)
                arr(j + 1) = arr(j)
                arr(j) = temp
                sorted = False
            End If
        Next j
        If sorted Then Exit For
    Next i
End Sub

'--- Helper Function to check if a dynamic array is allocated ---
Private Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = False
    If IsArray(arr) Then
       IsArrayAllocated = (LBound(arr, 1) <= UBound(arr, 1))
    End If
    If Err.Number = 9 Then
        IsArrayAllocated = False
        Err.Clear
    ElseIf Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo 0
End Function
