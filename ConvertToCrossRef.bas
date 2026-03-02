Attribute VB_Name = "ConvertToCrossRef"
Option Explicit

Sub ConvertToCrossRef()

    Dim RefList As Variant
    Dim LookUp As String
    Dim Ref As String
    Dim s As Integer, t As Integer
    Dim i As Integer
    Dim oRng As Range
    Dim oRngStr As String
    Dim spaceAfter As Boolean
    Dim numPart As String ' Variable to hold the extracted number part
    Dim delimiterPos As Integer ' Position of the first space or tab

    On Error GoTo CleanUp ' Changed error handler name for clarity

    Set oRng = Selection.Range
    oRngStr = Selection.Range.text

    ' Check if there was a space immediately after the original selection
    If Len(oRngStr) > 0 Then
        If oRng.Characters(Len(oRngStr)).text = Chr(32) Then spaceAfter = True
    End If

    ' --- Clean up the selected text ---
    With oRng
        ' Trim leading/trailing spaces from the range itself first
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32), wdForward ' More reliable way to trim leading spaces

        ' Refine trimming of trailing characters
        Do While .End > .Start
            Select Case Asc(Right(.text, 1))
                Case 13, 11, 32, 46 ' CR, VT, Space, Period
                    .MoveEnd wdCharacter, -1
                Case Else
                    Exit Do ' Stop if it's not one of the trailing chars we want to remove
            End Select
        Loop

        ' Check if selection is empty after cleaning
        If .End <= .Start Then GoTo ErrExitEmptySelection

        LookUp = .text ' Assign the cleaned text to LookUp
    End With
    ' --- End Selection Cleaning ---


    With ActiveDocument
        ' Use wdRefTypeNumberedItem to retrieve numbered paragraphs
        RefList = .GetCrossReferenceItems(wdRefTypeNumberedItem)

        If IsEmpty(RefList) Then GoTo ErrExitNoItems ' Check if any numbered items exist

        ' --- Loop through potential reference targets ---
        For i = 1 To UBound(RefList) ' Loop forward for clarity, index 'i' is needed later
            Ref = Trim(RefList(i)) ' Get the numbered item string, e.g., "3. Heading Text" or "3.2 Subheading"

            ' Find the position of the first space or tab, which separates the number from the text
            s = InStr(1, Ref, " ")
            t = InStr(1, Ref, Chr(9)) ' Chr(9) is Tab

            If s > 0 And t > 0 Then
                delimiterPos = IIf(s < t, s, t) ' Find the earlier delimiter
            ElseIf s > 0 Then
                delimiterPos = s ' Only space found
            ElseIf t > 0 Then
                delimiterPos = t ' Only tab found
            Else
                delimiterPos = 0 ' No delimiter found (might be just the number)
            End If

            ' Extract the number part
            If delimiterPos > 0 Then
                numPart = Trim(Left(Ref, delimiterPos - 1)) ' Get text before delimiter, trim spaces
            Else
                numPart = Ref ' The whole string is potentially the number part
            End If

            ' *** CORE FIX: Remove trailing period from the extracted number part if it exists ***
            If Right(numPart, 1) = "." Then
                numPart = Left(numPart, Len(numPart) - 1)
            End If

            ' Now compare the cleaned number part from the document (NumPart)
            ' with the cleaned selected text (LookUp)
            If StrComp(numPart, LookUp, vbTextCompare) = 0 Then ' Case-insensitive comparison
                ' Match found, exit the loop
                Exit For
            End If
        Next i
        ' --- End Loop ---


        ' --- Insert Cross Reference if match found ---
        If i <= UBound(RefList) Then ' Check if the loop completed because a match was found (i will be <= UBound)
            ' A match was found at index 'i'
            Selection.InsertCrossReference ReferenceType:="Numbered item", _
                                           ReferenceKind:=wdNumberFullContext, _
                                           ReferenceItem:=CStr(i), _
                                           InsertAsHyperlink:=True, _
                                           IncludePosition:=False, _
                                           SeparateNumbers:=False, _
                                           SeparatorString:=" "
            ' Add back the trailing space if it was originally present
            If spaceAfter Then Selection.Range.InsertAfter (Chr(32))
            ' Collapse selection after inserting
            Selection.Collapse wdCollapseEnd
        Else
            ' No match was found after checking all items
            MsgBox "A cross reference to """ & LookUp & """ couldn't be set." & vbCr & _
                   "A paragraph starting with that number" & vbCr & _
                   "couldn't be found in the document.", _
                   vbInformation, "Cross reference target not found"
        End If
        ' --- End Insert ---

    End With

    GoTo CleanUp ' Skip error handlers if successful

ErrExitEmptySelection:
    MsgBox "Please select a valid paragraph number reference.", _
           vbExclamation, "Invalid selection"
    GoTo CleanUp

ErrExitNoItems:
    MsgBox "There are no numbered items in this document to cross-reference.", _
            vbExclamation, "No Numbered Items Found"
    GoTo CleanUp

CleanUp:
    Set oRng = Nothing
    Set RefList = Nothing

End Sub

