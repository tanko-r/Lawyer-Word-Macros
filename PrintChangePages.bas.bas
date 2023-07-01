in file: word/vbaProject.bin - OLE stream: 'VBA/PrintChangePages'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub PrintOnlyMarkupPages()

If ActiveDocument.Revisions.count = 0 Then
    MsgBox "No revisions found."
    End
End If
    Dim strPages As String
    Dim currpg As String
    Dim prevpg As String
    Dim currText As String
    strPages = "1"
    prevpg = "1"
Selection.HomeKey (wdStory)
For Each myRev In ActiveDocument.Range.Revisions
    currpg = myRev.Range.Information(Word.WdInformation.wdActiveEndPageNumber)
    currText = myRev.Range.text
    If Len(currText) < 10 And "Article" = Left(currText, 7) Then
        myRev.Accept
        GoTo nextmyrev
    ElseIf myRev.FormatDescription <> "" And Left(currText, 1) <> "([0-9]{1,})" Then
        myRev.Accept
        GoTo nextmyrev
    ElseIf myRev.Range.Information(Word.WdInformation.wdInHeaderFooter) _
        And currpg = "1" Then
        myRev.Accept
        GoTo nextmyrev
    End If
    If prevpg <> currpg _
       And Not strPages Like "*" & currpg & "*" _
       And Not currpg = "1" Then 'add page to list, unless already there
        strPages = strPages & "," & currpg
    End If
nextmyrev:
prevpg = currpg
Next myRev
    'Print
    With Dialogs(wdDialogFilePrint)
       .Range = wdPrintRangeOfPages
       .Pages = strPages
       .Show
    End With
End Sub

-------------------------------------------------------------------------------
