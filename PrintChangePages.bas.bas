in file: word/vbaProject.bin - OLE stream: 'VBA/PrintChangePages'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub PrintOnlyMarkupPages()

If ActiveDocument.Revisions.count = 0 Then
    MsgBox "No revisions found."
    End
End If
    Dim strPages As String
    Dim currPg As String
    Dim currSec As String
    Dim oSec As Section
    Dim secPgs As PageNumbers
    Dim prevpg As String
    Dim currText As String
    strPages = "p1s1"
    prevpg = "1"
Selection.HomeKey (wdStory)
For Each myRev In ActiveDocument.Range.Revisions
    currPg = myRev.Range.Information(Word.WdInformation.wdActiveEndPageNumber)
    currSec = myRev.Range.Information(Word.WdInformation.wdActiveEndSectionNumber)
    currText = myRev.Range.text

'' I'm pretty sure all of this is unnecessary, but holding it just in case.''
'    If Len(currText) < 10 And "Article" = Left(currText, 7) Then
'        myRev.Accept
'        GoTo nextmyrev
'    ElseIf myRev.FormatDescription <> "" And Left(currText, 1) <> "([0-9]{1,})" Then
'        myRev.Accept
'        GoTo nextmyrev
'    ElseIf myRev.Range.Information(Word.WdInformation.wdInHeaderFooter) _
'        And currpg = "1" Then
'        myRev.Accept
'        GoTo nextmyrev
'    End If
    If prevpg <> currPg _
       And Not strPages Like "*" & currPg & "*" _
       And Not currPg = "1" Then 'add page to list, unless already there
        strPages = strPages & "," & "p" & currPg & "s" & currSec
    End If
nextmyrev:
prevpg = currPg
Next myRev
    'Print
    With Dialogs(wdDialogFilePrint)
       .Range = wdPrintRangeOfPages
       .Pages = strPages
       .Show
    End With
End Sub

Sub PrintOnlyMarkupPages2()

If ActiveDocument.Revisions.count = 0 Then
    MsgBox "No revisions found."
    Exit Sub
End If

Dim strPages As String
Dim currPg As Long
Dim currSec As Long
Dim prevpg As Long
Dim currSecPg As Long
Dim currText As String
Dim secStart As Long
Dim secEnd As Long

strPages = "p1s1"
prevpg = 1

Selection.HomeKey (wdStory)

For i = 1 To ActiveDocument.Range.Revisions.count
    myRev = ActiveDocument.Range.Revisions(i)
    If currPg = myRev.Range.Information(Word.WdInformation.wdActiveEndPageNumber) Then
        pgRevs = ActiveDocument.Range.
        i = ActiveDocument.Range.Revisions(myRev.Range.Information(Word.WdInformation.wdActiveEndPageNumber))
        Range.Revisions.
        GoTo nextmyrev
    Else
        currPg = myRev.Range.Information(Word.WdInformation.wdActiveEndPageNumber)
    End If
    currSec = myRev.Range.Information(Word.WdInformation.wdActiveEndSectionNumber)
    currText = myRev.Range.text
    
    With ActiveDocument.Sections(currSec)
        secStart = .Range.Characters.First.Information(Word.WdInformation.wdActiveEndPageNumber)
        secEnd = .Range.Characters.Last.Information(Word.WdInformation.wdActiveEndPageNumber)
    End With
    
    currSecPg = currPg - secStart + 1
    
    If prevpg <> currSecPg _
       And Not strPages Like "*p" & currSecPg & "s" & currSec & "*" _
       And Not currSecPg = 1 Then 'add page to list, unless already there
        strPages = strPages & "," & "p" & currSecPg & "s" & currSec
    End If
    
nextmyrev:
prevpg = currSecPg
Next i

'Print
With Dialogs(wdDialogFilePrint)
   .Range = wdPrintRangeOfPages
   .Pages = strPages
   .Show
End With

End Sub


-------------------------------------------------------------------------------
