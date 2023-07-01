Attribute VB_Name = "TermDefined"
Sub FindDef()
Dim wdFind As Find
Dim wdRng As Range
Dim wdDoc As Document
Dim wdSelect As Range
Dim openQuote As String

Set wdDoc = Application.ActiveDocument
Set wdRng = wdDoc.content
Set wdSelect = Selection.Range

'eliminate beginning and trailing spaces
With wdSelect
    .MoveEndWhile Chr(32), wdBackward
    .MoveStartWhile Chr(32)
End With

'Search parameters
With wdRng.Find
    .text = Chr(34) & wdSelect ' Somehow this returns both straight and curly quotes.  Not sure why but it works.
    '.Execute function returns TRUE/FALSE
    Searchresult = .Execute
End With

'Select entire sentence containing definition.
Dim wdDefinition As String
wdRng.Expand (wdSentence)
wdDefinition = wdRng

'Show message if searchresult = true
If Searchresult = True Then
    MsgBox Prompt:=wdDefinition, Title:="Yep, this is defined."
    Else: MsgBox "Nope, not defined."
End If

End Sub
