Attribute VB_Name = "CompareToForm"
Option Explicit

Sub FastCompareToForm()
Application.ScreenUpdating = False
Dim StrDocOld As String, DocOld As Document
Dim StrDocNew As String, DocNew As Document
Dim DocNewPath As String
Dim DocSummary As Shape
Dim IsBoxChecked As Boolean

Set DocNew = ActiveDocument
StrDocNew = DocNew.FullName
DocNewPath = DocNew.Path & "\"

On Error Resume Next ' create an error trap because there's no "exists" function for variables.
                     ' https://www.askwoody.com/forums/topic/check-to-see-if-a-docvariable-exists-before-running-line-of-vba-code/
Dim varCheck As String
varCheck = ActiveDocument.Variables("formPath").Value
Debug.Print ActiveDocument.Variables("formPath").Value
Debug.Print Err.Number

If Err.Number = 0 Then
    StrDocOld = ActiveDocument.Variables("formPath").Value
Else
    MsgBox "Looks like this document was not created with the FKSDO Save As button, so you'll have to use the regular Fast Compare button and navigage to the form."
    On Error GoTo 0
    GoTo ErrExit
End If
On Error GoTo 0 ' Reset error handler

'Confirm that the user wants to run this redline.
If MsgBox("Compare to this document?" & vbCr & vbNewLine & StrDocOld, vbYesNoCancel) <> vbYes Then GoTo ErrExit

Set DocOld = Documents.Open(StrDocOld) 'If the formPath variable is set, then use that to open the form.

' run comparison
Dim DocRev As Document

Set DocRev = Application.CompareDocuments( _
  OriginalDocument:=DocOld, RevisedDocument:=DocNew, _
  Destination:=wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
  CompareFormatting:=False, CompareCaseChanges:=True, CompareWhitespace:=False, _
  CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True, _
  CompareTextboxes:=True, CompareFields:=False, CompareComments:=True, _
  CompareMoves:=True, RevisedAuthor:="Author", IgnoreAllComparisonWarnings:=False)

If SummaryCheckboxValue = True Then
    Call FastCompare.AddBox(DocOld, DocNew, DocRev)               'Someday, I will figureout how to get a checkbox  on the ribbon and control the summary box there.
End If

' save comparison with filename of DocNew with "-redline" appended
StrDocNew = Left(StrDocNew, (InStrRev(StrDocNew, ".", -1, vbTextCompare) - 1))
StrDocNew = StrDocNew & "-redline to form"

With Application.Dialogs(wdDialogFileSaveAs)
    .Name = StrDocNew
    If .Show <> -1 Then GoTo ErrExit
End With

'Export to PDF
'    TrackChangesOptions
'    StrDocRev = DocRev.Name
'    DocRev.SaveAs2 FileName:=StrDocNew & ".pdf", _
'       FileFormat:=wdFormatPDF
'    DocRev.Save 'Have to save again for some reason.
    
ErrExit:
If IsEmpty(DocOld) = True Then DocOld.Close 'close DocOld
' Bring DocRev to front
If IsEmpty(DocRev) = True Then 'Insane that "IsEmpty" returns "False" if the object IS EMPTY!?!?
    DocRev.Activate
End If
Application.ScreenUpdating = True
Set DocOld = Nothing: Set DocNew = Nothing: Set DocRev = Nothing
End Sub

