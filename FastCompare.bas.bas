in file: word/vbaProject.bin - OLE stream: 'VBA/FastCompare'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub CheckboxTest()
MsgBox (SummaryCheckboxValue)
End Sub

Sub FastCompare()
Application.ScreenUpdating = False
Dim StrDocOld As String, DocOld As Document
Dim StrDocNew As String, DocNew As Document
Dim DocNewPath As String
Dim DocSummary As Shape
Dim IsBoxChecked As Boolean

Set DocNew = ActiveDocument
StrDocNew = DocNew.FullName
DocNewPath = DocNew.Path & "\"

' select "original" document to compare active document against
With Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
  .Title = "Select the Original Document"
  .AllowMultiSelect = False
  .Filters.Add "Documents", "*.doc; *.docx; *.docm", 1
  .InitialFileName = DocNewPath
  .ButtonName = "Compare"
  If .Show = -1 Then
    StrDocOld = .SelectedItems(1)
  Else
    Exit Sub
  End If
  If StrDocOld = StrDocNew Then
    MsgBox "The original document and revised document are the same.  Try again."
    Exit Sub
  End If
End With

Set DocOld = Documents.Open(StrDocOld)

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
    Call AddBox(DocOld, DocNew, DocRev)               'Someday, I will figureout how to get a checkbox  on the ribbon and control the summary box there.
End If

' save comparison with filename of DocNew with "-redline" appended
StrDocNew = Left(StrDocNew, (InStrRev(StrDocNew, ".", -1, vbTextCompare) - 1))
StrDocNew = StrDocNew & "-redline"

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
If StrDocOld <> "" Then DocOld.Close 'close DocOld
' Bring DocRev to front
DocRev.Activate
Application.ScreenUpdating = True
Set DocOld = Nothing: Set DocNew = Nothing: Set DocRev = Nothing
End Sub


Sub MTKFastCompare()
Application.ScreenUpdating = False
Dim StrDocOld As String, DocOld As Document
Dim StrDocNew As String, DocNew As Document
Dim DocNewPath As String
Set DocNew = ActiveDocument
StrDocNew = DocNew.FullName
DocNewPath = DocNew.Path & "\"

' select "original" document to compare active document against
With Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
  .Title = "Select the Original Document"
  .AllowMultiSelect = False
  .Filters.Add "Documents", "*.doc; *.docx; *.docm", 1
  .InitialFileName = DocNewPath
  .ButtonName = "Compare"
  If .Show = -1 Then
    StrDocOld = .SelectedItems(1)
  Else
    Exit Sub
  End If
  If StrDocOld = StrDocNew Then
    MsgBox "The original document and revised document are the same.  Try again."
    Exit Sub
  End If
End With

Set DocOld = Documents.Open(StrDocOld)

' run comparison
Dim DocRev As Document

Set DocRev = Application.CompareDocuments( _
  OriginalDocument:=DocOld, RevisedDocument:=DocNew, _
  Destination:=wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
  CompareFormatting:=False, CompareCaseChanges:=True, CompareWhitespace:=False, _
  CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True, _
  CompareTextboxes:=True, CompareFields:=False, CompareComments:=True, _
  CompareMoves:=True, RevisedAuthor:="Author", IgnoreAllComparisonWarnings:=False)

' *** save comparison with filename of DocNew with "-redline" appended\
    'Find the version number
    Dim sVer As String
    StrDocNew = Left(StrDocNew, (InStrRev(StrDocNew, ".", -1, vbTextCompare) - 1))
    StrDocOld = DocOld.Name
    DocOldNameOnly = Left(StrDocOld, InStrRev(StrDocOld, "(") - 1) 'Need shortened name to avoid conflict with dates
    vVer = Split(DocOldNameOnly, Chr(32))
    If (vVer(UBound(vVer))) = "" Then sVer = vVer(UBound(vVer) - 1) Else sVer = vVer(UBound(vVer))
    sVer = onlyDigits(sVer)
   
    DocOldLen = InStrRev(DocOldNameOnly, sVer, -1, vbTextCompare) - 1  'Find position of version number
    DocOldLen = (Len(StrDocOld) - DocOldLen) 'Find the position from the right of the version number
    StrDocOld = Right(StrDocOld, DocOldLen) 'Take the right side of the filename back to the version number

With Application.Dialogs(wdDialogFileSaveAs)
    .Name = StrDocNew & "-redline to v" & StrDocOld
    If .Show <> -1 Then GoTo ErrExit
End With

'Export to PDF
    TrackChangesOptions
    StrDocRev = Left(DocRev.Name, (InStrRev(DocRev.Name, ".", -1, vbTextCompare) - 1))
    DocRev.SaveAs2 FileName:=StrDocRev & ".pdf", _
       FileFormat:=wdFormatPDF
    DocRev.Save 'Have to save again for some reason.

ErrExit:
If StrDocOld <> "" Then DocOld.Close 'close DocOld
' Bring DocRev to front
DocRev.Activate
Application.ScreenUpdating = True
Set DocOld = Nothing: Set DocNew = Nothing: Set DocRev = Nothing
End Sub




Private Function onlyDigits(s As String) As String
'https://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string#7239408
'modified by Graham Mayor - https://www.gmayor.com - Last updated - 04 Apr 2021
' Variables needed (remember to use "option explicit").   '
Dim retval As String    ' This is the return string for the version number.      '
Dim i As Integer        ' Counter for character position. '

' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Or Mid(s, i, 1) = "." Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
       
    
End Function

Private Function TrackChangesOptions()

With Options
        'MARKUP group
        'Insertions
        .InsertedTextColor = wdClassicBlue
        'Deletions
        .DeletedTextColor = wdClassicRed
        
        'MOVES group
        'Moved from
        .MoveFromTextMark = wdMoveFromTextMarkDoubleStrikeThrough
        .MoveFromTextColor = wdGreen
        'Moved to
        .MoveToTextMark = wdMoveToTextMarkDoubleUnderline
        .MoveToTextColor = wdGreen
        
        'TABLE CELL HIGHLIGHTING group
        'Inserted cells
        .InsertedCellColor = wdCellColorLightBlue
        'Deleted cells
        .DeletedCellColor = wdCellColorPink
        'Merged cells
        .MergedCellColor = wdCellColorLightYellow
        'Split cells
        .SplitCellColor = wdCellColorLightOrange
    End With
End Function

Function AddBox(DocOld As Document, DocNew As Document, DocRev As Document)
'Add summary of comparison document to text box at top of first page
Dim BoxWidth As Long
Dim DocSummary As Shape

If Len(DocOld.Name) > Len(DocNew.Name) Then BoxWidth = Len(DocOld.Name) Else BoxWidth = Len(DocNew.Name)
DocRev.Activate
DocRev.TrackRevisions = False
Selection.HomeKey wdStory
Set DocSummary = ActiveDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Width:=BoxWidth * 7, Left:=25, top:=25, Height:=40) '
With DocSummary
    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .top = 25 ' Position from top of page
    .Left = 25 ' Position from left of page
        
    .Line.Visible = msoFalse
    .Shadow.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(255, 255, 255)
    With .TextFrame.TextRange
        .Style = wdStylePlainText
        .ParagraphFormat.Reset
        .ParagraphFormat.FirstLineIndent = 0
        .Font.Size = 8
        .Font.Name = "Tahoma"
        .InsertAfter "Redline" & Chr(13) & "Original:" & Chr(9) & DocOld.Name & Chr(13) & "Revised:" & Chr(9) & DocNew.Name 'how do I make "Redline" bold without a cumbersome for loop?
    End With

End With
End Function



-------------------------------------------------------------------------------
