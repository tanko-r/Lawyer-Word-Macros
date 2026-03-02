Attribute VB_Name = "FastCompare"
Option Explicit
Public CurrentDocNewFullName As String

Sub FastCompare()
    Application.ScreenUpdating = False                                                          ' Prevent screen flicker and speed up macro execution

    Dim StrDocOld As String, DocOld As Document                                                 ' Variables for the original document
    Dim StrDocNew As String, DocNew As Document                                                 ' Variables for the revised (active) document
    Dim DocNewPath As String                                                                    ' Path of the revised document
    Dim StrNewCaption As String                                                                 ' Window caption of the revised document
    Dim StrOldName As String
    Dim StrNewName As String
    Dim iManageTrue As Boolean                                                                  ' Flag for iManage integration


    Set DocNew = ActiveDocument                                                                 ' Set DocNew to the currently active document
    StrNewCaption = ActiveWindow.Caption                                                        ' Get the caption of DocNew's window. This is kept for iManage functionality.
    StrDocNew = DocNew.FullName                                                                 ' Get the full path and name of DocNew
    DocNewPath = DocNew.Path & "\"                                                              ' Get the path of DocNew
    StrDocOld = DocNewPath                                                                      ' Preload StrDocOld with DocNew's path for FileDialog default

    ' --- This check remains unchanged and continues to use the window CAPTION ---
    If iManageHelpers.iManTestCaption(StrNewCaption) = True Then iManageTrue = True              ' Check if iManage is active (assumes iManageHelpers.iManTestCaption function exists)

    Dim DocNewFormat As String                                                                  ' Determine file extension for DocNew
    If DocNew.SaveFormat = 0 Then DocNewFormat = ".doc" Else DocNewFormat = ".docx"

    Dim frmSelectDoc As Object                                                                  ' UserForm object for selecting an open document; late binding for robustness
    Dim bDocOldSetFromList As Boolean                                                           ' Flag to track if DocOld was set from the UserForm list
    bDocOldSetFromList = False

    ' Use UserForm for selecting an already open document
    On Error Resume Next                                                                        ' Handle potential errors if UserForm doesn't exist or fails to load
    CurrentDocNewFullName = StrDocNew                                                           ' Set the CurrentDocNewFullname variable so it is excluded from the list
    Set frmSelectDoc = New frmSelectOpenDocument                                                ' ASSUMES UserForm named frmSelectOpenDocument exists
    
    If Err.Number = 0 Then                                                                      ' UserForm object created successfully
        On Error GoTo 0                                                                         ' Reset error handling for UserForm operations
        frmSelectDoc.Show vbModal                                                               ' Display the UserForm

        If frmSelectDoc.UserCancelled Then                                                      ' User cancelled the UserForm
            Unload frmSelectDoc
            Set frmSelectDoc = Nothing
            Application.ScreenUpdating = True
            Exit Sub
        End If

        ' If user selected a document from the list AND did not choose to browse
        If Not frmSelectDoc.BrowseFileSystem And frmSelectDoc.SelectedDocumentFullName <> "" Then
            StrDocOld = frmSelectDoc.SelectedDocumentFullName                                   ' Get selected document's full name

            On Error Resume Next                                                                ' Handle error if FullName isn't found in Application.Documents
            Set DocOld = Application.Documents(StrDocOld)                                       ' Attempt to get a reference to the already open document
            On Error GoTo 0

            If DocOld Is Nothing Then                                                           ' Failed to get reference to the open document
                MsgBox "Error: Could not get a reference to the selected open document: " & StrDocOld & vbCrLf & "Please try browsing the filesystem.", vbExclamation
                                                                                                ' DocOld remains Nothing, will trigger FileDialog path below
            Else
                bDocOldSetFromList = True                                                       ' Flag that DocOld was successfully set from the list
            End If
        End If
        ' If BrowseFileSystem is true, or no selection from list, or error getting DocOld,
        ' it will fall through to the FileDialog logic (bDocOldSetFromList is False or DocOld is Nothing).

        Unload frmSelectDoc                                                                     ' Clean up UserForm
        Set frmSelectDoc = Nothing
    Else
        On Error GoTo 0                                                                         ' Reset error handling
        MsgBox "Note: The open documents list feature could not be initialized." & vbCrLf & "Please use the file browser to select the original document.", vbInformation
                                                                                                ' Fall through to FileDialog logic
    End If

    ' If DocOld was not set from the UserForm list (or UserForm failed/user chose browse)
    If Not bDocOldSetFromList Then
        With Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)                    ' Open FileDialog to select the original document
            .title = "Select the Original Document (from Filesystem)"
            .AllowMultiSelect = False
            .Filters.Add "Documents", "*.doc; *.docx; *.docm", 1
            .InitialFileName = StrDocOld                                                        ' Default to DocNew's path
            .ButtonName = "Compare"
            If .Show = -1 Then                                                                  ' User selected a file
                StrDocOld = .SelectedItems(1)                                                   ' Get the selected file's path
            Else
                Application.ScreenUpdating = True                                               ' User cancelled FileDialog
                Exit Sub
            End If
        End With

        Dim tempDoc As Document                                                                 ' Check if the document selected via FileDialog is already open
        Set tempDoc = Nothing
        On Error Resume Next                                                                    ' Handle if StrDocOld isn't an exact match for an open doc's FullName
        Set tempDoc = Application.Documents(StrDocOld)                                          ' Attempt to get reference if already open
        On Error GoTo 0

        If Not tempDoc Is Nothing Then
            Set DocOld = tempDoc                                                                ' Use the already open document reference
        Else
            On Error Resume Next                                                                ' Handle error during Document.Open
            Set DocOld = Documents.Open(StrDocOld)                                              ' Open the document selected from file system
            If Err.Number <> 0 Then
                MsgBox "Error opening selected document: " & StrDocOld & vbCrLf & Err.Description, vbExclamation
                On Error GoTo 0
                Application.ScreenUpdating = True
                Exit Sub
            End If
            On Error GoTo 0                                                                     ' Reset error handling
        End If
    End If

    ' Final check to ensure DocOld is set
    If DocOld Is Nothing Then
        MsgBox "The original document (DocOld) could not be determined or opened. Comparison cannot proceed.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If

    StrDocOld = DocOld.FullName                                                                 ' Ensure StrDocOld reflects the actual FullName of the DocOld object

    If StrDocOld = StrDocNew Then                                                               ' Check if original and revised documents are the same
        MsgBox "The original document and revised document are the same. Please try again."
        Application.ScreenUpdating = True
        Exit Sub                                                                                ' Exits, DocOld.Close will be called in ErrExit
    End If

    StrNewName = DocNew.Name
    StrOldName = DocOld.Name
    
    Dim DocRev As Document                                                                      ' Variable for the comparison result document

    Set DocRev = Application.CompareDocuments( _
        OriginalDocument:=DocOld, RevisedDocument:=DocNew, _
        Destination:=wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=False, CompareCaseChanges:=True, CompareWhitespace:=False, _
        CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True, _
        CompareTextboxes:=True, CompareFields:=False, CompareComments:=True, _
        CompareMoves:=True, RevisedAuthor:="Author", IgnoreAllComparisonWarnings:=False)

    ' Clean up document filenames for display
    If InStr(StrNewName, ".docx") Then
        StrNewName = Left(StrNewName, InStrRev(StrNewName, ".docx") - 1)
    End If
    If InStr(StrNewName, ".doc") Then
        StrNewName = Left(StrNewName, InStrRev(StrNewName, ".doc") - 1)
    End If
    If InStr(StrOldName, ".docx") Then
        StrOldName = Left(StrOldName, InStrRev(StrOldName, ".docx") - 1)
    End If
    If InStr(StrOldName, ".doc") Then
        StrOldName = Left(StrOldName, InStrRev(StrOldName, ".doc") - 1)
    End If
    
    ' The custom cleanup function is now called with the filenames.
    ' You may need to adjust the function if it's designed specifically for caption formatting.
    StrNewName = iManPolsinelliCaptionCleanup(StrNewName)
    StrOldName = iManPolsinelliCaptionCleanup(StrOldName)
    
    Call AddBox(StrOldName, StrNewName, DocRev)

    DocRev.Activate                                                                             ' Activate the comparison document (may be needed for iManage SaveAs)

    'Call the SaveAs dialog window and default it to the user's Downloads folder
    With Application.Dialogs(wdDialogFileSaveAs)
        .Name = Environ("USERPROFILE") & "\Downloads\" & StrNewName & "-redline"
        If .Show = -1 Then
             ' User clicked Save, do nothing more here as the dialog handles the save.
        Else
            ' If user cancels SaveAs, save a temporary copy and notify them.
             DocRev.SaveAs2 FileName:=Environ("TEMP") & "\" & StrNewName & "-redline.docx"
        End If
        GoTo ErrExit
    End With

ErrExit:                                                                                        ' Error handling and cleanup
    If Not DocOld Is Nothing Then                                                               ' Check if DocOld object exists
        If DocOld.Saved = False Then DocOld.Saved = True                                        ' Suppress save prompt if DocOld was not modified by this macro
        If StrDocOld <> "" Then                                                                 ' Close the original document if it was opened from the filesystem
            If bDocOldSetFromList = False Then
                DocOld.Close
            Else
                Dim closeOld As VbMsgBoxResult
                closeOld = MsgBox(prompt:="Close the old version?", _
                      Buttons:=vbYesNo, _
                      title:="Close Prior Version")
                If closeOld = vbYes Then DocOld.Close
            End If
        End If
    End If

    DocNew.Activate                                                                             ' Bring revised document to front
    If Not DocRev Is Nothing Then DocRev.Activate                                               ' Bring comparison document to front if it exists
    Application.ScreenUpdating = True                                                           ' Re-enable screen updating

    Set DocOld = Nothing                                                                        ' Release object variables
    Set DocNew = Nothing
    Set DocRev = Nothing
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

Private Function AddBox(StrOldCaption As String, StrNewCaption As String, DocRev As Document)
'Add summary of comparison document to text box at top of first page
Dim BoxWidth As Long
Dim DocSummary As Shape

If Len(StrOldCaption) > Len(StrNewCaption) Then BoxWidth = Len(StrOldCaption) Else BoxWidth = Len(StrNewCaption)

DocRev.Activate
DocRev.TrackRevisions = False
Selection.HomeKey wdStory
Set DocSummary = ActiveDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, Width:=BoxWidth * 4, Left:=11, Top:=11, Height:=40)
With DocSummary
    .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    .RelativeVerticalPosition = wdRelativeVerticalPositionPage
    .Top = 11 ' Position from top of page
    .Left = 11 ' Position from left of page
        
    .Line.Visible = msoTrue
    .Shadow.Visible = msoFalse
    .Fill.ForeColor.RGB = RGB(255, 255, 255)
    With .TextFrame.TextRange
        .Style = wdStylePlainText
        .ParagraphFormat.Reset
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.spaceAfter = 0
        .ParagraphFormat.spaceAfter = 0
        .Font.Size = 8
        .Font.Name = "Tahoma"
        .InsertAfter "Redline" & Chr(13) & "Original:" & Chr(9) & StrOldCaption & Chr(13) & "Revised:" & Chr(9) & StrNewCaption 'how do I make "Redline" bold without a cumbersome for loop?
    End With
    
    .TextFrame.TextRange.Paragraphs(1).Range.Font.Bold = True ' Bold "Redline"
End With
' AutoFit the textbox to its content
    With DocSummary
        .TextFrame.WordWrap = msoFalse  'Note: This must be set to false to make autosizing work.
        .TextFrame.AutoSize = 1         'NOTE: This must be set as a numerical value. Per Office VBA docs, this associates with msoAutoSizeShapeToFitText
    End With
End Function

