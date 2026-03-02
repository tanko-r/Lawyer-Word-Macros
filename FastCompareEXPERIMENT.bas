Attribute VB_Name = "FastCompareEXPERIMENT"
'Option Explicit
'
'' =========================================================================
'' Main Subroutine: FastCompare
'' Purpose: Compares the active document (revised) against a user-selected
''          original document, generates a new "redline" comparison document,
''          adds a summary box, and prompts to save the redline.
''          Includes logic potentially related to iManage DMS.
'' Source:  Google Gemini reorganized and commented the original Fast Compare macro and helper functions for readability.
'' =========================================================================
'Sub FastCompare()
'
'    ' --- Variable Declarations ---
'    Dim DocNew As Document      ' Represents the currently active (revised) document
'    Dim DocOld As Document      ' Represents the original document selected by the user
'    Dim DocRev As Document      ' Represents the newly created comparison (redline) document
'    Dim DocSummary As Shape     ' Represents the summary text box added to the redline doc
'
'    Dim StrDocNew As String     ' Full path and filename of the revised document
'    Dim StrDocOld As String     ' Full path and filename of the original document
'    Dim StrNewCaption As String ' Window caption of the revised document (may include iManage info)
'    Dim StrOldCaption As String ' Window caption of the original document (may include iManage info)
'    Dim DocNewPath As String    ' Path (folder) of the revised document
'    Dim DocNewFormat As String  ' File extension (.doc or .docx) of the revised document
'
'    ' Note: These variables are declared but not used in the provided snippet.
'    ' Dim StrDocNewId As String   ' Potentially for iManage Document ID
'    ' Dim StrDocNewVer As String  ' Potentially for iManage Document Version
'    ' Dim regex As Object         ' Potentially for regular expression operations (not used here)
'
'    Dim iManageTrue As Boolean  ' Flag indicating if the document seems to be managed by iManage
'    Dim IsBoxChecked As Boolean ' Declared but not used in the provided snippet. Likely for future use with a checkbox.
'
'    ' --- Setup & Performance ---
'    ' Turn off screen updating to speed up macro execution and prevent flickering
'    Application.ScreenUpdating = False
'
'    ' --- Identify Revised Document ---
'    Set DocNew = ActiveDocument          ' Get the currently open document
'    StrNewCaption = ActiveWindow.Caption ' Get the caption from the window title (might contain iManage info)
'    StrDocNew = DocNew.FullName          ' Get the full path and filename
'    DocNewPath = DocNew.Path & "\"       ' Extract the path and add a trailing slash
'
'    ' Preload the initial path for the file picker to the revised document's folder
'    StrDocOld = DocNewPath
'
'    ' --- Check for iManage Integration (Assumes iManageHelpers object exists) ---
'    ' This line checks if the caption indicates an iManage document.
'    ' Requires an external object or function named 'iManageHelpers' with an 'iManTestCaption' method.
'    ' If iManageHelpers.iManTestCaption(StrNewCaption) = True Then iManageTrue = True
'    ' --- Simplified iManage Check (Example - replace with your actual logic if needed) ---
'    ' A simpler check might just look for typical iManage patterns in the caption
'    If InStr(1, StrNewCaption, ".DOC - #") > 0 Or InStr(1, StrNewCaption, ".DOCX - #") > 0 Then
'        iManageTrue = True
'    Else
'        iManageTrue = False
'    End If
'
'    ' Determine the file format extension for later use
'    If DocNew.SaveFormat = 0 Then DocNewFormat = ".doc" Else DocNewFormat = ".docx"
'
'    ' --- Prompt User to Select Original Document ---
'    With Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
'        .title = "Select the Original Document" ' Set dialog title
'        .AllowMultiSelect = False              ' Only allow one file selection
'        .Filters.Clear                         ' Clear existing filters
'        .Filters.Add "Documents", "*.doc; *.docx; *.docm", 1 ' Add filter for Word documents
'        .InitialFileName = StrDocOld           ' Start Browse in the revised doc's folder
'        .ButtonName = "Compare"                ' Customize the button text
'
'        ' Show the dialog and check if the user selected a file (-1 means OK was clicked)
'        If .Show = -1 Then
'            StrDocOld = .SelectedItems(1) ' Get the full path of the selected file
'        Else
'            GoTo Cleanup ' Exit the subroutine cleanly
'        End If
'
'        ' Check if the user selected the same file as the revised document
'        If StrDocOld = StrDocNew Then
'            MsgBox "The original document and revised document are the same. Please select a different original document.", vbExclamation
'            GoTo Cleanup ' Exit the subroutine cleanly
'        End If
'    End With
'
'    ' --- Open Original Document ---
'    ' Error handling in case the selected file can't be opened
'    On Error Resume Next ' Temporarily ignore errors
'    Set DocOld = Documents.Open(StrDocOld, ReadOnly:=True) ' Open the selected original document read-only
'    If Err.Number <> 0 Then
'        MsgBox "Error opening the original document: " & vbCrLf & StrDocOld & vbCrLf & Err.Description, vbCritical
'        On Error GoTo 0 ' Restore default error handling
'        GoTo Cleanup ' Exit the subroutine cleanly
'    End If
'    On Error GoTo 0 ' Restore default error handling
'
'    DocOld.Activate ' Bring the original document window to the front briefly
'    StrOldCaption = ActiveWindow.Caption ' Get its window caption
'
'    ' --- Perform Document Comparison ---
'    ' Use Word's built-in compare feature
'    ' Note: RevisedAuthor is hardcoded. Consider making this dynamic or removing it.
'    Set DocRev = Application.CompareDocuments( _
'        OriginalDocument:=DocOld, _
'        RevisedDocument:=DocNew, _
'        Destination:=wdCompareDestinationNew, _
'        Granularity:=wdGranularityWordLevel, _
'        CompareFormatting:=False, _
'        CompareCaseChanges:=True, _
'        CompareWhitespace:=False, _
'        CompareTables:=True, _
'        CompareHeaders:=True, _
'        CompareFootnotes:=True, _
'        CompareTextboxes:=True, _
'        CompareFields:=False, _
'        CompareComments:=True, _
'        CompareMoves:=True, _
'        RevisedAuthor:="Author", _
'        IgnoreAllComparisonWarnings:=False)
'
'    ' DocRev now holds the new document showing the differences
'
'    ' --- Prepare Captions for Summary Box ---
'    ' If using iManage, potentially strip version info (e.g., "#12345.1") from the caption
'    If iManageTrue Then
'        If InStrRev(StrNewCaption, "#") > 0 Then
'           StrNewCaption = Trim(Left(StrNewCaption, InStrRev(StrNewCaption, "#") - 1))
'        End If
'         If InStrRev(StrOldCaption, "#") > 0 Then
'           StrOldCaption = Trim(Left(StrOldCaption, InStrRev(StrOldCaption, "#") - 1))
'        End If
'    End If
'
'    ' --- Add Summary Box to Redline Document ---
'    ' Calls a helper function to add a text box with comparison details
'    Call AddBox(StrOldCaption, StrNewCaption, DocRev)
'
'    ' --- Prepare Filename for Saving & Clipboard ---
'    Dim suggestedFilename As String
'    Dim obj As Object ' Using late binding for MSForms.DataObject
'
'    ' Create the suggested filename for the redline document
'    If iManageTrue Then
'        ' Use the (potentially stripped) caption for iManage docs
'        suggestedFilename = StrNewCaption & "-redline"
'    Else
'        ' Use the original filename (without extension) for non-iManage docs
'        If InStrRev(StrDocNew, ".") > 0 Then
'            suggestedFilename = Left(StrDocNew, InStrRev(StrDocNew, ".") - 1) & "-redline"
'        Else
'            suggestedFilename = StrDocNew & "-redline" ' Fallback if no extension
'        End If
'    End If
'
''    ' Copy the suggested filename to the clipboard (useful for pasting into Save As dialog)
''    ' Requires reference to 'Microsoft Forms 2.0 Object Library' or use late binding
''    On Error Resume Next ' In case DataObject isn't available
''    Set obj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' CLSID for DataObject
''    If Err.Number = 0 Then
''        obj.SetText suggestedFilename
''        obj.PutInClipboard
''    End If
''    On Error GoTo 0 ' Restore error handling
''    Set obj = Nothing ' Release DataObject
'
'    ' --- Save the Comparison Document ---
'    DocRev.Activate ' Ensure the comparison document is active before saving
'
'    ' Use the FileSaveAs dialog, pre-populating the filename
'    With Application.Dialogs(wdDialogFileSaveAs)
'        .Name = suggestedFilename ' Set the initial filename in the dialog
'        ' Show the dialog. If user cancels (.Show <> -1), jump to cleanup.
'        If .Show = -1 Then
'             ' Save successful or handled by user/iManage connector
'             ' No specific action needed here unless further processing depends on save path
'        Else
'            ' MsgBox "Save cancelled. The redline document remains open but unsaved.", vbInformation
'            ' Decide if you want to close the unsaved DocRev here or leave it open
'            ' Example: DocRev.Close SaveChanges:=wdDoNotSaveChanges
'            GoTo Cleanup ' Or just let it fall through to cleanup
'        End If
'    End With
'
'    ' --- (Commented Out) Optional: Export to PDF ---
'    ' This section was commented out in the original code.
'    ' To enable it, you would typically call TrackChangesOptions first
'    ' and then save as PDF.
'    ' Call TrackChangesOptions ' Apply desired track changes appearance (if needed for PDF)
'    ' Dim StrDocRev As String
'    ' StrDocRev = DocRev.Name ' Or DocRev.FullName if you need the path
'    ' DocRev.SaveAs2 FileName:=suggestedFilename & ".pdf", FileFormat:=wdFormatPDF
'    ' DocRev.Save ' May need to save the Word doc again after PDF export if changes occurred
'
'' --- Cleanup ---
'Cleanup:
'    ' Close the original document if it was opened successfully
'    If Not DocOld Is Nothing Then
'        If DocOld.Saved = False Then ' Optional: Check if needs saving (unlikely as opened ReadOnly)
'            ' Add logic here if needed, e.g., prompt user
'        End If
'        DocOld.Close SaveChanges:=wdDoNotSaveChanges ' Close without saving changes
'    End If
'
'    If Not DocRev Is Nothing Then DocRev.Activate ' Show the revised document to the user.
'
'    ' Bring the relevant windows to the front (optional, adjust as needed)
'    ' On Error Resume Next ' In case documents were closed manually or due to errors
'    ' If Not DocNew Is Nothing Then DocNew.Activate
'    ' If Not DocRev Is Nothing Then DocRev.Activate
'    ' On Error GoTo 0
'
'    ' Restore screen updating
'    Application.ScreenUpdating = True
'
'    ' Release object variables
'    Set DocOld = Nothing
'    Set DocNew = Nothing
'    Set DocRev = Nothing
'    Set DocSummary = Nothing
'
'    ' Optional: Notify user of completion
'    ' MsgBox "Comparison process complete.", vbInformation
'
'End Sub
'
'' =========================================================================
'' Helper Function: AddBox
'' Purpose: Adds a text box to the top-left of the specified document
''          containing details about the original and revised files.
''          Calculates width based on the longest filename line using
''          direct measurement, applies fixed tabs, and adjusts height.
''          Attempts to avoid overlapping main document text using text wrapping.
'' Parameters:
''   StrOldCaption: Caption/name of the original document.
''   StrNewCaption: Caption/name of the revised document.
''   DocRev: The document object (comparison document) to add the box to.
'' Returns:
''   Boolean: True if successful, False otherwise
'' =========================================================================
'Function AddBox(StrOldCaption As String, StrNewCaption As String, DocRev As Document) As Boolean
'    Dim DocSummary As Shape      ' The text box shape object
'    Dim tempRng As Range         ' Temporary range for measurement
'    Dim measuredWidth As Single  ' Measured width of the longest line
'    Dim finalWidth As Single
'    Dim finalHeight As Single
'    Dim lineCount As Long
'    Dim estLineHeight As Single
'    Dim headerLine As String
'    Dim originalLine As String
'    Dim revisedLine As String
'    Dim longestDataLine As String
'
'    ' --- Constants ---
'    ' Position
'    Const BOX_TOP As Single = 25    ' Distance from top of page in points
'    Const BOX_LEFT As Single = 25   ' Distance from left of page in points
'    ' Sizing & Formatting
'    Const FONT_SIZE As Single = 8
'    Const FONT_NAME As String = "Tahoma"
'    Const WIDTH_MARGIN As Single = 18  ' Extra horizontal space (padding) total (9 points each side approx)
'    Const HEIGHT_MARGIN As Single = 6   ' Extra vertical space (padding) total (3 points top/bottom approx)
'    Const FIXED_TAB_POS As Single = 9 ' Tab position in points (0.125 inches * 72 points/inch)
'    ' Initial size guess (will be overridden)
'    Const INITIAL_WIDTH As Single = 150
'    Const INITIAL_HEIGHT As Single = 50
'
'    On Error GoTo ErrorHandler ' Basic error handling
'
'    ' Ensure the target document is active
'    If DocRev Is Nothing Then GoTo ErrorHandler
'    DocRev.Activate
'
'    ' --- Prepare Text Lines ---
'    headerLine = "Redline Comparison Summary"
'    originalLine = "Original:" & vbTab & StrOldCaption
'    revisedLine = "Revised:" & vbTab & StrNewCaption
'
'    ' Determine which data line is likely longer (simple length check is usually sufficient)
'    If Len(originalLine) > Len(revisedLine) Then
'        longestDataLine = originalLine
'    Else
'        longestDataLine = revisedLine
'    End If
'    ' Also consider the header line for width calculation
'     If Len(headerLine) > Len(longestDataLine) Then
'        longestDataLine = headerLine ' In case the header is exceptionally long
'     End If
'
'
'    ' Temporarily disable track revisions while adding the shape
'    Dim trackRevisionsState As Boolean
'    trackRevisionsState = DocRev.TrackRevisions
'    DocRev.TrackRevisions = False
'
'    ' Add the text box shape with initial placeholder dimensions
'    Set DocSummary = DocRev.Shapes.AddTextbox( _
'        Orientation:=msoTextOrientationHorizontal, _
'        Left:=BOX_LEFT, _
'        Top:=BOX_TOP, _
'        Width:=INITIAL_WIDTH, _
'        Height:=INITIAL_HEIGHT)
'
'    ' --- Format Shape Appearance and Position ---
'    With DocSummary
'        .Name = "ComparisonSummaryBox" ' Give it a unique name
'        ' Anchoring relative to the page
'        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
'        .Top = BOX_TOP
'        .Left = BOX_LEFT
'        .LockAnchor = True
'        .WrapFormat.Type = wdWrapTopBottom ' Avoid overlap with main text flow
'
'        ' Appearance
'        .Line.Visible = msoTrue
'        .Line.ForeColor.RGB = RGB(0, 0, 0)
'        .Line.Weight = 0.75
'        .Shadow.Visible = msoFalse
'        .Fill.Visible = msoTrue
'        .Fill.ForeColor.RGB = RGB(255, 255, 255)
'        .Fill.Transparency = 0
'
'        ' --- Text Frame Setup ---
'        With .TextFrame
'            ' Set internal margins *before* measurement/text insertion
'            .MarginLeft = PointsToInches(WIDTH_MARGIN / 2) ' Convert points to inches for these properties
'            .MarginRight = PointsToInches(WIDTH_MARGIN / 2)
'            .MarginTop = PointsToInches(HEIGHT_MARGIN / 2)
'            .MarginBottom = PointsToInches(HEIGHT_MARGIN / 2)
'            .WordWrap = True ' Ensure text wraps if needed (shouldn't if width calc is right)
'            .VerticalAnchor = msoAnchorTop ' Align text to the top
'
'            ' --- Measure Required Width ---
'            ' Use a temporary range *within the textbox* to measure the longest line
'            Set tempRng = .TextRange
'            tempRng.text = longestDataLine ' Insert only the longest line for measurement
'            With tempRng
'                .Font.Name = FONT_NAME
'                .Font.Size = FONT_SIZE
'                .Font.Bold = False ' Ensure not bold for measurement unless header IS longest
'                .ParagraphFormat.Reset
'                .ParagraphFormat.SpaceBefore = 0
'                .ParagraphFormat.spaceAfter = 0
'                .ParagraphFormat.LeftIndent = 4
'                .ParagraphFormat.FirstLineIndent = 4
'                ' Add the fixed tab stop *if* measuring a data line
'                If InStr(longestDataLine, vbTab) > 0 Then
'                    .ParagraphFormat.TabStops.ClearAll
'                    .ParagraphFormat.TabStops.Add Position:=FIXED_TAB_POS, _
'                                                 Alignment:=wdAlignTabLeft, _
'                                                 Leader:=wdTabLeaderSpaces
'                Else
'                    .ParagraphFormat.TabStops.ClearAll ' Header line has no tabs
'                End If
'
'                ' Perform measurement using Selection (unavoidable for reliable width)
'                 measuredWidth = 0
'                 On Error Resume Next ' Handle cases where selection might fail
'                 .Select ' Select the temporary text
'                 measuredWidth = Selection.Information(wdHorizontalPositionRelativeToTextBoundary)
'                 Selection.Collapse wdCollapseStart ' Deselect
'                 On Error GoTo ErrorHandler ' Restore normal error handling
'
'                 ' Fallback if measurement failed (crude estimate)
'                 If measuredWidth <= 0 Then
'                    measuredWidth = Len(longestDataLine) * FONT_SIZE * 0.5 ' Estimate
'                 End If
'
'            End With ' End With tempRng
'
'            ' Calculate final width including internal margins
'            finalWidth = measuredWidth + .MarginLeft * 72 + .MarginRight * 72 ' Add internal margins IN POINTS
'
'            ' --- Insert and Format Final Text ---
'            Set tempRng = .TextRange ' Reset tempRng to the whole text frame
'            tempRng.text = headerLine & vbCrLf & originalLine & vbCrLf & revisedLine
'
'            ' Format the whole text range defaults first
'            With tempRng
'                .Font.Name = FONT_NAME
'                .Font.Size = FONT_SIZE
'                .Font.Bold = False
'                .ParagraphFormat.Reset
'                .ParagraphFormat.SpaceBefore = 0
'                .ParagraphFormat.spaceAfter = 3 ' Small space between lines
'                .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
'                .ParagraphFormat.Alignment = wdAlignParagraphLeft
'                .ParagraphFormat.LeftIndent = 0
'                .ParagraphFormat.FirstLineIndent = 0
'                .ParagraphFormat.TabStops.ClearAll ' Clear any previous tabs
'            End With
'
'            ' Format Header (Paragraph 1)
'            With .TextRange.Paragraphs(1)
'                .Range.Font.Bold = True
'                ' No tabs or indents needed (already reset)
'            End With
'
'            ' Format Data Lines (Paragraphs 2 and 3)
'            Dim i As Long
'            For i = 2 To 3
'                With .TextRange.Paragraphs(i)
'                     ' No indents needed
'                    .Format.TabStops.ClearAll ' Ensure clean slate
'                    .Format.TabStops.Add Position:=FIXED_TAB_POS, _
'                                         Alignment:=wdAlignTabLeft, _
'                                         Leader:=wdTabLeaderSpaces
'                End With
'            Next i
'
'            ' --- Calculate Required Height ---
'             lineCount = .TextRange.ComputeStatistics(wdStatisticLines)
'             ' Estimate line height (Font size + approximate leading)
'             estLineHeight = FONT_SIZE + 3 ' Adjust '+3' based on visual results
'             finalHeight = (lineCount * estLineHeight) + .MarginTop + .MarginBottom ' Add internal margins (in points)
'
'             ' --- Apply Calculated Size ---
'             DocSummary.Width = finalWidth
'             DocSummary.Height = finalHeight ' Set calculated height
'
'              ' Optional: Disable AutoSize to prevent Word interfering later
'              .AutoSize = msoAutoSizeNone
'
'        End With ' End With .TextFrame
'
'    End With ' End With DocSummary
'
'
'    ' Restore the original track revisions state
'    DocRev.TrackRevisions = trackRevisionsState
'
'    ' Deselect the shape if it was selected
'    Selection.HomeKey Unit:=wdStory
'    Selection.Collapse Direction:=wdCollapseStart
'
'    AddBox = True ' Indicate success
'    GoTo Cleanup
'
'ErrorHandler:
'    MsgBox "An error occurred in AddBox: " & Err.Description & vbCrLf & _
'           "(Error " & Err.Number & ")", vbExclamation, "AddBox Error"
'    AddBox = False ' Indicate failure
'    ' Attempt to restore track revisions even on error
'    On Error Resume Next ' Prevent error during cleanup
'    If Not DocRev Is Nothing Then
'        ' Check if trackRevisionsState was successfully captured before trying to restore
'        If TypeName(trackRevisionsState) = "Boolean" Then
'             If DocRev.TrackRevisions <> trackRevisionsState Then
'                 DocRev.TrackRevisions = trackRevisionsState
'             End If
'        Else
'             ' If state wasn't captured, maybe just turn it off if it's on? Risky.
'             ' Best to leave it as is if the initial state is unknown.
'        End If
'    End If
'    On Error GoTo 0 ' Resume normal error handling
'
'Cleanup:
'    ' Release object variables
'    Set DocSummary = Nothing
'    Set tempRng = Nothing
'    On Error GoTo 0 ' Ensure error handling is reset
'
'End Function
'
'' Helper function to convert points to inches for TextFrame margins
'Private Function PointsToInches(points As Single) As Single
'    PointsToInches = points / 72
'End Function
'
'' =========================================================================
'' Helper Function: onlyDigits
'' Purpose: Extracts only numeric digits (0-9) and periods (.) from a string.
'' Source: Modified version from Stack Overflow/Graham Mayor
'' Parameters:
''   s: The input string.
'' Returns: A string containing only the digits and periods from the input.
'' =========================================================================
'Private Function onlyDigits(s As String) As String
'    Dim retval As String ' String to build the result
'    Dim i As Integer     ' Loop counter for character position
'
'    retval = "" ' Initialize return value
'
'    ' Loop through each character in the input string
'    For i = 1 To Len(s)
'        ' Check if the character is a digit (0-9) or a period
'        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Or Mid(s, i, 1) = "." Then
'            ' Append the digit or period to the result string
'            retval = retval + Mid(s, i, 1)
'        End If
'    Next i
'
'    ' Return the resulting string of digits and periods
'    onlyDigits = retval
'
'End Function
'
'
'' =========================================================================
'' Helper Function: TrackChangesOptions
'' Purpose: Sets the visual options (colors, markings) for how tracked
''          changes are displayed in Word.
'' Note: This function is defined but *not called* in the FastCompare sub.
''       It would need to be called (e.g., before saving or PDF export)
''       if these specific settings are desired for the output.
'' =========================================================================
'Private Function TrackChangesOptions()
'
'    With Options
'        ' --- Markup Group ---
'        ' Insertions: How new text appears
'        .InsertedTextMark = wdInsertedTextMarkUnderline ' (Example, default is often color only)
'        .InsertedTextColor = wdBlue ' Using wdBlue constant
'
'        ' Deletions: How deleted text appears
'        .DeletedTextMark = wdDeletedTextMarkStrikeThrough ' (Example, default is often color only)
'        .DeletedTextColor = wdRed   ' Using wdRed constant
'
'        ' --- Moves Group (If Track Moves is enabled) ---
'        ' Moved from: How text that was moved appears at original location
'        .MoveFromTextMark = wdMoveFromTextMarkDoubleStrikeThrough
'        .MoveFromTextColor = wdGreen
'
'        ' Moved to: How text that was moved appears at new location
'        .MoveToTextMark = wdMoveToTextMarkDoubleUnderline
'        .MoveToTextColor = wdGreen
'
'        ' --- Formatting Changes (If Track Formatting is enabled) ---
'        ' .FormattingMark = wdFormattingMarkBold ' (Example)
'        ' .FormattingColor = wdColorIndexViolet ' (Example)
'
'        ' --- Changed Lines (Vertical lines in margin) ---
'        ' .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve ' (Example)
'        ' .RevisedLinesMark = wdRevisedLinesMarkLeftBorder ' (Example)
'        ' .RevisedLinesColor = wdColorIndexAuto ' (Example)
'
'        ' --- Table Cell Highlighting ---
'        .InsertedCellColor = wdCellColorLightBlue
'        .DeletedCellColor = wdCellColorPink
'        .MergedCellColor = wdCellColorLightYellow
'        .SplitCellColor = wdCellColorLightOrange
'
'        ' --- Balloons (How comments and some changes appear) ---
'        ' .RevisionsMode = wdBalloonRevisions ' (Example: Show revisions in balloons)
'        ' .RevisionsBalloonWidthType = wdBalloonWidthPercent ' (Example)
'        ' .RevisionsBalloonWidth = 50 ' (Example: 50% of window)
'
'    End With
'
'    ' Note: This function returns a Boolean by default (False if not explicitly set).
'    ' You could change it to Sub if no return value is needed, or return True on success.
'    TrackChangesOptions = True ' Indicate success (optional)
'
'End Function
'
'' =========================================================================
'' Debugging Subroutine: CheckboxTest (Seems unrelated to FastCompare)
'' Purpose: Lists all Document Variables stored in the active document
''          to the Immediate Window (Ctrl+G in VBA Editor).
'' =========================================================================
'Sub CheckboxTest()
'    Dim i As Long ' Loop counter
'
'    ' Check if there is an active document
'    If ActiveDocument Is Nothing Then
'        Debug.Print "No active document."
'        Exit Sub
'    End If
'
'    ' Check if the active document has any variables
'    If ActiveDocument.Variables.count = 0 Then
'        Debug.Print "Active document has no variables."
'        Exit Sub
'    End If
'
'    ' Loop through all document variables
'    Debug.Print "--- Document Variables (" & ActiveDocument.Name & ") ---"
'    For i = 1 To ActiveDocument.Variables.count
'        ' Print variable index, name, and value
'        Debug.Print "Var #" & i & " Name: " & ActiveDocument.Variables(i).Name & ", Value: " & ActiveDocument.Variables(i).value
'    Next i
'    Debug.Print "--- End of Variables ---"
'
'End Sub
'
