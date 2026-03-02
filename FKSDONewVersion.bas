Attribute VB_Name = "FKSDONewVersion"
Option Explicit

Sub SaveNewVersion_Word()
'PURPOSE: Save file, if already exists add a new version indicator to filename.
'         Correctly parses filenames with multiple parenthetical groups.
'         Always prompts for the author for the new version.

Dim sPath As String
Dim docCaption As String
Dim baseFilename As String      ' Raw filename part, from iManage processing or full caption (no extension)
Dim namePartToProcess As String ' Filename part intended to end with version, after stripping ALL parentheses
Dim fileStem As String          ' e.g., "LYH03 - Deed - RRSA Tank Lot"
Dim versionPrefix As String     ' e.g., "v" or "Version"
Dim versionNumberStr As String  ' e.g., "03" or "03.1" (as string)
Dim newVersionNumberStr As String ' e.g., "04" (as string, formatted "00")
Dim finalAuthor As String
Dim currentDateStr As String
Dim newFilename As String

' Intermediate parsing variables
Dim posFirstOpenParen As Long
Dim posLastDot As Long
Dim fileExtension As String
Dim vVerParts As Variant
Dim lastWord As String
Dim char As String
Dim k As Long
Dim extractedNumPart As String
Dim numPartStartIndex As Long
Dim verBeforeDot As String
Dim newVersionCoreNum As Long ' Numeric value of the new version

    ' --- Initial Check: Document must be saved at least once ---
    If ActiveDocument.Path = "" Then
        MsgBox "This file has not been initially saved. " & _
               "Cannot save a new version.", vbCritical, "Not Saved To Computer"
        GoTo lbl_Exit
    End If
    sPath = ActiveDocument.Path ' Used for check, not directly in naming if using caption

    ' --- 1. Get and Clean Initial Filename ---
    docCaption = ActiveWindow.Caption

    ' Check for iManage format using the helper (assumed to be available)
    ' If iManageHelpers is in a different module, ensure it's accessible.
    Dim isIManage As Boolean
    On Error Resume Next ' In case iManageHelpers doesn't exist or errors
    isIManage = iManageHelpers.iManTestCaption(docCaption)
    On Error GoTo 0

    If isIManage Then
        Dim delimiterPos As Long
        Const IMANAGE_DELIMITER As String = "<ACTIVE> - "

        delimiterPos = InStr(1, docCaption, IMANAGE_DELIMITER)

        If delimiterPos > 0 Then
            ' Extract the part of the filename after the delimiter.
            baseFilename = Trim(Mid(docCaption, delimiterPos + Len(IMANAGE_DELIMITER)))
        Else
            ' Fallback for safety, in case the delimiter is not found.
            MsgBox "Could not find the expected iManage delimiter '" & IMANAGE_DELIMITER & "' in the window title." & vbCrLf & vbCrLf & _
                   "Caption: " & docCaption, vbCritical, "iManage Parsing Error"
            GoTo lbl_Exit
        End If
    Else
        baseFilename = docCaption
        ' Strip common Word extension if present at the very end of a non-iManage caption
        posLastDot = InStrRev(baseFilename, ".")
        If posLastDot > 0 Then
            ' Ensure the dot is not part of something like (12345.6)
            ' A simple check: if there's no open paren after the last dot, it's likely an extension.
            If InStr(posLastDot, baseFilename, "(") = 0 Then
                fileExtension = Mid(baseFilename, posLastDot)
                Select Case LCase(fileExtension)
                    Case ".doc", ".docx", ".docm"
                        baseFilename = Trim(Left(baseFilename, posLastDot - 1))
                End Select
            End If
        End If
    End If
    ' baseFilename is now like "LYH03 - Deed - RRSA Tank Lot v03 (RRSA 05.27.25)(509252824.5)" (no extension)

    ' --- 2. Isolate Core Filename Part (before any parentheses) ---
    ' This part is assumed to contain the version number at its end.
    namePartToProcess = baseFilename ' Start with the full baseFilename

    posFirstOpenParen = InStr(1, baseFilename, "(")
    If posFirstOpenParen > 0 Then
        ' Part before any parens, e.g., "LYH03 - Deed - RRSA Tank Lot v03"
        namePartToProcess = Trim(Left(baseFilename, posFirstOpenParen - 1))
    End If
    ' namePartToProcess is now the segment expected to end with the version, e.g., "LYH03 - Deed - RRSA Tank Lot v03"

    ' --- 3. Extract Version Info (fileStem, versionPrefix, versionNumberStr) from namePartToProcess ---
    If Trim(namePartToProcess) = "" Then
        MsgBox "Cannot determine filename stem after parsing. Original: '" & docCaption & "'", vbCritical, "Parsing Error"
        GoTo lbl_Exit
    End If

    vVerParts = Split(namePartToProcess, " ")
    lastWord = vVerParts(UBound(vVerParts)) ' e.g., "v03", "Version2.1", or "Lot" if no version found there

    extractedNumPart = ""
    numPartStartIndex = 0 ' 1-based index of where the number starts in lastWord
    For k = Len(lastWord) To 1 Step -1
        char = Mid(lastWord, k, 1)
        ' Allow digits and a single period for version numbers like "3.1"
        If (char >= "0" And char <= "9") Or (char = "." And InStr(1, extractedNumPart, ".") = 0) Then
            extractedNumPart = char & extractedNumPart
            numPartStartIndex = k
        Else
            ' If we hit a non-digit (and it's not the first char for a potential prefix like 'v') break
            If numPartStartIndex > 0 Then Exit For
            ' If numPartStartIndex is still 0, it means we haven't found any digit yet from the right
            ' so continue scanning left in case the word is like "v03" vs "03v"
        End If
    Next k

    If extractedNumPart <> "" Then ' Numeric part found at the end of lastWord
        versionNumberStr = extractedNumPart ' e.g., "03" or "03.1"
        If numPartStartIndex > 1 Then
            versionPrefix = Left(lastWord, numPartStartIndex - 1) ' e.g. "v" or "Version"
        Else
            versionPrefix = "" ' Number was at the start of lastWord, e.g., "03" itself
        End If

        If UBound(vVerParts) >= 1 Then ' More than one word in namePartToProcess
            fileStem = Trim(Left(namePartToProcess, Len(namePartToProcess) - Len(lastWord) - 1)) ' Space before lastWord
        Else ' namePartToProcess was just the lastWord itself (e.g., "v01")
            fileStem = ""
        End If
    Else
        ' No numeric part found at the end of lastWord. Assume lastWord is part of the filestem.
        versionNumberStr = ""  ' Will trigger fallback to "01"
        versionPrefix = "v"    ' Default prefix for the new version number
        fileStem = namePartToProcess ' The whole namePartToProcess is the stem
    End If

    ' --- 4. Increment Version Number ---
    If versionNumberStr <> "" Then ' A version string was successfully parsed
        If InStr(1, versionNumberStr, ".") > 0 Then ' It's an incremental like "03.1"
            verBeforeDot = Left(versionNumberStr, InStr(1, versionNumberStr, ".") - 1)
            If IsNumeric(verBeforeDot) And verBeforeDot <> "" Then
                newVersionCoreNum = CLng(verBeforeDot) + 1
                newVersionNumberStr = Format(newVersionCoreNum, "00") ' Increments X.Y to (X+1) formatted "00"
            Else
                MsgBox "Could not parse major part of version '" & versionNumberStr & "'. Defaulting to version 01.", vbExclamation, "Version Warning"
                newVersionNumberStr = "01"
                If versionPrefix = "" Then versionPrefix = "v" ' Ensure 'v' prefix if defaulting
            End If
        Else ' It's a major version like "03"
            If IsNumeric(versionNumberStr) Then
                newVersionCoreNum = CLng(versionNumberStr) + 1
                newVersionNumberStr = Format(newVersionCoreNum, "00") ' Increments X to X+1 formatted "00"
            Else
                MsgBox "Parsed version '" & versionNumberStr & "' is not fully numeric. Defaulting to version 01.", vbExclamation, "Version Warning"
                newVersionNumberStr = "01"
                If versionPrefix = "" Then versionPrefix = "v"
            End If
        End If
    Else ' No version was parsed from filename, so this will be the first numbered version
        newVersionNumberStr = "01"
        ' versionPrefix is already "v" (default from parsing section, or ensure it)
        If versionPrefix = "" And fileStem <> namePartToProcess Then versionPrefix = "v"
    End If

    ' --- 5. ALWAYS Prompt for Final Author ---
    finalAuthor = InputBox("Whose draft is this? (e.g. Polsinelli, County, Initials)", _
                           "Author for New Version", Application.UserInitials)
    If Trim(finalAuthor) = "" Then GoTo lbl_Exit ' User cancelled or entered nothing

    ' --- 6. Get Current Date ---
    currentDateStr = Format(Date, "mm.dd.yy")

    ' --- 7. Construct New Filename ---
    If Trim(fileStem) <> "" Then
        newFilename = Trim(fileStem) & " " & versionPrefix & newVersionNumberStr & " (" & Trim(finalAuthor) & " " & currentDateStr & ")"
    Else ' fileStem was empty, e.g., original name might have been just "v01" or "Report"
        ' If fileStem is empty, ensure we don't have a leading space if versionPrefix is also empty
        If versionPrefix <> "" Then
            newFilename = versionPrefix & newVersionNumberStr & " (" & Trim(finalAuthor) & " " & currentDateStr & ")"
        Else
            newFilename = newVersionNumberStr & " (" & Trim(finalAuthor) & " " & currentDateStr & ")"
        End If
    End If

    ' --- 8. Copy to Clipboard and Attempt Save (using SendKeys) ---
    Dim oFilename As Object ' For late binding DataObject
    On Error Resume Next ' Handle if DataObject cannot be created
    Set oFilename = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") ' CLSID for MSForms.DataObject
    If Err.Number <> 0 Then
        MsgBox "Error creating DataObject for clipboard: " & Err.Description & vbCrLf & _
               "The new filename could not be copied to the clipboard automatically.", vbExclamation, "Clipboard Error"
        Err.Clear
    Else
        oFilename.SetText newFilename
        oFilename.PutInClipboard
        If Application.UserInitials <> "DSR" Then
            MsgBox "New filename copied to clipboard:" & vbCrLf & newFilename & vbCrLf & vbCrLf & _
                   "The macro will now attempt to trigger the iManage 'Save As' shortcut.  You may need to manually click the Save As New Version button.  " & _
                   "You will need to manually paste the filename into the dialog.", vbInformation, "Filename Ready"
        End If
    End If
    On Error GoTo 0 ' Reset error handling for DataObject

    ' WARNING: SendKeys is notoriously unreliable. It depends on the correct window having focus
    ' and the shortcut "%3" (ALT+3) being active and doing what's expected for that user.
    ' This line is specific to "DSR's computer" as per original comment.
    On Error Resume Next ' Use error trapping for SendKeys
    SendKeys "%4", False
    If Err.Number <> 0 Then
        MsgBox "Attempt to send ALT+3 keystroke failed (Error: " & Err.Description & "). " & _
               "Please manually invoke your 'Save As New Version' command and paste the filename from the clipboard.", vbExclamation, "SendKeys Warning"
        Err.Clear
    End If
    On Error GoTo 0 ' Reset error trapping

lbl_Exit:
    Set oFilename = Nothing
    Exit Sub

End Sub



