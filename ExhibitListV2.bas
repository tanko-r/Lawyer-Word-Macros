Attribute VB_Name = "ExhibitListV2"
Option Explicit ' Enforces variable declaration

Sub GenerateExhibitScheduleList_Regex_NoPageBreak()

    ' --- Add Reference ---
    ' Before running: Go to Tools > References in the VBA editor
    ' and check "Microsoft VBScript Regular Expressions 5.5"

    Dim regex As Object         ' RegExp object
    Dim matches As Object       ' Collection of matches found
    Dim match As Object         ' Individual match object
    Dim docText As String       ' Entire document text
    Dim exhibitList As Collection ' To store the found items "Identifier - Title"
    Dim identifier As String    ' Stores the found "Exhibit A", "Schedule 1", etc.
    Dim title As String         ' Stores the title found
    Dim item As Variant         ' Loop variable for the collection
    Dim pattern As String       ' The Regex pattern

    ' --- Initialization ---
    Set exhibitList = New Collection
    docText = ActiveDocument.content.text ' Get entire doc text

    ' --- Create and Configure Regex Object ---
    On Error Resume Next ' Handle error if reference wasn't added
    Set regex = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Then
        MsgBox "Error creating Regex object." & vbCrLf & _
               "Please ensure the 'Microsoft VBScript Regular Expressions 5.5' reference is added (Tools > References).", vbCritical
        Exit Sub
    End If
    On Error GoTo 0 ' Turn error handling off

    ' --- Define the Regex Pattern with Negative Lookahead to Exclude Page Breaks ---
    ' Explanation (Changes from previous version):
    ' (?!\f)             Negative lookahead assertion: The \r (paragraph break) MUST NOT be immediately followed by a \f (form feed/page break)
    '
    pattern = "^((?:Exhibit|Schedule)\s+[A-Za-z0-9-]+)[ \t]*\r(?!\f)(?:\s*\r)*\s*([^\r\s][^\r]*)"

    With regex
        .Global = True         ' Find all occurrences in the string
        .MultiLine = True      ' ESSENTIAL: Allows ^ to match the start of each line
        .IgnoreCase = True     ' ENSURES "Exhibit", "Schedule" match regardless of case
        .pattern = pattern     ' Set the pattern
    End With

    ' --- Execute the Search ---
    Set matches = regex.Execute(docText)

    ' --- Loop Through Matches ---
    If matches.count > 0 Then
        For Each match In matches
            ' Defensive check: Ensure we have the expected number of capturing groups
            If match.SubMatches.count = 2 Then
                ' Group 1 (index 0) is the identifier line (e.g., "Exhibit A")
                identifier = Trim(match.SubMatches(0))

                ' Group 2 (index 1) is the title line
                title = Trim(match.SubMatches(1))

                ' --- Store the Result ---
                 exhibitList.Add identifier & " - " & title
            Else
                ' Optional: Log if a match didn't capture expected groups
                Debug.Print "Warning: Regex match found value '" & match.value & "' but did not capture expected groups."
            End If
        Next match
    End If ' End matches.Count > 0

    ' --- Output the List ---
    If exhibitList.count > 0 Then
        ' Insert the list at the current cursor position
        Selection.TypeText text:="List of Exhibits and Schedules:" & vbCrLf ' Add a header
        For Each item In exhibitList
            Selection.TypeText text:=item & vbCrLf ' Insert item and paragraph break
        Next item
        Selection.TypeText text:=vbCrLf ' Add an extra line break after the list
        MsgBox exhibitList.count & " items added to the list.", vbInformation
    Else
        MsgBox "No Exhibit or Schedule headings matching the pattern were found.", vbInformation
    End If

    ' --- Cleanup ---
    Set regex = Nothing
    Set matches = Nothing
    Set match = Nothing
    Set exhibitList = Nothing

End Sub

