Attribute VB_Name = "NarrowCompare"
'================================================================================
' MACRO:         StartSelectionCompare
' PURPOSE:       Launches the UserForm to begin the text comparison process.
' INSTRUCTIONS:  Run this macro to start.
'================================================================================
Sub StartSelectionCompare()
    ' This line displays our custom form, but allows the user
    ' to continue interacting with the Word document.
    frmCompareTool.Show vbModeless
End Sub


'================================================================================
' SUB:           PerformTheComparison
' PURPOSE:       Called by the UserForm after both selections are captured.
'                Performs the document comparison.
' PARAMETERS:
'   - rngOriginal: A Range object containing the first selection.
'   - rngRevised:  A Range object containing the second selection.
'================================================================================
Public Sub PerformTheComparison(ByVal rngOriginal As Range, ByVal rngRevised As Range, ByVal sRngOriginalName As String, ByVal sRngRevisedName As String)

    ' --- Variable Declarations ---
    Dim docOriginal As Document
    Dim docRevised As Document
    Dim docCompare As Document
    Dim sCompareName As String
        
    On Error GoTo ErrorHandler

    ' --- Optimize Performance ---
    Application.ScreenUpdating = False
    
    ' --- Perform the Comparison ---
    Set docOriginal = Documents.Add(Visible:=False)
    Set docRevised = Documents.Add(Visible:=False)
    
    docOriginal.content.FormattedText = rngOriginal.FormattedText
    docRevised.content.FormattedText = rngRevised.FormattedText
    
    Set docCompare = Application.CompareDocuments( _
        OriginalDocument:=docOriginal, _
        RevisedDocument:=docRevised, _
        Destination:=wdCompareDestinationNew, _
        Granularity:=wdGranularityWordLevel, _
        CompareFormatting:=True, _
        CompareCaseChanges:=True, _
        CompareWhitespace:=False)

    ' --- Clean Up ---
    docOriginal.Close SaveChanges:=wdDoNotSaveChanges
    docRevised.Close SaveChanges:=wdDoNotSaveChanges
    
    'Name and save the comparison document
    If sRngOriginalName = sRngRevisedName Then
        sCompareName = "Text From " & sRngOriginalName
    Else
        sCompareName = "Text From " & sRngOriginalName & " +++and+++ " & sRngRevisedName
    End If
    
    docCompare.SaveAs2 FileName:=Environ("TEMP") & "\" & sCompareName
    
    docCompare.ActiveWindow.Visible = True
    Application.ScreenUpdating = True
    
'    MsgBox "Comparison complete!", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during comparison: " & Err.Description, vbCritical, "Error"
    Application.ScreenUpdating = True
    ' Ensure temp docs are closed on error
    If Not docOriginal Is Nothing Then docOriginal.Close wdDoNotSaveChanges
    If Not docRevised Is Nothing Then docRevised.Close wdDoNotSaveChanges
End Sub

