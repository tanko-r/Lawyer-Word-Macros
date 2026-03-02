VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompareTool 
   Caption         =   "Text Comparison"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "frmCompareTool.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCompareTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- Code for UserForm: frmCompareTool ---

' Module-level variables to store the selections
Private rngOriginal As Range
Private rngRevised As Range
Private sRngOriginalName As String
Private sRngRevisedName As String

' This runs when the form first opens
Private Sub UserForm_Initialize()
    ' Set the initial state: waiting for the first selection
    lblInstructions.Caption = "1. Highlight the FIRST (original) text in your document."
    btnCapture.Caption = "Capture Original Text"
End Sub

' This runs when the user clicks the button
Private Sub btnCapture_Click()
    ' Check if the user has selected anything
    If Selection.Type = wdSelectionIP Then
        MsgBox "No text is selected. Please highlight the text before clicking the button.", vbExclamation
        Exit Sub
    End If
    
    ' --- State Machine ---
    ' If rngOriginal is empty, we are capturing the first selection
    If rngOriginal Is Nothing Then
        ' Capture the first selection
        Set rngOriginal = Selection.FormattedText.Duplicate
        sRngOriginalName = rngOriginal.Document.Name
        
        ' Update the form for the next step
        lblInstructions.Caption = "2. Now highlight the SECOND (revised) text."
        btnCapture.Caption = "Capture Revised & Compare"
        
    ' Otherwise, we are capturing the second selection
    Else
        ' Capture the second selection
        Set rngRevised = Selection.FormattedText.Duplicate
        sRngRevisedName = rngRevised.Document.Name
        
        ' Hide the form so it's not in the way
        Me.Hide
        
        ' Call the separate module to do the actual comparison work
        Call NarrowCompare.PerformTheComparison(rngOriginal, rngRevised, sRngOriginalName, sRngRevisedName)
        
        ' Close the form completely now that we are done
        Unload Me
    End If
End Sub


' --- Code for a standard Module (e.g., Module1) ---

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
Public Sub PerformTheComparison(ByVal rngOriginal As Range, ByVal rngRevised As Range)

    ' --- Variable Declarations ---
    Dim docOriginal As Document
    Dim docRevised As Document
    Dim docCompare As Document
    
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
    
    docCompare.ActiveWindow.Visible = True
    Application.ScreenUpdating = True
    
    MsgBox "Comparison complete!", vbInformation, "Success"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during comparison: " & Err.Description, vbCritical, "Error"
    Application.ScreenUpdating = True
    ' Ensure temp docs are closed on error
    If Not docOriginal Is Nothing Then docOriginal.Close wdDoNotSaveChanges
    If Not docRevised Is Nothing Then docRevised.Close wdDoNotSaveChanges
End Sub

