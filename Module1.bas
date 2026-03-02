Attribute VB_Name = "Module1"
Sub FindDef()
    Dim originalSelection As Range
    Dim searchTerm As String
    Dim myForm As frmDefinition

    ' 1. Validation: Ensure text is selected
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select a term to find its definition.", vbInformation, "No Term Selected"
        Exit Sub
    End If
    
    ' 2. Capture and clean the selection (remove leading/trailing spaces)
    Set originalSelection = Selection.Range
    With originalSelection
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32), wdForward
    End With
    
    searchTerm = originalSelection.text

    ' 3. Initialize the UserForm and start the search
    ' Using 'New' ensures a fresh instance of the form
    Set myForm = New frmDefinition
    myForm.StartSearch searchTerm
End Sub
