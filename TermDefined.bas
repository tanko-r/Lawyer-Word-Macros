Attribute VB_Name = "TermDefined"
Sub FindDef()
    ' --- Variable Declarations ---
    Dim wdDoc As Document
    Dim originalSelection As Range
    Dim searchRange As Range
    Dim definitionText As String
    Dim userResponse As VbMsgBoxResult
    Dim wasFound As Boolean

    ' --- Initialization ---
    ' Set the macro to work on the currently active document.
    Set wdDoc = Application.ActiveDocument
    
    ' Check if the user has selected any text. If not, show a message and stop the macro.
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select a term to find its definition.", vbInformation, "No Term Selected"
        Exit Sub
    End If
    
    ' Store the user's currently selected text in a Range object.
    Set originalSelection = Selection.Range
    
    ' Trim any leading or trailing spaces from the user's selection to ensure a clean search.
    With originalSelection
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32), wdForward
    End With

    ' Define the entire body of the document as the area to search.
    Set searchRange = wdDoc.content

    ' --- Search for the Definition ---
    ' Configure the find parameters.
    With searchRange.Find
        .ClearFormatting
        ' We are searching for the selected text when it is preceded by a double-quote,
        ' as this is how definitions are typically formatted (e.g., "Term" means...).
        ' Word's Find function is usually smart enough to match both straight ("") and curly (“”) quotes.
        .text = Chr(34) & originalSelection.text
        .Forward = True
        .Wrap = wdFindStop ' Stop searching at the end of the document.
        .Format = False
        .MatchCase = False ' Set to True if you need case-sensitive searches.
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        
        ' Execute the search. The 'wasFound' variable will be True if a match is found, otherwise False.
        wasFound = .Execute
    End With

    ' --- Process the Search Results ---
    If wasFound Then
        ' This block runs if the search was successful.
        ' 'searchRange' now points to the location of the found text.
        
        ' Expand the range to include the entire sentence that contains the definition.
        searchRange.Expand Unit:=wdSentence
        definitionText = searchRange.text
        
        ' Display a message box showing the definition and asking the user for the next step.
        userResponse = MsgBox(prompt:="Definition found:" & vbCrLf & vbCrLf & definitionText & _
                                     vbCrLf & vbCrLf & "Go to this definition in a split screen?  Press Ctrl-Alt-S to exit split screen.", _
                              Buttons:=vbYesNo, _
                              title:="Definition Found")
                              
        If userResponse = vbYes Then
            ' This block runs if the user clicks "Yes".
            
            ' Split the active window into two panes.
            ActiveWindow.Split = True
            
            ' Select the definition's text. This automatically navigates the active pane
            ' to the definition's location in the document.
            searchRange.Select
            
        End If
        ' If the user clicks "No", the macro simply does nothing further and finishes.
        
    Else
        ' This block runs if the search did not find a matching definition.
        MsgBox prompt:="The term """ & originalSelection.text & """ does not appear to be defined in this document.", _
               title:="Not Defined"
    End If

    ' --- Cleanup ---
    ' Release the object variables from memory. This is good practice.
    Set wdDoc = Nothing
    Set originalSelection = Nothing
    Set searchRange = Nothing

End Sub




