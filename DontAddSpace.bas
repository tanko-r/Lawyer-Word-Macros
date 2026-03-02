Attribute VB_Name = "DontAddSpace"
Sub DontAddSpaceMacro()
    
    ' Declaring variables to hold our objects.
    Dim currentParagraph As Word.Paragraph
    Dim selectedStyles As New Collection
    Dim styleName As Variant
    Dim styleToModify As Word.Style
    
    ' We start by checking if the user has selected any text.
    If Selection.Type = wdSelectionNormal Then
        
        ' Temporarily stop screen updates to make the macro run faster and smoother.
        Application.ScreenUpdating = False
        
        ' 1. Loop through every paragraph in the selection.
        For Each currentParagraph In Selection.Paragraphs
            
            ' Get the name of the style applied to the current paragraph.
            styleName = currentParagraph.Style
            
            ' The "On Error Resume Next" is a trick to bypass the error
            ' that occurs if we try to add a style name that's already in the collection.
            On Error Resume Next
            
            ' 2. Add the style name to our collection. If it's already there,
            ' the error is skipped, so we only store unique style names.
            selectedStyles.Add item:=styleName, key:=styleName
            
            On Error GoTo 0 ' Always remember to turn error handling back on!
            
        Next currentParagraph
        
        ' 3. Now, we loop through our collection of unique style names to modify them.
        For Each styleName In selectedStyles
            
            ' Get the actual Style object from the document's Styles collection.
            Set styleToModify = ActiveDocument.Styles(styleName)
            
            ' The key line! Setting this property to False is the same as:
            ' UNCHECKING the box for "Don't add space between paragraphs of the same style"
            styleToModify.NoSpaceBetweenParagraphsOfSameStyle = False
            
            ' A little bonus: To ensure the space is respected, it's good practice
            ' to check the SpaceAfter property of the style too.
            ' We'll ensure it's not set to 0.
            If styleToModify.ParagraphFormat.spaceAfter = 0 Then
                 ' You can choose any value here, 6 or 12 points are common.
                 styleToModify.ParagraphFormat.spaceAfter = 6
            End If
            
        Next styleName
        
        ' 4. Clean up and inform the user.
        Application.ScreenUpdating = True
    Else
        MsgBox "Please select the text you want to modify the styles for first!", vbExclamation
    End If

End Sub
