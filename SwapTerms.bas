Attribute VB_Name = "SwapTerms"
Sub SwapTwoTerms()

    ' Declare variables
    Dim userInput As String
    Dim terms() As String
    Dim term1 As String
    Dim term2 As String
    Dim r As Range
    Dim text As String
    
    ' Ask user for input
    userInput = InputBox("Enter the terms to swap, separated by a comma(,) with no space", "Input two terms")
    If userInput = "" Then Exit Sub  ' If user does not provide any input
  
    ' Split the user input to get two terms
    terms = Split(userInput, ",")
    If UBound(terms) <> 1 Then
        MsgBox "Input not valid, please enter exactly two terms separated by a comma", vbExclamation
        Exit Sub
    End If
  
    term1 = Trim(terms(0))
    term2 = Trim(terms(1))
  
    ' Loop through each word in the selected text
    Set r = Selection.Range
    text = r.text
    
    ' Swap the terms
    text = Replace(text, term1, "<<temp>>", , , vbTextCompare)
    text = Replace(text, term2, term1, , , vbTextCompare)
    text = Replace(text, "<<temp>>", term2, , , vbTextCompare)

    r.text = text

End Sub
