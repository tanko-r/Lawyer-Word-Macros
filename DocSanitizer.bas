Attribute VB_Name = "DocSanitizer"
Sub DocSanitizer()
    Dim userInput As String
    Dim termList() As String, parts() As String
    Dim findTerm As String, replaceTerm As String
    Dim i As Integer
    Dim DataObj As Object

    ' Read the find/replace string from the clipboard (no length limit, no prompt)
    Set DataObj = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  ' MSForms.DataObject
    On Error Resume Next
    DataObj.GetFromClipboard
    userInput = DataObj.GetText(1)
    On Error GoTo 0

    If Trim(userInput) = "" Then
        MsgBox "Clipboard has no text. Click 'Copy' in the app first.", vbExclamation, "Nothing to replace"
        Exit Sub
    End If

    termList = Split(userInput, "^^")
    For i = LBound(termList) To UBound(termList)
        parts = Split(termList(i), "%%")
        If UBound(parts) >= 1 Then
            findTerm = Trim(parts(0))
            replaceTerm = Trim(parts(1))
            If Len(findTerm) > 0 Then
                With ActiveDocument.content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .text = findTerm
                    .Replacement.text = replaceTerm
                    .MatchCase = False
                    .MatchWildcards = False
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
    Next i

    MsgBox "Replacement complete.", vbInformation, "Done"
End Sub
