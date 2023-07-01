in file: word/vbaProject.bin - OLE stream: 'VBA/ApplyHeadingStyles'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub ChangeCaptionStyle()
    Dim para        As Word.Paragraph
    Dim testStyle As Word.Style

    On Error Resume Next
        Set testStyle = ActiveDocument.Styles("Level 2")
        If testStyle Is Nothing Then
            MsgBox "Level 2 is not in the document.  You have to apply Level 2 to the Level 2 sections first."
            GoTo lbl_End
        End If
    On Error GoTo 0
    
    For Each para In ActiveDocument.Paragraphs
        'paracheck = para.Range.Sentences(1)
        If para.Style = "Level 2" Then
            para.Range.Sentences(1).Select
            Selection.Range.Style = "Heading 2 Char"
        End If
    Next para

    Set para = Nothing
lbl_End:
End Sub
-------------------------------------------------------------------------------
