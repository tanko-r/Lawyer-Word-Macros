Attribute VB_Name = "AttachRedline"
Sub AttachRedline()

    Dim strFileName As String
    Dim baseName As String
    Dim strFolderPath As String
    Dim wOl As Object
    Dim wOlMail As Object
    Dim lastParenOpen As Integer
    Dim insideParens As String
    Dim i As Integer
    Dim isDocNum As Boolean
    
    ' Get the actual filename of the document
    strFileName = ActiveDocument.Name
    
    ' 1. Strip off the file extension (e.g., ".docx")
    If InStrRev(strFileName, ".") > 0 Then
        baseName = Left(strFileName, InStrRev(strFileName, ".") - 1)
    Else
        baseName = strFileName
    End If
    
    ' 2. Check if the name ends with a closing parenthesis
    If Right(baseName, 1) = ")" Then
        lastParenOpen = InStrRev(baseName, "(")
        
        If lastParenOpen > 0 Then
            ' Extract what is inside the last set of parentheses
            insideParens = Mid(baseName, lastParenOpen + 1, Len(baseName) - lastParenOpen - 1)
            
            ' Check if everything inside is ONLY a number or a period
            isDocNum = True
            For i = 1 To Len(insideParens)
                If Not (Mid(insideParens, i, 1) Like "#" Or Mid(insideParens, i, 1) = ".") Then
                    isDocNum = False
                    Exit For ' We found a letter or space, so it's not the iManage number
                End If
            Next i
            
            ' If it was just numbers/periods, remove it and trim any trailing spaces
            If isDocNum And Len(insideParens) > 0 Then
                baseName = Trim(Left(baseName, lastParenOpen - 1))
            End If
        End If
    End If
    
    ' 3. Construct the full file name for the temporary PDF
    strFolderPath = Environ("TEMP") & "\"
    strFileName = strFolderPath & baseName & "-redline.pdf"

    ' Convert the Word document to PDF
    On Error GoTo SaveAsError
    ActiveDocument.SaveAs2 FileName:=strFileName, FileFormat:=wdFormatPDF
    On Error GoTo 0 ' Reset error handler

    ' Create a new Outlook email safely
    On Error Resume Next
    Set wOl = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Set wOl = CreateObject("Outlook.Application")
    End If
    On Error GoTo SaveAsError ' Re-enable error handler

    Set wOlMail = wOl.CreateItem(0)

    With wOlMail
        .Attachments.Add strFileName
        .Display
    End With
    
    ' Clean up
    Set wOlMail = Nothing
    Set wOl = Nothing
    
    ' Delete the temporary PDF file after attachment
    On Error Resume Next
    Kill strFileName
    On Error GoTo 0
    
    Exit Sub

SaveAsError:
    MsgBox "Failed to process the document. Please ensure it is not protected or in use.", vbCritical
    Set wOlMail = Nothing
    Set wOl = Nothing
    
End Sub
