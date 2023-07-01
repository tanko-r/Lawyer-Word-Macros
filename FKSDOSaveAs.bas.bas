in file: word/vbaProject.bin - OLE stream: 'VBA/FKSDOSaveAs'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub SaveAsFKSDOFile()
Application.ScreenUpdating = False

Dim curFilename As String
Dim newFilename As String
Dim docTitle As String
Dim sDate As String
Dim sErr As Boolean

If curFilename <> "" Then
curFilename = Trim(Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1))
Else: curFilename = ""
End If

sDate = Format(Date, "mmddyy")
docTitle = InputBox("What is this document called?  E.g. 1AM to Lease", "Document Name", curFilename)
If docTitle = "" Then
    sErr = True
    Exit Sub
If docTitle = vbCancel Then Exit Sub
End If

newFilename = docTitle & "01 (" & Application.UserInitials & " " & sDate & ")"

With Application.Dialogs(wdDialogFileSaveAs)
    .Name = ActiveDocument.Path & "\" & newFilename
    If .Show <> 0 Then sErr = True
End With
        
FilePath.UpdatePathMacro
ActiveDocument.Save

ErrExit:
    Application.ScreenUpdating = True
    Exit Sub

End Sub

-------------------------------------------------------------------------------
