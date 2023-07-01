in file: word/vbaProject.bin - OLE stream: 'VBA/FKSDOSaveAs'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub SaveAsFKSDOFile()
Application.ScreenUpdating = False

Dim curFilename As String
Dim curFilePath As String
Dim newFilename As String
Dim docTitle As String
Dim sDate As String
Dim sErr As Boolean
'Dim testVar As String

curFilename = ActiveDocument.Path & "\" & ActiveDocument.Name 'save the filepath of the document in case the user wants a form redline

'testVar = ActiveDocument.Variables("formPath")

If curFilename <> "" Then
curFilename = Trim(Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1))
Else: curFilename = ""
End If

'Store the filepath so it can be added to the formPath variable.
curFilePath = ActiveDocument.FullName


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

On Error Resume Next ' create an error trap because there's no "exists" function for variables.
                     ' https://www.askwoody.com/forums/topic/check-to-see-if-a-docvariable-exists-before-running-line-of-vba-code/
Dim varCheck As String
varCheck = ActiveDocument.Variables("formPath").Value

If Err.Number = 0 Then
    ActiveDocument.Variables("formPath").Value = curFilePath
Else
    ActiveDocument.Variables.Add "formPath", curFilePath
End If
Debug.Print ActiveDocument.Variables("formPath").Value
On Error GoTo 0 ' Reset error handler

ActiveDocument.Save

ErrExit:
    Application.ScreenUpdating = True
    Exit Sub

End Sub

-------------------------------------------------------------------------------
