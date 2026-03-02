Attribute VB_Name = "FKSDOSaveAs"
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
    
    sDate = Format(Date, "mm.dd.yy")
    docTitle = InputBox("What is this document called?  E.g. 1AM to Lease", "Document Name", curFilename)
    If docTitle = "" Then
        sErr = True
        Exit Sub
    If docTitle = vbCancel Then Exit Sub
    End If
    
    newFilename = docTitle & " v01 (Polsinelli " & sDate & ")"
    
    'Copy the new filename to the clipboard
    Dim oFilename As DataObject
    Set oFilename = New DataObject
    oFilename.SetText newFilename
    oFilename.PutInClipboard
    
    SendKeys "^+s", True
    
'    With Application.Dialogs(wdDialogFileSaveAs)
'        .Name = ActiveDocument.Path & "\" & newFilename
'        If .Show <> 0 Then sErr = True
'    End With
            
ErrExit:
        Application.ScreenUpdating = True
        Exit Sub

End Sub

