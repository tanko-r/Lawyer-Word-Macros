in file: word/vbaProject.bin - OLE stream: 'VBA/CopyToForms'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub CopyToForms()
'Copies active document into a form folder.  Does not change the active document at all.

Dim sDoc As Document
Dim sDocPath As String
Dim sFolderPicker As FileDialog
Dim sFormPath As String
Dim sFormName As String
Dim fso As Object
Set fso = CreateObject("scripting.FileSystemObject")
Set sDoc = ActiveDocument

Select Case MsgBox("Save this document first? Only the saved version will be copied.", vbYesNoCancel)
    Case Is = vbYes
        ActiveDocument.Save
    Case Is = vbCancel
        Exit Sub
End Select

sFormName = InputBox("Name this form, e.g. 'Lease -- Tenant friendly'")
If Len(sFormName) = 0 Then Exit Sub

sDocPath = sDoc.FullName
Set sFolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
With sFolderPicker
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\DSR\Documents\Candidate Forms\"
        .Show
        sFormPath = .SelectedItems(1) & "\" & sFormName & ".docx"
End With

fso.CopyFile Source:=sDocPath, Destination:=sFormPath
fso.Quit

End Sub

-------------------------------------------------------------------------------
