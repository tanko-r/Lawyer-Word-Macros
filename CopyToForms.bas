Attribute VB_Name = "CopyToForms"
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
        If ActiveDocument.ReadOnly = True Then ActiveDocument.SaveAs2 ("C:\Users\dsrub\Downloads\" & ActiveDocument.Name) Else: ActiveDocument.Save
    Case Is = vbCancel
        Exit Sub
End Select

sFormName = InputBox("Name this form, e.g. 'Lease -- Tenant friendly'")
If Len(sFormName) = 0 Then Exit Sub

sDocPath = sDoc.FullName
Set sFolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
With sFolderPicker
        .title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\dsrub\Documents\Polsinelli Candidate Forms\Candidate Forms\"
        If .Show <> 0 Then sFormPath = .SelectedItems(1) & "\" & sFormName & ".docx" Else Exit Sub
End With

fso.CopyFile Source:=sDocPath, Destination:=sFormPath
Set fso = Nothing


End Sub

