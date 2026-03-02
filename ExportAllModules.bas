Attribute VB_Name = "ExportAllModules"
Sub ExportAllComponents()
    Dim strPath As String
    Dim vbc As Object
    Dim strExt As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then
            strPath = .SelectedItems(1)
            If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
            
            For Each vbc In ActiveDocument.VBProject.VBComponents
                ' Determine the correct extension based on component type
                Select Case vbc.Type
                    Case 1 ' Standard Module
                        strExt = ".bas"
                    Case 2 ' Class Module
                        strExt = ".cls"
                    Case 3 ' UserForm
                        strExt = ".frm"
                    Case Else ' Document objects (ThisDocument, Sheets, etc.)
                        strExt = ".cls"
                End Select
                
                ' Export with the proper name and extension
                vbc.Export strPath & vbc.Name & strExt
            Next vbc
            
            MsgBox "Export Complete!", vbInformation
        Else
            MsgBox "No folder specified...", vbExclamation
        End If
    End With
End Sub
