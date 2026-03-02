Attribute VB_Name = "AttachRedline"
Sub AttachRedline()

    Dim strFileName As String
    Dim strRedlineName As String
    Dim iManageTrue As Boolean
    Dim strDocCaption As String
    Dim objWord As Word.Application
    Dim objDoc As Word.Document
    Dim wOlMail As Object
    Dim wOlInsp As Object
    Dim wOlAttachment As Object
    Dim strFolderPath As String
    
    Set objDoc = ActiveDocument
    
    ' Get the name of the active Word document
    strFileName = objDoc.Name
    strDocCaption = ActiveWindow.Caption
    
    ' Assuming iManageHelpers module exists and functions are correctly implemented
     If iManageHelpers.iManTestCaption(strDocCaption) Then
         iManageTrue = True
         'Delete document number from caption
         strDocCaption = iManageHelpers.RemoveDocNo(strDocCaption)
     End If
    ' Commented out the iManageHelpers block as it's not provided and might cause errors
    ' if the module isn't present. Uncomment if you have this module.
            
    ' --- MODIFICATION START ---
    ' Get the path to the user's temporary folder
    ' This will return a path like C:\Users\<Username>\AppData\Local\Temp
    strFolderPath = Environ("TEMP") & "\"
    ' --- MODIFICATION END ---
    
    ' Construct the full file name for the saved PDF
    ' Use a cleaned version of the document caption for the filename
    ' Replace any invalid characters for filenames
    Dim cleanedDocCaption As String
    cleanedDocCaption = Replace(strDocCaption, ":", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, "\", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, "/", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, "*", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, "?", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, Chr(34), "_") ' Double quote
    cleanedDocCaption = Replace(cleanedDocCaption, "<", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, ">", "_")
    cleanedDocCaption = Replace(cleanedDocCaption, "|", "_")
    
    strFileName = strFolderPath & cleanedDocCaption & "-redline.pdf"

    ' Convert the Word document to PDF
    ' Add error handling for the SaveAs operation
    On Error GoTo SaveAsError
    ActiveDocument.SaveAs2 FileName:=strFileName, FileFormat:=wdFormatPDF
    On Error GoTo 0 ' Reset error handler

    ' Create a new Outlook email
    Set wOl = GetObject(Class:="Outlook.Application")
    Set wOlMail = wOl.CreateItem(0)

    With wOlMail
        Set wOlInsp = .GetInspector
        'If wOlInsp.EditorType = 4 Then 'Set wd = wOlInsp.WordEditor
        .Attachments.Add strFileName
        .Display
    End With
    
    ' Clean up
    Set wOlAttachment = Nothing
    Set wOlInsp = Nothing
    Set wOlMail = Nothing
    Set wOl = Nothing
    Set objDoc = Nothing
    
    ' Delete the temporary PDF file after attachment
    On Error Resume Next ' In case file is in use or doesn't exist
    Kill strFileName
    On Error GoTo 0
    
    Exit Sub ' Exit the sub cleanly

SaveAsError:
    MsgBox "Failed to save the document as PDF. Please ensure the document is not protected or in use, and you have write permissions to the temporary folder.", vbCritical
    ' Clean up in case of error before exiting
    Set wOlAttachment = Nothing
    Set wOlInsp = Nothing
    Set wOlMail = Nothing
    Set wOl = Nothing
    Set objDoc = Nothing
    
End Sub

