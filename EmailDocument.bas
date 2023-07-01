Attribute VB_Name = "EmailDocument"
Sub EmailDocumentMacro()
Dim wOl As Object
Dim wOlMail As Object
Dim wOlInsp As Object
Dim wOlAttachment As Object
Dim wd As Object
Dim wStrBody As String
Dim wPath As String

Debug.Print ActiveDocument.Saved

If ActiveDocument.Saved = False Then
    If MsgBox(Prompt:="Do you want to save the document first? Any changes since your last save will not be sent.", Buttons:=vbYesNoCancel) = vbYes Then
        ActiveDocument.Save
        Else: Exit Sub
    End If
End If

wPath = ActiveDocument.Path
oStrBody = "Here's a link to the document folder: <a href=" & Chr(34) & wPath & Chr(34) & ">" & wPath & "</a>"

Set wOl = GetObject(Class:="Outlook.Application")
Set wOlMail = wOl.CreateItem(0)

With wOlMail
    .HTMLbody = .HTMLbody & oStrBody
    Set wOlInsp = .GetInspector
    If wOlInsp.EditorType = 4 Then Set wd = wOlInsp.WordEditor
    .Attachments.Add ActiveDocument.FullName
    .Display
End With

End Sub
