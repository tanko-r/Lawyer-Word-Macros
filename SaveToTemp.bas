Attribute VB_Name = "SaveToTemp"
Sub SaveToTemp()
Dim sDocName As String
Dim sDocPath As String
Dim fso As FileSystemObject

Set fso = CreateObject("scripting.FileSystemObject")

sDocName = ActiveDocument.Name
sDocPath = ActiveDocument.FullName
Debug.Print sDocPath

'ActiveDocument.SaveAs2 FileName:=Environ("USERPROFILE") & "\Downloads\" & sDocName
fso.CopyFile Source:=sDocPath, Destination:=Environ("USERPROFILE") & "\Downloads\" & sDocName, OverWriteFiles:=True

Debug.Print Environ("USERPROFILE") & "\Downloads\"

End Sub
