Attribute VB_Name = "aaBackupMacros"
Option Explicit

Sub ExportAllModules()
    Dim VBComp As VBIDE.VBComponent
    Dim SaveToDirectory As String
    Dim fso As Object
    Dim templatePath As String
    Dim dateToday As String
    Dim doc As Document, docName As String
    
    Set fso = CreateObject("scripting.FileSystemObject")
    
    
    Set doc = ActiveDocument
    doc.Save
    templatePath = doc.FullName
    
    'check that the FKSDOMacros global template is open
    If InStr(templatePath, "STARTUP") = 0 Then
        MsgBox ("Please activate the Global Template and then rerun this.")
        Exit Sub
    End If
    
    SaveToDirectory = "C:\Users\david\Desktop\Tech Stuff\Git repos\FKSDO-Macros\"

    For Each VBComp In Application.VBE.ActiveVBProject.VBComponents
        If VBComp.Type = vbext_ct_StdModule Or VBComp.Type = vbext_ct_ClassModule Then
            VBComp.Export SaveToDirectory & "\" & VBComp.Name & ".bas"
        End If
    Next VBComp
    
    'copy the template into the "templates" folder
    
    dateToday = " export date " & Format(Date, "mm.dd.yyyy") & ", " & Format(Time, "hh.mm am/pm")
    docName = Left(doc.Name, InStr(1, doc.Name, ")", vbBinaryCompare))
    
    fso.CopyFile Source:=templatePath, Destination:=SaveToDirectory & "templates\" & docName & dateToday & ".dotm"
End Sub

