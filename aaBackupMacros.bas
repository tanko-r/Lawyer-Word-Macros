Attribute VB_Name = "aaBackupMacros"
Option Explicit

Sub ExportAllModules()
    Dim VBComp As VBIDE.VBComponent
    Dim SaveToDirectory As String
    Dim fso As Object
    Dim templatePath As String
    Dim dateToday As String
    
    Set fso = CreateObject("scripting.FileSystemObject")
    templatePath = ActiveDocument.FullName
    dateToday = Format(Date, "mm.dd.yyyy")
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
    
    fso.CopyFile Source:=templatePath, Destination:=SaveToDirectory & "templates\" & "FKSDO Macros " & dateToday & ".dotm"
End Sub

