in file: word/vbaProject.bin - OLE stream: 'VBA/FKSDOSaveInSequence'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


Sub SaveInSequence()
'PURPOSE: Save file in version sequence with existing file, add a new version indicator to filename
'by David Rubenstein - Copyright June 29, 2021

Dim sName As String, sNewName As String
Dim sDoc As Document
Dim sExt As String
Dim sVer As String
Dim vVer As Variant
Dim sDate As String
Dim newFilename As String
Dim sFormat As String
Dim verBeforeDot As String
Dim lastInSeq As String
Dim docTitle As String
Dim sPath As String
Dim myRange As Range
Dim inFolder As String

inFolder = FilePath.FooterCheck
If inFolder = "" Then inFolder = "G:\"
    
' Open filepicker to choose document to sequence from
With Application.FileDialog(msoFileDialogOpen)
    .Title = "Choose the most recent version"
    .AllowMultiSelect = False
    .InitialFileName = inFolder
    If .Show <> -1 Then
        Exit Sub
    End If
lastInSeq = .SelectedItems(1)
sPath = Left(lastInSeq, InStrRev(lastInSeq, "\"))
End With


' Get version and date information about lastInSeq
    sFormat = DateFormat(lastInSeq)
    sDate = Format(Date, sFormat) 'Format the date
    sNewName = Trim(Left(lastInSeq, InStrRev(lastInSeq, "(") - 1))
    sVer = sNewName
    vVer = Split(sVer, Chr(32))
    sVer = onlyDigits(CStr(vVer(UBound(vVer))))

sDocTitle = Left(sNewName, InStr(sNewName, sVer) - 1)

' If doc has an incremental version number, strip out the incremental portion, then set sNewName as next integer version
If InStrRev(sVer, ".") >= 1 Then
    verBeforeDot = Left(sVer, InStrRev(sVer, ".") - 1)
    sNewName = sDocTitle & Format(verBeforeDot + 1, "00")
Else
    sNewName = Replace(sNewName, sVer, Format(sVer + 1, "00"))
End If
sNewName = Mid(sNewName, InStrRev(sNewName, "\") + 1)

' Prompt for who created this document version
Dim Drafter As String
Drafter = InputBox("Whose draft is this?  E.g. 'Seller' or 'Tenant'", "Drafter")
If Drafter = "" Then Exit Sub

' Name doc using existing spacing and quirks
If InStrRev(lastInSeq, " (") > 0 Then
    newFilename = sNewName & " (" & Drafter & Chr(32) & sDate & ")"
Else
    newFilename = sNewName & "(" & Drafter & Chr(32) & sDate & ")"
End If

' Save document in folder where lastInSeq lives
With Application.Dialogs(wdDialogFileSaveAs)
    .Name = sPath & newFilename
    .Show
End With

FilePath.UpdatePathMacro

lbl_Exit:
    Exit Sub
    
    'Error Handler
'NotSavedYet:
'    MsgBox "This file has not been initially saved. " & _
'           "Cannot save a new version!", vbCritical, "Not Saved To Computer"    '
'    GoTo lbl_Exit

End Sub

Private Function DateFormat(sName As String) As String
' ©Graham Mayor - https://www.gmayor.com - Last updated - 04 Apr 2021
Dim sDate As String
sDate = Right(sName, Len(sName) - InStrRev(sName, "(") + 1)
    If InStr(1, sDate, "(") > 0 Then
        sDate = Split(sDate, "(")(1)
        sDate = Split(sDate, " ")(1)
        sDate = Split(sDate, ")")(0)
        If InStr(1, sDate, ".") > 0 Then
            DateFormat = "mm.dd.yy" 'note lower case!
        Else
            DateFormat = "mmddyy"
        End If
    Else
        DateFormat = "mm.dd.yy"
    End If
End Function

Private Function onlyDigits(s As String) As String
'https://stackoverflow.com/questions/7239328/how-to-find-numbers-from-a-string#7239408
'modified by Graham Mayor - https://www.gmayor.com - Last updated - 04 Apr 2021
' Variables needed (remember to use "option explicit").   '
Dim retval As String    ' This is the return string for the version number.      '
Dim i As Integer        ' Counter for character position. '

' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Or Mid(s, i, 1) = "." Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
       
    
End Function



-------------------------------------------------------------------------------
