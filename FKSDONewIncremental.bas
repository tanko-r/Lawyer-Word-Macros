Attribute VB_Name = "FKSDONewIncremental"
Sub SaveNewIncremental()
'PURPOSE: Same as FKSDONewVersion, but only saves an incremental (e.g. 01.2 --> 01.3)

Dim sPath As String
Dim sName As String, sNewName As String
Dim sExt As String
Dim sVer As String
Dim vVer As Variant
Dim sDate As String
Dim sInitials As String
Dim newFilename As String
Dim sFormat As String
Dim verIncremental As Boolean
Dim lenAfterDot As Long
Dim verAfterDot As String

    sInitials = Application.UserInitials
    sPath = ActiveDocument.Path
    If sPath = "" Then GoTo NotSavedYet
    sPath = sPath & "\"
    sName = ActiveDocument.Name
    sFormat = DateFormat(sName)
    sDate = Format(Date, sFormat) 'Format the date
    sExt = Right(sName, Len(sName) - InStrRev(sName, ".") + 1)
    sNewName = Trim(Left(sName, InStrRev(sName, "(") - 1))
    sVer = sNewName
    vVer = Split(sVer, Chr(32))
    sVer = onlyDigits(CStr(vVer(UBound(vVer))))

' Find number of digits after "." and find incremental version number, and set verIncremental
If InStrRev(sVer, ".") >= 1 Then
    verIncremental = True
    verAfterDot = Mid(sVer, InStr(1, sVer, ".") + 1)
Else
    verIncremental = False
End If
    
' Upversion document to a new incremental version.  If it's an integer version, save as new .1
If verIncremental = True Then
    sNewName = Replace(sNewName, "." & verAfterDot, "." & verAfterDot + 1)
Else
    sNewName = Replace(sNewName, sVer, Format(sVer + 0.1, "00.0"))
End If

' Name and save doc using existing spacing and quirks
If InStrRev(ActiveDocument.Name, " (") > 0 Then
    newFilename = sNewName & " (" & sInitials & Chr(32) & sDate & ")"
Else
    newFilename = sNewName & "(" & sInitials & Chr(32) & sDate & ")"
End If
Dim vShow As Variant
With Application.Dialogs(wdDialogFileSaveAs)
    .Name = ActiveDocument.Path & "\" & newFilename
    If .Show = 0 Then GoTo lbl_Exit
End With

FilePath.UpdatePathMacro


lbl_Exit:
    Exit Sub
    'Error Handler
NotSavedYet:
    MsgBox "This file has not been initially saved. " & _
           "Cannot save a new version.", vbCritical, "Not Saved To Computer"    '
    GoTo lbl_Exit

End Sub
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



