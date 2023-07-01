in file: word/vbaProject.bin - OLE stream: 'VBA/NumeralsOnly'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub NumeralsOnly()

Dim sSelect As Selection
Dim sWords As String
Dim sNumeral As String
Dim sLenSelect As Long
Dim sSelectStr As String
Set sSelect = Selection


If Len(sSelect) = 0 Then
    MsgBox "Nothing selected", vbCritical
    Exit Sub
End If

With sSelect
    'avoid inadvertently selected spaces at start and end of the selection
    .MoveEndWhile Chr(32), wdBackward
    .MoveStartWhile Chr(32)
End With

sLenSelect = Len(sSelect)
sLenRight = InStrRev(sSelect, "(") - 1
sSelectStr = Right(sSelect, sLenSelect - sLenRight)
For i = 1 To Len(sSelectStr)
        If Mid(sSelectStr, i, 1) >= "0" And Mid(sSelectStr, i, 1) <= "9" Or Mid(sSelectStr, i, 1) = "." Or Mid(sSelectStr, i, 1) = "," Or Mid(sSelectStr, i, 1) = "$" Or Mid(sSelectStr, i, 1) = "%" Then
            sNumeral = sNumeral + Mid(sSelectStr, i, 1)
        End If
Next

Debug.Print sSelect
Debug.Print Len(sSelect)
Selection.Collapse Direction:=wdCollapseStart
Selection.Delete Unit:=wdCharacter, count:=sLenSelect
Selection.TypeText sNumeral

End Sub
-------------------------------------------------------------------------------
