Attribute VB_Name = "RemoveLineBreaks"


Sub RemoveLineBreaks()
'A basic Word macro coded by Greg Maxey
Dim oRng As Word.Range
Set oRng = Selection.Range
  If oRng.Characters.Last = Chr(13) Or oRng.Characters.Last = Chr(11) Then
    oRng.End = oRng.End - 1
  End If
  oRng.text = Replace(Replace(oRng.text, Chr(11), " "), Chr(13), " ")
lbl_Exit:
  Exit Sub
End Sub
