Attribute VB_Name = "InsertPageNumber"
Sub InsertPageNumber()
  ' Inserts the PAGE field at the current cursor location.

  Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, text:="PAGE"
End Sub
