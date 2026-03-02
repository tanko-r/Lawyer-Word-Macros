Attribute VB_Name = "DeleteFooterBox"
Sub DeleteFooterBox()

  Dim objSect As Section
  Dim objHdrFtr As HeaderFooter
  Dim objShp As Shape
  Dim strText As String

  For Each objSect In ActiveDocument.Sections
    For Each objHdrFtr In objSect.Footers
      For Each objShp In objHdrFtr.Shapes
        If objShp.Type = msoTextBox Then
          strText = objShp.TextFrame.TextRange.text
          
          ' Check for the specific format
          If (strText Like "#########.#*" Or strText Like "#########.##*" _
                Or strText Like "##########.#*" Or strText Like "##########.##*") Then  ' Updated doc numbers after 6/2025 iManage upgrade
            objShp.Delete
          End If
        End If
      Next objShp
    Next objHdrFtr
  Next objSect

End Sub
