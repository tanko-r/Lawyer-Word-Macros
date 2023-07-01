in file: word/vbaProject.bin - OLE stream: 'VBA/RemoveManualNumbers'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit


Sub RemoveManualNumbersAndDoubleSpace()
    Dim i As Long
    Dim para As Paragraph
    
    Application.ScreenUpdating = False
    WordBasic.ToolsBulletsNumbers Replace:=0, Type:=1, Remove:=1
    With Selection.Range.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    For Each para In Selection.Range.Paragraphs
        para.Range.Select
        Selection.Collapse wdCollapseStart
        If para.Range.Characters(1) = vbTab Then para.Range.Characters(1).Delete
    Next
    Application.ScreenUpdating = True
End Sub

-------------------------------------------------------------------------------
