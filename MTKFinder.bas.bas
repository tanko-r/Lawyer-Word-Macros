in file: word/vbaProject.bin - OLE stream: 'VBA/MTKFinder'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 


Sub MTKFinder()

Dim StrFind As String
Dim StrRepl As String
Dim i As Long

' In StrFind and StrRepl, add words between the quote marks, separate with a comma, NO spaces
' To only highlight the found words (i.e. not replace with other words), either use StrRepl = StrFind OR use the SAME words in the same order in the StrRepl list as for the StrFind list; comment/uncomment to reflect the one you're using
' To replace a word with another and highlight it, put the new word in the StrRepl list in the SAME position as the word in the StrFind list you want to replace; comment/uncomment to reflect the one you're using

StrFind = "(the,(collectively,(1,(2,(3,(4,(5,(6,(7,(8,(9,($,latter,earlier,day of,"
StrRepl = StrFind
Set RngTxt = Selection.Range

' Set highlight color - options are listed here: https://docs.microsoft.com/en-us/office/vba/api/word.wdcolorindex
' main ones are wdYellow, wdTurquoise, wdBrightGreen, wdPink
Options.DefaultHighlightColorIndex = wdBrightGreen

Selection.HomeKey wdStory

' Clear existing formatting and settings in Find and Replace fields
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

With ActiveDocument.content.Find
  .Format = True
  .MatchWholeWord = False
  .MatchAllWordForms = False
  .MatchWildcards = False
  .Wrap = wdFindContinue
  .Forward = True
  For i = 0 To UBound(Split(StrFind, ","))
    .text = Split(StrFind, ",")(i)
    .Replacement.Highlight = True
    .Replacement.text = Split(StrRepl, ",")(i)
    .Execute Replace:=wdReplaceAll
  Next i
End With
End Sub

Sub MTKTheFinder()

Dim StrFind As String
Dim StrRepl As String
Dim i As Long

'Simplified version of MTK Finder, but only for "(the "[definition])" errors

StrFind = "(the """
StrRepl = StrFind
Set RngTxt = Selection.Range

Options.DefaultHighlightColorIndex = wdBrightGreen

Selection.HomeKey wdStory

' Clear existing formatting and settings in Find and Replace fields
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting

With ActiveDocument.content.Find
  .Format = True
  .MatchWholeWord = False
  .MatchAllWordForms = False
  .MatchWildcards = False
  .Wrap = wdFindContinue
  .Forward = True
  .text = StrFind
  .Replacement.Highlight = True
  .Replacement.text = StrFind
  .Execute Replace:=wdReplaceAll
End With
End Sub
-------------------------------------------------------------------------------
