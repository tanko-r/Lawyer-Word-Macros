Attribute VB_Name = "NewExhibit"
Sub NewExhibit()

' Insert a section break (next page) at the cursor.
Selection.Range.InsertBreak Type:=wdSectionBreakNextPage

' Unlink the footer from the previous section.
Selection.HeaderFooter.LinkToPrevious = False

' Copy the first three lines of text from the body of the previous section.
Dim previousSection As Range
Set previousSection = Selection.Sections(Selection.Range.SectionIndex - 1).Range
previousSection.Paragraphs.Range(1, 3).Copy

' Paste the copied text into the current section.
Selection.Paste

End Sub

