in file: word/vbaProject.bin - OLE stream: 'VBA/FilePath'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Sub FilePathMacro()
' Inserts filepath as plaintext plus dynamic filename.  This is to preserve the FKSDO filepath when converting to PDF or sending to outsiders.

    Selection.ClearFormatting
    Selection.Font.Name = "Tahoma"
    Selection.Font.Size = 6
    Selection.Font.Bold = False
    Selection.Font.AllCaps = False
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.ParagraphFormat.SpaceBefore = 6
    Selection.ParagraphFormat.spaceAfter = 0
    Selection.TypeText (Application.ActiveDocument.Path & "\")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldFileName, PreserveFormatting:=True
    'Selection.HomeKey Unit:=wdLine
    
'   Selection.Fields.Unlink 'Save for later.  Need to figure out how to loop through footers before this will work correctly.  Otherwise only one of the footers will update.
    
End Sub

Sub AddPath()
'Look for a footer and then prompt to add one, if missing
Dim oSec As Section
Dim oCursor As Range

Application.ScreenUpdating = False


Set oCursor = Selection.Range 'get current position
'Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst 'For use later when looping through sections.


'    For Each oSec In Application.ActiveDocument.Sections ' For use later--to loop through sections and add footer in all sections.
'           Application.ActiveDocument.Range.GoTo ' Incomplete, for use later-- to loop through sections.
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "\"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        If Selection.Find.Execute = True Then
            FilePath.UpdatePathMacro
'            Selection.HomeKey
'            Selection.EndKey Unit:=wdLine, Extend:=wdExtend
'            Selection.Delete
'            Selection.TypeText Chr(13)
'            FilePath.FilePathMacro
        End If
        If Selection.Find.Execute = False Then
            If MsgBox("Do you want to add a filepath footer?", vbYesNo) = vbYes Then
                Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst
                ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
                Selection.EndKey Unit:=wdStory, Extend:=wdMove
                Selection.TypeText Chr(13)
                FilePath.FilePathMacro
            End If
         End If
    End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Get view out of the footer
    ActiveDocument.Save
'    Next oSec

' oCursor.Select ' return cursor to original position  ''Currently unused because it changes the view to Draft for some reason.
Application.ScreenUpdating = True
End Sub

Sub UpdatePathMacro()
    Dim wFind As Find
    Dim wDoc As Document
    Dim newFooter As Long
    Dim pathExists As Boolean
    Dim footerExists As Boolean
    Dim footerRange As Range
    Dim sec As Section
    Dim loopBreak As Long
    Dim runFtrUPdate As Boolean
    Dim currentPosition As Range
    Dim oView As WdViewType
    Dim oldPath As String
    
    Application.ScreenUpdating = False
    
    Set currentPosition = Selection.Range 'store current cursor position
    oView = ActiveDocument.ActiveWindow.View.Type 'store the user's view because we have to use print layout
     
    Set wDoc = Application.ActiveDocument
    For Each sec In wDoc.Sections          'Start by looping through each footer and update fields.  At minimum, all existing footers should be updated.
        sec.Footers(wdHeaderFooterPrimary).Range.Fields.Update
        sec.Footers(wdHeaderFooterFirstPage).Range.Fields.Update 'Gotta update these separately, annoyingly.
        sec.Footers(wdHeaderFooterEvenPages).Range.Fields.Update
    Next sec

    If FooterCheck() <> "" Then
        footerExists = True                             'If the footercheck returned any results, then a footer exists.
        runFtrUPdate = True
    End If
                                                        
    If footerExists <> True Then                        'If a footer does not exist, then ask the user.
        newFooter = MsgBox("Do you want to add a filepath footer?", vbYesNo)
        If newFooter <> 6 Then GoTo ErrExit             'if the user does not agree to a new footer, exit the sub
        If newFooter = 6 Then runFtrUPdate = True
    End If
    
    If runFtrUPdate = True Then                              'If User wants a enw footer, add a new footer in each section
        ' Replace footer in first section
        Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst
        ActiveDocument.ActiveWindow.View.Type = wdPrintView
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter 'Go to the footer in the first section
        Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    
        'Start the find process
        Selection.Find.ClearFormatting
    
        Set wFind = Selection.Find                   'Look for existing footers and we will delete each footer path, then reinsert.
        With wFind                                   'Search down through the footers.  When accessing Find via Selection, the Selection is moved.
            .text = "\"                              'look for the backslash in a filepath, but note that this maybe overinclusive -- might be other kinds of text
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        
        If wFind.Execute = True Then                 'Add path to the first section
            pathExists = True
            Selection.HomeKey
            Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
            Selection.MoveEndWhile Cset:=Chr(13), count:=-1      'Don't delete the paragraph break
            Selection.Delete
            Selection.MoveLeft count:=1
            Selection.EndOf Unit:=wdStory, Extend:=wdMove 'go to the end of the footer
            Selection.TypeText " "
            Selection.MoveLeft count:=1  'The FILENAME field glitches in the FilePathMacro glitches if it is at the very end.
            FilePath.FilePathMacro
'            Selection.EndOf Unit:=wdStory, Extend:=wdMove
        Else
            Selection.EndOf Unit:=wdStory, Extend:=wdMove 'go to the end of the footer
            Selection.InsertAfter vbCr & " "
            Selection.Collapse wdCollapseEnd
            Selection.MoveLeft count:=1  'The FILENAME field glitches in the FilePathMacro glitches if it is at the very end.
            FilePath.FilePathMacro
            'Selection.TypeText Chr(13)
            'Dim rng As Range
            'Dim remainingText As Range
        End If
    
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' Not sure why but this is necessary to actually move to the next section.
        Selection.GoTo What:=wdGoToSection, Which:=wdGoToNext
        ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
   
        'Loop through each subsequent section in the document and find the \ character.  Delete all existing footers.
        Do

            If loopBreak > wDoc.Sections.count Then GoTo ErrExit     'Sometimes this seems to get caught in a loop.  Break the loop after total sections and exit.
            Selection.StartOf Unit:=wdStory, Extend:=wdMove
            If wFind.Execute = True Then
                pathExists = True
                Selection.HomeKey
                Selection.EndOf Unit:=wdParagraph, Extend:=wdExtend
                Selection.MoveEndWhile Cset:=Chr(13), count:=-1      'Don't delete the paragraph break
                Selection.Delete
                Selection.MoveLeft count:=1
                Selection.EndOf Unit:=wdParagraph, Extend:=wdMove 'go to the end of the line (inc. page number) and add a par break
                'Selection.TypeText Chr(13)
                Selection.TypeText " "
                Selection.MoveLeft count:=1  'The FILENAME field glitches in the FilePathMacro glitches if it is at the very end.
                FilePath.FilePathMacro
                Dim remainingText As Range
                Dim rng As Range
                Set rng = Selection.Range
                Set remainingText = rng.Document.Range(rng.End, rng.Document.Range.End)
                Selection.EndOf Unit:=wdStory, Extend:=wdMove
            Else
                Selection.EndOf Unit:=wdParagraph, Extend:=wdMove 'go to the end of the line (inc. page number) and add a par break
                Selection.InsertAfter (Chr(13) & " ")
                FilePath.FilePathMacro
                Selection.Collapse wdCollapseEnd
                Selection.TypeText vbLf
                'Dim rng As Range
                'Dim remainingText As Range
                'Set rng = Selection.Range
                'Set remainingText = rng.Document.Range(rng.End, rng.Document.Range.End)
                'If Len(remainingText.text) > 0 Then rng.InsertAfter vbCr  ' If there is text after the inserted text, insert a line break after the inserted text
            End If

            If Selection.Information(wdActiveEndSectionNumber) <> wDoc.Sections.count Then 'If not the last section, then go the next section
                ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument ' Not sure why but this is necessary to actually move to the next section.
                Selection.GoTo What:=wdGoToSection, Which:=wdGoToNext
                ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
            Else                                 'if the last section, exit the loop.
                ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
                Exit Do
            End If

            pathExists = False                      'pathExists should default to false unless found on the next loop, though this is probably unnecessary.
            loopBreak = loopBreak + 1               'increment the loop break
        Loop


        ' _______________________EXPERIMENTAL!!___________________________________
        ' Following is a new concept, which will create a unique docID based on the client/matter, the doc name, and the doc version
        ' ________________________END_____________________________________________


    End If
    
ErrExit:
    currentPosition.Select                                'return cursor to original position
    Selection.Find.text = ""                              'Clear the find text
    Selection.Find.Wrap = wdFindContinue                  'Continue is usually the expected Find behavior
    ActiveDocument.ActiveWindow.View.Type = oView              'restore the user's original view
    Application.ScreenUpdating = True
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'get out of the footer again
End Sub

Sub OpenFolder()
    Dim FolderName As String
    Dim docPath As String
    docPath = ActiveDocument.Path
    
    Application.ScreenUpdating = False
    
    If InStr(docPath, "G:") = 0 Then
        If InStr(docPath, "data") = 0 Then
            FolderName = FilePath.FooterCheck
        Else
            FolderName = docPath
        End If
    End If
    
    If FolderName = "" Then FolderName = docPath & "\"
    Shell "C:\WINDOWS\explorer.exe """ & FolderName & "", vbNormalFocus
    
    Application.ScreenUpdating = True
End Sub

Public Function FooterCheck() As String
'Returns a string with the folder path in the footer, without the filename
Dim inFolder As String
Dim ftrTarget As String
Dim currentPosition As Range

'Application.ScreenUpdating = False

    'keep track of the current position
    Set currentPosition = Selection.Range 'pick up current cursor position
    Selection.GoTo What:=wdGoToSection, Which:=wdGoToFirst 'Look for a footer on the first page
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
    Selection.Find.ClearFormatting
    With Selection.Find
        .text = "\"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    If Selection.Find.Execute = True Then
    'Select whole filepath and trim down to folder path.
        Selection.EndKey wdLine, wdMove
        Selection.HomeKey wdLine, wdExtend
        inFolder = Trim(Left(Selection, InStrRev(Selection, "\")))
    End If
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument 'Get view out of the footer
    
    'return to original cursor position.
    currentPosition.Select

FooterCheck = inFolder
End Function




-------------------------------------------------------------------------------
