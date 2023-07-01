in file: word/vbaProject.bin - OLE stream: 'VBA/StepThroughSentences'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub StepRightSentence()

    Selection.Sentences(1).Next(Unit:=wdSentence, count:=1).Select

End Sub


-------------------------------------------------------------------------------
