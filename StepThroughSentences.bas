Attribute VB_Name = "StepThroughSentences"
Sub StepRightSentence()

    Selection.Sentences(1).Next(Unit:=wdSentence, count:=1).Select

End Sub


