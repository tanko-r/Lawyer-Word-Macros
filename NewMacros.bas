Attribute VB_Name = "NewMacros"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    With Options
        .InsertedTextMark = wdInsertedTextMarkUnderline
        .InsertedTextColor =
        .DeletedTextMark = wdDeletedTextMarkStrikeThrough
        .DeletedTextColor = wdTurquoise
        .RevisedPropertiesMark = wdRevisedPropertiesMarkNone
        .RevisedPropertiesColor = wdByAuthor
        .RevisedLinesMark = wdRevisedLinesMarkOutsideBorder
        .CommentsColor = wdByAuthor
        .RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationPreserve
    End With
    ActiveWindow.View.RevisionsMode = wdMixedRevisions
    With Options
        .MoveFromTextMark = wdMoveFromTextMarkDoubleStrikeThrough
        .MoveFromTextColor = wdGreen
        .MoveToTextMark = wdMoveToTextMarkDoubleUnderline
        .MoveToTextColor = wdGreen
        .InsertedCellColor = wdCellColorLightBlue
        .MergedCellColor = wdCellColorLightYellow
        .DeletedCellColor = wdCellColorPink
        .SplitCellColor = wdCellColorLightOrange
    End With
    With ActiveDocument
        .TrackMoves = True
        .TrackFormatting = False
    End With
End Sub
