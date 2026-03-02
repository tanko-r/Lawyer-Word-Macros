Attribute VB_Name = "ShowHiddenText"

Sub ToggleHiddenText()
    ' Toggle the state of the ShowHiddenText property for the active window's view.
    ActiveWindow.View.ShowHiddenText = Not ActiveWindow.View.ShowHiddenText
End Sub
