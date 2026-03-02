Attribute VB_Name = "HideHiddenText"
Sub HideHiddenTextOption()
    ' This macro unchecks the "Always show hidden text" option in Word
    Application.ActiveWindow.View.ShowHiddenText = False
    
End Sub

