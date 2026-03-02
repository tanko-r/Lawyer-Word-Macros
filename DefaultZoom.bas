Attribute VB_Name = "DefaultZoom"
Public Sub AutoOpen()
    ActiveWindow.ActivePane.View.Zoom.Percentage = 100
    ActiveWindow.ActivePane.View.Zoom.PageColumns = 2
End Sub
