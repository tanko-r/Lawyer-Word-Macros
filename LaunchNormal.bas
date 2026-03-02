Attribute VB_Name = "LaunchNormal"
Sub LaunchNormal()


Application.NormalTemplate.OpenAsDocument

End Sub


Sub LaunchGlobalTemplate()

    Dim temp As Template
    
    'For Each temp In Application.Templates
    '    Debug.Print temp.FullName
    'Next temp
    
    For Each temp In Application.Templates
        If InStr(1, temp.Name, "DSR") > 0 Then
            temp.OpenAsDocument
            Exit Sub
        End If
    Next temp

End Sub

