Attribute VB_Name = "ChatGPT"
Option Explicit

' This is a macro that will communicate with the ChatGPT API in Microsoft Word.
'
'Just copy/paste this macro into Word following instructions in the Readme.md file.
' Don't forget to change the API key to your own.
' Author: Johann Dowa
' http://github.com/jddev273/chatgpt-word-macro

Function UnescapeString(ByVal str As String) As String
    Dim i As Integer
    Dim output As String
    For i = 1 To Len(str)
        If Mid(str, i, 2) = "\\" Then
            output = output & "\"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\/" Then
            output = output & "/"
            i = i + 1
        ElseIf Mid(str, i, 2) = "\n" Then
            output = output & vbCrLf
            i = i + 1
        ElseIf Mid(str, i, 2) = "\r" Then
            output = output & vbCr
            i = i + 1
        ElseIf Mid(str, i, 2) = "\t" Then
            output = output & vbTab
            i = i + 1
        ElseIf Mid(str, i, 2) = "\" & Chr(34) Then
            output = output & """"
            i = i + 1
        Else
            output = output & Mid(str, i, 1)
        End If
    Next i
    UnescapeString = output
End Function

Sub ChatGPT()
    '
    ' ChatGPT Macro
    '
    
    Dim apiUrl As String
    Dim requestPayload As String
    Dim apiKey As String
    Dim httpRequest As Object
    Dim responseText As String
    Dim content As String
    Dim startIndex As Integer
    Dim endIndex As Integer
    Dim startPos As Long
    Dim endPos As Long
    Dim oText As String
    
    content = InputBox("Enter the prompt for ChatGPT.", "Content")
    
    apiUrl = "https://api.openai.com/v1/chat/completions"
    apiKey = "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"
    requestPayload = "{""model"":""gpt-3.5-turbo-0301"",""messages"":[{""role"":""user"",""content"":""" & content & """}]}"
    
    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open "POST", apiUrl, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "Authorization", "Bearer " & apiKey
    On Error Resume Next
    httpRequest.send requestPayload
    On Error GoTo 0
    
    If httpRequest.Status <> 200 Then
        MsgBox "Error: " & httpRequest.Status & " " & httpRequest.StatusText
        Exit Sub
    End If
        
    responseText = httpRequest.responseText
    startPos = InStr(responseText, """content"":""") + 11
    endPos = InStr(responseText, """},""") - 1

    responseText = Trim(UnescapeString(Mid(responseText, startPos, endPos - startPos + 1)))
   MsgBox (responseText)
   ' Selection.HomeKey wdStory
   Selection.TypeText responseText
    
    Set httpRequest = Nothing
    

End Sub

Public Function RemoveNonASCII(inputStr As String) As String
    Dim outputStr As String
    Dim char As String
    Dim i As Integer

    For i = 1 To Len(inputStr)
        char = Mid(inputStr, i, 1)
        If Asc(char) >= 32 And Asc(char) <= 126 Then
            outputStr = outputStr & char
        End If
    Next i

    RemoveNonASCII = outputStr
End Function

Sub ASCIIfi_Document_Hide()
Application.ScreenUpdating = False
Dim i As Long
With ActiveDocument.Range
  .Font.Hidden = True
  For i = 1 To 255
    With .Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .text = Chr(i)
      .Replacement.text = "^&"
      .Font.Hidden = True
      .Replacement.Font.Hidden = False
      .Forward = True
      .Format = True
      .MatchWildcards = False
      .Wrap = wdFindContinue
      .Execute Replace:=wdReplaceAll
    End With
    DoEvents
  Next
  .Fields.Update
  DoEvents
End With
Application.ScreenUpdating = True
End Sub

Function CleanText(text As String) As String
    ' Remove control characters and escape double quotes
    Dim result As String
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String
        ch = Mid(text, i, 1)
        If Asc(ch) >= 32 And Asc(ch) <> 127 Then ' non-control character
            If ch = """" Then
                result = result & "\"""
            Else
                result = result & ch
            End If
        End If
    Next i
    CleanText = result
End Function

