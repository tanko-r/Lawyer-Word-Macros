Attribute VB_Name = "ChatGPTRedlineExperiment"
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

Sub ChatGPTRedline()
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
    Dim Prompt As String
    Dim oSelect As Selection
    
    Application.ScreenUpdating = False
    
    content = Selection.Range.text
    If Len(content) < 60 Then
        MsgBox ("Looks like you didn't select enough actual verbiage.  Try again")
        Exit Sub
    End If
    Prompt = "Act as a lawyer revising a contract. Revise this clause assuming you represent the party opposing the party who drafted the clause." & _
             "You should revise the langauge directly without commentary. Respond with the entire revised clause (not just the added or deleted portions). Take an aggressive stance." & _
             "Do not include linebreaks in your response.  Mark added langage like this 'ADD;;[NEW TEXT HERE];;ENDADD' and mark deleted language like this" & _
             "'DEL;;[DELETED TEXT HERE];;ENDDEL.' The clause is as follows:  "
    
    'cleaning
    content = RemoveNonASCII(content)  'ChatGPT only accepts ASCII content, so call the function that removes non-ASCII characters.
    content = Replace(content, Chr(34), Chr(39)) 'replace double quotes and smart quotes with single quote to avoid errors
    content = Replace(content, ChrW$(8220), Chr(39))
    content = Replace(content, ChrW$(8221), Chr(39))

    apiUrl = "https://api.openai.com/v1/chat/completions"
    apiKey = "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"
    requestPayload = "{""model"":""gpt-3.5-turbo"",""messages"":[{""role"":""user"",""content"":""" & Prompt & content & """}]}"
    Debug.Print Prompt & content
    'Debug.Print requestPayload

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
    Debug.Print responseText

    responseText = Trim(UnescapeString(Mid(responseText, startPos, endPos - startPos + 1)))
    'MsgBox (responseText)
    ActiveDocument.TrackRevisions = False
    Debug.Print responseText
    
    Set oSelect = Selection
    Selection.TypeText responseText
    
    ' The following (from ChatGPT itself) will find {{additions}} and ^^deletions^^ and convert them into track changes.
    Dim oDoc As Document
    Dim sFindText As String
    Dim sReplaceText As String
    Dim sPlaceholder As String
    
    ' Set active document
    Set oDoc = ActiveDocument

    ' Turn off ScreenUpdating
    Application.ScreenUpdating = False

    ' Create a new RegExp object
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
 ' Find text in ADD;; and ;;ENDADD and remove the markers
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "ADD;;(.*?);;ENDADD"
        
        Dim matches As Object
        Set matches = .Execute(responseText)
        
        Dim match As Variant
        For Each match In matches
            sFindText = match.Value
            sPlaceholder = "INSERTION GOES HERE"
                        
            ' Delete the text where it is found so that it can be reinserted as track changes
            oDoc.content.Find.Execute findText:=sFindText, ReplaceWith:=sPlaceholder, Replace:=wdReplaceOne, MatchCase:=True
            ' Turn on Track Changes and reinsert without the markers
            sReplaceText = Mid(sFindText, 6, Len(sFindText) - 13) 'Delete ADD;; and ;;ENDADD
            oDoc.TrackRevisions = True
            ' Replace the found text in the document
            oDoc.content.Find.Execute findText:=sPlaceholder, ReplaceWith:=sReplaceText, Replace:=wdReplaceOne, MatchCase:=True

            oDoc.TrackRevisions = False
        Next match
    End With
    oDoc.content.Find.Execute findText:=sPlaceholder, ReplaceWith:="", Replace:=wdReplaceAll, MatchCase:=True

    ' Find text in DEL;; and ;;ENDDEL and delete it
    With regEx
        .Pattern = "DEL;;(.*?);;ENDDEL"
        
        Set matches = .Execute(responseText)
        
        For Each match In matches
            sFindText = match.Value
            
            ' Delete the found text in the document
            oDoc.TrackRevisions = True
            oDoc.content.Find.Execute findText:=sFindText, ReplaceWith:="", Replace:=wdReplaceAll
        Next match
    End With

    ' Turn off Track Changes
    oDoc.TrackRevisions = False

    ' Turn on ScreenUpdating
    Application.ScreenUpdating = True

    ' Find and delete the custom markers returned by ChatGPT
    ' Array of special strings to delete
    Dim specialStrings() As Variant
    Dim stringIndex As Integer
    specialStrings = Array("ADD;;", ";;ENDADD", "DEL;;", ";;ENDDEL")
        For stringIndex = LBound(specialStrings) To UBound(specialStrings)
        With oDoc.content.Find
            .text = specialStrings(stringIndex)
            .Replacement.text = ""
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    Next stringIndex

    ' Turn on ScreenUpdating
    Application.ScreenUpdating = True

End Sub

Sub UpdateText()

    Dim responseText As String
    Dim oDoc As Document
    Dim sFindText As String
    Dim sReplaceText As String
    
    ' Your responseText here
    responseText = Selection.text

        
    ' Set active document
    Set oDoc = ActiveDocument

    ' Turn off ScreenUpdating
    Application.ScreenUpdating = False

    ' Turn on Track Changes
    oDoc.TrackRevisions = True
    
    ' Create a new RegExp object
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' Find text in ADD;; and ;;ENDADD and remove the markers
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "ADD;;(.*?);;ENDADD"
        
        Dim matches As Object
        Set matches = .Execute(responseText)
        
        Dim match As Variant
        For Each match In matches
            sFindText = match.Value
            sReplaceText = Mid(sFindText, 5, Len(sFindText) - 11)
            
            ' Replace the found text in the document
            oDoc.content.Find.Execute findText:=sFindText, ReplaceWith:=sReplaceText, Replace:=wdReplaceAll, MatchCase:=True
        Next match
    End With
    
    ' Find text in DEL;; and ;;ENDDEL and delete it
    With regEx
        .Pattern = "DEL;;(.*?);;ENDDEL"
        
        Set matches = .Execute(responseText)
        
        For Each match In matches
            sFindText = match.Value
            
            ' Delete the found text in the document
            oDoc.content.Find.Execute findText:=sFindText, ReplaceWith:="", Replace:=wdReplaceAll
        Next match
    End With

    ' Turn off Track Changes
    oDoc.TrackRevisions = False

    ' Turn on ScreenUpdating
    Application.ScreenUpdating = True

End Sub

