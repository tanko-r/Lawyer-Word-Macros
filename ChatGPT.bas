Attribute VB_Name = "ChatGPT"
'==================================================================================================
' MS Word VBA Macro to Interact with Google Gemini API
'
' Author: Google Gemini
' Version: 1.0
' Date: June 12, 2024
'
' Description:
' This macro sends a user-provided prompt, combined with predefined instructions,
' to the Google Gemini Pro API and inserts the response into the active Word document.
'
' --- PREREQUISITES ---
' 1. You must have a Google AI Studio API key.
' 2. You need to enable the "Microsoft XML, v6.0" reference in the VBA editor.
'    To do this: In the VBA Editor, go to "Tools" -> "References..." and check the box for
'    "Microsoft XML, v6.0". If you don't see v6.0, choose v3.0.
'
'==================================================================================================

' --- Paste your API Key here ---
Private Const API_KEY As String = "AIzaSyDyNx89DoJ-l9R_zgPlIQq0qO2RxrVSQ9g"

' --- Customize your pre-baked instructions here ---
' This text will be prepended to every prompt you send to Gemini.
Private Const PRE_BAKED_INSTRUCTIONS As String = "You are an expert editor and writing assistant. " & _
"Analyze the following text and perform the requested action. Your response should be " & _
"professional, clear, and ready to be inserted directly into a document. " & _
"Do not include any preamble like 'Sure, here is the response:'. Just provide the text. " & _
"The user's request is: "

' --- Main Subroutine ---
' This is the macro you will run in Word (e.g., by pressing Alt+F8).
Sub RunGeminiPrompt()
    ' --- Error handling ---
    If API_KEY = "PASTE_YOUR_API_KEY_HERE" Then
        MsgBox "Error: API Key is not set." & vbCrLf & vbCrLf & _
               "Please open the VBA editor (Alt+F11), find the 'RunGeminiPrompt' macro, " & _
               "and replace 'PASTE_YOUR_API_KEY_HERE' with your actual Google Gemini API key.", _
               vbCritical, "API Key Missing"
        Exit Sub
    End If

    ' --- Variable Declaration ---
    Dim userInput As String
    Dim fullPrompt As String
    Dim geminiResponse As String

    ' --- Get User Input ---
    ' First, check if the user has selected any text. Use that as the default input.
    Dim defaultInput As String
    If Selection.Type = wdSelectionNormal Then
        defaultInput = Selection.text
    Else
        defaultInput = ""
    End If

    userInput = InputBox("Enter your prompt for Gemini:", "Gemini Prompt", defaultInput)

    ' Exit if the user clicks Cancel or enters nothing.
    If userInput = "" Then
        Exit Sub
    End If

    ' --- Combine Instructions and User Input ---
    fullPrompt = PRE_BAKED_INSTRUCTIONS & userInput

    ' --- Show a status message in the Word status bar ---
    Application.StatusBar = "Contacting Google Gemini... Please wait."

    ' --- Call the API and get the response ---
    On Error GoTo ApiErrorHandler
    geminiResponse = GetGeminiApiResponse(fullPrompt)
    On Error GoTo 0 ' Reset error handler

    ' --- Process and Insert Response ---
    If geminiResponse <> "" Then
        ' Clean the response to remove potential leading/trailing whitespace
        geminiResponse = Trim(geminiResponse)
        ' Insert the text at the current cursor position
        Selection.TypeText text:=geminiResponse
    Else
        MsgBox "The API returned an empty or invalid response. Please check your prompt or API key.", _
               vbExclamation, "No Response"
    End If

    ' --- Clear the status bar ---
    Application.StatusBar = ""
    Exit Sub

' --- Error Handling Block ---
ApiErrorHandler:
    MsgBox "An error occurred while contacting the API." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "API Communication Error"
    Application.StatusBar = "" ' Clear status bar on error
End Sub


' --- Helper Function to Call the API ---
Private Function GetGeminiApiResponse(ByVal prompt As String) As String
    ' --- Constants for the API ---
    Const API_URL As String = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key="
    
    ' --- Variable Declaration ---
    Dim http As Object ' MSXML2.XMLHTTP60
    Dim jsonBody As String
    Dim responseText As String
    Dim extractedText As String

    ' --- Create the HTTP object ---
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP60")
    If Err.Number <> 0 Then
        MsgBox "Could not create the 'Microsoft XML, v6.0' object." & vbCrLf & _
               "Please ensure it is enabled in Tools -> References.", vbCritical, "Object Creation Failed"
        GetGeminiApiResponse = ""
        Exit Function
    End If
    On Error GoTo 0

    ' --- Build the JSON payload ---
    ' We need to escape special characters in the prompt for valid JSON.
    jsonBody = "{""contents"":[{""parts"":[{""text"": """ & EscapeJsonString(prompt) & """}]}]}"

    ' --- Send the request ---
    http.Open "POST", API_URL & API_KEY, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send jsonBody

    ' --- Check the response ---
    If http.Status = 200 Then
        responseText = http.responseText
        ' Simple parsing to find the generated text.
        ' A full JSON parser is safer but this is more portable for a simple macro.
        extractedText = ParseSimpleJson(responseText, "text")
        GetGeminiApiResponse = extractedText
    Else
        ' Provide a more detailed error message
        MsgBox "API Error: " & http.Status & " - " & http.StatusText & vbCrLf & vbCrLf & _
               "Response: " & http.responseText, vbCritical, "API Error"
        GetGeminiApiResponse = ""
    End If

    Set http = Nothing
End Function


' --- Helper Function to escape characters for JSON compatibility ---
Private Function EscapeJsonString(ByVal text As String) As String
    text = Replace(text, "\", "\\") ' Escape backslashes
    text = Replace(text, """", "\""") ' Escape quotes
    text = Replace(text, vbCr, "\r")   ' Escape carriage returns
    text = Replace(text, vbLf, "\n")   ' Escape line feeds
    text = Replace(text, vbTab, "\t")  ' Escape tabs
    EscapeJsonString = text
End Function


' --- Helper Function for basic JSON parsing ---
' This is a simple, non-robust parser to avoid external dependencies.
' It finds the *first* instance of a key and returns its value.
Private Function ParseSimpleJson(ByVal jsonString As String, ByVal key As String) As String
    Dim keyPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim valueStart As Long

    ' The pattern we expect is: "key": "value"
    keyPattern = """" & key & """: """
    startPos = InStr(1, jsonString, keyPattern)

    If startPos > 0 Then
        ' Find the start of the actual value (after the key pattern)
        valueStart = startPos + Len(keyPattern)
        ' Find the closing quote of the value
        endPos = InStr(valueStart, jsonString, """")
        
        If endPos > 0 Then
            ParseSimpleJson = Mid(jsonString, valueStart, endPos - valueStart)
        Else
            ParseSimpleJson = "" ' Value not found or malformed
        End If
    Else
        ParseSimpleJson = "" ' Key not found
    End If
End Function



