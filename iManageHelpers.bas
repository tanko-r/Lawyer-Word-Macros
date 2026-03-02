Attribute VB_Name = "iManageHelpers"
Option Explicit

Public Function RemoveDocNo(StrDocName As String) As String
' After checking if this is an iManage doc using iManageTest, use this function to delete the document number from the filename or window caption.

RemoveDocNo = Left(StrDocName, InStrRev(StrDocName, "#") - 1)


End Function

Public Function iManTestCaption(ByRef docCaption As String) As Boolean

 Dim intWindowCaptionLength As Integer
  Dim strWindowCaption As String
  Dim strWindowCaptionEnd As String
  Dim strFileName As String
  Dim strFileExtension As String
  Dim strFileExtensionLength As Integer
    
  strWindowCaptionEnd = Right(docCaption, 14) ' Adjust length to 14

  If docCaption Like "*#########v##*" Then
    iManTestCaption = True
  ElseIf docCaption Like "*#########v#*" Then
    iManTestCaption = True
  Else
    iManTestCaption = False
  End If
  
End Function
' THIS PROBABLY APPLIES ONLY TO KLG FILES.  I SHOULD REVISE FOR POLSINELLI.
Public Function CaptionCleanup(ByRef docCaption As String) As String
'Probably not helpful.  Need to rewrite this to process file name rather than captions.  Check how it is used in SaveNewVersion
Dim baseFilename As String      ' Raw filename part, from iManage processing or full caption (no extension)
' Intermediate parsing variables
Dim posFirstOpenParen As Long
Dim posLastDot As Long
Dim fileExtension As String

        baseFilename = Trim(Left(docCaption, InStr(1, docCaption, "#") - 1))
        ' If baseFilename still ends with .doc(x/m) because it was part of name before #
        posLastDot = InStrRev(baseFilename, ".")
        If posLastDot > 0 Then
            fileExtension = Mid(baseFilename, posLastDot)
            Select Case LCase(fileExtension)
                Case ".doc", ".docx", ".docm"
                    baseFilename = Trim(Left(baseFilename, posLastDot - 1))
            End Select
        End If
    Else
        baseFilename = docCaption
        ' Strip common Word extension if present at the very end of a non-iManage caption
        posLastDot = InStrRev(baseFilename, ".")
        If posLastDot > 0 Then
            ' Ensure the dot is not part of something like (12345.6)
            ' A simple check: if there's no open paren after the last dot, it's likely an extension.
            If InStr(posLastDot, baseFilename, "(") = 0 Then
                fileExtension = Mid(baseFilename, posLastDot)
                Select Case LCase(fileExtension)
                    Case ".doc", ".docx", ".docm"
                        baseFilename = Trim(Left(baseFilename, posLastDot - 1))
                End Select
            End If
        End If
End Function


' Takes a window caption string as input and returns the descriptive
' part of the filename, assuming the new Polsinelli iManage format.
'
' PARAMETERS:
'   sourceCaption (String): The full window title to be parsed.
'
' RETURNS:
'   (String): The extracted descriptive filename, or the original
'             string if the iManage delimiter is not found.
' ==========================================================================
Public Function iManPolsinelliCaptionCleanup(ByVal sourceCaption As String) As String
    ' --- VARIABLES ---
    Dim cleanedName As String     ' To hold the final, extracted filename
    Dim delimiterPos As Long      ' To store the position of our search text
    
    ' The unique text that separates the iManage info from the descriptive filename.
    Const IMANAGE_DELIMITER As String = "<ACTIVE> - "

    ' --- 1. Find the delimiter in the provided source string ---
    delimiterPos = InStr(1, sourceCaption, IMANAGE_DELIMITER)

    ' --- 2. Extract the filename ---
    If delimiterPos > 0 Then
        ' If the delimiter is found, extract the text that comes after it.
        cleanedName = Trim(Mid(sourceCaption, delimiterPos + Len(IMANAGE_DELIMITER)))
    Else
        ' If the delimiter isn't found, return the original string.
        ' This makes the function safe to use on non-iManage filenames.
        cleanedName = sourceCaption
    End If

    ' --- 3. Return the result ---
    ' Assign the result to the function's name.
    iManPolsinelliCaptionCleanup = cleanedName
End Function

Function ListProperties()
Dim prop As Object

For Each prop In ActiveDocument.BuiltInDocumentProperties
    Debug.Print "Property" & "_____" & prop.Name & " \ "; prop.value; ""
Next prop

End Function
