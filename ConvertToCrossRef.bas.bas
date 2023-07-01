in file: word/vbaProject.bin - OLE stream: 'VBA/ConvertToCrossRef'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Sub ConvertToCrossRefSimplified()

    Dim LookUp As String
    Dim RefList As Variant
    Dim i As Integer
    Dim refIndex As Integer
    Dim ref As Variant
    Dim spaceAfter As Boolean

    ' Store the selected text and check if it ends with a space
    LookUp = Trim(Selection.text)
    spaceAfter = Right(Selection.text, 1) = " "

    ' Get the list of all numbered items in the document
    RefList = ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)

    ' Loop through all numbered items
    For i = 1 To UBound(RefList)
        ' Check if the selected text matches the start of the reference item
        ref = Trim(RefList(i))
        If InStr(1, RefList(i), LookUp & " ", vbTextCompare) = 1 Or _
           InStr(1, RefList(i), LookUp & ".", vbTextCompare) = 1 Or _
           InStr(1, RefList(i), LookUp & Chr(9), vbTextCompare) = 1 Or _
           InStr(1, RefList(i), LookUp & vbCr, vbTextCompare) = 1 Then
            refIndex = i
            Exit For
        End If
    Next i

    ' If a matching reference was found, insert the cross reference
    If refIndex > 0 Then
        Selection.InsertCrossReference ReferenceType:="Numbered item", _
                                       ReferenceKind:=wdNumberFullContext, _
                                       ReferenceItem:=CStr(refIndex), _
                                       InsertAsHyperlink:=True, _
                                       IncludePosition:=False, _
                                       SeparateNumbers:=False, _
                                       SeparatorString:=" "
        If spaceAfter Then Selection.Range.InsertAfter " "
    Else
        MsgBox "A cross reference to """ & LookUp & """ couldn't be set" & vbCr & _
               "because a paragraph with that number couldn't" & vbCr & _
               "be found in the document.", _
               vbInformation, "Invalid cross reference"
    End If

End Sub



Sub ConvertToCrossRefOG()
    
    Dim RefList As Variant
    Dim LookUp As String
    Dim ref As String
    Dim spaceLoc As Integer, tabLoc As Integer
    Dim i As Integer
    Dim oRng As Range
    Dim oRngStr As String
    Dim spaceAfter As Boolean
    
    On Error GoTo ErrExit
    Set oRng = Selection.Range
    oRngStr = Selection.Range.text
    
    If oRng.Characters(Len(oRngStr)) = Chr(32) Then spaceAfter = True
    
    With oRng
        .MoveEndWhile Chr(32), wdBackward
        .MoveStartWhile Chr(32)
        ' discard leading blank spaces
        Do While (Asc(.text) = 32) And (.End > .Start)
            .MoveStart wdCharacter
        Loop
        ' discard trailing blank spaces, full stops and CRs
        Do While ((Asc(Right(.text, 1)) = 46) Or _
            (Asc(Right(.text, 1)) = 32) Or _
            (Asc(Right(.text, 1)) = 11) Or _
            (Asc(Right(.text, 1)) = 13)) And _
            (.End > .Start)
            .MoveEnd wdCharacter, -1
        Loop
        
ErrExit:
        If Len(.text) = 0 Then
            MsgBox "Please select a reference.", _
            vbExclamation, "Invalid selection"
            Exit Sub
        End If
        
        LookUp = .text
    End With
    On Error GoTo 0
    
    With ActiveDocument
        ' Use WdRefTypeNumberedItem to retrieve numbered items (including single numbers)
        RefList = .GetCrossReferenceItems(wdRefTypeNumberedItem)
        For i = UBound(RefList) To 1 Step -1
            ref = Trim(RefList(i))
            If InStr(1, ref, LookUp, vbTextCompare) = 1 Then
                spaceLoc = InStr(2, ref, " ")
                tabLoc = InStr(2, ref, Chr(9))
                If (spaceLoc = 0) Or (tabLoc = 0) Then
                    spaceLoc = IIf(spaceLoc > 0, spaceLoc, tabLoc)
                Else
                    spaceLoc = IIf(spaceLoc < tabLoc, spaceLoc, tabLoc)
                End If
                
                If LookUp = Left(ref, spaceLoc - 1) Then Exit For
            End If
        Next i
        
        If i Then
            Selection.InsertCrossReference ReferenceType:="Numbered item", _
            ReferenceKind:=wdNumberFullContext, _
            ReferenceItem:=CStr(i), _
            InsertAsHyperlink:=True, _
            IncludePosition:=False, _
            SeparateNumbers:=False, _
            SeparatorString:=" "
            If spaceAfter = True Then Selection.Range.InsertAfter (Chr(32))
        Else
            MsgBox "A cross reference to """ & LookUp & """ couldn't be set" & vbCr & _
            "because a paragraph with that number couldn't" & vbCr & _
            "be found in the document.", _
            vbInformation, "Invalid cross reference"
        End If
    End With
End Sub
-------------------------------------------------------------------------------
