Attribute VB_Name = "ConvertToCrossRef"
Option Explicit

Sub ConvertToCrossRef()

    Dim RefList As Variant
    Dim LookUp As String
    Dim Ref As String
    Dim s As Integer, t As Integer
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
        ' Use WdRefTypeHeading to retrieve Headings
        RefList = .GetCrossReferenceItems(wdRefTypeNumberedItem)
        For i = UBound(RefList) To 1 Step -1
            Ref = Trim(RefList(i))
            If InStr(1, Ref, LookUp, vbTextCompare) = 1 Then
                s = InStr(2, Ref, " ")
                t = InStr(2, Ref, Chr(9))
                If (s = 0) Or (t = 0) Then
                    s = IIf(s > 0, s, t)
                Else
                    s = IIf(s < t, s, t)
                End If
                If LookUp = Left(Ref, s - 1) Then Exit For
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

