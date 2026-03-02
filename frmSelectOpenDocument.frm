VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectOpenDocument 
   Caption         =   "Select the old document to compare against"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   OleObjectBlob   =   "frmSelectOpenDocument.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectOpenDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- Code for UserForm: frmSelectOpenDocument ---

Option Explicit

Public SelectedDocumentFullName As String
Public BrowseFileSystem As Boolean
Public UserCancelled As Boolean


Private Sub UserForm_Initialize()
    Dim doc As Document

    ' Initialize public properties
    Me.BrowseFileSystem = False
    Me.UserCancelled = False
    Me.SelectedDocumentFullName = ""

    Me.Caption = "Select Original Document"
    Me.lstOpenDocs.Clear

    If Application.Documents.count > 0 Then
        For Each doc In Application.Documents
            ' Exclude the document that will be DocNew (passed via CurrentDocNewFullName)
            If doc.FullName <> FastCompare.CurrentDocNewFullName Then
                Me.lstOpenDocs.AddItem doc.Name ' Display friendly name
            End If
        Next
        Debug.Print FastCompare.CurrentDocNewFullName
        
        If Me.lstOpenDocs.listCount > 0 Then
            Me.lstOpenDocs.ListIndex = 0 ' Select the first item by default
            Me.cmdSelectFromList.Enabled = True
        Else
            ' No *other* documents open
            Me.cmdSelectFromList.Enabled = False
            Me.cmdBrowseFiles.Caption = "Browse Filesystem... (No other documents open)"
        End If
    Else
        ' This case (no documents open at all) shouldn't happen if FastCompare starts with an ActiveDocument
        Me.cmdSelectFromList.Enabled = False
        Me.cmdBrowseFiles.Enabled = False ' No sensible action if no documents are open at all
        Me.cmdBrowseFiles.Caption = "Browse Filesystem... (Error: No documents open)"
    End If
    
    ' Default captions (can be overridden above if no other docs)
    If Me.cmdSelectFromList.Enabled Then Me.cmdSelectFromList.Caption = "Use Selected Document"
    If Me.cmdBrowseFiles.Caption = "Browse Filesystem..." And Me.lstOpenDocs.listCount = 0 Then
         ' Keep the more informative caption if no other docs
    Else
        Me.cmdBrowseFiles.Caption = "Browse Filesystem..."
    End If
    Me.cmdCancel.Caption = "Cancel"
End Sub

Private Sub cmdSelectFromList_Click()
    Dim selectedIdx As Long
    Dim doc As Document
    Dim tempSelectedName As String
    
    selectedIdx = Me.lstOpenDocs.ListIndex
    If selectedIdx > -1 Then
        tempSelectedName = Me.lstOpenDocs.List(selectedIdx) ' This is doc.Name
        
        ' Find the corresponding FullName by iterating through open documents
        For Each doc In Application.Documents
            If doc.Name = tempSelectedName And doc.FullName <> FastCompare.CurrentDocNewFullName Then
                Me.SelectedDocumentFullName = doc.FullName
                Exit For
            End If
        Next

        If Me.SelectedDocumentFullName = "" Then
            MsgBox "Error: It looks like you are trying to compare the same document.  Try again.", vbExclamation
        Else
            Me.Hide
        End If
    Else
        MsgBox "Please select a document from the list or choose to browse.", vbInformation
    End If
End Sub

Private Sub cmdBrowseFiles_Click()
    Me.BrowseFileSystem = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Me.UserCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If the user clicks the 'X' button on the form's title bar
    If CloseMode = vbFormControlMenu Then
        Me.UserCancelled = True
        ' Me.Hide will be called automatically after this event
    End If
End Sub

