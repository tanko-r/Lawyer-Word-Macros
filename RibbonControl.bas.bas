in file: word/vbaProject.bin - OLE stream: 'VBA/RibbonControl'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Public MyRibbon As IRibbonUI
Public SummaryCheckboxValue As Boolean ' A variable to hold the checkbox status


Sub RibbonOnLoad(ribbon As IRibbonUI)
    'Create a ribbon instance for use in this project
    Set MyRibbon = ribbon
    SummaryCheckboxValue = True
End Sub

Sub MyBtnMacro(ByVal control As IRibbonControl)
Select Case control.ID
    Case Is = "MakeDef"
        MakeDefinition.MakeDefinition
    Case Is = "BoldDefTerms"
        BoldDefinedTerms.BoldDefinedTerms
    'Case Is = "CrossRef"
        'ConvertToCrossRef.ConvertToCrossRef
    Case Is = "HowDefined"
        TermDefined.FindDef
    Case Is = "RmvLnBrks"
        RemoveLineBreaks.RemoveLineBreaks
    Case Is = "NxtSnt"
        StepThroughSentences.StepRightSentence
    Case Is = "FilePath"
        FilePath.FilePathMacro
    Case Is = "FastCompare"
        FastCompare.FastCompare
    Case Is = "FKSDOSave"
        FKSDOSaveAs.SaveAsFKSDOFile
    Case Is = "SaveNewVer"
        FKSDONewVersion.SaveNewVersion_Word
    Case Is = "SaveInSeq"
        FKSDOSaveInSequence.SaveInSequence
    Case Is = "HighlightTerm"
        HilightTerm.HighlightSelectedTerm
    Case Is = "BigHighlightTerm"
        HilightTerm.HighlightSelectedTerm
    Case Is = "UnhighlightTerm"
        HilightTerm.UnHighlightSelectedTerm
    Case Is = "HighlightBrackets"
        HighlightBrackets.HighlightBrackets
    Case Is = "CopyToForms"
        CopyToForms.CopyToForms
    Case Is = "NumeralsOnly"
        NumeralsOnly.NumeralsOnly
    Case Is = "FKSDONote"
        FKSDONote.FKSDONoteMacro
    Case Is = "PrintChangePages"
        PrintChangePages.PrintOnlyMarkupPages
    Case Is = "DontAddSpace"
        DontAddSpace.DontAddSpaceMacro
    Case Is = "UpdatePath"
        FilePath.UpdatePathMacro
    Case Is = "NewIncremental"
        FKSDONewIncremental.SaveNewIncremental
    Case Is = "EmailDoc"
        EmailDocument.EmailDocumentMacro
    Case Is = "OpenFolder"
        FilePath.OpenFolder
    Case Is = "MTKFastCompare"
        FastCompare.MTKFastCompare
        
End Select
End Sub


'Callback for DropDown onAction
Sub myDDMacro(ByVal control As IRibbonControl, selectedID As String, _
               selectedIndex As Integer)
Select Case selectedIndex
  Case 1
    Macros.Macro1
  Case 2
    Macros.Macro2
  Case 3
    Macros.Macro3
End Select
'Force the ribbon to restore the control to original state
MyRibbon.InvalidateControl control.ID
End Sub
'Callback for DropDown GetItemCount
Sub GetItemCount(ByVal control As IRibbonControl, ByRef count)
count = 4
End Sub
'Callback for DropDown GetItemLabel
Sub GetItemLabel(ByVal control As IRibbonControl, _
    Index As Integer, ByRef label)
label = Choose(Index + 1, "Select from list", "Macro 1", "Macro 2", "Macro 3")
End Sub
'Callback DropDown GetSelectedIndex
Sub GetSelectedItemIndex(ByVal control As IRibbonControl, ByRef Index)
'This procedure is used to ensure the first item in the dropdown is displayed.
Select Case control.ID
  Case Is = "DD1"
    Index = 0
  Case Else
End Select
End Sub
'Callback for Button onAction

'Callback for Toggle onAction
Sub ToggleonAction(control As IRibbonControl, pressed As Boolean)
Select Case control.ID
  Case Is = "TB1"
    ActiveWindow.View.ShowBookmarks = Not ActiveWindow.View.ShowBookmarks
  Case Is = "TB2"
  'Note:  "pressed" returns the toggle state.  So we could use this instead.
    If pressed Then
      ActiveWindow.View.ShowHiddenText = False
    Else
      ActiveWindow.View.ShowHiddenText = True
    End If
    If Not ActiveWindow.View.ShowHiddenText Then
      ActiveWindow.View.ShowAll = False
    End If
End Select
'Force the ribbon to redefine the control wiht correct image and label
MyRibbon.InvalidateControl (control.ID)
End Sub
'Callback for togglebutton getLabel
Sub getLabel(control As IRibbonControl, ByRef returnedVal)
Select Case control.ID
  Case Is = "TB1"
    If Not ActiveWindow.View.ShowBookmarks Then
      returnedVal = "Show Bookmarks"
    Else
      returnedVal = "Hide Bookmarks"
    End If
  Case Is = "TB2"
    If Not ActiveWindow.View.ShowHiddenText Then
      returnedVal = "Show Text"
    Else
      returnedVal = "Hide Text"
    End If
End Select
End Sub
'Callback for togglebutton getImage
Sub GetImage(control As IRibbonControl, ByRef returnedVal)
Select Case control.ID
  Case Is = "TB1"
   If ActiveWindow.View.ShowBookmarks Then
      returnedVal = "_3DTiltRightClassic"
    Else
      returnedVal = "_3DTiltLeftClassic"
   End If
  Case Is = "TB2"
    If ActiveWindow.View.ShowHiddenText Then
      returnedVal = "WebControlHidden"
    Else
      returnedVal = "SlideShowInAWindow"
    End If
End Select
End Sub
'Callback for togglebutton getPressed
Sub buttonPressed(control As IRibbonControl, ByRef toggleState)
'toggleState is used tp set the toggle state (i.e., true or false) and determine how the
'toggle appears on the ribbon (i.e., flusn or sunken).
Select Case control.ID
  Case Is = "TB1"
    If Not ActiveWindow.View.ShowBookmarks Then
      toggleState = True
    Else
      toggleState = False
    End If
  Case Is = "TB2"
    If Not ActiveWindow.View.ShowHiddenText Then
      toggleState = True
    Else
      toggleState = False
    End If
End Select
End Sub


' This function is called when the UI is invalidated. It will return the current status of the checkbox.
Public Sub IncludeSummary_getPressed(control As IRibbonControl, ByRef returnedVal)
  If control.ID = "IncludeSummary_Checkbox" Then
    returnedVal = SummaryCheckboxValue
  End If
End Sub

' This function is called when the checkbox is clicked. It toggles the value of the checkbox.
Public Sub IncludeSummary_onAction(control As IRibbonControl, pressed As Boolean)

    If control.ID = "IncludeSummary_Checkbox" Then
        SummaryCheckboxValue = pressed
        ' Invalidate the UI so that the IncludeSummary_getPressed function is called
    If Not MyRibbon Is Nothing Then
            MyRibbon.Invalidate
        End If
        ' Now you can use the value of SummaryCheckboxValue in your macro to determine its functionality
  End If
End Sub

-------------------------------------------------------------------------------
