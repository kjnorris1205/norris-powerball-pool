Attribute VB_Name = "modFormEvents"
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

' ======================================================================================
'  NAVIGATION FUNCTIONS (called from frmMainDashboard button OnClick events)
'  Must be Public Functions returning Variant for =Expression() event binding.
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : OpenParticipants
' Purpose    : Open the Participants form from the dashboard
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function OpenParticipants() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmParticipants"
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: OpenParticipants" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : OpenTicketEntry
' Purpose    : Open the Ticket Entry form from the dashboard
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function OpenTicketEntry() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmTicketEntry"
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: OpenTicketEntry" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : OpenDrawResults
' Purpose    : Open the Draw Results form from the dashboard
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function OpenDrawResults() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmDrawResults"
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: OpenDrawResults" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : OpenMatchResults
' Purpose    : Open the Match Results form from the dashboard
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function OpenMatchResults() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmMatchResults"
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: OpenMatchResults" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : OpenSettings
' Purpose    : Open the Settings form from the dashboard
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function OpenSettings() As Variant
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmSettings"
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: OpenSettings" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

' ======================================================================================
'  COMMON FORM FUNCTIONS
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CloseCurrentForm
' Purpose    : Close whichever form is currently active
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function CloseCurrentForm() As Variant
    On Error GoTo ErrorHandler
    DoCmd.Close acForm, Screen.ActiveForm.Name
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: CloseCurrentForm" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : SaveSettingsAndCloseForm
' Purpose    : Save settings, update globals and dashboard, then close form
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function SaveSettingsAndCloseForm() As Variant
    On Error GoTo ErrorHandler

    DoCmd.RunCommand acCmdSaveRecord

    ' Reload global settings from the saved record
    gstrPoolName = Nz(Forms!frmSettings!txtPoolName, "")
    gstrAdminName = Nz(Forms!frmSettings!txtAdminName, "")
    gstrStateOfPlay = Nz(Forms!frmSettings!cboStateOfPlay, "")

    ' Refresh the dashboard subtitle if it is open
    If CurrentProject.AllForms("frmMainDashboard").IsLoaded Then
        Forms!frmMainDashboard!txtSubtitle.Requery
    End If

    DoCmd.Close acForm, "frmSettings"

Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: SaveSettingsAndCloseForm" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : GoToNewRecord
' Purpose    : Navigate to a new blank record in the current form
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function GoToNewRecord() As Variant
    On Error GoTo ErrorHandler
    DoCmd.GoToRecord , , acNewRec
Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: GoToNewRecord" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : DeleteCurrentRecord
' Purpose    : Delete the current record after user confirmation
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function DeleteCurrentRecord() As Variant
    On Error GoTo ErrorHandler

    Dim lngResponse As Long
    lngResponse = MsgBox("Are you sure you want to delete this record?", _
                         vbYesNo + vbQuestion, APP_TITLE)
    If lngResponse = vbYes Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If

Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: DeleteCurrentRecord" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

' ======================================================================================
'  MATCH RESULTS FUNCTIONS
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : RefreshMatchResults
' Purpose    : Requery the match results listbox on frmMatchResults.
'              Called from cboDrawing AfterUpdate and cmdCheckMatches Click.
' Parameters : None
' Returns    : Variant
'---------------------------------------------------------------------------------------
Public Function RefreshMatchResults() As Variant
    On Error GoTo ErrorHandler

    Dim frm As Form
    Set frm = Screen.ActiveForm

    If IsNull(frm!cboDrawing.Value) Then
        frm!lstResults.RowSource = ""
        Exit Function
    End If

    frm!lstResults.Requery

Exit_Function:
    Exit Function
ErrorHandler:
    MsgBox "An error occurred in: RefreshMatchResults" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function
