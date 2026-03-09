Attribute VB_Name = "modCreateForms"
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

'---------------------------------------------------------------------------------------
' Name       : CreateAllForms
' Purpose    : Orchestrate creation of all MVP forms
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub CreateAllForms()
    On Error GoTo ErrorHandler

    CreateForm_frmMainDashboard
    CreateForm_frmSettings
    CreateForm_frmParticipants
    CreateForm_frmTicketEntry
    CreateForm_frmDrawResults
    CreateForm_frmMatchResults

    MsgBox "All forms created successfully.", vbInformation, APP_TITLE

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateAllForms" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  HELPER FUNCTIONS
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : FormExists
' Purpose    : Check if a form already exists in the current database
' Parameters : strFormName (String) - Name of the form to check
' Returns    : Boolean - True if the form exists
'---------------------------------------------------------------------------------------
Private Function FormExists(ByVal strFormName As String) As Boolean
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        If obj.Name = strFormName Then
            FormExists = True
            Exit Function
        End If
    Next obj
    FormExists = False
End Function

'---------------------------------------------------------------------------------------
' Name       : HideAttachedLabel
' Purpose    : Hide the auto-created label attached to a control
' Parameters : ctl (Control) - The control whose attached label to hide
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub HideAttachedLabel(ctl As Control)
    On Error Resume Next
    ctl.Controls(0).Visible = False
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Name       : StyleLabel
' Purpose    : Apply standard font styling to a label control
' Parameters : ctl (Control) - The label control
'              lngFontSize (Long) - Font size
'              blnBold (Boolean) - Whether to bold
'              lngForeColor (Long) - Text color
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub StyleLabel(ctl As Control, ByVal lngFontSize As Long, _
                       ByVal blnBold As Boolean, ByVal lngForeColor As Long)
    ctl.FontName = UI_FONT_NAME
    ctl.FontSize = lngFontSize
    ctl.FontBold = blnBold
    ctl.ForeColor = lngForeColor
End Sub

'---------------------------------------------------------------------------------------
' Name       : StyleTextBox
' Purpose    : Apply standard font styling to a text box control
' Parameters : ctl (Control) - The text box control
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub StyleTextBox(ctl As Control)
    ctl.FontName = UI_FONT_NAME
    ctl.FontSize = UI_FONT_SIZE_BODY
End Sub

'---------------------------------------------------------------------------------------
' Name       : StyleButton
' Purpose    : Apply standard styling to a command button
' Parameters : ctl (Control) - The button control
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub StyleButton(ctl As Control)
    ctl.FontName = UI_FONT_NAME
    ctl.FontSize = UI_FONT_SIZE_BODY
    ctl.FontBold = True
End Sub

' ======================================================================================
'  FORM: frmMainDashboard
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmMainDashboard
' Purpose    : Create the main navigation dashboard (unbound, no record source)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmMainDashboard()
    On Error GoTo ErrorHandler

    If FormExists("frmMainDashboard") Then
        Debug.Print "Form frmMainDashboard already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String
    Dim lngBtnLeft As Long
    Dim lngBtnWidth As Long
    Dim lngTop As Long

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = APP_TITLE
    frm.RecordSource = ""
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.Width = 7200
    frm.Section(acDetail).Height = 5400
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND
    frm.OnOpen = "=InitializeApp()"

    ' Button positioning: centered in 7200-wide form
    lngBtnWidth = 4320
    lngBtnLeft = (7200 - lngBtnWidth) / 2

    ' --- Title ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 1000, 300, 5200, 600)
    ctl.Name = "lblTitle"
    ctl.Caption = "Norris Powerball Pool"
    ctl.TextAlign = 2
    StyleLabel ctl, UI_FONT_SIZE_TITLE, True, UI_COLOR_PRIMARY

    ' --- Subtitle (dynamic pool name) ---
    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 1000, 950, 5200, 360)
    ctl.Name = "txtSubtitle"
    ctl.ControlSource = "=DLookup(""PoolName"",""tblSystemSettings"")"
    ctl.TextAlign = 2
    ctl.BackStyle = 0
    ctl.BorderStyle = 0
    ctl.Locked = True
    ctl.TabStop = False
    StyleTextBox ctl
    ctl.ForeColor = UI_COLOR_TEXT_LIGHT
    HideAttachedLabel ctl

    ' --- Navigation buttons ---
    lngTop = 1700

    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            lngBtnLeft, lngTop, lngBtnWidth, 480)
    ctl.Name = "cmdParticipants"
    ctl.Caption = "Manage Participants"
    ctl.OnClick = "=OpenParticipants()"
    StyleButton ctl

    lngTop = lngTop + 600
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            lngBtnLeft, lngTop, lngBtnWidth, 480)
    ctl.Name = "cmdTicketEntry"
    ctl.Caption = "Enter Tickets"
    ctl.OnClick = "=OpenTicketEntry()"
    StyleButton ctl

    lngTop = lngTop + 600
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            lngBtnLeft, lngTop, lngBtnWidth, 480)
    ctl.Name = "cmdDrawResults"
    ctl.Caption = "Enter Draw Results"
    ctl.OnClick = "=OpenDrawResults()"
    StyleButton ctl

    lngTop = lngTop + 600
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            lngBtnLeft, lngTop, lngBtnWidth, 480)
    ctl.Name = "cmdMatchResults"
    ctl.Caption = "Check Matches"
    ctl.OnClick = "=OpenMatchResults()"
    StyleButton ctl

    lngTop = lngTop + 600
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            lngBtnLeft, lngTop, lngBtnWidth, 480)
    ctl.Name = "cmdSettings"
    ctl.Caption = "Settings"
    ctl.OnClick = "=OpenSettings()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmMainDashboard", acForm, strName

    Debug.Print "Form frmMainDashboard created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmMainDashboard" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  FORM: frmSettings
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmSettings
' Purpose    : Create the settings form (bound to tblSystemSettings, single record)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmSettings()
    On Error GoTo ErrorHandler

    If FormExists("frmSettings") Then
        Debug.Print "Form frmSettings already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = "Pool Settings"
    frm.RecordSource = "tblSystemSettings"
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.AllowAdditions = False
    frm.AllowDeletions = False
    frm.ScrollBars = 0
    frm.Width = 7200
    frm.Section(acDetail).Height = 3600
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND

    ' --- Title label ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 200, 6600, 500)
    ctl.Name = "lblTitle"
    ctl.Caption = "Pool Settings"
    StyleLabel ctl, UI_FONT_SIZE_HEADING, True, UI_COLOR_PRIMARY

    ' --- Pool Name ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 950, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblPoolName"
    ctl.Caption = "Pool Name:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 950, 4200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtPoolName"
    ctl.ControlSource = "PoolName"
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- Admin Name ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1450, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblAdminName"
    ctl.Caption = "Admin Name:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 1450, 4200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtAdminName"
    ctl.ControlSource = "AdminName"
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- State of Play (combo box) ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1950, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblStateOfPlay"
    ctl.Caption = "State of Play:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acComboBox, acDetail, , , 2400, 1950, 4200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cboStateOfPlay"
    ctl.ControlSource = "StateOfPlay"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT StateCode, StateName FROM tlkpStates ORDER BY StateName"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "600;3500"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- Save & Close button ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            2400, 2700, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdSaveClose"
    ctl.Caption = "Save && Close"
    ctl.OnClick = "=SaveSettingsAndCloseForm()"
    StyleButton ctl

    ' --- Close (cancel) button ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            4950, 2700, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=CloseCurrentForm()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmSettings", acForm, strName

    Debug.Print "Form frmSettings created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmSettings" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  FORM: frmParticipants
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmParticipants
' Purpose    : Create the participants continuous form with header labels
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmParticipants()
    On Error GoTo ErrorHandler

    If FormExists("frmParticipants") Then
        Debug.Print "Form frmParticipants already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = "Pool Participants"
    frm.RecordSource = "tblParticipants"
    frm.DefaultView = 1    ' Continuous Forms
    frm.NavigationButtons = True
    frm.RecordSelectors = True
    frm.DividingLines = True
    frm.Width = 9600

    ' Enable form header/footer
    On Error Resume Next
    frm.Section(acHeader).Height = 500
    frm.Section(acFooter).Height = 600
    If Err.Number <> 0 Then
        Err.Clear
        DoCmd.RunCommand acCmdFormHdrFtr
        frm.Section(acHeader).Height = 500
        frm.Section(acFooter).Height = 600
    End If
    On Error GoTo ErrorHandler

    frm.Section(acDetail).Height = 400
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND
    frm.Section(acHeader).BackColor = UI_COLOR_HEADER_BG
    frm.Section(acFooter).BackColor = UI_COLOR_HEADER_BG

    ' --- Header labels ---
    Set ctl = CreateControl(strName, acLabel, acHeader, , , 100, 100, 1800, 300)
    ctl.Name = "lblFirstName"
    ctl.Caption = "First Name"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_PRIMARY

    Set ctl = CreateControl(strName, acLabel, acHeader, , , 2000, 100, 1800, 300)
    ctl.Name = "lblLastName"
    ctl.Caption = "Last Name"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_PRIMARY

    Set ctl = CreateControl(strName, acLabel, acHeader, , , 3900, 100, 2400, 300)
    ctl.Name = "lblEmail"
    ctl.Caption = "Email"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_PRIMARY

    Set ctl = CreateControl(strName, acLabel, acHeader, , , 6400, 100, 1800, 300)
    ctl.Name = "lblPhone"
    ctl.Caption = "Phone"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_PRIMARY

    Set ctl = CreateControl(strName, acLabel, acHeader, , , 8300, 100, 900, 300)
    ctl.Name = "lblActive"
    ctl.Caption = "Active"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_PRIMARY

    ' --- Detail controls (bound) ---
    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 100, 30, 1800, 330)
    ctl.Name = "txtFirstName"
    ctl.ControlSource = "FirstName"
    StyleTextBox ctl
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2000, 30, 1800, 330)
    ctl.Name = "txtLastName"
    ctl.ControlSource = "LastName"
    StyleTextBox ctl
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 3900, 30, 2400, 330)
    ctl.Name = "txtEmail"
    ctl.ControlSource = "Email"
    StyleTextBox ctl
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 6400, 30, 1800, 330)
    ctl.Name = "txtPhone"
    ctl.ControlSource = "Phone"
    StyleTextBox ctl
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acCheckBox, acDetail, , , 8500, 70, 300, 300)
    ctl.Name = "chkIsActive"
    ctl.ControlSource = "IsActive"
    HideAttachedLabel ctl

    ' --- Footer: Close button ---
    Set ctl = CreateControl(strName, acCommandButton, acFooter, , , _
                            3600, 100, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=CloseCurrentForm()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmParticipants", acForm, strName

    Debug.Print "Form frmParticipants created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmParticipants" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  FORM: frmTicketEntry
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmTicketEntry
' Purpose    : Create the ticket entry form (bound to tblTickets)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmTicketEntry()
    On Error GoTo ErrorHandler

    If FormExists("frmTicketEntry") Then
        Debug.Print "Form frmTicketEntry already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String
    Dim i As Integer
    Dim lngBallLeft As Long

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = "Enter Ticket"
    frm.RecordSource = "tblTickets"
    frm.NavigationButtons = True
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.Width = 7200
    frm.Section(acDetail).Height = 4850
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND
    frm.OnOpen = "=TogglePlayOptions()"
    frm.OnCurrent = "=TogglePlayOptions()"
    frm.BeforeUpdate = "=ValidateTicketBeforeUpdate()"

    ' --- Title ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 200, 6600, 500)
    ctl.Name = "lblTitle"
    ctl.Caption = "Enter Ticket"
    StyleLabel ctl, UI_FONT_SIZE_HEADING, True, UI_COLOR_PRIMARY

    ' --- Drawing combo ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 900, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblDrawing"
    ctl.Caption = "Drawing:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acComboBox, acDetail, , , 2400, 900, 4200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cboDrawingID"
    ctl.ControlSource = "DrawingID"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT DrawingID, Format(DrawDate,'mm/dd/yyyy (ddd)') AS DisplayDate FROM tblDrawings ORDER BY DrawDate DESC"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;4100"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- Purchased By combo ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1350, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblPurchasedBy"
    ctl.Caption = "Purchased By:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acComboBox, acDetail, , , 2400, 1350, 4200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cboPurchasedBy"
    ctl.ControlSource = "ParticipantID"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT ParticipantID, FirstName & ' ' & LastName AS FullName FROM tblParticipants WHERE IsActive = True ORDER BY LastName, FirstName"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;4100"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- White Balls label ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1950, 6600, 300)
    ctl.Name = "lblWhiteBalls"
    ctl.Caption = "White Balls (1-69):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- WB1 through WB5 ---
    lngBallLeft = 300
    For i = 1 To 5
        Set ctl = CreateControl(strName, acTextBox, acDetail, , , _
                                lngBallLeft, 2350, 1000, UI_TEXTBOX_HEIGHT)
        ctl.Name = "txtWB" & i
        ctl.ControlSource = "WB" & i
        StyleTextBox ctl
        ctl.TextAlign = 2  ' Center
        HideAttachedLabel ctl
        lngBallLeft = lngBallLeft + 1100
    Next i

    ' --- Powerball ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 2950, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblPB"
    ctl.Caption = "Powerball (1-26):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 2950, 1000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtPB"
    ctl.ControlSource = "PB"
    StyleTextBox ctl
    ctl.TextAlign = 2
    HideAttachedLabel ctl

    ' --- Power Play checkbox ---
    Set ctl = CreateControl(strName, acCheckBox, acDetail, , , 300, 3550, 300, 300)
    ctl.Name = "chkPowerPlay"
    ctl.ControlSource = "IsPowerPlay"
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acLabel, acDetail, , , 700, 3550, 1800, 300)
    ctl.Name = "lblPowerPlay"
    ctl.Caption = "Power Play"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- Double Play checkbox ---
    Set ctl = CreateControl(strName, acCheckBox, acDetail, , , 2800, 3550, 300, 300)
    ctl.Name = "chkDoublePlay"
    ctl.ControlSource = "IsDoublePlay"
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acLabel, acDetail, , , 3200, 3550, 1800, 300)
    ctl.Name = "lblDoublePlay"
    ctl.Caption = "Double Play"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- Buttons ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            300, 4200, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdNew"
    ctl.Caption = "New Ticket"
    ctl.OnClick = "=GoToNewRecord()"
    StyleButton ctl

    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            2850, 4200, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdDelete"
    ctl.Caption = "Delete Ticket"
    ctl.OnClick = "=DeleteCurrentRecord()"
    StyleButton ctl

    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            5400, 4200, 1500, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=CloseCurrentForm()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmTicketEntry", acForm, strName

    Debug.Print "Form frmTicketEntry created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmTicketEntry" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  FORM: frmDrawResults
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmDrawResults
' Purpose    : Create the draw results entry form (bound to tblDrawings)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmDrawResults()
    On Error GoTo ErrorHandler

    If FormExists("frmDrawResults") Then
        Debug.Print "Form frmDrawResults already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String
    Dim i As Integer
    Dim lngBallLeft As Long

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = "Draw Results"
    frm.RecordSource = "tblDrawings"
    frm.NavigationButtons = True
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.Width = 7200
    frm.Section(acDetail).Height = 7500
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND

    ' --- Title ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 200, 6600, 500)
    ctl.Name = "lblTitle"
    ctl.Caption = "Draw Results"
    StyleLabel ctl, UI_FONT_SIZE_HEADING, True, UI_COLOR_PRIMARY

    ' --- Draw Date ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 900, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblDrawDate"
    ctl.Caption = "Draw Date:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 900, 2400, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtDrawDate"
    ctl.ControlSource = "DrawDate"
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- White Balls label ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1500, 6600, 300)
    ctl.Name = "lblWhiteBalls"
    ctl.Caption = "Winning White Balls (1-69):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- WB1 through WB5 ---
    lngBallLeft = 300
    For i = 1 To 5
        Set ctl = CreateControl(strName, acTextBox, acDetail, , , _
                                lngBallLeft, 1900, 1000, UI_TEXTBOX_HEIGHT)
        ctl.Name = "txtWB" & i
        ctl.ControlSource = "WB" & i
        StyleTextBox ctl
        ctl.TextAlign = 2
        HideAttachedLabel ctl
        lngBallLeft = lngBallLeft + 1100
    Next i

    ' --- Powerball ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 2500, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblPB"
    ctl.Caption = "Powerball (1-26):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 2500, 1000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtPB"
    ctl.ControlSource = "PB"
    StyleTextBox ctl
    ctl.TextAlign = 2
    HideAttachedLabel ctl

    ' --- Power Play Multiplier ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 3100, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblPPMultiplier"
    ctl.Caption = "PP Multiplier:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acComboBox, acDetail, , , 2400, 3100, 1200, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cboPowerPlayMultiplier"
    ctl.ControlSource = "PowerPlayMultiplier"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = ";2;3;4;5;10"
    ctl.LimitToList = True
    StyleTextBox ctl
    ctl.TextAlign = 2
    HideAttachedLabel ctl

    ' --- Jackpot Amount ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 4000, 3100, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblJackpot"
    ctl.Caption = "Jackpot Amount:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 4000, 3500, 2400, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtJackpotAmount"
    ctl.ControlSource = "JackpotAmount"
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- Double Play Section ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 4100, 6600, 400)
    ctl.Name = "lblDPSection"
    ctl.Caption = "Double Play Drawing"
    StyleLabel ctl, UI_FONT_SIZE_HEADING, True, UI_COLOR_ACCENT

    ' --- DP White Balls label ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 4600, 6600, 300)
    ctl.Name = "lblDPWhiteBalls"
    ctl.Caption = "DP White Balls (1-69):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- DPWB1 through DPWB5 ---
    lngBallLeft = 300
    For i = 1 To 5
        Set ctl = CreateControl(strName, acTextBox, acDetail, , , _
                                lngBallLeft, 5000, 1000, UI_TEXTBOX_HEIGHT)
        ctl.Name = "txtDPWB" & i
        ctl.ControlSource = "DPWB" & i
        StyleTextBox ctl
        ctl.TextAlign = 2
        HideAttachedLabel ctl
        lngBallLeft = lngBallLeft + 1100
    Next i

    ' --- DP Powerball ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 5600, 2000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblDPPB"
    ctl.Caption = "DP Powerball (1-26):"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acTextBox, acDetail, , , 2400, 5600, 1000, UI_TEXTBOX_HEIGHT)
    ctl.Name = "txtDPPB"
    ctl.ControlSource = "DPPB"
    StyleTextBox ctl
    ctl.TextAlign = 2
    HideAttachedLabel ctl

    ' --- Verified checkbox ---
    Set ctl = CreateControl(strName, acCheckBox, acDetail, , , 300, 6200, 300, 300)
    ctl.Name = "chkVerified"
    ctl.ControlSource = "IsVerified"
    HideAttachedLabel ctl

    Set ctl = CreateControl(strName, acLabel, acDetail, , , 700, 6200, 1800, 300)
    ctl.Name = "lblVerified"
    ctl.Caption = "Verified"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    ' --- Buttons ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            300, 6800, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdNew"
    ctl.Caption = "New Drawing"
    ctl.OnClick = "=GoToNewRecord()"
    StyleButton ctl

    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            2850, 6800, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdDelete"
    ctl.Caption = "Delete Drawing"
    ctl.OnClick = "=DeleteCurrentRecord()"
    StyleButton ctl

    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            5400, 6800, 1500, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=CloseCurrentForm()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmDrawResults", acForm, strName

    Debug.Print "Form frmDrawResults created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmDrawResults" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  FORM: frmMatchResults
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateForm_frmMatchResults
' Purpose    : Create the match results form (unbound, with drawing combo and
'              results listbox)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateForm_frmMatchResults()
    On Error GoTo ErrorHandler

    If FormExists("frmMatchResults") Then
        Debug.Print "Form frmMatchResults already exists - skipped."
        Exit Sub
    End If

    Dim frm As Form
    Dim ctl As Control
    Dim strName As String
    Dim strRowSource As String

    Set frm = CreateForm
    strName = frm.Name

    ' --- Form properties ---
    frm.Caption = "Match Results"
    frm.RecordSource = ""
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    frm.ScrollBars = 0
    frm.Width = 10200
    frm.Section(acDetail).Height = 8600
    frm.Section(acDetail).BackColor = UI_COLOR_BACKGROUND

    ' --- Title ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 200, 9600, 500)
    ctl.Name = "lblTitle"
    ctl.Caption = "Match Results"
    StyleLabel ctl, UI_FONT_SIZE_HEADING, True, UI_COLOR_PRIMARY

    ' --- Drawing combo ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 900, 1800, UI_TEXTBOX_HEIGHT)
    ctl.Name = "lblDrawing"
    ctl.Caption = "Drawing:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, False, UI_COLOR_TEXT

    Set ctl = CreateControl(strName, acComboBox, acDetail, , , 2200, 900, 3600, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cboDrawing"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT DrawingID, Format(DrawDate,'mm/dd/yyyy (ddd)') AS DisplayDate FROM tblDrawings ORDER BY DrawDate DESC"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;3500"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    ctl.AfterUpdate = "=RefreshMatchResults()"
    StyleTextBox ctl
    HideAttachedLabel ctl

    ' --- Check Matches button ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            6100, 900, 2400, UI_TEXTBOX_HEIGHT)
    ctl.Name = "cmdCheckMatches"
    ctl.Caption = "Check Matches"
    ctl.OnClick = "=RefreshMatchResults()"
    StyleButton ctl

    ' --- Results label ---
    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 1500, 9600, 300)
    ctl.Name = "lblResults"
    ctl.Caption = "Powerball Winning Tickets:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_TEXT

    ' --- Results listbox (with Power Play adjusted amount) ---
    strRowSource = "SELECT wt.TicketID, " & _
                   "wt.PurchasedBy AS [Purchased By], " & _
                   "wt.WB1 & '-' & wt.WB2 & '-' & wt.WB3 & '-' & wt.WB4 & '-' & wt.WB5 AS [White Balls], " & _
                   "wt.PB AS [PB], " & _
                   "wt.WhiteBallMatches AS [WB Matches], " & _
                   "IIf(wt.PowerballMatch,'Yes','No') AS [PB Match], " & _
                   "wt.PrizeName AS [Prize Tier], " & _
                   "IIf(wt.IsPowerPlay,'Yes','No') AS [PP], " & _
                   "Format(wt.AdjustedPrizeAmount,'Currency') AS [Prize] " & _
                   "FROM qryWinningTickets AS wt " & _
                   "WHERE wt.DrawingID = Nz(Forms!frmMatchResults!cboDrawing,0) " & _
                   "ORDER BY wt.AdjustedPrizeAmount DESC"

    Set ctl = CreateControl(strName, acListBox, acDetail, , , 300, 1900, 9600, 2800)
    ctl.Name = "lstResults"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = strRowSource
    ctl.ColumnCount = 9
    ctl.ColumnWidths = "600;1600;2200;500;800;700;1400;500;1300"
    ctl.ColumnHeads = True
    ctl.FontName = UI_FONT_NAME
    ctl.FontSize = UI_FONT_SIZE_BODY
    HideAttachedLabel ctl

    ' --- Double Play Results label ---
    Dim strDPRowSource As String

    Set ctl = CreateControl(strName, acLabel, acDetail, , , 300, 4900, 9600, 300)
    ctl.Name = "lblDPResults"
    ctl.Caption = "Double Play Winning Tickets:"
    StyleLabel ctl, UI_FONT_SIZE_BODY, True, UI_COLOR_TEXT

    ' --- Double Play Results listbox ---
    strDPRowSource = "SELECT dp.TicketID, " & _
                     "dp.PurchasedBy AS [Purchased By], " & _
                     "dp.WB1 & '-' & dp.WB2 & '-' & dp.WB3 & '-' & dp.WB4 & '-' & dp.WB5 AS [White Balls], " & _
                     "dp.PB AS [PB], " & _
                     "dp.WhiteBallMatches AS [WB Matches], " & _
                     "IIf(dp.PowerballMatch,'Yes','No') AS [PB Match], " & _
                     "dp.PrizeName AS [Prize Tier], " & _
                     "Format(dp.DefaultPrizeAmount,'Currency') AS [Prize] " & _
                     "FROM qryDoublePlayWinningTickets AS dp " & _
                     "WHERE dp.DrawingID = Nz(Forms!frmMatchResults!cboDrawing,0) " & _
                     "ORDER BY dp.DefaultPrizeAmount DESC"

    Set ctl = CreateControl(strName, acListBox, acDetail, , , 300, 5300, 9600, 2400)
    ctl.Name = "lstDPResults"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = strDPRowSource
    ctl.ColumnCount = 8
    ctl.ColumnWidths = "700;1800;2200;600;900;800;1400;1200"
    ctl.ColumnHeads = True
    ctl.FontName = UI_FONT_NAME
    ctl.FontSize = UI_FONT_SIZE_BODY
    HideAttachedLabel ctl

    ' --- Close button ---
    Set ctl = CreateControl(strName, acCommandButton, acDetail, , , _
                            3900, 7900, UI_BUTTON_WIDTH, UI_BUTTON_HEIGHT)
    ctl.Name = "cmdClose"
    ctl.Caption = "Close"
    ctl.OnClick = "=CloseCurrentForm()"
    StyleButton ctl

    ' --- Save and rename ---
    DoCmd.Close acForm, strName, acSaveYes
    DoCmd.Rename "frmMatchResults", acForm, strName

    Debug.Print "Form frmMatchResults created successfully."

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateForm_frmMatchResults" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    On Error Resume Next
    DoCmd.Close acForm, strName, acSaveNo
    On Error GoTo 0
    Resume Exit_Procedure
End Sub
