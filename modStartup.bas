Attribute VB_Name = "modStartup"
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

' Global application settings (loaded on startup)
Public gstrPoolName As String
Public gstrAdminName As String
Public gstrStateOfPlay As String

'---------------------------------------------------------------------------------------
' Name       : InitializeApp
' Purpose    : Load global settings from tblSystemSettings. Called automatically
'              when frmMainDashboard opens (via its OnOpen event).
' Parameters : None
' Returns    : Variant (required for =Expression() event binding)
'---------------------------------------------------------------------------------------
Public Function InitializeApp() As Variant
    On Error GoTo ErrorHandler

    EnsureDefaultSettings

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT TOP 1 * FROM tblSystemSettings", dbOpenSnapshot)

    If Not rs.EOF Then
        gstrPoolName = Nz(rs!PoolName, "")
        gstrAdminName = Nz(rs!AdminName, "")
        gstrStateOfPlay = Nz(rs!StateOfPlay, "")
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    MsgBox "An error occurred in: InitializeApp" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : EnsureDefaultSettings
' Purpose    : Create a default settings row if tblSystemSettings is empty.
'              Uses TX as a sensible default state (no state income tax).
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub EnsureDefaultSettings()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset( _
        "SELECT COUNT(*) AS Cnt FROM tblSystemSettings", dbOpenSnapshot)

    If rs!Cnt = 0 Then
        db.Execute "INSERT INTO tblSystemSettings " & _
                   "(PoolName, AdminName, StateOfPlay) " & _
                   "VALUES ('My Powerball Pool', 'Admin', 'TX')", _
                   dbFailOnError
    End If

    rs.Close

Exit_Procedure:
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    ' Fail silently — user can enter settings manually via frmSettings
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : ConfigureStartup
' Purpose    : Set database startup properties so the app opens frmMainDashboard
'              automatically and looks like a finished application
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub ConfigureStartup()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb()

    ' Open frmMainDashboard on database open
    SetDatabaseProperty db, "StartUpForm", dbText, "frmMainDashboard"

    ' Set the title bar text
    SetDatabaseProperty db, "AppTitle", dbText, APP_TITLE

    ' Hide the navigation pane for a cleaner look
    SetDatabaseProperty db, "StartUpShowDBWindow", dbBoolean, False

    ' Hide the status bar
    SetDatabaseProperty db, "ShowStatusBar", dbBoolean, True

    ' Allow full menus (so users can access design tools if needed)
    SetDatabaseProperty db, "AllowFullMenus", dbBoolean, True

    MsgBox "Startup configured successfully." & vbCrLf & vbCrLf & _
           "On next database open:" & vbCrLf & _
           "  - frmMainDashboard opens automatically" & vbCrLf & _
           "  - Title bar shows '" & APP_TITLE & "'" & vbCrLf & _
           "  - Navigation pane is hidden" & vbCrLf & vbCrLf & _
           "Tip: Hold SHIFT while opening the database to " & _
           "bypass startup and show the navigation pane.", _
           vbInformation, APP_TITLE

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: ConfigureStartup" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : SetDatabaseProperty
' Purpose    : Set a database property, creating it if it does not exist
' Parameters : db (DAO.Database) - The database object
'              strName (String) - Property name
'              lngType (Long) - DAO data type constant
'              varValue (Variant) - Property value
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub SetDatabaseProperty(db As DAO.Database, _
                                 ByVal strName As String, _
                                 ByVal lngType As Long, _
                                 ByVal varValue As Variant)
    On Error Resume Next
    db.Properties(strName) = varValue
    If Err.Number = 3270 Then
        Err.Clear
        Dim prp As DAO.Property
        Set prp = db.CreateProperty(strName, lngType, varValue)
        db.Properties.Append prp
    End If
    On Error GoTo 0
End Sub
