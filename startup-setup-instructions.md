# Startup Setup Instructions

Set up the application startup module that loads global settings and configures the database to open `frmMainDashboard` automatically.

## Setup Steps

1. Open `NorrisPowerballPool.accdb` in Microsoft Access.
2. Press **Alt+F11** to open the VBA editor.
3. Go to **Insert** → **Module**. A new module window opens.
4. Copy the entire VBA code block below and paste it into the new module.
5. In the **Properties** window (press **F4** if not visible), change the module `(Name)` to `modStartup`.
6. Press **Ctrl+S** to save.
7. Press **Ctrl+G** to open the **Immediate Window**.
8. Type `ConfigureStartup` and press **Enter**.
9. A confirmation message lists the startup settings applied.

## VBA Code: `modStartup`

```vb
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
```

## Notes

- **InitializeApp** is called automatically via the `frmMainDashboard` OnOpen event (set up by `modFormSetup`). It loads settings into public variables (`gstrPoolName`, `gstrAdminName`, `gstrStateOfPlay`) for use throughout the app.
- **EnsureDefaultSettings** inserts a placeholder row into `tblSystemSettings` if empty, so the settings form always has a record to edit.
- **ConfigureStartup** only needs to be run once. It sets the startup form, app title, and hides the navigation pane.
- **To access design tools after startup is configured,** hold **Shift** while opening the database to bypass the startup form and show the navigation pane.
