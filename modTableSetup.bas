Attribute VB_Name = "modTableSetup"
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

'---------------------------------------------------------------------------------------
' Name       : CreateAllTables
' Purpose    : Orchestrate creation of all tables and relationships
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub CreateAllTables()
    On Error GoTo ErrorHandler

    ' Create tables in dependency order (parents before children)
    CreateTable_tlkpStates
    CreateTable_tlkpPrizeTiers
    CreateTable_tlkpAppVersion
    CreateTable_tblParticipants
    CreateTable_tblDrawings
    CreateTable_tblSystemSettings
    CreateTable_tblTickets
    CreateTable_tblContributions

    ' Create relationships after all tables exist
    CreateAllRelationships

    MsgBox "All tables and relationships created successfully.", _
           vbInformation, APP_TITLE

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateAllTables" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : TableExists
' Purpose    : Check if a table already exists in the current database
' Parameters : strTableName (String) - Name of the table to check
' Returns    : Boolean - True if the table exists
'---------------------------------------------------------------------------------------
Private Function TableExists(ByVal strTableName As String) As Boolean
    On Error Resume Next
    Dim td As DAO.TableDef
    Set td = CurrentDb.TableDefs(strTableName)
    TableExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Name       : RelationExists
' Purpose    : Check if a relationship already exists in the current database
' Parameters : strRelName (String) - Name of the relationship to check
' Returns    : Boolean - True if the relationship exists
'---------------------------------------------------------------------------------------
Private Function RelationExists(ByVal strRelName As String) As Boolean
    On Error Resume Next
    Dim rel As DAO.Relation
    Set rel = CurrentDb.Relations(strRelName)
    RelationExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Name       : SetFieldProperty
' Purpose    : Set a custom property on a DAO field (creates property if it
'              does not already exist, updates it if it does)
' Parameters : fld (DAO.Field) - The field to set the property on
'              strName (String) - Property name (e.g., "Caption", "Format")
'              lngType (Long) - DAO data-type constant for the property value
'              varValue (Variant) - Property value
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub SetFieldProperty(fld As DAO.Field, ByVal strName As String, _
                             ByVal lngType As Long, ByVal varValue As Variant)
    On Error Resume Next
    fld.Properties(strName) = varValue
    If Err.Number <> 0 Then
        Err.Clear
        Dim prp As DAO.Property
        Set prp = fld.CreateProperty(strName, lngType, varValue)
        fld.Properties.Append prp
    End If
    On Error GoTo 0
End Sub

' ======================================================================================
'  LOOKUP / REFERENCE TABLES
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tlkpStates
' Purpose    : Create the tlkpStates lookup table (50 US states + DC)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tlkpStates()
    On Error GoTo ErrorHandler

    If TableExists("tlkpStates") Then
        Debug.Print "Table tlkpStates already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tlkpStates")

    ' --- StateCode (PK, Short Text, 2) ---
    Set fld = td.CreateField("StateCode", dbText, 2)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- StateName (Short Text, 50) ---
    Set fld = td.CreateField("StateName", dbText, 50)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- FederalTaxRate (Double) ---
    Set fld = td.CreateField("FederalTaxRate", dbDouble)
    fld.Required = True
    fld.DefaultValue = "0.24"
    fld.ValidationRule = ">=0 And <=1"
    fld.ValidationText = "Federal tax rate must be between 0% and 100%."
    td.Fields.Append fld

    ' --- StateTaxRate (Double) ---
    Set fld = td.CreateField("StateTaxRate", dbDouble)
    fld.Required = True
    fld.DefaultValue = "0"
    fld.ValidationRule = ">=0 And <=1"
    fld.ValidationText = "State tax rate must be between 0% and 100%."
    td.Fields.Append fld

    ' --- HasStateLottery (Yes/No) ---
    Set fld = td.CreateField("HasStateLottery", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- HasPowerPlay (Yes/No) ---
    Set fld = td.CreateField("HasPowerPlay", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- HasDoublePlay (Yes/No) ---
    Set fld = td.CreateField("HasDoublePlay", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("StateCode")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties (must be done after table is appended) ---
    Set td = db.TableDefs("tlkpStates")

    Set fld = td.Fields("StateCode")
    SetFieldProperty fld, "Description", dbText, "Two-letter state/territory code"
    SetFieldProperty fld, "Caption", dbText, "State Code"
    SetFieldProperty fld, "InputMask", dbText, ">LL"

    Set fld = td.Fields("StateName")
    SetFieldProperty fld, "Description", dbText, "Full state or territory name"
    SetFieldProperty fld, "Caption", dbText, "State Name"

    Set fld = td.Fields("FederalTaxRate")
    SetFieldProperty fld, "Description", dbText, "Federal tax withholding rate as a decimal"
    SetFieldProperty fld, "Caption", dbText, "Federal Tax Rate"
    SetFieldProperty fld, "Format", dbText, "Percent"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 2

    Set fld = td.Fields("StateTaxRate")
    SetFieldProperty fld, "Description", dbText, "State tax withholding rate as a decimal"
    SetFieldProperty fld, "Caption", dbText, "State Tax Rate"
    SetFieldProperty fld, "Format", dbText, "Percent"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 4

    Set fld = td.Fields("HasStateLottery")
    SetFieldProperty fld, "Description", dbText, "Whether the state participates in Powerball"
    SetFieldProperty fld, "Caption", dbText, "Has State Lottery"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Set fld = td.Fields("HasPowerPlay")
    SetFieldProperty fld, "Description", dbText, "Whether Power Play is available in this state"
    SetFieldProperty fld, "Caption", dbText, "Has Power Play"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Set fld = td.Fields("HasDoublePlay")
    SetFieldProperty fld, "Description", dbText, "Whether Double Play is available in this state"
    SetFieldProperty fld, "Caption", dbText, "Has Double Play"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Debug.Print "Table tlkpStates created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tlkpStates" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tlkpPrizeTiers
' Purpose    : Create the tlkpPrizeTiers lookup table (9 Powerball prize tiers)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tlkpPrizeTiers()
    On Error GoTo ErrorHandler

    If TableExists("tlkpPrizeTiers") Then
        Debug.Print "Table tlkpPrizeTiers already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tlkpPrizeTiers")

    ' --- PrizeTierID (AutoNumber PK) ---
    Set fld = td.CreateField("PrizeTierID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- WhiteBallMatches (Integer) ---
    Set fld = td.CreateField("WhiteBallMatches", dbInteger)
    fld.Required = True
    fld.ValidationRule = ">=0 And <=5"
    fld.ValidationText = "White ball matches must be between 0 and 5."
    td.Fields.Append fld

    ' --- PowerballMatch (Yes/No) ---
    Set fld = td.CreateField("PowerballMatch", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- PrizeName (Short Text, 50) ---
    Set fld = td.CreateField("PrizeName", dbText, 50)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- DefaultPrizeAmount (Currency) ---
    Set fld = td.CreateField("DefaultPrizeAmount", dbCurrency)
    fld.Required = True
    fld.DefaultValue = "0"
    fld.ValidationRule = ">=0"
    fld.ValidationText = "Default prize amount cannot be negative."
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("PrizeTierID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tlkpPrizeTiers")

    Set fld = td.Fields("PrizeTierID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated tier identifier"
    SetFieldProperty fld, "Caption", dbText, "Prize Tier ID"

    Set fld = td.Fields("WhiteBallMatches")
    SetFieldProperty fld, "Description", dbText, "Number of white balls matched (0-5)"
    SetFieldProperty fld, "Caption", dbText, "White Ball Matches"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    Set fld = td.Fields("PowerballMatch")
    SetFieldProperty fld, "Description", dbText, "Whether the Powerball was also matched"
    SetFieldProperty fld, "Caption", dbText, "Powerball Match"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Set fld = td.Fields("PrizeName")
    SetFieldProperty fld, "Description", dbText, "Display name (e.g., ""Jackpot"", ""Match 4+PB"")"
    SetFieldProperty fld, "Caption", dbText, "Prize Name"

    Set fld = td.Fields("DefaultPrizeAmount")
    SetFieldProperty fld, "Description", dbText, "Default fixed prize amount ($0 for jackpot)"
    SetFieldProperty fld, "Caption", dbText, "Default Prize Amount"
    SetFieldProperty fld, "Format", dbText, "Currency"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 2

    Debug.Print "Table tlkpPrizeTiers created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tlkpPrizeTiers" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tlkpAppVersion
' Purpose    : Create the tlkpAppVersion lookup table
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tlkpAppVersion()
    On Error GoTo ErrorHandler

    If TableExists("tlkpAppVersion") Then
        Debug.Print "Table tlkpAppVersion already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tlkpAppVersion")

    ' --- VersionID (AutoNumber PK) ---
    Set fld = td.CreateField("VersionID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- VersionNumber (Short Text, 20) ---
    Set fld = td.CreateField("VersionNumber", dbText, 20)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- ReleaseDate (Date/Time) ---
    Set fld = td.CreateField("ReleaseDate", dbDate)
    fld.Required = True
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("VersionID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tlkpAppVersion")

    Set fld = td.Fields("VersionID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated version identifier"
    SetFieldProperty fld, "Caption", dbText, "Version ID"

    Set fld = td.Fields("VersionNumber")
    SetFieldProperty fld, "Description", dbText, "Semantic version string (e.g., ""1.0.0"")"
    SetFieldProperty fld, "Caption", dbText, "Version Number"

    Set fld = td.Fields("ReleaseDate")
    SetFieldProperty fld, "Description", dbText, "Date this version was released"
    SetFieldProperty fld, "Caption", dbText, "Release Date"
    SetFieldProperty fld, "Format", dbText, "Short Date"
    SetFieldProperty fld, "InputMask", dbText, "99/99/0000;0;_"

    Debug.Print "Table tlkpAppVersion created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tlkpAppVersion" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  SYSTEM SETTINGS TABLE
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tblSystemSettings
' Purpose    : Create the tblSystemSettings single-row config table
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tblSystemSettings()
    On Error GoTo ErrorHandler

    If TableExists("tblSystemSettings") Then
        Debug.Print "Table tblSystemSettings already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tblSystemSettings")

    ' --- SettingsID (AutoNumber PK) ---
    Set fld = td.CreateField("SettingsID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- PoolName (Short Text, 100) ---
    Set fld = td.CreateField("PoolName", dbText, 100)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- AdminName (Short Text, 100) ---
    Set fld = td.CreateField("AdminName", dbText, 100)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- StateOfPlay (Short Text, 2 — FK to tlkpStates.StateCode) ---
    Set fld = td.CreateField("StateOfPlay", dbText, 2)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("SettingsID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' --- Index on StateOfPlay (Duplicates OK) ---
    Set idx = td.CreateIndex("StateOfPlay")
    idx.Unique = False
    Set fld = idx.CreateField("StateOfPlay")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tblSystemSettings")

    Set fld = td.Fields("SettingsID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated settings identifier"
    SetFieldProperty fld, "Caption", dbText, "Settings ID"

    Set fld = td.Fields("PoolName")
    SetFieldProperty fld, "Description", dbText, "Name of the lottery pool"
    SetFieldProperty fld, "Caption", dbText, "Pool Name"

    Set fld = td.Fields("AdminName")
    SetFieldProperty fld, "Description", dbText, "Pool administrator's name"
    SetFieldProperty fld, "Caption", dbText, "Admin Name"

    Set fld = td.Fields("StateOfPlay")
    SetFieldProperty fld, "Description", dbText, "Foreign Key to tlkpStates.StateCode"
    SetFieldProperty fld, "Caption", dbText, "State of Play"
    SetFieldProperty fld, "InputMask", dbText, ">LL"

    Debug.Print "Table tblSystemSettings created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tblSystemSettings" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  CORE DATA TABLES
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tblParticipants
' Purpose    : Create the tblParticipants table for pool members
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tblParticipants()
    On Error GoTo ErrorHandler

    If TableExists("tblParticipants") Then
        Debug.Print "Table tblParticipants already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tblParticipants")

    ' --- ParticipantID (AutoNumber PK) ---
    Set fld = td.CreateField("ParticipantID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- FirstName (Short Text, 50) ---
    Set fld = td.CreateField("FirstName", dbText, 50)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- LastName (Short Text, 50) ---
    Set fld = td.CreateField("LastName", dbText, 50)
    fld.Required = True
    fld.AllowZeroLength = False
    td.Fields.Append fld

    ' --- Email (Short Text, 100) — optional ---
    Set fld = td.CreateField("Email", dbText, 100)
    fld.Required = False
    td.Fields.Append fld

    ' --- Phone (Short Text, 20) — optional ---
    Set fld = td.CreateField("Phone", dbText, 20)
    fld.Required = False
    td.Fields.Append fld

    ' --- IsActive (Yes/No) ---
    Set fld = td.CreateField("IsActive", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "True"
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("ParticipantID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tblParticipants")

    Set fld = td.Fields("ParticipantID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated participant identifier"
    SetFieldProperty fld, "Caption", dbText, "Participant ID"

    Set fld = td.Fields("FirstName")
    SetFieldProperty fld, "Description", dbText, "Participant's first name"
    SetFieldProperty fld, "Caption", dbText, "First Name"

    Set fld = td.Fields("LastName")
    SetFieldProperty fld, "Description", dbText, "Participant's last name"
    SetFieldProperty fld, "Caption", dbText, "Last Name"

    Set fld = td.Fields("Email")
    SetFieldProperty fld, "Description", dbText, "Participant's email address"
    SetFieldProperty fld, "Caption", dbText, "Email"

    Set fld = td.Fields("Phone")
    SetFieldProperty fld, "Description", dbText, "Participant's phone number"
    SetFieldProperty fld, "Caption", dbText, "Phone"
    SetFieldProperty fld, "InputMask", dbText, "!\(999"") ""000\-0000;0;_"

    Set fld = td.Fields("IsActive")
    SetFieldProperty fld, "Description", dbText, _
        "Whether this participant is currently active in the pool"
    SetFieldProperty fld, "Caption", dbText, "Active"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Debug.Print "Table tblParticipants created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tblParticipants" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tblDrawings
' Purpose    : Create the tblDrawings table for official Powerball draw results
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tblDrawings()
    On Error GoTo ErrorHandler

    If TableExists("tblDrawings") Then
        Debug.Print "Table tblDrawings already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tblDrawings")

    ' --- DrawingID (AutoNumber PK) ---
    Set fld = td.CreateField("DrawingID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- DrawDate (Date/Time) ---
    Set fld = td.CreateField("DrawDate", dbDate)
    fld.Required = True
    fld.ValidationRule = "Weekday([DrawDate]) In (2,4,7)"
    fld.ValidationText = "Draw date must be a Monday, Wednesday, or Saturday."
    td.Fields.Append fld

    ' --- WB1 through WB5 (Integer, 1-69) ---
    Dim i As Integer
    For i = 1 To 5
        Set fld = td.CreateField("WB" & i, dbInteger)
        fld.Required = True
        fld.ValidationRule = ">=1 And <=69"
        fld.ValidationText = "White ball must be between 1 and 69."
        td.Fields.Append fld
    Next i

    ' --- PB (Integer, 1-26) ---
    Set fld = td.CreateField("PB", dbInteger)
    fld.Required = True
    fld.ValidationRule = ">=1 And <=26"
    fld.ValidationText = "Powerball must be between 1 and 26."
    td.Fields.Append fld

    ' --- JackpotAmount (Currency) — optional ---
    Set fld = td.CreateField("JackpotAmount", dbCurrency)
    fld.Required = False
    fld.DefaultValue = "0"
    fld.ValidationRule = ">=0"
    fld.ValidationText = "Jackpot amount cannot be negative."
    td.Fields.Append fld

    ' --- IsVerified (Yes/No) ---
    Set fld = td.CreateField("IsVerified", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("DrawingID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' --- Unique Index on DrawDate ---
    Set idx = td.CreateIndex("DrawDate")
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("DrawDate")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tblDrawings")

    Set fld = td.Fields("DrawingID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated drawing identifier"
    SetFieldProperty fld, "Caption", dbText, "Drawing ID"

    Set fld = td.Fields("DrawDate")
    SetFieldProperty fld, "Description", dbText, _
        "Official draw date. Must be Mon, Wed, or Sat"
    SetFieldProperty fld, "Caption", dbText, "Draw Date"
    SetFieldProperty fld, "Format", dbText, "Short Date"
    SetFieldProperty fld, "InputMask", dbText, "99/99/0000;0;_"

    For i = 1 To 5
        Set fld = td.Fields("WB" & i)
        SetFieldProperty fld, "Description", dbText, "Winning white ball " & i
        SetFieldProperty fld, "Caption", dbText, "WB " & i
        SetFieldProperty fld, "DecimalPlaces", dbByte, 0
    Next i

    Set fld = td.Fields("PB")
    SetFieldProperty fld, "Description", dbText, "Winning Powerball number"
    SetFieldProperty fld, "Caption", dbText, "Powerball"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    Set fld = td.Fields("JackpotAmount")
    SetFieldProperty fld, "Description", dbText, _
        "Estimated or actual jackpot for this drawing"
    SetFieldProperty fld, "Caption", dbText, "Jackpot Amount"
    SetFieldProperty fld, "Format", dbText, "Currency"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 2

    Set fld = td.Fields("IsVerified")
    SetFieldProperty fld, "Description", dbText, _
        "Whether results have been officially confirmed"
    SetFieldProperty fld, "Caption", dbText, "Verified"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Debug.Print "Table tblDrawings created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tblDrawings" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tblTickets
' Purpose    : Create the tblTickets table for pool ticket entries
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tblTickets()
    On Error GoTo ErrorHandler

    If TableExists("tblTickets") Then
        Debug.Print "Table tblTickets already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tblTickets")

    ' --- TicketID (AutoNumber PK) ---
    Set fld = td.CreateField("TicketID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- DrawingID (Long — FK to tblDrawings.DrawingID) ---
    Set fld = td.CreateField("DrawingID", dbLong)
    fld.Required = True
    td.Fields.Append fld

    ' --- WB1 through WB5 (Integer, 1-69) ---
    Dim i As Integer
    For i = 1 To 5
        Set fld = td.CreateField("WB" & i, dbInteger)
        fld.Required = True
        fld.ValidationRule = ">=1 And <=69"
        fld.ValidationText = "White ball must be between 1 and 69."
        td.Fields.Append fld
    Next i

    ' --- PB (Integer, 1-26) ---
    Set fld = td.CreateField("PB", dbInteger)
    fld.Required = True
    fld.ValidationRule = ">=1 And <=26"
    fld.ValidationText = "Powerball must be between 1 and 26."
    td.Fields.Append fld

    ' --- IsPowerPlay (Yes/No) ---
    Set fld = td.CreateField("IsPowerPlay", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- IsDoublePlay (Yes/No) ---
    Set fld = td.CreateField("IsDoublePlay", dbBoolean)
    fld.Required = True
    fld.DefaultValue = "False"
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("TicketID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' --- Index on DrawingID (Duplicates OK) ---
    Set idx = td.CreateIndex("DrawingID")
    idx.Unique = False
    Set fld = idx.CreateField("DrawingID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tblTickets")

    Set fld = td.Fields("TicketID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated ticket identifier"
    SetFieldProperty fld, "Caption", dbText, "Ticket ID"

    Set fld = td.Fields("DrawingID")
    SetFieldProperty fld, "Description", dbText, "Foreign Key to tblDrawings.DrawingID"
    SetFieldProperty fld, "Caption", dbText, "Drawing ID"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    For i = 1 To 5
        Set fld = td.Fields("WB" & i)
        SetFieldProperty fld, "Description", dbText, "White ball " & i
        SetFieldProperty fld, "Caption", dbText, "WB " & i
        SetFieldProperty fld, "DecimalPlaces", dbByte, 0
    Next i

    Set fld = td.Fields("PB")
    SetFieldProperty fld, "Description", dbText, "Powerball"
    SetFieldProperty fld, "Caption", dbText, "Powerball"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    Set fld = td.Fields("IsPowerPlay")
    SetFieldProperty fld, "Description", dbText, _
        "Whether this ticket includes Power Play"
    SetFieldProperty fld, "Caption", dbText, "Power Play"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Set fld = td.Fields("IsDoublePlay")
    SetFieldProperty fld, "Description", dbText, _
        "Whether this ticket includes Double Play"
    SetFieldProperty fld, "Caption", dbText, "Double Play"
    SetFieldProperty fld, "Format", dbText, "Yes/No"

    Debug.Print "Table tblTickets created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tblTickets" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateTable_tblContributions
' Purpose    : Create the tblContributions table for participant payments
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateTable_tblContributions()
    On Error GoTo ErrorHandler

    If TableExists("tblContributions") Then
        Debug.Print "Table tblContributions already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index

    Set db = CurrentDb()
    Set td = db.CreateTableDef("tblContributions")

    ' --- ContributionID (AutoNumber PK) ---
    Set fld = td.CreateField("ContributionID", dbLong)
    fld.Attributes = dbAutoIncrField
    fld.Required = True
    td.Fields.Append fld

    ' --- ParticipantID (Long — FK to tblParticipants.ParticipantID) ---
    Set fld = td.CreateField("ParticipantID", dbLong)
    fld.Required = True
    td.Fields.Append fld

    ' --- DrawingID (Long — FK to tblDrawings.DrawingID) ---
    Set fld = td.CreateField("DrawingID", dbLong)
    fld.Required = True
    td.Fields.Append fld

    ' --- AmountPaid (Currency) ---
    Set fld = td.CreateField("AmountPaid", dbCurrency)
    fld.Required = True
    fld.ValidationRule = ">0"
    fld.ValidationText = "Amount paid must be greater than zero."
    td.Fields.Append fld

    ' --- DatePaid (Date/Time) ---
    Set fld = td.CreateField("DatePaid", dbDate)
    fld.Required = True
    fld.DefaultValue = "=Date()"
    td.Fields.Append fld

    ' --- Primary Key ---
    Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Unique = True
    idx.Required = True
    Set fld = idx.CreateField("ContributionID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' --- Index on ParticipantID (Duplicates OK) ---
    Set idx = td.CreateIndex("ParticipantID")
    idx.Unique = False
    Set fld = idx.CreateField("ParticipantID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' --- Index on DrawingID (Duplicates OK) ---
    Set idx = td.CreateIndex("DrawingID")
    idx.Unique = False
    Set fld = idx.CreateField("DrawingID")
    idx.Fields.Append fld
    td.Indexes.Append idx

    ' Append table to database
    db.TableDefs.Append td
    db.TableDefs.Refresh

    ' --- Set custom properties ---
    Set td = db.TableDefs("tblContributions")

    Set fld = td.Fields("ContributionID")
    SetFieldProperty fld, "Description", dbText, "Auto-generated contribution identifier"
    SetFieldProperty fld, "Caption", dbText, "Contribution ID"

    Set fld = td.Fields("ParticipantID")
    SetFieldProperty fld, "Description", dbText, "Foreign Key to tblParticipants.ParticipantID"
    SetFieldProperty fld, "Caption", dbText, "Participant ID"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    Set fld = td.Fields("DrawingID")
    SetFieldProperty fld, "Description", dbText, "Foreign Key to tblDrawings.DrawingID"
    SetFieldProperty fld, "Caption", dbText, "Drawing ID"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 0

    Set fld = td.Fields("AmountPaid")
    SetFieldProperty fld, "Description", dbText, "Amount contributed by this participant"
    SetFieldProperty fld, "Caption", dbText, "Amount Paid"
    SetFieldProperty fld, "Format", dbText, "Currency"
    SetFieldProperty fld, "DecimalPlaces", dbByte, 2

    Set fld = td.Fields("DatePaid")
    SetFieldProperty fld, "Description", dbText, "Date payment was received"
    SetFieldProperty fld, "Caption", dbText, "Date Paid"
    SetFieldProperty fld, "Format", dbText, "Short Date"
    SetFieldProperty fld, "InputMask", dbText, "99/99/0000;0;_"

    Debug.Print "Table tblContributions created successfully."

Exit_Procedure:
    Set fld = Nothing
    Set idx = Nothing
    Set td = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateTable_tblContributions" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

' ======================================================================================
'  RELATIONSHIPS
' ======================================================================================

'---------------------------------------------------------------------------------------
' Name       : CreateAllRelationships
' Purpose    : Create all foreign key relationships with referential integrity
'              and cascade updates
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateAllRelationships()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rel As DAO.Relation
    Dim fld As DAO.Field

    Set db = CurrentDb()

    ' --- tlkpStates.StateCode -> tblSystemSettings.StateOfPlay ---
    If Not RelationExists("rel_tlkpStates_tblSystemSettings") Then
        Set rel = db.CreateRelation("rel_tlkpStates_tblSystemSettings", _
                                    "tlkpStates", "tblSystemSettings", _
                                    dbRelationUpdateCascade)
        Set fld = rel.CreateField("StateCode")
        fld.ForeignName = "StateOfPlay"
        rel.Fields.Append fld
        db.Relations.Append rel
        Debug.Print "Relationship rel_tlkpStates_tblSystemSettings created."
    End If

    ' --- tblDrawings.DrawingID -> tblTickets.DrawingID ---
    If Not RelationExists("rel_tblDrawings_tblTickets") Then
        Set rel = db.CreateRelation("rel_tblDrawings_tblTickets", _
                                    "tblDrawings", "tblTickets", _
                                    dbRelationUpdateCascade)
        Set fld = rel.CreateField("DrawingID")
        fld.ForeignName = "DrawingID"
        rel.Fields.Append fld
        db.Relations.Append rel
        Debug.Print "Relationship rel_tblDrawings_tblTickets created."
    End If

    ' --- tblDrawings.DrawingID -> tblContributions.DrawingID ---
    If Not RelationExists("rel_tblDrawings_tblContributions") Then
        Set rel = db.CreateRelation("rel_tblDrawings_tblContributions", _
                                    "tblDrawings", "tblContributions", _
                                    dbRelationUpdateCascade)
        Set fld = rel.CreateField("DrawingID")
        fld.ForeignName = "DrawingID"
        rel.Fields.Append fld
        db.Relations.Append rel
        Debug.Print "Relationship rel_tblDrawings_tblContributions created."
    End If

    ' --- tblParticipants.ParticipantID -> tblContributions.ParticipantID ---
    If Not RelationExists("rel_tblParticipants_tblContributions") Then
        Set rel = db.CreateRelation("rel_tblParticipants_tblContributions", _
                                    "tblParticipants", "tblContributions", _
                                    dbRelationUpdateCascade)
        Set fld = rel.CreateField("ParticipantID")
        fld.ForeignName = "ParticipantID"
        rel.Fields.Append fld
        db.Relations.Append rel
        Debug.Print "Relationship rel_tblParticipants_tblContributions created."
    End If

Exit_Procedure:
    Set fld = Nothing
    Set rel = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateAllRelationships" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub
