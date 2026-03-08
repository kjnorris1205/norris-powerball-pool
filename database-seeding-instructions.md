# Database Seeding Instructions

Populate the lookup tables (`tlkpStates` and `tlkpPrizeTiers`) with default data using DAO. Tables must already exist — run `CreateAllTables` from [database-setup-instructions.md](database-setup-instructions.md) first.

## Setup Steps

1. Open `NorrisPowerballPool.accdb` in Microsoft Access.
2. Press **Alt+F11** to open the VBA editor.
3. Go to **Insert** → **Module**. A new module window opens.
4. Copy the entire VBA code block below and paste it into the new module.
5. In the **Properties** window (press **F4** if not visible), change the module `(Name)` to `modSeedData`.
6. Press **Ctrl+S** to save.
7. Press **Ctrl+G** to open the **Immediate Window**.
8. Type `SeedAllLookupTables` and press **Enter**.
9. A confirmation message appears when seeding is complete (or reports which tables were skipped because they already contain data).

## VBA Code: `modSeedData`

```vb
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

'---------------------------------------------------------------------------------------
' Name       : SeedAllLookupTables
' Purpose    : Orchestrate seeding of all lookup tables. Safe to run multiple times —
'              skips any table that already contains records.
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub SeedAllLookupTables()
    On Error GoTo ErrorHandler

    Dim strResults As String
    strResults = ""

    strResults = strResults & SeedStates() & vbCrLf
    strResults = strResults & SeedPrizeTiers() & vbCrLf

    MsgBox "Seed Results:" & vbCrLf & vbCrLf & strResults, _
           vbInformation, APP_TITLE

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: SeedAllLookupTables" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : RecordCount
' Purpose    : Return the number of records in a table (0 if table is empty or missing)
' Parameters : strTableName (String) - Name of the table to count
' Returns    : Long - Number of records
'---------------------------------------------------------------------------------------
Private Function RecordCount(ByVal strTableName As String) As Long
    On Error Resume Next
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT COUNT(*) AS Cnt FROM [" & strTableName & "]", _
                                     dbOpenSnapshot)
    If Err.Number <> 0 Then
        RecordCount = 0
        Err.Clear
    Else
        RecordCount = rs!Cnt
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Name       : AddStateRow
' Purpose    : Add a single row to tlkpStates using a DAO recordset
' Parameters : rs (DAO.Recordset) - Open recordset on tlkpStates
'              strCode (String) - Two-letter state code
'              strName (String) - Full state name
'              dblFedRate (Double) - Federal tax withholding rate
'              dblStateRate (Double) - State tax withholding rate
'              blnLottery (Boolean) - Has state lottery
'              blnPowerPlay (Boolean) - Has Power Play
'              blnDoublePlay (Boolean) - Has Double Play
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub AddStateRow(rs As DAO.Recordset, _
                        ByVal strCode As String, _
                        ByVal strName As String, _
                        ByVal dblFedRate As Double, _
                        ByVal dblStateRate As Double, _
                        ByVal blnLottery As Boolean, _
                        ByVal blnPowerPlay As Boolean, _
                        ByVal blnDoublePlay As Boolean)
    rs.AddNew
    rs!StateCode = strCode
    rs!StateName = strName
    rs!FederalTaxRate = dblFedRate
    rs!StateTaxRate = dblStateRate
    rs!HasStateLottery = blnLottery
    rs!HasPowerPlay = blnPowerPlay
    rs!HasDoublePlay = blnDoublePlay
    rs.Update
End Sub

'---------------------------------------------------------------------------------------
' Name       : SeedStates
' Purpose    : Insert all 51 rows into tlkpStates (50 states + DC) using DAO.
'              Skips if the table already contains any records.
' Parameters : None
' Returns    : String - Status message describing what happened
'---------------------------------------------------------------------------------------
Private Function SeedStates() As String
    On Error GoTo ErrorHandler

    If RecordCount("tlkpStates") > 0 Then
        SeedStates = "tlkpStates: SKIPPED (already contains data)"
        Exit Function
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tlkpStates", dbOpenTable)

    AddStateRow rs, "AL", "Alabama", 0.24, 0#, False, False, False
    AddStateRow rs, "AK", "Alaska", 0.24, 0#, False, False, False
    AddStateRow rs, "AZ", "Arizona", 0.24, 0.05, True, True, False
    AddStateRow rs, "AR", "Arkansas", 0.24, 0.055, True, True, False
    AddStateRow rs, "CA", "California", 0.24, 0#, True, True, False
    AddStateRow rs, "CO", "Colorado", 0.24, 0.04, True, True, True
    AddStateRow rs, "CT", "Connecticut", 0.24, 0.0699, True, True, False
    AddStateRow rs, "DE", "Delaware", 0.24, 0#, True, True, False
    AddStateRow rs, "DC", "District of Columbia", 0.24, 0.0875, True, True, False
    AddStateRow rs, "FL", "Florida", 0.24, 0#, True, True, True
    AddStateRow rs, "GA", "Georgia", 0.24, 0.055, True, True, False
    AddStateRow rs, "HI", "Hawaii", 0.24, 0#, False, False, False
    AddStateRow rs, "ID", "Idaho", 0.24, 0.058, True, True, False
    AddStateRow rs, "IL", "Illinois", 0.24, 0.0495, True, True, False
    AddStateRow rs, "IN", "Indiana", 0.24, 0.0323, True, True, True
    AddStateRow rs, "IA", "Iowa", 0.24, 0.06, True, True, False
    AddStateRow rs, "KS", "Kansas", 0.24, 0.05, True, True, False
    AddStateRow rs, "KY", "Kentucky", 0.24, 0.05, True, True, False
    AddStateRow rs, "LA", "Louisiana", 0.24, 0.05, True, True, False
    AddStateRow rs, "ME", "Maine", 0.24, 0.05, True, True, False
    AddStateRow rs, "MD", "Maryland", 0.24, 0.0875, True, True, False
    AddStateRow rs, "MA", "Massachusetts", 0.24, 0.05, True, True, False
    AddStateRow rs, "MI", "Michigan", 0.24, 0.0425, True, True, False
    AddStateRow rs, "MN", "Minnesota", 0.24, 0.0785, True, True, False
    AddStateRow rs, "MS", "Mississippi", 0.24, 0.05, False, False, False
    AddStateRow rs, "MO", "Missouri", 0.24, 0.0495, True, True, True
    AddStateRow rs, "MT", "Montana", 0.24, 0.069, True, True, False
    AddStateRow rs, "NE", "Nebraska", 0.24, 0.0684, True, True, False
    AddStateRow rs, "NV", "Nevada", 0.24, 0#, False, False, False
    AddStateRow rs, "NH", "New Hampshire", 0.24, 0#, True, True, False
    AddStateRow rs, "NJ", "New Jersey", 0.24, 0.08, True, True, False
    AddStateRow rs, "NM", "New Mexico", 0.24, 0.059, True, True, False
    AddStateRow rs, "NY", "New York", 0.24, 0.0882, True, True, False
    AddStateRow rs, "NC", "North Carolina", 0.24, 0.0525, True, True, False
    AddStateRow rs, "ND", "North Dakota", 0.24, 0.029, True, True, False
    AddStateRow rs, "OH", "Ohio", 0.24, 0.04, True, True, False
    AddStateRow rs, "OK", "Oklahoma", 0.24, 0.0475, True, True, False
    AddStateRow rs, "OR", "Oregon", 0.24, 0.09, True, True, False
    AddStateRow rs, "PA", "Pennsylvania", 0.24, 0.0307, True, True, True
    AddStateRow rs, "RI", "Rhode Island", 0.24, 0.0599, True, True, False
    AddStateRow rs, "SC", "South Carolina", 0.24, 0.07, True, True, False
    AddStateRow rs, "SD", "South Dakota", 0.24, 0#, True, True, False
    AddStateRow rs, "TN", "Tennessee", 0.24, 0#, True, True, True
    AddStateRow rs, "TX", "Texas", 0.24, 0#, True, True, True
    AddStateRow rs, "UT", "Utah", 0.24, 0#, False, False, False
    AddStateRow rs, "VT", "Vermont", 0.24, 0.06, True, True, False
    AddStateRow rs, "VA", "Virginia", 0.24, 0.04, True, True, False
    AddStateRow rs, "WA", "Washington", 0.24, 0#, True, True, False
    AddStateRow rs, "WV", "West Virginia", 0.24, 0.065, True, True, False
    AddStateRow rs, "WI", "Wisconsin", 0.24, 0.0765, True, True, False
    AddStateRow rs, "WY", "Wyoming", 0.24, 0#, True, True, False

    rs.Close
    SeedStates = "tlkpStates: Seeded with 51 rows."
    Debug.Print SeedStates

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    SeedStates = "tlkpStates: ERROR - " & Err.Description
    Debug.Print SeedStates
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : AddPrizeTierRow
' Purpose    : Add a single row to tlkpPrizeTiers using a DAO recordset
' Parameters : rs (DAO.Recordset) - Open recordset on tlkpPrizeTiers
'              intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
'              strName (String) - Display name for this prize tier
'              curAmount (Currency) - Default prize amount
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub AddPrizeTierRow(rs As DAO.Recordset, _
                            ByVal intWhiteMatches As Integer, _
                            ByVal blnPBMatch As Boolean, _
                            ByVal strName As String, _
                            ByVal curAmount As Currency)
    rs.AddNew
    rs!WhiteBallMatches = intWhiteMatches
    rs!PowerballMatch = blnPBMatch
    rs!PrizeName = strName
    rs!DefaultPrizeAmount = curAmount
    rs.Update
End Sub

'---------------------------------------------------------------------------------------
' Name       : SeedPrizeTiers
' Purpose    : Insert the 9 Powerball prize tiers into tlkpPrizeTiers using DAO.
'              Skips if the table already contains any records.
' Parameters : None
' Returns    : String - Status message describing what happened
'---------------------------------------------------------------------------------------
Private Function SeedPrizeTiers() As String
    On Error GoTo ErrorHandler

    If RecordCount("tlkpPrizeTiers") > 0 Then
        SeedPrizeTiers = "tlkpPrizeTiers: SKIPPED (already contains data)"
        Exit Function
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb()
    Set rs = db.OpenRecordset("tlkpPrizeTiers", dbOpenTable)

    AddPrizeTierRow rs, 5, True, "Jackpot (5+PB)", CCur(0)
    AddPrizeTierRow rs, 5, False, "Match 5", CCur(1000000)
    AddPrizeTierRow rs, 4, True, "Match 4+PB", CCur(50000)
    AddPrizeTierRow rs, 4, False, "Match 4", CCur(100)
    AddPrizeTierRow rs, 3, True, "Match 3+PB", CCur(100)
    AddPrizeTierRow rs, 3, False, "Match 3", CCur(7)
    AddPrizeTierRow rs, 2, True, "Match 2+PB", CCur(7)
    AddPrizeTierRow rs, 1, True, "Match 1+PB", CCur(4)
    AddPrizeTierRow rs, 0, True, "Match PB Only", CCur(4)

    rs.Close
    SeedPrizeTiers = "tlkpPrizeTiers: Seeded with 9 rows."
    Debug.Print SeedPrizeTiers

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    SeedPrizeTiers = "tlkpPrizeTiers: ERROR - " & Err.Description
    Debug.Print SeedPrizeTiers
    Resume Exit_Function
End Function
```

## Notes

- **Safe to re-run.** Each seed function checks whether its table already contains records and skips if so. No duplicates will be created.
- **To re-seed from scratch,** first clear the table: `CurrentDb.Execute "DELETE FROM tlkpStates", dbFailOnError` (same for `tlkpPrizeTiers`), then run `SeedAllLookupTables` again.
- **Data review:** After seeding, open each table in Datasheet View to verify the data looks correct. Pay attention to `FederalTaxRate`, `StateTaxRate`, and the play-option flags.
