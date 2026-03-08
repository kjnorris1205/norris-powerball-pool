# Batch SQL Runner Utility

A developer utility module (`modBatchSQL`) that executes multiple SQL statements in batch from a `.sql` file. Useful for seeding tables, running migrations, or applying bulk changes without manually executing statements one at a time.

## How It Works

- Accepts a block of SQL containing multiple statements separated by semicolons.
- Splits the block on `;` and executes each non-empty statement individually via `CurrentDb.Execute` with `dbFailOnError`.
- Reports a summary: total statements executed, number succeeded, number failed, and details for any failures.
- Entry point: **`RunSQLFromFile`** — opens a file-picker dialog to select a `.sql` or `.txt` file.

## Setup Instructions

1. Open `NorrisPowerballPool.accdb` in Microsoft Access.
2. Press **Alt+F11** to open the VBA editor.
3. In the VBA editor, go to **Insert** → **Module**. A new module window opens.
4. Copy the entire VBA code block below and paste it into the new module.
5. In the **Properties** window (press **F4** if not visible), change the module `(Name)` to `modBatchSQL`.
6. Press **Ctrl+S** to save.

## VBA Code: `modBatchSQL`

```vb
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Name       : ExecuteBatchSQL
' Purpose    : Split a multi-statement SQL string on semicolons and execute each
'              statement individually. Returns a summary of results.
' Parameters : strSQL (String) - One or more SQL statements separated by semicolons
' Returns    : None (displays summary via MsgBox)
'---------------------------------------------------------------------------------------
Public Sub ExecuteBatchSQL(ByVal strSQL As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim arrStatements() As String
    Dim strStatement As String
    Dim lngTotal As Long
    Dim lngSuccess As Long
    Dim lngFail As Long
    Dim strFailures As String
    Dim i As Long

    Set db = CurrentDb()

    ' Split on semicolons
    arrStatements = Split(strSQL, ";")

    lngTotal = 0
    lngSuccess = 0
    lngFail = 0
    strFailures = ""

    For i = LBound(arrStatements) To UBound(arrStatements)
        strStatement = arrStatements(i)

        ' Strip spaces, tabs, and line breaks
        strStatement = Replace(strStatement, vbCrLf, "")
        strStatement = Replace(strStatement, vbCr, "")
        strStatement = Replace(strStatement, vbLf, "")
        strStatement = Trim$(strStatement)

        ' Skip empty or whitespace-only segments
        If Len(strStatement) = 0 Then GoTo NextStatement

        ' Skip lines that are only comments
        If Left$(strStatement, 2) = "--" And InStr(strStatement, vbCrLf) = 0 Then GoTo NextStatement

        lngTotal = lngTotal + 1

        On Error Resume Next
        db.Execute strStatement, dbFailOnError
        If Err.Number <> 0 Then
            lngFail = lngFail + 1
            strFailures = strFailures & "Statement #" & lngTotal & ": " & _
                          Err.Description & vbCrLf & _
                          "  SQL: " & Left$(strStatement, 120) & vbCrLf & vbCrLf
            Err.Clear
        Else
            lngSuccess = lngSuccess + 1
        End If
        On Error GoTo ErrorHandler

NextStatement:
    Next i

    ' Build summary message
    Dim strSummary As String
    strSummary = "Batch SQL Execution Complete" & vbCrLf & vbCrLf & _
                 "Total statements: " & lngTotal & vbCrLf & _
                 "Succeeded: " & lngSuccess & vbCrLf & _
                 "Failed: " & lngFail

    If lngFail > 0 Then
        strSummary = strSummary & vbCrLf & vbCrLf & "--- Failure Details ---" & vbCrLf & strFailures
        MsgBox strSummary, vbExclamation, "Norris Powerball Pool"
    Else
        MsgBox strSummary, vbInformation, "Norris Powerball Pool"
    End If

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: ExecuteBatchSQL" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, "Norris Powerball Pool"
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : RunSQLFromFile
' Purpose    : Open a file dialog to select a .sql or .txt file, read its contents,
'              and execute the SQL statements in batch
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub RunSQLFromFile()
    On Error GoTo ErrorHandler

    Dim fd As Object
    Dim strFilePath As String
    Dim intFileNum As Integer
    Dim strSQL As String
    Dim strLine As String

    ' Use Application.FileDialog for file picker
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker = 3
    With fd
        .Title = "Select a SQL File to Execute"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "SQL Files", "*.sql"
        .Filters.Add "Text Files", "*.txt"
        .Filters.Add "All Files", "*.*"
        .InitialFileName = CurrentProject.Path & "\"

        If .Show = 0 Then
            ' User cancelled
            GoTo Exit_Procedure
        End If

        strFilePath = .SelectedItems(1)
    End With

    ' Read the file contents
    intFileNum = FreeFile
    Open strFilePath For Input As #intFileNum

    strSQL = ""
    Do Until EOF(intFileNum)
        Line Input #intFileNum, strLine
        strSQL = strSQL & strLine & vbCrLf
    Loop

    Close #intFileNum

    If Len(Trim$(strSQL)) = 0 Then
        MsgBox "The selected file is empty: " & strFilePath, _
               vbExclamation, "Norris Powerball Pool"
        GoTo Exit_Procedure
    End If

    ' Confirm before executing
    Dim lngResponse As Long
    lngResponse = MsgBox("Execute SQL from file:" & vbCrLf & _
                         strFilePath & vbCrLf & vbCrLf & _
                         "Preview (first 300 chars):" & vbCrLf & _
                         Left$(strSQL, 300), _
                         vbYesNo + vbQuestion, "Norris Powerball Pool")

    If lngResponse = vbYes Then
        ExecuteBatchSQL strSQL
    End If

Exit_Procedure:
    Set fd = Nothing
    Exit Sub

ErrorHandler:
    ' Ensure file is closed on error
    If intFileNum > 0 Then Close #intFileNum
    MsgBox "An error occurred in: RunSQLFromFile" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, "Norris Powerball Pool"
    Resume Exit_Procedure
End Sub
```

## How to Use

1. Save your SQL statements to a `.sql` or `.txt` file. Statements must be separated by semicolons.
2. In Access, press **Alt+F11** to open the VBA editor.
3. Press **Ctrl+G** to open the **Immediate Window**.
4. Type `RunSQLFromFile` and press **Enter**.
5. A file-picker dialog opens. Navigate to your SQL file and click **Open**.
6. A confirmation dialog shows the file path and a preview. Click **Yes** to execute.
7. A summary dialog reports how many statements succeeded or failed.

## SQL File Format

Statements must be separated by semicolons. Each statement is executed individually. Example:

```sql
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AL','Alabama',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AK','Alaska',0.24,0.00,FALSE,FALSE,FALSE);
DELETE FROM tblContributions WHERE DrawingID = 99;
UPDATE tblParticipants SET IsActive = FALSE WHERE ParticipantID = 5;
```

## Notes

- **Semicolon delimiter:** Every statement must end with a `;`. The runner splits on this character.
- **No SELECT statements:** `CurrentDb.Execute` runs action queries only (INSERT, UPDATE, DELETE, DDL). SELECT statements will error — use Access queries for those.
- **Transaction behavior:** Each statement executes independently. If statement #3 fails, statements #1 and #2 are already committed. The summary tells you exactly which ones failed and why.
- **Safety:** A confirmation dialog shows the file path and a preview before executing anything.