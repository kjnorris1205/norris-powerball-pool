Attribute VB_Name = "modBatchSQL"
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
