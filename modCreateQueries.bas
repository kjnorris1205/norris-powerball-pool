Attribute VB_Name = "modCreateQueries"
Option Compare Database
Option Explicit

Private Const APP_TITLE As String = "Norris Powerball Pool"

'---------------------------------------------------------------------------------------
' Name       : CreateAllQueries
' Purpose    : Orchestrate creation of all MVP queries
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub CreateAllQueries()
    On Error GoTo ErrorHandler

    CreateQuery_qryMatchCheck
    CreateQuery_qryWinningTickets
    CreateQuery_qryUnpaidParticipants
    CreateQuery_qryTicketsByDrawing

    MsgBox "All queries created successfully.", vbInformation, APP_TITLE

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateAllQueries" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : QueryExists
' Purpose    : Check if a query already exists in the current database
' Parameters : strQueryName (String) - Name of the query to check
' Returns    : Boolean - True if the query exists
'---------------------------------------------------------------------------------------
Private Function QueryExists(ByVal strQueryName As String) As Boolean
    On Error Resume Next
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs(strQueryName)
    QueryExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Name       : CreateQuery_qryMatchCheck
' Purpose    : Create query that compares ticket entries against drawing results.
'              Counts matching white balls (unordered set comparison) and checks
'              for exact Powerball match.
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateQuery_qryMatchCheck()
    On Error GoTo ErrorHandler

    If QueryExists("qryMatchCheck") Then
        Debug.Print "Query qryMatchCheck already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT t.TicketID, t.DrawingID, d.DrawDate, " & _
             "t.ParticipantID, " & _
             "pa.FirstName & ' ' & pa.LastName AS PurchasedBy, " & _
             "t.WB1, t.WB2, t.WB3, t.WB4, t.WB5, t.PB, " & _
             "t.IsPowerPlay, t.IsDoublePlay, " & _
             "IIf(t.WB1=d.WB1 Or t.WB1=d.WB2 Or t.WB1=d.WB3 Or t.WB1=d.WB4 Or t.WB1=d.WB5,1,0) + " & _
             "IIf(t.WB2=d.WB1 Or t.WB2=d.WB2 Or t.WB2=d.WB3 Or t.WB2=d.WB4 Or t.WB2=d.WB5,1,0) + " & _
             "IIf(t.WB3=d.WB1 Or t.WB3=d.WB2 Or t.WB3=d.WB3 Or t.WB3=d.WB4 Or t.WB3=d.WB5,1,0) + " & _
             "IIf(t.WB4=d.WB1 Or t.WB4=d.WB2 Or t.WB4=d.WB3 Or t.WB4=d.WB4 Or t.WB4=d.WB5,1,0) + " & _
             "IIf(t.WB5=d.WB1 Or t.WB5=d.WB2 Or t.WB5=d.WB3 Or t.WB5=d.WB4 Or t.WB5=d.WB5,1,0) " & _
             "AS WhiteBallMatches, " & _
             "IIf(t.PB=d.PB,True,False) AS PowerballMatch " & _
             "FROM (tblTickets AS t " & _
             "INNER JOIN tblDrawings AS d ON t.DrawingID = d.DrawingID) " & _
             "INNER JOIN tblParticipants AS pa ON t.ParticipantID = pa.ParticipantID " & _
             "WHERE d.WB1 Is Not Null AND d.WB2 Is Not Null AND d.WB3 Is Not Null " & _
             "AND d.WB4 Is Not Null AND d.WB5 Is Not Null AND d.PB Is Not Null"

    db.CreateQueryDef "qryMatchCheck", strSQL
    Debug.Print "Query qryMatchCheck created successfully."

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateQuery_qryMatchCheck" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateQuery_qryWinningTickets
' Purpose    : Create query that filters qryMatchCheck to only rows that match
'              a prize tier in tlkpPrizeTiers (0+PB or better)
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateQuery_qryWinningTickets()
    On Error GoTo ErrorHandler

    If QueryExists("qryWinningTickets") Then
        Debug.Print "Query qryWinningTickets already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT mc.TicketID, mc.DrawingID, mc.DrawDate, " & _
             "mc.ParticipantID, mc.PurchasedBy, " & _
             "mc.WB1, mc.WB2, mc.WB3, mc.WB4, mc.WB5, mc.PB, " & _
             "mc.IsPowerPlay, mc.IsDoublePlay, " & _
             "mc.WhiteBallMatches, mc.PowerballMatch, " & _
             "p.PrizeName, p.DefaultPrizeAmount " & _
             "FROM qryMatchCheck AS mc " & _
             "INNER JOIN tlkpPrizeTiers AS p " & _
             "ON mc.WhiteBallMatches = p.WhiteBallMatches " & _
             "AND mc.PowerballMatch = p.PowerballMatch"

    db.CreateQueryDef "qryWinningTickets", strSQL
    Debug.Print "Query qryWinningTickets created successfully."

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateQuery_qryWinningTickets" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateQuery_qryUnpaidParticipants
' Purpose    : Create parameterized query that finds active participants with no
'              contribution record for a given DrawingID
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateQuery_qryUnpaidParticipants()
    On Error GoTo ErrorHandler

    If QueryExists("qryUnpaidParticipants") Then
        Debug.Print "Query qryUnpaidParticipants already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "PARAMETERS [prmDrawingID] Long; " & _
             "SELECT p.ParticipantID, p.FirstName, p.LastName, " & _
             "p.Email, p.Phone " & _
             "FROM tblParticipants AS p " & _
             "LEFT JOIN (SELECT ContributionID, ParticipantID FROM tblContributions WHERE DrawingID = [prmDrawingID]) AS c " & _
             "ON p.ParticipantID = c.ParticipantID " & _
             "WHERE p.IsActive = True " & _
             "AND c.ContributionID Is Null " & _
             "ORDER BY p.LastName, p.FirstName"

    db.CreateQueryDef "qryUnpaidParticipants", strSQL
    Debug.Print "Query qryUnpaidParticipants created successfully."

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateQuery_qryUnpaidParticipants" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub

'---------------------------------------------------------------------------------------
' Name       : CreateQuery_qryTicketsByDrawing
' Purpose    : Create parameterized query that lists all tickets for a given DrawingID
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Private Sub CreateQuery_qryTicketsByDrawing()
    On Error GoTo ErrorHandler

    If QueryExists("qryTicketsByDrawing") Then
        Debug.Print "Query qryTicketsByDrawing already exists - skipped."
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "PARAMETERS [prmDrawingID] Long; " & _
             "SELECT t.TicketID, t.DrawingID, " & _
             "t.ParticipantID, " & _
             "pa.FirstName & ' ' & pa.LastName AS PurchasedBy, " & _
             "t.WB1, t.WB2, t.WB3, t.WB4, t.WB5, t.PB, " & _
             "t.IsPowerPlay, t.IsDoublePlay " & _
             "FROM tblTickets AS t " & _
             "INNER JOIN tblParticipants AS pa ON t.ParticipantID = pa.ParticipantID " & _
             "WHERE t.DrawingID = [prmDrawingID] " & _
             "ORDER BY t.TicketID"

    db.CreateQueryDef "qryTicketsByDrawing", strSQL
    Debug.Print "Query qryTicketsByDrawing created successfully."

Exit_Procedure:
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: CreateQuery_qryTicketsByDrawing" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, APP_TITLE
    Resume Exit_Procedure
End Sub
