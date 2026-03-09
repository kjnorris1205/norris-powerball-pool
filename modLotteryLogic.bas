Attribute VB_Name = "modLotteryLogic"
Option Compare Database
Option Explicit

' Powerball game constants
Public Const MAX_WHITE_BALLS As Integer = 5
Public Const MIN_WHITE_BALL As Integer = 1
Public Const MAX_WHITE_BALL As Integer = 69
Public Const MIN_POWERBALL As Integer = 1
Public Const MAX_POWERBALL As Integer = 26
Public Const TOTAL_PRIZE_TIERS As Integer = 9

' Power Play constants
Public Const MIN_POWER_PLAY_MULTIPLIER As Integer = 2
Public Const MAX_POWER_PLAY_MULTIPLIER As Integer = 10
Public Const POWER_PLAY_MATCH5_AMOUNT As Currency = 2000000

Private Const APP_TITLE As String = "Norris Powerball Pool"

'---------------------------------------------------------------------------------------
' Name       : ValidateWhiteBallsUnique
' Purpose    : Check that all five white ball values are distinct
' Parameters : intWB1 through intWB5 (Integer) - The five white ball numbers
' Returns    : Boolean - True if all five values are unique
'---------------------------------------------------------------------------------------
Public Function ValidateWhiteBallsUnique(ByVal intWB1 As Integer, _
                                         ByVal intWB2 As Integer, _
                                         ByVal intWB3 As Integer, _
                                         ByVal intWB4 As Integer, _
                                         ByVal intWB5 As Integer) As Boolean
    Dim arrBalls(1 To 5) As Integer
    Dim i As Integer
    Dim j As Integer

    arrBalls(1) = intWB1
    arrBalls(2) = intWB2
    arrBalls(3) = intWB3
    arrBalls(4) = intWB4
    arrBalls(5) = intWB5

    For i = 1 To 4
        For j = i + 1 To 5
            If arrBalls(i) = arrBalls(j) Then
                ValidateWhiteBallsUnique = False
                Exit Function
            End If
        Next j
    Next i

    ValidateWhiteBallsUnique = True
End Function

'---------------------------------------------------------------------------------------
' Name       : ValidateBallInRange
' Purpose    : Check that a ball number falls within the valid range
' Parameters : intBall (Integer) - The ball number to validate
'              intMin (Integer) - Minimum allowed value
'              intMax (Integer) - Maximum allowed value
' Returns    : Boolean - True if the ball is within range
'---------------------------------------------------------------------------------------
Public Function ValidateBallInRange(ByVal intBall As Integer, _
                                    ByVal intMin As Integer, _
                                    ByVal intMax As Integer) As Boolean
    ValidateBallInRange = (intBall >= intMin And intBall <= intMax)
End Function

'---------------------------------------------------------------------------------------
' Name       : ValidateTicketNumbers
' Purpose    : Full validation of a ticket's ball numbers (range + uniqueness)
' Parameters : intWB1 through intWB5 (Integer) - White ball numbers
'              intPB (Integer) - Powerball number
'              strErrorMsg (String) - ByRef, receives error description if invalid
' Returns    : Boolean - True if all numbers are valid
'---------------------------------------------------------------------------------------
Public Function ValidateTicketNumbers(ByVal intWB1 As Integer, _
                                       ByVal intWB2 As Integer, _
                                       ByVal intWB3 As Integer, _
                                       ByVal intWB4 As Integer, _
                                       ByVal intWB5 As Integer, _
                                       ByVal intPB As Integer, _
                                       ByRef strErrorMsg As String) As Boolean
    Dim i As Integer
    Dim arrBalls(1 To 5) As Integer

    arrBalls(1) = intWB1
    arrBalls(2) = intWB2
    arrBalls(3) = intWB3
    arrBalls(4) = intWB4
    arrBalls(5) = intWB5

    ' Validate each white ball is in range
    For i = 1 To MAX_WHITE_BALLS
        If Not ValidateBallInRange(arrBalls(i), MIN_WHITE_BALL, MAX_WHITE_BALL) Then
            strErrorMsg = "White ball " & i & " must be between " & _
                          MIN_WHITE_BALL & " and " & MAX_WHITE_BALL & "."
            ValidateTicketNumbers = False
            Exit Function
        End If
    Next i

    ' Validate Powerball is in range
    If Not ValidateBallInRange(intPB, MIN_POWERBALL, MAX_POWERBALL) Then
        strErrorMsg = "Powerball must be between " & _
                      MIN_POWERBALL & " and " & MAX_POWERBALL & "."
        ValidateTicketNumbers = False
        Exit Function
    End If

    ' Validate white balls are unique
    If Not ValidateWhiteBallsUnique(intWB1, intWB2, intWB3, intWB4, intWB5) Then
        strErrorMsg = "All five white balls must be different numbers."
        ValidateTicketNumbers = False
        Exit Function
    End If

    strErrorMsg = ""
    ValidateTicketNumbers = True
End Function

'---------------------------------------------------------------------------------------
' Name       : CountWhiteBallMatches
' Purpose    : Count how many of the ticket's white balls appear in the drawing's
'              white balls (unordered set comparison)
' Parameters : intTkt1..intTkt5 (Integer) - Ticket white balls
'              intDrw1..intDrw5 (Integer) - Drawing white balls
' Returns    : Integer - Number of matching white balls (0-5)
'---------------------------------------------------------------------------------------
Public Function CountWhiteBallMatches(ByVal intTkt1 As Integer, _
                                       ByVal intTkt2 As Integer, _
                                       ByVal intTkt3 As Integer, _
                                       ByVal intTkt4 As Integer, _
                                       ByVal intTkt5 As Integer, _
                                       ByVal intDrw1 As Integer, _
                                       ByVal intDrw2 As Integer, _
                                       ByVal intDrw3 As Integer, _
                                       ByVal intDrw4 As Integer, _
                                       ByVal intDrw5 As Integer) As Integer
    Dim arrTicket(1 To 5) As Integer
    Dim arrDraw(1 To 5) As Integer
    Dim intMatches As Integer
    Dim i As Integer
    Dim j As Integer

    arrTicket(1) = intTkt1: arrTicket(2) = intTkt2: arrTicket(3) = intTkt3
    arrTicket(4) = intTkt4: arrTicket(5) = intTkt5

    arrDraw(1) = intDrw1: arrDraw(2) = intDrw2: arrDraw(3) = intDrw3
    arrDraw(4) = intDrw4: arrDraw(5) = intDrw5

    intMatches = 0
    For i = 1 To MAX_WHITE_BALLS
        For j = 1 To MAX_WHITE_BALLS
            If arrTicket(i) = arrDraw(j) Then
                intMatches = intMatches + 1
                Exit For
            End If
        Next j
    Next i

    CountWhiteBallMatches = intMatches
End Function

'---------------------------------------------------------------------------------------
' Name       : GetPrizeTierName
' Purpose    : Look up the prize tier name for a given match result
' Parameters : intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
' Returns    : String - Prize tier name (empty string if no matching tier)
'---------------------------------------------------------------------------------------
Public Function GetPrizeTierName(ByVal intWhiteMatches As Integer, _
                                  ByVal blnPBMatch As Boolean) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT PrizeName FROM tlkpPrizeTiers " & _
             "WHERE WhiteBallMatches = " & intWhiteMatches & _
             " AND PowerballMatch = " & CInt(blnPBMatch)

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        GetPrizeTierName = rs!PrizeName
    Else
        GetPrizeTierName = ""
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetPrizeTierName = ""
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : GetPrizeAmount
' Purpose    : Look up the default prize amount for a given match result
' Parameters : intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
' Returns    : Currency - Default prize amount (0 if no matching tier)
'---------------------------------------------------------------------------------------
Public Function GetPrizeAmount(ByVal intWhiteMatches As Integer, _
                                ByVal blnPBMatch As Boolean) As Currency
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT DefaultPrizeAmount FROM tlkpPrizeTiers " & _
             "WHERE WhiteBallMatches = " & intWhiteMatches & _
             " AND PowerballMatch = " & CInt(blnPBMatch)

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        GetPrizeAmount = rs!DefaultPrizeAmount
    Else
        GetPrizeAmount = 0
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetPrizeAmount = 0
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : GetPowerPlayPrizeAmount
' Purpose    : Calculate the Power Play adjusted prize amount.
'              Rules: Jackpot is not affected. Match 5 (no PB) is always
'              $2,000,000. All other tiers multiply by the drawn multiplier.
' Parameters : intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
'              intMultiplier (Integer) - The Power Play multiplier drawn (2-10)
' Returns    : Currency - Adjusted prize amount
'---------------------------------------------------------------------------------------
Public Function GetPowerPlayPrizeAmount(ByVal intWhiteMatches As Integer, _
                                         ByVal blnPBMatch As Boolean, _
                                         ByVal intMultiplier As Integer) As Currency
    On Error GoTo ErrorHandler

    Dim curBase As Currency
    curBase = GetPrizeAmount(intWhiteMatches, blnPBMatch)

    ' Jackpot (5+PB) is not affected by Power Play
    If intWhiteMatches = MAX_WHITE_BALLS And blnPBMatch Then
        GetPowerPlayPrizeAmount = curBase
        Exit Function
    End If

    ' Match 5 (no PB) is always $2,000,000 with Power Play
    If intWhiteMatches = MAX_WHITE_BALLS And Not blnPBMatch Then
        GetPowerPlayPrizeAmount = POWER_PLAY_MATCH5_AMOUNT
        Exit Function
    End If

    ' All other tiers multiply by the drawn multiplier
    If intMultiplier >= MIN_POWER_PLAY_MULTIPLIER And _
       intMultiplier <= MAX_POWER_PLAY_MULTIPLIER Then
        GetPowerPlayPrizeAmount = curBase * intMultiplier
    Else
        GetPowerPlayPrizeAmount = curBase
    End If

Exit_Function:
    Exit Function

ErrorHandler:
    GetPowerPlayPrizeAmount = 0
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : GetDoublePlayPrizeTierName
' Purpose    : Look up the Double Play prize tier name for a given match result
' Parameters : intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
' Returns    : String - Prize tier name (empty string if no matching tier)
'---------------------------------------------------------------------------------------
Public Function GetDoublePlayPrizeTierName(ByVal intWhiteMatches As Integer, _
                                            ByVal blnPBMatch As Boolean) As String
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT PrizeName FROM tlkpDoublePlayPrizeTiers " & _
             "WHERE WhiteBallMatches = " & intWhiteMatches & _
             " AND PowerballMatch = " & CInt(blnPBMatch)

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        GetDoublePlayPrizeTierName = rs!PrizeName
    Else
        GetDoublePlayPrizeTierName = ""
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetDoublePlayPrizeTierName = ""
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : GetDoublePlayPrizeAmount
' Purpose    : Look up the Double Play prize amount for a given match result
' Parameters : intWhiteMatches (Integer) - Number of white balls matched
'              blnPBMatch (Boolean) - Whether Powerball was matched
' Returns    : Currency - Default prize amount (0 if no matching tier)
'---------------------------------------------------------------------------------------
Public Function GetDoublePlayPrizeAmount(ByVal intWhiteMatches As Integer, _
                                          ByVal blnPBMatch As Boolean) As Currency
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String

    Set db = CurrentDb()

    strSQL = "SELECT DefaultPrizeAmount FROM tlkpDoublePlayPrizeTiers " & _
             "WHERE WhiteBallMatches = " & intWhiteMatches & _
             " AND PowerballMatch = " & CInt(blnPBMatch)

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)

    If Not rs.EOF Then
        GetDoublePlayPrizeAmount = rs!DefaultPrizeAmount
    Else
        GetDoublePlayPrizeAmount = 0
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    GetDoublePlayPrizeAmount = 0
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : IsStateFeatureAvailable
' Purpose    : Check whether a specific feature (Power Play/Double Play)
'              is available for a given state
' Parameters : strStateCode (String) - Two-letter state code
'              strFeature (String) - "PowerPlay" or "DoublePlay"
' Returns    : Boolean - True if the feature is available
'---------------------------------------------------------------------------------------
Public Function IsStateFeatureAvailable(ByVal strStateCode As String, _
                                         ByVal strFeature As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strField As String

    Select Case strFeature
        Case "PowerPlay"
            strField = "HasPowerPlay"
        Case "DoublePlay"
            strField = "HasDoublePlay"
        Case Else
            IsStateFeatureAvailable = False
            Exit Function
    End Select

    Set db = CurrentDb()

    Dim qdf As DAO.QueryDef
    Set qdf = db.CreateQueryDef("")
    qdf.SQL = "PARAMETERS prmStateCode Text(2); " & _
              "SELECT " & strField & " FROM tlkpStates " & _
              "WHERE StateCode = prmStateCode"
    qdf.Parameters("prmStateCode") = strStateCode
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)

    If Not rs.EOF Then
        IsStateFeatureAvailable = rs.Fields(strField)
    Else
        IsStateFeatureAvailable = False
    End If

    rs.Close

Exit_Function:
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    IsStateFeatureAvailable = False
    Resume Exit_Function
End Function

'---------------------------------------------------------------------------------------
' Name       : ValidateTicketPlayOptions
' Purpose    : Validate that Power Play and Double Play selections are allowed
'              for the current state of play
' Parameters : blnIsPowerPlay (Boolean) - Whether Power Play is selected
'              blnIsDoublePlay (Boolean) - Whether Double Play is selected
'              strErrorMsg (String) - ByRef, receives error description if invalid
' Returns    : Boolean - True if selections are valid
'---------------------------------------------------------------------------------------
Public Function ValidateTicketPlayOptions(ByVal blnIsPowerPlay As Boolean, _
                                           ByVal blnIsDoublePlay As Boolean, _
                                           ByRef strErrorMsg As String) As Boolean
    strErrorMsg = ""

    If blnIsPowerPlay And Not gblnHasPowerPlay Then
        strErrorMsg = "Power Play is not available in " & gstrStateOfPlay & "."
        ValidateTicketPlayOptions = False
        Exit Function
    End If

    If blnIsDoublePlay And Not gblnHasDoublePlay Then
        strErrorMsg = "Double Play is not available in " & gstrStateOfPlay & "."
        ValidateTicketPlayOptions = False
        Exit Function
    End If

    ValidateTicketPlayOptions = True
End Function
