# Database Seeding

## **SQL: Seed `tlkpStates`**

Below are ready-to-run Access SQL `INSERT` statements that seed `tlkpStates` with the 50 US states plus DC using the default values shown earlier. Review and edit the tax rates and play-option flags as needed for your distribution.

```sql
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AL','Alabama',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AK','Alaska',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AZ','Arizona',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AR','Arkansas',0.24,0.055,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('CA','California',0.24,0.00,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('CO','Colorado',0.24,0.04,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('CT','Connecticut',0.24,0.0699,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('DE','Delaware',0.24,0.00,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('DC','District of Columbia',0.24,0.0875,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('FL','Florida',0.24,0.00,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('GA','Georgia',0.24,0.055,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('HI','Hawaii',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('ID','Idaho',0.24,0.058,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('IL','Illinois',0.24,0.0495,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('IN','Indiana',0.24,0.0323,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('IA','Iowa',0.24,0.06,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('KS','Kansas',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('KY','Kentucky',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('LA','Louisiana',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('ME','Maine',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MD','Maryland',0.24,0.0875,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MA','Massachusetts',0.24,0.05,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MI','Michigan',0.24,0.0425,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MN','Minnesota',0.24,0.0785,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MS','Mississippi',0.24,0.05,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MO','Missouri',0.24,0.0495,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('MT','Montana',0.24,0.069,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NE','Nebraska',0.24,0.0684,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NV','Nevada',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NH','New Hampshire',0.24,0.00,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NJ','New Jersey',0.24,0.08,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NM','New Mexico',0.24,0.059,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NY','New York',0.24,0.0882,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('NC','North Carolina',0.24,0.0525,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('ND','North Dakota',0.24,0.029,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('OH','Ohio',0.24,0.04,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('OK','Oklahoma',0.24,0.0475,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('OR','Oregon',0.24,0.09,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('PA','Pennsylvania',0.24,0.0307,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('RI','Rhode Island',0.24,0.0599,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('SC','South Carolina',0.24,0.07,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('SD','South Dakota',0.24,0.00,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('TN','Tennessee',0.24,0.00,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('TX','Texas',0.24,0.00,TRUE,TRUE,TRUE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('UT','Utah',0.24,0.00,FALSE,FALSE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('VT','Vermont',0.24,0.06,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('VA','Virginia',0.24,0.04,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('WA','Washington',0.24,0.00,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('WV','West Virginia',0.24,0.065,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('WI','Wisconsin',0.24,0.0765,TRUE,TRUE,FALSE);
INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('WY','Wyoming',0.24,0.00,TRUE,TRUE,FALSE);
```

How to run these statements in Access:

1. Open `NorrisPowerballPool.accdb` in Microsoft Access.
2. Create the `tlkpStates` table (see schema above) if you haven't already.
3. In the **Create** tab, click **Query Design** → close the Add Tables dialog.
4. Switch the query to **SQL View** (View → SQL View) and paste the SQL block above.
5. Click the red exclamation mark (Run) to execute the INSERTs. Confirm that 51 records were added.

Alternative: Run from VBA (single-run example). Open the VBA editor (Alt+F11), insert a new module, and run:

```vb
'---------------------------------------------------------------------------------------
' Name       : SeedStates
' Purpose    : Insert default rows into tlkpStates
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub SeedStates()
	On Error GoTo ErrorHandler

	Dim db As DAO.Database
	Set db = CurrentDb()

	db.Execute "INSERT INTO tlkpStates (StateCode, StateName, FederalTaxRate, StateTaxRate, HasStateLottery, HasPowerPlay, HasDoublePlay) VALUES ('AL','Alabama',0.24,0.00,FALSE,FALSE,FALSE);", dbFailOnError
	' ...repeat for each INSERT or build and execute a single SQL string

Exit_Procedure:
	Exit Sub

ErrorHandler:
	MsgBox "An error occurred in: SeedStates" & vbCrLf & vbCrLf & _
		   "Error #: " & Err.Number & vbCrLf & _
		   "Description: " & Err.Description, _
		   vbCritical, "Norris Powerball Pool"
	Resume Exit_Procedure
End Sub
```

Notes:
- Verify `tlkpStates.StateCode` is the primary key before inserting to avoid duplicates.
- If you re-run the SQL, remove or change existing rows first (e.g., `DELETE FROM tlkpStates;`).
- Review and adjust `FederalTaxRate` and `StateTaxRate` values to match your deployment requirements.

## **SQL: Seed `tlkpPrizeTiers`**

Use the following Access SQL `INSERT` statements to seed `tlkpPrizeTiers` with the nine Powerball prize tiers described above. Omit `PrizeTierID` (AutoNumber) from the INSERTs.

```sql
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (5, TRUE, 'Jackpot (5+PB)', 0.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (5, FALSE, 'Match 5', 1000000.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (4, TRUE, 'Match 4+PB', 50000.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (4, FALSE, 'Match 4', 100.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (3, TRUE, 'Match 3+PB', 100.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (3, FALSE, 'Match 3', 7.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (2, TRUE, 'Match 2+PB', 7.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (1, TRUE, 'Match 1+PB', 4.00);
INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (0, TRUE, 'Match PB Only', 4.00);
```

How to run these statements in Access:

1. Open `NorrisPowerballPool.accdb` in Microsoft Access.
2. Create the `tlkpPrizeTiers` table (see schema above) if you haven't already.
3. In the **Create** tab, click **Query Design** → close the Add Tables dialog.
4. Switch the query to **SQL View** (View → SQL View) and paste the SQL block above.
5. Click the red exclamation mark (Run) to execute the INSERTs. Confirm that 9 records were added.

Alternative: Run from VBA (single-run example). Open the VBA editor (Alt+F11), insert a new module, and run:

```vb
'---------------------------------------------------------------------------------------
' Name       : SeedPrizeTiers
' Purpose    : Insert default rows into tlkpPrizeTiers
' Parameters : None
' Returns    : None
'---------------------------------------------------------------------------------------
Public Sub SeedPrizeTiers()
	On Error GoTo ErrorHandler

	Dim db As DAO.Database
	Set db = CurrentDb()

	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (5, TRUE, 'Jackpot (5+PB)', 0.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (5, FALSE, 'Match 5', 1000000.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (4, TRUE, 'Match 4+PB', 50000.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (4, FALSE, 'Match 4', 100.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (3, TRUE, 'Match 3+PB', 100.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (3, FALSE, 'Match 3', 7.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (2, TRUE, 'Match 2+PB', 7.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (1, TRUE, 'Match 1+PB', 4.00);", dbFailOnError
	db.Execute "INSERT INTO tlkpPrizeTiers (WhiteBallMatches, PowerballMatch, PrizeName, DefaultPrizeAmount) VALUES (0, TRUE, 'Match PB Only', 4.00);", dbFailOnError

Exit_Procedure:
	Exit Sub

ErrorHandler:
	MsgBox "An error occurred in: SeedPrizeTiers" & vbCrLf & vbCrLf & _
		   "Error #: " & Err.Number & vbCrLf & _
		   "Description: " & Err.Description, _
		   vbCritical, "Norris Powerball Pool"
	Resume Exit_Procedure
End Sub
```

Notes:
- Verify `tlkpPrizeTiers` exists and has the expected fields before inserting.
- If you re-run the SQL, remove or change existing rows first (e.g., `DELETE FROM tlkpPrizeTiers;`).
- Adjust `DefaultPrizeAmount` values if you intend to use custom default amounts.