# Norris Powerball Pool

A professional-grade Microsoft Access application (.accdb) for managing Powerball lottery pool group entries, matching results, and tracking winnings. Designed for use in all 50 US states plus DC with configurable state-of-play settings.

## Requirements

- Microsoft Access 2016+ or Microsoft 365
- Single `.accdb` file (no external dependencies)
- DAO data access (Microsoft Office xx.0 Access Database Engine Object Library)

## Powerball Rules Reference

| Rule | Value |
|------|-------|
| White balls | Pick 5 from 1–69 (order does not matter) |
| Powerball | Pick 1 from 1–26 |
| Draw schedule | Monday, Wednesday, Saturday |
| Prize tiers | 9 total (0+PB through 5+PB) |

---

## Database Setup Instructions

Open Microsoft Access → **Blank desktop database** → name it `NorrisPowerballPool.accdb` → **Create**.

For every table below, go to **Create** tab → **Table Design**, add the listed fields, set the primary key, then save with the indicated name. Configure each field's properties exactly as shown — empty cells mean "leave at Access default."

---

### 1. Lookup / Reference Tables

#### `tlkpStates`

Stores all 50 US states plus DC. Seed this table with data after creation.

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `StateCode` | Short Text | **Primary Key.** Two-letter state/territory code | 2 | | | >LL | State Code | | | | Yes | Yes (No Duplicates) |
| `StateName` | Short Text | Full state or territory name | 50 | | | | State Name | | | | Yes | No |
| `FederalTaxRate` | Number | Federal tax withholding rate as a decimal | Double | Percent | 2 | | Federal Tax Rate | 0.24 | >=0 And <=1 | Federal tax rate must be between 0% and 100%. | Yes | No |
| `StateTaxRate` | Number | State tax withholding rate as a decimal | Double | Percent | 4 | | State Tax Rate | 0 | >=0 And <=1 | State tax rate must be between 0% and 100%. | Yes | No |
| `HasStateLottery` | Yes/No | Whether the state participates in Powerball | | Yes/No | | | Has State Lottery | No | | | Yes | No |
| `HasPowerPlay` | Yes/No | Whether Power Play is available in this state | | Yes/No | | | Has Power Play | No | | | Yes | No |
| `HasDoublePlay` | Yes/No | Whether Double Play is available in this state | | Yes/No | | | Has Double Play | No | | | Yes | No |

**Seed data for `tlkpStates`:**

| StateCode | StateName | FederalTaxRate | StateTaxRate | HasStateLottery | HasPowerPlay | HasDoublePlay |
|---|---|---|---|---|---|---|
| AL | Alabama | 0.24 | 0.00 | No | No | No |
| AK | Alaska | 0.24 | 0.00 | No | No | No |
| AZ | Arizona | 0.24 | 0.05 | Yes | Yes | No |
| AR | Arkansas | 0.24 | 0.055 | Yes | Yes | No |
| CA | California | 0.24 | 0.00 | Yes | Yes | No |
| CO | Colorado | 0.24 | 0.04 | Yes | Yes | Yes |
| CT | Connecticut | 0.24 | 0.0699 | Yes | Yes | No |
| DE | Delaware | 0.24 | 0.00 | Yes | Yes | No |
| DC | District of Columbia | 0.24 | 0.0875 | Yes | Yes | No |
| FL | Florida | 0.24 | 0.00 | Yes | Yes | Yes |
| GA | Georgia | 0.24 | 0.055 | Yes | Yes | No |
| HI | Hawaii | 0.24 | 0.00 | No | No | No |
| ID | Idaho | 0.24 | 0.058 | Yes | Yes | No |
| IL | Illinois | 0.24 | 0.0495 | Yes | Yes | No |
| IN | Indiana | 0.24 | 0.0323 | Yes | Yes | Yes |
| IA | Iowa | 0.24 | 0.06 | Yes | Yes | No |
| KS | Kansas | 0.24 | 0.05 | Yes | Yes | No |
| KY | Kentucky | 0.24 | 0.05 | Yes | Yes | No |
| LA | Louisiana | 0.24 | 0.05 | Yes | Yes | No |
| ME | Maine | 0.24 | 0.05 | Yes | Yes | No |
| MD | Maryland | 0.24 | 0.0875 | Yes | Yes | No |
| MA | Massachusetts | 0.24 | 0.05 | Yes | Yes | No |
| MI | Michigan | 0.24 | 0.0425 | Yes | Yes | No |
| MN | Minnesota | 0.24 | 0.0785 | Yes | Yes | No |
| MS | Mississippi | 0.24 | 0.05 | No | No | No |
| MO | Missouri | 0.24 | 0.0495 | Yes | Yes | Yes |
| MT | Montana | 0.24 | 0.069 | Yes | Yes | No |
| NE | Nebraska | 0.24 | 0.0684 | Yes | Yes | No |
| NV | Nevada | 0.24 | 0.00 | No | No | No |
| NH | New Hampshire | 0.24 | 0.00 | Yes | Yes | No |
| NJ | New Jersey | 0.24 | 0.08 | Yes | Yes | No |
| NM | New Mexico | 0.24 | 0.059 | Yes | Yes | No |
| NY | New York | 0.24 | 0.0882 | Yes | Yes | No |
| NC | North Carolina | 0.24 | 0.0525 | Yes | Yes | No |
| ND | North Dakota | 0.24 | 0.029 | Yes | Yes | No |
| OH | Ohio | 0.24 | 0.04 | Yes | Yes | No |
| OK | Oklahoma | 0.24 | 0.0475 | Yes | Yes | No |
| OR | Oregon | 0.24 | 0.09 | Yes | Yes | No |
| PA | Pennsylvania | 0.24 | 0.0307 | Yes | Yes | Yes |
| RI | Rhode Island | 0.24 | 0.0599 | Yes | Yes | No |
| SC | South Carolina | 0.24 | 0.07 | Yes | Yes | No |
| SD | South Dakota | 0.24 | 0.00 | Yes | Yes | No |
| TN | Tennessee | 0.24 | 0.00 | Yes | Yes | Yes |
| TX | Texas | 0.24 | 0.00 | Yes | Yes | Yes |
| UT | Utah | 0.24 | 0.00 | No | No | No |
| VT | Vermont | 0.24 | 0.06 | Yes | Yes | No |
| VA | Virginia | 0.24 | 0.04 | Yes | Yes | No |
| WA | Washington | 0.24 | 0.00 | Yes | Yes | No |
| WV | West Virginia | 0.24 | 0.065 | Yes | Yes | No |
| WI | Wisconsin | 0.24 | 0.0765 | Yes | Yes | No |
| WY | Wyoming | 0.24 | 0.00 | Yes | Yes | No |

> **Notes:** Double Play is available in a limited number of states. State tax rates reflect general lottery withholding and may vary by prize amount — these are user-editable. States without a lottery (AL, AK, HI, MS, NV, UT) have all play options set to No.

#### `tlkpPrizeTiers`

Defines the 9 Powerball prize tiers. Seed this table with data after creation.

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `PrizeTierID` | AutoNumber | **Primary Key.** Auto-generated tier identifier | Long Integer | | | | Prize Tier ID | | | | Yes | Yes (No Duplicates) |
| `WhiteBallMatches` | Number | Number of white balls matched (0–5) | Integer | | | | White Ball Matches | | >=0 And <=5 | White ball matches must be between 0 and 5. | Yes | No |
| `PowerballMatch` | Yes/No | Whether the Powerball was also matched | | Yes/No | | | Powerball Match | No | | | Yes | No |
| `PrizeName` | Short Text | Display name (e.g., "Jackpot", "Match 4+PB") | 50 | | | | Prize Name | | | | Yes | No |
| `DefaultPrizeAmount` | Currency | Default fixed prize amount ($0 for jackpot) | | Currency | 2 | | Default Prize Amount | 0 | >=0 | Default prize amount cannot be negative. | Yes | No |

**Seed data for `tlkpPrizeTiers`:**

| WhiteBallMatches | PowerballMatch | PrizeName | DefaultPrizeAmount |
|---|---|---|---|
| 5 | Yes | Jackpot (5+PB) | $0.00 |
| 5 | No | Match 5 | $1,000,000.00 |
| 4 | Yes | Match 4+PB | $50,000.00 |
| 4 | No | Match 4 | $100.00 |
| 3 | Yes | Match 3+PB | $100.00 |
| 3 | No | Match 3 | $7.00 |
| 2 | Yes | Match 2+PB | $7.00 |
| 1 | Yes | Match 1+PB | $4.00 |
| 0 | Yes | Match PB Only | $4.00 |

#### `tlkpAppVersion`

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `VersionID` | AutoNumber | **Primary Key.** Auto-generated version identifier | Long Integer | | | | Version ID | | | | Yes | Yes (No Duplicates) |
| `VersionNumber` | Short Text | Semantic version string (e.g., "1.0.0") | 20 | | | | Version Number | | | | Yes | No |
| `ReleaseDate` | Date/Time | Date this version was released | | Short Date | | 99/99/0000;0;_ | Release Date | | | | Yes | No |

---

### 2. System Settings Table

#### `tblSystemSettings`

Single-row table that stores global configuration. Loaded into a public variable on startup.

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `SettingsID` | AutoNumber | **Primary Key.** Auto-generated settings identifier | Long Integer | | | | Settings ID | | | | Yes | Yes (No Duplicates) |
| `PoolName` | Short Text | Name of the lottery pool | 100 | | | | Pool Name | | | | Yes | No |
| `AdminName` | Short Text | Pool administrator's name | 100 | | | | Admin Name | | | | Yes | No |
| `StateOfPlay` | Short Text | **Foreign Key** → `tlkpStates.StateCode` | 2 | | | >LL | State of Play | | | | Yes | Yes (Duplicates OK) |

---

### 3. Core Data Tables

#### `tblParticipants`

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `ParticipantID` | AutoNumber | **Primary Key.** Auto-generated participant identifier | Long Integer | | | | Participant ID | | | | Yes | Yes (No Duplicates) |
| `FirstName` | Short Text | Participant's first name | 50 | | | | First Name | | | | Yes | No |
| `LastName` | Short Text | Participant's last name | 50 | | | | Last Name | | | | Yes | No |
| `Email` | Short Text | Participant's email address | 100 | | | | Email | | | | No | No |
| `Phone` | Short Text | Participant's phone number | 20 | | | !\(999") "000\-0000;0;_ | Phone | | | | No | No |
| `IsActive` | Yes/No | Whether this participant is currently active in the pool | | Yes/No | | | Active | Yes | | | Yes | No |

#### `tblDrawings`

Stores official Powerball draw results. **One field per ball.**

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `DrawingID` | AutoNumber | **Primary Key.** Auto-generated drawing identifier | Long Integer | | | | Drawing ID | | | | Yes | Yes (No Duplicates) |
| `DrawDate` | Date/Time | Official draw date. Must be Mon, Wed, or Sat | | Short Date | | 99/99/0000;0;_ | Draw Date | | Weekday([DrawDate]) In (2,4,7) | Draw date must be a Monday, Wednesday, or Saturday. | Yes | Yes (No Duplicates) |
| `WB1` | Number | Winning white ball 1 | Integer | | 0 | | WB 1 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB2` | Number | Winning white ball 2 | Integer | | 0 | | WB 2 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB3` | Number | Winning white ball 3 | Integer | | 0 | | WB 3 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB4` | Number | Winning white ball 4 | Integer | | 0 | | WB 4 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB5` | Number | Winning white ball 5 | Integer | | 0 | | WB 5 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `PB` | Number | Winning Powerball number | Integer | | 0 | | Powerball | | >=1 And <=26 | Powerball must be between 1 and 26. | Yes | No |
| `JackpotAmount` | Currency | Estimated or actual jackpot for this drawing | | Currency | 2 | | Jackpot Amount | 0 | >=0 | Jackpot amount cannot be negative. | No | No |
| `IsVerified` | Yes/No | Whether results have been officially confirmed | | Yes/No | | | Verified | No | | | Yes | No |

> **Additional rule:** All five white ball values (`WB1`–`WB5`) must be distinct. Enforce via VBA validation in `modLotteryLogic` before saving, since Access table-level validation cannot easily cross-reference five fields for uniqueness.

#### `tblTickets`

Stores pool ticket entries (purchased numbers). **One field per ball.**

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `TicketID` | AutoNumber | **Primary Key.** Auto-generated ticket identifier | Long Integer | | | | Ticket ID | | | | Yes | Yes (No Duplicates) |
| `DrawingID` | Number | **Foreign Key** → `tblDrawings.DrawingID` | Long Integer | | 0 | | Drawing ID | | | | Yes | Yes (Duplicates OK) |
| `WB1` | Number | White ball 1 | Integer | | 0 | | WB 1 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB2` | Number | White ball 2 | Integer | | 0 | | WB 2 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB3` | Number | White ball 3 | Integer | | 0 | | WB 3 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB4` | Number | White ball 4 | Integer | | 0 | | WB 4 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `WB5` | Number | White ball 5 | Integer | | 0 | | WB 5 | | >=1 And <=69 | White ball must be between 1 and 69. | Yes | No |
| `PB` | Number | Powerball | Integer | | 0 | | Powerball | | >=1 And <=26 | Powerball must be between 1 and 26. | Yes | No |
| `IsPowerPlay` | Yes/No | Whether this ticket includes Power Play | | Yes/No | | | Power Play | No | | | Yes | No |
| `IsDoublePlay` | Yes/No | Whether this ticket includes Double Play | | Yes/No | | | Double Play | No | | | Yes | No |

> **Additional rule:** All five white ball values (`WB1`–`WB5`) must be distinct. Enforce via VBA validation in `modLotteryLogic` before saving.

#### `tblContributions`

Tracks participant payments per drawing.

| Field Name | Data Type | Description | Field Size | Format | Decimal Places | Input Mask | Caption | Default Value | Validation Rule | Validation Text | Required | Indexed |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `ContributionID` | AutoNumber | **Primary Key.** Auto-generated contribution identifier | Long Integer | | | | Contribution ID | | | | Yes | Yes (No Duplicates) |
| `ParticipantID` | Number | **Foreign Key** → `tblParticipants.ParticipantID` | Long Integer | | 0 | | Participant ID | | | | Yes | Yes (Duplicates OK) |
| `DrawingID` | Number | **Foreign Key** → `tblDrawings.DrawingID` | Long Integer | | 0 | | Drawing ID | | | | Yes | Yes (Duplicates OK) |
| `AmountPaid` | Currency | Amount contributed by this participant | | Currency | 2 | | Amount Paid | | >0 | Amount paid must be greater than zero. | Yes | No |
| `DatePaid` | Date/Time | Date payment was received | | Short Date | | 99/99/0000;0;_ | Date Paid | =Date() | | | Yes | No |

---

### 4. Define Relationships

Go to **Database Tools** → **Relationships** and create the following:

| Parent Table | Parent Field | Child Table | Child Field | Enforce RI | Cascade Update |
|---|---|---|---|---|---|
| `tlkpStates` | `StateCode` | `tblSystemSettings` | `StateOfPlay` | Yes | Yes |
| `tblDrawings` | `DrawingID` | `tblTickets` | `DrawingID` | Yes | Yes |
| `tblDrawings` | `DrawingID` | `tblContributions` | `DrawingID` | Yes | Yes |
| `tblParticipants` | `ParticipantID` | `tblContributions` | `ParticipantID` | Yes | Yes |

---

### 5. Startup Configuration

Create an **AutoExec macro** (`mcrAutoExec`) that calls `modStartup.InitializeApp` on database open. This procedure loads `tblSystemSettings` values into public variables for use throughout the application.

---

### 6. MVP Forms

| Form Name | Purpose |
|---|---|
| `frmMainDashboard` | Central navigation with command buttons |
| `frmSettings` | Edit pool name, admin name, state of play |
| `frmParticipants` | Add/edit/remove pool members |
| `frmTicketEntry` | Record purchased ticket numbers for a drawing |
| `frmDrawResults` | Enter official winning numbers for a draw date |
| `frmMatchResults` | View match-checking results by drawing |

---

### 7. MVP Queries

| Query Name | Purpose |
|---|---|
| `qryMatchCheck` | Compare `tblTickets` entries against `tblDrawings` results for a given `DrawingID`. Count matching white balls (unordered set comparison) and check Powerball exact match. |
| `qryWinningTickets` | Filter `qryMatchCheck` results to only rows with at least one prize-tier match (0+PB or better). |
| `qryUnpaidParticipants` | Find active participants with no contribution record for a given `DrawingID`. |
| `qryTicketsByDrawing` | List all tickets for a selected drawing. |

---

### 8. Naming Conventions

| Object Type | Prefix | Example |
|---|---|---|
| Tables | `tbl` | `tblParticipants` |
| Lookup Tables | `tlkp` | `tlkpStates` |
| Queries | `qry` | `qryWinningTickets` |
| Forms | `frm` | `frmMainDashboard` |
| Reports | `rpt` | `rptTicketLog` |
| Standard Modules | `mod` | `modLotteryLogic` |
| Class Modules | `cls` | `clsTicketValidator` |
| Macros | `mcr` | `mcrAutoExec` |

---

### 9. Coding Standards

- **VBA only.** All modules require `Option Explicit`.
- **DAO** for all recordset operations (not ADO).
- **Error handling** in every procedure: `On Error GoTo ErrorHandler` with `MsgBox` title `"Norris Powerball Pool"`.
- **No magic numbers.** Use constants (e.g., `Const MAX_WHITE_BALLS As Integer = 5`).
- **No hard-coded file paths.** Use `CurrentProject.Path`.
- **No `DoCmd.RunSQL`.** Use `CurrentDb.Execute` with `dbFailOnError`.
- **Parameterized queries** for any user-supplied values (no SQL string concatenation).
- **CamelCase** naming for all objects — no spaces or special characters.

---

### 10. Match-Checking Logic (Overview)

The core matching algorithm compares a ticket's five white balls against the drawing's five white balls as **unordered sets** and checks for an **exact Powerball match**:

1. Count how many of the ticket's `WB1`–`WB5` values appear in the drawing's `WB1`–`WB5` values.
2. Check if the ticket's `PB` equals the drawing's `PB`.
3. Look up the resulting (white ball count, Powerball match) pair in `tlkpPrizeTiers` to determine the prize tier.

This logic lives in `modLotteryLogic`, not in form code-behind.
