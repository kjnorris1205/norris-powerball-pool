# Norris Powerball Pool

A professional-grade Microsoft Access application (.accdb) for managing Powerball lottery pool group entries, matching results, and tracking winnings. Designed for use in all 50 US states plus DC with configurable state-of-play settings.

## Requirements

- Microsoft Access 2016+ or Microsoft 365
- Single `.accdb` file (no external dependencies)
- DAO data access (Microsoft Office xx.0 Access Database Engine Object Library)

## Powerball Rules Reference

| Rule | Value |
|---|---|
| White balls | Pick 5 from 1–69 (order does not matter) |
| Powerball | Pick 1 from 1–26 |
| Draw schedule | Monday, Wednesday, Saturday |
| Prize tiers | 9 total (0+PB through 5+PB) |
| Power Play | Optional multiplier (2x–10x, excludes jackpot) — do NOT implement until asked |

## Database Creation

See [DATABASE-CREATION.md](DATABASE-CREATION.md) for complete step-by-step instructions to build the database from scratch, including table schemas, field specifications, and relationships.

---

## VBA Module Files

| File | Module Name | Purpose |
|---|---|---|
| `modCreateTables.bas` | `modCreateTables` | Create all tables and relationships via DAO |
| `modCreateSeedData.bas` | `modCreateSeedData` | Seed lookup tables with default data |
| `modCreateQueries.bas` | `modCreateQueries` | Create all MVP queries via DAO |
| `modLotteryLogic.bas` | `modLotteryLogic` | Runtime validation and match-checking logic |
| `modUIConstants.bas` | `modUIConstants` | UI color, font, and layout constants |
| `modFormEvents.bas` | `modFormEvents` | Navigation and form event handler functions |
| `modCreateForms.bas` | `modCreateForms` | Programmatic form creation via DAO |
| `modStartup.bas` | `modStartup` | App initialization and startup configuration |

---

## State Data Model

- Store all 50 states + DC in `tlkpStates` with fields: `StateCode` (PK, text 2), `StateName`, `FederalTaxRate` (Double), `StateTaxRate` (Double), `HasStateLottery` (Boolean), `HasPowerPlay` (Boolean), `HasDoublePlay` (Boolean).
- `HasPowerPlay` and `HasDoublePlay` control whether those ticket options are available for the selected state of play.
- Tax rates and state lottery participation should be updateable by the user.
- The selected state in `tblSystemSettings` drives which tax rates and play options apply.

## Forms

| Form Name | Purpose |
|---|---|
| `frmMainDashboard` | Central navigation with command buttons |
| `frmSettings` | Edit pool name, admin name, state of play |
| `frmParticipants` | Add/edit/remove pool members |
| `frmTicketEntry` | Record purchased ticket numbers for a drawing |
| `frmDrawResults` | Enter official winning numbers for a draw date |
| `frmMatchResults` | View match-checking results by drawing |

## Queries

| Query Name | Purpose |
|---|---|
| `qryMatchCheck` | Compare `tblTickets` entries against `tblDrawings` results for a given `DrawingID`. Count matching white balls (unordered set comparison) and check Powerball exact match. |
| `qryWinningTickets` | Filter `qryMatchCheck` results to only rows with at least one prize-tier match (0+PB or better). |
| `qryUnpaidParticipants` | Find active participants with no contribution record for a given `DrawingID`. |
| `qryTicketsByDrawing` | List all tickets for a selected drawing. |

## Naming Conventions

### Object Prefixes

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
| Constants | `ALL_CAPS` | `MAX_POWERBALL` |
| Enums | `e` | `ePrizeTier` |

### Variable Prefixes (Hungarian Notation)

`str` String, `int` Integer, `lng` Long, `dbl` Double, `cur` Currency, `bln` Boolean, `dt` Date, `rs` DAO.Recordset, `db` DAO.Database, `qdf` DAO.QueryDef, `frm` Form reference, `ctl` Control reference.

## Coding Standards

- **VBA only.** All modules require `Option Explicit`.
- **DAO** for all recordset operations (not ADO).
- **Error handling** in every procedure: `On Error GoTo ErrorHandler` with `MsgBox` title `"Norris Powerball Pool"`.
- **No magic numbers.** Use constants (e.g., `Const MAX_WHITE_BALLS As Integer = 5`).
- **No hard-coded file paths.** Use `CurrentProject.Path`.
- **No `DoCmd.RunSQL`.** Use `CurrentDb.Execute` with `dbFailOnError`.
- **Parameterized queries** for any user-supplied values (no SQL string concatenation).
- **CamelCase** naming for all objects — no spaces or special characters.

### Procedure Header Template

```vb
'---------------------------------------------------------------------------------------
' Name       : [Procedure/Function Name]
' Purpose    : [Brief description of logic]
' Parameters : [ParamName] ([Type]) - [Description]
' Returns    : [Type] - [Description]
'---------------------------------------------------------------------------------------
```

No author names or date stamps in headers.

### Error Handling Pattern

Every procedure follows this structure:

```vb
Public Sub ExampleProcedure()
    On Error GoTo ErrorHandler

    ' --- procedure logic here ---

Exit_Procedure:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in: ExampleProcedure" & vbCrLf & vbCrLf & _
           "Error #: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description, _
           vbCritical, "Norris Powerball Pool"
    Resume Exit_Procedure
End Sub
```

## Form & UI Standards

- **Navigation:** Use a main dashboard (`frmMainDashboard`) with command buttons — not a switchboard or navigation pane.
- **Form Design:** Use `Detail` section only where possible. Anchor controls for basic resize support.
- **Consistent Look:** Use a shared color scheme and font. Define colors as constants in `modUIConstants`.
- **User Feedback:** Use status bar text (`SysCmd acSysCmdSetStatus`) for non-critical info. Use `MsgBox` only for errors, confirmations, and important alerts.
- **No tab controls in MVP** — keep forms flat and simple.

## Match-Checking Logic (Overview)

The core matching algorithm compares a ticket's five white balls against the drawing's five white balls as **unordered sets** and checks for an **exact Powerball match**:

1. Count how many of the ticket's `WB1`–`WB5` values appear in the drawing's `WB1`–`WB5` values.
2. Check if the ticket's `PB` equals the drawing's `PB`.
3. Look up the resulting (white ball count, Powerball match) pair in `tlkpPrizeTiers` to determine the prize tier.

This logic lives in `modLotteryLogic`, not in form code-behind.
