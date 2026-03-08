# Copilot Instructions: Norris Powerball Pool (MS Access)

## Project Overview
- **Name:** Norris Powerball Pool
- **Goal:** A professional-grade MS Access database for managing Powerball lottery pools, sold on a marketplace for use in all 50 US states plus DC.
- **Key Feature:** Global settings to select a "State of Play" upfront, which dictates tax withholding rates and jurisdiction-specific rules.
- **Phasing:** Start with a working MVP, then layer in complexity. Do not over-engineer early features.

## Team & Workflow
- **Two-person team:** Kevin (logic/VBA) and Daniel (UI/form design).
- **AI role:** Generate clean, working code that either team member can understand. Favor clarity over cleverness.
- **Iteration style:** Build features end-to-end in small vertical slices. Each slice should be testable on its own before moving to the next.

## MVP Scope (Phase 1)
Focus only on these capabilities first:
1. **Global Settings** — Select state of play, pool name, admin name.
2. **Participants** — Add/edit/remove pool members.
3. **Pool Entries** — Record ticket purchases (white balls 1-69, Powerball 1-26) tied to a draw date. Allow for double play and power play options per state rules.
4. **Draw Results** — Enter official winning numbers for a draw date.
5. **Match Checking** — Compare entries against results and flag winners by prize tier.
6. **Dashboard** — Simple main navigation form.

Do NOT generate code for these until asked: tax calculations, payout splitting, reporting, Power Play logic, state law lookups, import/export, or multi-pool management.

## Environment & Platform
- **Platform:** Microsoft Access (.accdb format), targeting Access 2016+/Microsoft 365.
- **Data Access:** Use **DAO** (not ADO) for all recordset operations. Reference: `Microsoft Office xx.0 Access Database Engine Object Library`.
- **Distribution Model:** Single-file `.accdb` for MVP. Plan for front-end/back-end split (FE/BE) in a later phase.
- **No external dependencies** in MVP — no linked tables, no ODBC, no COM add-ins.
- **Future Extensions:** C# or JavaScript integrations will come later. Keep VBA modules self-contained so they can be wrapped or called externally.

## General Coding Standards
- **Language:** VBA (Visual Basic for Applications).
- **Naming Convention:** Use **CamelCase** for all objects. No spaces or special characters in names (e.g., `tblDrawResults` not `Draw Results`).
- **Every module** must include `Option Explicit` at the top.
- **Every procedure** must have `On Error GoTo ErrorHandler` with a clean exit path.
- **Variable names** must be self-documenting (e.g., `strStateCode` not `s`, `dblTaxRate` not `t`).
- **No magic numbers.** Use named constants or an Enum (e.g., `Const MAX_WHITE_BALLS As Integer = 5`).
- **No hard-coded file paths.** Use `CurrentProject.Path` for any file references.

## Object Naming Prefixes
| Object Type     | Prefix     | Example                  |
|-----------------|------------|--------------------------|
| Tables          | `tbl`      | `tblParticipants`        |
| Queries         | `qry`      | `qryWinningTickets`      |
| Forms           | `frm`      | `frmMainDashboard`       |
| Reports         | `rpt`      | `rptPoolPayouts`         |
| Standard Modules| `mod`      | `modLotteryLogic`        |
| Class Modules   | `cls`      | `clsTicketValidator`     |
| Macros          | `mcr`      | `mcrAutoExec`            |
| Constants       | `ALL_CAPS` | `MAX_POWERBALL`          |
| Enums           | `e`        | `ePrizeTier`             |

### Variable Prefixes (Hungarian Notation)
`str` String, `int` Integer, `lng` Long, `dbl` Double, `cur` Currency, `bln` Boolean, `dt` Date, `rs` DAO.Recordset, `db` DAO.Database, `qdf` DAO.QueryDef, `frm` Form reference, `ctl` Control reference.

## Database Architecture
- **Global Settings:** Single-row table `tblSystemSettings` stores state of play, app config. Load into a public variable on startup via `AutoExec` macro → `modStartup.InitializeApp`.
- **Separation of Concerns:** Form code-behind handles UI events only. All business logic (matching, validation, calculations) lives in standard modules (`mod` prefix).
- **Data Integrity:** Every table must have a primary key (AutoNumber `ID` or natural key). Enforce referential integrity with cascading updates where appropriate.
- **Lookup/Reference Tables:** Use `tlkp` prefix for lookup tables (e.g., `tlkpStates`, `tlkpPrizeTiers`). Seed these with data — do not rely on user entry for reference data.

## Powerball Domain Rules
These are fixed rules the AI must respect when generating lottery logic:
- **White balls:** Pick 5 from 1–69. Order does not matter for matching.
- **Powerball:** Pick 1 from 1–26.
- **Prize Tiers (9 total):** Match 0+PB through 5+PB. The jackpot is 5 white + Powerball.
- **Power Play:** Optional multiplier (2x–10x, excludes jackpot). Do NOT implement until asked.
- **Draw Schedule:** Monday, Wednesday, and Saturday nights.
- **Matching Logic:** Compare entry white balls to draw white balls as unordered sets. Powerball is an exact match. Count matching white balls + whether Powerball matches to determine prize tier.

## State Data Model
- Store all 50 states + DC in `tlkpStates` with fields: `StateCode` (PK, text 2), `StateName`, `FederalTaxRate` (Double), `StateTaxRate` (Double), `HasStateLottery` (Boolean), `HasPowerPlay` (Boolean), `HasDoublePlay` (Boolean).
- `HasPowerPlay` and `HasDoublePlay` control whether those ticket options are available for the selected state of play.
- Tax rates and state lottery participation should be updateable by the user.
- The selected state in `tblSystemSettings` drives which tax rates and play options apply.

## Form & UI Standards
- **Navigation:** Use a main dashboard (`frmMainDashboard`) with command buttons — not a switchboard or navigation pane.
- **Form Design:** Use `Detail` section only where possible. Anchor controls for basic resize support.
- **Consistent Look:** Use a shared color scheme and font. Define colors as constants in `modUIConstants`.
- **User Feedback:** Use status bar text (`SysCmd acSysCmdSetStatus`) for non-critical info. Use `MsgBox` only for errors, confirmations, and important alerts.
- **No tab controls in MVP** — keep forms flat and simple.

## Documentation & Error Handling

### Procedure Header Template
```VB
'---------------------------------------------------------------------------------------
' Name       : [Procedure/Function Name]
' Purpose    : [Brief description of logic]
' Parameters : [ParamName] ([Type]) - [Description]
' Returns    : [Type] - [Description]
'---------------------------------------------------------------------------------------
```
No author names or date stamps in headers.

### Error Handling Pattern
Every procedure follows this structure. Use "Norris Powerball Pool" as the MsgBox title everywhere.

```VB
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

## Security & Distribution
- **No sensitive data in code** — no API keys, passwords, or connection strings embedded in modules.
- **Input validation:** Validate all user inputs on forms before writing to tables (e.g., ball numbers in range, required fields not empty).
- **SQL injection prevention:** Use parameterized queries (`QueryDef.Parameters`) instead of concatenating user input into SQL strings.
- **Marketplace distribution:** The `.accdb` file must be self-contained. Include a `tlkp` table with app version info. Provide a "Reset to Defaults" option in settings.

## What NOT to Generate
- Do not generate test data unless explicitly asked.
- Do not add features beyond what is requested in the current task.
- Do not refactor or "improve" existing working code unless asked.
- Do not create external files (`.txt`, `.csv`, `.xml`) unless the task requires it.
- Do not use `DoCmd.RunSQL` for action queries — use `CurrentDb.Execute` with `dbFailOnError` instead.