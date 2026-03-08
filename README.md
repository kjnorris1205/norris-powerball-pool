# Norris Powerball Pool
A simple, robust Microsoft Access application for tracking, managing, and automating Powerball pool group entries and winnings.

## Development Guide

### Overview

A workplace lottery pool program in Microsoft Access is built using four primary database objects—

**Tables**, **Queries**, **Forms**, and **Reports**—to automate participant tracking, ticket management, and winnings distribution.

#### Data Structure (Tables) 

Tables are the foundation where all pool data is stored. A typical setup includes: 

- **Participants Table:** Stores employee names, IDs, contact information, and their current "active" status in the pool.
- **Drawings Table:** Tracks specific lottery dates, jackpot amounts, and the game type (e.g., Powerball or Mega Millions).
- **Tickets Table:** Records the actual numbers purchased for each draw. It links to the Drawings table via a "Drawing ID".
- **Contributions Table:** Tracks who paid, how much, and when. This is critical for reconciling sales and ensuring only paid participants share in a win.

#### User Interface (Forms) 

Forms provide a user-friendly way for the pool "captain" or leader to interact with the data:

- **Data Entry Form:** For adding new participants or recording weekly payments.
- **Ticket Entry Form:** A specialized interface to quickly type in purchased ticket numbers.
- **Dashboard:** A central menu to navigate between different functions of the program.

#### Processing Logic (Queries) 

Queries perform the "heavy lifting" by filtering and calculating data:

- **Payment Verification:** Identifies which participants have not yet paid for the upcoming draw.
- **Winning Checker:** Compares the purchased ticket numbers against the official winning numbers to identify "hits" or matches.
- **Payout Calculator:** Divides a winning amount by the number of active, paid participants for that specific draw. 

#### Output and Verification (Reports) 

Reports are used to share information with the rest of the pool to maintain transparency:

- **Participant List:** A printable summary of everyone currently in the pool.
- **Ticket Log:** A report showing all ticket numbers for the next drawing, often distributed to members before the draw happens as a security measure.
- **Financial History:** Summarizes total funds collected and any small winnings that may be rolled over into future draws.

#### Automation (Macros & Modules) 

For more advanced programs, **VBA (Visual Basic for Applications)** or **Macros** can automate repetitive tasks, such as generating random numbers for a "Quick Pick" or sending automated email reminders to participants who haven't paid.

### Implementation

To implement a workplace lottery pool program in Microsoft Access, follow these steps to build the structure, interface, and logic based on the overview.

#### 1. Create the Database and Tables 

Start by opening Microsoft Access and selecting **Blank desktop database**. Give it a name and click **Create**.

- **Participants Table (`tblParticipants`):** Go to the **Create** tab and click **Table Design**.
    - `ParticipantID`: **AutoNumber** (Set as **Primary Key**).
    - `FirstName`, `LastName`: **Short Text**.
    - `IsActive`: **Yes/No** (To track current members).
- **Drawings Table (`tblDrawings`):**
    - `DrawingID`: **AutoNumber** (Primary Key).
    - `DrawDate`: **Date/Time**.
    - `WinningNumbers`: **Short Text** (or separate fields for each ball).
- **Contributions Table (`tblContributions`):**
    - `ContributionID`: **AutoNumber** (Primary Key).
    - `ParticipantID`: **Number** (Foreign Key linking to `tblParticipants`).
    - `DrawingID`: **Number** (Foreign Key linking to `tblDrawings`).
    - `AmountPaid`: **Currency**.

#### 2. Define Relationships 

Go to **Database Tools** > **Relationships**. Drag the `ParticipantID` from `tblParticipants` to `tblContributions`, and `DrawingID` from `tblDrawings` to `tblContributions`. Enable **Enforce Referential Integrity** to ensure data consistency.

#### 3. Build Data Entry Forms 

Use the **Form Wizard** on the **Create** tab to quickly build interfaces for your tables.

- **Participant Management Form:** A simple form to add or edit employee details.
- **Weekly Contribution Form:** Create a form based on `tblContributions` where the `ParticipantID` and `DrawingID` can be selected from dropdowns (Lookup fields).

#### 4. Create Queries for Logic 

Go to **Create** > **Query Design**.

- **Unpaid Members Query:** Link `tblParticipants` and `tblContributions`. Filter for participants where no payment record exists for a specific `DrawingID`.
- **Winning Matcher:** Use a query to compare a user-entered winning number against the numbers stored in your tickets table.

#### 5. Generate Reports for Transparency 

Use the **Report Wizard** to create structured documents.

- **Weekly Ticket Log:** A report showing all purchased numbers for the upcoming draw to share with the group.
- **Contribution Summary:** A report grouped by `DrawDate` to show total funds collected.

#### 6. Best Practices for Implementation 

- **Standardized Naming:** Avoid spaces in field names (e.g., use `FirstName` instead of `First Name`) to prevent issues with queries or code later.
- **Security:** If shared on a network, consider using the **Access Runtime** to allow others to use the program without needing a full Access license.
- **Transparency:** Use **Electronic Documents** or printed reports to provide every participant with a copy of the purchased tickets before the drawing occurs.
