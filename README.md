# Productdevtracker

Google Apps Script helpers for the product development tracker.

## Own Brand upcoming deadlines report

The script in [`apps-script/Code.gs`](apps-script/Code.gs) builds a sheet called
`Own Brand-Upcoming Deadlines` from the `Own Brand Stage and Gates` tab.

The generated report columns are:

| Column | Header | Source |
| --- | --- | --- |
| A | Reference Number | `Own Brand Stage and Gates!C`, from row 6 |
| B | Component Name | `Own Brand Stage and Gates!L` |
| C | Launch Date | `Own Brand Stage and Gates!I`, unless it says `no linked bouquet IDs`; then `J` |
| D | Status | To be confirmed in `REPORT_CONFIG.columns.status` |
| E | Current Stage | To be confirmed in `REPORT_CONFIG.columns.currentStage` |
| F | Stage | Configured per deadline in `REPORT_CONFIG.deadlineStages` |
| G | Deadline | Configured per deadline in `REPORT_CONFIG.deadlineStages` |

### Setup

1. Open the Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Paste the contents of `apps-script/Code.gs` into the Apps Script editor.
4. Update `REPORT_CONFIG`:
   - Set the `status` source column.
   - Set the `currentStage` source column.
   - Add every stage/deadline column to `deadlineStages`.
5. Run `buildOwnBrandUpcomingDeadlinesReport`.

After the spreadsheet is reloaded, the script also adds an **Own Brand Reports**
menu with a **Build Upcoming Deadlines** item.

### Questions to finish the report configuration

To complete the script, please confirm:

1. Which column in `Own Brand Stage and Gates` contains **Status**?
2. Which column contains **Current Stage**?
3. Which columns contain the stage deadline dates that should be pulled into the report?
4. For each deadline column, should the **Stage** value come from:
   - the column header,
   - a separate stage-name column in the same row, or
   - a fixed stage name that we define in the script?
5. Should the report include every populated deadline, or only future/upcoming deadlines?
6. Should rows with blank reference numbers be skipped? The current script skips them.
7. Should the target sheet be fully refreshed each time, or should it preserve old rows/formatting?
