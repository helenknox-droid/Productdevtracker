# Productdevtracker

Google Apps Script helpers for the product development tracker.

## Own Brand upcoming deadlines report

The script in [`apps-script/Code.gs`](apps-script/Code.gs) builds a sheet called
`Own Brand - Upcoming Deadlines` from the `Own-Brand Stage & Gates` tab.

The generated report columns are:

| Column | Header | Source |
| --- | --- | --- |
| A | Reference Number | Header-matched from row 5; originally column C |
| B | Component Name | Header-matched from row 5; originally column L |
| C | Launch Date | Header-matched primary launch date; if it says `no linked bouquet IDs`, use the fallback launch date |
| D | Status | Header-matched from row 5; originally column F |
| E | Current Stage | Header-matched from row 5; originally column D |
| F | Stage | Derived from the deadline column header, e.g. `Gate 6 Deadline` becomes `Gate 6` |
| G | Deadline | Any header containing `deadline`, limited to the next 4 weeks |
| H | Comments | Blank comment column refreshed on every run |

Source data starts at row 6, so the script treats row 5 as the header row.
The target tab is cleared and rebuilt each time the report runs.

### Header names currently used

The script anchors on header names instead of fixed column letters. Please review
these expected names against the real row 5 source headers:

| Field | Accepted source header names |
| --- | --- |
| Reference Number | `Reference Number`, `Reference No`, `Ref Number`, `Ref No` |
| Current Stage | `Current Stage` |
| Status | `Status` |
| Primary Launch Date | `Launch Date`, `Launch Week` |
| Fallback Launch Date | `Fallback Launch Date`, `Launch Date Fallback`, `Manual Launch Date` |
| Component Name | `Brief Name`, `Component Name` |
| Deadline columns | Any header containing `deadline` |

Deadline values are expected in year-week format, for example `2026-W43`.
The report includes deadlines from the current ISO week through 4 weeks ahead,
inclusive.

### Setup

1. Open the Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Paste the contents of `apps-script/Code.gs` into the Apps Script editor.
4. Review `REPORT_CONFIG.columns` and adjust the accepted header names if the
   real sheet uses different labels.
5. Run `buildOwnBrandUpcomingDeadlinesReport`.

After the spreadsheet is reloaded, the script also adds an **Own Brand Reports**
menu with a **Build Upcoming Deadlines** item.

### Remaining check

Please confirm whether the header names listed above match row 5 in the source
sheet. If any differ, update the matching `headers` list in `REPORT_CONFIG`.
