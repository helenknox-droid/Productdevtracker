# Productdevtracker

Google Apps Script helpers for the product development tracker.

## Own Brand upcoming deadlines report

The script in [`apps-script/Code.gs`](apps-script/Code.gs) builds a sheet called
`Own Brand - Upcoming Deadlines` from the `Own-Brand Stage & Gates` tab.

The generated report columns are:

| Column | Header | Source |
| --- | --- | --- |
| A | Reference Number | Header-matched from row 5; originally column C |
| B | Component Name | `NS NAME (Once Created)` when populated; otherwise `Brief Name` |
| C | Launch Date | `Earliest BQID Launch Date`; if it says `no linked bouquet IDs`, use `Target Launch Date (pre BQID Link)` |
| D | Status | Header-matched from row 5; originally column F |
| E | Current Stage | Header-matched from `Current Gate`; originally column D |
| F | Upcoming Gate | Derived from the deadline column header, e.g. `Gate 6 Deadline` becomes `Gate 6` |
| G | Deadline | Gate deadline headers such as `Gate 0 Deadline`, limited to the next 4 weeks |
| H | Comments | Blank comment column refreshed on every run |

Source data starts at row 6, so the script treats row 5 as the header row.
The target report starts at row 2: row 2 is the header row and returned
deadline rows begin at row 3.
The target tab is cleared and rebuilt from row 2 downward each time the report
runs, leaving row 1 untouched.
If no deadlines match the four-week window, the target tab shows a short
message and you can run the diagnostics report described below.

### Header names currently used

The script anchors on header names instead of fixed column letters. Please review
these expected names against the real row 5 source headers:

| Field | Accepted source header names |
| --- | --- |
| Reference Number | `Reference Number` |
| Current Stage | `Current Gate` |
| Status | `Status` |
| Primary Launch Date | `Earliest BQID Launch Date` |
| Fallback Launch Date | `Target Launch Date (pre BQID Link)` |
| Created Component Name | `NS NAME (Once Created)`; optional, originally column P |
| Brief Name fallback | `Brief Name`; originally column L |
| Deadline columns | Gate deadline headers matching `Gate <number> Deadline`, for example `Gate 0 Deadline` through `Gate 6 Deadline` |

Deadline values are expected in year-week format, for example `2026-W43`.
The parser also accepts common variants such as `2026-WK43`, `2026 Week 43`,
and `W43 2026`.
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
menu with:

- **Build Upcoming Deadlines**
- **Build Deadline Diagnostics**

### Troubleshooting an empty report

If the report returns no deadline rows:

1. Run **Own Brand Reports > Build Deadline Diagnostics**.
2. Check the `Own Brand - Deadline Diagnostics` tab.
3. Review:
   - how many deadline columns were found,
   - how many populated deadline cells were checked,
   - how many values parsed as year-week dates,
   - how many were inside the upcoming four-week window,
   - sample dates before the window, after the window, or not parseable.

For example, if today is in `2026-W22`, the report window is roughly
`2026-W22` through `2026-W26`. A deadline like `2026-W43` is valid, but it will
not appear because it is outside the next four weeks.

### Remaining check

Please confirm whether the header names listed above match row 5 in the source
sheet. If any differ, update the matching `headers` list in `REPORT_CONFIG`.
