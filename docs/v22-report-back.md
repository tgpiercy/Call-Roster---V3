# v22 report-back (before any code)

Date: 2026-10-13. Baseline: tag `M1` (v21 + M1). Brief: `briefs/v22.md` (the kit's brief 03). Decisions taken on Oct 13: sick window = any shift starting within 24 h after the sick shift ends; calendar label "Sick"; WSIB and LTD added as leave types.

## What the code has today (the brief's "verify" points)

1. **Hours History** (`tblHoursHistory`, ImportKronosData.bas): Import ID, Imported At, Kronos File Date, Employee, Worked Hours, Imported By, **Data Through** (v21; blank on v20 rows, derive as file date minus one). Nothing reads it yet. One row per person per import, so "worked as of d" = the row with the latest Data Through <= d for that person.
2. **Absences** (`tblAbsence`, 13 columns, `AB_*` in modCalendar.bas:35-48): Absence ID, Employee, Unit, Type (3-letter code), Start, End (blank = open), Hours / Shift, Times, Note, Status, Logged By, Timestamp, Void Note. **No Shift, HomeUnit or LeaveGroup column.** Rows are written by `SaveAbsence` (646-698) with `ReDim rec(1 To 1, 1 To AB_N)`, so appending columns means bumping `AB_N` and every reader keeps working (all reads use the named constants). Types today: VAC, TRN, LOA, STD, PAR, OTH (`AbsenceTypes`, line 107); `TypeTag`/`TypeLong`/`TypeColor` map LOA/STD/PAR to "Leave".
3. **Booked off = date range only.** `AbsentOn` (380-389) checks Start <= d <= End, no shift. A sick call needs a shift-aware rule, so `NotFreeReason` gains one extra check for SIC rows.
4. **Imported lines** live in `tblCoverage` with `Source = "Import yyyy-mm-dd"`, `Type` from `GuessType` (VAC unless the text says PHA/train/op/S-S), `Off / Need` = the text as written ("Paul M"), `Note` = the month tab name, `Absence ID` blank (column `CV_ABS` exists, unused for imports). `CalName` turns a roster name into "First L", so matching runs the other way: first name + last initial against the roster.
5. **Needs Attention** already has a LEAVE section (returning within the window, or open-ended; modCalendar DrawNeeds section D). The "Returning from leave" list is that section plus a double-click action.
6. **Adjustments**: `AddAdjustment` writes one row (Ref `ADJ-####`, Type "Adjustment", type word in Reason, calculation text in Note). Leave return reuses the same write path with Ref `LV-####` and Reason "Leave return".
7. **Forms**: `frmTimeOff` is the code-built pattern (controls created in `UserForm_Initialize`, prefixed names, `ApplyFormTheme`). `frmOvertimeEntry` has no pre-fill entry point; Log Call-Out is opened by `ShowOvertimeForm`. The sick call needs a `Prefill(unit, date, shift)` on that form.
8. **Buttons**: the LOG group is three rows (Log Call-Out, Shift Switch, Book Time Off). Two more buttons need a home.

## Approach

### 1. `TotalHoursAsOf(emp, d, snapUsed, note)` (modOTLog or a new modHoursAsOf)
- Worked: Hours History row for emp with the latest Data Through <= d (v20 rows: file date minus 1). None: earliest row, `note` = "start figure estimated from <date>"; flag if the snapshot is more than 7 days from d.
- Plus non-void Hours Charged with Shift Date <= d (OB/MAN rows: Period Start) within the roster year.
- Plus provisional as of d: non-void Accepted/Awarded rows with Shift Date <= d whose shift had not ended by that snapshot's Data Through.
- Pure core `HoursAsOfCore(...)` over arrays so SelfTest and tests/ can check it; the sheet reader wraps it.

### 2. Leave (`frmLeave`, code-built, Going / Returning)
- Verified column add on `tblAbsence`: **HomeUnit, LeaveGroup, Shift** appended (14-16), done once by `UpgradeAbsenceColumns` (backup copy, add, verify row count and every roster total unchanged, block save on mismatch). `AB_N` becomes 16.
- Types: PAR, STD, LOA, **WSIB, LTD**, OTH all shown as "Leave" everywhere outside Absences.
- Going: writes the absence with HomeUnit copied from the roster and a new `LV-####` group id; off call lists from the first day; Needs Attention LEAVE section unchanged.
- Returning: pick from open leave absences; window = first day off to return, clipped to the year start; peers = HomeUnit's people with a code and grade <= 2, excluding the returner, anyone with a leave-type absence overlapping the window, and anyone with no Hours History before the window start; average of `TotalHoursAsOf(return) - TotalHoursAsOf(firstDayOff)`, rounded to 0.1; confirm shows unit, peers, dates, snapshot dates, average, before -> after; on OK one Adjustment row (Ref `LV-####`, Reason "Leave return", Note = the calculation), absence End = return - 1. Admin override value with a required reason.
- Void `LV-####` removes the top-up (existing Void Entry).

### 3. Sick Call (`frmSickCall`, code-built)
- Person, date, shift (from rotation), unit (Home Unit), hours (DST-aware default), note. Saves a `tblAbsence` row Type **SIC** with Shift, Ref `SC-####` in Absence ID.
- Eligibility: in `NotFreeReason`, a SIC row blocks any candidate shift that starts before the sick shift's end + 24 h. Pure rule `SickBlocks(sickDate, sickIsNight, d, isNight)` (days D sick -> off nights D and days D+1; nights D sick -> off days D+1 and nights D+1). Mirrored in tests/ and SelfTest.
- Adds an OPEN coverage line (Type SIC, text "<name> (Sick)"), `MarkCalDirty`, then opens Log Call-Out pre-filled. Calendar shows "Sick".
- A second sick call for the next shift is just another SC row.

### 4. Convert imported lines (`frmConvert`, code-built)
- Candidates: coverage rows with Source "Import*", Type VAC, date >= today, Absence ID blank, non-void.
- Match "Paul M" to the roster by first name + last initial (roster names are "Last, First"); Match / Ambiguous (pick) / No match.
- Tick rows; consecutive dates for one person merge into one VAC absence (Shift column records Days/Nights/Both per the lines); each converted line gets the new Absence ID so it is never converted twice; the line keeps its display.
- Needs Attention: **IMPORTED, NOT BOOKED** section (next 56 days) until converted or dismissed (dismiss = Absence ID "n/a").

### 5. Buttons and plumbing
- New command-bar group **PEOPLE** (admin): Sick Call, Leave, Convert imports. `ApplyRosterLayout` rebuilds it (install step).
- `frmOvertimeEntry.Prefill(unit, date, shift)`.
- SystemCheck lists the new module (`modLeave` holding TotalHoursAsOf, leave maths, sick rule, convert logic); forms get version lines.
- SelfTest + tests/: TotalHoursAsOf (snapshot choice, year clip, void excluded, provisional as of d), leave average with an on-leave peer and a grade-3 peer excluded, sick window for days and nights, convert merging a 5-day run.

### Risks
- `AB_N` grows: every `ReDim ... AB_N` already uses the constant, but SelfTest's fixture rows and the member publish's absence reads must be checked.
- Hours History before v21 has no Data Through; the derived value (file date minus 1) is the same rule.
- Leave maths depends on Hours History having started (Oct 9); anyone already on leave before that gets the "estimated from" note.
