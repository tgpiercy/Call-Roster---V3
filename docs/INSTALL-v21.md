# Call Out Roster v21 - install

Target: installed before Sat Oct 31 (the DST night). Takes about 15 minutes.

**Back up first.** Save a copy of the workbook somewhere safe before you start.

## 1. Replace the modules
In the VBA editor (Alt+F11), for EACH of these, right-click the module > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modOTLog`, `ImportKronosData`, `modShiftCycle`, `modHolidays`, `modShiftSwitch`, `modPostings`, `modCalendar`, `modUI`, `modSystem`, `Call_Out_Roster`

(`Call_Out_Roster` replaces the old sort-and-copy call out roster macro. If your project has that code under another module name, remove that module too, so there is only one `Call_Out`.)

## 2. Import the two new modules
**File > Import File:** `modules/modGradeLog.bas`, then `modules/modSelfTest.bas`.

## 3. Re-paste the code (open the code window, Ctrl+A, paste)
- **ThisWorkbook** from `paste-in/ThisWorkbook.txt`
- **frmOvertimeEntry** from `paste-in/frmOvertimeEntry.txt`
- **frmNewPosting** from `paste-in/frmNewPosting.txt`

## 4. Settings (in modOvertimeEntry, near the top)
- `BACKUP_FOLDER` - the folder for the daily backup copy, e.g. `"P:\...\Roster Backups\"`. Leave it `""` for now if the folder isn't decided; SystemCheck will keep reminding you.
- `DST_USE_ACTUAL_HOURS` is `True` (13 h / 11 h on the clock-change nights). Set it to `False` to keep 12 h and only see the warning.

## 5. Debug > Compile VBAProject
It should finish with no message. If it stops on a line, send me a screenshot.

## 6. One-time macros (Alt+F8)
1. **RefreshRosterFormulas** - rewrites Total Hours as Worked + Refused + Provisional and checks every row. Expect "Formulas refreshed for all N roster rows" and the provisional total (probably 0 until something is accepted).
2. **SetupHomeUnit** - saves a backup copy, adds the Home Unit column (B) with its drop-down, fills in everyone who has a code in exactly one unit, and lists who is left for you.
3. **ApplyRosterLayout** - refreshes the roster formatting (the new column and the hidden provisional helper).

## 7. Save, close and reopen
Opening as an admin now takes today's backup (if `BACKUP_FOLDER` is set) and builds the grade-change snapshot.

## 8. SystemCheck (Alt+F8)
Expect **"Modules: 12 of 12 on v21"**. The only CHECK lines should be:
- people with no Home Unit (fill them in from the drop-down in column B, or from Needs Attention), and
- `BACKUP_FOLDER isn't set` (until it is).

Anything else listed: send me the message.

## 9. SelfTest (Alt+F8)
Expect **"147 passed, 0 failed"**. It writes nothing and doesn't save.

## 10. Import Kronos
Run **Import Kronos** once so the status line shows "data through <date>" and the provisional hours start dropping off as Kronos catches up.

## Smoke test
1. SystemCheck: all modules on v21; only the Home Unit and BACKUP_FOLDER lines (see step 8).
2. SelfTest: all cases pass.
3. On a COPY of the file: Log Call-Out, mark someone near the top of the list Accepted and submit. Open Log Call-Out again for the same unit: they have moved down and their total shows `*`. Void the entry (Void Entry > the S- id): they move back.
4. New Posting, Shift = Nights, Shift date = Sat Oct 31: Hours changes to 13 and the red line reads "DST night: 13 h (clocks go back)". Change the date to Oct 30: it goes back to 12.
5. On the roster change one person's grade code (e.g. D2 -> D1). Alt+F8 > ShowGradeLog: one new row with the old and new code and your login. Click a sort button: no new rows appear. Put the code back (that is one more row).
6. Click **Call Out Roster**. The Blank Call Out Roster tab is redrawn (one section per unit, in call order, hours / name / crew / phone) without the OT Roster re-sorting, and you're asked whether to print it. Roster columns: A hidden, B = Home Unit.
7. Close, and reopen as an admin. If `BACKUP_FOLDER` is set, today's `Call Out Roster yyyy-mm-dd.xlsm` is in it; if the folder can't be reached you get ONE warning and the roster opens as usual.

## What changed (short)
- Accepted call-outs and awarded postings count in Total Hours straight away, until Kronos covers the shift.
- The two clock-change nights default to 13 h / 11 h on New Posting and Log Call-Out.
- Home Unit column (B) on the roster, filled automatically where it's obvious.
- Add Adjustment asks for a type.
- Grade code changes are logged (Alt+F8 > ShowGradeLog).
- A dated backup copy is saved on the first admin open each day.
- Alt+F8 > SelfTest checks the rules.
- The Call Out Roster button draws the printable roster directly from the data (no sorting or copying) and offers to print it.
