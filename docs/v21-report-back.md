# v21 report-back (before any code)

Date: 2026-10-11. Baseline: tag `v20` (commit 1969609). Repo: `C:\Users\Graeme\Claude.code\callout-roster`.

## Step 0 result

- v20 master unzipped: 17 modules, 6 paste-ins, UPDATE-20.md (now `docs/UPDATE-20.md`). All files are CRLF, ASCII, `Option Explicit`, and every `.bas` starts with `Attribute VB_Name`. Committed and tagged `v20`.
- **Gate: all three Oct 9 fixes are present**, so no `v20.1` is needed.
  - Out-of-department drop-down has an hours column, lowest first: `OutsideCandidates` (modPostings.bas:1153-1189) sorts by hours then seniority; `cmbOutside` is 4 columns (frmPostingEntry.txt:152-163).
  - Printed blank lines ask for name and initials only: modPostings.bas:606 ("Out of department - print your name" / "Initials to sign up"); no phone or date column.
  - Two-way swap payback blocked unless the payer is free: `PaybackProblem` (frmShiftSwitch.txt:469-498) calls `NotFreeReason` and Save stops on it (lines 665-669), then `EligibilityProblems` (line 689).
- **Still missing:** `deep-dive-review-2026-10-09.md` and `transition-plan-2026-10.md` were not in the zip. The brief carries every decided rule, so they are not blocking, but `docs/` is incomplete until they arrive.

## Answers to the six questions

### 1. Roster columns: fixed letters, not header lookup
Every roster column is a constant in modOvertimeEntry.bas:48-62: D name, E shift, F phone, **G total, H worked, I refused, J seniority, K email**, T:AS pay periods, AZ right edge of sorts. Unit codes L:S are a fixed `Select Case` in `DeptCol` (lines 174-186): L QC Derivatives, M Derivatives, N Cypure, O Phosphine, P Material Handler, Q QC Gases, R Utilities, S SBP. The only header-text lookups are: `RosterSheet()` finds the *sheet* by "EMPLOYEE" in D2 and "SENIORITY" in J2 (lines 371-389), and `SetupEmailColumn` checks K2 = "EMAIL". Nothing finds a roster column by header, so v21 follows the constant pattern.

Spare columns: A and B were cleared by the migration (`CleanupRoster`, modMigration.bas:178) and are never read or written since; C is hidden by `FormatRoster` as a "duplicate of Total Hours" (modUI.bas:512). All three sit inside the sort range (A:AZ) and the filter range (A:AS), so values there travel with the row on every sort.

### 2. Total Hours (G)
G is a formula, written once by `ConvertRosterToFormulas` (modOTLog.bas:144): `=H5+I5`. I is `=SUM(T5:AS5)`; each T:AS cell is a SUMIFS over `tblOTLog` (Hours Charged by Employee, Period Start, Status<>VOID; `PeriodFormula`, modOvertimeEntry.bas:467). H is written by the Kronos import.

Readers of G, all through one path:
- `LoadEligible` reads `C_TOTAL` into `EmpRec.TotalHrs` (modOvertimeEntry.bas:311). That feeds the call-out form list and "Now calling" label, the posting snapshot (`PL_HRS`, modPostings.bas:309), the printed roster (Call_Out_Roster.bas:71), and the "Total Hrs Then" log column.
- `SortRoster` sorts the sheet by code, G, J (Sort_Hours_Dept.bas:74).
- `RosterLookup` reads G for an out-of-department award (modPostings.bas:1201).
- `PublishMemberView` copies G (modOTLog.bas:447).
- `ConvertRosterToFormulas` snapshots G before and verifies it after (modOTLog.bas:68, 167).
So adding a provisional term to the G formula reaches everything.

### 3. What the Kronos import stores about dates
- Status line date = **the Kronos file's saved timestamp** (`FileDateTime(KRONOS_PATH)`), falling back to the import time (ImportKronosData.bas:84-88, 179-182). Both are kept as hidden workbook names `KronosFileDate` and `KronosImportedAt` holding the date serial (`SetDateName`, line 161).
- Hours History (`tblHoursHistory`, hidden sheet) columns: Import ID (KR-0001), Imported At, Kronos File Date, Employee, Worked Hours, Imported By (line 250). **No "data through" date exists yet.**

### 4. OT Log Type column
Values written: `Callout` (frmOvertimeEntry.txt:441), `Posting` (modPostings.bas:1287, 1321), `Adjustment` (modOTLog.bas:272), `Opening Balance` and `Manual Entry` (via `ConvertRosterToFormulas`). The only code that *reads* Type is `UpgradeOTLogLayout` (modLogLayout.bas:122, looks for "Posting"); voids match on Ref ID, not Type. So a new value in Type breaks nothing. Add Adjustment writes one row with Ref ADJ-####, Type, Shift Date, Employee, Note (the free-text reason), Hours Charged, Period Start, user, stamp (modOTLog.bas:269-279). It leaves column 7 "Reason" blank.

### 5. Version pattern and SystemCheck
Each module: `Public Const X_VERSION As String = "Name v20 (Oct 9)"`. Forms: a `' @version v20 (Oct 9)` comment line near the top (frmPostingEntry and frmTimeOff have none; SystemCheck only checks the three forms v20 changed). ThisWorkbook exposes `CodeVersion()`. `SystemCheck` (modSystem.bas:31-40) lists 9 modules by name and counts "N of 9 on EXPECTED_RELEASE"; `VerCheck` passes only when the constant contains "v21". **Consequence:** every listed module must ship with a v21 constant, including ones whose only change is the bump, or SystemCheck flags them. New modules go on that list.

### 6. Rules computed more than once
- **Call order:** `LoadEligible`/`SortRecs`/`RecBefore` is the one implementation for the form, posting list, printed roster and log order. But the roster sort buttons (`SortRoster`, Sort_Hours_Dept.bas:65-85) are a second copy: an Excel sort by code *text*, G, J, with no name tie-break. It agrees with `RecBefore` for today's codes but is a separate rule. `OutsideCandidates` sorts by hours then seniority on purpose (no grade), so it is a different rule, not a duplicate.
- **Charging:** posting charges live in `ComputeAward` and `ComputeNoTakers` (pure). The out-of-department award repeats the no-takers loop inline (modPostings.bas:1101-1113) with the signed-up exemption. The call-out charge (Refused/Left Message, grade cap, x1.75) is inline in `btnSubmit_Click` (frmOvertimeEntry.txt:434-437), so there is no callable call-out charge rule.
- **Eligibility:** `NotFreeReason` (modCalendar.bas:404-420) is the one place and the shift switch form uses it. `WorksConflict` in modShiftSwitch.bas:228-247 re-states the rest rule for switches (partial duplicate, not touched by v21).

## Proposed approach per item

### 1. Provisional accepted OT
- **Data-through date.** Store it on every Hours History snapshot as a new last column `Data Through`, added by an idempotent `EnsureHoursHistoryColumns` the import calls (no roster/log column is touched, so the upgrade-macro rule does not apply). Also keep the latest value as a hidden name `KronosDataThrough`. *Decision needed:* the brief says import date minus one; I recommend **Kronos file date minus one**, since an old file imported late would otherwise drop provisional hours Kronos does not have yet.
- **The amount.** Column C (hidden, currently unused) becomes the provisional helper with header C2 = "PROV". Each row is a formula over the log, same pattern as T:AS: `SUMIFS(Hours, Employee=D5, Status<>"VOID", Outcome="Accepted", Shift Date within T$2..AS$4, Shift<>"Nights", Shift Date>KronosDataThrough)` plus the same for Nights with `Shift Date>KronosDataThrough-1`, plus both again for `Awarded`. That is exactly the rule in the brief: a day shift is covered when D <= through, a night when D+1 <= through. Blank Shift on pre-upgrade call-out rows counts as Days (they are all covered by now anyway).
- **G** becomes `=H5+I5+C5`, written by `WritePeriodRow`'s neighbour in `ConvertRosterToFormulas` and verified the same way. One change, every reader picks it up. Void sets Status=VOID, so the SUMIFS drops it at once.
- **Visible where.** `EmpRec` gets `ProvHrs` (read from C). The call-out form shows "(84 hrs, 12 provisional)" in Now calling and the list's Total column; the posting form's eligible-count line says how many have provisional hours. No new roster column.
- **Done-when checks** map directly: an Accepted row recalculates C and G immediately (`AppendLogRows` already calls `Application.Calculate`); the import sets the name and calls `CalculateFull`; voids already recalc.
- Pure rule for SelfTest and Python: `ProvisionalCovered(shiftDate, isNight, dataThrough) As Boolean` in modOvertimeEntry; the SUMIFS encodes the same inequality.

### 2. DST nights
- New pure functions in modShiftCycle beside the rotation: `DstFallBack(y)` = Saturday before the first Sunday in November, `DstSpringForward(y)` = Saturday before the second Sunday in March, `DstNightHours(d) As Double` (13 / 11 / 12) and `DstNote(d)`. Setting `DST_USE_ACTUAL_HOURS = True` in the modOvertimeEntry config block.
- frmNewPosting: when date or shift changes and Hours still holds the last auto value (12 or a DST value), set 13/11 and show the one-line warning in the crew hint; otherwise leave what was typed. frmOvertimeEntry: same on date / Days-Nights change when Hours is blank or still the auto value; warning in `lblRot`. With the setting False, hours stay 12 and only the warning shows. Day shifts untouched.
- No change to charging: No Takers already charges `p.Hours x 1.75` = 22.75.

### 3. Home Unit column
- *Decision needed.* **Recommended: column B** (free since the migration): header B2 "HOME UNIT", list validation from `DeptList()`, width set, formatted like the other detail columns. No column insert, so no upgrade macro and no formula re-pointing. Alternative: a real insert via a macro modelled on `UpgradeOTLogLayout` (back up, move, verify every period total and Hours Charged, block save on mismatch).
- One-time `SetupHomeUnit` (Alt+F8): writes the header and validation, fills anyone with a code in exactly one of L:S, leaves others blank. Runs under the bulk-change rule (backup copy first, verify names and hours untouched after).
- SystemCheck and Needs Attention (new section E, admin only) list blank Home Units. Nothing else reads it in v21.

### 4. Adjustment types
- *Decision needed.* **Recommended:** keep Type = "Adjustment" and put the pick in column 7 "Reason" (blank on adjustment rows today, and already the event-reason column for call-outs and postings). Reports split on Reason; filters on Type still group all adjustments. Alternative per the brief: Type = "Adjustment - Correction".
- Required pick via a numbered `Application.InputBox` (1 Leave return, 2 Unit move, 3 New hire, 4 Correction, 5 Other), validated, before the existing reason/confirm prompts. Existing rows untouched.

### 5. Grade change log
- New module `modGradeLog`. `Workbook_SheetChange` passes `Sh, Target` to `GradeSheetChanged` as well as `CalSheetChanged`. It returns immediately unless `Sh` is the roster, `Target` intersects L:S below row 4, and a module flag `gSuppressGradeLog` is off.
- Old values come from a very-hidden sheet `GradeSnap` keyed by **employee name** (name, L..S codes). A sort changes row positions but not name-to-code pairs, so a sort produces no entries; filters fire no change event at all. Each row in the intersection is compared to its name's snapshot; only real differences are written to `tblGradeChange` (hidden admin sheet `Grade Log`: Timestamp, Employee, Unit, Old Code, New Code, Windows Login), then the snapshot row is refreshed. Macro writes set the flag or run with events off.
- Names absent from the snapshot (new hire, just renamed) are added silently without an entry. Snapshot is (re)built on open if missing and after `SetupHomeUnit`.

### 6. Daily backup on first open
- `BackupOnOpen` in modSystem, called first thing in `Workbook_Open`: skip quietly if read-only, not admin, or `BACKUP_FOLDER` blank; warn once and carry on if the folder is unreachable; skip if `Call Out Roster yyyy-mm-dd.xlsm` for today exists; else `SaveCopyAs`, then delete files in that folder matching exactly `Call Out Roster ####-##-##.xlsm` whose date is more than 30 days old. `SaveCopyAs` fires no save events, so the copy is the file as opened. SystemCheck flags a blank or unreachable folder and no backup today. **Path still needed from Graeme.**

### 7. SelfTest v1
- New module `modSelfTest`, Alt+F8 > `SelfTest`, pass/fail per case in one message box, nothing saved. Needs these small refactors so the rules are callable without sheets: make `SortRecs` public (`SortEmpRecs`), extract `CalloutCharge(outcome, grade, hours)` from the form and use it there (closes the duplicate in question 6), and split `NotFreeReason` into a pure `RotationFreeReason` (scheduled / rest / holiday) plus the table checks. Booked-off and switch cases use injected test fixtures (`SetTestFixtures` loads the absence and switch arrays the cache reads; cleared in a `Finally` even on error), so no table is written.
- `tests/` (pytest): `rules.py` mirrors rotation, holidays, DST, call order, charging and provisional coverage; `cases.csv` is the single source of expected answers; `gen_selftest_cases.py` emits the VBA constant block for modSelfTest from the CSV, so the two sides cannot drift.

## Package notes
- Modules changing: modOvertimeEntry, modOTLog, ImportKronosData, modShiftCycle, modSystem, modCalendar (Needs Attention section), modUI (column format), plus new modGradeLog and modSelfTest; version bumps on every SystemCheck-listed module. Paste-ins: ThisWorkbook, frmNewPosting, frmOvertimeEntry (each gets `@version v21`).
- One-time macros: `SetupHomeUnit`, and `RefreshRosterFormulas` (rewrites G with the provisional term and verifies).
