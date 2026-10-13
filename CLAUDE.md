# Call Out Roster: CLAUDE.md

Context for every coding session on this repo. Read this first, then the brief for the update you're building (`briefs/`). Reconciled on 2026-10-13 from the project's code kit (`docs/code-kit-2026-10/CLAUDE-kit.md`) and what is actually built through v21 and M1.

## What this is
An Excel desktop `.xlsm` (VBA) overtime call-out system for Syensqo's Niagara Falls phosphine plant.
- **Master file:** `V2 - Call Out Roster.xlsm`. Main tab `OT Roster` (setting `ROSTER_SHEET`; the code also finds it by headers).
- **Owner / maintainer:** Graeme (shift supervisor, admin). 8-10 supervisors use it, one at a time.
- **Units:** Phosphine, Derivatives, Cypure, Material Handler (Fumigation, called "Gases"), QC Lab (QC Derivatives), QC Gases, Utilities, SBP.
- **Plant terms:** NCN = No Coverage Needed. S/S or SS = shift switch. PHA = hazard-analysis coverage. 12-hour shifts, days 7 AM-7 PM, nights 7 PM-7 AM, nights dated by the evening they start.

**Nothing in this repo can be run.** There is no Excel here. Graeme imports the code, runs Debug > Compile, and tests by hand on a copy. So:
- Write compile-safe VBA (`Option Explicit` everywhere) and review it as if no compiler will check it.
- Prefer small, verifiable changes over rewrites.
- Anything that changes data runs a before/after check and blocks the save on a mismatch.
- **This repo is the single master for the code.** Nothing is installed from anywhere else; every release is built from a tagged commit.

## Repo layout
```
modules/     *.bas  standard modules (file name = VB_Name), imported via File > Import
paste-in/    *.txt  code for ThisWorkbook and each UserForm (pasted with Ctrl+A); no Attribute lines
member/      the members' workbook: modMember.bas + its ThisWorkbook.txt (a separate .xlsm, no master names)
updater/     modUpdater.bas: the Roster Updater, imported into its OWN workbook (never into the roster)
briefs/      one brief per update (v21.md, M1.md, v22.md ...)
docs/        planning docs, install notes (INSTALL-vNN.md), report-backs, the code kit as supplied
tests/       Python mirrors of the pure rules + known-answer cases + static checks (see Tooling)
dist/        built zips (gitignored)
CHANGELOG.md
```
**Never commit data:** no `*.xls*`, Kronos exports, posting PDFs or plant CSVs (names, phones, emails, leave codes). Test fixtures use made-up names only.

## Workbook map (as built, through v24.3)
### Modules (22 checked by SystemCheck, plus 6 legacy; the updater lives in its own workbook)
| Module | Holds |
|---|---|
| `modOvertimeEntry` | **All settings** at the top (`ADMIN_USERS`, folders, Kronos path, `KRONOS_STALE_DAYS`, `DST_USE_ACTUAL_HOURS`, `BACKUP_FOLDER`, member file names ...), roster column constants, `EmpRec`, `LoadEligible` (the one call order), `CalloutCharge`, provisional-OT formulas, OT Log access (`AppendLogRows`), protection, `ApplyView`. Every other module depends on it. |
| `modOTLog` | Roster formulas (`ConvertRosterToFormulas`), `RefreshRosterFormulas`, `AddAdjustment` (type required), `VoidEntry`, `RenameEmployee` (v24: opens the Employee pop-up on Rename), `PublishMemberView` (wrapper) |
| `modLogLayout` | OT Log column layout, `FormatOTLog`, `UpgradeOTLogLayout` (the model for verified upgrades) |
| `modMigration` | Original `RunMigration` (keep: the backup / move / verify / block-save pattern) |
| `modPostings` | Postings, awards, charging (`ComputeAward`, `ComputeNoTakers`, `ComputeOutside`, `ComputeCancel`), emails (CC constants, `POSTING_CLOSE_TIME`, `POSTING_CLOSE_DAYS` 14, `POSTING_OUTSIDE_LINES` 3), `ReopenPosting`, Outlook + mailto fallback |
| `modShiftCycle` | 28-day rotation (two 28-letter strings; day 1 = Sun Dec 28, 2025, B days / A nights), `CrewShiftOn`, `AvailableCrew`, DST rules (`DstNightHours`, `DefaultShiftHours`) |
| `modShiftSwitch` | Shift switches, `SwitchPickList`, `PersonFlag` (returns a code), `SwitchWorkingOn` / `SwitchOffOn`, publish batch cache |
| `modCalendar` | Absences / Coverage tables, eligibility (`NotFreeReason`, `RotationFreeReason`, `FilterNotFree`), `SlotsFor`, Calendar + Needs Attention drawing, CalMap change tracking, Import Old Cal |
| `modHolidays` | 11 plant holidays for any year + Holidays sheet "Observed on" |
| `modSystem` | `SystemCheck`, `OpenChecks`, daily backup (`BackupOnOpen`, `PruneBackups`) |
| `modUI` | Command bar (`BuildCommandBar`), status line, `ApplyRosterLayout`, `SetupHomeUnit`, form theme |
| `ImportKronosData` | Kronos import, Hours History snapshot (+ Data Through), `KronosDataThrough`, `EnsureDataThrough` |
| `modGradeLog` | Grade change log (`tblGradeChange`, snapshot `GradeSnap`) |
| `modSelfTest` | Alt+F8 > SelfTest: known-answer cases generated from `tests/cases.csv` |
| `modPublish` | `PublishMembers` (M1): builds the members' tables as values, privacy + order self-checks, fills the template, saves the members' file |
| `modLeave` | v22: sick calls (`SaveSickCall`, `SickBlocks`), leave going / returning (`ComputeLeaveReturn`, `TotalHoursAsOf`), convert imported calendar lines |
| `modAttendance` | v22.1: the sick / personal day rule (`AttendanceCore`), `AttendanceFor`, Sick Call preview, past sick days, the Attendance sheet |
| `modIntegrity` | v23: `IntegrityCheck` / `RunIntegrity` (read-only; PROBLEM / Check / Note), `IntegrityOnOpen`, the Integrity sheet, pure `PeriodCellCode`, `DuplicatesIn` |
| `modRollover` | v23: `RolloverYear` (the only code that removes OT Log rows, verified), Year End sheet, frozen FINAL copy, archive sheets, the 27th period column, `ArchiveHasActiveRef` |
| `modPeople` | v24: crew history (`tblCrewHistory`, `CrewOn`, pure `CrewOnCore`, the typed-crew prompt), the Employee pop-up's work (`SaveNewHire`, `SaveUnitMove`, `RenameEverywhere`, `RemoveEmployee`, `BringBack`), the Former sheet (`tblFormer`, `FormerNameSet`), `SetupPeople`, the shared before/after check (`RosterTotalsMap`, `TotalsDifference`). v24.3: planned unit moves (`tblPlannedMoves`, pure `PlannedCodeCore` / `PlannedListHours`, `UnitCodeOn`, `SavePlannedMove`, `CancelPlannedMove`, `ApplyDuePlannedMoves` on open, `ApplyPlannedMoves`, `PlannedMoveMenu`) |
| `modReports` | v24.1: Refused hours by kind, display only (`RefusedBucket`, `SplitTotals`, the status-bar line for the selected row, `RefusalsReport`) |
| `modSignups` | v24.1: the electronic sign-up inbox (`ReadSignups`, pure `ReduceSignups`, `ApplySignups` for the Award pop-up, `ArchiveSignups`); off unless `SIGNUPS_ON` |
| `modErrorLog` | v24.2: `LogError` (never raises), `ReportError` (the entry points' handler: logs, closes pop-ups, plain message), `ShowErrorLog`, `ErrorLogTest`; also writes `Roster error log.txt` beside the roster |
| `modSmokeTest` | v24.2: `SmokeTest`, practice copies only, never saves: 14 steps (v24.3: step 9 plans a unit move and cancels it) through the pop-ups' own functions, totals must come back. Holds no rules |
| `Call_Out_Roster` | Printed call-out roster (`Call_Out`), redrawn from the data |
| `Sort_Hours_Dept`, `OT_Posting`, `LoginFullAccess`, `modCleanup` | Sort buttons, legacy posting print, admin, one-time cleanup (not version-checked) |

### Forms
`frmOvertimeEntry` and `frmPostingEntry` have designer layouts (don't touch the designs). `frmNewPosting`, `frmShiftSwitch`, `frmTimeOff`, `frmSickCall`, `frmLeave`, `frmConvert`, `frmEmployee` (v24) are **blank in the designer and built entirely in code**. New forms follow the code-built pattern so install is "Insert > UserForm, set (Name), paste code". Every form and ThisWorkbook carries a `' @version vNN (date)` line; SystemCheck checks each against the release that last changed it (`FORM_RELEASES` / `THISWB_RELEASE` in modSystem), so bump that table when a form changes.

### Sheets and tables
| Sheet | Notes |
|---|---|
| `OT Roster` | Rows 1-4 command bar + navy date block (period start / "to" / end); frozen panes. **A = Payroll # (v24.2, `C_PAYROLL`: admin view only, filled by the Kronos import, not a key yet), B = Home Unit, C hidden = Provisional accepted OT (formula)**, D name, E crew, F phone, G Total Hours = H + I + C, H Worked (Kronos), I Refused = SUM(periods), J Plant Seniority, K EMAIL, L:S unit code columns (grade codes like D1.5, P2), T:AS pay-period columns = **SUMIFS over the OT Log**. |
| `OT Log` (`tblOTLog`) | **Append-only, newest first.** Event: Ref ID, Type, Shift Date, Unit, Shift, Crew, Reason, Hours. Person: Call Order, Employee, Dept Code, Total Hrs Then, Outcome, Note, Hours Charged, Period Start. Audit (grouped): Entry ID, Logged By, Timestamp, Status, Void Note. Adjustment rows carry their type in **Reason** (column 7). |
| `Postings` (`tblPostings`, `tblPostingList`) | Posting header + frozen list with hours at posting time. v24: the list's 10th column `Signup Hours` (blank = whole shift); a Result starting `Exempt` = moved unit or left, never charged on that posting |
| `Absences` (`tblAbsence`) | One row per booking; start/end (blank end = open-ended), hours per shift, times note. Admin only, hidden in saved copies. |
| `Coverage` (`tblCoverage`) | One row per date + shift line |
| `Shift Switches` (`tblShiftSwitch`) | SS-#### with payback links |
| `Calendar`, `Needs Attention` | Drawn, never typed into. Needs Attention admin-only. |
| `Holidays` | Admin; Observed on column for weekend holidays |
| `Hours History` (hidden) | Dated snapshot of worked hours per import, with Data Through (v21) |
| `Grade Log`, `GradeSnap` | Grade change log (admin, hidden) and its very hidden name-keyed snapshot |
| `Attendance` | v22.1: drawn sick / personal day counters per person (admin, hidden in saved copies). v24.1: one row per person with a column per occurrence, only people with something logged, three sorts. The SHEET is never published; since M2 the five COUNTS per person are (see Privacy) |
| `Refusal Report` | v24.1: output-only, redrawn by `RefusalsReport` (admin) |
| `Integrity` | v23: the last IntegrityCheck report (admin, drawn) |
| `Error Log` (`tblErrorLog`) | v24.2: When, User, Procedure, Error, Detail, Release, Seen. Admin, never published |
| `Year End` (`tblYearEnd`) | v23: each person's final G / H / I / C per closed year (admin) |
| `OT Log 2026` ... (`tblOTLog2026`) | v23: one per closed year, the archived log rows; protected even in admin view |
| `Crew History` (`tblCrewHistory`) | v24: Employee, Crew, Effective From (+ who, when). Admin. Roster column E shows the newest row |
| `Planned Moves` (`tblPlannedMoves`) | v24.3: one row per unit move entered ahead of time: ID `PM-0001`, Employee, New Unit, Grade Code, Codes To Clear, New Crew, Old Crew, Effective From, Status (Planned / Applied / Cancelled), who, when, Closed On, Note. Admin, never published, made on first use; rows are never removed |
| `Former` (`tblFormer`) | v24: one row per person taken off the roster, as values (final figures, Left On, Status Former / Returned). Admin, never published; no phone or email |
| `CalMap` (hidden) | Calendar change counters |
| `Overtime Posting`, `Blank Call Out Roster` | Output only, redrawn each time |
| `Legacy Entry Log` (hidden) | Pre-migration history |

### Ref ID prefixes
`OB-` opening balance, `MAN-` manual cells, `S-` call-out session, `P-` posting, `ADJ-` adjustment, `SS-` shift switch, `KR-` Kronos snapshot, `#n` = Entry ID. `LV-` leave return, `SC-` sick call (v22), `UM-` unit move, `NH-` new hire (v24). `PM-` is a planned move's ID in `tblPlannedMoves` (v24.3): it is never an OT Log ref.

## Canonical business rules
These win over anything older in `docs/`. Each rule lives in **one** function that forms, printouts, Needs Attention and the member publish all call. If you find a second copy, consolidate it instead of patching both.

**Call order:** grade code (in that unit's column) -> Total Hours G (lowest first) -> Plant Seniority J (earliest first) -> name. Implemented once: `LoadEligible` / `SortEmpRecs` (modOvertimeEntry). Used by the call-out form, the posting list, the printed roster, the member publish.

**Total Hours** = Kronos worked hours (H) + non-void Hours Charged in the OT Log for the current year (I) + provisional accepted OT (C): Accepted call-outs and Awarded postings count from the moment they are logged, even for a future shift, until Kronos covers the shift. Kronos data runs **through the day before the file date** (`KronosDataThrough`). A day shift on D is covered when D <= through; a night on D when D + 1 <= through.

**Who's listed (call-outs and postings):** anyone with a code in the unit's column, **any grade**, who is free for that date and shift (`NotFreeReason` in modCalendar). Not free when:
| Reason | Rule |
|---|---|
| Scheduled | Crew on that shift per rotation. A/B and C/D work their pair's **day** shifts whenever either is on days, never nights. DAYS = Mon-Fri days, **off on plant holidays** (observed date). **The crew is the one the person is on THAT DATE** (v24: `CrewOn`, crew history). |
| Rest rule | Day OT: crew coming off nights that morning. Night OT: crew starting days next morning. A switch worked the night before / day after counts. |
| Booked off | Absences row covering the date (not imported calendar lines, which are display only until converted in v22). |
| Switch | Working a switch that day, or gave up a shift that day. |

**Charging (only grade <= 2, x 1.75, never out-of-department people):**
| Event | Charged | Period |
|---|---|---|
| Call-out | Refused or Left Message (`CalloutCharge`) | callout date |
| Posting: Submit Award | Ahead of the **last** awarded person in call order **and didn't sign up** (`ComputeAward`) | shift date |
| Posting: Award Out of Dept | Everyone eligible in dept who didn't sign up (`ComputeOutside`) | shift date |
| Posting: No Takers | Everyone eligible on the list (`ComputeNoTakers`) | shift date |
| Call-out: part accepted (v24.1) | Someone takes part of the hours; everyone BELOW them in call order is offered what is left, and a Refused / Left Message there is charged on the hours **offered** x 1.75. A whole acceptance changes nothing for those below (`CalloutOffers`). The Accepted row carries the hours taken | callout date |
| Posting: Cancel | Nobody (`ComputeCancel`) | - |
| Posting: partial sign-ups (v24) | Same people as Submit Award, charged the posting's **full** hours x 1.75. Full-shift sign-ups are awarded first; a partial only when no full-shift sign-up is left for the position (`AwardOrderProblem`, blocked otherwise). A partial sign-up that isn't awarded: not charged. Hours nobody covers: dropped, uncharged. The Awarded log row carries the hours awarded | shift date |
| Posting: moved unit / left (v24) | Never charged on an open posting for a shift on or after the date (list row marked `Exempt - ...`) | - |
| DST nights | **Actual hours** (fall night = 13 h -> 22.75; spring night = 11 h -> 19.25), `DST_USE_ACTUAL_HOURS` | - |

Postings compare using the **posted** totals (frozen at creation). Awarded is chosen in the Award pop-up, not "first sign-up".

**Year:** totals reset to zero each year; no carry-forward. Pay periods are biweekly, Sunday start; 2026 periods start Dec 28, 2025; 2027 starts Dec 27, 2026. **The year is one cell:** the roster's period-1 start (T2); every header, formula, the provisional window, leave top-ups and the publish read it (`RosterYearFrom` / `RosterYearTo` / `RosterYear`). A year has 26 periods, or 27 when one more fits before the next Jan 1 (first: 2033); the 27th period is column AT and `PeriodLastCol()` says which is in use, so never loop to `C_PERIODN` directly. `RolloverYear` (v23) moves the year: see `docs/rollover-one-pager.md`.

**Privacy:** STD, parental, LOA show only as "Leave" everywhere outside the Absences sheet. Never publish phone, email, leave type, leave notes, void notes, Logged By or Timestamp, the Windows login behind an electronic sign-up, or the DATE of any sick or personal day. The members' file gets values only, pushed by the master, and contains no rules. **One deliberate exception (M2, Oct 12, 2026):** each person's sick-time COUNTS are published in `dAttend` (sick days used, paid days left, occurrences used, personal days used, one of four status phrases; the only date is the "until" date) and shown on the card of whoever picks that name; `AttendPrivacy` in modPublish refuses anything else.

**Refused by kind (v24.1):** Refused stays ONE figure and drives Total Hours and call order. Call-out / Posting / Other is a display split of the same log rows by the Type column (`modReports`): status bar for the selected row, the members' file, the Refusals Report. No roster column, no formula.

**Electronic sign-ups (v24.1, off until `SIGNUPS_ON`):** the members' file writes one small request file per click into the `Signups` folder; the master reads them when a posting is loaded for award and pre-ticks Signed up. The supervisor's tick is what counts; the paper sheet stays valid.

**People changes (v22/v24):** every rule-driven hours change is an **Adjustment** row in the OT Log with the calculation in Note, its type in Reason, shown in a confirm before saving, undoable by Void Entry. v24's Employee pop-up does Add (`NH-`: the home unit's average, grade <= 2, not on leave on the start date), Move to unit (`UM-`: up to the highest among grade <= 2 in the new unit, never lowered), Rename (every table together), Remove (the row goes to the Former sheet) and Bring back. Peers are everyone with a code in that unit's column, as the leave top-up counts them.

**Planned unit moves (v24.3):** Move to unit with a date after today changes nothing on the roster; it writes one Planned row. For a shift **on or after** the date the person is off the cleared units' lists and on the new unit's with the planned code (`UnitCodeOn`, called by `LoadEligible`; with no date the date is today), **listed at the hours they would have after the top-up** (`PlannedListHours`: for call order only, nothing logged). The first admin open on or after the date applies it through `SaveUnitMove` with the hours **as of the effective date**; a move that can't be applied stays Planned with the reason in its Note. Never edited: cancel and re-enter. One waiting move per person. Rename carries it, Remove cancels it. A crew change that comes with it is written to the crew history at entry, dated the move's date. Old-unit open postings from the date are marked `Exempt - moves <date> (PM-000n)`; a cancel removes only marks with that ID.

## Engineering conventions
1. **Versioning.** Every module has `Public Const X_VERSION As String = "Name vNN (Mon d)"`; forms and ThisWorkbook have the `' @version` line; `modSystem.EXPECTED_RELEASE` is the release and `N_MODULES` the count. Bump every module that ships, and the expected release, in the same update. SystemCheck must say "N of N on vNN".
2. **Append-only log.** Never edit or delete OT Log rows. Corrections are voids (Status VOID, Void Note) plus a new entry. Never type hours into roster cells; they are formulas over the log. The ONE sanctioned removal is `RolloverYear` moving a closed year's rows to its `OT Log YYYY` sheet, verified. (Roster ROWS are removed only by `RemoveEmployee`, verified, and names are changed only by `RenameEverywhere`, which rewrites the Employee column of the log and every other table together.) Because of it, IDs come from `NextRefID` / `NextEntryID` (which honour the `IdMark_*` high-water marks): never work out a next ID by scanning the log yourself.
3. **Verified data changes.** Any macro that restructures data: dated backup first -> change -> compare every roster period, Refused (I) and Total (G) value and the log's total Hours Charged before vs after -> on mismatch, `MarkNoSave` and tell the user to close without saving. Copy `UpgradeOTLogLayout` / `RunMigration` / `SetupHomeUnit`.
4. **Read every row, whatever the view.** Use `ListObject.DataBodyRange` or the roster's cell reads, never last-visible-row. Handle `DataBodyRange Is Nothing`.
5. **Columns.** Roster and log columns are constants in modOvertimeEntry (`C_*`, `L_*`); use them. For anything new, find by header text where the code already does that. **Crew (v24): never read roster column E to answer a question about a date.** Use `CrewOn(emp, d)`, `EmpCrewCode(emp, d)`, `EmpShiftOn(emp, d)` or `LoadEligible(..., asOf:=d)`. Column E is the newest crew, for display. Code that writes column E turns events off and adds the crew row itself (`AddCrewRow`). **Unit codes (v24.3): a roster unit column is the code TODAY.** To ask whether someone holds a code in a unit on a date, use `LoadEligible(..., asOf:=d)` or `UnitCodeOn(emp, unit, d, rosterCode)`, which count a waiting planned move from its date. Display code (sorts, the Former sheet, the publish's Roster table, the peer lists for a top-up) reads the column as it is.
6. **Bulk work:** turn off ScreenUpdating, EnableEvents and calculation, and restore them in an error handler on every exit path. Capture the "was" values **before** `On Error GoTo`.
7. **Calendar dirty flag:** writes to Absences, Coverage, Postings or Shift Switches with events off must call `MarkCalDirty`.
8. **Late binding only** (Outlook, Scripting). No new references.
9. **Dates:** build with `DateSerial`; parse typed dates through the existing parser in frmShiftSwitch (handles "Oct 12" and year rollover). Never `CDate` a locale-ambiguous string. Cells that hold month or date **text** get `NumberFormat = "@"` first, or Excel converts them.
10. **Admin gating** via `ADMIN_USERS`; admin buttons are hidden for others (`ui_adm_` shapes).
11. **Flat, list-shaped tables** (one header row, stable names, no merged cells) so they can become SharePoint Lists later.
12. **Messages** tell a supervisor what to do next, in plain words.
13. **Names are an interface.** Public procedure names that buttons call (OnAction in `BuildCommandBar`) and macro names in install notes don't change. If a rename is unavoidable, `ApplyRosterLayout` must rebuild the buttons and the install note must say so.
15. **Entry points catch errors (v24.2).** A button macro, a macro that opens a pop-up, or a workbook event starts with `On Error GoTo Oops` and ends `Exit Sub / Oops: / ReportError "Name"`. Inside it, restore a handler with `On Error GoTo Oops`, never `On Error GoTo 0`. A handler that already restores state and shows its own message adds `LogError "Name", Err.Number, em` right AFTER it has captured `em = Err.Description` (LogError clears Err). An error inside a pop-up travels up to the macro that opened it.
16. **SmokeTest switch (v24.2).** `gSmokeRun` is True only while SmokeTest runs on a practice copy. New confirms on a path SmokeTest uses go through `Ask(prompt, buttons, title, smokeAnswer)` and success messages through `Tell`; `SaveIfWritable` does nothing while it is on. Never test `gSmokeRun` to change a rule.
14. **File format:** CRLF and ASCII only in `.bas` and paste-in `.txt` (`.gitattributes` enforces CRLF; write `->`, `-`, `...`). Every `.bas` starts with `Attribute VB_Name`. Every file has `Option Explicit`.

### VBA compile and runtime traps to check in review
- `Dim a, b As Long` makes `a` a Variant. Declare each.
- Max 24 line continuations per statement; 64 KB per procedure. Split big builders.
- Use `Long` for row counts, not `Integer`.
- Duplicate procedure names across modules -> "Ambiguous name". Search before adding a Public name.
- **A local or parameter with the same name as a procedure hides it** (names are case-insensitive): `tests/shadowcheck.py` finds these. Avoid helpers named `S`, `N`, `D`, `P`, `F`.
- `Lock`, `Unlock` and other reserved words can't be procedure names. Data-type and built-in names can't be variables either: `Single`, `Date`, `String`, `Long`, `Len`, `Int`, `Abs`, `Input`, `Type`, `Write` ... (`Dim single As Long` is a syntax error).
- One-line `If ... Then x ElseIf ...` is a compile error; use a block If.
- `ReDim Preserve` can only grow the **last** dimension; size 2-D tables up front.
- Module-level variables must be declared above the first procedure.
- Code-built forms: control names must be unique; `WithEvents` handlers need the module-level `WithEvents` pattern already used.
- `Application.Match`/`VLookup` return errors, not raise them: test with `IsError`.
- Excel auto-converts text like "Oct 2026" or "Mon Oct 12" written to a General cell into a date.

## Tooling (`tests/`)
- `py -3.14 -m pytest tests -q`: the Python mirror of the pure rules (`rules.py`) against `cases.csv`, the same table SelfTest uses. Add cases for every rule touched; run `python tests/gen_selftest_cases.py` to paste them into modSelfTest.
- `python tests/vbacheck.py .`: blocks, Option Explicit, CRLF/ASCII, duplicate Publics (its hits on one-line `Sub X(): ...: End Sub` are false positives).
- `python tests/shadowcheck.py .`: locals hiding procedures.
- `python tests/undeclcheck.py .`: names used but declared nowhere ("Variable not defined"), and declared names that are VBA reserved words ("Syntax error").
- `python tests/callcheck.py .`: duplicate declarations, missing GoTo labels, ambiguous names inside a module, wrong argument counts on calls to our own procedures, `Foo(a, b)` called as a statement, ByRef argument type mismatches, fields a Type doesn't have.
- Patch scripts: write them with the Write tool into the scratchpad and run them as files. Heredocs in this shell mangle quotes and backslashes.
- **Run all four checkers (vbacheck, shadowcheck, undeclcheck, callcheck) before every review and every zip.** Two compile errors shipped in M1 through v23 and were found by Graeme at install (Oct 12, 2026: an undeclared `cap` and a variable named `single` in modPublish). Three rounds of reviewers reading the code missed both; the checkers find both. A clean run is good evidence, not proof: nothing here is a compiler.
- `python tests/build_zip.py vNN vPREV`: builds `dist/Callout_Roster_vNN.zip` from `docs/INSTALL-vNN.md`. From v22 on every zip holds the **full** module, paste-in and member set, so it can serve as the master.
- Python 3.12 is at `C:\Users\Graeme\AppData\Local\Programs\Python\Python312\python.exe` (no pytest); pytest is on `py -3.14`. GNU `sed -i` strips CRs in this shell: re-run the CRLF fix on any file it touches.

## How every update ships
1. Start from the current master (the latest tag in this repo). Read the brief and every module it touches.
2. **Report back before writing code:** check every (verify) item against the code; ask real decisions as **multiple-choice questions** (AskUserQuestion), offering the brief's default first. Wait for a go-ahead.
3. Build to the brief, one commit per item.
4. **Independent review** of every changed file by a separate read-only agent that didn't write it: compile traps above, rule consistency with this file, filter/visible-row bugs, events/calc restored on every exit path, privacy, no new write paths to the roster or OT Log outside the existing ones. Fix, then re-review the fixes.
5. Add SelfTest + `tests/` cases for every rule touched.
6. Package `dist/Callout_Roster_vNN.zip` (full set) with `INSTALL-vNN.md`: Remove/Import list, re-paste list, new forms, Compile, macros to run once, SystemCheck expectation, smoke test. INSTALL lists only what changed.
7. Update `CHANGELOG.md` and the affected `docs/`; tag `vNN`.

**Install freezes:** no updates Dec 20 - Jan 4 except the rollover; none in the week before Jan 23, 2027.

## Decisions already made (don't re-ask)
- Postings list any grade; only grade <= 2 charged.
- Totals reset yearly, no carry-forward.
- Stat holidays: DAYS off and free; A/B, C/D stay scheduled. Weekend holidays use the plant's announced weekday (Holidays sheet).
- Provisional accepted OT counts **as soon as it's logged, even for a future shift**, until Kronos covers it. Data-through = Kronos **file date** minus one (chosen over import date).
- DST nights charge **actual** hours (`DST_USE_ACTUAL_HOURS = True`).
- Home Unit lives in **roster column B** (free since the migration; no column insert). Graeme fills the blanks.
- Adjustment type is stored in the OT Log **Reason** column (7); Type stays "Adjustment".
- Grade changes are logged by comparing to a name-keyed snapshot (sorts log nothing).
- Partial sign-ups (v24): a full-shift sign-up beats a partial one even with more hours; a partial is awarded only if nobody takes the whole shift; uncovered remainder drops off, uncharged.
- Leave return (v22): + home unit's average Total Hours gained between first day off and return (grade <= 2, not on leave). Leave spanning the year reset tops up only from the reset.
- Full unit move (v24): up to the highest Total Hours among grade <= 2 in the new unit; never lowered. Extra sign-offs don't trigger it.
- New hire (v24): home unit's average (grade <= 2, not on leave).
- Sick call-in (v22): opens a line, starts a call-out; off call lists for any shift starting within 24 h after the sick shift ends; shown as "Sick" on the calendar; WSIB and LTD are added as leave types.
- Attendance (v22.1): 5 paid sick days / calendar year in up to 3 occurrences (day = scheduled shift, partial counts; occurrence = consecutive scheduled shifts), 5 unpaid personal days; over 5 combined = no paid sick until a year after the 2nd occurrence's first day, then restart to Dec 31; a sick call on an unscheduled shift counts; members see "Personal"; vacation and personal days only on scheduled shifts (refuse none, offer to trim ends). Tracking only: never touches hours. **No back-fill:** counting starts at the v22.1 install; 2027 is the first complete year.
- Member workbook (M1, built): plant PCs, desktop Excel, everyone's hours visible, master pushes values into the template next to the master, published to `PUBLISH_FOLDER`, My history shows anyone, member file has no rules.
- Platform: stay on the Excel desktop master through Q1 2027; keep tables platform-ready.
- Release zips hold the full set from v22 on.
- `PUBLISH_FOLDER = "."` (Oct 12, 2026): Member View saves the members' file next to the master for now (`modPublish.PublishFolderPath`). It moves to the members' network folder at the next roll-out wave; the P: path is kept in the comment above the setting.
- `BACKUP_FOLDER = "."`: daily backups go into the roster's own folder (`ResolveFolder` handles "." and ".\sub"). Can be pointed at a network folder later. `ARCHIVE_FOLDER = "."` the same way.
- v24 (Oct 12, 2026): refusers are charged the posting's full hours when only a partial is awarded; a partial award while a full-shift sign-up is free is **blocked**; Remove **moves the row to a Former sheet** (Graeme chose it over the brief's inactive flag: not date-aware, so it is run on or after the last day); someone who moves or leaves is never charged on open postings from that date; a rename leaves closed-year sheets (`OT Log 2026`, `Year End`) under the old name with a note on the Grade Log; a unit move's date is today or earlier (unit codes aren't dated), a crew change may be future-dated; one v24 install.
- v24.1 + M2 (Oct 12, 2026): Refused split shown on the **status bar + members' file + report, no roster columns**; "other" only when not zero; members see **who** signed up (as on paper); call-out refusers after a partial acceptance are charged **the hours on offer** (postings stay at the full posting hours); Find me does **not** remember the person; members' postings = the ones they are qualified for (listed, then out of department = a code in the unit) plus one plain table; **sick-time counts only** are published for the picked name; planned unit moves are **v24.2**.
- v24.2 (Oct 12, 2026): install kit first, planned moves as v24.3; Payroll # captured now, lookups switch in v25.1; **the updater never lives in the roster** (v1 ran inside it and crashed Excel: it is a separate workbook, and it swaps code in place rather than remove + import); rules that live only in pop-ups stay there until v25 and SmokeTest lists them as not covered; a practice copy publishes only `PRACTICE - ...` beside itself; `Roster error log.txt` sits beside the roster (decide where it goes when `PUBLISH_FOLDER` moves to the members' network folder: it must not end up where members can read it).
- v24.3 (Oct 12, 2026, decided before any code): a planned mover is listed in the new unit, for shifts on or after the date, **at the hours they would have after the top-up** (an estimate for order only); a planned move is **cancelled and re-entered**, never edited; a rename carries it, Remove cancels it, a second move is refused until it is cancelled; the top-up is worked out **as of the effective date**; the members' file shows it **in Who's next only**; a crew change that comes with it is **dated at entry**. v25, v26, calendar speed and the Payroll # key are all **moved ahead of go-live**.
- v24.3 build (Oct 13, 2026): go-ahead given after v24.2 installed clean. The planned-move apply runs inside `PeopleOnOpen` (no ThisWorkbook change); the `Planned Moves` sheet is made on first use (no one-time macro); the exempt marks keep the text `moves <date> (PM-000n)` after the move is applied; the updater identifies a roster by `modOvertimeEntry` OR the `tblOTLog` table.
- Year end (v23): old log rows are archived AND removed at the rollover (Graeme chose the brief's option over keeping one growing log); 26- and 27-period years are both handled now; nothing is built for OT worked Dec 27-31 or for Kronos year-to-date until the first January 2027 export is seen; the rollover is run by Graeme with his boss; IntegrityCheck PROBLEMs block it, Checks don't; the live file can't be rolled before the new year's first period starts.

## Still open (use the default unless Graeme answers)
| Question | Default |
|---|---|
| Boxing Day 2026 observed weekday | **Answered:** Monday Dec 28, 2026 for day workers. Entered on the Holidays sheet ("Observed on"), no code. |
| Network paths (`REPORT_FOLDER`, `PUBLISH_FOLDER`, `KRONOS_PATH` on P:) | Everything is built and tested locally on Graeme's desktop; the P: links are wired in at go-live. SystemCheck will flag them as unreachable until then. |
| Who runs the rollover / where the frozen copy lives | **Answered:** Graeme runs it together with his boss (two people present; v23 makes it a guided, verified macro with a frozen copy). Frozen-copy location still a settings constant, blank until given. |
| Does Kronos export an employee number | **Answered Oct 12, 2026: YES.** The Kronos report's first column is `Payroll #` (8 digits), then `Employee Name`, ten pay-code columns and `Grand Total` (the import reads the name from column 2 and the total from column 13 today). This replaces the earlier "name only" answer: the employee key in v25.1 is the Payroll #, not a generated `E-0001`. It is a personal identifier: admin-only, never published (the publish's phone check already stops any 7+ digit run). |
| Kronos worked hours after Jan 1 (YTD or cumulative), and OT worked Dec 27-31 | Check the first January 2027 export; the import shows a reminder on that first run. A small follow-up if hours are missing or cumulative. |
| `ADMIN_USERS` logins | Graeme's only for now; the other supervisors are added later (settings constant, no code) |

## Roadmap
| Brief | Target | Theme | State |
|---|---|---|---|
| v20 | Oct 9 | Safety and stats | installed; gate passed |
| M1 | live Nov 15 | Member workbook + Publish | built, tagged `M1` |
| v21 | before Oct 31 | Provisional OT, DST, Home Unit, adjustment types, grade log, backup, SelfTest, printable roster | built, tagged `v21`, installed |
| v22 | mid-Nov | Sick Call, Leave going/returning, convert imported lines | built, tagged `v22` |
| v22.1 | Oct 2026 | Sick / personal day counters, Attendance sheet, vacation on scheduled shifts | built, tagged `v22.1` |
| v23 | build by Nov 30, live Dec 27 | RolloverYear, Integrity check, docs | built, tagged `v23`; dry run on a copy first week of December |
| v24 | install Jan 5 - 15, 2027 | Employee pop-up, crew history, partial OT | built, tagged `v24`; needs the rollover done first |
| v24.1 + M2 | with v24, or after Jan 23 | Refused by kind + report, Attendance tab, part-accepted call-outs, sign-up inbox, publish reminder; members: Find me, month grid, my postings, Sign me up, sick-time counts | built, tagged `v24.1` |
| v24.2 + M2.1 | with or after v24.1 | Install kit: error log, SmokeTest, practice copy, Roster Updater (own workbook); Payroll # capture; Refusals Report range; members' fixes from the Oct 12 test | built, tagged `v24.2`; **installed by hand on Graeme's copy Oct 13, 2026: SystemCheck, SelfTest and SmokeTest passed**; the updater's `UpdateMemberTemplate` said "No roster is open" (fixed in v24.3's Updater v3) |
| v24.3 | with or after v24.2 | Unit moves entered ahead of time (planned moves, applied on open); Roster Updater v3 | built, tagged `v24.3`; brief `briefs/v24.3.md`, report-back `docs/v24.3-report-back.md` |
| v25 | before go-live (moved up Oct 12, 2026) | Pop-up rules into modules (SmokeTest covers them), protection without the VBA password | kit brief 06; starts when v24.2's SmokeTest has passed |
| v26 | before go-live (moved up) | Employee history report, "why was I charged", board print, payback aging | kit brief 07; minimum staffing waits for the minimums |
| v25.2 | before go-live (moved up) | Calendar speed, alone | kit brief 06 item 3 |
| v25.1 | before go-live (moved up), last | Payroll # as the key in every table, alone | needs column A filled by a Kronos import, and a go-live date |
| go-live | not set | VBA password, P: paths, `ADMIN_USERS`, members' folder | |

**Pre-release rule (Oct 12, 2026):** nothing is built more than one release ahead of what has run in Excel. Review and order: `docs/v24.3-report-back.md` part 2.

## Working with Graeme
- Ask questions as multiple choice (AskUserQuestion), not in plain text.
- Start each update by reading the brief and every module it touches. Report anything that changes the plan, and wait for a go-ahead before writing code.
- If a brief conflicts with this file, ask.
