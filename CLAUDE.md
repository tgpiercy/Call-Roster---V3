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
briefs/      one brief per update (v21.md, M1.md, v22.md ...)
docs/        planning docs, install notes (INSTALL-vNN.md), report-backs, the code kit as supplied
tests/       Python mirrors of the pure rules + known-answer cases + static checks (see Tooling)
dist/        built zips (gitignored)
CHANGELOG.md
```
**Never commit data:** no `*.xls*`, Kronos exports, posting PDFs or plant CSVs (names, phones, emails, leave codes). Test fixtures use made-up names only.

## Workbook map (as built, v21 + M1)
### Modules (13 checked by SystemCheck, plus 5 legacy)
| Module | Holds |
|---|---|
| `modOvertimeEntry` | **All settings** at the top (`ADMIN_USERS`, folders, Kronos path, `KRONOS_STALE_DAYS`, `DST_USE_ACTUAL_HOURS`, `BACKUP_FOLDER`, member file names ...), roster column constants, `EmpRec`, `LoadEligible` (the one call order), `CalloutCharge`, provisional-OT formulas, OT Log access (`AppendLogRows`), protection, `ApplyView`. Every other module depends on it. |
| `modOTLog` | Roster formulas (`ConvertRosterToFormulas`), `RefreshRosterFormulas`, `AddAdjustment` (type required), `VoidEntry`, `RenameEmployee`, `PublishMemberView` (wrapper) |
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
| `Call_Out_Roster` | Printed call-out roster (`Call_Out`), redrawn from the data |
| `Sort_Hours_Dept`, `OT_Posting`, `LoginFullAccess` | Sort buttons, legacy posting print, admin (not version-checked) |

### Forms
`frmOvertimeEntry` and `frmPostingEntry` have designer layouts (don't touch the designs). `frmNewPosting`, `frmShiftSwitch`, `frmTimeOff` are **blank in the designer and built entirely in code**. New forms follow the code-built pattern so install is "Insert > UserForm, set (Name), paste code". Every form and ThisWorkbook carries a `' @version vNN (date)` line that SystemCheck reads.

### Sheets and tables
| Sheet | Notes |
|---|---|
| `OT Roster` | Rows 1-4 command bar + navy date block (period start / "to" / end); frozen panes. **A hidden (empty), B = Home Unit, C hidden = Provisional accepted OT (formula)**, D name, E crew, F phone, G Total Hours = H + I + C, H Worked (Kronos), I Refused = SUM(periods), J Plant Seniority, K EMAIL, L:S unit code columns (grade codes like D1.5, P2), T:AS pay-period columns = **SUMIFS over the OT Log**. |
| `OT Log` (`tblOTLog`) | **Append-only, newest first.** Event: Ref ID, Type, Shift Date, Unit, Shift, Crew, Reason, Hours. Person: Call Order, Employee, Dept Code, Total Hrs Then, Outcome, Note, Hours Charged, Period Start. Audit (grouped): Entry ID, Logged By, Timestamp, Status, Void Note. Adjustment rows carry their type in **Reason** (column 7). |
| `Postings` (`tblPostings`, `tblPostingList`) | Posting header + frozen list with hours at posting time |
| `Absences` (`tblAbsence`) | One row per booking; start/end (blank end = open-ended), hours per shift, times note. Admin only, hidden in saved copies. |
| `Coverage` (`tblCoverage`) | One row per date + shift line |
| `Shift Switches` (`tblShiftSwitch`) | SS-#### with payback links |
| `Calendar`, `Needs Attention` | Drawn, never typed into. Needs Attention admin-only. |
| `Holidays` | Admin; Observed on column for weekend holidays |
| `Hours History` (hidden) | Dated snapshot of worked hours per import, with Data Through (v21) |
| `Grade Log`, `GradeSnap` | Grade change log (admin, hidden) and its very hidden name-keyed snapshot |
| `CalMap` (hidden) | Calendar change counters |
| `Overtime Posting`, `Blank Call Out Roster` | Output only, redrawn each time |
| `Legacy Entry Log` (hidden) | Pre-migration history |

### Ref ID prefixes
`OB-` opening balance, `MAN-` manual cells, `S-` call-out session, `P-` posting, `ADJ-` adjustment, `SS-` shift switch, `KR-` Kronos snapshot, `#n` = Entry ID. Planned: `LV-` leave return, `SC-` sick call (v22), `UM-` unit move, `NH-` new hire (v24), `YE-` year-end (if needed).

## Canonical business rules
These win over anything older in `docs/`. Each rule lives in **one** function that forms, printouts, Needs Attention and the member publish all call. If you find a second copy, consolidate it instead of patching both.

**Call order:** grade code (in that unit's column) -> Total Hours G (lowest first) -> Plant Seniority J (earliest first) -> name. Implemented once: `LoadEligible` / `SortEmpRecs` (modOvertimeEntry). Used by the call-out form, the posting list, the printed roster, the member publish.

**Total Hours** = Kronos worked hours (H) + non-void Hours Charged in the OT Log for the current year (I) + provisional accepted OT (C): Accepted call-outs and Awarded postings count from the moment they are logged, even for a future shift, until Kronos covers the shift. Kronos data runs **through the day before the file date** (`KronosDataThrough`). A day shift on D is covered when D <= through; a night on D when D + 1 <= through.

**Who's listed (call-outs and postings):** anyone with a code in the unit's column, **any grade**, who is free for that date and shift (`NotFreeReason` in modCalendar). Not free when:
| Reason | Rule |
|---|---|
| Scheduled | Crew on that shift per rotation. A/B and C/D work their pair's **day** shifts whenever either is on days, never nights. DAYS = Mon-Fri days, **off on plant holidays** (observed date). |
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
| Posting: Cancel | Nobody (`ComputeCancel`) | - |
| DST nights | **Actual hours** (fall night = 13 h -> 22.75; spring night = 11 h -> 19.25), `DST_USE_ACTUAL_HOURS` | - |

Postings compare using the **posted** totals (frozen at creation). Awarded is chosen in the Award pop-up, not "first sign-up".

**Year:** totals reset to zero each year; no carry-forward. Pay periods are biweekly, Sunday start; 2026 periods start Dec 28, 2025; 2027 starts Dec 27, 2026. Rollover is v23.

**Privacy:** STD, parental, LOA show only as "Leave" everywhere outside the Absences sheet. Never publish phone, email, leave type, leave notes, void notes, Logged By or Timestamp. The members' file gets values only, pushed by the master, and contains no rules.

**People changes (v22/v24):** every rule-driven hours change is an **Adjustment** row in the OT Log with the calculation in Note, its type in Reason, shown in a confirm before saving, undoable by Void Entry.

## Engineering conventions
1. **Versioning.** Every module has `Public Const X_VERSION As String = "Name vNN (Mon d)"`; forms and ThisWorkbook have the `' @version` line; `modSystem.EXPECTED_RELEASE` is the release and `N_MODULES` the count. Bump every module that ships, and the expected release, in the same update. SystemCheck must say "N of N on vNN".
2. **Append-only log.** Never edit or delete OT Log rows. Corrections are voids (Status VOID, Void Note) plus a new entry. Never type hours into roster cells; they are formulas over the log.
3. **Verified data changes.** Any macro that restructures data: dated backup first -> change -> compare every roster period, Refused (I) and Total (G) value and the log's total Hours Charged before vs after -> on mismatch, `MarkNoSave` and tell the user to close without saving. Copy `UpgradeOTLogLayout` / `RunMigration` / `SetupHomeUnit`.
4. **Read every row, whatever the view.** Use `ListObject.DataBodyRange` or the roster's cell reads, never last-visible-row. Handle `DataBodyRange Is Nothing`.
5. **Columns.** Roster and log columns are constants in modOvertimeEntry (`C_*`, `L_*`); use them. For anything new, find by header text where the code already does that.
6. **Bulk work:** turn off ScreenUpdating, EnableEvents and calculation, and restore them in an error handler on every exit path. Capture the "was" values **before** `On Error GoTo`.
7. **Calendar dirty flag:** writes to Absences, Coverage, Postings or Shift Switches with events off must call `MarkCalDirty`.
8. **Late binding only** (Outlook, Scripting). No new references.
9. **Dates:** build with `DateSerial`; parse typed dates through the existing parser in frmShiftSwitch (handles "Oct 12" and year rollover). Never `CDate` a locale-ambiguous string. Cells that hold month or date **text** get `NumberFormat = "@"` first, or Excel converts them.
10. **Admin gating** via `ADMIN_USERS`; admin buttons are hidden for others (`ui_adm_` shapes).
11. **Flat, list-shaped tables** (one header row, stable names, no merged cells) so they can become SharePoint Lists later.
12. **Messages** tell a supervisor what to do next, in plain words.
13. **Names are an interface.** Public procedure names that buttons call (OnAction in `BuildCommandBar`) and macro names in install notes don't change. If a rename is unavoidable, `ApplyRosterLayout` must rebuild the buttons and the install note must say so.
14. **File format:** CRLF and ASCII only in `.bas` and paste-in `.txt` (`.gitattributes` enforces CRLF; write `->`, `-`, `...`). Every `.bas` starts with `Attribute VB_Name`. Every file has `Option Explicit`.

### VBA compile and runtime traps to check in review
- `Dim a, b As Long` makes `a` a Variant. Declare each.
- Max 24 line continuations per statement; 64 KB per procedure. Split big builders.
- Use `Long` for row counts, not `Integer`.
- Duplicate procedure names across modules -> "Ambiguous name". Search before adding a Public name.
- **A local or parameter with the same name as a procedure hides it** (names are case-insensitive): `tests/shadowcheck.py` finds these. Avoid helpers named `S`, `N`, `D`, `P`, `F`.
- `Lock`, `Unlock` and other reserved words can't be procedure names.
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
- Member workbook (M1, built): plant PCs, desktop Excel, everyone's hours visible, master pushes values into the template next to the master, published to `PUBLISH_FOLDER`, My history shows anyone, member file has no rules.
- Platform: stay on the Excel desktop master through Q1 2027; keep tables platform-ready.
- Release zips hold the full set from v22 on.
- `BACKUP_FOLDER = "."`: daily backups go into the roster's own folder (`BackupFolderPath` resolves "." and ".\sub"). Can be pointed at a network folder later.

## Still open (use the default unless Graeme answers)
| Question | Default |
|---|---|
| Boxing Day 2026 observed weekday | **Answered:** Monday Dec 28, 2026 for day workers. Entered on the Holidays sheet ("Observed on"), no code. |
| Network paths (`REPORT_FOLDER`, `PUBLISH_FOLDER`, `KRONOS_PATH` on P:) | Everything is built and tested locally on Graeme's desktop; the P: links are wired in at go-live. SystemCheck will flag them as unreachable until then. |
| Who runs the rollover / where the frozen copy lives | Settings constants, blank until given |
| Does Kronos export an employee number | Needed before v25 |
| Kronos worked hours after Jan 1 (YTD or cumulative) | Check the first January export (v23) |

## Roadmap
| Brief | Target | Theme | State |
|---|---|---|---|
| v20 | Oct 9 | Safety and stats | installed; gate passed |
| M1 | live Nov 15 | Member workbook + Publish | built, tagged `M1` |
| v21 | before Oct 31 | Provisional OT, DST, Home Unit, adjustment types, grade log, backup, SelfTest, printable roster | built, tagged `v21`, installed |
| v22 | mid-Nov | Sick Call, Leave going/returning, convert imported lines | `briefs/v22.md` |
| v23 | build by Nov 30, live Dec 27 | RolloverYear, Integrity check, docs | kit brief 04 |
| v24 | before Jan 16 | Employee pop-up, crew history, partial OT | kit brief 05 |
| v25 | Feb-Mar 2027 | Employee ID key, protection, calendar speed | kit brief 06 |
| v26 | Mar 2027+ | History report, staffing check, board print, payback aging | kit brief 07 |

## Working with Graeme
- Ask questions as multiple choice (AskUserQuestion), not in plain text.
- Start each update by reading the brief and every module it touches. Report anything that changes the plan, and wait for a go-ahead before writing code.
- If a brief conflicts with this file, ask.
