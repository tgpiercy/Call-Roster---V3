# The Oct 19-20, 2026 reviews: gaps, opportunities and where each one goes

Three reviews done against v26.5 (the Members M4 mock-up commit), kept here as the source for `briefs/v26.5.5.md` and `briefs/v27.md`. Every line number is from that commit. Decisions taken on Oct 19-20 are in CLAUDE.md ("Decisions already made", first bullet) and are not repeated as questions here.

- **Part 0:** the review of CLAUDE.md itself (what was fixed Oct 20, what is left: the split).
- **Part 1:** process gaps from the project's planning docs (Oct 19): the rules and jobs the forms cannot record yet.
- **Parts A-C:** the code review (Oct 20): three read-only reviewers on the VBA, data integrity / ease of use / testing and tooling; every HIGH item re-checked by hand.
- **Part 5:** making the outcome obvious: adoption and understanding for the transition.
- **Part D:** the release map.

Effort everywhere: **S** under an hour, **M** one update item, **L** a release on its own.

---

## Part 0. CLAUDE.md

Verified Oct 20: every claim about stamps, counts, forms, docs paths and checkers holds (25 stamps on v26.5, 649 cases, 663 pytest, four checkers clean with the three named false positives). Fixed the same day: the reconciliation date, the numbering rule restated to match practice, the SmokeTest row rewritten as the current 20 steps, the gating convention reduced to the current three levels, the timeline brought to its Oct 20 state, the shipped roadmap rows collapsed to one line, the checker false positives named, the Linux pytest line, `LoadCases`' size.

Still to do, in a session with no code change: split the file. `CLAUDE.md` (20-25 KB: what this is, layout, canonical rules, conventions, traps, tooling, how an update ships, working with Graeme, still open, the unshipped roadmap), `docs/workbook-map.md` (the module / form / sheet / Ref ID tables as they are NOW, one line each; the per-release narrative in the map rows is CHANGELOG content), `docs/decisions.md` (the decision log by date), `docs/settings.md` (every `*_FOLDER`, `*_USERS`, `*_DAYS`, `UI_SCALE`, `SIGNUPS_ON`, `DST_USE_ACTUAL_HOURS`, the posting constants: name, meaning, current value, who changes it; the go-live checklist and the Settings-sheet item both need it). The privacy paragraph becomes bullets. "Decisions already made" keeps one line per outcome and points at the brief for the rationale.

## Part 1. Process gaps (Oct 19, from the planning docs)

Confirmed Oct 19: no forced / mandated OT (illegal in Canada); personal Windows logins on plant PCs; a 16 h maximum consecutive hours rule; Left Message charged immediately. A partial-day absence already blocks both shifts (v25 D-C1).

### 1.1 Safety: the system cannot see fatigue
The rest rule reads the rotation only. It does not know what the person actually worked.

| Gap | Example | Fix | Effort | Goes to |
|---|---|---|---|---|
| **Accepted OT is not a shift** | A accepts 12 h Tuesday night; Wednesday's day call-out still lists A. | Eligibility treats accepted call-outs, awarded postings and switches as worked shifts: nobody listed whose OT would run past **16 consecutive hours** with their scheduled shift on either side, and nobody with less than 11 h between that OT and their next scheduled shift. One rule in `NotFreeReason`. | M | **v26.5.5** (decided Oct 20) |
| **Partial OT makes 16 h reachable** | Scheduled 12 h, then a 4 h partial is fine; 8 h is not. | The form shows each person's max hours for that shift and greys them when the hours on offer exceed it; partial awards use the same number. | S with the above | v26.5.5 |
| **Weekly cap** | Five OT shifts in a pay period. | A setting for max hours per 7 days (ESA 48 h unless an agreement exists): warn, never block, an override reason into the log Note. | S | v27 |
| **OT restrictions** | Modified duties, WSIB return-to-work, a doctor's note, "do not call for 30 days". | An **OT Restricted** status with reason, date range and whether charged while restricted; off the lists; "Restricted" on the form. | S-M | v27 |

### 1.2 Jobs the forms cannot record
| Gap | Now | Fix | Effort | Goes to |
|---|---|---|---|---|
| **Same shift, two units** | A person signs two postings for one date/shift, or accepts a call-out after being awarded elsewhere. | Award and Accept check Coverage for an existing assignment that date/shift and warn (the 16 h check catches most). | S | v26.5.5 |
| **Call attempt detail** | The session's outcome and timestamp only. | Per-person call time and method; "no answer" vs "left message" vs "declined". Grievances ask "what time did you call?". | S | v26.5.5 |
| **Unfilled shift** | Everyone refuses; the line just stays open. | A **Not filled** session outcome, so the gap is visible in reports. | S | v26.5.5 |
| **Award withdrawn** | The awarded person calls in sick, or the company cancels after awarding. | **Withdraw award** on a closed posting: voids the Awarded row, re-opens the line, optional re-award or call-out; records whether the person is still paid (CBA). | M | v27 |
| **Posting edits** | Hours, date or positions change after creation: cancel + recreate loses sign-ups. | **Edit posting**: re-draw, re-print, the old PDF marked superseded. | S-M | v27 |
| **Multi-shift postings** | A 7-shift vacation block = 7 postings and 7 sheets. | **Post all lines** from Book Time Off; one combined print. | M | v27 |
| **Emergency exceptions** | A plant upset: the nearest competent person is called, out of order, and logged wrong or not at all. | An **Exception** outcome: reason required, nobody charged, flagged in Needs Attention for the superintendent. | S | v27 |

### 1.3 Evidence for a grievance
| Gap | Fix | Effort | Goes to |
|---|---|---|---|
| **No sealed record** | **Period close**: each pay period, the OT Log, roster totals and postings to a dated PDF + CSV in a write-once folder, the CSV's hash in the PDF footer and in the next period's export. | M | v27 |
| **The paper sign-up sheet is the evidence** | A scan saved as `P-####.pdf` in a Scans folder, linked from the posting row, one click from the Award pop-up. | S | v27 |
| **Disputes are not tracked** | A **Disputed** status on a log row (who, date, outcome, note); Needs Attention lists open ones; the history report shows them. | S-M | v27 |
| **Rules are not tied to the CBA** | A **Rules** sheet: each rule with its CBA article or the date and by whom it was agreed; SelfTest cases reference the rule IDs. | S | v27 |
| **Config lives in code** | An admin **Settings** sheet with an effective date and a change log (see Part 0's settings table). | M | v27 |

### 1.4 Operations, communication, reports
| Item | Fix | Effort | Goes to |
|---|---|---|---|
| Kronos import sanity | Columns by heading (done); flag any person whose hours decreased or rose by more than 60 h since the last import; totals before / after with a confirm. | S | v26.5.5 (with A6) |
| Corruption-proof backup | Daily `ExportTables`: every table to CSV in a Backups folder; a rebuild without the workbook becomes possible; the first step of any platform move. | S | v27 |
| Roster hygiene | A Data check section in Needs Attention: missing or malformed phone / email, duplicate names, blank crew, a code that is not a grade. | S | v27 |
| Unit quirks | A per-unit Notes box on the call-out form ("Material Handler needs a forklift ticket"). | S | v27 |
| Open postings digest | One email (or Teams post) to the crew list when a posting is created; a weekly summary of what is still open. | S-M | v27 (with Members M4's digest) |
| Shift handover view | Everything logged in the last 12 h on one printable page. | S | v27 |
| Equalization report | Per unit: spread of Total Hours (min, max, mean, SD), refusal rate, first-called vs accepted. The union's number. | M | v27 Stats report |
| Monthly OT pack | OT hours and cost estimate per unit by cause, trend by month. | M | v27 Stats report |
| Posting outcomes | Fill rate, no-takers, unfilled, out-of-dept, days to award, by unit. | S | v27 Stats report |
| Switch ledger | Owed paybacks by age and person; a printable statement. | S | v27 |
| Replay the legacy Entry Log | Run the call-order rule over the old log and compare with what was called; each difference is a bug or an unwritten rule. | S | before Nov 15, with the Reconcile sheet |
| Data retention, certification expiry, vacation-to-demand forecast | Decide and document; build only if asked. | - | later |

## Part A. Data integrity (code review)

These can lose or miscount hours silently. All the S items ship in **v26.5.5**.

| # | Gap | What the code does | Fix | Effort |
|---|---|---|---|---|
| A1 | **A failed save is swallowed; "Session submitted" shows anyway** | `SaveIfWritable` (`modOvertimeEntry.bas:1335-1338`): `On Error Resume Next: ThisWorkbook.Save`. A locked file, SharePoint upload failure or dropped drive raises inside Save, is swallowed, and the call-out pop-up shows "Session S-0012 submitted". The rows exist only in memory. `BeforeClose` (`ThisWorkbook.txt:97-113`) does the same. `RolloverYear` already has the right pattern (`modRollover.bas:312-319`: FileDateTime before/after, "THE FILE DID NOT SAVE"). | Capture `Err` and compare `FileDateTime` before/after in `SaveIfWritable`; on failure a blocking "NOT SAVED to disk: do not close Excel; try File > Save or tell <admin>" and keep the form's entries. A red line on the Desk header while `ThisWorkbook.Saved` is False after a submit. | S |
| A2 | **No stale-copy detection; two writable copies both mint the same IDs** | `OpenChecks` (`modSystem.bas:483-506`) only reads Excel's ReadOnly flag. A OneDrive/SharePoint sync client on two PCs, or a copy taken home, both pass `RequireWritable`; both compute the next S-/P-/Entry ID from their own log (`:1717-1733`); last writer wins silently. | Record the file's disk mtime at `Workbook_Open` and after each own save; in `BeforeSave`, if the on-disk mtime differs, **cancel the save** (decided Oct 20): "Someone else saved this file at 3:12 AM. Close without saving, reopen, and log again." | S |
| A3 | **Esc / Ctrl+Break mid-write leaves half data with no save block** | `Application.EnableCancelKey = xlErrorHandler` exists only in the rollover (`modRollover.bas:224`). Esc during `AppendLogRows` (`modOvertimeEntry.bas:1759-1766`), `VoidRef` (`modOTLog.bas:475-481`), `RenameEverywhere` (`modPeople.bas:2861-2868`), `RemoveEmployee` (`:3042-3046`) or the Kronos H loop (`ImportKronosData.bas:215-265`) stops the code: no handler, no `MarkNoSave`, events and calculation left off, `BeforeClose` saves the half state. | Set `EnableCancelKey = xlErrorHandler` in `QuietOn` (the writers already call it) and restore in `QuietOff` and the handlers; `ReportError` logs Err 18 like any other. | S |
| A4 | **An error in Total Hours puts that person FIRST in the call order** | `LoadEligible` reads `.TotalHrs = NumOr0(v(i, C_TOTAL))` (`modOvertimeEntry.bas:499`); `NumOr0` returns 0 for `#NAME?`, `#REF!`, `#VALUE!` (`:1806-1812`). `RefusedThrough` is referenced as a bare name (`:905-910`) and not re-created on a read-only copy (`:888-898`), so a `#NAME?` is realistic. The person is listed lowest with no message. | In `LoadEligible` and `BuildRoster` (publish): if any row has `IsError` on G, I or C, refuse to load: "Total Hours shows an error for <name>: run RefreshRosterFormulas". | S |
| A5 | **Log row saved before the calendar; coverage can be skipped; nothing reconciles the two** | `SubmitCalloutSession` (`modOvertimeEntry.bas:1202-1212`): `AppendLogRows`, save, then `CalendarAddAccepted` per person, which opens a typed-number `InputBox` (`modCalendar.bas:1304-1327`); Cancel or a typo = "not put on the calendar". `WriteOutcome` then `CalendarOutcome` the same way (`modPostings.bas:1319-1332`). `VoidRef` releases coverage under `Resume Next` (`modOTLog.bas:486-496`). So the log can say Accepted while the line is OPEN (the Desk tile, the short flag and Needs Attention invite a second call-out for a filled shift, charging everyone again), or VOID while the line shows the person covering. IntegrityCheck's calendar check (`modIntegrity.bas:702-743`) tests only that a Source exists. | (a) Auto-pick the calendar line when exactly one OPEN line matches unit/date/shift, and pass the sick call's slot through `Prefill`; ask with a list only when ambiguous; on Cancel say in red "NOT on the calendar: open the Calendar and Set coverage". (b) IntegrityCheck "Coverage = log": every non-void Accepted/Awarded row needs a non-void line with that Source naming the person, and vice versa (PROBLEM). (c) `SubmitCalloutSession` asks when a non-void S- session for the same unit/date/shift was logged in the last 12 h. (d) `VoidRef` reports a failed release. | M |
| A6 | **Kronos export summed with no date window; H never cross-checked** | `PivotFromTransactions` (`ImportKronosData.bas:435-458`) adds every transaction whatever its Date; only the per-period side uses `PeriodStartFor`. An export covering Dec 27-31 or a wider range puts another year's hours into Worked and says "complete". Data-through = file date - 1 (`:174-177`): an export re-saved later moves "through" forward, so provisional OT stops counting shifts H does not hold. | Skip and report rows dated outside the roster year; compare each H with the person's in-year period sum (PROBLEM; also in IntegrityCheck "Kronos"); warn when file date - 1 is more than a few days past the newest transaction date; the sanity flags from 1.4. | M |
| A7 | **Kronos H written row by row: no backup, no verify, no NoSave** | `ImportKronosData.bas:215-265` writes `Cells(r, C_WORKED)` per row with events on; the handler (`:348-351`) only re-protects. An error at row 40 of 90 leaves half the roster on the new export, the stamps not moved, no Hours History row, and the file saves on close. People missing from the export keep OLD H while the provisional window advances, so their Total silently falls. No procedure writes H back from a snapshot. | `PeopleBackup("kronos")` + `RosterTotalsMap` / `TotalsDifference` on every column except H and G; H written from an array in one assignment; `MarkNoSave` on failure; a "stale since <date>" IntegrityCheck line for unmatched people; `RestoreWorkedFromSnapshot(id)` with the same verify. | M |
| A8 | **Worked Periods outside every "all tables" path** | `NameColumns` (`modPeople.bas:2663-2676`) lists nine tables; `tblWorkedPeriods` is matched by name (`ImportKronosData.bas:563`) and is absent; not covered by `RemoveEmployee`, the rollover or IntegrityCheck's orphans. Pending sign-up files carry `employee=<old name>` after a rename. | Add `AddNameCol ... WorkedPeriodsTable(False), 2, "Worked Periods"`; orphans include it; SmokeTest's rename step asserts the old name is gone from every `NameColumns` entry; `RenamePreview` counts inbox files under the old name. | S |
| A9 | **`ApplyRosterLayout` lifts a save block it did not set** | `FlipPeriodRows` ends with `ClearNoSave` on every success (`modUI.bas:247`). An admin whose rename / remove / edit failed (`MarkNoSave` at `modPeople.bas:2946, 3057, 3331`) and then runs ApplyRosterLayout (an install-panel step) can save the half-changed workbook. | `ClearNoSave` only when `mFlipStarted` was set in this call. | S |
| A10 | **Backups: admin-only, taken at open, 30-day retention, never proven, no restore** | `BackupOnOpen` runs only when `IsLiveRoster` = admin (`modSystem.bas:566-572, 629-656`); a master corrupted at 6 PM loses the day; `PruneBackups` deletes after 30 days (`:734-756`); no `Restore*` procedure. | **Daily backup on any writable open (decided)**; a `BeforeSave` backup when the newest is older than 4 h (hourly slot name, pruned at 7 days); the first backup of each pay period kept 400 days. `RestoreDrill` (admin): copy the newest backup to the Practice folder under a dry-run name, open it, run SystemCheck + IntegrityCheck + SelfTest, report "N log rows and these totals would be lost since the backup", one Install Log row; in every release's group 4. | M |
| A11 | **Rule-driving fields edited in place with no audit** | `SaveEmployeeEdit` writes phone, email, Plant Seniority (the tie-break) and home unit with no history row (`modPeople.bas:3300-3310`); `SetAbsenceEndTo` overwrites an absence's End (`modCalendar.bas:938-948`); `ReleaseCoverage` edits Covered By in place (`:1232-1238`). | A Field column on `tblGradeChange` (or a small `tblPeopleLog`) with old, new, who, when; the previous End kept in the absence Note; a VOID + new coverage line instead of an edit. | S |
| A12 | **Sign-up inbox trusts the file's own stamp and name** | `ReduceSignups` orders and tests "late" by the `stamp=` line inside the file (`modSignups.bas:78, 91, 257`), not its mtime; the `employee=` name has no tie to the login (`:303-321`); equal stamps resolve in `Dir$` order. | Read `FileDateTime`; a content stamp earlier than mtime - 10 min is "suspect" and named; flag a sign-up whose login differs from an earlier file for that name; sort by (stamp, file name). | S |
| A13 | **Members write into the master's folder during the beta** | `SIGNUP_FOLDER = ".\Signups"`, `PUBLISH_FOLDER = "."` (`modOvertimeEntry.bas:68, 83, 94`): every member PC needs write rights on the folder holding the master, its backups and the error log. | Point `SIGNUP_FOLDER` and `PUBLISH_FOLDER` at a separate members' folder now. | S |
| A14 | Small | `AppendLogRows` can say "Nothing was saved" after rows were written (`modOvertimeEntry.bas:1759-1784`: a duplicate session on re-submit); `VoidRef` has no handler or NoSave (`modOTLog.bas:460-500`); `RefusedThrough` follows the opening PC's clock with no bounds (`:888-898`); "on the roster twice" is a Check, not a PROBLEM (`modIntegrity.bas:211`). | One to five lines each. | S |

Checked and fine: void-twice refused; `UndoLast` picks the right row; `RenameEverywhere`, `RemoveEmployee`, `SaveEmployeeEdit`, `SavePlannedMove`, `WriteOutcome`, `SaveSwitch` all backup + verify + NoSave; the rollover is the model; IdMark counters; formula checks; read-only copies refuse every writer; a Kronos file imported twice is idempotent; the publish self-checks before any write.

## Part B. Ease of use (code review)

| # | Gap | What the code does | Fix | Effort | Goes to |
|---|---|---|---|---|---|
| B1 | **Between midnight and 7 AM the night in progress is dated yesterday** | Every pop-up pre-fills `DateText(Date)` (`frmOvertimeEntry.txt:219`, `frmSickCall.txt:118`). At 3 AM Tue a call-out or sick call for the shift under way, with Nights + the default date, books TOMORROW night. Two date formats in use (`MM/DD/YYYY` at `frmOvertimeEntry.txt:739`). | When Nights is picked the hint reads "Night of Mon Oct 19 (7 PM Oct 19 - 7 AM Oct 20)"; between 00:00 and 07:00 the date defaults to yesterday with "It is 3:10 AM: tonight's shift is dated Mon Oct 19"; one date format. | S | v26.5.5 |
| B2 | **The calendar is a dead end for a supervisor; the cover tile opens the wrong month** | `ShowLineMenu`: a non-admin double-clicking an OPEN line gets Details only (`modCalendar.bas:4967-4970`); no "Call out for this line" for anyone. `DeskCalendar` (`modDesk.bas:545-551`) opens the calendar with no date or unit; `CalShowDate(d, unit)` exists (`:4878`). | Supervisor line menu: "Fill this shift now (call-out)" through `Prefill` with the slot key, "Post this line", "Award posting P-nnnn". The cover tile calls `CalShowDate`. | M | v27 |
| B3 | **The printed roster has no phones for a supervisor and nowhere to log outcomes** | `Call_Out_Roster.bas:102-103` hides the phone column unless admin, though supervisors print it (`:24`). No A/R/L/U, time or session columns. | Phones for supervisors; tick columns + time per row; a "Session S-____ Shift ____ Hours ____" line per unit. | S | v27 |
| B4 | **A call-out session lives only in memory** | The form is modal (`modOvertimeEntry.bas:192`); Esc = "N unsaved entries will be lost" (`frmOvertimeEntry.txt:786-788`); `ReportError` unloads every form (`modErrorLog.bas:167-172`), so an error or a crash throws away every outcome so far. | A draft to a hidden sheet on every `RecordOutcome`; on the next open by the same login "Resume the Phosphine Nights Oct 19 session (4 of 12 logged)?"; cleared on submit. | M | v27 |
| B5 | **"Kronos data is old" stops a supervisor with a remedy only an admin has** | `KronosOkToContinue` (`ImportKronosData.bas:719-737`) on every LOAD LIST and CREATE POSTING past 7 days: Yes/No, No the default, "run Import Kronos first" (admin-only). | Non-admin: a red line on the form and carry on. Admin: keep the question, default Yes. | S | v26.5.5 |
| B6 | **"Why isn't Joe on the list?"** | `FilterNotFree` (`modCalendar.bas:755-756`) returns counts for two reasons only; names only for booked off and switches; clipped in a 9.5 pt label (`frmOvertimeEntry.txt:377-380`). | A "Not listed (N)..." link opening a box: name and reason, grouped; the same on the status-bar detail. | S-M | v27 |
| B7 | **The error box does not say whether the entry was saved** | `modErrorLog.bas:175-180`; the Desk's "Something went wrong" email is never offered from the box. | "Was it saved? Newest OT Log row: S-0012 by you at 3:14 AM / none" + "Email <admin> the details now?" | S | v27 |
| B8 | **Hours typed on every call-out; a sick call's hours dropped** | `txtHours` empty (`frmOvertimeEntry.txt:106`), LOAD refuses without it; `frmSickCall.txt:225` `Prefill unit, d, sh` passes no hours. | Default from `DefaultShiftHours`; add `hrs` to `Prefill`. | S | v26.5.5 |
| B9 | **A sick call saves with no confirm** | `frmSickCall.txt:206-212`. Undo stays admin (decided). | "Book Joe Smith sick for Mon Oct 19 Nights, Phosphine? They come off call lists until Wed 7 AM." The supervisor's message names the admin for an undo. | S | v26.5.5 |
| B10 | **Message audit** | Internal names shown ("(SUPERVISOR_USERS)" `modOvertimeEntry.bas:1272`; "modPostings" `modPostings.bas:2499`; "Alt+F8 > ..." `:189`, `modCalendar.bas:881`; `ByHandHint` tells a user to type into the VBA editor `modOvertimeEntry.bas:1407-1411`); "me" with no name (`Call_Out_Roster.bas:65`, `modPublish.bas:269, 280`); "Graeme" hard-coded (`modErrorLog.bas:178`); Yes/No with invented meanings (`modCalendar.bas:2914`, `modReports.bas:728-731`); a supervisor told to use NO TAKERS / CANCEL REQUEST which are greyed for them (`modPostings.bas:1253-1255`); "Days" as an unavailability reason (`:254-257`); 51 MsgBoxes with no title. | One text pass: plain words, `PRIMARY_ADMIN_NAME`, branch on `IsAdmin`, two named buttons instead of Yes/No meanings, titles everywhere. | S | v26.5.5 |
| B11 | **Award pop-up** | "REFUSED 21" for people who did not sign up (`frmPostingEntry.txt:301`); 8 pt hints; greyed buttons with no reason (`:427-441`); no keys. | "not signed up: charged 21"; 10 pt; a note under the greyed buttons; Space = signed up, A = award. | S | v27 |
| B12 | **The same unit has three names** | `modUI.bas:288-290` vs `modOvertimeEntry.bas:222-224` vs `modCalendar.bas:154-156`. | One `UnitDisplayName(unit)`; the stored value stays. | S | v27 |
| B13 | **The submit summary never says what to do next** | `modOvertimeEntry.bas:1217-1221`. | A last line by outcome: "Nobody accepted: the line stays OPEN. Next: call back later, post the shift, or ask <admin> to award out of department." | S | v26.5.5 |
| B14 | **Type sizes; `UI_SCALE` is a recompile** | Bar 8.5 / 7.5 pt (`modUI.bas:361, 382, 398`); hints 8 pt; `UI_SCALE` a `Public Const` (`modOvertimeEntry.bas:44`). | Floor at 10 pt; a per-login hidden name read by `UiScale()` with a Desk "Larger type" button. | S / M | v27 |
| B15 | **A plain user gets no Help, no Desk, no explanation** | `modUI.bas:296-322`, `modDesk.bas:434`; "Access denied ... (SUPERVISOR_USERS)". | A Help button for everyone; "Signed in as campeauc: supervisor" in the status cell; a first-open box pointing at Help. | S | v27 |
| B16 | **Employee history is five boxes in a row; Find a person answers on the status bar only** | `modReports.bas:719-741`, `:376-436`. | One small code-built form; Find a person writes to the roster's status cell and lists multiple matches. | S-M | v27 |
| B17 | **Members: "Not listed in the next 14 days" with no why; file age not flagged** | `modMember.bas:2121`, `:744-749`. | Publish `dNotFree` (name, date, shift, reason word) for 14 days; colour the Updated line past a threshold. | M | Members M4 |

## Part C. Testing, tooling and handover (code review)

| # | Gap | Found | Fix | Effort | Goes to |
|---|---|---|---|---|---|
| C1 | **IntegrityCheck has no row-level charge invariant** | The 11 checks (`modIntegrity.bas:145-158`) compare totals and IDs. Not caught: Hours Charged <> Hours x 1.75; a charged row with grade > 2; a charged Accepted row; Period Start <> `PeriodStartFor(Shift Date)` (load-bearing since v26.5); a VOID row with no Void Note; an Entry ID gap; overlapping Absences; a payback before the shift; a Former name on an open list; one person charged twice for one unit/date/shift. | A "Log rows" check recomputing the expected charge from Type / Outcome / Hours / Dept Code with `CalloutCharge` and the posting rule; the table checks. PROBLEM. | M | v27 |
| C2 | **The members' file reads published tables by column position** | `modPublish.bas:809-811` vs `modMember.bas:2179, 2210-2212`; nine sheet-name literals (`modMember.bas:341`); the sign-up format written (`modMember.bas:1917-1924`) and parsed (`modSignups.bas:190-240`) as two copies. | `ColOf(v, "Employee")` header lookups; `test_members_contract.py`. | S-M | Members M4 |
| C3 | **SmokeTest's end state is "totals came back"** | `modSmokeTest.bas:137-143`. No step runs `BringBack`, the planned-move apply, `UndoLast`, `PivotFromTransactions`, `BoardPrint`, the report draws, the masked PDF, the min-staff flag, a DST charge, a 27-period year. | Asserts at the end (log row count = before + rows logged; zero non-void rows under the test names; Check lines not up); steps for BringBack, apply, a 6-row transactions array, draw-only steps. | M | v27 |
| C4 | **`LoadCases` at 53 KB of the 64 KB limit** | `modSelfTest.bas:883`. | `LoadCases1..N` of 200 cases from the generator; a drift assert under 56 KB. | S | v27 (first) |
| C5 | **Pure rules trapped inside sheet readers** | `StaffCounts`, `PivotFromTransactions` (`CDate` at `:446`), `KronosHeadingProblem`'s same-row rule, `PublicWhy`, `SwitchSaveStops` and kin, `ScheduledSpan`, `UnitCodeOn`, `SyncPersonOnPostings`, `CalloutClosed`, `RefusedFormula`. | Extract cores; suites `staffcount`, `kronostx`, `switchstop`, `refformula`, `callclosed`. | M | v27.5 |
| C6 | **Five checker classes** | MACROSTR (60 `"Module.Name"` strings unguarded), ORNSWALLOW (70 procedures), MSGTITLE (51), PROCLEN over 300 (3), DEADPUB (2). | Add to `callcheck`; a `' never raises` marker. | S | v27 |
| C7 | **Install: the verification record is a MsgBox** | SystemCheck truncated at 700 chars (`modSystem.bas:256, 277`); the Install Log holds "OK" (`modUpdater.bas:1933-1934`); the backup never verified (`:347-385`); the template overwritten with no copy (`:741`). | Full text to the Install Log note; a `diff` step; `FileLen` + reopen read-only; `FileCopy` the template first. | M | v27 (updater change) |
| C8 | **Mirrored constants `test_drift` does not compare** | `HOLIDAY_NAMES`, `LEAVE_CODES`, `POSTING_CLOSE_HOUR`, `EXEMPT_TAG`, the sick window and shift starts, `PAYBACK_WARN_DAYS`, `SIGNUP_STALE_DAYS`, `POSTING_CLOSE_DAYS`; counts typed by hand in the steps files. | Six asserts; `{CASES}` / `{SUITES}` / `{MODULES}` tokens in `build_zip`. | S | v27 |
| C9 | **Docs headers lag** | Fixed Oct 20 by hand. | `tests/test_docs.py` so they cannot lag again; run in `build_zip`. | S | v27 |
| C10 | **Duplicated form-builder code across seven forms** | `AddCtl`, `AddLabel`, `FillCombo` x7, `C()` x9, `RowT` x4, `PickedDate` x2; four table finders; `AddValueCF` twice. | `modUI.FormAddCtl` and kin; one `TableByName`. | M | v27.5 |
| C11 | **Module splits** | modCalendar 5,540 lines (Board `2971-3181`, Min staff `3182-3527`, Import Old Cal `3528-3756`, Needs `3979-4773`, Menus `4774-5480`); modPeople (dates `652-834`, planned `1770-2657`, Former `1143-1322` + `2974-3429`); modPostings (mail `2417-2619`, sheet `505-794`, sync `1896-2416`). | Cut and paste, a `_VERSION` each, `N_MODULES`. | M | v27.5 |
| C12 | **No record of who opened the file or how errors trend** | `Workbook_Open` logs failures only; the Error Log's Release column is never counted. | An `Opens` table, admin-only, pruned at 90 days; SystemCheck adds "errors this release / opens last 7 days". | S | v27 |
| C13 | **No build manifest** | `build_zip` records the tag only. | `manifest.txt` with SHA-256 per file; the panel hashes the folder. | S | v27 |

## Part 5. Making the outcome obvious (adoption)

One test for every item: **after any action or on any output, can a person who never read the manual say in one sentence what happened and what it means for them?**

### 5.1 Every save ends with a receipt (decided: one layout, every form)
| Item | Change | Effort | Goes to |
|---|---|---|---|
| **One receipt card** | After a call-out, award, no takers, sick call, switch, time off, leave or edit: *What was recorded* (one line per row written, with its Ref ID) / *Who it changed* ("Joe Smith 120 -> 141 h, charged 21; Ann Lee 98 -> 98 h, not charged: signed up") / *Where to see it* (the calendar line, the OT Log row, the members' file after the next update) / *If this is wrong* (the admin's name; Undo last for an admin). Built once in modUI from a `Receipt` type; every save path's `Tell` goes through it. | M | v27 |
| **Printable and emailable** | "Print receipt" / "Email me this" on the card (the mailto fallback exists). | S | v27 |
| **Before/after on the roster** | For 60 s after a save the changed Total Hours cells carry a highlight and a note "was 120, +21 from S-0012". | S | v27 |
| **Outcome keys say their consequence** | "A Accepted", "R Refused (charged x1.75)", "L Left message (charged x1.75)", "U Unavailable (not charged)", "P Part of the hours (closes the shift)"; the award pop-up: "Signed up / Not signed up: will be charged". | S | v26.5.5 |

### 5.2 Every list shows its own reasoning
| Item | Change | Effort | Goes to |
|---|---|---|---|
| **Order columns on every list** | Grade code, Total Hours, seniority date, in that order, on the call-out list, the posting list, the printed roster and Who's next; "=" between tied rows; a footer "Order: grade, then lowest Total Hours, then earliest seniority. Hours as of <Kronos through> + <N> provisional." | S | v27 |
| **The total is clickable** | On the roster and the member card, Total Hours opens the pay-period breakdown the history report already draws. | M | v27 / Members M4 |
| **"What changed since the paper" column** | For the first two pay periods the printed roster carries the person's position on the last paper roster vs now, with the reason. Removal date in the brief. | M | v27 |
| **The trust strip** | One line on every output: "Worked hours through Sun Oct 18 (Kronos) | 3 provisional | Refused counted to the current pay period | Members' file updated Mon 3:12 AM." | S | v26.5.5 |

### 5.3 A parallel run that builds trust (decided: at least one pay period)
| Item | Change | Effort | Goes to |
|---|---|---|---|
| **Reconcile sheet** | Admin types the paper call order per unit once per period; the sheet lists every position that differs with the roster's reason; each line marked paper error / rule missing / data error. The Nov 15 pack's evidence. | M | v27, built BEFORE Nov 15 |
| **The old sheets as prints** | The Overtime Posting sheet and the Refusal Report keep their old layout for the first months, plus "Produced by the roster from S-0012 / P-0031". | S | v27 |
| **"On paper this would have been ..."** | When a charge applies a rule paper never had (partial OT, a planned move, the sick window, DST actual hours), one receipt line names the rule and its agreed date (from the Rules sheet). | S | v27, after the Rules sheet |

### 5.4 Fear-free learning
| Item | Change | Effort | Goes to |
|---|---|---|---|
| **Practice in one click, loudly marked** | A "Practice" button on the Desk for supervisors: the copy opens with a red banner on every sheet "PRACTICE COPY: nothing here counts", the title bar prefixed, the publish disabled; deleted on close. | S | v26.5.5 |
| **Guided first session** | The first three opens for a new login show a four-step strip on the Desk (log a call-out, book a sick call, post a shift, print the roster), ticking off in the practice copy; dismissible. | S | v27 |
| **The next thing, named** | Every Desk tile carries the verb that resolves it ("Fill it now", "Post it", "Award it"); a tile with no action is a report. | S | v27 (Desk v2's tile rule) |
| **"Was this right?" on every receipt** | Yes / Not sure / No with an optional line, into a `Feedback` table with the Ref ID, login and release; read from the Desk. | S | v27 |

### 5.5 Language and shape
| Item | Change | Effort | Goes to |
|---|---|---|---|
| **Old term -> new place glossary** | "Overtime Posting sheet -> New Posting / the printed posting sheet"; "Refusal Report -> Refusals Report and the Why column"; "the binder -> OT Log"; "call list -> Fill a shift now". A Help section and a one-page print. | S | v27 |
| **One status vocabulary** | One legend (word, colour, icon) for OPEN / POSTED / AWARDED / COVERED / NCN / CANCELLED and Accepted / Refused / Left Message / Unavailable / Awarded / Charged, printed on the calendar page, on the Desk and in the members' file. | S | v27 / Members M4 |
| **Numbers with units and basis** | "141 h (to Oct 18)", "21 charged (12 h x 1.75)": `ChargedText` and `TotalText` beside `HoursText`, used everywhere a number is shown. | S | v26.5.5 |

## Part D. The release map (decided Oct 20)

| Release | Takes |
|---|---|
| **v26.5.5** (before Nov 1, after v26 + v26.5 are on the real roster) | A1, A2, A3, A4, A8, A9, A13, A14; B1, B5, B8, B9, B10, B13; the 16 h rule + max-hours grey-out + same-shift check (1.1, 1.2); call attempt detail and Not filled (1.2); Kronos sanity flags (1.4, with the date-window part of A6 if the first real import has run); outcome-key labels, the trust strip, `ChargedText` / `TotalText`, the Practice button (5). Form re-pastes: frmOvertimeEntry, frmSickCall. |
| **v27** (about Nov 20-25) | The planned items (Desk v2, Stats, tiles, since-last-here, pop-ups); A5, A6, A7, A10, A11, A12; B2, B3, B4, B6, B7, B11, B12, B14, B15, B16; C1, C3, C4, C6, C7, C8, C9, C12, C13; 1.1 weekly cap + OT Restricted; 1.2 withdrawal, posting edit, multi-shift, Exception; 1.3 all five; 1.4 ExportTables, hygiene, unit notes, digest, handover, switch ledger; the Stats report absorbs equalization / monthly pack / posting outcomes; 5.1 receipts, 5.2 order columns + paper column, 5.3 all three (Reconcile first, before Nov 15), 5.4 guided session + tile verbs + feedback, 5.5 glossary + legend. |
| **Members M4** (one members' install with v27) | B17, C2, the clickable total on the card, the trust strip on Home, the legend. |
| **v27.5** (its own release) | C5, C10, C11. |
| **December** | `RestoreDrill` at the dry run; C7's install record. |
