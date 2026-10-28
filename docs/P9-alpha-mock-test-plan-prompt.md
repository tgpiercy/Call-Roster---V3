# P9 - The alpha's mock data set and testing plan (session prompt, written Oct 28, 2026)

Paste everything below the line into a fresh session. It briefs the session that builds the mock data set, the
test-only loader, the expected figures and the day-by-day testing plan for the alpha of Nov 1-14, 2026, which runs on
v29 with MOCK data on the real roster file. Decided by Graeme on Oct 28, 2026 (multiple choice): the mock set IS the
alpha (not a side pass on a copy); the data goes in through a test-only loader module; the mock plant has the full
shape (8 units, every crew, about 70 people); the test window uses the real calendar, Nov 1-14, 2026. The data reset
at the close-out takes the mock set off and puts the real roster back.

---

## Where things stand (read CLAUDE.md first; this block is what it does not say yet)

- **v29 is on the real roster since Oct 28, 2026** ("v29 executed well no issues"; main tip 7410d68, tag v29 = 4b6eb17).
  Nothing is installed during the alpha except an S1 hotfix (the roster is broken) as v29.0.x through a practice copy.
  **No roster code changes for this job**: the loader is TEST-ONLY (never in `modules/`, never counted by SystemCheck,
  packed under `test-only/` like `tests/snapshot/modSnapshot.bas`) and calls only Public procedures that exist in v29.
- **The alpha runs on MOCK people, hours and events** on the real file, under its real name, with the real supervisor
  logins. The plant's real call-outs run on paper for Nov 1-14 (paper governs, as planned); the real roster file is set
  aside before Day 0 and catches up from paper at the close-out (decision D1 below).
- Graeme executes every step by hand in Excel and ticks it against the figure this session computed. **Your expected
  figures must be right to the hour**: they come from `tests/rules.py` (the Python mirror of the pure rules), never
  from your head. Where the mirror has no function for a rule a step exercises, add the mirror function + `cases.csv`
  rows (and the SelfTest regeneration) - that is a test change, not a roster change.
- Ray (lanthierr) works from the Admin Lite Desk; three supervisors (mihaychukc, sheltond, campeauc) and Graeme
  (piercyg, admin) take the other chairs; members are played by the testers on the members' file (`PRACTICE -` is NOT
  used: the real paths are the point).

## Standing rules (CLAUDE.md "Working with Graeme" + "Engineering conventions")
Report back before code and put every decision to Graeme as multiple choice (the D-list below, defaults first). One
commit per deliverable. The four checkers + docscheck + pytest before every review and the zip. Read-only Opus
reviewers in isolated worktrees with named scenarios; fix, re-review. `git add` by file name, never `-A`. CRLF + ASCII
in `.bas`; `.md` files LF. Patch scripts through the Write tool (heredocs mangle backslashes). **Never commit data that
could be real**: every mock name, phone, email and Payroll # is made up (names from a word list, phones 905-555-01nn,
emails `first.last@example.test`, Payroll # `9` + 7 digits); `tests/test_members_contract.py`'s privacy screens must
pass on the mock set as they do on the real one.

## Deliverables (one commit each, in this order)

### A. The mock plant - `tests/mock/*.csv` (+ `tests/mock/README.md` describing every column)
The full shape: all 8 units (Phosphine, Derivatives, Cypure, Material Handler, QC Lab, QC Gases, Utilities, SBP),
crews A / B / C / D on the 28-day rotation and DAYS, about 70 people. Build it so EVERY rule in CLAUDE.md's canonical
section has at least one person and one date that exercises it in Nov 1-14:
- grades 1-3 in each unit (grade <= 2 charged; a grade 3 listed but never charged), two cross-coded people (counted in
  both units for staffing; out-of-department candidates for an award), Plant Seniority with at least two exact
  Total Hours ties broken by seniority and one by name;
- `people.csv` (roster row: Payroll #, Home Unit, name, crew, phone, Worked H opening, seniority, email, the unit
  codes), `crew_history.csv` (one dated row per person; one person whose crew changed Oct 20 so `CrewOn` matters),
  `log_opening.csv` (an `OB-` opening row per person with Hours Charged chosen so the call orders are interesting, plus
  about 25 historic `S-` / `P-` rows dated in periods 21-22 (Oct 4-31) with real outcomes so the history report, the
  Refusals Report, Stats and "Since you were last here" have content; `#` Entry IDs and Ref numbers consecutive, the
  `IdMark_` counters in `idmarks.csv`), `absences.csv` + `coverage.csv` (a vacation Nov 2-5 with open lines, one
  person on STD leave since Oct 5 returning Mon Nov 9 (`LV-` top-up on the return), two past sick occurrences for one
  person so a third in the window flips the attendance phrase, a personal day), `switches.csv` (one payback owed since
  Oct 10 = over `PAYBACK_WARN_DAYS`), `postings.csv` + `posting_list.csv` (one OPEN posting for Tue Nov 3 closing Mon
  Nov 2 with two sign-up files already in `Signups\`, one full-shift and one partial; one OPEN posting for Fri Nov 20 -
  the NEXT pay period, so its refusals land in period 24 and do not count in Total Hours until Nov 15:
  `RefusedThrough`), `planned_moves.csv` (one move effective Sun Nov 8 - the first admin open on or after it offers to
  apply), `restrictions.csv` (one active OT restriction Nov 9-13), `min_staff` = the plant's numbers (already in the
  roster; the loader leaves the Min Staff, Holidays, Rules and Settings sheets alone).
- Pay periods: 2026's period 23 is **Nov 1-14** (Dec 28, 2025 + 22 x 14 days); period 24 starts Nov 15. **Clocks fall
  back on the night of Sat Oct 31 -> Sun Nov 1, 2026: that night is 13 h** (`DST_USE_ACTUAL_HOURS`), so Day 1's sick call
  and call-out on it charge 13 x 1.75 = 22.75. **Wed Nov 11 is Remembrance Day**: DAYS is off and free for a call-out.
  Put both on the plan.
- `kronos/`: the mock Kronos data (decision D4 fixes the form): the cumulative "Totals" through Sat Oct 31 that gives
  every person their opening Worked H, then one per morning Nov 1-14 (a one-day report each, if D4 keeps the v27
  ledger path) with codes that exercise the multipliers (OVERTIME 1.75, OT - DT Split 2, LIEU OT 1.5, CALLIN OT 1.75,
  MIN 4 1.5) and one deliberate timecard EDIT on Nov 6 of a Nov 3 day, so `tblWorkedChanges` gets a row. Generated by
  `tests/mock/make_kronos.py` (openpyxl 3.1.5 is on `py -3.14`) in the layout `briefs/v27-kronos.md` section 1 describes.
- Member files for the scripted member steps are NOT pre-made: the testers make them on the members' file (sign up,
  withdraw, request time off, request a switch, the partner agrees). `tests/mock/member_steps.csv` lists each click
  with its expected file name pattern, so the plan can say what to look for in `Signups\` / `Requests\`.

### B. The loader - `tests/mock/modMockData.bas` (test-only; stamp `Public Const MOCK_VERSION`)
- `LoadMockData`: refuses unless the file is a practice copy (`IsDryRunName`) OR the admin types the word MOCK in an
  input box (the live file on Day 0); dated backup first (`SaveRestrictedCopy` into `Backups\` under a name containing
  "Backup"); then, events / screen / calc off with a handler that restores them: wipe the data tables (the OT Log,
  Absences, Coverage, Postings + list, Shift Switches, Crew History, Grade Log + GradeSnap, Planned Moves, Former,
  Restrictions, Requests, Feedback, Evidence, Closed Periods, Opens, Draft, Hours History, Worked Periods, Kronos
  Totals, Worked Days, Worked Changes, Kronos Files, Attendance's drawn sheet, the Error Log's rows) and the roster
  rows; write the people through `AddRosterRow`, the log rows through `AppendLogRows` (the OB- rows and the historic
  rows, newest first as the log is kept), the counters through `SetIdMark`, the other tables through their
  `ListObject` (`UnlockSheet` ... `ProtectSheet`, `MarkCalDirty`), `FillPayrollKeys`, `RefreshRosterFormulas`,
  `EnsureRefusedThrough`; the hidden names the mock needs (`RefusedThrough`, the Kronos through-dates) through the
  existing setters. **Verify** before the save: `RosterTotalsMap` and `LogChargedTotal` against `tests/mock/expected_load.csv`
  (G / H / I / C per person and the log's total, computed by C), and `RunIntegrity` with 0 PROBLEM; a mismatch =
  `MarkNoSave` + the box names the first person that differs. Nothing in it writes a formula by hand or edits a log row.
- `ResetRosterData`: the wipe alone (the same list, the roster rows too), verified empty, settings sheets untouched. It is
  for the practice rehearsal of the loader; the LIVE reset at the close-out is a file swap (E), not this macro.
- `MockMemberFile(kind, ...)`: a fallback that writes one sign-up / request file the way the members' file does (read
  `member/modMember.bas` for the exact format), for a step where no member tester is at hand. The plan prefers the real
  members' file.
- Checkers: vbacheck / shadowcheck / undeclcheck / callcheck run over `tests/mock/` too (add the folder to their walk
  only if they skip `tests/`; `test-only/` in the zip). `build_zip` packs it beside `modSnapshot.bas`.

### C. The expected figures - `tests/mock/expected.py` + `tests/test_mock.py`
- `expected.py` reads A's CSVs and the plan's steps (`tests/mock/steps.csv`: id, date, shift, who, action, inputs) and
  computes, with `rules.py` only (crews per date, `crew_on`, the call order grade -> G -> J -> name, `callout_charge`,
  the 16 h rule's `fatigue_max_hours` / `fatigue_block_core`, `dst_night_hours`, `attendance`, `refused_counts`, the
  leave / unit-move / new-hire top-ups, `staff_short`), every figure the plan asserts: the call order a pop-up shows,
  who is charged what, each named person's Total Hours before -> after, Worked H after each Kronos morning, the Lite
  Desk's four card counts and the three daily-check marks per day, Needs Attention's section counts, the attendance
  phrase, the posting outcomes, the Reconcile differences for the two paper orders, the Stats group A/B/C/D figures at
  the close-out. It writes `tests/mock/expected.csv` (step id, what, figure) and regenerates the EXPECT column of D.
- `test_mock.py`: the set is internally consistent - every scripted person exists, is on the crew the step assumes,
  is free (or deliberately not free, and the step says why), every date is in the window, the opening totals add up,
  the Kronos mock's weighted totals equal the Worked H the plan expects, the privacy screens pass on every mock table,
  and `expected.csv` is reproduced byte for byte by `expected.py` (drift).

### D. The plan - `docs/alpha-test-plan.md`
One page per day. Each step: **id** (`D3.2`), **when** (date, shift, before / after the morning import), **who** (login
and chair: Graeme admin / Ray Lite / a named supervisor / a member on the members' file), **where** (the button or the
pop-up, the sheet, the members' page), **inputs** (exactly what to type or pick, in order), **EXPECT** (the figures,
filled by C), **CHECK** (what to look at: the receipt's lines, the Total Hours cell's before -> after mark, the status
bar's Why, the calendar line, Needs Attention's line, the Lite card's count, the members' file after the publish, the
Integrity sheet), a tick box, and **if different** = a row in `docs/alpha-findings.md` with the step id (P1 triage).
- **Day 0 - Sat Oct 31:** set the real file aside (D1), `LoadMockData` typed MOCK, the Kronos opening import (D4),
  `AssignPayrollKeys` is NOT run (the loader stamps the keys), SystemCheck "45 of 45 on v29" + every path line,
  IntegrityCheck (0 PROBLEM; the expected Check / Note lines listed), SelfTest 1133 / 121, `PublishMembers` (M6) and the
  members' file opened by each tester once, the two carried-in sign-up files seen on the Nov 3 posting at LOAD, the
  Desk for each chair (Ray's Lite cards: requests 0, exceptions 1 - the Nov 3 posting closes Nov 2 = "closing soon" -,
  corrections 0, returning 1), the Opens row.
- **Days 1-14:** every morning the Kronos import (EXPECT: three named people's H, the pill's state), the three
  daily-check marks going green on Ray's Desk, Needs Attention's counts, the publish; then the day's scripted events.
  The skeleton below is the minimum (the session places the people and exact hours so the orders and charges are
  interesting; one rule per step, so a wrong figure points at one function):
  - **D1 Sun Nov 1** (the 13 h night): a sick call on the Oct 31 night (`frmSickCall`, the sick window), the call-out it
    opens (Refused, Left Message, a PARTIAL Accepted that closes the shift: 22.75 charged above, nothing below), the
    receipt + "Was this right?" No on purpose (a Corrections row for Ray), the session draft (close the pop-up mid-way
    once and reload it).
  - **D2 Mon Nov 2:** the Nov 3 posting closes: Ray awards it from the Lite card (the full-shift sign-up beats the
    partial; the refusers ahead charged the full hours; `AwardOrderProblem` once by trying the partial first); a member
    request for time off arrives (members' file) -> Ray books it from Needs Attention (`RequestBooked`); a member
    switch request + the partner's Agree -> booked; a vacation already on the calendar makes a person not free.
  - **D3 Tue Nov 3:** Add a person (`NH-` at the home unit's average), a shift switch with a payback (the 16 h stop
    once on purpose, then a legal one), Fix an entry on D1's No (void + re-enter: the request status, the Feedback row
    resolved, the Lite count back to 0).
  - **D4 Wed Nov 4:** the 16 h rule: a person coming off a worked switch is OFFERED their maximum ("max N h"), refuses,
    is charged on it; a NO TAKERS posting (everyone on the list charged); an out-of-department award.
  - **D5 Thu Nov 5:** Edit a posting (hours, then the date) and Withdraw an award; Cancel a posting (nobody charged;
    the Reason stays on the line); the weekly cap is OFF (nothing warns).
  - **D6 Fri Nov 6:** the Kronos EDIT row appears in Worked Changes; a planned move entered for Nov 8; a time-off
    booking trimmed to scheduled shifts; `ShiftHandover`, `SinceLastHere`.
  - **D7-D8 Sat-Sun Nov 7-8:** the same-day sister swap (EXEMPT from the 16 h arithmetic; a two-way payback on the
    adjoining shift stopped); the rest rule (a crew coming off nights not listed for the day call-out); Nov 8's first
    admin open OFFERS the planned move (Yes: `UM-` top-up, the lists follow); below-minimum staffing shown and the
    Board printed.
  - **D9 Mon Nov 9:** the leave return (`LV-` = the home unit's average gained since Oct 5; `ComputeLeaveReturn`); the OT
    restriction starts (not listed, never charged); a rename (every table follows, by key).
  - **D10 Tue Nov 10:** a member withdraws a sign-up then signs up again (the newest wins; the suspect flag stays
    clear); a request declined with a reason (never published); `RequestBookingVoided` by voiding a booked absence,
    then re-booked.
  - **D11 Wed Nov 11 (Remembrance Day):** DAYS is off and free: a DAYS person on a call-out list and accepted; the
    holiday's observed date on the Holidays sheet; Needs Attention's HOLIDAYS line.
  - **D12 Thu Nov 12:** Reconcile: two paper call orders typed (one agreeing, one with a planted difference) ->
    Compare -> the PDF; the Refusals Report; the Employee history (member-safe and FULL).
  - **D13 Fri Nov 13:** Remove a person (to Former) and Bring back; an adjustment (`ADJ-`) and Undo last; IntegrityCheck
    17 (a booked request whose Ref was voided = Note once past); the ID counters.
  - **D14 Sat Nov 14:** the period's last day: the Nov 20 posting's refusals still OUT of Total Hours (period 24); the
    Paybacks ledger; Stats drawn mid-way; every tester's Feedback answered.
- **Coverage matrix** at the end: every canonical rule (CLAUDE.md) x the step ids that exercise it; a rule with no
  step is written as such (the rollover, year-end and the guided move are NOT in this plan - December).
- **The close-out, Nov 16-17 (E below)** as the plan's last page.

### E. The data reset - the last page of D, plus `docs/alpha-data-reset.md` (one page, by itself)
- **Day 0, before anything:** the REAL roster is closed, copied, and the copy named so nothing treats it as live and
  nothing prunes it (verify `IsBackupCopyName`, `IsLiveRoster`, `PrunePreChangeCopies` and `PruneBackups` against the
  name and folder you choose; default: `Pre-alpha\Backup - PRE-ALPHA 2026-10-31 - V2 - Call Out Roster.xlsm`, a
  subfolder beside the roster: the prunes look only at the roster's own folder and `Backups\`). The real members' file,
  `Signups\` and `Requests\` of that moment are copied with it (the mock run reuses the folders).
- **Nov 16-17:** on the MOCK file: the mock cumulative Kronos through Nov 14, IntegrityCheck, `ClosePeriod` Nov 1-14
  (its CSV + PDF + hash row are the alpha's evidence, labelled mock), Stats, the Feedback rows, every plan tick and
  every alpha-findings row triaged. Then the swap: the mock file renamed with "Backup" in its name
  (`Backup - ALPHA MOCK 2026-11-14 - ...`) into `Pre-alpha\`; the pre-alpha copy renamed back to the live name;
  opened (SystemCheck, IntegrityCheck); **caught up from paper**: every real Nov 1-14 event entered as real rows dated
  on its day, in order (call-outs through the pop-up with the real outcomes, postings, sick calls, switches), the real
  cumulative Kronos report imported, IntegrityCheck 0 PROBLEM, `RolloverProblems` glanced at, `PublishMembers` (the real
  members' file is back, M6), the testers' member-file sign-ups from the mock run are NOT in `Signups\` (the folder
  was restored). The runbook names who does what and the box each step should show.

## Report-back decisions to put to Graeme (multiple choice, defaults first)
- **D1 The real file during the alpha.** Default: set aside Oct 31 under a Backup name in `Pre-alpha\`; the plant's
  real call-outs run on paper; the real file is caught up from paper Nov 16-17. Alternatives: the real file keeps
  running on a second PC for the real events (two live rosters: the stale-copy guard will bite - say so); or the
  real events are entered into the MOCK file too and copied over by hand (no).
- **D2 Logins.** Default: the real logins in their real chairs (Graeme admin, Ray Lite, the three supervisors); the
  member testers are those five on the members' file (the sign-up files carry their Windows logins - fine on a mock
  set). Alternative: Graeme alone in every chair (the login-dependent steps then say "as <login>" and he switches PCs).
- **D3 How many people.** Default about 70 across the 8 units in the plant's real proportions (`min_staff` needs
  Phosphine >= 2 per shift etc.). Alternative: 40 (one crew thinner; staffing flags fire more).
- **D4 Kronos.** Verify first what `KRONOS_PATH` points at on the real roster (the constant names the processing
  workbook `Kronos - Raw Data Processing.xlsm` = the pivot-sheet fall-back; v27's folder / wildcard form reads the
  "Employee Transactions and Totals" reports). Default: the mock data as ONE-DAY REPORT .xlsx files each morning plus
  the cumulative on Day 0 and at the close-out, placed where `KronosFilesToImport` looks - which means a
  `KRONOS_PATH` that names a folder on the real roster; if today's constant names the pivot workbook, the choice is
  (a) a mock pivot workbook put at that path for the fortnight (cumulative only; the ledger path untested; the real
  reports go on landing beside it), or (b) a v29.0.1 whose ONE change is `KRONOS_PATH` = the mock folder, installed
  Oct 31 through the panel (one constant; the pre-release rule allows a hotfix during the alpha only for S1 - this is
  not one, so it needs Graeme's explicit yes), or (c) the loader writes the Kronos tables directly from the CSVs (no
  import code exercised: last resort). Ask.
- **D5 Members' file.** Default: published beside the master as today (`PUBLISH_FOLDER "."`), the testers open it from
  there; the mock run's `Signups\` / `Requests\` are the real folders, emptied on Day 0 after the Day-0 copy. Alternative:
  the folder split rehearsed now with `MoveRoster` into test folders (that is December's job; no).
- **D6 The paper side.** Default: the mock plant has NO paper; the two Reconcile orders on D12 are typed from the plan
  (one planted difference). Alternative: the real paper of a real day typed against the mock roster (meaningless; no).
- **D7 What Ray tests.** Default: Ray's steps are the Lite Desk's cards end to end (request -> book, exception ->
  award, correction -> Fix an entry, returning -> leave return), the daily check and `MovePathLines` on SystemCheck;
  Graeme's are everything else. Alternative: Ray shadows every step (the plan then marks "Ray watches").

## Done means
- `tests/mock/` (A, B, C) committed, `py -3.14 -m pytest tests -q` green with `test_mock.py` in it, the four checkers
  + docscheck clean, `build_zip` unchanged for the roster but packing `test-only/modMockData.bas` next time.
- `docs/alpha-test-plan.md` with every step's EXPECT filled from `expected.csv`, the coverage matrix, the close-out page;
  `docs/alpha-data-reset.md`; `docs/alpha-findings.md`'s "Sources" line gains "a plan step id"; CLAUDE.md's roadmap
  gets one row (the alpha mock set, test-only) and the timeline line says the alpha runs on mock data; NEXT-SESSION-
  PROMPT.md's NEXT list points at the plan.
- Two read-only Opus reviews (one for the loader's write paths + the wipe list + the guards against running on the
  wrong file; one for the expected figures: the reviewer re-derives five steps by hand from the rotation and the
  CSVs, incl. the 13 h night and a 16 h cap) and their fixes.
- A rehearsal of the loader on a PRACTICE copy by Graeme before Oct 31 (the plan's "Day -1"): `LoadMockData`, the
  verify box, SystemCheck, IntegrityCheck, the Desk for each chair, `ResetRosterData`, then `LoadMockData` again - the
  same figures twice.
- Nothing in `modules/`, `paste-in/` or `member/` changed (git shows it); SystemCheck still says 45 of 45 on v29.
