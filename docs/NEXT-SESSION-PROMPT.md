# Paste this into a fresh session to carry on (updated Oct 28, 2026, later: **v29 is TAGGED on 4b6eb17, ZIPPED, PUSHED and INSTALLED on the real roster the same day ("v29 executed well no issues"). THE ALPHA (Nov 1-14) RUNS ON v29.** An alpha S1 is a v29.0.x hotfix; S2 / S3 rows go to docs/alpha-findings.md for v29.5 (or v28.1's slot). v28.1 (the second refactor) may now be built on v29 - report-back first, never installed during the alpha. See "v29 STATE" first; the older blocks are history.)

```
v29 STATE (Oct 28, 2026): BUILT on main in eight item commits (2cac2bf, e5ffc90, 412dead, 9a61bde, 2ebd706, e6def18, 8fceb66 + the
report-back 080daaa), reviewed by four read-only Opus reviewers in isolated worktrees (fixes 705f05d, 674844e), re-reviewed by
two fresh ones (fixes ddcb923), read once more (1 MED / 6 LOW, taken in the release-prep commit), then release-prepped (stamps v29 on all 45 modules, EXPECTED_RELEASE v29, the
What's New box ON with four lines, FORM_RELEASES frmFixEntry / frmTimeOff / frmShiftSwitch v29, INSTALL-v29.md, steps-v29.txt,
the CHANGELOG's v29 entry incl. the review record, CLAUDE.md's rows, the four end-user docs' counts, the Desk guide, NEW
docs/admin-lite-cheat-sheet.md + docs/guide-08-admin-lite.md + docs/members-what-changes.md). Two NEW modules modLite + modMove
(N_MODULES 45), ONE NEW UserForm frmFixEntry (Insert > UserForm BEFORE the panel: modLite names it), two forms re-pasted,
ThisWorkbook unchanged, Members M6 (template rebuilt at install). The brief said N_MODULES 43 - two modules were cleaner.
Tests: 1133 cases / 121 suites, pytest 1196, SmokeTest 22; checkers + docscheck clean.
WHAT v29 IS: briefs/v29.md (items 1-6, 9, 10 built; 7 = the alpha's S2/S3 fixes NOT built - after Nov 18 as a v29 item or
v29.5; 8 = item 39 NOT built - unsigned) + briefs/admin-lite.md. Graeme's 12 decisions in the brief's table; built ON v28.
KEY DESIGN FACTS: the move stores the members' folder IN the file (hidden name MembersFolder "L|path" live / "R|path"
rehearsal; PublishFolderPath reads it; a COPY honours only R); MoveRoster = Alt+F8 admin, a CUTOVER step (Jan 11-15), never an
install step; the frozen old file (MovedTo) never saves; IntegrityCheck 17 Requests never PROBLEM; the void reason is NEVER
published (Answer = 'booking voided'); the Lite Desk draws on the FULL admin view (his call).
DONE Oct 28 (later): tagged v29 on 4b6eb17, dist/Callout_Roster_v29.zip built, main + tag pushed (e446b42..4b6eb17), INSTALLED
on the real roster by Graeme the same day, no issues. THE ALPHA RUNS ON v29.
NEXT: (0) FIRST, before Oct 31: the P9 session - docs/P9-alpha-mock-test-plan-prompt.md (Graeme, Oct 28: the alpha runs on a
MOCK data set loaded by a test-only module, full plant shape, real calendar; expected figures from rules.py; a data reset at
the close-out) - report-back D1-D7 first, then build; (3) the alpha Nov 1-14 on v29 (P1 triage into docs/alpha-findings.md; S1 -> a v29.0.x hotfix through a practice copy);
(4) the alpha close-out Nov 16-17 on v29 (ClosePeriod Nov 1-14 on v29); the Nov 18-20 install slot is free - v28.1 if its
review is clean and Graeme says so, else nothing; the S2/S3 fixes become v29.5 (or ride with v28.1) - decide with him;
(5) v28.1 (the second refactor) built ON v29 (v29 has run): report-back by the P2 pattern, build during the alpha, NO install in it;
(6) docs/december-rehearsal.md (P6) and docs/go-live-runbook.md (P7, by Jan 5; it must say: rename the old roster and
members' file MOVED, the IT rights, the two test folders for the rehearsal).
```

```
v29 REPORT-BACK DONE (Oct 27, 2026, later in the day): Graeme said "Lets pull ahead 29" and chose: v29 NEXT, built on v28 (v28 has run), installed in
the Nov 18-20 slot after the alpha closes; v28.1 (the refactor) BEHIND it, built on v29, installed before the Dec 20 freeze; training Dec 1 and the
December rehearsal on v29. He pasted briefs/admin-lite.md (saved as sent; read "v28.1" there as v28). Three read-only maps (folders; logins + Desk;
Requests + IntegrityCheck) -> twelve decisions, all in briefs/v29.md's table. THE ONES THAT CHANGE THE PLAN: the move STORES the members' folder
path in the file (hidden name MembersFolder read by PublishFolderPath; constants otherwise; no January install); the roster drives the move, the
reopen is by hand (Excel cannot open two same-named files); a rehearsal mode on a COPY (a practice copy always publishes beside itself today);
the old file frozen by a MovedTo marker + a rename; THE LITE DESK DRAWS ON THE FULL ADMIN VIEW (not desk mode - his call); all four sourceless
Lite cards built; Fix an entry in v29 (frmFixEntry, a NEW code-built form); SUPERVISOR_USERS stays the three; a void marks its Booked request
"Booking voided" (new final status, Members M6); check 17 "Requests" = Check while the dates are ahead, Note once past, never PROBLEM; the two
RequestBooked gaps fixed; My history's PDF to the user's Downloads. BUILD ORDER (one commit each, checkers + docscheck + pytest after each,
read-only Opus reviewers with the P5 scenarios, fix + re-review): item 2 logins/ViewModeFor lite/SystemCheck login lines -> 3a the card
sources -> 3b DrawLiteDesk (own Sub: DrawDesk is 284/300 lines) + 3-state preview + DeskShowFull + the landing -> 4 Fix an entry -> 5 Needs
Attention for Lite -> 6 Requests -> 1 the move (MembersFolder, MoveRoster + rehearsal, MovedTo freeze, SystemCheck path lines reach+write,
wording, IntegrityCheck 15 Note) -> 7 alpha fixes after Nov 18 -> 8 item 39 if signed -> 10 release prep (stamps v29, N_MODULES 43, What's New
YES, FORM_RELEASES + frmFixEntry, INSTALL-v29.md incl. the new UserForm + the members' template rebuild, steps-v29.txt, CHANGELOG, docs incl. NEW
admin-lite-cheat-sheet / guide-08-admin-lite / members-what-changes, the snapshot steps admin + supervisor-via-preview + Lite); tag v29;
build_zip v29 v28; push from his Terminal. MoveRoster is NOT run at the install - a cutover step (docs/go-live-runbook.md by Jan 5). The build
starts on his go (asked at the end of the report-back session). Scratch maps of the three agents are NOT in the repo (session scratchpad).
```

```
v28 INSTALLED (Oct 27, 2026): Graeme: "V28 executed flawlessly." v28 is on the REAL roster; v27.5 rode in; v27.1 has been on it
since Oct 25. Every check passed: SystemCheck 43 of 43 on v28, SelfTest 1053 cases / 113 suites, SmokeTest 22 steps, the group 4
rows, AssignPayrollKeys clean, the members' template rebuilt (Members M5). In code: ADMIN_USERS = piercyg, lanthierr; DESK_USERS =
lanthierr; SUPERVISOR_USERS = mihaychukc, sheltond, campeauc. NOT in v28 (the go-live kit, now v29): the folder split, the final
logins, Ray's Admin Lite Desk. CLAUDE.md's roadmap and "Decisions already made" carry the Oct 27 decisions (the first bullet).
THE CHAIN (as first planned Oct 27; RE-ORDERED later that day - see the v29 block above): v28 (real, Oct 27) -> alpha Nov 1-14 -> v29 (the
go-live kit + the alpha fixes, installed Nov 18-20) -> v28.1 (the refactor, built on v29, installed Dec 14-18 if clean) -> the real rollover
on v29 in the freeze -> the cutover Jan 11-15 -> v29.5 (Feb). There is NO v28.5.
THE CALENDAR: Oct 27-31 housekeeping (this commit) + the alpha gate re-run on v28 (Graeme's runbook, project-side) | Nov 1-14 the
alpha on v28: paper governs, the daily admin check (Ray when Graeme is on nights), findings triaged into docs/alpha-findings.md as
they come (S1 = broken -> a v28.0.x hotfix through a practice copy; S2 = wrong hours / charge / order -> void and re-enter now, the
code fix in v29; S3 = confusing / slow / cosmetic -> v29 or later); NO other install during it | Nov 2-6 the v28.1 report-back
(maps by read-only agents: modMember's sections and cut lines, the 116 Consts' readers, every by-name lookup and where KeyMatchCore
goes, TableByName vs TableData; multiple choice; NO code until his go) | Nov 6-13 v28.1 build + review + tag (v27.1's method:
splitmodule.py + movecheck.py, R3 its own commit with cases + a Python mirror, builders in isolation worktrees, five read-only
reviewers, modSnapshot before / after; stamps v28.1, no What's New box, INSTALL-v28.1.md + steps-v28.1.txt, build_zip v28.1 v28;
ask before the push; NOT installed until the alpha closes) | Nov 9-13 the v29 report-back part A (the kit; save his updated
brief-admin-lite.md into briefs/admin-lite.md first) | Nov 16-17 the alpha close-out on v28 (the cumulative Kronos report ->
import -> IntegrityCheck -> ClosePeriod Nov 1-14 -> Reconcile -> Stats -> Feedback) | Nov 18-20 v28.1 installed (practice copy with
the snapshot, then the real roster; SystemCheck "N of N on v28.1") | Nov 18 the v29 report-back part B (the S2 / S3 rows) ->
briefs/v29.md | Nov 20 - Dec 4 v29 build + review + tag on v28.1 | Dec 1 supervisor training on v28.1 | Dec 4-7 write
docs/december-rehearsal.md | Dec 7-11 the rehearsal on a v29 practice copy (install, rollover dry run, RestoreDrill, Ray's first
week, the guided move into test folders); CUT-OFF Dec 12: a fail moves v29 to Jan 5-8 and the rollover runs on v28.1 | Dec 14-18
v29 on the real roster, folders at "." | Dec 20 - Jan 4 the freeze, the real rollover (2027 period 1 starts Sun Dec 27) | Jan 5-8
the first January Kronos export (year-to-date or cumulative?) | by Jan 5 docs/go-live-runbook.md | Jan 11-15 the cutover | Jan
16-22 no installs, Ray's solo check | Jan 23-30 Graeme away | Feb v29.5.
STILL OPEN (defaults in CLAUDE.md's Still-open table): G1 the go-live date (Jan 11-15); G2 the P: paths + members' write rights
from IT by Nov 20; G3 who else goes on SUPERVISOR_USERS (the four shift supervisors only); G4 paper past Nov 14 (stop after the
pack if the Reconcile differences are explained); G5 item 39's wording (not built until signed); G6 signing the VBA project (ask IT
in November). Ask each ONCE, as multiple choice, when its phase comes.
NEXT SESSION: alpha triage as he brings findings (no code; file:line, roster-or-paper against the canonical rules, S1 / S2 / S3,
a row in docs/alpha-findings.md; an S1 stops and puts the hotfix choice to him), and from Nov 2 the v28.1 report-back.
```

```
v28 STATE (Oct 26, 2026): read docs/v28-report-back.md (Graeme's decisions: build now - the pre-release rule waived -,
scope = members' requests + the Payroll # key, the phone page CANCELLED, go-live facts as placeholders; admins only
act on a request; a switch request needs the partner's Agree in the file first; the key carried everywhere with the
rename by it, the SUMIFS by name; a new hire's number typed at Add), briefs/v28.md ("As built"), the CHANGELOG's
v28 entry (the design + the review record), docs/INSTALL-v28.md and briefs/steps-v28.txt. On main: modRequests
(NEW) + the members' file's Requests page (Members M5, the template rebuilt at install), modKeys (NEW), 43 modules,
ThisWorkbook v28 (open step 10 = RequestsOnOpen), frmTimeOff / frmShiftSwitch (Prefill) / frmEmployee (Payroll #) v28,
AssignPayrollKeys = the run-once in group 3. Two read-only reviews took 7 HIGH / 8 MED (5c63a9f; the CHANGELOG lists
them - the biggest: a 21-column log deadlocked behind the layout guard, now LogTable() adds the key column; 24
module-level Consts after procedures in the members' module = a compile error vbacheck could not see behind an
open one-line Sub - fixed, and vbacheck now closes one-liners: a clean run says OK). 1053 cases / 113 suites, pytest
1116, SmokeTest 22, checkers + docscheck clean.
DONE Oct 26: the re-review's 2 MED / 10 LOW taken (the commit after 5c63a9f); TAGGED v28 (annotated) on that commit;
ZIPPED dist/Callout_Roster_v28.zip. main + the tag PUSHED to origin from Graeme's Terminal on Oct 27 (83fef17..f490805). NOT installed.
NEXT: (b) Graeme's INSTALL - the ORDER: v27.5 first (or straight to v28: v27.5 rides in), the members' template rebuilt by
hand (INSTALL-v28.md), AssignPayrollKeys once; the freeze Dec 20 - Jan 4; (c) the alpha Nov 1-15 -> v27.5.5 (fixes
found at the v27.5 install go there, never alone); (d) item 39 (the email digest) when he gives the wording; the
December rollover dry run; the Reconcile differences for the Nov 15 pack. Follow-ups noted in briefs/v28.md: an
IntegrityCheck line for a Booked request whose Ref is void; the lookups (call order, lists, calendar) still match by
name - a later release may switch them to the key.
```

```
v27.5 STATE (Oct 25, 2026): read docs/v27.5-report-back.md (Graeme's decisions incl. the two taken AFTER the reviews),
briefs/v27.5.md, the CHANGELOG's v27.5 entry (the design + the review record), docs/INSTALL-v27.5.md and
briefs/steps-v27.5.txt. Three rule fixes on main: (1) every switch runs the 16 h stop by the rule's own arithmetic
(modCalendar.FatigueShiftReason; a scheduled shift beside the one worked counts; the shifts a switch gives up are set
aside through SetPendingGiveUp) - BUT the same-day sister swap of BOTH shifts is EXEMPT from the arithmetic (Graeme:
"the only time an A/B pair can swap"; only "on OT" stops it); (2) a missed Kronos morning writes "gap" rows at the
year-to-date import, counted in Worked Periods per person + pay code until that day's rows arrive; the one-day reports
of a run are taken before its year-to-date; (3) the staffing count adds partial covers once ("; " and " / " parts),
rounds down, says "(part)"; a switch saved over the warning counts its worker. Review record in the CHANGELOG: item 2
had a HIGH (SerialSet called itself - callcheck SELFCALL now flags a Function's own name with an argument list on the
left of =), item 3 two HIGHs (the switch-giver no-op, the "; " join). 1010 cases / 109 suites, pytest 1066, SmokeTest
22, checkers + docscheck clean. Every stamp v27.5, EXPECTED_RELEASE v27.5, the What's New box shows (a rule change).
DONE Oct 25 (late): the re-review's 2 MED / 6 LOW taken (1814e88); TAGGED v27.5 = 1814e88 (annotated); ZIPPED
dist/Callout_Roster_v27.5.zip (67 files, sent to Graeme); main + the tag PUSHED to origin from his Terminal. NOT
installed yet.
NEXT: (a) his install on a practice copy, then the real roster (steps-v27.5.txt group 4: the one-way switch stops, the
same-day swap does not, the What's New box shows once) - when he says "it doesn't work", ask which file and version is
on screen first; fixes go in v27.5.5, never alone; (b) the alpha Nov 1-15 -> v27.5.5; item 39 (the email digest) when
he gives the wording; the Reconcile differences for the Nov 15 pack; the December rollover dry run on v27.5.
```

```
v27.1 STATE (Oct 22, 2026): read docs/v27.1-report-back.md (the plan, the maps, Graeme's eight decisions), the
CHANGELOG's v27.1 entry, docs/INSTALL-v27.1.md and briefs/steps-v27.1.txt. All code is on main (local, not pushed):
11 new modules split out of modCalendar / modPeople / modPostings by tests/splitmodule.py and proved by
tests/movecheck.py; pure cores SwitchStopsCore, the Kronos ledger cores, StaffCountsCore; SyncKronosData / DrawCalendar
/ DrawNeeds in steps (callcheck KNOWN_LONG empty); shared form helpers in modUI (10 forms re-pasted); the admin Desk
preview; stamps v27.1, N_MODULES 41, WHATS_NEW_RELEASE v27 (no box re-shown). 962 cases / 106 suites, pytest 1017,
SmokeTest 22. Five read-only Opus reviews + a re-review of the release prep (see the CHANGELOG). The builders' branches
v271/b1-calendar, b1-wave2, b2-people, b3-postings, b4-kronos, b5-forms are merged (their worktrees under
.claude/worktrees can be cleaned up).

DONE Oct 22-23: Graeme chose no What's New box for v27.1 (as built) and OK'd the release; tagged, zipped, pushed.
DONE Oct 25 (Graeme's answers): **v27.1 IS ON THE REAL ROSTER** and passed his testing; **the Kronos section-0
re-import is DONE** (Worked H on the real roster is correct; nothing owed there any more). The merged v271/*
builder branches and their .claude/worktrees were removed with his OK.
NEXT (he chose it Oct 25): v27.5 = the alpha fixes + the three gaps found while reading (the same-day swap's 16 h
stop - put to Graeme first; the missed-morning double count in Worked Periods; the staffing count's partial-line
rounding / empty " / " part), built ON v27.1 (a v27 zip cannot go over v27.1); report back before code. Still
open: item 39's digest wording + recipients (ask when he is ready), the alpha testing Nov 1-15 with paper
alongside, the Reconcile differences for the Nov 15 pack.
```

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first (the numbering rule, the three gating levels, the
canonical business rules including the v27 rows - OT restricted, the weekly cap OFF, the Exception outcome -, the
"Decisions already made" bullets). Then briefs/v27.md (every item's BUILT note, review record and Graeme's decisions),
CHANGELOG.md's v27 entry, docs/INSTALL-v27.md and briefs/steps-v27.txt. This file's sections below hold the history.

STATE (Oct 22, 2026, end of the v27 build): v27 is built, every commit reviewed read-only (Opus) and fixed, tagged
v27 (annotated) and pushed with main to origin (https://github.com/tgpiercy/Call-Roster---V3); the release zip is
dist/Callout_Roster_v27.zip (full set + INSTALL.md + install-steps.txt + manifest.txt), built by
tests/build_zip.py v27 v26.5.5. Checkers: vbacheck shows only the three named one-line-Sub false positives
(Sort_Hours_Dept, frmLeave, modMember); shadowcheck / undeclcheck / callcheck / docscheck 0; pytest 914; SelfTest
877 cases in 100 suites (LoadCases1..5); SmokeTest 22 steps. N_MODULES 30 (five new: modReconcile, modStats,
modReceipt, modRules, modEvidence); three new code-built forms (frmReceipt, frmHistory, frmTableEdit); Roster
Updater v4.2; Members M4 (MEM_VERSION "Members M4 (Oct 22)", the master's MEMBER_NEEDS "Members M4").

FIRST, ASK GRAEME (multiple choice, once each):
1. The Kronos section-0 re-import on the REAL roster (briefs/v27-kronos.md section 0) - done? It must happen
   BEFORE v27 goes on the real roster (v26.5's import summed the timecard edits: Worked (H) is wrong until then).
2. Has v27 been installed on a practice copy yet (INSTALL-v27.md: the Updater's modUpdater.bas by hand FIRST, then
   three empty UserForms frmReceipt / frmHistory / frmTableEdit, then the Install panel)? What did SmokeTest,
   SystemCheck ("30 of 30 on v27") and the group 4 rows say? When he says "it doesn't work", ask which file and
   version is on screen first (SystemCheck's first line).
3. Item 39 (the publish digest by email): his wording and recipient list - not built until he signs them.

DECIDED BY GRAEME Oct 22 (don't re-ask): the weekly 48 h warning is OFF ("48h is scheduled work. OT is requested";
WEEK_MAX_HOURS 0, a Settings number turns it on); the daily CSV export stays supervisor-level; no members' dStats
for now; the paper-position column is dropped; the FULL history PDF stays beside the roster; an OT-restricted person
is left out of the members' "not listed" table.

THEN, in order: (a) whatever his practice-copy install turns up - fixes go in a v27.5 BUNDLE, never a single small
install (the numbering rule: FEWER INSTALLS); (b) the alpha testing Nov 1-15 with paper alongside, the Reconcile
sheet's differences for the Nov 15 pack; (c) v27 on the real roster about Nov 20-25 before the Dec 1 training;
(d) v27.5 = the behaviour-free refactors (CLAUDE.md roadmap: the pure cores, the module splits - DrawCalendar,
DrawNeeds and SyncKronosData are callcheck's KNOWN_LONG list), plus item 39 once signed; (e) the December rollover
dry run on v27; (f) v28 in January with go-live.

Known LOWs left (briefs/v27.md has them at their items): the first days strip shows for existing logins too and
ticks only through the Desk's buttons; the members' trust strip names the previous publish; the "Backup before ...
TEMPLATE" copies are not pruned; MACROSTR cannot see a macro name passed into a builder (AddChoice / AddBtn); after a
withdraw the posting stays Status Awarded; Post all lines uses the default shift hours; the Install panel's diff step
is not built; dStats waits.

How every update ships (CLAUDE.md): report back first; one commit per item; the four checkers + docscheck + pytest
after each; read-only Opus reviewers in isolation worktrees given the commit hashes (their worktree lands on an old
commit - tell them to read `git show main:<path>`; they must not mutate either tree); fix, re-review; stamps,
INSTALL note, steps file, CHANGELOG, tag, zip, push from Graeme's Terminal panel with his approval. git add files BY
NAME, never git add -A. Write patch scripts with the Write tool (the Bash heredoc mangles backslashes); a new .bas
written by the Write tool is LF - convert to CRLF. Put every decision to Graeme as a multiple-choice question.
Update docs/NEXT-SESSION-PROMPT.md at the end of the session.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## The Oct 21 session in one paragraph
Graeme sent the Kronos files (two one-day PDFs, the cumulative .xlsx, his processing .xlsm). Reading the real export showed v26.5's reader had summed the wrong block - the brief `briefs/v27-kronos.md` was written from the files, Graeme answered its open points (sections 3a, 3b, 4), and the live roster's H turned out wrong (section 0, the remedy, still owed). The v27 report-back ran as three read-only code checks + two rounds of multiple-choice questions (the decisions are in `briefs/v27.md`). Then the build: Desk v2, the LoadCases split, the Kronos import, the Reconcile sheet, the Stats report, each its own commit. Two read-only Opus reviews: the Kronos commit (5 HIGH / 10 MED, all fixed in 96f86b7, re-reviewed) and the Desk / Reconcile / Stats commits (its findings and fixes are recorded below when they land).

## Review record (v27 so far; two read-only Opus reviewers, every round re-reviewed until clean)
- **Kronos import 1388f89, reviewer 1, four rounds.** Round 1 (5 HIGH / 10 MED): the rename arrays one slot short; tables written before the sanity confirm (No left them in); the rollover kept last year's snapshot; an older or empty cumulative accepted; one-day reports alone setting H; edits dated against the new Through; a re-imported day keeping other people's rows; the Payroll # match without the name check; a roster person missing from the report keeping old H; exits that did not block the save; a wrong date in the message; quiet days read as missing; the pivot fall-back skipped on an unreadable Totals; no test of the reader; files picked by saved time only -> 96f86b7. Round 2 (1 HIGH / 6 MED): a clash or an unmatched row zeroed H; the save block set only after the ingest returned; IntegrityCheck PROBLEMs with one-day reports alone; edit rows into periods before the ledger start and after the rollover; a re-pulled day counted twice; H holding days past a gap; every file reopened -> 1d40a38 (+ the Kronos Files record). Round 3 (3 MED): a clash forgotten after one run; an edit dropped when its day was first imported later; January reopening a year of files -> 7fc840c. Round 4 (1 MED): a refused year-to-date dropping the held clashes -> 3629ec8. Verdict after round 4: closed (two cosmetic LOWs noted: a rename leaves the old name in a "clash:" row until the next clean year-to-date; a file literally named "clash: ..." would read as a clash).
- **Desk v2 / Reconcile / Stats (0da0e38, b2f13c2, 333d678), reviewer 2, two rounds.** Round 1 (1 HIGH / 5 MED): the Stats fairness line's 0 / 0 inside an IIf stopped the report for any unit with no offers (certain after the rollover); the roster year's late-December days in Q4; hours offered understated; the Reconcile "as of that day" claim false for the hours; long one-cell lines cut off in the PDFs; cross-coded people double-counted in the all-units sick denominator; + LOWs -> 8fac4a1 (Stats) and 7ba246a (Desk + Reconcile). Round 2: all fixed, no new HIGH / MED; the LOWs taken in the commit after (jan1 once per report, wrapped section titles, the differ count over every position, the Reconcile frame redrawn when its version in L1 changes with the typed cells kept, the stale comments). Left as noted: `TableByName` duplicates `TableData`; the denominator loop's cost (a few seconds); the mirrors' `.strip()` / `float` vs `Trim$` / `Val` (no case hits them); the Reconcile PDF prompt's "since it was last changed" wording; NBSP in typed names.
- **Items 8 + 9 (e5b30c8, d189b1f), reviewer 3, three rounds.** Round 1 (4 MED / 8 LOW): EnableEvents captured after `On Error` (an early error left events off for the session); a call-out that did not save left a fresh map a later card repainted under its own ref; the row banding's conditional format hid the fill on every second row; the 60 s ran from the paint, not the card's close; + the box past MsgBox's 1,024 characters, a foreign note deleted, the structure left unlocked on a failed sheet make, an answer "kept" on a copy that will not be saved, Enter closing the card from the note box, no LogError, a wrong colour comment, docs -> 052e10b. Round 2 (1 HIGH / 4 LOW): the NoSave refusal made SmokeTest step 4 fail every run (SmokeTest blocks saving from its first line) -> the write split into `WriteFeedbackRow`; a threaded comment stopping the loop; marks painted before an error never expiring; xlAutomatic = black not the roster's ink; the clock under the modal call-out form -> 112b997. Round 3 (112b997): clean, no finding (one doc nit taken: CLAUDE.md names WriteFeedbackRow).
- **Items 10, 14-22 (Oct 22, session 3), one Opus reviewer per round, every round re-reviewed.** Item 10: 1 HIGH (a Public Const between procedures - a compile error; now a vbacheck class), 5 MED (the auto-pick by unit not group, the reverse coverage check, the forward PROBLEM only from today, a support award asks, the unit change drops the sick key) -> 857ef4d; re-review 1 MED (a NO TAKERS key dropped) -> e44e2ff. Items 14-16: 1 HIGH (SignupSuspect: 10.000000001 min for ten exact minutes - whole seconds now), 5 MED (clock-change / skew, the POSTED / NO TAKERS key, the phone column saved visible, the print's row 3 and footer) -> e44e2ff; re-review 1 MED (a supervisor's mid-session saves) + 8 LOW -> e6b1ab4. Item 17: 3 MED (durability, a kept draft overwritten, a reload leaving the draft) -> e6b1ab4. Items 18, 19, 20A, 20B + e6b1ab4: review IN FLIGHT at the end of the session (its findings are NOT taken yet - take them first). Items 20C, 21, 22: review IN FLIGHT too.
- **Rules that came out of the reviews (now in the code and the briefs):** a roster row is set to 0 only when it HAS a Payroll # the year-to-date report does not hold (no number = "not found", H kept); a Payroll # clash leaves both people's H alone, is named, and is held in Kronos Files until a clean year-to-date; the one-day rows count only along the unbroken run of days after the Totals, and Data Through stops at the gap; the first v27 import needs the year-to-date report (one-day reports alone stay in the ledger).

## The members' side
Unchanged since Members M3.1 (v26.5). M4 (items 34-39) adds the me-first Home, the header band, dNotFree, dStats, the contract test and the MsgBox titles, as ONE members' install with v27.
