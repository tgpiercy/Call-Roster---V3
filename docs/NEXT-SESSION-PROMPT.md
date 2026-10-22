# Paste this into a fresh session to carry on (updated Oct 22, 2026: v27 is TAGGED, PUSHED and ZIPPED - dist/Callout_Roster_v27.zip; NOT installed yet; the Kronos section-0 re-import is still OWED)

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
