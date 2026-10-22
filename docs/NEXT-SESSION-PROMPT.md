# Paste this into a fresh session to carry on (updated Oct 22, 2026, end of v27 build session 3 - items 8 + 9; first written Oct 21 at the end of session 1: v26.5.5 is ON THE REAL ROSTER; the Kronos import was found WRONG and rebuilt; v27 items 1, 4, 11, 26 and the LoadCases split are BUILT and reviewed on main; the live re-import is still OWED by Graeme)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first (the numbering rule, the three gating levels,
the 20 SmokeTest steps, the timeline, the "Decisions already made" bullets for Oct 21 and Oct 19-20, the "16 h
rule" row of the not-free table, the Total Hours rule's v27 sentence). Then briefs/v27.md (the "Report-back DONE
Oct 21" block holds every decision; items 1, 4, 11, 26 are marked BUILT), briefs/v27-kronos.md (sections 0, 1,
1b, 2, 3 and the Status line), docs/NEXT-SESSION-PROMPT.md (this file: the review record below). Confirm the
four checkers (py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest
tests -q are green on main (793) and the tree is clean. vbacheck's only hits must be the three named false
positives (Sort_Hours_Dept, frmLeave, modMember: one-line Subs). SelfTest cases: 807 in 96 suites, emitted as
LoadCases1..4. Tags: `git ls-remote --tags origin` shows v26.5 and v26.5.5; main is AHEAD of v26.5.5 by the v27
commits (not pushed yet unless Graeme pushed from his Terminal panel: ask).

THE ONE URGENT THING (ask about it first, once; asked again Oct 22: STILL NOT DONE, and main was still not pushed):
the real roster's Worked (H) has been WRONG since the Oct 21
v26.5 import - the export's dated "Transactions" block is timecard EDITS (435 of 629 rows negative, ~800 h net)
and v26.5 summed it instead of the undated "Totals" block (~29,500 h, the real year-to-date). Graeme confirmed
the live import read the transactions. The no-code remedy is briefs/v27-kronos.md section 0 (delete the pasted
export sheet in his processing workbook, Import Kronos must say "Read from the pivot sheet", clear Worked
Periods through UnlockForRepair, republish). He said "Yes, doing it now" at 1 PM but had NOT done it by the end
of the session. Ask: done, and what did the finishing box say? If not done, it comes before anything else.

WHAT v27 HOLDS SO FAR (main, one commit per item, all four checkers + pytest green after each):
- 0da0e38 item 1 Desk v2 (the approved mock-up: three header pills, verbs on the tiles, NOW / PLAN / FINISH
  side by side, the grey footer row; UI-only, same macros, same rights; Help paragraph + docs/desk-guide.md).
- 738c2e9 item 30a the LoadCases split (gen_selftest_cases emits LoadCases1..N of 200 inside the one marker
  pair; test_drift asserts each chunk < 56 KB).
- 1388f89 + 96f86b7 + 1d40a38 + 7fc840c + 3629ec8 item 11 the Kronos import REBUILT from the report's Totals block (briefs/v27-kronos.md):
  ReadKronosReport / ReadKronosArray, the Worked Days ledger (one-day reports), the Kronos Totals snapshot
  (the year to date), the Worked Changes log (edits dated by the Transactions block), H = snapshot + ledger
  days after KronosCumThrough, hours x per-code multiplier (PREMIUM_CODES / PREMIUM_MULTIPLIERS in
  modOvertimeEntry: OVERTIME 1.75, OT - DT Split 2, LIEU OT 1.5, LIEU DT 2 [corrected from his workbook's
  swapped pair], LIEU OT - DT Split 1.5, LIEU ST 1, CALLIN OT 1.75, MIN 4 1.5, STATWK 2.0 2; DBLTIME /
  Uanpproved OT / LIEU TKN never an operator's: reported, not counted), rounded whole per person
  (RoundHalfUp = Excel's ROUND), PivotFromLedger feeds the unchanged name / Payroll # matching, the verified
  H write (PeopleBackup, the sanity confirm through Ask, one array, Refused + period cells compared,
  MarkNoSave), Data Through = the Time Period end, every newer report oldest first; the pivot sheet is the
  fall-back; ClearKronosTables at the rollover; IntegrityCheck "Kronos" (H = tables, codes not counted,
  missing days, orphans); SmokeTest step 15 reads a made-up three-block report; suites kmult / kweight /
  kperiod / kkind; the Kronos Files record (every file read, taken or skipped, + the held clashes). Four review
  rounds, all taken (the record below).
- b2f13c2 item 4 the Reconcile sheet (modReconcile, NEW MODULE; N_MODULES 26): the paper call order typed
  beside the roster's (LoadEligible "ALL" as of the typed day), the reason per differing position, Decision +
  Note, Compare / Save PDF / Clear, admin Desk FINISH button; "From the old log" NOT built (the legacy Entry
  Log's layout is not in the repo - ask Graeme for its columns if he wants it).
- 333d678 item 26 the Stats report (modStats, NEW MODULE; N_MODULES 27): unit x quarter + year to date, four
  groups to the signed rules (fairness over Total Hours first, premium hours second); StatsUnitRows ready for
  M4's dStats. Not built: occurrences vs days, "how covered". + 8fac4a1 / 7ba246a / the LOW commit = the
  review fixes for Stats, Desk v2 and Reconcile (the record below).
- OCT 22 (session 2): 7cc658c item 7 the RECEIPTS (modReceipt, NEW MODULE, N_MODULES 28; frmReceipt, a NEW
  code-built form in FORM_RELEASES; one card after every save: call-out, the four posting finishers, sick call
  save-only, switch, time off, leave going / returning, Adjustment, Void, the Employee pop-up; Print = PDF via the
  Receipt output sheet, Email = a mail draft; SmokeTest logs the text and step 4 asserts it) + fixes 5a379f3 /
  207d292; b619023 item 12 BACKUPS (daily backup on any writable open, hourly backup before a save when the newest
  is older than 4 h kept 7 days, the first daily backup of each pay period kept 400 days, RestoreDrill with an
  Install Log row, LogInstallRow) + fixes 207d292 / fe96720 (a HIGH: PruneBackups could have deleted every
  backup on a bad header); 72ff653 item 13 the AUDIT OF IN-PLACE EDITS (Grade Log Field column + LogFieldChange,
  SetAbsenceEnd / SetAbsenceEndTo keep the old end in the Note, ReleaseCoverage voids and re-writes) + its fix
  round. All reviewed read-only (one Opus reviewer, every round re-read until clean or LOW-only).
- OCT 22 (session 3): e5b30c8 item 8 BEFORE / AFTER ON THE ROSTER (modReceipt.MarkChangedTotals paints every Total
  Hours cell a save moved from the receipt's own before / after map: a note "Was 120, +21 from S-0012 (3:10 AM)" +
  amber fill + amber ink [the row banding's conditional format hides a fill alone]; ClearTotalMarks at ApplyViewCore,
  the next save that reads a before map [ReceiptBefore], or the first roster click 60 s after the card closes
  [RestartMarkClock / TotalMarksExpired in RosterSelectionChanged; NO OnTime]; pure FirstRefIn, suite firstref 16
  cases; the Total Hours DOUBLE-CLICK for any login = modReports.ShowTotalBreakdown, a box with the parts and the
  charged hours per pay period; SmokeTest step 4 asserts the acceptor's "+8" mark) + d189b1f item 9 "WAS THIS
  RIGHT?" (Yes / Not sure / No + an optional line on frmReceipt -> tblFeedback on the admin-only Feedback sheet
  [When, Login, Release, Ref ID, Answer, Note, Seen; made on first use], UnseenFeedbackCount on a Needs Attention
  line [key F:], ShowFeedback marks seen; SaveFeedback REFUSES on a read-only / NoSave copy and the write is
  WriteFeedbackRow so SmokeTest step 4 tests both; the Desk's morning tile [item 2, not built] takes the count
  when built) + fixes 052e10b (4 MED / 8 LOW) + 112b997 (1 HIGH the fix round had made: step 4 vs the NoSave
  refusal; 4 LOW). frmOvertimeEntry re-pasted again (NOT SAVED branch paints its marks; QueryClose restarts the clock).
- OCT 22 (session 3, continued - Graeme: "No let's build v27. Continue with this"; a v27-pre1 preview zip was started and
  REVERTED on his word): THE 3 AM PASS IS BUILT - 2f23d7c item 10 log/calendar reconciled (auto-pick the UNIT's one open
  line, the sick call's key rides Prefill -> SubmitCalloutSession -> CalendarAddAccepted, IntegrityCheck 12 "Coverage =
  log" [PROBLEM from today on, Check for the past], the 12 h second-session question, VoidRef reports a failed release)
  + fixes 857ef4d; f75babf item 14 the sign-up inbox (SG_MTIME, rows sorted, SR_SUSPECT in WHOLE SECONDS w/ a clock-change
  allowance, SR_LOGINMIX; suites sgsuspect / sgmix); 8efc66b item 15 the calendar's supervisor menu + CalMenuCallout +
  NextOpenSlot for the Desk tile (NO TAKERS lines filled IN PLACE, no call-out item on a POSTED line); 8ce0a48 item 16
  the outage print (ten columns, the session line, the legend footer; the phone column hidden again after the print);
  31b6236 item 17 the session DRAFT (very-hidden Draft sheet, one row per login, saved to disk every 2 min, the resume
  offer in UserForm_Activate, Esc = keep / DISCARD / stay) + fixes e44e2ff / e6b1ab4; 5b86cf2 item 18 "Not listed (N)..."
  (FilterNotFree's detail); 46985af item 19 the error box (NewestLogRowByMe, the email offer); item 20 in three parts
  364551c / 2092571 / f140613 (the award pop-up's keys and 10 pt, UnitDisplayName, Find a person; type floors, UiScale()
  per login + the Desk's Larger type, "How do I ..." on the bar for everyone, "Signed in as" on the status cell, the
  first-open box; frmHistory = a NEW code-built form); 392de3e item 21 OT restrictions (tblRestrictions, never charged,
  RestrictOT / EndRestriction, the Employee pop-up's Edit button) + the 48 h weekly cap (WeekCapOver, suite weekcap);
  965fe3d item 22 modRules = a NEW MODULE (N_MODULES 29): the Rules sheet (23 rules in code, Basis typed) + the Settings
  sheet (six tunables through wrappers, SetSetting). TOOLING: fa67911 vbacheck flags a module-level Const / Dim / Type
  after the first procedure (a compile error three review rounds read past; it then caught me twice more).
  NEW FORMS now TWO (frmReceipt, frmHistory); NEW MODULES now FOUR (modReconcile, modStats, modReceipt, modRules).
Checkers and pytest after the last commit (965fe3d): four checkers clean (the three vbacheck false positives), 852
passed, 832 cases / 100 suites (firstref, sgsuspect, sgmix, weekcap new). The re-pasted forms (frmOvertimeEntry, frmSickCall, frmShiftSwitch, frmTimeOff,
frmLeave, frmEmployee, frmReceipt) and ThisWorkbook changed: release prep bumps FORM_RELEASES / THISWB_RELEASE and the
INSTALL note lists THREE new modules (modReconcile, modStats, modReceipt) and ONE new form (frmReceipt).
Stamps are still v26.5.5 everywhere (test_drift ties every stamp to EXPECTED_RELEASE): the release-prep
commit bumps all 27 + the two new modules' stamps, FORM_RELEASES / THISWB_RELEASE as touched.

DECIDED Oct 21 (the report-back; briefs/v27.md has the block; don't re-ask): build order Desk v2 -> Kronos
import -> Reconcile -> Stats -> receipts -> the 3 AM pass -> process items -> Desk tiles / Since you were last
here / the two pop-ups -> the testing items -> language -> Members M4; Stats rules signed with fairness over
Total Hours; OT Restricted in its own table tblRestrictions, never charged while restricted; a second call-out
for the same shift inside 12 h ASKS; backups 4 h / 7 days / 400 days + RestoreDrill in group 4; period close
.\Closed + certutil; A11 audit = a Field column on tblGradeChange; the paper column removed after the second
full pay period; the M4 digest's wording signed before it is built. Kronos (Graeme Oct 21): a one-day snapshot
every morning + the cumulative at each period end; Excel export; Welland Hourly query; no backfill before the
first snapshot ("not available before <date>"); no union-folder PDF (the members' pay-period block shows the
premium hours); the live re-import OWED.

NEXT in order (briefs/v27.md): the 3 AM pass is DONE (10-20); next = item 23 (evidence: period close, scan link, disputes), the process items (21-25; 21 as tblRestrictions), the Desk tiles (2:
"Import now" = KronosFilesToImport has unread files), 3, 33, the testing items (28, 29, 30 b-e, 31, 32), 27
language, Members M4 (34-39, + dStats from StatsUnitRows, + the MsgBox titles), then 40 release prep
(stamps, INSTALL-v27.md: TWO NEW MODULES to import - modReconcile, modStats - the updater's swap step must
handle a module that does not exist yet: check updater/modUpdater.bas, say so in the note; steps-v27.txt
group 4 = Desk v2 on a supervisor login, a receipt, the Reconcile sheet with one typed order, the Stats PDF,
the Kronos acceptance test: after the section-0 re-import, importing the real cumulative on a copy must
change EXACTLY the people with LIEU OT / LIEU DT hours and nobody else; the session draft; the supervisor
menu; the outage print; RestoreDrill), four read-only reviewers on main, the zip, the push.

v26.5.5 leftovers still open for v27: ChargedText / TotalText not in the call-out summary / status bar; the
practice banner deleted by the three full-shape-delete redraws (RenderCalloutRoster, ShowAttendance,
RenderPostingSheet: delete by prefix); NotFreeReason's 16 h readers per free person (measure a draw);
the stale-copy guard's DST hour (accepted).

How every update ships (CLAUDE.md): report back first; one commit per item; the four checkers after each;
read-only Opus reviewers (commit range + "quote file:line"; they may not edit; their worktree may be behind -
tell them the hash); fix, re-review; stamps, INSTALL note, steps file, CHANGELOG, tag, zip
(tests/build_zip.py vNN vPREV on a clean tagged tree), push from Graeme's Terminal panel. Never git add -A.
Update docs/NEXT-SESSION-PROMPT.md at the end of the session. Write files with the Write/Edit tools (the
Bash heredoc mangles backslashes); a NEW .bas written by the Write tool comes out LF: convert to CRLF
(python) before vbacheck.
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
