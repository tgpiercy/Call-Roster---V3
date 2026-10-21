# Paste this into a fresh session to carry on (written Oct 21, 2026, end of the v27 build session 1: v26.5.5 is ON THE REAL ROSTER; the Kronos import was found WRONG and rebuilt; v27 items 1, 4, 11, 26 and the LoadCases split are BUILT and reviewed on main; the live re-import is still OWED by Graeme)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first (the numbering rule, the three gating levels,
the 20 SmokeTest steps, the timeline, the "Decisions already made" bullets for Oct 21 and Oct 19-20, the "16 h
rule" row of the not-free table, the Total Hours rule's v27 sentence). Then briefs/v27.md (the "Report-back DONE
Oct 21" block holds every decision; items 1, 4, 11, 26 are marked BUILT), briefs/v27-kronos.md (sections 0, 1,
1b, 2, 3 and the Status line), docs/NEXT-SESSION-PROMPT.md (this file: the review record below). Confirm the
four checkers (py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest
tests -q are green on main (793) and the tree is clean. vbacheck's only hits must be the three named false
positives (Sort_Hours_Dept, frmLeave, modMember: one-line Subs). SelfTest cases: 776 in 93 suites, emitted as
LoadCases1..4. Tags: `git ls-remote --tags origin` shows v26.5 and v26.5.5; main is AHEAD of v26.5.5 by the v27
commits (not pushed yet unless Graeme pushed from his Terminal panel: ask).

THE ONE URGENT THING (ask about it first, once): the real roster's Worked (H) has been WRONG since the Oct 21
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
Checkers and pytest after the last commit: four checkers clean (the three vbacheck false positives), 793
passed, 776 cases / 93 suites.
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

NEXT in order (briefs/v27.md): item 7 the receipts (one card after every save; every save path outside
modPeople must gain a RosterTotalsMap before/after - the Oct 21 code check found none has it), items 8-9,
the 3 AM pass (10, 12-20; 11 is done), the process items (21-25; 21 as tblRestrictions), the Desk tiles (2:
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
- **Rules that came out of the reviews (now in the code and the briefs):** a roster row is set to 0 only when it HAS a Payroll # the year-to-date report does not hold (no number = "not found", H kept); a Payroll # clash leaves both people's H alone, is named, and is held in Kronos Files until a clean year-to-date; the one-day rows count only along the unbroken run of days after the Totals, and Data Through stops at the gap; the first v27 import needs the year-to-date report (one-day reports alone stay in the ledger).

## The members' side
Unchanged since Members M3.1 (v26.5). M4 (items 34-39) adds the me-first Home, the header band, dNotFree, dStats, the contract test and the MsgBox titles, as ONE members' install with v27.
