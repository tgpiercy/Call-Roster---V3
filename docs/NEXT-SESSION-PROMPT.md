# Paste this into a fresh session to carry on (written Oct 20, 2026: v26.5 built, tagged and pushed; the Oct 19-20 reviews done; next = Graeme's installs, then v26.5.5 before Nov 1, then v27 before the Dec 1 training, v28 in January)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first (it was brought up to date Oct 20: the
numbering rule, the three gating levels, the 20 SmokeTest steps, the timeline, the shipped releases collapsed
to one line, the Oct 19-20 decisions as the first bullet of "Decisions already made"). Then
docs/review-2026-10-20-gaps.md (the three reviews with every finding numbered A1.., B1.., C1.., 1.x, 5.x and
its release), briefs/v26.5.5.md and briefs/v27.md (the two briefs written from it), docs/timeline-2026-10-18.md,
the v26.5 entry at the top of CHANGELOG.md. Confirm the four checkers (py -3.14 tests/vbacheck.py . /
shadowcheck / undeclcheck / callcheck; on Linux python3) and py -3.14 -m pytest tests -q are green on main and
the tree is clean. vbacheck's only hits must be the three named false positives (Sort_Hours_Dept, frmLeave,
modMember: one-line Subs). Tags: `git ls-remote --tags origin` should show v26 and v26.5.

Where things stand: v25.5 (the ALPHA BUILD) is INSTALLED on the real roster (Oct 18). v26 and v26.5 are BUILT
and TAGGED, NOT installed: docs/INSTALL-v26.md then docs/INSTALL-v26.5.md (v26.5 needs v26 first; both rebuild
the members' template; v26.5 re-pastes ThisWorkbook, frmEmployee, frmLeave and MUST run RefreshRosterFormulas
in group 3). Ask Graeme first whether they are on the real roster; nothing in v26.5.5 is built before they are
(the pre-release rule), and the first REAL Kronos export import is the test of the per-period hours (ask for
the "Premium hours per pay period" line).

NEXT BUILD = v26.5.5 (briefs/v26.5.5.md), installed BEFORE NOV 1 so the alpha runs on it: the integrity fixes
(a failed save reported and the form kept; a stale copy BLOCKS the save; Esc-safe writes; an error in Total
Hours refuses to load; Worked Periods in every rename path; ClearNoSave only for its own block; the small
ones; the members' folders moved), THE 16 H FATIGUE RULE (the one rule change: accepted OT / awards / switches
count as worked shifts for eligibility; max hours per person on the pop-ups; suite `fatigue`), the 3 AM fixes
(the night in progress dated yesterday between 00:00 and 07:00; the sick-call confirm - undo stays admin;
default hours and the sick call's hours carried; Kronos-stale as a red line for supervisors; the next-step line
after a submit; the text pass with a title on every box and no internal names), the outcome-key labels, the
trust strip + ChargedText / TotalText, the Practice button. 19 items, one commit each, the four checkers after
each, two read-only Opus reviewers in isolation worktrees told to review `main`, re-review the fixes, INSTALL
note, steps file, stamps, tag, zip, push from Graeme's Terminal panel. Step 2 of "How every update ships"
first: the report-back with the (verify) items checked and the open points as multiple choice (the brief lists
them). Never git add -A. Update docs/NEXT-SESSION-PROMPT.md at the end of the session.

THEN v27 (briefs/v27.md), one master install about Nov 20-25 + ONE members' install (Members M4), built during
the alpha once v26.5.5 has run on the real roster, with the alpha findings folded in: build order Desk v2
(the approved mock-up) -> the Reconcile sheet (paper runs alongside for at least one pay period; the
differences go into the Nov 15 pack, so this one may be installed early on a copy) -> the Stats report
(Graeme signs the counting rules before any code) -> the receipts (one card after every save, every form) ->
the 3 AM pass (log/calendar reconciliation, Kronos hardening, backups on any open + RestoreDrill, the audit of
in-place edits, the inbox hardening, the calendar's supervisor menu, the outage print, the session draft,
"why isn't Joe listed", the error box's saved line, the pop-up polish) -> the process items (OT Restricted +
the weekly-cap warning, the Rules and Settings sheets, period close / scan link / disputes, withdraw award /
edit posting / post all lines / Exception, ExportTables / data check / unit notes / digest / handover / switch
ledger) -> the Desk tiles, "Since you were last here", the Holidays + Min Staff pop-ups -> the testing items
(IntegrityCheck "Log rows", SmokeTest asserts and steps, LoadCases split, five checker classes, drift and docs
tests, the manifest, the install record in the updater, the Opens table) -> language (glossary, legend, guided
first session, order columns on every list, the transition-only paper column) -> Members M4 (me-first Home,
highlight me, How it works, Sign me up from the tile, since-you-last-looked, dNotFree, the contract test, the
email digest only after Graeme signs the wording). Four read-only reviewers, one master install, the December
rollover dry run on v27 with RestoreDrill in its group 4.

AFTER v27: v27.5 = the behaviour-free refactors alone (pure cores extracted + suites; the shared form-builder
helpers; the module splits: modCalendar -> Board / MinStaff / CalImport / Needs / CalMenus, modPeople -> Dates /
PlannedMoves / Former, modPostings -> Mail / Sheet / Sync); and the CLAUDE.md split into CLAUDE.md +
docs/workbook-map.md + docs/decisions.md + docs/settings.md (review part 0), in a session with no code change.
v28 in January after the rollover, with go-live: members' requests (switches + time off), the phone page,
v25.4 Payroll # as the key, the P: paths.

Decided Oct 19-20 (don't re-ask; CLAUDE.md has the list): 16 h max consecutive hours; no forced OT; Left
Message charged at once; personal logins; a stale copy blocks; backups on any writable open; sick-call undo
admin-only; paper alongside for a pay period; receipts on every form; the 16 h rule in v26.5.5; the weekly cap
warns with a logged override.

In parallel, no code (docs/timeline-2026-10-18.md): the what-if scenario sheets, the four guides refreshed
for v26.5 / v26.5.5, the Nov 15 presentation pack (with the Reconcile PDF), the Dec 1 training plan - start
them when Graeme says.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v26.5 holds (built and tagged Oct 19; `docs/INSTALL-v26.5.md`)
Graeme's five changes from the v26 testing plus the masked month PDF: nothing on the roster is typed into (the Employee pop-up's **Edit**, opened by a double-click on the row; `SaveEmployeeEdit` is the one write path); **Refused counts the pay periods up to the current one** (`RefusedFormula` = the row's period cells whose header date <= the hidden name `RefusedThrough`; `EnsureRefusedThrough` on open, in `LoadEligible` and in `RosterTotalsMap`; the rollover, new hire, bring back, IntegrityCheck, the split and the members' figures follow); the short flag counts every unit; My history's pay-period block = Charged, Posting refusals, Call-out refusals, Worked (Kronos); the Kronos EXPORT read directly (`TransactionsSheet` / `PivotFromTransactions` / the admin-only Worked Periods sheet, published as `dWorked`; the pivot workbook is the fall-back); `SIGNUPS_ON` True; `CalPrintPdf` draws the month masked. ThisWorkbook, frmEmployee and frmLeave re-pasted; Members M3.1 (template rebuilt); 25 of 25 stamps on v26.5; SelfTest 649 / 75; SmokeTest 20 steps; pytest 663.

## Review record (v26.5)
Four read-only Opus reviewers in isolation worktrees on 4febe25 (compile traps; rules / gates / privacy; events / state / the screen / the install; completeness / tests / docs), then two re-reviewers on the fix commit 4ff50cb. All four found the same root defect: the first Refused formula (a SUMIFS over the log) was unbounded at the year end, so from Dec 27 it counted next year's rows and would have blocked the rollover, and four older checks (the rollover's, new hire, bring back, IntegrityCheck's split) still summed every period. Fixed by summing the row's own period cells up to `RefusedThrough` and capping the checks the same way. Also fixed: the masked PDF's text colour gave the leave kind away; an unfinished masked draw was exported; a crew date before the newest crew row; the pop-up keeping the previous person's home unit; a text seniority value blocking every Edit save; the members' "Sent" / Withdraw reading the wrong folder; total lines counted in the export; non-roster people on Worked Periods. Lesson kept: when a rule changes a roster figure, grep every place that recomputes it (the rollover, the people saves, IntegrityCheck, the members' build) before the reviewers do.

## The members' side
Members M3.1: the template is rebuilt by the panel's `template` row (first row of group 3; publishing refuses until it is done). `INSTALL-members.md` has the M3.1 paragraph.
