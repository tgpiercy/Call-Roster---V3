# Paste this into a fresh session to carry on (written Oct 19, 2026: v26.5 built and tagged; the push, the installs and the alpha testing are next)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/timeline-2026-10-18.md,
the v26.5 entry at the top of CHANGELOG.md and briefs/v26.5.md (the brief: the seven items and the
decisions inside). Confirm the four checkers (py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck /
callcheck) and py -3.14 -m pytest tests -q are green on main and the tree is clean (main carries tags v26
and v26.5; check `git ls-remote --tags origin` - if v26.5 is missing, the push from my Terminal panel has
not happened yet: `git push origin main` and `git push origin v26.5`).

Where things stand: v25.5 (the ALPHA BUILD) is INSTALLED on the real roster (Oct 18). v26 (the beta build)
and v26.5 (the revisions after it: Employee > Edit instead of typing on the roster, Refused counts the pay
periods up to the current one = a RULE CHANGE, the short flag for every unit, the per-period split +
Worked (Kronos) in My history, the Kronos export read directly into a Worked Periods sheet, sign-ups ON,
the month PDF masked) are BUILT and TAGGED, NOT installed: docs/INSTALL-v26.md then docs/INSTALL-v26.5.md
(v26.5 needs v26 first; both rebuild the members' template; v26.5 re-pastes ThisWorkbook, frmEmployee,
frmLeave and MUST run RefreshRosterFormulas in group 3).

Decided Oct 19 after the reviews (briefs/v26.5.md, the last section): one rule for call-outs and
postings (a charge carries its shift's pay period and counts once that period starts); the sign-up
files' login in a folder members can list is accepted for the beta; KRONOS_PATH = the exports' folder
(newest file). Nothing is owed on v26.5.

Then: install v26 + v26.5 on a copy, then the real roster; the first REAL export import is the test of
the per-period hours (send me the "Premium hours per pay period" line). Alpha findings from Nov 1-15 =
v26.5.5 or v27 (one commit per finding, the four checkers after each, read-only Opus reviewers in isolation
worktrees told to review `main`, re-review the fixes, INSTALL note, steps file, stamps, tag, zip, push from
my Terminal panel). Never git add -A. Update docs/NEXT-SESSION-PROMPT.md at the end of the session.

Desk v2 (decided Oct 19): docs/desk-v2-mockup-2026-10.html is the mock-up (four bands NOW / PLAN / FINISH
+ TODAY, three colours, header pills, action tiles; same buttons and rights). No code until the boss's first
alpha week (Nov 1-7) has shaped it; then a UI-only item (DrawDesk redraw, no form re-paste) in v26.5.5.
UI_SCALE and the call-out form's fit wait for the plant screen measurement.

In parallel, no code (docs/timeline-2026-10-18.md): the what-if scenario sheets, the four guides refreshed
for v25.5 / v26 / v26.5, the Nov 15 presentation pack, the Dec 1 training plan - start them when I say.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v26.5 holds (built and tagged Oct 19; `docs/INSTALL-v26.5.md`)
Graeme's five changes from the v26 testing plus the masked month PDF: nothing on the roster is typed into (the Employee pop-up's **Edit**, opened by a double-click on the row; `SaveEmployeeEdit` is the one write path); **Refused counts the pay periods up to the current one** (`RefusedFormula` = the row's period cells whose header date <= the hidden name `RefusedThrough`; `EnsureRefusedThrough` on open, in `LoadEligible` and in `RosterTotalsMap`; the rollover, new hire, bring back, IntegrityCheck, the split and the members' figures follow); the short flag counts every unit; My history's pay-period block = Charged, Posting refusals, Call-out refusals, Worked (Kronos); the Kronos EXPORT read directly (`TransactionsSheet` / `PivotFromTransactions` / the admin-only Worked Periods sheet, published as `dWorked`; the pivot workbook is the fall-back); `SIGNUPS_ON` True; `CalPrintPdf` draws the month masked. ThisWorkbook, frmEmployee and frmLeave re-pasted; Members M3.1 (template rebuilt); 25 of 25 stamps on v26.5; SelfTest 649 / 75; SmokeTest 20 steps; pytest 663.

## Review record (v26.5)
Four read-only Opus reviewers in isolation worktrees on 4febe25 (compile traps; rules / gates / privacy; events / state / the screen / the install; completeness / tests / docs), then two re-reviewers on the fix commit 4ff50cb. All four found the same root defect: the first Refused formula (a SUMIFS over the log) was unbounded at the year end, so from Dec 27 it counted next year's rows and would have blocked the rollover, and four older checks (the rollover's, new hire, bring back, IntegrityCheck's split) still summed every period. Fixed by summing the row's own period cells up to `RefusedThrough` and capping the checks the same way. Also fixed: the masked PDF's text colour gave the leave kind away; an unfinished masked draw was exported; a crew date before the newest crew row; the pop-up keeping the previous person's home unit; a text seniority value blocking every Edit save; the members' "Sent" / Withdraw reading the wrong folder; total lines counted in the export; non-roster people on Worked Periods. Lesson kept: when a rule changes a roster figure, grep every place that recomputes it (the rollover, the people saves, IntegrityCheck, the members' build) before the reviewers do.

## The members' side
Members M3.1: the template is rebuilt by the panel's `template` row (first row of group 3; publishing refuses until it is done). `INSTALL-members.md` has the M3.1 paragraph.
