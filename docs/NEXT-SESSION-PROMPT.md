# Paste this into a fresh session to carry on (written Oct 18, 2026, night: v26 built and tagged; the push, the install and alpha testing are next)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/timeline-2026-10-18.md,
the v26 entry at the top of CHANGELOG.md and briefs/v26.md (the brief: every item checked against the code,
the eight decisions and the staffing minimums inside). Confirm the four checkers (py -3.14 tests/vbacheck.py .
/ shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest tests -q are green on main and the tree is
clean (main carries tag v26; check `git ls-remote --tags origin` for v26 - if it is missing, the push from
my Terminal panel has not happened yet: `git push origin main` and `git push origin v26`).

Where things stand: v25.5 (the ALPHA BUILD) is INSTALLED on the real roster (Oct 18), testing clean;
alpha testing runs Nov 1-15, its findings batched into v26 (the beta build) before Nov 15, or v26.5 after.
v26 is BUILT and TAGGED (Oct 18): docs/INSTALL-v26.md, briefs/steps-v26.txt, dist/Callout_Roster_v26.zip
(the members' template is rebuilt by the panel: Members M3). It is NOT installed yet. Four read-only
reviews + one re-review: four HIGH found and fixed before the tag (the CHANGELOG's review paragraph).

Decided Oct 18 night (briefs/v26.md decision 9): a cross-coded person counts in BOTH units' minimum
staffing; a sick call = off on the booked shift only; the month calendar's own PDF (CalPrintPdf prints
"(Sick)") is MASKED to "Off" in v26.5 = the first v26.5 item. Still open: whether the call-out pop-up's
bottom is cut off on any plant PC and whether the FIND group is on screen for a supervisor.

Then: install v26 on a copy, then the real roster (the panel; group 3 asks you to check the Min Staff
sheet's numbers); alpha findings + what the Nov 15 review asks for = v26.5 (one commit per finding, the
four checkers after each, read-only Opus reviewers in isolation worktrees told to review `main`, re-review
the fixes, INSTALL-v26.5.md, briefs/steps-v26.5.txt, stamps, tag, zip, push from my Terminal panel).
Never git add -A. Update docs/NEXT-SESSION-PROMPT.md at the end of the session.

In parallel, no code (docs/timeline-2026-10-18.md): the what-if scenario sheets, the four guides refreshed
for v25.5 / v26, the Nov 15 presentation pack, the Dec 1 training plan - start them when I say.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v26 holds (built and tagged Oct 18; `docs/INSTALL-v26.md`)
The reports (employee history admin + member-safe, "why was I charged" on the status bar / the Refusals Report / the members' My history, the two-week board, payback aging, the IdMark-vs-archive integrity line + `RepairIdMarks`, backup housekeeping, `LastReason` on Attendance), **minimum staffing** with the plant's numbers (Phosphine 2, Derivatives 3, Cypure 1, Material Handler 1, QC Derivatives 1, QC Gases / Utilities / SBP 1 on days only; the Min Staff sheet is admin-typed), the members' follow-up M3 (one template rebuild: Am I next?, Larger type, My history's Call order / Why / Print, Exempt rows, `~`, withdrawals at LOAD, the PC in the sign-up file name). No form and no ThisWorkbook change; 25 of 25 stamps on v26; SelfTest 644 / 74; SmokeTest 19 steps; pytest 658.

## Review record
Four read-only Opus reviewers in isolation worktrees (compile traps; gates / rules / privacy; events / state / the screen / performance; completeness and tests), then one re-review of the fixes. The worktrees again landed on an old commit: every reviewer was told to review `main` and did (`git show main:` / `git archive main`). Four HIGH before the tag: `DashedSeven`'s `IIf` evaluated `Mid$(s, 0, 1)` (a runtime error, two SelfTest cases failing at install); every shift switch counted twice in the staffing count; the board printed "(Sick)" where members read it; the SHORT list clipped the calendar's day number. Lesson kept: an `IIf` whose unused branch can raise, and a module-level Const or variable after a procedure, are the two traps no checker sees - grep for both in every review.

## The members' side
Members M3: the template is rebuilt by the panel's `template` row; the master refuses an older template. `INSTALL-members.md` section 4 has the new smoke-test rows.
