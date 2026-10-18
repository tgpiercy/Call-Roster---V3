# Paste this into a fresh session to carry on (written Oct 18, 2026, night: v25.5 installed, alpha testing next)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/timeline-2026-10-18.md,
the v25.5 entry at the top of CHANGELOG.md, and docs/code-kit-2026-10/briefs/07-v26-reporting.md with
docs/deep-review-2026-10-13.md sections 4.5 and 4.6. Confirm the four checkers (py -3.14 tests/vbacheck.py .
/ shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest tests -q are green on main and the tree is
clean (main is at a216d56 or later; everything is pushed to origin; tag v25.5 is on origin).

Where things stand: v25.5 (the ALPHA BUILD) is INSTALLED on the real roster (Oct 18): no issues, the
install's testing ran clean. Alpha testing runs Nov 1-15; its findings are batched into v25.5.5 (nothing
is installed on its own unless the roster is broken). Still open from the install: whether the call-out
pop-up's bottom is cut off on any plant PC (500 x 615 pt; if so v25.5.5 shrinks it to about 540 pt), and
whether the FIND group at the bar's right end is on screen for a supervisor.

Now: v26 = the BETA BUILD by Nov 15, presented to management and the union that day. Write
briefs/v26.md FIRST from the kit brief 07 + deep review 4.6 (employee history report, admin and
member-safe; "why was I charged" on the status bar and the Refusals Report; the two-week board print;
payback aging; the IdMark-vs-archive integrity line; backup housekeeping; LastReason on Attendance;
minimum staffing only if the minimums have arrived) and deep review 4.5 (M3: the members' template
rebuilt - withdrawals and login mismatches reported at LOAD, the PC name in the sign-up file name, Kind
in dLog, Exempt in dPostList, EstHrs with ~, the month picker, the phone pattern, an "Am I next?" tile, a
large-type mode; one members' install). Check every (verify) item against the code, ask the real
decisions as multiple-choice questions (recommended option first), and WAIT for my go-ahead before
writing code. Then build to the brief one commit per item, the four checkers after each; read-only
reviewers (opus, isolation worktrees, told to review `main`) before the tag; fix, re-review the fixes;
INSTALL-v26.md, briefs/steps-v26.txt, CHANGELOG, CLAUDE.md, every stamp to v26, FORM_RELEASES /
THISWB_RELEASE for whatever changed, tag v26 (annotated), build the zip, push main + the tag from my
Terminal panel. Never git add -A. Update docs/NEXT-SESSION-PROMPT.md at the end of the session.

In parallel, no code (docs/timeline-2026-10-18.md): the what-if scenario sheets, the four guides refreshed
for v25.5 / v26, the Nov 15 presentation pack, the Dec 1 training plan - start them when I say.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v25.5 holds (installed Oct 18; `docs/INSTALL-v25.5.md`)
The bundle (`briefs/v25.2.5.md` items 0-2, 6-9) folded into the pop-up release (`briefs/v25.5.md` items 1-7): supervisors by login (`SUPERVISOR_USERS`), supervisors publish, click details on the calendar with no hover boxes anywhere, the legacy writers and SmokeTest steps 19-20 out, the greyed admin award buttons, the Help tab, the call-out session reworked (the card, one key per outcome, Call back, the four-line summary), `UI_SCALE` on every pop-up, plain language, Ctrl+Shift shortcuts + the what's-new box, Desk refinements, Undo my last log entry, Find a person. ThisWorkbook and all nine forms re-pasted; 25 of 25 stamps; SelfTest 581 / 68; SmokeTest 19 steps; pytest 595.

## Review record
Four read-only reviewers in isolation worktrees (compile traps; gates / rules / privacy; events / state / the screen; completeness), then one re-review of the fixes. Two HIGH found and fixed before the tag: two Consts declared after procedures in modDesk (would not have compiled; no checker sees it - grep for it in every review), and the partial-acceptance close bypassable by re-recording Accepted on a row logged before the close (a hole since v25 through a list click). The CHANGELOG's v25.5 review paragraph lists the rest. Lesson for the next review: the isolation worktrees landed on an old commit; tell reviewers to review `main`.

## The members' side
Unchanged by v25.5 apart from who may publish (supervisors too). M3 (the members' ease-of-use follow-up) is in v26 as one members' install.
