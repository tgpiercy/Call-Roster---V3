# Paste this into a fresh session to carry on (written Oct 18, 2026, late: v25.5 tagged, alpha testing next)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/timeline-2026-10-18.md,
docs/INSTALL-v25.5.md and the v25.5 entry at the top of CHANGELOG.md. Confirm the four checkers
(py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest tests -q
are green on main and the tree is clean (main is at the commit after tag v25.5, or later).

Where things stand: v25.5 = the ALPHA BUILD is built, reviewed (four read-only agents + a re-review of
the fixes), tagged and zipped (dist/Callout_Roster_v25.5.zip, built from the tag). The real roster is on
v25.3. The push of main + the tag to origin, and the install of v25.5 on a copy, are Graeme's; tell me
what the install's group 3-4 rows said (SmokeTest 19 steps, the call-out pop-up's keys, the screen fit).

Answered by Graeme Oct 18 (late): the plant screens are MIXED / not sure - v25.5 ships as built (the call-out
pop-up 500 x 615 pt = 820 px tall: fits 1920 x 1080, cut at the bottom on 1366 x 768; UI_SCALE 1); the
install's hand check reports whether the bottom is cut off, and v25.5.5 shrinks the form to about 540 pt
(list 115 tall, the card tighter) if it is. The FIND group stays at the bar's right end (it moves into the
ADMIN slot when that group is hidden); move it only if the install shows it off-screen.

Next work, in order (one commit per item, the four checkers after each, read-only reviewers before a tag):
 1. v25.5.5 = whatever alpha testing (Nov 1-15) turns up, batched; the screen-fit change above if needed.
 2. v26 = the BETA BUILD by Nov 15: the reports (docs/code-kit-2026-10/briefs/07-v26-reporting.md + deep
    review 4.6: employee history, "why was I charged", the two-week board print, payback aging, the
    IdMark-vs-archive integrity line, backup housekeeping, LastReason on Attendance), the members' follow-up
    M3 (deep review 4.5, one members' install), the alpha's findings. Write the brief first, report back,
    wait for the go-ahead.
 3. During alpha (no code): the what-if scenario sheets, the four guides refreshed, the Nov 15 presentation
    pack, the Dec 1 training plan (docs/timeline-2026-10-18.md).
Ask every decision as a multiple-choice question with the recommended option first; never git add -A.
Update docs/NEXT-SESSION-PROMPT.md at the end of the session.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v25.5 holds (one install; `docs/INSTALL-v25.5.md`)
The bundle (`briefs/v25.2.5.md` items 0-2, 6-9) folded into the pop-up release (`briefs/v25.5.md` items 1-7): supervisors by login (`SUPERVISOR_USERS`), supervisors publish, click details on the calendar with no hover boxes anywhere, the legacy writers and SmokeTest steps 19-20 out, the greyed admin award buttons, the Help tab, the call-out session reworked (the card, one key per outcome, Call back, the four-line summary), `UI_SCALE` on every pop-up, plain language, Ctrl+Shift shortcuts + the what's-new box, Desk refinements, Undo my last log entry, Find a person. ThisWorkbook and all nine forms are re-pasted; 25 of 25 stamps; SelfTest 581 / 68; SmokeTest 19 steps; pytest 595.

## Review record
Four read-only reviewers in isolation worktrees (compile traps; gates / rules / privacy; events / state / the screen; completeness), then one re-review of the fixes. Two HIGH found and fixed before the tag: two Consts declared after procedures in modDesk (would not have compiled; no checker sees it - a reviewer lens to keep), and the partial-acceptance close bypassable by re-recording Accepted on a row logged before the close (a hole since v25 through a list click). The CHANGELOG's v25.5 review paragraph lists the rest.

## The members' side
Unchanged by v25.5 apart from who may publish (supervisors too). M3 (the members' ease-of-use follow-up) is in v26.
