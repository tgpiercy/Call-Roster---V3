# Paste this into a fresh session to carry on (written Oct 21, 2026: v26.5 is ON THE REAL ROSTER; v26.5.5 built, reviewed three times, tagged and pushed; next = Graeme installs v26.5.5 before Nov 1, and sends the two Kronos extracts, then the alpha Nov 1-15, then v27 before the Dec 1 training, v28 in January)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first (the numbering rule, the three gating levels,
the 20 SmokeTest steps, the timeline, the "Decisions already made" bullets for Oct 21 and Oct 19-20, the "16 h
rule" row of the not-free table). Then the v26.5.5 entry at the top of CHANGELOG.md (20 items + two review
rounds), docs/INSTALL-v26.5.5.md, briefs/steps-v26.5.5.txt, briefs/v27.md, docs/timeline-2026-10-18.md. Confirm
the four checkers (py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck / callcheck) and
py -3.14 -m pytest tests -q are green on main (713) and the tree is clean. vbacheck's only hits must be the
three named false positives (Sort_Hours_Dept, frmLeave, modMember: one-line Subs). Tags: `git ls-remote --tags
origin` should show v26.5 and v26.5.5 (if v26.5.5 is missing, the push from Graeme's Terminal panel is still
owed: `git push origin main v26.5.5`).

Where things stand: v26.5 is INSTALLED on the real roster (Graeme, Oct 21) and a real Kronos export has been
imported. v26.5.5 is BUILT and TAGGED (dist/Callout_Roster_v26.5.5.zip; docs/INSTALL-v26.5.5.md): the integrity
fixes, the 3 AM fixes, THE 16 H RULE (a rule change), the text pass, the trust strip, the Practice button. It
re-pastes ThisWorkbook and every form but frmEmployee; no members' change; SystemCheck "25 of 25 on v26.5.5",
SelfTest 698 / 81, SmokeTest 20 steps. NOT installed yet: ask first whether it is on the real roster.

DECIDED BY GRAEME Oct 21 (after the reviews; don't re-ask): a person the 16 h rule caps is OFFERED their
maximum, so a refusal / left message is charged on those hours (max 4 h -> 7 h, not 21) and a posting they
did not sign up for charges the hours they could have taken (none = not charged); a capped person who takes
their whole maximum (P 4 of 12) CLOSES the shift like any partial (v25 D-E4). KRONOS (v27 item, his words):
"I want to send you 2 files that are extracted from Kronos. We will then import this data to fill in premium
hours. The change in hours day to day is logged" - i.e. the Worked Periods / Worked (H) figures come from two
Kronos extracts he will send, with a day-to-day delta kept; the v26.5 export import is the fall-back until
then. FIRST THING NEXT SESSION: ask for the two files (uploads folder, never in the repo), read their LAYOUT
only, and write the import's brief (which file feeds H, which the per-period hours, what the daily delta log
is for) before any code.

THEN the install: he runs the panel on a copy, then the real roster (docs/INSTALL-v26.5.5.md; group 4 has
eleven rows, three of which may not be stageable - the failed save, the two-PC stale copy, a supervisor login:
he notes them). Ask for the SmokeTest box, the IntegrityCheck result, and any group-4 row that failed.

THEN v27 (briefs/v27.md), one master install about Nov 20-25 + ONE members' install (Members M4), built during
the alpha once v26.5.5 has run on the real roster, with the alpha findings folded in: build order Desk v2
(the approved mock-up docs/desk-v2-mockup-2026-10.html) -> the Reconcile sheet (paper runs alongside for at
least one pay period; the differences go into the Nov 15 pack, so this one may be installed early on a copy)
-> the Stats report (Graeme signs the counting rules before any code) -> the receipts (one card after every
save, every form) -> the 3 AM pass (log/calendar reconciliation, Kronos hardening, backups on any open +
RestoreDrill, the audit of in-place edits, the inbox hardening, the calendar's supervisor menu, the outage
print, the session draft, "why isn't Joe listed", the error box's saved line, the pop-up polish) -> the
process items (OT Restricted + the weekly-cap warning, the Rules and Settings sheets, period close / scan
link / disputes, withdraw award / edit posting / post all lines / Exception, ExportTables / data check / unit
notes / digest / handover / switch ledger) -> the Desk tiles, "Since you were last here", the Holidays + Min
Staff pop-ups -> the testing items (IntegrityCheck "Log rows", SmokeTest asserts and steps, the LoadCases
split - it is at 57 KB of the 64 KB limit with 698 cases: SPLIT IT FIRST, five checker classes, drift and docs
tests, the manifest, the install record in the updater, the Opens table) -> language (glossary, legend, guided
first session, order columns on every list, the transition-only paper column) -> Members M4 (me-first Home,
highlight me, How it works, Sign me up from the tile, since-you-last-looked, dNotFree with the "16 h rule" and
"on OT" words, the contract test, the email digest only after Graeme signs the wording; the members' MsgBox
titles - MSGTITLE exempts member/ until then). Four read-only reviewers, one master install, the December
rollover dry run on v27 with RestoreDrill in its group 4.

v26.5.5 leftovers for v27 (small, noted in the reviews): ChargedText / TotalText are not yet in the call-out
summary or the status bar; the practice banner sits over the roster's command bar (top left) and Call_Out's
redraw deletes it from the printed sheet; LoadCache no longer reads the log but NotFreeReason now runs the
16 h readers for every free person (measure a calendar draw and a publish on the real roster: ShowCalDrawTimes);
the stale-copy guard treats a save exactly one hour later as the DST shift (accepted).

AFTER v27: v27.5 = the behaviour-free refactors alone (pure cores extracted + suites; the shared form-builder
helpers; the module splits: modCalendar -> Board / MinStaff / CalImport / Needs / CalMenus, modPeople -> Dates /
PlannedMoves / Former, modPostings -> Mail / Sheet / Sync); and the CLAUDE.md split into CLAUDE.md +
docs/workbook-map.md + docs/decisions.md + docs/settings.md (review part 0), in a session with no code change.
v28 in January after the rollover, with go-live: members' requests (switches + time off), the phone page,
v25.4 Payroll # as the key, the P: paths.

Decided (don't re-ask; CLAUDE.md has the list): Oct 21 - only the work BEFORE an offer counts toward the 16 h,
the work after it is governed by the 11 h gap to the next work; the master is on a share (FileDateTime guard);
"Days" -> "Working days elsewhere". Oct 19-20 - 16 h max consecutive hours; no forced OT; Left Message charged
at once; personal logins; a stale copy blocks; backups on any writable open (v27); sick-call undo admin-only;
paper alongside for a pay period; receipts on every form; the weekly cap warns with a logged override (v27).

How every update ships (CLAUDE.md): report back first with the (verify) items and the open points as multiple
choice; one commit per item; the four checkers after each; read-only Opus reviewers (give them the commit range
and "quote file:line"; they may not edit); fix, re-review the fixes; stamps, INSTALL note, steps file, CHANGELOG,
tag, zip (tests/build_zip.py vNN vPREV on a clean tagged tree), push from Graeme's Terminal panel. Never git
add -A. Update docs/NEXT-SESSION-PROMPT.md at the end of the session.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## What v26.5.5 holds (built, reviewed and tagged Oct 21; `docs/INSTALL-v26.5.5.md`)
Nineteen items from `briefs/v26.5.5.md` plus two review rounds (CHANGELOG items 1-20). The rule change: **the 16 h rule** (`modCalendar` FatigueMaxFor and friends; `NotFreeReason` "16 h rule" / "on OT"; `EmpRec.MaxHours`, `PListRec.MaxHours`; the pop-ups' "max N h"; a capped person offered their maximum and charged on it; switches a hard stop). The integrity fixes: a failed save reported (`SaveIfWritable` a Function, `SaveBlockText`, the Desk header), a stale copy blocks (`RecordDiskStamp` / `StaleOnDisk`, DST-safe), Esc-safe writes (`EnableCancelKey`, callcheck RESTORE), an error in Total Hours refuses to load, Worked Periods in the rename, the NoSave fix, the small ones. The 3 AM fixes: `DefaultShiftDate` (by Hour), the sick-call confirm (`SickFreeFrom`), default hours, Kronos-stale as a red line, `NextStepText`, the text pass (147 titles, MSGTITLE, `UnlockForRepair`), the key labels, `TrustStrip` / `ChargedText` / `TotalText`, the Practice button (`PracticeBannerOnOpen`, `PrunePracticeCopies`). New suites: fatigue, fatblock, postcap, shiftdate, nextstep, practice (698 cases, 81 suites); pytest 713.

## Review record (v26.5.5)
Two read-only Opus reviewers on the 20-commit range (compile / state / install; rules / tests / docs): 3 HIGH, 8 MED, 20 LOW, all taken in one fix commit; a re-reviewer on that commit: 0 HIGH, 4 MED, 5 LOW, all taken; a final check on the second fix commit. The HIGHs: a capped person was charged on the full offer (now offered their maximum); the award pop-up read a stale OT map (now lazy + reset on every log write, reloaded at LOAD); a save refused by the stale-copy guard cleared the call-out form with the wrong instruction (now keeps the entries, `SaveBlockText`). Lessons kept: a pure function that takes pre-computed numbers tests the arithmetic, not the decision - give the readers a core too (`FatigueBlockCore`); a 07:00 serial can sit under 7/24 - compare clock parts with `Hour`; a batch flag needs an error path; `FileDateTime` moves one hour across DST.

## The members' side
Unchanged since Members M3.1 (v26.5). The two members' boxes without a title and the "16 h rule" / "on OT" words in dNotFree come with Members M4 in v27.
