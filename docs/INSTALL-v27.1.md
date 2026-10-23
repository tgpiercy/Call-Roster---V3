# Call Out Roster v27.1 - install (the Install panel)

The refactor release: the code reorganised behind the scenes, **nothing a supervisor or member sees changes**, plus one new admin link (the Desk preview). Its own install before the Nov 1 alpha; v27.5 stays the alpha fixes. Plan and decisions: `docs/v27.1-report-back.md`.

Needs **v27 installed**. Nothing to do by hand before the panel this time: the Roster Updater is unchanged (v4.2), no new form, no members' template rebuild, no run-once macro.

**Back up first. Do it on a practice copy before the real roster** - and on that copy, take the before snapshot first.

## Before the install, on the practice copy only (the proof that nothing changed)
1. Desk > Practice on a copy (on v27). In the copy: Alt+F11, File > Import File > `test-only/modSnapshot.bas` from this folder (outside `modules/`, so the panel never installs it), Ctrl+S.
2. Alt+F8 > `TakeSnapshot` > Yes. Answer **No** to the two print questions (the call-out roster, the Board). It writes `snapshot-v27.txt` beside the copy (hashes only - no names, no phones).
3. Install v27.1 on THAT copy through the panel (below). Group 4's FIRST row is the compare: TakeSnapshot again, then CompareSnapshots, before anything else is done in the copy. Expect **IDENTICAL** except one item, the admin Desk's new "See the supervisor Desk" link. Do both snapshots the same day, with nothing logged in between. On this rehearsal copy, Skip "Make a practice copy" (it is itself a copy) and run SmokeTest in it after the compare.
4. On the real roster: no snapshot. Skip that row.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v27.1` folder.
3. Click the one blue button, read its box, click the next. The swap replaces every changed module, **imports the 11 new ones** (41 of 41) and re-pastes the ten code-built forms. Group 3: System check, Self test. Group 4: the practice copy + SmokeTest, the snapshot compare (practice copy only), each re-pasted form, the Desk preview, the Calendar's and Needs Attention's buttons, IntegrityCheck.

By hand instead: in the VBA editor **replace modCalendar, modPeople and modPostings FIRST** (Remove, then File > Import), then import the 11 new modules, then replace every other module in `modules/`; re-paste the ten forms from `paste-in/` (not frmOvertimeEntry, frmPostingEntry or ThisWorkbook); Debug > Compile; save, close, reopen; then the group 3-4 rows yourself. Importing a new module while the old big one is still in gives "Ambiguous name detected" - replace the old one first.

## What's new
- **The code is split into smaller modules** - modCalendar into modCalendar + modBoard, modMinStaff, modCalImport, modNeeds, modCalMenus; modPeople into modPeople + modDates, modFormer, modPlannedMoves; modPostings into modPostings + modPostingSheet, modPostingSync, modPostingMail. Moved by script, never retyped; `tests/movecheck.py` proves every procedure is word for word what it was.
- **Pure rules with known-answer cases** - the shift-switch hard stops, the Kronos ledger rule, the staffing count (SelfTest gains them: 962 cases in 106 suites).
- **The three longest procedures in named steps** - DrawCalendar, DrawNeeds, SyncKronosData; every procedure is now under 300 lines.
- **The ten code-built forms build through four shared helpers** in modUI.
- **The Desk preview (admin)** - See the supervisor Desk > on the admin Desk; a red PREVIEW pill; < Back to my Desk. This session only.
- **Not shown again**: the What's New box. Logins that saw v27's box are not shown it again (v27.1 has nothing new for them).

## Known, on purpose
- Three internal helpers now appear in Alt+F8's macro list (`AfterPeopleChange`, `LoadCache`, `LoadPostCache`). Running one by hand writes no data (it reloads a cache, or marks the calendar and the members' file to redraw / republish).
- The members' file shows the master's release in its version sentence ("v27.1 (Core v27.1 ...)") after the next publish; display only. SystemCheck's "Errors this release" restarts at 0.
- **The only way back to v27 is the backup copy.** The install never removes a module, so a v27 zip cannot go over v27.1. From now on every release is built on v27.1 or later.
- Gaps found while reading, kept as they are for v27.5 (your decision): a same-day sister swap does not run the 16 h stop; a missed Kronos morning that arrives after the year-to-date report counts twice in Worked Periods (Worked H is right); the staffing count rounds two partial lines up separately and counts an empty " / " part.

## Then
Send the SmokeTest box, the CompareSnapshots box (and `snapshot-compare.txt` if it is not IDENTICAL), the System check line and the IntegrityCheck result.
