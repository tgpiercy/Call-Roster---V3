# Call Out Roster v27.5 - install (the Install panel)

Three rule fixes found while reading for v27.1, decided by Graeme Oct 25, 2026 (`docs/v27.5-report-back.md`, `briefs/v27.5.md`). The alpha findings (Nov 1-15) follow as v27.5.5.

Needs **v27.1 installed** (the install never removes a module, so this zip cannot go over v27). Nothing to do by hand before the panel: the Roster Updater is unchanged (v4.2), no new module, no new form, no form re-paste, ThisWorkbook unchanged, no members' template rebuild, no run-once macro.

**Back up first. Do it on a practice copy before the real roster.**

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v27.5` folder.
3. Click the one blue button, read its box, click the next. The swap replaces every module (every stamp moves to v27.5; the rules changed in modShiftSwitch, modCalendar, ImportKronosData, modMinStaff; the cases in modSelfTest; the Rules sheet text in modRules; one Help sentence in modDesk; the release constants in modSystem). Group 3: System check (41 of 41 on v27.5), Self test (993 cases, 108 suites). Group 4: the practice copy + SmokeTest (22 steps), the switch stop by hand, the What's New box, IntegrityCheck.

By hand instead: in the VBA editor replace every module in `modules/` (Remove, then File > Import); nothing to paste; Debug > Compile; save, close, reopen; then the group 3-4 rows yourself.

## What's new (the three rule fixes)
- **Every shift switch runs the 16 h stop by the rule's own arithmetic.** The one WORKING a shift must fit the 16 h rule with the shifts beside it - a SCHEDULED shift now counts like overtime accepted or a switch. Nights last night then days today is a hard stop ("can't work Nov 2 days under the 16 h rule ..."), where before it was the "coming off nights that morning" warning that could be saved over. A **same-day sister swap** (days for nights) is checked for BOTH people on the shift each one works; the shift each one gives up is not counted (it is the other's now). The same relief removes a false stop on a two-way payback on the adjoining shift. (Rules sheet EL-4 says so.)
- **A missed Kronos morning never counts twice in Worked Periods** (members' My history, the history report's pay-period block; Worked H itself was always right). At the year-to-date import, when a day between the two Throughs had no one-day report, the differences are logged as **gap** rows in Worked Changes ("no one-day report for 2026-10-14; ...") instead of timecard edits; a gap row counts in Worked Periods only until that person's rows for one of those days are imported, then the day's own rows count. The import's message names the missed days. Known: with two missed days and only one backfilled, that person's period reads LOW until the other arrives (never high).
- **The staffing count adds partial covers once and rounds down.** "Joe D 6h" is 6/12 of a person; the shares of every line on a shift are added once (two 6 h lines = one person, as one line "Joe D 6h / Sam K 6h" always was), a blank " / " part adds nobody, and the calendar's flag and Needs Attention say **(part)** - "SHORT 1/2 (part)" - when a fraction was dropped. A switch saved over the "isn't on that shift" warning no longer counts one high.
- **The What's New box** shows these three lines once per login (v27.5 is a rule change, unlike v27.1).

## Known, on purpose
- For the one WORKING a switch, "coming off nights that morning" / "starts days the next morning" is now the 16 h stop, not a warning: with 12 h shifts the adjoining shift leaves 0 h rest. The rotation warning still shows for the one GIVING UP a shift they are not on.
- The members' file shows "v27.5" in its version sentence after the next publish; display only. SystemCheck's "Errors this release" restarts at 0.
- CAL_VERSION moved, so the Calendar's and Needs Attention's saved buttons are rebuilt on the first draw (as at v27.1).

## Then
Send the SmokeTest box, the System check line, the switch-stop box from the by-hand row, and the IntegrityCheck result.
