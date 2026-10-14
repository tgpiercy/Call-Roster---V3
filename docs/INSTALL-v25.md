# Call Out Roster v25 - install (the Install panel)

Needs **v24.4 installed** on the roster (the Install Log sheet exists) and **Roster Updater.xlsm** on your PC. Full module, paste-in, member and updater set in the zip, plus `install-steps.txt` at its root (the panel reads it). **This release changes the updater itself** (modUpdater v4.1: a catch-all error handler, the backup taken in the restricted view through the roster's own macro, the template version and the Install Log shown through the roster's macros, a plain message when the roster's code can't run), so that one file goes in by hand first; everything else goes in through the panel. The members' template is not touched.

**Back up first. Do it on a copy of the roster before the real one** - name that copy without the words COPY, DRY RUN, FINAL or Backup (for example `V2 - Call Out Roster TEST.xlsm`; a Windows "- Copy" name makes the practice-copy row refuse). "Trust access to the VBA project object model" stays on.

## What's new in v25 (brief: `briefs/v25.md`; the ten items, each its own commit)
- **No rule lives in a pop-up any more.** The shift-switch stops and payback checks, the time-off trim, the award pop-up's sign-up / award / partial-hours rules and one typed-date parser (every date box: "Oct 12", "12 Oct 2026", a typed time ignored, "Jan 5" more than 90 days back reads as next January, the box pre-fills the full date) are module code with known-answer cases. A return date in the future is refused.
- **A partial acceptance CLOSES the call-out shift** (decided Oct 14: reverses v24.1): nobody below is offered the rest, nobody is charged for it, the pop-up refuses a second Accepted; refusals logged before the partial keep their charge.
- **Every sheet stays protected in every view.** Admins type only into roster B, E, F, H (Worked: typed from the Kronos PDF until the daily import runs), J, K and the unit code columns L:S, and Holidays "Observed on"; totals, period cells and every other sheet go through the buttons; a block of cells can't be sorted by hand (the sort buttons do it). Sheets can't be added, deleted or unhidden by hand. Pre-change copies, practice copies and the updater's backup are saved in the restricted view. SystemCheck shows a "Protection:" line.
- **Planned moves:** a Needs Attention line applies only its own move through two plain questions ("Apply it now?", then "Cancel the move?"); the batch summary and the planned row's Note carry the warnings and the UM- ref; a cancel before the date deletes the never-effective crew row, after it dates the undo today; IntegrityCheck flags two waiting moves for one person and orphan future crew rows.
- **Postings:** one closing test everywhere (closed from the day after the closing date); Reopen refuses a shift before the roster year and names anyone it marks not chargeable; out-of-department candidates are listed in THE call order (grade first).
- **Switches:** inside the 24 h sick window a person can neither work nor give up a shift; a shift given up in a switch relieves the next day's rest rule (they did not work it).
- **Small rules:** a 0-hour leave top-up still logs its LV- row; the past-sick save says how many days it skipped and why; a frozen FINAL / backup copy opened read-only says so; Book Time Off refuses a date before the roster year's Jan 1; IntegrityCheck "Check" lines for a non-date Plant Seniority and for mixed code letters in one unit column.
- **Error frames** on the 18 CalMenu macros, CalendarCheck, ShowHolidays, GoToHoliday, SyncSwitchesToCalendar and the updater's run.
- **Tests:** SelfTest compares case-sensitively; 552 cases in 67 suites (ten new suites for rules that had no known answer); SmokeTest 18 steps (no takers / out of dept / void / reopen, a leave going and returning, the Kronos heading rule on a made-up sheet, a planted leave word stopping the publish); three more static checks on the code.
- **From the v24.4 first run:** the panel's System check passes on the version lines (the untrusted-forms line no longer fails it).

## A. Once by hand: modUpdater v4.1 into Roster Updater.xlsm (3 minutes)
1. Open `Roster Updater.xlsm` alone. Alt+F11. Under its project, right-click `modUpdater` > Remove > No.
2. File > Import File > `updater/modUpdater.bas` from the unzipped folder. Debug > Compile VBAProject: no message. Ctrl+S.
3. Close it.

## B. The panel (everything else)
1. Open the roster copy from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front, ADMIN view on.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25` folder. The Install sheet fills: **1. Prepare** (backup, check), **2. Install code** (swap: every module, ThisWorkbook and every form are replaced - 40 files, of which 24 modules, ThisWorkbook and seven forms changed; compile; finish + reopen), **3. Run once** (System check, Self test, Kronos), **4. Checks** (practice copy + SmokeTest, typing on the sheets, the partial-acceptance close, typed dates, New Posting from the panel, the Install Log).
3. Click the one blue button, read its box, click the next. A FAILED step blocks the rest: **Retry** or **Skip this step...** with a note.
   - **System check** should say **"Modules: 24 of 24 on v25"** and a **"Protection: N of N sheets protected (the Install Log open by design); roster totals locked; tabs locked"** line; the forms line says v25 for frmOvertimeEntry, frmNewPosting, frmPostingEntry, frmShiftSwitch, frmTimeOff, frmLeave, frmSickCall (frmConvert and frmEmployee stay v24.3.1).
   - **Self test:** 552 passed, 67 suites.
   - **Smoke test** in the practice copy: SMOKE TEST PASSED, 18 steps. A step may say SKIPPED with its reason (8 or 9 when no second unit has people, 11 or 15 without the members' template, 13 when the leave rule declines on that copy); that is not a fail. CLOSE THE COPY WITHOUT SAVING, then click Done.
   - **Typing on the sheets** is the one new hand check that matters: the row says what must be refused and what must be allowed, in both views.
4. **Show the log** opens the roster's Install Log sheet.

If Excel is closed part-way: open the roster and the updater again, Alt+F8 > InstallPanel, Refresh. It resumes at the next step.

## C. Then
Tell me "v25 OK" with the Install Log's rows for v25 (a photo of the sheet is enough), or the box of the step that failed. Things to watch in the first week on the real file: a cell you used to type into that is now refused (say which: the list of unlocked cells is deliberate but can grow), and any pop-up date that reads differently from what you typed.
