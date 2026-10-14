# Call Out Roster v24.4 - install (the Install panel)

The Install panel: from this release on, an install is a list of buttons in Roster Updater.xlsm, clicked in order, each one recorded on the roster's hidden Install Log sheet. v24.4 itself goes in through it, after one step by hand: the panel's code must be in the updater workbook first.

Needs **v24.3.3 installed** on the roster and **Roster Updater.xlsm** (v3) on your PC. Full module, paste-in, member and updater set in the zip, plus `install-steps.txt` at its root (the panel reads it). ThisWorkbook, the forms and the members' template are not touched. One new roster module, `modSmoke` (the panel imports it).

**Back up first. Do it on a copy of the roster before the real one.** "Trust access to the VBA project object model" stays on (it is).

## A. Once by hand: modUpdater v4 into Roster Updater.xlsm (3 minutes)
1. Open `Roster Updater.xlsm` alone. Alt+F11. Under its project, right-click `modUpdater` > Remove > No.
2. File > Import File > `updater/modUpdater.bas` from the unzipped folder. Debug > Compile VBAProject: no message. Ctrl+S (the updater is saved from now on: its Install sheet remembers the folder you picked).
3. Close it.

## B. The panel (everything else)
1. Open the roster COPY from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front, ADMIN view on.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. It asks for the update folder: pick the unzipped `Callout_Roster_v24.4` folder. The Install sheet fills with the steps, grouped: **1. Prepare** (backup, check), **2. Install code** (swap, compile, finish + reopen), **3. Run once** (SystemCheck, SelfTest, Kronos), **4. Checks** (SmokeTest, three pre-filled pop-ups, Member View with the members' file open, the Install Log).
3. Click the one blue button, read its box, click the next. Each step writes OK / FAILED / PASS / FAIL / DONE and the time on its row. A FAILED step blocks the rest: **Retry** (its button again) or **Skip this step...** with a note.
   - **finish** saves, closes and reopens the roster (its open boxes run as usual), then the panel carries on.
   - **System check** should say **"Modules: 24 of 24 on v24.4"**; **Self test** 432 passed, 50 suites (no rule changed).
   - **Smoke test**: the step runs MakePracticeCopy for you; in the practice copy run Alt+F8 > SmokeTest, close that copy WITHOUT saving, come back to the Install sheet and answer Yes.
   - **The three pop-up rows** open Log Call-Out, New Posting and Shift Switch pre-filled; cancel each and click Pass (or Fail with what you saw).
4. When every row is done the status line says so. **Show the log** opens the roster's Install Log sheet (ADMIN view lists it from now on; non-admins never see it).

If Excel is closed part-way: open the roster and the updater again, Alt+F8 > InstallPanel, Refresh. It resumes at the next step (the log is the record).

If "Trust access" is ever off (another PC): the swap and compile rows show the exact remove / import / paste list instead and ask Done.

## C. Then
Tell me "v24.4 OK" with the Install Log's rows for v24.4 (a photo of the sheet is enough), or the box of the step that failed. The C6 rerun from v24.3.3 (four PAST sick days) still stands if not done.

From v25 on each release's note is "what's new" plus this: unzip, open the roster and the updater, Alt+F8 > InstallPanel, pick the folder, click through.
