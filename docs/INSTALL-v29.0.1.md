# Call Out Roster v29.0.1 - install (the Install panel) - THE ALPHA'S MOCK FILE ONLY

v29 with ONE setting changed: `KRONOS_PATH` (modOvertimeEntry) names the local folder the mock Kronos reports are copied
into during the alpha of Nov 1-14, 2026 - `C:\Users\piercyg\Desktop\Kronos mock\` (Graeme, Oct 28: a local path for now).
Nothing else changes; the What's New box is not shown again. **It goes on the MOCK file only** (the real roster is set
aside under `Pre-alpha\` before this install - docs/alpha-data-reset.md page 1 - and stays on v29 with the P: path).

Needs **v29 installed** (on the real roster since Oct 28) and **Roster Updater.xlsm v4.2**. The updater does not change.
The panel's swap replaces every module (all 45 stamps move to v29.0.1); ThisWorkbook, the forms and the members'
template are not touched.

**Order on Oct 31 (Day 0 of docs/alpha-test-plan.md):** the real file set aside FIRST (page 1 of the reset runbook),
then a practice copy of what is now the mock-to-be file, this install on the copy, then on the file itself, then
`LoadMockData` (the test-only module, below), then the Day-0 Kronos import.

## What changed since v29 (`CHANGELOG.md` has the detail)
- `KRONOS_PATH` = `C:\Users\piercyg\Desktop\Kronos mock\`. Make that folder before the install; copy
  `00 cumulative 2026-01-11 to 2026-10-30.xlsx` into it (from `tests/mock/kronos/`, made by `py -3.14 tests/mock/make_kronos.py`).
  SystemCheck's paths box must read "Kronos reports: C:\Users\piercyg\Desktop\Kronos mock\ - reachable". If the login
  on that PC is not `piercyg`, say so: the path is the one constant and moves in a minute.
- No rule changed: SelfTest 1133 / 121, SmokeTest 22, the four checkers at baseline.

## The panel
1. Open the roster (ADMIN view) from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v29.0.1` folder. Groups: **1. Prepare**, **2. Install code**, **3. Run once** (System check, Self test), **4. Checks**.
3. Click the one blue button, read its box, click the next.
   - **System check:** "Modules: 45 of 45 on v29.0.1"; the paths box names the Kronos mock folder as reachable.
   - **Self test:** 1133 passed, 121 suites.
4. **Show the log** opens the roster's Install Log sheet.

By hand instead: remove and import all 45 stamped modules from `modules/`; leave the six unstamped modules, ThisWorkbook
and the forms alone; Debug > Compile; run the group 3 rows yourself.

## The test-only module (not installed by the panel)
`test-only/modMockData.bas` in the zip is the alpha's mock-plant loader. File > Import it by hand into the mock file
ONLY (never into the real roster, never into `modules/`); SystemCheck does not count it. Alt+F8 > `LoadMockData`
(type `MOCK` on the live-named file) does the rest - docs/alpha-data-reset.md page 1, step 8; the rehearsal on a
practice copy first is plan step D-1.1. Remove the module at the close-out, or leave it: the file it sits in becomes
the `Backup - ALPHA MOCK ...` copy and never runs again.

## Then
Tell me "v29.0.1 OK" with the Install Log's rows and the paths box's Kronos line, or the box of the step that failed.
