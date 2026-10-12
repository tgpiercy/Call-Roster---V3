# Call Out Roster v24.2 + members' M2.1 - install

The install kit: an error log, a smoke test, a practice-copy maker and an updater. Also the Payroll # captured from Kronos, a pay-period range on the Refusals Report, and the fixes from the Oct 12 members' file testing.

Needs **v24.1 installed**. Full module, paste-in, member and updater set in the zip.

**Back up first** (the daily backup next to the roster counts if it ran today). **Do the whole install on a copy before the real roster.**

There are two ways to put the code in. Part A1 is the updater; part A2 is by hand, as before. Use ONE of them.

## A1. With the Roster Updater (try it on a COPY first)
The first updater ran inside the roster and crashed Excel (Oct 12). This one runs from its own workbook, so none of the roster's code is running while it is replaced. It has still never been run in Excel. The roster file on disk is not touched until you say so.

Needs Excel's "Trust access to the VBA project object model" ticked, and the roster opened from an ordinary folder.

1. **Make the updater workbook (once).** New blank workbook > Alt+F11 > File > Import File > `updater/modUpdater.bas` > save as **Roster Updater.xlsm** (Excel Macro-Enabled Workbook), anywhere. Keep it for later releases.
2. If `modUpdater` is still inside the roster from the Oct 12 attempt: Alt+F11, right-click it under the ROSTER's project > Remove > No. Save the roster. (The updater refuses to start while it is there, and says so.)
3. Open the roster copy AND Roster Updater.xlsm. Press Ctrl+S in the roster.
4. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped v24.2 folder. Nothing changes; it lists what it would do. Expect **2 new** (`modErrorLog`, `modSmokeTest`) and the rest "to replace". A legacy module you removed earlier (`modCleanup`, for example) also shows as new and comes back: that is harmless.
5. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. It saves a backup copy next to the roster, swaps the code and stops. **Nothing is saved yet.**
6. Alt+F11, click the ROSTER's project in the list on the left, then **Debug > Compile VBAProject**.
   - **No message:** Alt+F8 > `UpdateFinished` > Yes. It saves.
   - **A message:** write down the line it highlights, then Alt+F8 > `UpdateDrop`. It closes the roster without saving; the file is exactly as it was. Send me the line. Do not press Ctrl+S in between: with a compile error the save block may not hold.
7. Close the roster and open it again. Carry on at part B.

If Excel crashes or anything looks wrong: do not keep a recovered copy, open the roster from its folder (it is unchanged), and use part A2.

## A2. By hand
### 1. Replace the modules that changed
In the VBA editor (Alt+F11), for EACH of these: right-click > **Remove** (No to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`, `modReports`, `modSignups`, `modOTLog`, `modLogLayout`, `Sort_Hours_Dept`, `LoginFullAccess`

(`OT_Posting`, `modMigration` and `modCleanup` are unchanged.)

### 2. Import the two new modules
**File > Import File:** `modules/modErrorLog.bas` and `modules/modSmokeTest.bas`.

Do NOT import `updater/modUpdater.bas` into the roster. It belongs in its own workbook (part A1).

### 3. Re-paste ThisWorkbook and two forms
- Double-click **ThisWorkbook**, Ctrl+A, paste `paste-in/ThisWorkbook.txt`.
- For each of **frmOvertimeEntry** and **frmPostingEntry**: right-click > View Code, Ctrl+A, paste the file of the same name from `paste-in/`. Do not change either form's design.

### 4. Debug > Compile VBAProject
No message expected. If one appears, stop and send me the line it highlights.

## B. After the code is in (both ways)
1. Alt+F8 > **ApplyRosterLayout** (column A becomes the Payroll # column, shown in the admin view only; the command bar is rebuilt beside it).
2. Alt+F8 > **SystemCheck**: expect **"Modules: 22 of 22 on v24.2"** and no ThisWorkbook or form line.
3. Alt+F8 > **SelfTest**: expect **"364 passed, 0 failed"** and "all 45 suites passed".
4. Alt+F8 > **ErrorLogTest**: a box says "This didn't finish ... Error Log line N". That is the test working. Then Alt+F8 > **ShowErrorLog**: the line is there.

## C. The members' template (M2.1)
The members' code changed, and the roster now refuses a template whose code or sheets are older.

**With the updater:** close the template and the members' file, then Alt+F8 > `'Roster Updater.xlsm'!UpdateMemberTemplate` > the v24.2 folder. Answer OK to "Build the members' template in THIS workbook?" and OK to "Template built (Members M2.1 (Jan 5))". It is saved for you.

**By hand:**
1. Open `Call Out Roster - Members TEMPLATE.xlsm`. Remove **every** module whose name starts with `modMember`, then import `member/modMember.bas`. (Its ThisWorkbook code is unchanged.)
2. Debug > Compile VBAProject.
3. Alt+F8 > **BuildMemberTemplate** > OK. It must end "Template built (Members M2.1 (Jan 5))".
4. Ctrl+S, close the template.

Then click **Member View** in the roster. Its message now names the members' version and the file it wrote.

## D. Smoke test (replaces the old list of things to try by hand)
1. Alt+F8 > **MakePracticeCopy**. It saves `... COPY <date>.xlsm` in a `Practice` folder with a copy of the template beside it, and opens it. Check the title bar says COPY.
2. In the copy: Alt+F8 > **SmokeTest** > OK. About a minute. It must end **"SMOKE TEST PASSED: all steps OK, totals unchanged."**
3. If it says FAILED, send me the box (a photo is fine). It names the step.
4. Close the copy. Nothing in it is saved.

In a practice copy Member View writes `PRACTICE - Call Out Roster - Members.xlsm` beside the copy, never the real members' file.

Three things only a person can judge, on the copy or the real roster:
- **Calendar:** open the Calendar tab. Does the month look right?
- **Call Out Roster printout:** does it still fit the page?
- **Members' file:** pick a name on Home. The arrow beside Find me is now visible; the card, Roster, Who's next, Calendar and Open Postings follow the person.

## E. What changes day to day
- **Errors.** A macro that fails now says "This didn't finish ... Error Log line N" instead of showing the VBA debug box, and any pop-up it happened in closes. Needs Attention and SystemCheck show a line while there are errors nobody has looked at. `ShowErrorLog` opens the list. The same lines are also written to `Roster error log.txt` beside the roster (logins, procedure names and error text; no hours, phones or leave details), because an error that blocks saving would otherwise take its own log line with it. Still on the VBA debug box for now: the eight department sort buttons and the calendar's right-click menu.
- **Payroll #.** The next Import Kronos fills roster column A for everyone it matches by name, and lists who still has none. From then on a row is matched by its number first. "Same Payroll #, different name" means nothing was changed for that row: rename the person (Employee > Rename) or correct column A. If the Kronos sheet has no "Payroll #" header the import says so and works as before. Column A is hidden in every saved copy and is never published.
- **Refusals Report** asks for the pay periods after the unit. It opens on year to date.
- **Members' file.** A file filled by an older roster, or new code on old sheets, now says so in red on Home instead of showing wrong figures. A sign-up can be sent for another name; a crowded calendar day shows "... and N more".

## Not in this release
Unit moves entered ahead of time (v24.3).
