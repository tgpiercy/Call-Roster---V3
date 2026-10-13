# Call Out Roster v24.3.1 - install

The hotfix from the Oct 13 deep review, on top of v24.3. Install it INSTEAD of re-installing v24.3: the planned-move fixes from your Oct 13 findings (posting lists follow a move) are in it, and every module, form and ThisWorkbook now carries the v24.3.1 stamp, so SystemCheck can tell this file from a v24.3 one.

Needs **v24.3 installed** (your desktop copy). Full module, paste-in, member and updater set in the zip. The members' template and the Roster Updater are **not** touched by this release: nothing to do for either.

**Back up first** (the daily backup next to the roster counts if it ran today). **Do the whole install on a copy before the real roster.** Close the VBA editor (Alt+F11 window) before the updater runs.

Use part A1 or part A2, not both.

## A1. With the Roster Updater (v3, already in Roster Updater.xlsm)
1. Open the roster copy AND Roster Updater.xlsm in the same Excel window. Press Ctrl+S in the roster. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped v24.3.1 folder. Nothing changes. Expect **0 new** and the rest "to replace" (23 modules, ThisWorkbook, 9 forms).
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. It saves a backup copy, swaps the code and stops. **Nothing is saved yet.**
4. Alt+F11, click the ROSTER's project on the left, then **Debug > Compile VBAProject**.
   - **No message:** Alt+F8 > `UpdateFinished` > Yes. It saves.
   - **A message:** write down the line it highlights, then Alt+F8 > `UpdateDrop`. Send me the line. Do not press Ctrl+S in between.
5. Close the roster and open it again. Carry on at part B.

If the updater says "No roster is open", the roster is in a second Excel window: close both, open the roster first from its folder, then the updater from the same window (File > Open). Send me a photo of its box either way.

## A2. By hand (about 40 minutes)
### 1. Replace the modules: 23 of them
Every module carries the v24.3.1 stamp, so SystemCheck wants all 23. In the VBA editor (Alt+F11), for EACH of these: right-click > **Remove** (No to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modOTLog`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`, `modReports`, `modSignups`, `modErrorLog`, `modSmokeTest`

**`modOTLog` is new to this list.** It has never been removed and imported by version before: do it like the others. The six without a stamp (`modLogLayout`, `modMigration`, `Sort_Hours_Dept`, `OT_Posting`, `LoginFullAccess`, `modCleanup`) are in the zip for completeness; replace `Sort_Hours_Dept` too (its code changed), leave the other five.

The ones whose code really changed: `modOTLog`, `modSystem`, `modPublish`, `modAttendance`, `modLeave`, `ImportKronosData`, `modErrorLog`, `modPostings`, `modPeople`, `modIntegrity`, `modCalendar`, `modSelfTest`, `modSmokeTest`, `modGradeLog`, `modOvertimeEntry`, `Call_Out_Roster`, `Sort_Hours_Dept`. The rest changed only their version line.

Do NOT import `updater/modUpdater.bas` into the roster.

### 2. Re-paste ThisWorkbook and all nine forms
Only their version line changed, but SystemCheck checks it. For each of **ThisWorkbook**, **frmOvertimeEntry**, **frmPostingEntry**, **frmNewPosting**, **frmShiftSwitch**, **frmTimeOff**, **frmLeave**, **frmSickCall**, **frmConvert**, **frmEmployee**: right-click > View Code, Ctrl+A, paste the file of the same name from `paste-in/`. Do not change the two designer forms' designs.

### 3. Debug > Compile VBAProject
No message expected. If one appears, stop and send me the line it highlights.

## B. After the code is in (both ways)
1. Alt+F8 > **SystemCheck**: expect **"Modules: 23 of 23 on v24.3.1"**, no ThisWorkbook or form line.
2. Alt+F8 > **SelfTest**: expect **"432 passed, 0 failed"** and "all 50 suites passed".

No macro has to be run once.

## C. Smoke test
1. Alt+F8 > **MakePracticeCopy**, then in the copy Alt+F8 > **SmokeTest** > OK.
2. Still **14 steps**; three grew. Step 5 now checks a sick call draws exactly one calendar line. Step 8 checks the moved person is on the list of every open posting in the new unit they are free for. Step 9 now makes an open posting in the first unit, plans the move (the person is added to that list), puts a filter on the posting list, cancels the move (the row comes off through the filter; the code clears it), then cancels that posting. It must end **"SMOKE TEST PASSED: all steps OK, totals unchanged."**
3. If it says FAILED, send me the box.

## D. The manual test: cancel a planned move on a FILTERED posting list
This is the one piece of code that removes a row from a posted list, and it has never run on real data. SmokeTest runs it with its own filter; do it once by hand on the practice copy too (a fresh one, or the same copy before SmokeTest).

1. **New Posting** in any unit, for a shift about 3 weeks out. Note its number (P-00nn). Leave it open.
2. Open the **Postings** sheet. On the posting LIST table (the right-hand one), click the filter arrow on **Rank** and tick only "1", so most rows are hidden.
3. **ADMIN > Employee > Move to unit.** Pick someone from ANOTHER unit who is free for that shift, the posting's unit, a grade code, and a date of **tomorrow**. Calculate, Save (the confirm is headed PLANNED MOVE). The box names P-00nn under "added to the list".
4. Clear the filter and look: the person is on P-00nn's list, in call order; the ranks run 1, 2, 3 ... with no gap. Put the Rank filter back on.
5. **Needs Attention** > PLANNED UNIT MOVES > double-click the line > Yes to cancel.
6. Check, in this order: the filter is gone (the code clears it); the person is OFF P-00nn's list and the ranks run 1, 2, 3 ... again; every OTHER posting's header row in the left-hand table is still there (count them before and after); Alt+F8 > **IntegrityCheck** has no PROBLEM and no "list rows but no posting row" line.
7. Tell me "filtered cancel OK" or send the boxes.

## E. What changes day to day
- **Opening the file with a planned move due:** an admin is asked "N planned unit move(s) are due: apply them now?" Yes applies them as before (backup first, one box). No leaves them waiting; Needs Attention lists them and the next open asks again. Before this the open applied them by itself, and a move whose check failed blocked saving on every open.
- **Members' calendar:** a sick call or a personal day is published as **"Off"** with the name on the day. The five sick-time counts on the Where I stand card stay. (This reverses the Oct 13 decision to show "Sick".)
- **Member View refuses** a calendar line or posting reason that contains a leave word (STD, LOA, WSIB, LTD, OLV, parental, disability, "leave of absence") or the word sick or personal, naming the line: reword it and publish again. Before this only a bracketed code in one column was caught.
- **Member View and the daily backup refuse a copy:** a dated backup, the FINAL or a "Backup before" file opened to look at never takes today's backup slot and never publishes. A practice copy still publishes PRACTICE - ... beside itself.
- **"Paid sick days left"** is 0 once a 4th occurrence has started or the 3rd has ended (a scheduled shift worked after its last sick day). Four single sick days used to show "1 left".
- **A leave return dated after the roster's last pay period** (Dec 27, 2026 on this year's roster) is refused: "run the year-end rollover first, then log the return".
- **Kronos import** finds `Employee Name`, `Grand Total` and `Payroll #` by their headings and refuses, changing nothing, when a heading is missing or the report layout changed. The first January export is when that matters.
- **A sick call on a shift that isn't the person's scheduled one** draws one calendar line, not two.
- **A posting whose header was deleted by hand** no longer has its number reused; IntegrityCheck names the orphan list rows.
- **Every button that failed with a raw VBA box** (the department sort buttons, SystemCheck, IntegrityCheck, SelfTest, ShowGradeLog, RefreshRosterFormulas, Leave from the calendar menu) now shows "This didn't finish ... Error Log line N", and a failed sort leaves the roster protected. The call-out roster print and the calendar draws log their errors.
- **Cancel Request** that fails part-way blocks saving, like an award that fails.

## Not in this release
The admin-view backup copies saved restricted (v25, once the master has its own folder); the sign-up inbox items (sign-ups stay off); one posting closing test; one typed-date parser; the planned-move follow-ups (two-question box, apply one line only, cancel after the date). All on the v25 list from the review.
