# Call Out Roster v24.3 - install

Unit moves entered ahead of time (planned moves), and the fix for the Roster Updater saying "No roster is open".

Needs **v24.2 installed**. Full module, paste-in, member and updater set in the zip. The members' template is **not** touched by this release.

**Back up first** (the daily backup next to the roster counts if it ran today). **Do the whole install on a copy before the real roster.**

## 0. Replace the updater's own code first (2 minutes)
The updater that said "No roster is open" is fixed, but it cannot update itself.

1. Open **Roster Updater.xlsm**. Alt+F11.
2. Right-click `modUpdater` > **Remove** > No.
3. **File > Import File** > `updater/modUpdater.bas` from this zip.
4. Ctrl+S. It is now "Updater v3".

**What changed in it.** It now recognises the roster by its OT Log table as well as by its code. If it still does not recognise any workbook, the box lists every workbook Excel can see with the reason beside each, and asks whether the workbook in front is the roster: **Yes** carries on. Send me a photo of that box if you see it, either way. It tells me why v2 missed the roster.

Then use part A1 or part A2, not both.

## A1. With the Roster Updater
1. Open the roster copy AND Roster Updater.xlsm. Press Ctrl+S in the roster. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped v24.3 folder. Nothing changes. Expect **0 new** and the rest "to replace".
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. It saves a backup copy, swaps the code and stops. **Nothing is saved yet.**
4. Alt+F11, click the ROSTER's project on the left, then **Debug > Compile VBAProject**.
   - **No message:** Alt+F8 > `UpdateFinished` > Yes. It saves.
   - **A message:** write down the line it highlights, then Alt+F8 > `UpdateDrop`. Send me the line. Do not press Ctrl+S in between.
5. Close the roster and open it again. Carry on at part B.

If anything looks wrong, close without saving and use part A2.

## A2. By hand
### 1. Replace the modules
Every module carries the v24.3 stamp, so SystemCheck wants all 22. In the VBA editor (Alt+F11), for EACH of these: right-click > **Remove** (No to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`, `modReports`, `modSignups`, `modErrorLog`, `modSmokeTest`

The ones whose code really changed are `modPeople`, `modOvertimeEntry`, `modPostings`, `modPublish`, `modCalendar`, `modIntegrity`, `modSystem`, `modSmokeTest`, `modSelfTest` and `Call_Out_Roster`; the rest changed only their version line.

No new modules. Do NOT import `updater/modUpdater.bas` into the roster.

### 2. Re-paste two forms
For each of **frmEmployee** and **frmOvertimeEntry**: right-click > View Code, Ctrl+A, paste the file of the same name from `paste-in/`. Do not change frmOvertimeEntry's design. ThisWorkbook and the other forms are unchanged.

### 3. Debug > Compile VBAProject
No message expected. If one appears, stop and send me the line it highlights.

## B. After the code is in (both ways)
1. Alt+F8 > **SystemCheck**: expect **"Modules: 22 of 22 on v24.3"**, no form line, and a new line "Planned unit moves waiting: 0".
2. Alt+F8 > **SelfTest**: expect **"392 passed, 0 failed"** and "all 47 suites passed".

No macro has to be run once. The `Planned Moves` sheet is made the first time a planned move is saved.

## C. Smoke test
1. Alt+F8 > **MakePracticeCopy**, then in the copy Alt+F8 > **SmokeTest** > OK.
2. It now has **14 steps**. Step 9 is new: "plan a unit move for tomorrow, then cancel it". It must end **"SMOKE TEST PASSED: all steps OK, totals unchanged."**
3. If it says FAILED, send me the box.

## D. Try a planned move yourself (on the practice copy, before SmokeTest or in a fresh copy)
1. **ADMIN > Employee > Move to unit.** Pick someone, a new unit, a grade code, and a date **next week**. Calculate.
   The result starts "PLANNED:" and ends "the hours are worked out again on the day".
2. **Save.** The confirm is headed "PLANNED MOVE - nothing changes on the roster today". OK.
3. Check the roster row: codes, Home Unit and hours are exactly as before.
4. **Needs Attention** has a new block, "PLANNED UNIT MOVES", with the person on it.
5. **New Posting** in the NEW unit for a shift after the date: the person is on the list, at the hours they would have after the top-up. On the call-out pop-up and the printed call-out roster that figure carries a **~**. In the OLD unit for the same shift: not on the list. For a shift before the date it is the other way round.
6. Double-click the line on Needs Attention > Yes to cancel. Lists made from now on go back. The box also tells you about any posting made while the move was waiting: on the new unit's the person is marked not chargeable; on the old unit's they are missing, and the box names those postings so you can cancel and re-create one if they should be offered the shift.

To see one applied: enter a move dated **tomorrow**, close the copy, and open it the next day. One box lists what was applied.

## E. What changes day to day
- **A unit move can be entered before the day.** Employee > Move to unit with a date after today plans it. Nothing on the roster changes. From that date the person is on the new unit's call-out and posting lists and off the old unit's. Until the move is applied they are listed in the new unit at the hours they would have after the top-up, so nobody jumps the queue. That figure is marked **~** where a supervisor sees it; the OT Log always records their real total.
- **It is applied by itself.** The first time an admin opens the file on or after the date, the codes, Home Unit, Grade Log and the top-up (`UM-`) go in, with one box saying what was done. The hours are worked out as of the move's date, even if the file is opened a few days late. `ApplyPlannedMoves` (Alt+F8) does the same by hand.
- **A move that cannot be applied stays waiting** and says why (the person is no longer on the roster, for example). It shows on Needs Attention as "DUE - not applied".
- **To change a planned move, cancel it and enter it again.** Needs Attention > Planned unit moves > double-click the line.
- **One move at a time per person.** A second move is refused until the first is cancelled or applied.
- **Rename** carries the planned move with the person. **Remove** cancels it.
- **A crew change entered with a planned move** goes into the Crew History straight away, dated the day of the move. Column E shows the new crew from now on; lists use the old crew until the date. Cancelling puts the old crew back.
- **Open postings in the old unit** for shifts on or after the date show "Exempt - moves <date>" for the person: they are never charged on those.
- **The members' file** shows a planned move only through Who's next, which follows the same lists by date. The Roster table and the person's card stay on today's units until the move is applied.
- **Year end.** A move planned for early January is applied after the rollover, when everyone's hours are zero, so little or nothing is added.

## Not in this release
Editing a planned move in place. Moving the pop-up rules into modules, protection, the reports, calendar speed and the Payroll # key are the next releases, in that order.
