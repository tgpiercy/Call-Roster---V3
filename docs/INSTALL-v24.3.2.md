# Call Out Roster v24.3.2 - install

Two findings from Graeme's Oct 13 run of the v24.3.1 test guide, fixed on top of v24.3.1. Install it the same way; it is small.

Needs **v24.3.1 installed** (the copy you tested on, or the real roster once v24.3.1 is there). Full module, paste-in, member and updater set in the zip. **ThisWorkbook, the nine forms, the members' template and the Roster Updater are not touched**: nothing to re-paste, nothing to do for the template or the updater.

**Back up first** (the daily backup next to the roster counts if it ran today). **Do it on a copy before the real roster.** Close the VBA editor (Alt+F11 window) before the updater runs.

## A1. With the Roster Updater
1. Open the roster copy AND Roster Updater.xlsm in the same Excel window. Ctrl+S in the roster. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped v24.3.2 folder. Expect **0 new**; it lists the modules, ThisWorkbook and the forms "to replace" (the paste-ins in the folder are the v24.3.1 ones, unchanged: replacing them changes nothing).
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. Backup copy saved, code swapped, nothing saved yet.
4. Alt+F11, click the ROSTER's project, **Debug > Compile VBAProject**. No message: Alt+F8 > `UpdateFinished` > Yes. A message: note the line, Alt+F8 > `UpdateDrop`, send me the line.
5. Close the roster and open it again.

## A2. By hand (about 25 minutes)
Every module's stamp moves to v24.3.2 (SystemCheck wants all 23), so remove and import all 23 again from `modules/`, exactly as the v24.3.1 note lists them:

`modOvertimeEntry`, `modOTLog`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`, `modReports`, `modSignups`, `modErrorLog`, `modSmokeTest`

The ones whose code really changed: `modCalendar`, `modPublish`, `modPostings`, `modSystem`, `modSmokeTest`. The other 18 changed only their version line. Leave the six unstamped modules alone. **Nothing to re-paste.** Do NOT import `updater/modUpdater.bas` into the roster.

Then **Debug > Compile VBAProject**: no message expected.

## B. After the code is in
1. Alt+F8 > **SystemCheck**: **"Modules: 23 of 23 on v24.3.2"**, no ThisWorkbook or form line.
2. Alt+F8 > **SelfTest**: **"432 passed, 0 failed"**, "all 50 suites passed" (no rule changed).
3. Alt+F8 > **MakePracticeCopy**, then in the copy **SmokeTest**: **"SMOKE TEST PASSED: all steps OK, totals unchanged."**, 14 steps. Step 9 now also changes the words on its posting's calendar line (a leave word refused, then the line and the Reason changed together) before it cancels the posting. The box ends "Close it WITHOUT saving ... Make the next practice copy from the real roster". Close that copy without saving.

No macro has to be run once.

## C. What changed, and the two checks to do by hand (on a fresh practice copy made from the roster)
1. **Change the text of this line.** New Posting with Reason "Covering J. Smith STD". Member View: refused, and the box now says *"That is calendar line CV-00nn (from P-00nn): on the Calendar, double-click the line > Change the text of this line."* Do that: type "Covering J. Smith LOA" > OK: refused, the word named. Type "Covering J. Smith" > OK: the line and the posting's Reason change; Member View publishes. (Cancelling the posting does not clear the words: the line keeps them with NCN as the coverage. That is what the Oct 13 test saw.)
2. **MakePracticeCopy inside a copy.** In the practice copy, Alt+F8 > MakePracticeCopy: refused, "This file is itself a copy ... run MakePracticeCopy there." Nothing made.

Then tell me "v24.3.2 OK" or send the boxes. The v24.3.1 test guide (`docs/v24.3.1-install-and-test-guide.md`) is corrected for both and has the full list again (C3's expectation was wrong; row C11 is new).
