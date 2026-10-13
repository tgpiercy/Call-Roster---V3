# Call Out Roster v24.3.2 - install

Two findings from Graeme's Oct 13 run of the v24.3.1 test guide, fixed on top of v24.3.1. Install it the same way; it is small.

Needs **v24.3.1 installed** (the copy you tested on, or the real roster once v24.3.1 is there). Full module, paste-in, member and updater set in the zip. **ThisWorkbook, eight of the nine forms, the members' template and the Roster Updater are not touched.** One form is re-pasted: **frmNewPosting** (it now refuses a leave word in the Reason). Nothing to do for the template or the updater.

**Back up first** (the daily backup next to the roster counts if it ran today). **Do it on a copy before the real roster.** Close the VBA editor (Alt+F11 window) before the updater runs.

## A1. With the Roster Updater
1. Open the roster copy AND Roster Updater.xlsm in the same Excel window. Ctrl+S in the roster. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped v24.3.2 folder. Expect **0 new**; it lists the modules, ThisWorkbook and the forms "to replace" (only frmNewPosting's code changed; the other paste-ins in the folder are the v24.3.1 ones: replacing them changes nothing).
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. Backup copy saved, code swapped, nothing saved yet.
4. Alt+F11, click the ROSTER's project, **Debug > Compile VBAProject**. No message: Alt+F8 > `UpdateFinished` > Yes. A message: note the line, Alt+F8 > `UpdateDrop`, send me the line.
5. Close the roster and open it again.

## A2. By hand (about 25 minutes)
Every module's stamp moves to v24.3.2 (SystemCheck wants all 23), so remove and import all 23 again from `modules/`, exactly as the v24.3.1 note lists them:

`modOvertimeEntry`, `modOTLog`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`, `modReports`, `modSignups`, `modErrorLog`, `modSmokeTest`

The ones whose code really changed: `modCalendar`, `modPublish`, `modPostings`, `modSystem`, `modSmokeTest`. The other 18 changed only their version line. Leave the six unstamped modules alone. **Re-paste one form: frmNewPosting** (right-click > View Code, Ctrl+A, paste `paste-in/frmNewPosting.txt`). Do NOT import `updater/modUpdater.bas` into the roster.

Then **Debug > Compile VBAProject**: no message expected.

## B. After the code is in
1. Alt+F8 > **SystemCheck**: **"Modules: 23 of 23 on v24.3.2"**, no ThisWorkbook or form line (a frmNewPosting line = it was not re-pasted).
2. Alt+F8 > **SelfTest**: **"432 passed, 0 failed"**, "all 50 suites passed" (no rule changed).
3. Alt+F8 > **MakePracticeCopy**, then in the copy **SmokeTest**: **"SMOKE TEST PASSED: all steps OK, totals unchanged."**, 14 steps. Step 9 now also changes the words on its posting's calendar line and then the posting's Reason (a leave word refused each time, then line and Reason changed together) before it cancels the posting. The box ends "Close it WITHOUT saving ... Make the next practice copy from the real roster". Close that copy without saving.

No macro has to be run once.

## C. What changed, and the checks to do by hand (on a fresh practice copy made from the roster you installed on)
"The roster you installed on" = the file with v24.3.2 in it, whatever its name, as long as the name has no COPY / DRY RUN / FINAL / Backup in it (if Windows named your install copy "- Copy", rename it: those words mark a practice or backup file to every check). Make the practice copy from it, not from another practice copy.

1. **New Posting refuses the word.** New Posting, Reason "Covering J. Smith STD", Create: "Please fill in: Reason without 'STD' (members see it; Member View would refuse it)". It is not created.
2. **Change the text of this line.** To get a posting with the word in (as the Oct 13 test had): on the Postings sheet, in admin view, type "Covering J. Smith STD" into an OPEN posting's Reason cell, and on the Coverage sheet into its line's Need cell (the row whose Source is that P-00nn), then Member View. Refused, and the box ends *"That is calendar line CV-00nn (from P-00nn): on the Calendar, double-click the line > Change the text of this line."* Do that: type "Covering J. Smith LOA" > OK: refused, the word named. Double-click the line again, type "Covering J. Smith" > OK: "The line now reads ... Posting P-00nn's Reason too". Member View publishes. (Cancelling a posting does not clear the words: the line keeps them with NCN as the coverage. That is what the Oct 13 test saw.)
3. **Change the Reason of posting.** Type the STD text into an open posting's Reason cell only (Postings sheet). Member View: refused, *"That is posting P-00nn's Reason: on Needs Attention (or its calendar line), double-click the posting > Change the Reason of posting P-00nn."* Needs Attention > that posting > double-click > Change the Reason of posting > type "Covering J. Smith" > OK. Member View publishes.
4. **MakePracticeCopy inside a copy.** In the practice copy, Alt+F8 > MakePracticeCopy: refused, "This file is itself a copy ... run MakePracticeCopy from the roster you are testing." Nothing made. (In a copy where SmokeTest ran it says saving is blocked there instead.)

Then tell me "v24.3.2 OK" or send the boxes. The v24.3.1 test guide (`docs/v24.3.1-install-and-test-guide.md`) is corrected for both and has the full list again (C3's expectation was wrong; rows C11 and C12 are new).
