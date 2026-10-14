# Call Out Roster v24.3.3 - install

One finding from Graeme's Oct 14 run of the test guide on v24.3.2 (row C6). Install it the same way as v24.3.2.

Needs **v24.3.2 installed**. Full module, paste-in, member and updater set in the zip. **The code change is in `modPublish` only.** ThisWorkbook, every form, the members' template and the Roster Updater are not touched. Every module's version stamp moves to v24.3.3 so SystemCheck can tell the file apart, which is why the updater replaces all 23 (and the by-hand route imports all 23).

**Back up first. Do it on a copy before the real roster.** Close the VBA editor before the updater runs.

## A1. With the Roster Updater
1. Roster copy and Roster Updater.xlsm open in the same Excel window; Ctrl+S in the roster; click the roster.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > the unzipped v24.3.3 folder. Expect **0 new**.
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK.
4. Alt+F11, the ROSTER's project, **Debug > Compile VBAProject**. No message: Alt+F8 > `UpdateFinished` > Yes. A message: note the line, `UpdateDrop`, send me the line.
5. Close and reopen the roster.

## A2. By hand
Remove and import all 23 stamped modules from `modules/` (the list in the v24.3.1 note). Only `modPublish` and `modSystem` carry changed code; the rest changed their version line. Nothing to re-paste. Debug > Compile: no message.

## B. After the code is in
1. **SystemCheck**: "Modules: 23 of 23 on v24.3.3", no ThisWorkbook or form line.
2. **SelfTest**: 432 passed, 50 suites (no rule changed).
3. **MakePracticeCopy** > **SmokeTest** in the copy: PASSED, 14 steps. Close that copy without saving.

## C. The check (fresh practice copy made from the roster you installed on)
1. **Member View** once, then open `PRACTICE - Call Out Roster - Members.xlsm` beside the copy and leave it open. Back in the roster: **Sick Call** for someone, then **Member View** again. Expect one box: *"The members' file is open in this Excel: ... Close it and publish?"* Yes: it closes, the publish runs, the usual "Members' file published" box. Open the members' file again: the new sick call is on the card. **No file with a date in its name appeared** in the folder.
2. Same, but answer **No**: nothing is written, the open file stays open.
3. (Only if you have a second PC or can open the members' file read-only from another Excel instance.) With it open there, Member View in the roster: *"NOT published. The members' file couldn't be written ... It is probably open on a plant PC ..."*, and no dated copy.

Then rerun **C6** from the guide (four sick calls, publish, Find me) and the two Calendar rows **C11 / C12**, which were not in the table you used. Tell me "v24.3.3 OK" or send the boxes.
