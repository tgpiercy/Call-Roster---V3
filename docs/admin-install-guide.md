# Call Out Roster: Admin Install Guide

How the admin installs an update, checks it, and looks after the roster file.

Current as of v24.3.1 (Oct 2026).

## 1. Before you start

| Check | Detail |
|---|---|
| Install freeze | No installs from Dec 20 to Jan 4, except the year-end rollover. |
| Install freeze | No installs in the week before Jan 23, 2027. |
| Who is in the file | Everyone else must be out. You need the writable copy. |
| Your login | Your Windows login must be on the admin list (section 7). |
| What you need | The release zip and its INSTALL note, and `Roster Updater.xlsm` (section 3A). |
| Time | About 5 minutes with the Roster Updater, about 40 by hand. |
| The VBA editor | Close it (Alt+F11 window) before running the updater. |

## 2. Where the releases are

| Item | Where |
|---|---|
| Release zips | `dist/Callout_Roster_vNN.zip` in the code repository |
| What a zip holds | `modules/` (the .bas files), `paste-in/` (code for ThisWorkbook and each form), `member/` (the members' file code), `updater/` (the Roster Updater's own module) and `INSTALL.md` |
| Full set | From v22 on, every zip holds the full set of code. It can serve as the master copy. |
| INSTALL notes | `INSTALL.md` inside each zip. The same note is `docs/INSTALL-vNN.md` in the repository. |
| What each release does | `CHANGELOG.md` |

The INSTALL note lists only what changed. Follow it exactly. This guide gives the pattern behind every note.

## 3A. With the Roster Updater (since v24.2)

`Roster Updater.xlsm` is a separate workbook that holds one module, `updater/modUpdater.bas`. It is never imported into the roster (the first version ran inside it and crashed Excel). It replaces the roster's code from a release folder, module by module, with a backup first and saving blocked until you say the compile passed.

| Need | Detail |
|---|---|
| Trust access | File > Options > Trust Center > Trust Center Settings > Macro Settings > tick "Trust access to the VBA project object model" (section 9). Without it the updater cannot see the roster's code. |
| Both files open | The roster (a copy first) and Roster Updater.xlsm in the SAME Excel window. A roster opened from Explorer or an email can land in a second Excel; then the updater says "No roster is open" and lists what it can see. |
| The VBA editor closed | Close Alt+F11 before `UpdateFromFolder`. |
| The roster's own password | None. The updater refuses a password-locked project. The roster has no VBA password (decided Oct 13, 2026). |
| Its admin gate | Advisory when the roster's code cannot run (it asks "Continue?"). Only an admin should answer Yes. |

1. Unzip the release. Open the roster copy, press Ctrl+S in it, open Roster Updater.xlsm, click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!CheckUpdateFolder` > pick the unzipped folder. Nothing changes; it lists what it would replace.
3. Alt+F8 > `'Roster Updater.xlsm'!UpdateFromFolder` > same folder > OK. It saves a backup copy, swaps the code and stops. Nothing is saved yet.
4. Alt+F11, click the ROSTER's project, Debug > Compile VBAProject. No message: Alt+F8 > `UpdateFinished` > Yes (it saves). A message: note the line, Alt+F8 > `UpdateDrop` (it puts the old code back), send the line.
5. Close the roster and reopen it. Carry on at section 4.

When the updater's own code changes, the INSTALL note says so: open Roster Updater.xlsm, remove `modUpdater`, import `updater/modUpdater.bas`, save. `UpdateMemberTemplate` rebuilds the members' template from `member/` the same way.

## 3. The install pattern (by hand)

Do the steps in this order. Skip a step only when the INSTALL note has nothing for it.

1. **Back up.** Save a copy of the workbook somewhere safe. Today's daily backup counts if it ran today.
2. **Open the VBA editor.** Press Alt+F11.
3. **Replace the modules the note lists.** For each one: right-click the module, choose Remove, and answer No to export. Then use File > Import File and pick the file from `modules/`.
4. **Import any new modules** the note lists, with File > Import File.
5. **Re-paste ThisWorkbook and forms the note lists.** Open the code window (F7). Press Ctrl+A. Paste the whole file from `paste-in/`.
6. **Add any new forms.** Choose Insert > UserForm. Set its (Name) in the Properties window to the name in the note. Leave the form blank. Open its code (F7) and paste the file.
7. **Check the settings.** See section 8. A fresh `modOvertimeEntry` brings the settings that shipped with the release.
8. **Compile.** Choose Debug > Compile VBAProject. It should finish with no message.
9. **Run the one-time macros the note lists.** Press Alt+F8, pick the macro, click Run. Run them in the order given.
10. **Save, close and reopen** the workbook.
11. **Run SystemCheck** (Alt+F8). See section 4.
12. **Run SelfTest** (Alt+F8). See section 5.
13. **Update the members' template** if the note says so. See section 10.
14. **Make a practice copy and run the smoke test.** Alt+F8 > `MakePracticeCopy` (it saves a `... COPY <date>.xlsm` in a `Practice` folder with the members' template beside it and opens it). In the copy: Alt+F8 > `SmokeTest` > OK. It must end "SMOKE TEST PASSED". Then do whatever the INSTALL note asks you to try by hand in the same copy. A practice copy can never reach the real roster or the real members' file, and it cannot be saved after SmokeTest.

| If this happens | Do this |
|---|---|
| Compile stops on a line | Note the line and the message. Do not use the file until it compiles. Go back to your backup if the roster is needed. |
| A module shows twice, such as `modCalendar1` | You imported without removing first. Remove the copy with the number. |
| A one-time macro says "Saving is blocked" | See section 11. |

### One-time macros you may be asked to run

The ones that change data check the roster totals before and after. They block saving if the totals do not match. Some save a backup copy first and say where it is.

| Macro | What it does |
|---|---|
| `RefreshRosterFormulas` | Rewrites the roster's total formulas and checks every row against the OT Log. |
| `ApplyRosterLayout` | Rebuilds the button bar and the roster formatting. Needed when a release adds a button. |
| `SetupHomeUnit` | Adds the Home Unit column (B) and fills in anyone with a code in exactly one unit. |
| `UpgradeAbsenceColumns` | Adds the Home Unit, Leave Group and Shift columns to the Absences table. |
| `UpgradeOTLogLayout` | Moves an old OT Log to the current column layout. |
| `SetupPeople` | v24: adds the Crew History sheet (one row per person), the Former sheet and the Signup Hours column on the posted lists. Saves a backup first. |

No release since v24 has needed a one-time macro: the Planned Moves, Error Log and Refusal Report sheets are made on first use.

### Macros you run when needed

| Macro | What it does |
|---|---|
| `SystemCheck` | Section 4 |
| `SelfTest` | Section 5 |
| `IntegrityCheck` | Section 13. Reads only. |
| `MakePracticeCopy` | Saves and opens a practice copy (COPY in its name) in a `Practice` folder, with the members' template beside it. |
| `SmokeTest` | In a practice copy only: 14 steps through the roster's own functions; every total must come back. Never saves. |
| `ShowErrorLog` | Opens the Error Log sheet and marks every line seen. `ErrorLogTest` writes one test line so you can see the box a supervisor gets. |
| `ApplyPlannedMoves` | Applies the planned unit moves that are due (the open asks the same question). |
| `ApplyPlannedMoves` / Needs Attention > Planned unit moves | Cancel a planned move, or apply one that is due. |
| `RolloverYear` | Once a year, with a second person. See docs/rollover-one-pager.md. Never part of an install. |

## 4. SystemCheck

Run it after every install. Problems are listed first, under CHECK. With no problems it says "Everything looks right."

| Line | Meaning | What to do |
|---|---|---|
| Modules: N of N on vNN | Every checked module carries this release's stamp (23 since v24.3.1). | Nothing. Both numbers must match the INSTALL note. |
| [module] is ...: remove it and import the vNN file | That module is from an older release. | Remove it and import the new file. |
| ThisWorkbook code is older than vNN | ThisWorkbook was not re-pasted. | Re-paste `ThisWorkbook.txt`. |
| [form] code is older than vNN | That form was not re-pasted. | Re-paste its `.txt` file. |
| form [name] is missing | A form was never added. | Insert > UserForm, name it, paste its file. |
| '[name]' looks like a duplicate from an import | A module was imported twice. | Remove the duplicate. |
| '[name]' is left over from an earlier version | An old module is still there. | Remove it. |
| Forms / leftovers: not checked (VBA project access is off) | Excel will not let the check read the code. | Turn on VBA project access (section 9). Module versions are still checked without it. |
| OT Log is in the old layout | The log needs its one-time upgrade. | Run `UpgradeOTLogLayout`. |
| Kronos data is old | The Kronos data is more than 7 days old, or was never imported. | Click Import Kronos. |
| Total Hours doesn't include provisional accepted OT yet | The total formulas are from before v21. | Run `RefreshRosterFormulas`. |
| [holiday] is on a weekend | The weekday off is not entered. | Enter it in "Observed on" on the Holidays sheet. |
| Home Unit column not set up | Column B is not ready. | Run `SetupHomeUnit`. |
| N with no Home Unit (column B) | Those people have no Home Unit. | Pick it from the drop-down in column B. |
| Absences table needs its v22 columns | The table is from before v22. | Run `UpgradeAbsenceColumns`. |
| Crew History / Former sheets not set up | v24's one-time setup has not been run. | Run `SetupPeople`. |
| posting lists need their Signup Hours column | The same. | Run `SetupPeople`. |
| sign-up inbox folder not found | `SIGNUPS_ON` is True and the folder isn't there yet. | Click Member View (it makes the folder), then give members write access to it. |
| N new line(s) in the Error Log | A button or pop-up failed part-way since you last looked. | Run `ShowErrorLog`, read the lines, send them on. |
| Planned unit moves waiting: N | Information: moves entered for a later date. | Nothing. Needs Attention lists them; a due one is offered on open. |
| Backup: off - this file is itself a backup or frozen copy | You opened a dated backup, the FINAL or a "Backup before" copy. | Nothing is wrong. Close it and open the live roster to log anything. |
| N imported vacation line(s) ... not booked yet | Old-calendar lines are display only. | Click Convert imports. |
| members' template not found | The template is not next to the roster. | See section 10. |
| BACKUP_FOLDER isn't set | Daily backups are off. | Set it (section 6). |
| backup folder can't be reached | The folder is missing or offline. | Check the folder and the setting. |
| no backup copy taken today yet | No admin has opened the writable file today. | It is taken on the next admin open. |
| This copy is read-only | Someone else has the file. | Open it again when they are out. |

| archive folder for the year-end frozen copy can't be reached | `ARCHIVE_FOLDER` points at a missing folder. | Fix the setting or create the folder before the rollover. |
| this year's reports folder doesn't exist yet | The roster year changed and the new year's folder is not there. | Create the folder named in the line. PDFs go next to the workbook until then. |
| the integrity check found problems | IntegrityCheck found a PROBLEM this session. | Run `IntegrityCheck` and read the Integrity sheet (section 13). |

The lines under the list are for information: roster year and number of pay periods, integrity result, OT Log row count, Kronos date, data-through date, Holidays sheet, grade change log, planned moves waiting, sign-ups (on or off), members' file and backup.

After a rollover the Kronos line reads "Kronos: awaiting 2027 data (Worked = 0)" until the first export of the new year is imported.

## 5. SelfTest

SelfTest runs known-answer cases against the rules: rotation, holidays, clock-change nights, call order, charging, who is free, provisional hours, leave top-ups, attendance, the pay-period year, ID numbering and the integrity helpers. It works in memory. It writes nothing and saves nothing.

| Line | Meaning |
|---|---|
| SelfTest: N passed, 0 failed of N cases. | Every rule gave its known answer. N must match the INSTALL note. |
| all N suites passed | Every group of rules passed. |
| [suite]: n failed of m | Shown only for a group with a failure. |
| Failed: [suite] / [case]: got "x", expected "y" | A rule in the installed code gave a wrong answer. |
| Stopped by an error | The test could not finish. |
| Nothing was written or saved. | Always shown. |

If any case fails, do not log on that file. Go back to your backup and report the failed lines.

## 6. The daily backup

| Item | Detail |
|---|---|
| When | The first time an admin opens the writable file each day |
| File name | `Call Out Roster yyyy-mm-dd.xlsm` |
| Where | The folder in `BACKUP_FOLDER` |
| `BACKUP_FOLDER = "."` | The folder the roster is in. This is the shipped setting. |
| `BACKUP_FOLDER = ".\Backups"` | A subfolder of the roster's folder. It must already exist. |
| `BACKUP_FOLDER = ""` | Backups are off. SystemCheck says so. |
| A full path | That folder, for example a network folder |
| How many are kept | 30 days (`BACKUP_KEEP_DAYS`). Older copies in that folder are deleted. |
| If the folder cannot be reached | You get one warning. The roster opens as usual. |
| Opening a backup to look at it | Fine. Since v24.3.1 a dated backup, the FINAL or a "Backup before" copy never takes that day's backup slot and Member View refuses to publish from it. Nothing you log in it reaches the roster. |

The dated copies sit next to the roster when the setting is ".". Always open the roster itself to log anything. At go-live the master moves to a supervisor-only folder, with the backups, the FINAL and the error log beside it, and the members' file to a separate members' folder (decided Oct 13, 2026).

## 7. ADMIN_USERS

| Item | Detail |
|---|---|
| What it is | The list of Windows logins allowed to use the admin functions |
| Where | Near the top of `modOvertimeEntry` |
| Format | Login names separated by commas. Upper or lower case does not matter. |
| What it controls | The admin buttons, the admin sheets, the phone and email columns, and the daily backup |
| If a login is not listed | That person sees Log Call-Out, Calendar and the sort buttons only. Admin macros answer "Access denied". |

To add or remove a login:
1. Open the VBA editor (Alt+F11) and open `modOvertimeEntry`.
2. Edit the list inside the quotes on the `ADMIN_USERS` line.
3. Choose Debug > Compile VBAProject.
4. Save, close and reopen.
5. Make the same change in the repository copy. If not, the next install puts the old list back.

Void Entry, Add Adjustment and Employee are open to every login on this list.

## 8. Settings

The settings are constants at the top of `modOvertimeEntry`. Change them there and nowhere else. After any change: compile, save, close and reopen. Passwords and folder paths are not printed in this guide. Read them in the module.

| Setting | Shipped value | What it is for |
|---|---|---|
| `ROSTER_SHEET` | OT Roster | Name of the roster tab. The code also finds it by its headers. |
| `LOG_SHEET`, `LOG_TABLE` | OT Log, tblOTLog | Name of the log tab and its table |
| `CALLOUT_SHEET` | Blank Call Out Roster | The tab the printed call-out roster is drawn on |
| `CALLOUT_PHONE_COL` | 5 | Phone column on the printed roster. Hidden for non-admins. |
| `POSTING_SHEET` | Overtime Posting | The tab the posting sheet is drawn on |
| `SHEET_PASSWORD` | (in the module) | Password that protects the sheets |
| `ADMIN_USERS` | (in the module) | See section 7 |
| `CHARGE_MULTIPLIER` | 1.75 | Rate for charged hours |
| `CHARGE_MAX_GRADE` | 2 | Grades above this are called but never charged |
| `POSTING_USE_GRADE_CUTOFFS` | False | False lists anyone with a code in the unit on a posting |
| `CREW_ALL` | All not scheduled | Default choice in the Crew box |
| `CALLOUT_INCLUDE_DAYS` | True | DAYS people can appear on call-out lists for any shift |
| `DST_USE_ACTUAL_HOURS` | True | True fills in 13 or 11 hours on clock-change nights. False keeps 12 and only warns. |
| `USUAL_SHIFT_HOURS` | 12 | Normal shift length |
| `REPORT_FOLDER` | A network folder ending in `{YEAR}` | Where posting PDFs and calendar PDFs go. `{YEAR}` is replaced by the roster year. If the folder cannot be reached, PDFs go next to the roster. |
| `PUBLISH_FOLDER` | . | Where the members' file is published. "." is the roster's own folder (the setting for now). At the network roll-out it becomes the members' network folder. |
| `PUBLISH_FILE` | (old file name) | The old Member View file. No longer written. |
| `PUBLISH_REASONS` | False | False keeps log notes out of the members' file (True sends them, screened for leave words) |
| `REFUSAL_SHEET` | Refusal Report | The tab the Refusals Report is drawn on |
| `SIGNUPS_ON` | False | True switches on electronic sign-ups from the members' file |
| `SIGNUP_FOLDER` | Signups | The inbox folder for sign-up request files, inside the members' folder |
| `SIGNUP_STALE_DAYS` | (in the module) | A request file older than this is ignored |
| `MEMBER_TEMPLATE_NAME` | Call Out Roster - Members TEMPLATE.xlsm | The template. It must sit next to the roster. |
| `MEMBER_FILE_NAME` | Call Out Roster - Members.xlsm | The published members' file |
| `MEMBER_PASSWORD` | (in the module) | Sheet protection in the members' file. The members' code uses the same one. |
| `MEMBERS_STALE_DAYS` | 1 | Needs Attention reminds you after this many days without a publish |
| `MEMBER_NEXT_DAYS` | 14 | How many days ahead "Who's next" covers |
| `KRONOS_PATH` | A network file | The Kronos workbook that Import Kronos reads |
| `KRONOS_SHEET` | Sorted Data Here | The sheet in that workbook. The import finds the `Employee Name`, `Grand Total` and `Payroll #` columns by their headings (first 10 rows) and refuses, changing nothing, when a heading is missing. |
| `KRONOS_STALE_DAYS` | 7 | Kronos data older than this gives a warning |
| `BACKUP_FOLDER` | . | See section 6 |
| `BACKUP_KEEP_DAYS` | 30 | Days of backups kept |
| `BACKUP_BASENAME` | Call Out Roster | First part of each backup file name |
| `ARCHIVE_FOLDER` | . | Where the frozen year-end copy goes at the rollover. "." is the roster's folder. |
| `PAYBACK_WARN_DAYS` | 42 | A shift-switch payback owed longer than this is flagged by the integrity check |
| `PUB_NAME`, `PROV_NAME`, `AWAIT_NAME` | (hidden names) | Names the workbook uses to remember dates. Do not change them. |

Two other things live outside this module.

| Where | What |
|---|---|
| `modPostings` | The CC lists for award emails, the posting closing time (8:00 AM), the default closing lead (14 days) and the number of blank out-of-department lines (3) |
| `modAttendance` | The attendance limits: 5 paid sick days, 3 occurrences, 5 personal days, 5 combined |

Importing a new copy of a module replaces its settings with the shipped ones. Keep the repository copy right, and check the settings after every install.

## 9. Macro security

| Topic | What to do |
|---|---|
| Where the master lives | Keep it in one fixed supervisor-only folder. Open it from there. The members' file goes to a separate members' folder (decided Oct 13, 2026). |
| VBA password | None, now and at go-live. The sheets and the workbook structure are protected instead (v25). A password would stop the Roster Updater. |
| Never | Open the master from an email attachment or a browser download. Macros may be turned off. |
| Trusted location | Ask IT to make the roster's folder a Trusted Location. The setting is under File > Options > Trust Center > Trust Center Settings > Trusted Locations. |
| Enable Content | If Excel shows the yellow bar, click Enable Content. Without macros, nothing in the roster works. |
| SharePoint or OneDrive | If the file is opened from there, the roster turns AutoSave off. Every entry saves itself. |

### VBA project access

SystemCheck can only check the forms and look for leftover modules when Excel allows it.

1. Choose File > Options > Trust Center > Trust Center Settings.
2. Open Macro Settings.
3. Tick "Trust access to the VBA project object model".

Without it, SystemCheck still checks the module versions. It says "Forms / leftovers: not checked".

## 10. The members' file

### Build the template (once)

1. Start a blank workbook. Save it with the name in `MEMBER_TEMPLATE_NAME`, as a macro-enabled workbook, in the same folder as the roster.
2. Press Alt+F11. Use File > Import File and pick `member/modMember.bas`.
3. Open ThisWorkbook in that project. Press Ctrl+A and paste `member/ThisWorkbook.txt`.
4. Choose Debug > Compile VBAProject.
5. Press Alt+F8, run `BuildMemberTemplate`, and click OK.
6. Save and close it. The template holds no data. That is correct.

### Update the template (when an INSTALL note says so)

1. Open the template. Remove `modMember` and import the new one.
2. Re-paste `member/ThisWorkbook.txt` if the note lists it.
3. Compile, save and close.
4. Click Member View once in the roster.

### Publish

1. Click **Member View** on the roster.
2. Wait for "Members' file published" and the counts.
3. Publish again whenever Needs Attention shows "Members file: last published ...". Double-click that line to publish.

| Message | Meaning | What to do |
|---|---|---|
| The members' template wasn't found | The template is missing or has the wrong name. | Put it next to the roster with the exact name. |
| The members' folder can't be reached | `PUBLISH_FOLDER` is offline or wrong. | Check the drive and the setting. |
| NOT published. Something private would have gone in | An email, phone number, employee number or leave code is in a published field, or a calendar line or posting reason contains a leave word (STD, WSIB, parental ...) or the word sick or personal. | The message names the table, the row and the word. Reword that entry and publish again. |
| This file is a backup or frozen copy of the roster | You clicked Member View in a dated backup, the FINAL or a "Backup before" copy. | Open the live roster and publish there. |
| NOT published. The call positions ... don't match | The publish check failed. Nothing was written. | Run SystemCheck and SelfTest, and report the message. |
| The members' file is open in this Excel ... Close it and publish? | You (or this PC) had the members' file open. | Yes closes it without saving and publishes into it. No writes nothing. |
| NOT published. The members' file couldn't be written | It is open on a plant PC, or read-only in the folder. Nothing was written; members see the old one. | Have it closed there. Publish again. (v24.3.3: no dated copy is saved any more.) |

The members' file is values only. It never reads the master. Members open it on a plant PC and click Enable Content.

## 11. When "Saving is blocked" appears

The roster blocks saving when a check finds that totals do not match after a change. This protects the saved file.

1. Read the message. Write down what it says, or take a screenshot.
2. Close the workbook WITHOUT saving. If Excel asks, choose Don't Save.
3. Reopen it. The file is as it was at the last good save.
4. Run SystemCheck.
5. If a one-time macro caused it, do not run it again until the cause is known. Report the message.
6. If a posting award caused it, do that posting again.

Restore the backup only if the reopened file is wrong.

1. Close the roster. Make sure nobody has it open.
2. Rename the bad file so you can keep it.
3. Copy the newest good backup. Give the copy the roster's own file name.
4. Open it and run SystemCheck.
5. Log again anything entered after that backup was taken.

| Backup | Where |
|---|---|
| Daily backup | `Call Out Roster yyyy-mm-dd.xlsm` in the backup folder |
| Backup taken by a one-time macro or an Employee change | Next to the roster. The name starts with "Backup before". The macro's message gives the full name. |
| Backup taken by the Roster Updater | Next to the roster, named in its report. |

## 12. Admin tools

| Button or macro | What it does |
|---|---|
| **Import Kronos** | Reads worked hours from the Kronos workbook and saves a dated snapshot |
| **Add Adjustment** | Adds or removes charged hours for the person in the selected roster row |
| **Void Entry** | Voids a session, posting or adjustment by its number, or one row by `#` and its Entry ID |
| **Employee** | Add a person, move one to another unit, rename, remove (to the Former sheet) or bring back. Each shows what it will do first, saves a backup where rows move, and checks every total after. |
| Crew column (E) on the roster | Type the new crew. A box asks from which date; Cancel puts the old crew back. Earlier time off and switches stay on the old crew's shifts. |
| **Full Access** | Puts the admin view back for a login on the admin list |
| **Attendance** | Draws the sick and personal day sheet: one row per person, a column per occurrence, sorts and a filter |
| **Refusals Report** (PRINT / PUBLISH) | Refused hours split into call-out and posting refusals, per person and per unit; printable, PDF |
| Status bar (bottom left) | Click a roster row to see that person's Refused hours by kind |
| **Convert imports** | Turns imported old-calendar vacation lines into real bookings |
| Employee > Move to unit with a later date | Plans the move. Nothing changes until the date; the first admin open on or after it asks "apply now?". Needs Attention > Planned unit moves cancels or applies one. |
| `MakePracticeCopy`, `SmokeTest` (Alt+F8) | Section 3, step 14 |
| `ShowErrorLog`, `ErrorLogTest` (Alt+F8) | The Error Log sheet (admin only, never published) |
| `ReopenPosting` (Alt+F8) | Reopens a closed posting after its log rows are voided |
| `ShowGradeLog` (Alt+F8) | Shows the log of grade code changes |
| `SystemCheck`, `SelfTest` (Alt+F8) | See sections 4 and 5 |

## 13. IntegrityCheck

IntegrityCheck reads the roster, the OT Log and the other tables. It reports anything that does not add up. It never changes data.

| When it runs | What you see |
|---|---|
| Each time an admin opens the file | The status line on the roster: "Integrity: OK", "Integrity: OK (n to check)" or "Integrity: n PROBLEMS" in red |
| Alt+F8 > `IntegrityCheck` | The full list on the **Integrity** sheet, with what to do for each line |
| At the start and end of `RolloverYear` | The rollover stops on any PROBLEM |

| Level | Meaning |
|---|---|
| PROBLEM | Totals cannot be trusted until it is fixed. It blocks the rollover. |
| Check | Worth a look. Nothing is wrong with the hours. |
| Note | Information only |

| Check | Level | It reports |
|---|---|---|
| Roster = log | PROBLEM | A pay-period cell, Refused, Provisional or Total that differs from what the OT Log gives |
| Typed values | PROBLEM | A pay-period cell that holds a typed number, text, or no formula |
| IDs | PROBLEM | An Entry ID used twice, the same posting, switch, absence or calendar-line ID on two rows, or a leave return logged twice |
| IDs | Check | One call-out or posting number covering rows with different dates, units or shifts |
| Not on roster | Check | A name in this year's OT Log, an active booking or a coming switch that is not on the roster |
| Postings | Check | An Open posting that already has log rows, a posting whose own list and the log disagree on hours charged, or list rows whose posting row is gone (a header deleted by hand) |
| Payroll # | Check | One Payroll # on two roster rows |
| Planned moves | Check | A waiting move for someone not on the roster, to an unknown unit, with no readable date, or more than 3 days overdue |
| Calendar | Check | A calendar line that points at a posting or switch that does not exist, or a void switch |
| Paybacks | Check | A shift-switch payback owed for more than `PAYBACK_WARN_DAYS` days |
| Kronos | Check | Kronos data older than `KRONOS_STALE_DAYS` |

For most PROBLEM lines the fix is `RefreshRosterFormulas`, which rewrites the formulas and moves typed numbers into the OT Log. Run IntegrityCheck again after any fix. If a PROBLEM stays, do not fix cells by hand. Report the line.

The Integrity sheet is admin only and is hidden in the saved file.

## 14. Year end

The year-end rollover is described in docs/rollover-one-pager.md. It is run once a year by the admin with a second person. Practise it on a copy in the first week of December.

| Sheet the rollover adds | What it holds |
|---|---|
| Year End | Each person's final totals for each closed year. Admin only. |
| OT Log 2026 (one per year) | The closed year's OT Log rows. Locked, even for admins. |

Other admin-only sheets made on first use: Error Log (v24.2), Planned Moves (v24.3), Refusal Report (v24.1), Former and Crew History (v24, by `SetupPeople`).

| Setting | Use |
|---|---|
| `ARCHIVE_FOLDER` | Where the frozen "Call Out Roster 2026 FINAL" copy is saved. "." is the roster's folder. |
| `REPORT_FOLDER` | Keep `{YEAR}` in the path. It is replaced by the roster year, so reports follow the rollover. |

## The Install panel (v24.4)
From v24.4 every release is installed from one sheet of buttons in `Roster Updater.xlsm`: open the roster (ADMIN view) and the updater in the same Excel window, Alt+F8 > `InstallPanel`, pick the unzipped release folder (it holds `install-steps.txt`), then click the one blue button at a time. Groups: Prepare (backup, check), Install code (swap, compile, finish + reopen), Run once (SystemCheck, SelfTest, the release's macros), Checks (SmokeTest, Pass / Fail rows that open the pop-ups pre-filled). Every click writes a row to the roster's hidden **Install Log** sheet (ADMIN view lists it), so closing Excel part-way and opening again resumes at the next step; **Skip this step...** records a note for a step done by hand. A FAILED step blocks the rest until Retry or Skip. If "Trust access to the VBA project object model" is off, the swap and compile rows show the exact remove / import / paste list instead. The old macros (`UpdateFromFolder`, `UpdateFinished`, `UpdateDrop`, `UpdateMemberTemplate`) still work for a release without a step list. When a release changes the updater itself, its note says to import `updater/modUpdater.bas` into Roster Updater.xlsm by hand first (remove the old module, import, Compile, save).
