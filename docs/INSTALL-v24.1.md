# Call Out Roster v24.1 + members' M2 - install

Refused hours by kind, the new Attendance tab, part-accepted call-outs, electronic sign-ups, and the rebuilt members' workbook.

Needs **v24 installed and `SetupPeople` run**. Install window: with v24 (Jan 5 - 15, 2027) or after Jan 23. Full module, paste-in and member set in the zip. About 40 minutes including the members' template.

**Installing v24 and v24.1 in one sitting?** This zip is a full set. Do the v24 install note's steps with THIS zip's files (stop before its SystemCheck step), then carry on from step 3 below. The numbers to expect are the ones on this page.

**Back up first** (the daily backup next to the roster counts if it ran today).

## A. The roster (master)

### 1. Replace every module
In the VBA editor (Alt+F11), for EACH module, right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`, `modPeople`

(Code changed in `modOvertimeEntry`, `modPostings`, `modShiftSwitch`, `modSystem`, `modUI`, `modSelfTest`, `modPublish`, `modAttendance`, `modIntegrity`, `modPeople`, and `modCleanup` if you still have it: replace that one too. The others are the same code with a v24.1 stamp. `modOTLog`, `modLogLayout`, `modMigration`, `Sort_Hours_Dept`, `OT_Posting` and `LoginFullAccess` are unchanged from v24.)

### 2. Import the two new modules
**File > Import File:** `modules/modReports.bas` and `modules/modSignups.bas`.

### 3. Re-paste ThisWorkbook and two forms
- Double-click **ThisWorkbook**, Ctrl+A, paste `paste-in/ThisWorkbook.txt`.
- For each of **frmOvertimeEntry** and **frmPostingEntry**: right-click > View Code, Ctrl+A, paste the file of the same name from `paste-in/`. Do not change either form's design.

### 4. Debug > Compile VBAProject
No message expected. If one appears, stop and send me the line it highlights.

### 5. Run once: ApplyRosterLayout (Alt+F8)
Adds the **Refusals Report** button under PRINT / PUBLISH.

### 6. SystemCheck (Alt+F8)
Expect **"Modules: 20 of 20 on v24.1"**, no ThisWorkbook or form line, and a new line **"Sign-ups: off"**.

### 7. SelfTest (Alt+F8)
Expect **"320 passed, 0 failed"** and "all 39 suites passed".

### 8. Settings (top of `modOvertimeEntry`)
| Setting | Shipped value | Change it? |
|---|---|---|
| `SIGNUPS_ON` | `False` | Leave it off until the union and management agree an electronic sign-up counts like initialling the posted sheet. See part C. |
| `SIGNUP_FOLDER` | `".\Signups"` | A folder beside the members' file. Leave it. |

## B. The members' template (rebuild it)
The members' file gained data and sheets, so the old template can't be filled any more: Member View will say "the template is older than this roster" until this is done.

1. Open `Call Out Roster - Members TEMPLATE.xlsm` (next to the roster). In the VBA editor remove **modMember** (No to export) and import `member/modMember.bas` from the zip. Then double-click its **ThisWorkbook**, Ctrl+A, and paste `member/ThisWorkbook.txt`.
2. Debug > Compile VBAProject. No message expected.
3. Alt+F8 > **BuildMemberTemplate** > OK. It rebuilds every sheet and button.
4. **Save** the template (Ctrl+S), then close it.
5. In the roster, click **Member View**. It should publish as before, in about the same time.
6. Open the published `Call Out Roster - Members.xlsm` and try part D's members' checks.

## C. Switching electronic sign-ups on (later, not at install)
1. Set `SIGNUPS_ON` to `True` (top of `modOvertimeEntry`), Compile, save.
2. Click **Member View**. The publish makes the `Signups` folder beside the members' file.
3. Give every member **write** access to that `Signups` folder (and read access to the members' file, as now). Supervisors need write access to move files into `Signups\Done`.
4. From then on: members click **Sign me up** in their file; when you LOAD a posting in Award Posting the ticks are already there, with a line saying how many were read. You still confirm each one against the paper sheet.

## D. Smoke test (on a copy with COPY in the name, in its own folder, with a copy of the template beside it)
1. **Refused by kind:** click a person's row on the roster. Excel's status bar (bottom left of the window) shows `Refused 63 h = call-out 42 / posting 21`. The two always add up to the Refused column, plus "other" for anyone with an adjustment.
2. **Refusals Report** (PRINT / PUBLISH): pick 0 for all units. A sheet lists everyone with charged hours, sub-totalled by Home Unit; the last column equals each person's Refused on the roster.
3. **Attendance** (ADMIN): only people with a sick or personal day show, A - Z. Each occurrence has its own column with its dates and paid days. Try **By department**, **Worst first**, **Show everyone**, and the filter arrows on the header.
4. **Part of a call-out:** Log Call-Out, load a list, mark the first person Refused, select the second and click **ACCEPTED PART OF THE HOURS**, enter 8 (of 12). The third person now shows "Offering 4 of 12 h"; mark them Refused and submit. In the OT Log: the first refusal is charged 21, the acceptance shows Hours 8, the last refusal is charged 7.
5. **Out of department, part of a shift:** Award Posting, LOAD a posting nobody signed, pick an out-of-department name, type 6 in the **hrs** box, AWARD OUT OF DEPT. The log row shows Hours 6; in-department people are charged the posting's full hours as before.
6. **Publish reminder:** this one can't be seen on the copy. It is deliberately silent on a file with COPY or DRY RUN in its name, on a daily backup and on a FINAL year file, so an old copy can never overwrite the members' file. Check it on the live roster the first time you create a posting and close: it asks whether to publish first; No closes as usual.
7. **Members' file:** open the published copy.
   - **Find me** on Home: pick a name. The card below shows hours, Refused by kind, call positions and the sick-time counts (no dates). The Roster tab opens on that person's home unit with their row highlighted; Who's next and Calendar open on the same unit.
   - **Calendar:** a month grid, days above nights, like the roster's Calendar tab. Try `< Month`, `Month >` and Print.
   - **Open Postings:** the person's postings first (listed, then out of department), then one table of every open posting with filter arrows; pick a posting in the box at the top to see its list.
   - **Not me (clear)** on Home puts every sheet back to the neutral view.
   - **My history:** pick a name. It lists that person's entries this year. (Before v24.1 this sheet came out empty: the publish was dropping every log row.)
8. **Sign-up (only if you set `SIGNUPS_ON = True` on the copy and published):** in the members' file click **Sign me up** for a posting, choose 8 hours. A file appears in the `Signups` folder. In the roster copy, Award Posting > LOAD that posting: the name is ticked with 8 h and the line under the list says "1 electronic sign-up ticked (1 partial)". Submit the award: the file moves to `Signups\Done\<posting>\`.
9. **IntegrityCheck:** no new PROBLEM lines.
