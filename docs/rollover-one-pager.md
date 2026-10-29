# Year-end rollover: one page

The steps for rolling the Call Out Roster from one year to the next. Current as of v29.0.3 (Oct 28, 2026; v29.0.3 = v29 with the alpha's changes, on the MOCK file only - `KRONOS_PATH` names the local Kronos mock folder, `REPORT_FOLDER` an `Alpha reports` subfolder beside the roster, and Act as lets an admin play a chair with a password - plus the Oct 28 review's fixes: a partial acceptance closes a call-out by the order the outcomes were LOGGED, one default closing date for every posting path, Fix an entry re-enters a session at its hours on offer and a personal day as a personal day, Bring back restores Worked, the Kronos age reads from the data, the hourly backup is saved in the restricted view, Save As and a read-only copy's save are refused, a backup-named copy is frozen; it supersedes v29.0.2, never installed; v29 adds the Admin Lite Desk for the second admin, Fix an entry, the members' requests' Booking voided status with IntegrityCheck 17, SystemCheck's path and login lines, and the guided move that splits the folders at the January cutover; v28 adds the Payroll # column on the end of the OT Log - the archive sheet carries it too; run AssignPayrollKeys once at the install BEFORE the rollover so the column exists): v27 adds IntegrityCheck checks 14 ("Log rows": every call-out and posting row charged what its own outcome, hours and code give, in its shift's pay period) and 15 ("Records") - a PROBLEM in either stops the rollover like any other; the rollover empties the Kronos Totals and Worked Days tables and keeps the Kronos Files record; close the year's last pay period (Close the pay period) before rolling. From v26.5: Refused counts the pay periods up to the current one, so a December practice run on a copy dated into the new year gives 0 for the kept rows.

**Two people.** The admin drives. The second person reads each check aloud and ticks it off. Allow 30 minutes. Do it with nobody else in the file.

**When.** On or after the first day of the new year's first pay period. For 2027 that is **Sunday Dec 27, 2026**. The macro refuses to run on the live file before that date.

## What the rollover does
| Step | What happens |
|---|---|
| 1 | Runs the integrity check. It stops if there is any PROBLEM. |
| 2 | Writes each person's final Total, Worked, Refused and Provisional hours to the **Year End** sheet. |
| 3 | Saves a frozen copy, **Call Out Roster 2026 FINAL.xlsm**, next to the roster. It reopens the copy, checks it against step 2, and marks it read-only. |
| 4 | Moves the old year's OT Log rows to a sheet called **OT Log 2026**. Rows already logged for the new year stay. ID numbers carry on. They do not restart. |
| 5 | Sets the roster's first pay-period date to the new year. Every period column follows. |
| 6 | Sets Worked hours (column H) to 0. They stay 0 until the first Kronos export of the new year is imported. |
| 7 | Checks everything. If anything does not match, saving is blocked. |

Nothing is saved unless step 7 passes.

## Before you start
| # | Check | Done |
|---|---|---|
| 1 | Everyone else has closed the roster. You have it open for editing, not read-only. | |
| 2 | Run **Import Kronos** one last time for the old year, so the final Worked hours are current. | |
| 3 | Finish or cancel any posting for an old-year shift that is still Open. | |
| 4 | Alt+F8 > **IntegrityCheck**. The message must not say PROBLEM. Fix anything it lists, then run it again. | |
| 5 | Today's backup exists next to the roster: `Call Out Roster yyyy-mm-dd.xlsm` with today's date. It was taken when you opened the file this morning, BEFORE step 2's import: if you need to restore it, that import is redone. | |
| 6 | No file called `Call Out Roster 2026 FINAL.xlsm` exists next to the roster yet. | |
| 7 | Needs Attention > PLANNED UNIT MOVES shows nothing "DUE - not applied". Apply it (Alt+F8 > ApplyPlannedMoves) or cancel it first. A move dated in the new year can wait. | |
| 8 | No leave return is waiting with a return date in the new year: it is refused until the roll is done, then logged on the new roster. | |

## Run it
1. Alt+F8 > **RolloverYear** > Run.
2. Read the summary box aloud. Check the new year, the start date and the number of log rows moving.
3. Click **OK**.
4. Type `ROLL 2027` in the box and click OK. Use the year the summary shows.
5. Wait. The screen may flicker while the frozen copy opens and closes. Do not click anything.
6. The last message says **"Rolled over to 2027 and verified."** Read it aloud.
7. If that message also says **"THE FILE DID NOT SAVE"**, press Ctrl+S straight away and check it saves before you close.

## Check the result
| # | Check | Done |
|---|---|---|
| 1 | The roster's first period column shows Dec 27 (for 2027). | |
| 2 | Total Hours are zero, or small for anyone already awarded a new-year shift (provisional) or refused one in the first period; a refusal for a later period shows in its column and counts from that period (v26.5). | |
| 3 | Worked (H) is 0 for everyone. The status line says "Kronos: awaiting 2027 data". | |
| 4 | The **OT Log 2026** sheet exists and holds last year's rows. | |
| 5 | The **Year End** sheet lists every person with their 2026 finals. | |
| 6 | `Call Out Roster 2026 FINAL.xlsm` is next to the roster. Open it: it says "Someone else has the roster open for editing, so this copy is read-only" (that box is what a read-only file gets; it is right), shows the 2026 totals, takes no backup and refuses Member View. Close it. | |
| 7 | Alt+F8 > **SystemCheck**: "Roster year: 2027". Alt+F8 > **IntegrityCheck**: no PROBLEM. | |
| 8 | Click **Member View** to publish the new year to the members. | |
| 9 | Create the new year's reports folder (SystemCheck names it if it is missing). | |

## If it stops
**A message before anything changed** ("did NOT start", "too early", "already exists"): nothing was changed. Read the message, fix the cause, run it again.

**"The rollover STOPPED" or "did NOT verify":**
1. Do not save. The macro has blocked saving.
2. Close the file. Choose **Don't Save** if asked.
3. Reopen it. It is exactly as it was before you started. The frozen copy from the failed try was removed.
4. Send the message to the admin who maintains the code before trying again.

**If the file was saved in a bad state somehow:** close it, rename it, and copy today's backup (`Call Out Roster yyyy-mm-dd.xlsm`) back to the roster's name. Redo the final Kronos import (the backup was taken before it).

**If Excel crashes mid-run:** the file on disk is untouched (nothing is saved until step 7 passes). The read-only frozen file from the failed try may be left behind: delete it, or the next run refuses "already exists". If Excel's Document Recovery offers a recovered roster, do not keep it: open the real file.

**A second run on a practice copy** that is already rolled to 2027 would roll it on to 2028 after a prompt. Answer No.

## After the rollover
| When | What |
|---|---|
| First Kronos export saved after Jan 1 | Run **Import Kronos**. Earlier files are refused, because they still hold last year's hours. |
| Straight after that first import | Check two or three people's Worked hours against Kronos. They should be this year's hours only. If they look like last year's total, stop and tell the code admin. |
| Same check | See whether overtime worked Dec 27-31 is in the new Kronos figures. If it is missing, tell the code admin. It is five days of hours for a few people. |
| Any time | A late entry for the old year can still be logged. The form warns that the period is not on this roster. It will not change the frozen file. |

## Things to know
- **Postings made in December for January shifts** keep the totals they were posted with. Their charges show on the new roster in their period's column, and in Refused / Total Hours from that period's start (v26.5).
- **The old year is closed.** Old-year rows are on the OT Log 2026 sheet and cannot be voided. The frozen file is the record.
- **Members see the new year only.** The frozen file is the record of 2026.
- **Sick and personal day counts** run on the calendar year. The rollover does not touch them.
- **A leave that spans the rollover** is topped up from the new year's start only.

## Practice run (first week of December)
1. Alt+F8 > **MakePracticeCopy** in the roster. It saves a copy with COPY in its name in a `Practice` folder and opens it (a file with DRY RUN in its name works the same way).
2. In the copy, log a test posting for a January shift and award it.
3. Run **RolloverYear** in the copy. It says "PRACTICE RUN" and names its frozen file "(dry run)".
4. Go through "Check the result" above. The January posting's charges must show in their period's column; on a practice run before Dec 27, Refused and Total Hours stay 0 for them (their period hasn't started: v26.5).
5. Delete the copy and its frozen file when done.
