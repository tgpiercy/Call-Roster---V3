# The alpha's data reset - the real roster set aside (Oct 31) and put back (Nov 16-17)

The alpha of Nov 1-14, 2026 runs on the MOCK plant (tests/mock) on the real roster FILE under its real name and folder,
with the real logins. The plant's real call-outs run on paper for the fortnight (paper governs, as planned). This page
is the runbook for the two swaps: who does what, and the box each step should show. Decided by Graeme Oct 28, 2026
(D1: set aside, catch up from paper; D5: the members' file and its folders stay where they are).

## Page 1 - Day 0, Sat Oct 31, before anything else (Graeme)

| # | step | the box / what to see |
|---|---|---|
| 1 | Close the real roster (nobody else has it open: the lock file `~$V2 - Call Out Roster.xlsm` is gone) | - |
| 2 | In the roster's folder make the subfolder `Pre-alpha\` | - |
| 3 | COPY (not move) `V2 - Call Out Roster.xlsm` into `Pre-alpha\` and rename the copy **`Backup - PRE-ALPHA 2026-10-31 - V2 - Call Out Roster.xlsm`** | the name holds "Backup": opened by mistake it reads as a backup copy (`IsBackupCopyName`) - no daily backup is taken from it, nothing is published from it, it is not the live roster (`IsLiveRoster` False); it holds no "COPY" / "DRY RUN", so it is not a practice copy either; `PrunePreChangeCopies` matches only `Backup before * - <name>` beside the roster and `PruneBackups` only the exact daily / hourly names in the roster's own folder (`BACKUP_FOLDER` is "."): a subfolder is never pruned |
| 4 | COPY the members' file `Call Out Roster - Members.xlsm`, the `Signups\` folder and the `Requests\` folder into `Pre-alpha\` as they are (`Signups\Done\`, `Requests\done\` included) | the real sign-ups and requests of that moment are kept whole |
| 5 | Empty `Signups\` and `Requests\` (the files only; keep the folders and their `Done` / `done` subfolders empty) | the mock run starts with no inbox |
| 6 | Make the Kronos mock folder v29.0.1's `KRONOS_PATH` names (`<the roster's folder>\Kronos mock\`) and copy `00 cumulative 2026-01-11 to 2026-10-30.xlsx` into it (from tests/mock/kronos, made by make_kronos.py) | - |
| 7 | Install v29.0.1 on a practice copy first, then on the roster (docs/INSTALL-v29.0.1.md: the one change is `KRONOS_PATH`) | SystemCheck "45 of 45 on v29.0.1"; the paths box: Kronos reports = the mock folder - reachable |
| 8 | Open the roster (the live file), Alt+F8 > `LoadMockData`, type `MOCK` | the box names the "Backup before MOCK LOAD ..." copy, then: 69 people, 94 OT Log rows, every figure matches expected_load, IntegrityCheck 0 PROBLEM, 2 sign-up files for P-0003 |
| 9 | Carry on with docs/alpha-test-plan.md Day 0 (D0.3 onward) | - |

During the fortnight the mock file takes the real daily and hourly backup slots (`Call Out Roster yyyy-mm-dd.xlsm`,
`Call Out Roster backup yyyy-mm-dd hh.xlsm`) in the roster's folder: they hold mock data under the real names, and the
first of pay period 23 (Nov 1) would be the long-kept one. Page 2 moves them out.

## Page 2 - the close-out, Mon Nov 16 and Tue Nov 17

### On the MOCK file first (Graeme; Ray watches the Lite Desk steps)
| # | step | the box / what to see |
|---|---|---|
| 1 | Copy `15 cumulative 2026-01-11 to 2026-11-14.xlsx` into the Kronos mock folder; Desk > Import Kronos | read from the report; Data Through Sat Nov 14; H per person = the ledger's figure plus Nov 14's hours; no sanity confirm |
| 2 | Alt+F8 > IntegrityCheck | 0 PROBLEM (the Check / Note lines as the plan's D13.8) |
| 3 | Alt+F8 > ClosePeriod, the period starting Nov 1 | the CSV, the PDF and the SHA-256 row in Closed Periods (expected.csv row C.3: the row count and the charged sum); label the three files MOCK when they are filed |
| 4 | Alt+F8 > StatsReport; Needs Attention > the F: line > ShowFeedback | every Feedback row seen |
| 5 | Every plan tick made; every docs/alpha-findings.md row triaged (P1) | - |
| 6 | Close the roster. Nobody opens it again until step 10 | - |

### The swap (Graeme)
| # | step | the box / what to see |
|---|---|---|
| 7 | Rename the mock file `V2 - Call Out Roster.xlsm` to **`Backup - ALPHA MOCK 2026-11-14 - V2 - Call Out Roster.xlsm`** and MOVE it into `Pre-alpha\` | the mock evidence stays readable (opened, it is a backup copy: nothing publishes, nothing prunes) |
| 8 | MOVE the mock run's daily and hourly backups of Oct 31 - Nov 17 (`Call Out Roster 2026-10-31.xlsm` ... `Call Out Roster backup 2026-11-17 ....xlsm`) and the `Backup before MOCK LOAD ...` copy into `Pre-alpha\` | the real file's next open takes a clean daily slot; the long-kept first backup of period 23 will be the real file's |
| 9 | Delete the mock `Call Out Roster - Members.xlsm` and empty `Signups\` and `Requests\` (the testers' mock files and their `Done` / `done` contents); copy the Day-0 `Pre-alpha\Call Out Roster - Members.xlsm`, `Signups\` and `Requests\` back | the real inbox is as it was on Oct 31; the members' file is the Day-0 one until step 12 publishes |
| 10 | COPY `Pre-alpha\Backup - PRE-ALPHA 2026-10-31 - V2 - Call Out Roster.xlsm` back beside the folders and rename the copy `V2 - Call Out Roster.xlsm` (the Pre-alpha copy stays) | the real roster is back, on v29 (NOT v29.0.1: it was set aside before the install) |
| 11 | Open it; SystemCheck; IntegrityCheck | 45 of 45 on v29; the Kronos reports path = the P: processing workbook as before; 0 PROBLEM; the Opens row says v29 |

### Catch up from paper (Graeme, with the supervisors' paper sheets; one event at a time, in date order)
| # | step | the box / what to see |
|---|---|---|
| 12 | Every real Nov 1-14 event as a real row dated on its day: call-outs through Log Call-Out with the paper's outcomes (the date box set to the shift's date; the pop-up fills the crew from the rotation), postings through New Posting / the Award pop-up with the paper sign-ups ticked, sick calls through Sick call (the date box), switches through the Shift Switch pop-up, bookings through Book Time Off. Each receipt's "Was this right?" = Yes | the receipts; IntegrityCheck 12 names any accepted row whose calendar line is missing - fix with Set coverage |
| 13 | The REAL cumulative Kronos report (Jan 11 - Nov 14, through the processing workbook as before: paste the Totals block, run its macro, save) ; Desk > Import Kronos | "Read from the pivot sheet 'Sorted Data Here'"; the sanity confirm lists anyone whose H rose by more than 60 h in the fortnight (expected for nobody) |
| 14 | Alt+F8 > IntegrityCheck; Alt+F8 > RolloverYear and CANCEL at its first box (a glance at RolloverProblems) | 0 PROBLEM; the rollover's first box lists nothing unexpected; nothing rolled |
| 15 | Desk > Update the members' file | published; the members' file is the real one again (Members M6) |
| 16 | `Signups\` holds no file named with a tester's login from the mock run | the folder was restored in step 9 |
| 17 | Alt+F8 > ClosePeriod for Nov 1 on the REAL file (the real period's evidence), once paper and roster agree | the real period's CSV / PDF / hash row |

What is NOT done at the close-out: no install (v28.1 waits for its own slot), no rollover, no guided move. The mock
Kronos folder may stay (v29 on the real file reads the P: workbook). The `Pre-alpha\` folder is kept until the real
period close of Nov 15-28 has run clean, then archived with the alpha's evidence.
