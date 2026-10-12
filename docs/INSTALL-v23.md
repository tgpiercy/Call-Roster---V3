# Call Out Roster v23 - integrity check and year-end rollover - install

Needs v22.1 installed and checked. Full module, paste-in and member set in the zip; the steps below list only what to replace. About 20 minutes.

**Back up first** (the daily backup next to the roster counts if it ran today).

**Do not run RolloverYear on the live file as part of this install.** On the live file it refuses before Dec 27, 2026 anyway. Practise on a copy (step 9).

## 1. Replace every module
In the VBA editor (Alt+F11), for EACH module, right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modOTLog`, `modLogLayout`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`

(Code changed in `modOvertimeEntry`, `modOTLog`, `modLogLayout`, `modPostings`, `modShiftSwitch`, `modCalendar`, `modSystem`, `modUI`, `ImportKronosData`, `modSelfTest`, `modPublish`, `modLeave`. The others are the same code with a v23 stamp. `modOTLog` and `modLogLayout` carry no stamp but DID change this time: replace them. `modMigration`, `modCleanup`, `Sort_Hours_Dept`, `OT_Posting` and `LoginFullAccess` are unchanged.)

## 2. Import the two new modules
**File > Import File:** `modules/modIntegrity.bas` and `modules/modRollover.bas`.

## 3. Re-paste ThisWorkbook
Double-click **ThisWorkbook**, Ctrl+A, paste `paste-in/ThisWorkbook.txt`. (It now runs the quiet integrity check when an admin opens the file.) No form changed.

## 4. Debug > Compile VBAProject
No message expected.

## 5. Settings to look at (top of `modOvertimeEntry`)
| Setting | Shipped value | Change it? |
|---|---|---|
| `ARCHIVE_FOLDER` | `"."` | The frozen year file goes next to the roster. Leave it for now. |
| `REPORT_FOLDER` | ends in `\{YEAR}\` | The year now follows the roster. Leave `{YEAR}` in place. |
| `PAYBACK_WARN_DAYS` | `42` | Leave. |

## 6. SystemCheck (Alt+F8)
Expect **"Modules: 17 of 17 on v23"**, no ThisWorkbook or form line, and a new line **"Roster year: 2026 (26 pay periods from Dec 28, 2025)"**.

## 7. SelfTest (Alt+F8)
Expect **"248 passed, 0 failed"** and "all 26 suites passed".

## 8. IntegrityCheck (Alt+F8)
Run it once. It only reads. The **Integrity** sheet lists what it found:
- **PROBLEM** lines mean roster totals can't be trusted until fixed. Each line says what to do. Tell me what it shows before fixing anything by hand.
- **Check** lines are worth a look (someone no longer on the roster, an old payback, old Kronos data). They do not block anything.

From now on the roster's status line shows "Integrity: OK" each time an admin opens the file.

## 9. Practise the rollover on a copy (first week of December, with your boss)
Follow `docs/rollover-one-pager.md`, "Practice run". The copy must have `DRY RUN` or `COPY` in its file name.

## 10. Members' template (optional)
`member/modMember.bas` only has a new version stamp. No need to re-import it.

## Smoke test (on a copy with DRY RUN in the name, in its own folder)
1. **RolloverYear:** log a posting for a January 2027 shift and submit an award that charges someone. Run RolloverYear, confirm with `ROLL 2027`.
   - The final message says "PRACTICE RUN finished" and "verified".
   - The roster's first period is Dec 27; Worked is 0; the January charge shows in the right period.
   - Sheet **OT Log 2026** holds the old rows; the live OT Log holds only the January rows.
   - **Year End** lists everyone's 2026 finals.
   - `Call Out Roster 2026 FINAL (dry run).xlsm` sits next to the copy and is read-only.
2. **IDs carry on:** log one call-out. Its Ref ID continues from the last one of 2026 (not `S-0001`).
3. **Kronos:** click Import Kronos with the current (2026) Kronos file. It is refused with "still holds last year's hours".
4. **Run it twice:** run RolloverYear again. On a practice copy it offers 2027 -> 2028 (the date rule is off for copies): click **Cancel**. The live file refuses with "It's too early".
5. **IntegrityCheck catches a typed number (do this last):** in admin view, type `5` over one pay-period formula and run IntegrityCheck. It lists PROBLEMs for that person, one of them "holds a typed number". The file saves itself on close, so this copy is now spoiled.
6. Delete the copy and its frozen file.
