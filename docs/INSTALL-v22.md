# Call Out Roster v22 - sick calls and leave - install

Needs v21 and M1 installed and checked. This zip holds the **full** module, paste-in and member set (it is the master from v22 on); the steps below list only what to replace. About 20 minutes.

**Back up first.**

## 1. Replace every module
In the VBA editor (Alt+F11), for EACH module listed, right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`. Every module carries a v22 version stamp, so all of them are replaced:

`modOvertimeEntry`, `modOTLog`, `modLogLayout`, `modMigration`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`

(Only `modCalendar`, `modSystem`, `modUI`, `modPublish`, `modSelfTest` and the version stamps changed in code; the rest are the same code with a v22 stamp. `Sort_Hours_Dept`, `OT_Posting`, `LoginFullAccess`, `modLogLayout` and `modMigration` are included for completeness.)

## 2. Import the new module
**File > Import File:** `modules/modLeave.bas`.

## 3. New forms (three)
For each: **Insert > UserForm**, set its **(Name)** in the Properties window, leave it blank, open its code (F7), paste the file:
- `frmSickCall` from `paste-in/frmSickCall.txt`
- `frmLeave` from `paste-in/frmLeave.txt`
- `frmConvert` from `paste-in/frmConvert.txt`

## 4. Re-paste
- **frmOvertimeEntry** from `paste-in/frmOvertimeEntry.txt` (adds the pre-fill a sick call uses).

## 5. Debug > Compile VBAProject
No message expected.

## 6. One-time macros (Alt+F8)
1. **UpgradeAbsenceColumns** - adds Home Unit, Leave Group and Shift to the Absences table (on the end; backup copy first; row count and roster totals checked).
2. **ApplyRosterLayout** - rebuilds the command bar with the new PEOPLE group (Sick Call, Leave, Convert imports).

## 7. SystemCheck (Alt+F8)
Expect **"Modules: 14 of 14 on v22"**. CHECK lines you may see: blank Home Units, `BACKUP_FOLDER`, and "imported vacation lines ... not booked yet" if old-calendar imports exist.

## 8. SelfTest (Alt+F8)
Expect **"178 passed, 0 failed"**.

## 9. Members' template
`member/modMember.bas` changed (sick calls show in orange). In the template: Remove + Import `modMember`, re-paste `member/ThisWorkbook.txt`, Compile, save. Then click Member View once.

## Smoke test (on a copy)
1. **Sick Call:** pick someone scheduled today, Save + Start Call-Out. Log Call-Out opens with the unit, date and shift filled in; the person is not on the list. The calendar shows them as "Sick" with an OPEN line. Log Call-Out for the next day's day shift: they are still off the list; for the shift after that, they are back.
2. **Leave, going:** book a test leave from a past date (e.g. 6 weeks ago). The person drops off call lists; the calendar shows "Leave".
3. **Leave, returning:** Leave > Returning, pick that booking, return date today, Work out the top-up. The result names the peers, the average and before -> after. Save. The OT Log has one Adjustment `LV-####` with Reason "Leave return" and the calculation in Note; the roster total rose by the average; the booking's end date is yesterday. Void Entry `LV-####`: the total drops back.
4. **Convert imports:** if old-calendar lines exist, Convert imports lists them; book one person's run; the Absences sheet gets one VAC booking covering the run and those dates drop off call lists. Needs Attention's IMPORTED, NOT BOOKED section shrinks.
5. Needs Attention: double-click a leave in the LEAVE section: "Return from leave now (top-up)..." is on the menu.
