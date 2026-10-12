# Call Out Roster v22.1 - sick and personal day tracking - install

Needs v22 installed and checked. Full module, paste-in and member set in the zip; the steps below list only what to replace. About 15 minutes.

**Back up first** (the daily backup in the roster's folder counts if it ran today).

## 1. Replace every stamped module
In the VBA editor (Alt+F11), for EACH module, right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`

(Code changed in `modCalendar`, `modSystem`, `modUI`, `modOvertimeEntry`, `modSelfTest`; the others are the same code with a v22.1 stamp. `modOTLog`, `modLogLayout`, `modMigration`, `modCleanup`, `Sort_Hours_Dept`, `OT_Posting` and `LoginFullAccess` have no stamp and are unchanged.)

## 2. Import the new module
**File > Import File:** `modules/modAttendance.bas`.

## 3. Re-paste two forms
Open each form's code (F7), Ctrl+A, paste:
- **frmTimeOff** from `paste-in/frmTimeOff.txt`
- **frmSickCall** from `paste-in/frmSickCall.txt`

ThisWorkbook and the other forms are unchanged.

## 4. Debug > Compile VBAProject
No message expected.

## 5. One-time macro (Alt+F8)
**ApplyRosterLayout** - rebuilds the command bar; the ADMIN group gains an **Attendance** button (next to Full Access).

## 6. SystemCheck (Alt+F8)
Expect **"Modules: 15 of 15 on v22.1"** and no form or ThisWorkbook line.

## 7. SelfTest (Alt+F8)
Expect **"200 passed, 0 failed"** (18 new `attend` cases).

## 8. Members' template (optional)
`member/modMember.bas` only gained a colour for personal days. Re-import it in the template when convenient (Remove + Import `modMember`, Compile, save), then click Member View once.

## 9. Counting starts here (decided Oct 12, 2026)
No back-fill: the counters start at this install. Sick and personal days taken earlier in 2026 are not in the system, so for the rest of 2026 the Attendance sheet shows what was logged from now on. From Jan 1, 2027 the counts are complete.

Book Time Off > **SIC - Sick day (past)** stays for the odd correction (a sick day that wasn't logged at the time): one row per scheduled shift in the range, no OPEN line, no call-out.

## Smoke test (on a copy)
1. **Sick Call:** pick someone. Under the Note box a line shows "2026: N sick days (N paid of 5) in N of 3 occurrences; personal N of 5." and "This shift: Paid sick day ..." (or why not). Save only; the Absences row's Note ends with that outcome.
2. **Book Time Off > PER - Personal day:** a range with days off at either end offers to trim; a range with no scheduled shift is refused. The calendar shows "Personal".
3. **Book Time Off > VAC:** same trim offer; saving as typed still works (answer No).
4. **Attendance** button: the sheet lists everyone, worst first, with the counters and status. It is hidden again when the file is saved and reopened by a non-admin.
5. **Needs Attention:** anyone at or over a limit appears under ATTENDANCE; double-click opens the Attendance sheet.
6. **Void** a sick row from Book Time Off: the counters drop.
