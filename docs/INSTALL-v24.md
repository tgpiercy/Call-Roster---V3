# Call Out Roster v24 - employee changes, crew history, partial OT - install

Needs v23 installed and **the 2027 rollover done**. Install window: **Jan 5 - 15, 2027**. Full module, paste-in and member set in the zip; the steps below list only what to replace. About 30 minutes.

**Back up first** (the daily backup next to the roster counts if it ran today).

## 1. Replace every module
In the VBA editor (Alt+F11), for EACH module, right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:

`modOvertimeEntry`, `modOTLog`, `modPostings`, `modShiftCycle`, `modShiftSwitch`, `modCalendar`, `modHolidays`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modPublish`, `Call_Out_Roster`, `modLeave`, `modAttendance`, `modIntegrity`, `modRollover`

(Code changed in `modOvertimeEntry`, `modOTLog`, `modPostings`, `modShiftSwitch`, `modCalendar`, `modSystem`, `modUI`, `ImportKronosData`, `modGradeLog`, `modSelfTest`, `modIntegrity`, `modRollover`. The others are the same code with a v24 stamp. `modOTLog` carries no stamp but DID change: replace it. `modLogLayout`, `modMigration`, `modCleanup`, `Sort_Hours_Dept`, `OT_Posting` and `LoginFullAccess` are unchanged.)

## 2. Import the new module
**File > Import File:** `modules/modPeople.bas`.

## 3. Re-paste ThisWorkbook and one form
- Double-click **ThisWorkbook**, Ctrl+A, paste `paste-in/ThisWorkbook.txt`.
- Double-click **frmPostingEntry** > View Code, Ctrl+A, paste `paste-in/frmPostingEntry.txt`. Do not change the form's design.

## 4. Add the new form
**Insert > UserForm**. In the Properties window set **(Name)** to `frmEmployee`. View Code, Ctrl+A, paste `paste-in/frmEmployee.txt`. Leave the form itself blank: the code builds it.

## 5. Debug > Compile VBAProject
No message expected. If one appears, stop and send me the line it highlights.

## 6. Run once: SetupPeople (Alt+F8)
It saves a backup copy, then:
- adds the **Crew History** sheet with one row per person: their crew from the start of the roster year;
- adds the empty **Former** sheet;
- adds a **Signup Hours** column on the end of the posted lists (Postings sheet).

It finishes with "Roster totals and the OT Log are unchanged". If it says it did NOT verify, close without saving and send me the message.

## 7. Run once: ApplyRosterLayout (Alt+F8)
The ADMIN button **Rename Employee** becomes **Employee**.

## 8. SystemCheck (Alt+F8)
Expect **"Modules: 18 of 18 on v24"**, no ThisWorkbook or form line, and a new line **"Crew History: n rows"**.

## 9. SelfTest (Alt+F8)
Expect **"299 passed, 0 failed"** and "all 35 suites passed".

## 10. IntegrityCheck (Alt+F8)
Same result as before the install.

## Members' template
`member/modMember.bas` is unchanged. Nothing to do.

## Smoke test (on a copy with COPY in the name, in its own folder)
Use made-up people. Delete the copy afterwards.
1. **Add:** ADMIN > Employee > Add. Name `Test, Pat`, a crew, a seniority date with the year, a home unit, a grade code. Click **Work out the hours**: it lists the unit's peers and their average. Click **ADD EMPLOYEE**.
   - The new row is at the bottom of the roster with the same look as the others; Refused shows the average; Total = Refused.
   - The OT Log's top row is an Adjustment, Ref `NH-0001`, Reason `New hire`.
2. **Crew typed on the roster:** change Pat's crew in column E. A box asks from which date. Enter a date two weeks ahead. Book Pat a vacation day on a day the OLD crew works next week: it lands on the old crew's shift on the Calendar. Press Cancel on a second crew change: the old crew comes back.
3. **Move:** Employee > Move to unit, pick Pat, another unit, a grade code. **Work out the hours** shows the highest total in that unit. Save. Pat's old code is gone, the new one is there, Home Unit changed, the Grade Log has two new rows, and (if Pat was below the highest) the OT Log has `UM-0001`.
4. **Rename:** Employee > Rename, Pat -> `Test, Patricia`. The confirm lists the rows per table. After it, Pat's vacation day on the Absences sheet and the two Adjustment rows carry the new name, and roster totals are unchanged.
5. **Partial OT:** New Posting for Pat's unit. Award Posting > LOAD. Select someone low on the list, click **PARTIAL HOURS**, enter 8. Mark them AWARD. If anyone is marked signed up for the whole shift, SUBMIT AWARD refuses ("comes first"). With no full sign-up it submits: the OT Log row for the awarded person shows Hours 8, people ahead who didn't sign up are charged the posting's full hours x 1.75, and the Calendar line shows the name with `8h`.
6. **Remove / Bring back:** Employee > Remove, pick Patricia, today's date. The row leaves the roster and appears on the **Former** sheet; everyone else's totals are unchanged. Employee > Bring back returns the row with the same Refused hours.
7. **IntegrityCheck:** no new PROBLEM lines.
