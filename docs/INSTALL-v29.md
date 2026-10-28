# Call Out Roster v29 - install (the Install panel) + the members' install (Members M6)

The go-live kit, built on v28 (Graeme, Oct 27, 2026: `briefs/v29.md`, `briefs/admin-lite.md`): **Ray's Admin Lite Desk**, **Fix an entry**, **Needs Attention in Ray's order**, **the members' requests follow a void** (+ IntegrityCheck 17), and **the guided move** that splits the folders at the January cutover. Install in the **Nov 18-20 slot**, after the alpha closes (ClosePeriod for Nov 1-14 runs on v28 FIRST). The freeze is Dec 20 - Jan 4; the week before Jan 23 stays free.

Needs **v28 installed** (v27.5 rode in it). The Roster Updater is unchanged (v4.2).

**Back up first. Do it on a practice copy before the real roster. Take the snapshot on the copy BEFORE the install and compare after (`tests/snapshot/modSnapshot.bas`, test-only: TakeSnapshot, install, TakeSnapshot, CompareSnapshots - expect ONE "differs" item on the admin Desk, the relabelled See Desk as link, and two new items, Desk[supervisor] and Desk[lite]).**

## Before the panel, by hand (two things)
1. **Two NEW modules**, `modules/modLite.bas` and `modules/modMove.bas`, come in with the swap (45 of 45). **One NEW UserForm**: in the roster's VBA editor, Insert > UserForm, set its (Name) to `frmFixEntry`, then paste `paste-in/frmFixEntry.txt` into its code (Ctrl+A first). The panel's paste step checks it is there.
2. **The members' template** must be REBUILT (Members M6): open `Call Out Roster - Members TEMPLATE.xlsm`, Alt+F11, replace modMember (Remove, then Import `member/modMember.bas`), Alt+F8 > `BuildMemberTemplate` > OK, save. The master's publish refuses a template older than "Members M6".

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v29` folder.
3. Click the one blue button, read its box, click the next. The swap replaces every module (every stamp moves to v29), imports the two new ones, and re-pastes the forms that changed (**frmFixEntry** - new -, **frmTimeOff**, **frmShiftSwitch**). ThisWorkbook is unchanged. Group 3: System check (45 of 45 on v29; the new path lines and the login line), Self test (1128 cases, 121 suites). Group 4: the practice copy + SmokeTest (22 steps), the Lite Desk through See Desk as, Fix an entry on the copy, the move REHEARSED on the copy into test folders, a void that marks its request, the What's New box, IntegrityCheck (17 checks).

By hand instead: in the VBA editor replace every module in `modules/`, import `modLite.bas` and `modMove.bas`, insert the UserForm frmFixEntry and paste its code, re-paste frmTimeOff and frmShiftSwitch from `paste-in/`; Debug > Compile; save, close, reopen; then the group 4 rows yourself.

**`MoveRoster` is NOT run at the install.** It is the January cutover's step (Alt+F8, admin, from the live roster once IT has the two P: folders ready): `docs/go-live-runbook.md`. The folders stay as they are until then.

## What's new
- **The Admin Lite Desk (Ray).** An admin on `LITE_DESK_USERS` (lanthierr) opens on a simpler Desk drawn on the full admin view: four bands in fixed places - NEEDS YOUR DECISION (members' requests, posting exceptions, corrections asked for, returning from leave), BOOK AND CHANGE, DAILY CHECK (three steps, green when done today, one line when all three are), STAFFING AND OT (below minimum ahead, this period's OT by unit as bars) - over a trust strip (the hours, today's backup, the members' file, the system checks; a red item says call Graeme). **Show full admin Desk >** draws the full Desk for the session; **See Desk as: Supervisor / Admin Lite** lets any admin look at either. Every Lite button runs the same macro as its full-Desk twin; rights are unchanged. A "?" on each band explains it (`docs/admin-lite-cheat-sheet.md`, `docs/guide-08-admin-lite.md`). SystemCheck's new **Logins** line names the admins, the Lite login and the supervisors, and flags a Lite login that is not an admin.
- **Fix an entry.** The guided correction: pick a receipt answered No / Not sure (or type a Ref), see every row it wrote and whose Total Hours it moved, **Void** with a required reason through the roster's own void (VoidRef for OT Log refs with its backup, verify and calendar release; VoidSwitch for a switch; VoidAbsence for time off), **Re-enter it now?** through the pop-up prefilled; the receipt's answer is marked resolved with the new Ref (the Feedback table gains Resolved By / Resolved Ref). An award is withdrawn through Withdraw an award, never voided raw.
- **Needs Attention for the Lite login:** his sections first (requests, open lines, leave, staffing, TO SETTLE), and DATA CHECK + errors + receipt answers folded under **MORE (system)** - double-click the heading to show them. Admins and supervisors see Needs Attention exactly as before.
- **Requests follow a void.** Voiding the absence or switch a Booked request became sets the request to **Booking voided** (the member sees the word after the next publish). **IntegrityCheck 17 "Requests"**: a Booked request whose Ref is blank, of the wrong kind, missing, VOID or naming other people - a Check while the request's dates are ahead, a Note once past, never a PROBLEM. Two v28 gaps fixed: a switch request is matched on the shift too; a request booked as a past sick day becomes Booked (SC- id).
- **The guided move (`MoveRoster`, admin, Alt+F8; a cutover step).** The move copies the roster and the template to the new supervisor-only folder, moves the Signups\ and Requests\ files to the new members' folder, writes the members' folder INTO the new copy (the one stored value - `PUBLISH_FOLDER` and every other setting stay constants; no January install), freezes the old file (every open says it moved; nothing can be logged, saved or published in it) and says close this / open the new one. The new file's first admin open asks for SystemCheck, IntegrityCheck and a publish. On a practice COPY the same macro runs a **REHEARSAL** into test folders you name (the rehearsal week). **SystemCheck lists every path** - master, members', sign-ups, requests, backups, archive, closed, scans, reports, Kronos, the guide, the error log - reachable and writable, the members' file's age, and where the roster moved from.
- **Members M6:** the status word Booking voided; **My history's PDF saves to the user's Downloads folder** (the members' folder will be read-only for members).
- **The What's New box** shows four lines once per login.

## Known, on purpose
- The OT by unit and Posting outcomes REPORT buttons on the Lite Desk stay hidden until those reports exist (the bars and the counts are there).
- `SUPERVISOR_USERS` stays mihaychukc, sheltond, campeauc (Graeme, Oct 27); more are added as constants at go-live if asked.
- A practice copy made from a MOVED roster still publishes beside itself (a copy honours only a rehearsal's stored folder), so nothing from a copy reaches the real members' folder.
- SystemCheck's path lines warn; they never fail the Install panel's check step (that stays versions only).
- The members' folder rights (IT, by Nov 20): the folder root READ for members; `Signups\` and `Requests\` LIST + READ + WRITE for members (the members' file reads them back); supervisors MODIFY on the members' folder.
- The alpha's S2 / S3 fixes (`docs/alpha-findings.md`) and the email digest (item 39) are not in this zip unless the CHANGELOG's v29 entry lists them.

## Then
Send the SmokeTest box, the System check text (the Paths and Logins lines), the Self test line, the rehearsal move's final box from the practice copy, the Lite Desk screenshot (See Desk as: Admin Lite), and the IntegrityCheck result.
