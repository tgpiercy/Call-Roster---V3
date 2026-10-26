# Call Out Roster v28 - install (the Install panel) + the members' install (Members M5)

The go-live release (January 2027, after the rollover): **members' requests** from their file and **the Payroll # as the key**. Decided by Graeme Oct 25, 2026 (`docs/v28-report-back.md`, `briefs/v28.md`). The install freeze is Dec 20 - Jan 4 and the week before Jan 23 stays free: install in the window between.

Needs **v27.5 installed** (or v27.1: v27.5's three rule fixes ride in this zip - the install never removes a module). The Roster Updater is unchanged (v4.2).

**Back up first. Do it on a practice copy before the real roster.**

## Before the panel, by hand (two things)
1. **Two NEW modules**: the panel imports `modules/modRequests.bas` and `modules/modKeys.bas` (43 of 43). Nothing to do by hand for them.
2. **The members' template** must be REBUILT (Members M5): open `Call Out Roster - Members TEMPLATE.xlsm`, Alt+F11, replace modMember (Remove, then Import `member/modMember.bas`), re-paste `member/ThisWorkbook.txt`, Alt+F8 > `BuildMemberTemplate` > OK, save. The master's publish refuses a template older than "Members M5".

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v28` folder.
3. Click the one blue button, read its box, click the next. The swap replaces every module (every stamp moves to v28), imports the two new ones, re-pastes **ThisWorkbook** (open step 10: the requests, before the Desk lands) and the three forms that changed (**frmTimeOff**, **frmShiftSwitch** - the prefill from a request - and **frmEmployee** - the Payroll # box). Group 3: System check (43 of 43 on v28), Self test (1053 cases, 113 suites), **AssignPayrollKeys** (run once: the Payroll # column on every people table, filled from the roster, verified). Group 4: the practice copy + SmokeTest (22 steps), the Requests folder, a request end to end, the Employee pop-up's Payroll # box, the What's New box, IntegrityCheck (16 checks).

By hand instead: in the VBA editor replace every module in `modules/`, import `modRequests.bas` and `modKeys.bas`, re-paste ThisWorkbook, frmTimeOff, frmShiftSwitch and frmEmployee from `paste-in/`; Debug > Compile; save, close, reopen; Alt+F8 > `AssignPayrollKeys` > OK; then the group 4 rows yourself.

## What's new
- **Members' requests (admins act).** The members' file (M5) gets a **Requests** page (Home > MY REQUESTS): Kind (Time off / Switch), From, To, Shift, Partner, Note, **SEND REQUEST**. One small file per click goes into `Requests\` beside the members' file (as the sign-ups into `Signups\`). A **switch** request waits until the named partner clicks **Agree** on their own PC (**ASKED OF ME** on their Requests page; an Agree from the same login does not count). The master reads the folder on an admin's open and from the admin Desk's new tile; Needs Attention lists the waiting ones under **MEMBERS' REQUESTS**: a **double-click books it** through the time-off or switch pop-up, PREFILLED (every rule of those pop-ups applies: bookings, the 16 h stop, the sick window) - the save marks the request Booked with its AB- / SS- id; a **right-click declines** with a reason (no leave words - it reaches every member). The members' file shows **My requests** with the status (Sent / Waiting for partner / Waiting / Booked / Declined / Declined by partner / Withdrawn / Expired), the admin's answer and the decided date after the next publish, and offers **Withdraw** while a request waits. A request books nothing by itself; the member's note is read by the admin only. Settings: `REQUESTS_ON` True, `REQUEST_FOLDER` ".\Requests". The admin-only **Requests** sheet (`tblRequests`) keeps every request.
- **The Payroll # as the key.** Every people table (the OT Log, the posting lists, Absences, Shift Switches - two columns -, the Grade Log, Hours History, Crew History, Planned Moves, Restrictions, Evidence, Requests) carries a **Payroll #** column on its END, filled from the roster's column A (then the Former sheet) by **AssignPayrollKeys** once at this install and kept current before every admin save, after a Kronos import and after an Add / Edit (a corrected number moves onto the rows that carried the old one). **A rename follows the number** where a row carries it (a row whose name was mistyped by hand is found too). The Employee pop-up takes the number at **Add** and **Edit** (digits only, at most 10, never on another row; the Kronos import confirms or fills it). **IntegrityCheck 16 "Keys"**: a number on two roster rows (PROBLEM), a key the roster and Former do not hold (Check), rows without a key (Note). The roster's pay-period formulas stay by name; the members' file never sees the number. The call order, the lists and the calendar still match by name in v28.
- **The What's New box** shows three lines once per login.

## Known, on purpose
- Before `AssignPayrollKeys` runs, IntegrityCheck 16 says which tables have no Payroll # column yet (one line) - run it once (group 3) and it goes. The OT Log's column is added by the roster itself on the first open after the install (its layout check needs it).
- The key columns are hidden for supervisors and plain users (as the roster's column A is) and left out of the daily CSV export; the period-close CSV, an admin's evidence file, carries the OT Log's key column.
- A member whose PC cannot reach the `Requests\` folder is told so and asks the usual way; the folder needs the same rights as `Signups\`.
- A switch request between two sister-crew people for the same day is booked through the switch pop-up with its own rules (v27.5: the same-day swap of both shifts is exempt from the 16 h arithmetic; a one-way switch is not).
- The go-live settings (`ADMIN_USERS`, `SUPERVISOR_USERS`, the P: paths, the members' folder) are constants set at install, as every release so far. The email digest (item 39) is not in v28.

## Then
Send the SmokeTest box, the System check line, the AssignPayrollKeys receipt, the request you sent end to end (the members' file's My requests line after the publish) and the IntegrityCheck result.
