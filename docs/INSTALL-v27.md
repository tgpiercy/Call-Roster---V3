# Call Out Roster v27 - install (the Install panel)

The pre-training bundle: everything in `briefs/v27.md` (items 1-38; item 39, the publish digest by email, waits for your signed wording) and the Kronos import rebuilt from the report's Totals block (`briefs/v27-kronos.md`). One master install plus one members' template rebuild (Members M4), as planned for about Nov 20-25, before the Dec 1 training.

Needs **v26.5.5 installed** and **the section-0 Kronos re-import done** (`briefs/v27-kronos.md` section 0) before the real roster takes this.

**Back up first. Do it on a copy before the real roster** (Desk > Practice on a copy).

## Before the panel (by hand, once)
1. **The Roster Updater changes (v4.2).** Open `Roster Updater.xlsm`, Alt+F11, remove `modUpdater`, File > Import `updater/modUpdater.bas` from this folder, Debug > Compile, save. The panel checks the backup copy, copies the members' template before rebuilding it, and keeps SystemCheck's whole text in the Install Log.
2. **Three new forms in the roster.** Open the roster, Alt+F11, Insert > UserForm three times; set the (Name) of each to **frmReceipt**, **frmHistory** and **frmTableEdit**. Leave them empty (the panel pastes their code). Debug > Compile, save.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v27` folder.
3. Click the one blue button, read its box, click the next. The swap replaces every module and **imports the five new ones** (modReconcile, modStats, modReceipt, modRules, modEvidence: 30 of 30), re-pastes ThisWorkbook and every form (the three new ones too). Group 3: the members' template rebuilt (Members M4), the command bar rebuilt (ApplyRosterLayout), System check, Self test. Group 4: the practice copy + SmokeTest, the Desk, a receipt, the draft, the calendar menu and the outage print, the Reconcile sheet, the Stats report, the Kronos acceptance test, the two pop-ups, the OT restriction and the weekly cap, a period close, the handover and the ledger, the members' file, the restore drill, IntegrityCheck.

By hand instead: remove and import every module in `modules/` (the five new ones are new imports), re-paste ThisWorkbook and every form from `paste-in/`, Debug > Compile, run ApplyRosterLayout; in the members' template remove and import `member/modMember.bas`, re-paste its ThisWorkbook, run BuildMemberTemplate, save; then the group 3-4 rows yourself.

## What's new (the headlines; the brief has every item and its review record)
- **The Desk v2** - three pills, TODAY tiles with what a click does, NOW / PLAN / FINISH side by side; the admin's Kronos pill imports in one click; "This morning" (integrity, today's backup, unread errors); "Since you were last here"; the legend; the first days strip.
- **A receipt after every save** - what was recorded, whose Total Hours changed (before -> after, and the cells marked on the roster), where to look, how to undo, and "On paper this would have been different" when a rule paper never had was used; "Was this right?" answers go to an admin-only Feedback sheet.
- **The Kronos import rebuilt** - the report's Totals block, the one-day ledger, the change log; Worked = hours x each code's weight.
- **The 3 AM pass** - the call-out covers its calendar line by itself, a second session for the same shift asks first, the session draft survives a crash, the supervisor's line menu, the outage print, "Not listed (N)...", the error box says what was saved, larger type per login, the history pop-up.
- **The process items** - OT restrictions (never charged) and the 48 h weekly cap (a warning with a reason), the Rules and Settings sheets, the period close with a SHA-256 hash chain, scans and disputes, Withdraw an award / Edit a posting / Post all lines / the plant-upset Exception, the daily CSV export, the data check, the shift handover, the payback ledger, the Holidays and Min Staff pop-ups (nothing is typed into any sheet any more).
- **The testing items** - IntegrityCheck checks 12-15 (coverage = log, settings, log rows, records), SmokeTest 22 steps, four more checker classes, the drift and docs tests, the Opens table.
- **The Reconcile sheet and the Stats report.**
- **Members M4** - the members' file opens on the member (remembered on that PC), four tiles, Sign me up from Home, since you last looked, the header band on every sheet, How it works, who is not listed and why.

## Open points (your answers change one function each)
The six questions in `docs/NEXT-SESSION-PROMPT.md`: the weekly cap's meaning, the FULL history PDF's folder, a full CSV export's folder, the paper-position column, the members' Stats cut, the email digest's wording.

## Then
Send the SmokeTest box, the Kronos acceptance test's names, the group 4 rows that could not be staged, and the IntegrityCheck result at the end.
