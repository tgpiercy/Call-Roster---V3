# Update 20: holidays, Kronos age, safety checks

## What's new
**Stat holidays.**
- DAYS workers are treated as off on plant holidays. They appear on call-out and posting lists that day, and shift switches treat them as free.
- A/B, C/D and the A-D crews are unchanged.
- All 11 holidays are worked out for any year (New Year's, Family Day, Good Friday, Victoria Day, Canada Day, Civic Holiday, Labour Day, Thanksgiving, Remembrance Day, Christmas, Boxing Day). There's nothing to update each January.
- The holiday name shows top-right on its date on the Calendar.
- **Weekend holidays:** a new admin **Holidays** sheet lists this year and next. When a holiday falls on a weekend, enter the weekday the plant gives off in **Observed on**. Until you do, Needs Attention lists it under **HOLIDAYS TO CONFIRM**; double-click the row to jump to the sheet.
  - **Boxing Day 2026 is a Saturday.** It will be listed until its weekday is entered.
- The floater is booked as vacation, as before.

**Kronos.**
- The status line has a third line: **"Kronos: Oct 7 (2 days ago)"**. It turns red once the data is more than 7 days old (`KRONOS_STALE_DAYS` in modOvertimeEntry).
- Log Call-Out (on Load) and New Posting (on Create) warn when the data is old. **No** stops so you can import first.
- Every import saves a dated copy of everyone's worked hours on a new hidden **Hours History** sheet. Later updates use it for the leave top-up and provisional accepted OT, so it needs to start recording now.
- The import message now also lists Kronos names that aren't on the roster.

**Safety.**
- **Alt+F8 > SystemCheck:** shows any problems first: modules not on v20, leftover or duplicate modules, forms or ThisWorkbook not re-pasted, OT Log layout, old Kronos data, holidays to confirm.
  - Checking forms and leftovers needs **File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model**. Without it, the module versions are still checked.
- **Read-only copies** now say who has the file open for editing. Logging was already blocked in a read-only copy; that's unchanged.
- If the file is opened from SharePoint/OneDrive, AutoSave is turned off. Every submit already saves, and this stops two people's copies being merged in the background.
- **Award emails:** if Outlook can't be driven (for example after a move to the new Outlook), the email opens in your default mail app instead. If no mail app opens, the text is shown so you can copy it.

**Shift Switch form.** Includes the Oct 9 fixes (January dates typed in December, unit change clears names, payback warnings before Save, void clears the form, Enter doesn't save), plus:
- DAYS workers get their shift filled in and the same warnings as A/B and C/D.
- The form no longer depends on the wording of a message.

## Checks done on the current code
- **Void by `#245`** matches the Entry ID, not the sheet row. It's safe with the newest-first log.
- **`POSTING_INCLUDE_DAYS`** was never used anywhere. Postings already list DAYS people whenever they're free. The setting is removed.

## Install
**Back up first.** In the VBA editor (Alt+F11):

1. **Replace these modules** (Remove, answer No to export, then File > Import from `modules/`):
   `modOvertimeEntry`, `modCalendar`, `modShiftCycle`, `modShiftSwitch`, `modPostings`, `ImportKronosData`, `modUI`
2. **Import two new modules:** `modHolidays.bas`, `modSystem.bas`.
3. **Re-paste the code** (open the code, Ctrl+A, paste) for:
   - ThisWorkbook, from `paste-in/ThisWorkbook.txt`
   - frmShiftSwitch, from `paste-in/frmShiftSwitch.txt`
   - frmOvertimeEntry, from `paste-in/frmOvertimeEntry.txt`
   - frmNewPosting, from `paste-in/frmNewPosting.txt`
4. **Debug > Compile VBAProject.** It should finish with no message.
5. **Save, close and reopen.** Opening creates the Holidays sheet (admins only).
6. **Alt+F8 > SystemCheck.** Expect "9 of 9 on v20". The only items it should list are Boxing Day and the Kronos import (until step 7).
7. **Import Kronos** once, so the status line and Hours History start.
8. On the Holidays sheet, enter Boxing Day's weekday once the plant announces it.

The other modules and forms are unchanged. The zip has the full set so it can serve as the master copy.

**Don't import** any `modCalendar.bas` other than the one in `modules/`. The old update folder had an outdated copy at the top level; it isn't in this zip.

## Smoke test
- **Thanksgiving:** Log Call-Out for Mon Oct 12, Days. Free DAYS people are listed. The Calendar shows "Thanksgiving" on the 12th.
- **Status line:** the three lines show, and the Kronos line goes red if the data is old.
- **Read-only:** open the file while someone else has it. The message names them.
- **Shift Switch:** type "Jan 5". The hint reads Jan 5, 2027.
