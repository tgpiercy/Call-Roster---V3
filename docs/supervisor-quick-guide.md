# Call Out Roster: Supervisor Quick Guide

How to do each everyday job in the roster. Current as of v23 (Dec 2026).

The buttons are in the bar at the top of the **OT Roster** tab. **Log Call-Out** and **Calendar** show for everyone. The other buttons show only for logins on the admin list. If a button is missing, ask the admin.

## Open the file
1. Open the roster from its usual folder. Click **Enable Content** if Excel asks.
2. A notice that someone "has the roster open for editing, so this copy is read-only" means you can look but not log. Close it. Open it again when they are out.
3. Close the file as soon as you are done, so the next supervisor can get in.

## Log a call-out
1. Click **Log Call-Out**. Pick Department, Shift and Date. Type the Hours. Leave Crew on "All not scheduled".
2. Click **LOAD LIST**. The list is in call order and holds only people free for that shift.
3. Call the name shown at "Now calling". Click **ACCEPTED**, **REFUSED**, **LEFT MSG** or **UNAVAIL.** For Unavailable, pick the reason first.
4. When the shift is filled, click **SUBMIT ALL ENTRIES**. Note the session number (S-0012) in the message.
5. If asked which calendar line the person covers, type its number.

Refused and Left Message are charged (grade 2 and below). `>>` means above grade 2, called but not charged. `*` means the total includes accepted overtime not yet in Kronos. If warned that Kronos data is old, No stops so the admin can import first.

## Postings
**Post a shift**
1. Click **New Posting**. Fill in Unit, Shift date, Shift, Reason, Hours and Positions. Crew and "Posting closes" fill in.
2. Leave "Print it" and "Save PDF" ticked. Click **CREATE POSTING**, then OK.
3. Put up the printed sheet. Sign-up is on paper.

**Load a posting:** click **Load Posting**, pick it, then **View on sheet**, **Print** or **Save PDF**.

**Award it**
1. Click **Award Posting**. Pick the posting and click **LOAD**.
2. Double-click each name that signed up.
3. If someone signed up for only part of the shift: select the name, click **PARTIAL HOURS** and enter the hours.
4. Select the person who gets it and click **AWARD (on / off)**. Repeat for more positions. Anyone who signed up for the whole shift comes before a partial sign-up.
5. Read the charge summary. Click **SUBMIT AWARD**, then OK.
6. An email draft opens. Check it and send it.

| Case | Do this | Charged |
|---|---|---|
| Someone from another department signed a blank line | Pick the name in the "Out of department" box. Click **AWARD OUT OF DEPT**. | In-department people who did not sign up |
| Nobody took it | Click **NO TAKERS**. | Everyone eligible on the list |
| Overtime no longer needed | Click **CANCEL REQUEST**. | Nobody |
| Only part of the shift was taken | Award the partial sign-up(s). The rest of the hours drop off. | People ahead who did not sign up at all, the posting's full hours |
| Awarded wrong | Ask the admin to void the posting number (P-0003) and reopen it. Then award it again. | - |

## Shift switch and payback
1. Click **Shift Switch**. Type the Shift date and pick the Unit.
2. Pick "Gives up shift" and "Works the shift". The Shift fills in.
3. Read the hint lines. "NOT ELIGIBLE" stops the save. "check" is a warning only.
4. Click **SAVE SWITCH**, then OK. It stays on the OWED list until paid back.

- **Payback:** click the switch in the OWED list. Type the payback date. Click **SAVE SWITCH**.
- **Both dates known now:** tick **Two-way swap**, type the payback date, and save once.
- **Wrong switch:** select it in the list and click **Void selected**. Give the reason.

## Book time off
1. Click **Book Time Off**. Pick Employee, Unit and Type. Type From and To (last day).
2. Types: VAC vacation, PER personal day (unpaid), TRN training, or a leave. For a leave, the **Leave** button is better (see below).
3. Vacation and personal days go on scheduled shifts only. A range with no scheduled shift is refused.
4. If the range starts or ends on days off, you are asked to trim it. Yes trims. No saves as typed. Cancel goes back.
5. No return date yet: tick "Open-ended". Part of a shift: set "Hours / shift".
6. Click **SAVE**. To fix a booking, select it under BOOKED TIME OFF. Click **Change end date** or **Void**.

## Sick call
1. Click **Sick Call**. Pick who called in. Check the Shift date, Shift, Unit to cover and Hours.
2. Read the line under the Note box. It gives the person's counts and "This shift: Paid sick day ..." or "Unpaid - ..." with the reason.
3. Click **SAVE + START CALL-OUT**. Log Call-Out opens filled in. Type the hours, click **LOAD LIST** and call as usual.
4. **Save only** books the sick shift without starting a call-out.

The person is off call lists for that shift and any shift starting within 24 hours after it ends. Log each sick shift as its own sick call.

## Leave: going and returning
- **Going:** click **Leave**, choose "Going on leave". Pick Employee and Type. Type the first day off. Type the expected return, or leave it blank. Click **SAVE LEAVE**.
- **Returning:** click **Leave**, choose "Returning from leave (top-up)". Pick who is back and type the return date (first day back). Click **Work out the top-up**. Read the result: the peers, the average, and hours before and after. Click **SAVE TOP-UP**, then OK.
- Use "Override hours" only if the average is wrong. A reason is required.

## Needs Attention
Open it from the Calendar with the **Needs attention >** button (it reads "N need cover >" when lines are open), or click the Needs Attention tab. It lists open lines, postings ready to award, paybacks owed, leaves to watch and other items. Double-click any row for its options.

## Calendar
Click **Calendar**. Pick the month and the unit at the top. OPEN means the line still needs cover. Double-click a line for its details and options. **Print / PDF** prints the month or saves it as a PDF.

## Print the call-out roster
Click **Call Out Roster**. It is redrawn in call order. Answer Yes to print.

## Publish Member View
Click **Member View**. Wait for "Members' file published". If the message says the file is open somewhere, have it closed and publish again.

## Asking for a void
**Void Entry** is an admin button. To correct a logged call-out or award:
1. Tell the admin the number (S-0012, P-0003) and what was wrong.
2. The admin voids it. The entry stays in the log, marked VOID.
3. Log the correct entry.

## The status line
The three lines above the roster show the pay period, the Kronos date and the log size. If it says "Integrity: n PROBLEMS" in red, tell the admin before logging.

## Close the file
- Saving is automatic after each entry, and again when you close. Just close the file.
- If a red "Saving is blocked" message appeared, close WITHOUT saving. Reopen the file. Check your last entry and redo it if it is missing. Tell the admin.
