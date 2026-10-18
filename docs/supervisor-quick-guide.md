# Call Out Roster: Supervisor Quick Guide

How to do each everyday job in the roster. Current as of v25.1 (Oct 2026).

The buttons are in the bar at the top of the **OT Roster** tab. **Log Call-Out**, **Calendar**, **Sick Call**, **New Posting**, **Award Posting** and **Call Out Roster** show for everyone (v25.3). The other buttons show only for logins on the admin list. If a button is missing, ask the admin.

## Open the file
1. Open the roster from its usual folder. Click **Enable Content** if Excel asks.
2. A notice that someone "has the roster open for editing, so this copy is read-only" means you can look but not log. Close it. Open it again when they are out. "This is a frozen copy of the roster" means you opened a dated backup or the year-end FINAL file: look, but log in the live roster.
3. An admin may see "N planned unit move(s) are due: apply them now?" Yes applies them (a backup is saved first). No leaves them waiting; the question comes back at the next open.
4. Close the file as soon as you are done, so the next supervisor can get in.
5. Every sheet is protected. Nothing is typed into cells: the buttons do the writing. (Admins can type into the roster's Home Unit, Crew, Phone, Worked, Seniority, EMAIL and unit code columns, and the Holidays sheet's "Observed on"; everything else, totals included, is refused, and so is sorting a block of cells by hand: the sort buttons do it.)

## Typing a date
Every date box takes "Oct 12", "12 Oct", "Oct 12, 2026" or "10/12" (month then day), and shows the full date it understood. With no year typed, a date more than three months back means next year ("Jan 5" typed in October is next January) and a date more than nine months ahead means last year ("Dec 20" typed on January 3 is the December just gone). A time typed after the date is ignored. The box tells you when it can't read what you typed.

## If a red "This didn't finish" box appears
Something stopped part-way. The box says "Error Log line N" and the name of what you were doing. Nothing you did before the error is lost; it is saved when you close (unless the box says this copy can't be saved). Check what you were doing (run it again if it was an import), and tell the admin the line number. If the same thing fails twice, stop and ask.

## Log a call-out
1. Click **Log Call-Out**. Pick Department, Shift and Date. Type the Hours. Leave Crew on "All not scheduled".
2. Click **LOAD LIST**. The list is in call order and holds only people free for that shift.
3. Call the name shown at "Now calling". Click **ACCEPTED**, **REFUSED**, **LEFT MSG** or **UNAVAIL.** For Unavailable, pick the reason first.
4. When the shift is filled, click **SUBMIT ALL ENTRIES**. Note the session number (S-0012) in the message.
5. If asked which calendar line the person covers, type its number.

Refused and Left Message are charged (grade 2 and below). `>>` means above grade 2, called but not charged. `*` means the total includes accepted overtime not yet in Kronos. `~` means the hours are an estimate because the person has a unit move planned for a later date (the admin entered it ahead of time). If warned that Kronos data is old, No stops so the admin can import first.

## Part of a call-out
If someone takes only part of the hours, select them and click **ACCEPTED PART OF THE HOURS**, then enter the hours. **That closes the shift:** nobody else is called, nobody below is charged, and the pop-up refuses any further outcome (the partial taker's own entry can still be changed). Anyone who refused BEFORE the partial keeps their charge. Leave the rest of the list uncalled and click **SUBMIT ALL ENTRIES**.

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
| Someone from another department signed a blank line | Pick the name in the "Out of department" box (listed in call order, grade first). Click **AWARD OUT OF DEPT**. | In-department people who did not sign up |
| Nobody took it | Click **NO TAKERS**. | Everyone eligible on the list |
| Overtime no longer needed | Click **CANCEL REQUEST**. | Nobody |
| Only part of the shift was taken | Award the partial sign-up(s). The rest of the hours drop off. | People ahead who did not sign up at all, the posting's full hours |
| Out of department takes part of the shift | Type the hours in the **hrs** box beside the name. Only allowed when nobody in the department signed up. | In-department people who did not sign up, the posting's full hours |
| Awarded wrong | Ask the admin to void the posting number (P-0003) and reopen it. The reopen names anyone who is no longer chargeable on it (moved, left). Then award it again. | - |

## Shift switch and payback
1. Click **Shift Switch**. Type the Shift date and pick the Unit.
2. Pick "Gives up shift" and "Works the shift". The Shift fills in.
3. Read the hint lines. "NOT ELIGIBLE" stops the save. "check" is a warning only. Someone inside the 24 hours after a sick shift can neither work nor give up a shift (the pick list leaves them out). A shift given up in a switch still keeps the person off the other shift that day.
4. Click **SAVE SWITCH**, then OK. It stays on the OWED list until paid back.

- **Payback:** click the switch in the OWED list. Type the payback date. Click **SAVE SWITCH**.
- **Both dates known now:** tick **Two-way swap**, type the payback date, and save once.
- **Wrong switch:** select it in the list and click **Void selected**. Give the reason.

## Book time off
1. Click **Book Time Off**. Pick Employee, Unit and Type. Type From and To (last day). A date before this roster year's January 1 is refused.
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
5. **Past sick days** (a range already gone): the box says how many days were recorded and how many were skipped because there was no scheduled shift or the day was already booked sick.

The person is off call lists for that shift and any shift starting within 24 hours after it ends. Log each sick shift as its own sick call. The calendar shows one "Sick" line on the shift called in; members see it as "Off".

## Leave: going and returning
- **Going:** click **Leave**, choose "Going on leave". Pick Employee and Type. Type the first day off. Type the expected return, or leave it blank. Click **SAVE LEAVE**.
- **Returning:** click **Leave**, choose "Returning from leave (top-up)". Pick who is back and type the return date (first day back). Click **Work out the top-up**. Read the result: the peers, the average, and hours before and after. Click **SAVE TOP-UP**, then OK.
- Use "Override hours" only if the average is wrong. A reason is required.
- A return dated after the roster's last pay period is refused until the year-end rollover has been run. A return date in the future is refused: log the return on or after the first day back.
- A top-up of 0 hours is still logged (an LV- row with "no hours added"), so the return is on record and can't be logged twice.

## Needs Attention
Open it from the Calendar with the **Needs attention >** button (it reads "N need cover >" when lines are open), or click the Needs Attention tab. It lists open lines, postings ready to award, paybacks owed, leaves to watch, planned unit moves and other items. Double-click any row for its options. A due planned move's row asks "Apply it now?": Yes applies that one move; No asks "Cancel the move?".

## The Desk
Every supervisor opens on the **Desk** tab: tiles for what is waiting today and the buttons for a supervisor's jobs (Someone called in sick, Fill a shift now, What needs cover, Post a shift, Award a posting, Print the call-out roster, help). Since v25.3 a supervisor may book a sick call, post a shift and award a posting; the bar shows those buttons too. Back to the Desk: click its tab. If your login is on the admin Desk list you open on the fuller Desk: four tiles for what is waiting today and fourteen plain buttons (Someone called in sick, Fill a shift now, Post a shift, Award a posting, Book time off, Shift switch or payback, Leave, Update the members' file, Print the call-out roster ...). Each opens the usual pop-up. You see a supervisor's tabs plus the Desk and Needs Attention; **Show everything** at the bottom gives the full roster for the session and **< Desk** on the bar brings the Desk back. `docs/desk-guide.md` is the one-page version.

## Calendar
Click **Calendar**. Pick the month and the unit at the top. OPEN means the line still needs cover. Double-click a line for its details and options. **Print / PDF** prints the month or saves it as a PDF.

## Print the call-out roster
Click **Call Out Roster**. It is redrawn in call order. Answer Yes to print.

## Publish Member View
Click **Member View**. Wait for "Members' file published". If it asks "The members' file is open in this Excel ... Close it and publish?", Yes closes it and publishes. If it says the file couldn't be written, it is open on a plant PC: have it closed there and publish again (nothing was written). "NOT published. Something private would have gone into the members' file" names a calendar line or posting reason that mentions a kind of leave, sick or personal: reword it and publish again. Member View refuses to run from a dated backup or the frozen year-end file: open the live roster.

## Postings when someone moves unit
A person who moves into a unit (today, or on a date the admin entered ahead of time) is added to that unit's open postings for shifts from the move that can still be signed up for, in call order. Print those sheets again; the box after the move names them. If a planned move is cancelled before the day, they come off again.

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
