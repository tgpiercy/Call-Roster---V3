# Call Out Roster: Supervisor Quick Guide

How to do each everyday job in the roster. Current as of v27.5 (Oct 25, 2026; v27.5 changes one job here - the shift switch's 16 h stop - and the staffing flag's (part) word). The same text is on the roster's **Help** tab (Desk > How do I ...).

The buttons are in the bar at the top of the **OT Roster** tab. **Log Call-Out** and **Calendar** show for everyone. **Sick Call**, **New Posting**, **Award Posting**, **Call Out Roster**, **Member View** and **Find a person** show for supervisors (a login on the supervisor list) and admins (v25.5). The other buttons show only for logins on the admin list. If a button is missing, ask the admin. Keys while the roster is the active workbook: Ctrl+Shift+S sick call, Ctrl+Shift+C call-out, Ctrl+Shift+K calendar, Ctrl+Shift+D the Desk.

## Open the file
1. Open the roster from its usual folder. Click **Enable Content** if Excel asks.
2. A notice that someone "has the roster open for editing, so this copy is read-only" means you can look but not log. Close it. Open it again when they are out. "This is a frozen copy of the roster" means you opened a dated backup or the year-end FINAL file: look, but log in the live roster.
3. An admin may see "N planned unit move(s) are due: apply them now?" Yes applies them (a backup is saved first). No leaves them waiting; the question comes back at the next open.
4. Close the file as soon as you are done, so the next supervisor can get in.
5. Every sheet is protected. Nothing is typed into cells: the buttons do the writing. (Since v26.5 nothing on the roster is typed into either: an admin double-clicks a person's row and the Employee pop-up opens on Edit. Since v27 the Holidays and the minimum staffing have their own pop-up too (the admin Desk's Holidays / minimum staffing). Only the Reconcile sheet's typed paper order and the Rules sheet's Basis column take typing; everything else, totals included, is refused, and so is sorting a block of cells by hand: the sort buttons do it.)

## Typing a date
Every date box takes "Oct 12", "12 Oct", "Oct 12, 2026" or "10/12" (month then day), and shows the full date it understood. With no year typed, a date more than three months back means next year ("Jan 5" typed in October is next January) and a date more than nine months ahead means last year ("Dec 20" typed on January 3 is the December just gone). A time typed after the date is ignored. The box tells you when it can't read what you typed.

## After every save: the receipt (v27)
Every job that writes something ends with one card: **What was recorded** (the rows and their numbers), **Who it changed** (each person whose Total Hours moved, before -> after; the cells are marked amber on the roster with a note for a minute or until the next save), **Where to look**, and **If it is wrong** (who to ask). A line "On paper this would have been different" names a rule the paper roster never had (a part acceptance closing the shift, the 16 h rule, a DST night's actual hours). **Print receipt** saves it as a PDF, **Email me this** opens a mail draft. **Was this right?** Yes / Not sure / No (with a line if you like) goes to the admin. Double-click any Total Hours cell for how it is made up.

## If a red "This didn't finish" box appears
Something stopped part-way. The box says "Error Log line N", the name of what you were doing, and **whether it was saved** ("Newest OT Log row: S-0012 Accepted ... at 3:10 AM", or none by you today); it offers to email the admin the details. Nothing you did before the error is lost; it is saved when you close (unless the box says this copy can't be saved). Check what you were doing (run it again if it was an import), and tell the admin the line number. If the same thing fails twice, stop and ask.

## Log a call-out
1. Click **Log Call-Out** (or Fill a shift now on the Desk, or Ctrl+Shift+C). Pick Department, Shift and Date. Type the Hours. Leave Crew on "All not scheduled".
2. Click **LOAD LIST**. The list is in call order and holds only people free for that shift. **Not listed (N)...** beside the Hours box says who was left off and why (scheduled, rest rule, booked off, switch, already on OT, 16 h rule, OT restricted). A unit's standing note (set by the admin) shows in the line under the date. If a call-out for the same unit, date and shift was logged in the last 12 hours, you are asked first. The card under the list shows who to call: name, phone, and one line with their code, hours and anything that matters.
3. Call the person on the card. Press **A** (accepted), **R** (refused, charged x1.75), **L** (left message, charged x1.75) or **U** (unavailable, not charged: the reason list gets the focus, pick one, Enter records it), or click the button - each button says what it costs. The next person comes up. **K** skips for now (the row keeps "-"), **B** goes back one, **P** is part of the hours (that closes the shift). **Call back (?)** goes down the pile of people you left a message with (marked "?"). Someone marked `[max N h]` is offered only those hours: a refusal or a message is charged on them, not on the whole shift. **U** with the reason "Plant upset ..." is the **Exception** outcome: nobody is charged, and the admin sees it on Needs Attention.
   **The draft (v27):** every outcome is kept as you go and the file is saved every 2 minutes, so if Excel closes mid-call the next Log Call-Out offers "Resume the ... session (N logged)?" - Yes puts the list and the outcomes back.
4. When the shift is filled, press **Enter** or click **SUBMIT ALL ENTRIES**. The receipt card says what was logged and what happens to the line next (covered; closed at N h; or stays OPEN - call back, post it, or ask the admin to award out of department).
5. The acceptance covers the calendar line by itself when the unit has ONE open line for that date and shift (or the call-out was started from that line); with several, you are asked which.

Refused and Left Message are charged (grade 2 and below). `>>` means above grade 2, called but not charged. `*` means the total includes accepted overtime not yet in Kronos. `~` means the hours are an estimate because the person has a unit move planned for a later date (the admin entered it ahead of time). **`[max N h]` (v26.5.5)** means the 16 h rule leaves them less than the offer: overtime they already accepted, an award or a switch touches this shift, so **A** is refused and **P** up to that many hours is allowed. **Before 7 AM the date box is yesterday** - the night under way is dated by the evening it started; the Hours box fills itself (12, or the DST night's hours) unless you type. The grey line at the bottom of the pop-up says how current the worked hours are and when the members' file was last updated; if it says in red that the Kronos hours are old, the call order may be off - tell the admin to import.

## Part of a call-out
If someone takes only part of the hours, select them and press **P** (Part of the hours), then enter the hours. **That closes the shift:** nobody else is called, nobody below is charged, and the pop-up refuses a further Accepted (the partial taker's own entry can still be changed; a refusal or message logged before the close can be corrected to a refusal, a message or unavailable). Anyone who refused BEFORE the partial keeps their charge. Leave the rest of the list uncalled and click **SUBMIT ALL ENTRIES**.

## Postings
**Post a shift**
1. Click **New Posting**. Fill in Unit, Shift date, Shift, Reason, Hours and Positions. Crew and "Posting closes" fill in.
2. Leave "Print it" and "Save PDF" ticked. Click **CREATE POSTING**, then OK.
3. Put up the printed sheet. Sign-up is on paper or, since v26.5, from the members' file (Sign me up).

**Load a posting:** click **Load Posting**, pick it, then **View on sheet**, **Print** or **Save PDF**.

**Award it**
1. Click **Award Posting**. Pick the posting and click **LOAD** (keys: Space = signed up, A = award). Electronic sign-ups are ticked already; the summary line names late ones, withdrawals and any name not on the roster.
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
| The award has to be taken back (v27) | The admin's **Withdraw an award**: the awarded rows are voided, the refusals stay charged, the calendar line opens again and a call-out is offered. | The refusals stay |
| The hours, date or positions changed (v27) | The admin's **Edit a posting** (an open posting; a new date that changes who is free needs a new posting). | - |

## Electronic sign-ups (v26.5)
Members can sign up from the members' file: on its Postings sheet, **Sign me up** on an open posting sends a small request file into the Signups folder beside the members' file (Withdraw takes it back). When you LOAD that posting in Award Posting, those sign-ups are ticked already and the summary line names any that came after the closing time, any withdrawals (with the login that withdrew) and any file naming someone not on the roster. Your tick is what counts, and the paper sheet stays valid.

## Shift switch and payback
1. Click **Shift Switch**. Type the Shift date and pick the Unit.
2. Pick "Gives up shift" and "Works the shift". The Shift fills in.
3. Read the hint lines. "NOT ELIGIBLE" stops the save. "check" is a warning only. v27.5: the save also stops when the one WORKING the shift would run past 16 h with the shift beside it - a scheduled shift counts (nights last night, then days today); in a same-day days-for-nights swap both people are checked on the shift each one works, and the shift each gives up does not count. Someone inside the 24 hours after a sick shift can neither work nor give up a shift (the pick list leaves them out). A shift given up in a switch still keeps the person off the other shift that day.
4. Click **SAVE SWITCH**, then OK. It stays on the OWED list until paid back.

- **Payback:** click the switch in the OWED list. Type the payback date. Click **SAVE SWITCH**.
- **Both dates known now:** tick **Two-way swap**, type the payback date, and save once.
- **Wrong switch:** select it in the list and click **Void the selected switch**. Give the reason.

## Book time off
1. Click **Book Time Off**. Pick Employee, Unit and Type. Type From and To (last day). A date before this roster year's January 1 is refused.
2. Types: VAC vacation, PER personal day (unpaid), TRN training, or a leave. For a leave, the **Leave** button is better (see below).
3. Vacation and personal days go on scheduled shifts only. A range with no scheduled shift is refused.
4. If the range starts or ends on days off, you are asked to trim it. Yes trims. No saves as typed. Cancel goes back.
5. No return date yet: tick "Open-ended". Part of a shift: set "Hours / shift".
6. Click **SAVE**. To fix a booking, select it under BOOKED TIME OFF. Click **Change end date** or **Void this booking**.
7. **Post all lines (v27):** after the save you may be offered one posting per scheduled shift of the booking (each on its own calendar line, a PDF each); one summary says what was made.

## Sick call
1. Click **Sick Call**. Pick who called in. Check the Shift date, Shift, Unit to cover and Hours.
2. Read the line under the Note box. It gives the person's counts and "This shift: Paid sick day ..." or "Unpaid - ..." with the reason.
3. Click **SAVE + START CALL-OUT**. A box reads the booking back first (who, the date and shift, the unit, until when they are off the lists): **Yes** books it, **No** saves nothing (v26.5.5). Log Call-Out then opens filled in, hours included; click **LOAD LIST** and call as usual.
4. **Save, no call-out** books the sick shift without starting a call-out (the same read-back box first). Wrong person? Only an admin can void it: ask. Before 7 AM the date box is yesterday: the night under way is dated by the evening it started.
5. **Past sick days** (a range already gone): the box says how many days were recorded and how many were skipped because there was no scheduled shift or the day was already booked sick.

The person is off call lists for that shift and any shift starting within 24 hours after it ends. Log each sick shift as its own sick call. The calendar shows one "Sick" line on the shift called in; members see it as "Off".

## Leave: going and returning
- **Going:** click **Leave**, choose "Going on leave". Pick Employee and Type. Type the first day off. Type the expected return, or leave it blank. Click **SAVE LEAVE**.
- **Returning:** click **Leave**, choose "Returning from leave (top-up)". Pick who is back and type the return date (first day back). Click **Work out the top-up**. Read the result: the peers, the average, and hours before and after. Click **SAVE TOP-UP**, then OK.
- Use "Override hours" only if the average is wrong. A reason is required.
- A return dated after the roster's last pay period is refused until the year-end rollover has been run. A return date in the future is refused: log the return on or after the first day back.
- A top-up of 0 hours is still logged (an LV- row with "no hours added"), so the return is on record and can't be logged twice.

## Needs Attention
Open it from the Calendar with the **Needs attention >** button (it reads "N need cover >" when lines are open), or click the Needs Attention tab. It lists open lines, **shifts below their minimum staffing (v26: the next 14 days, every unit whatever group is picked; double-click opens the calendar on that date)**, postings ready to award, paybacks owed (oldest first; one owed over 42 days is in red on amber since v26), leaves to watch, planned unit moves and other items. Double-click any row for its options. A due planned move's row asks "Apply it now?": Yes applies that one move; No asks "Cancel the move?".

## The Desk
A supervisor (a login on the supervisor list) opens on the **Desk** tab. The three pills at the top say which crews are on today, the pay period, and how old the Kronos worked hours are (red when they need an import; an admin's pill says "N new Kronos report(s) - click to import" when one is waiting). The **TODAY** tiles say what is waiting and what a click does: shifts needing cover (the next one named), postings ready to award, the members' file, and **Since you were last here** (what others logged since your last saved visit; click for the list). The jobs sit in three bands by when you reach for them: **NOW** (Someone called in sick, Fill a shift now, What needs cover), **PLAN** (Post a shift, Award a posting, Paybacks owed - the ledger) and **FINISH** (Update the members' file, Print the call-out roster, Employee history, Shift handover - the last 12 hours on one page). The grey row is help: How do I ..., Something went wrong, Suggest an improvement, Larger type (your pop-ups at 120 %), Practice on a copy (a throw-away copy with a red PRACTICE COPY banner - nothing done in it counts). Under it: the keys and the legend; for your first three visits a "Your first days" line ticks off four things to try. If the Desk header says **NOT SAVED since <time>** in red, the last save failed: try File > Save, and tell the admin if it fails again.

An admin on the Desk list gets the fuller Desk: six tiles (+ paybacks owed and **This morning** - the integrity check, today's backup, unread errors) and the admin's jobs (time off, switches, leave, Holidays / minimum staffing, Reconcile, Stats, the Rules sheet, Close the pay period, Withdraw an award / edit a posting). **Show everything** gives the full roster for the session and **< Desk** on the bar brings the Desk back. A login on neither list opens on the roster and can log call-outs and look at the calendar. `docs/desk-guide.md` is the one-page version.

## Calendar
Click **Calendar** (or Ctrl+Shift+K). Pick the month and the unit at the top. OPEN means the line still needs cover. Click a line: its details show in the row under the title and on the status bar (no hover boxes since v25.5). Double-click a line for its options; a supervisor's menu on an OPEN line has **Fill this shift now (call-out)...** (the pop-up opens on that unit, date and shift, and the acceptance covers that line), Post this line and Award posting. **Print / PDF** prints the month or saves it as a PDF (since v26.5 as the members' copy: a sick call or a personal day reads Off on the printed page). **v26:** a date in the next 14 days where any unit is below its minimum staffing shows a red "N short" flag in its date cell, whatever unit group the calendar shows (click the date: the row under the title says which unit and shift); **Board** draws the two-week board (today and the next 13 days, one page per unit group) and saves it as a PDF, then asks whether to print it.

## Employee history (v26)
**Employee history** (the FIND group on the bar, or the Desk) opens a pop-up (v27): pick or type the name, From and To (the whole roster year to start), member-safe or (an admin) FULL, and whether to save a PDF. The History Report tab lists every OT Log entry for that person in the range, newest first: call-outs with their place in the call order, postings, adjustments, the hours charged and a **Why** line for each ("Charged 21 h: 3rd on the call list, Refused, 12 h offered x 1.75 (grade P2)"), then the charged hours per pay period and the total. The PDF tick puts it beside the roster. It is the same report the members' file shows as My history, row for row, so a question at the board can be answered from either. An admin is also offered the full report (who logged each entry and when). **Why on the status bar:** click any row on the OT Log tab and the status bar at the bottom explains that row's charge. Since v26.5 the pay-period block shows the charged hours, the posting refusals, the call-out refusals and Worked (Kronos).

## Print the call-out roster
Click **Call Out Roster**. It is redrawn in call order, with the phone numbers, A / R / L / U tick boxes and a Time column, a session line to fill in under each unit, an Order line (grade, then the lowest Total Hours, then seniority; `=` marks a row tied with the one above on grade and hours) and the legend in the footer - the paper to run a call-out from when the roster is down. Answer Yes to print. The phone column is hidden again before the file is saved.

## Shift handover and paybacks (v27)
**Shift handover (last 12 h)** (the Desk's FINISH band) draws one page of everything logged in the last 12 hours - call-outs, awards, adjustments, time off and sick calls booked (leave shows as Leave), switches entered - and offers to print it. **Paybacks owed (ledger)** (PLAN) lists every switch still waiting for its payback, oldest first, the ones over the warning age in amber; type part of a name for that person's statement and save it as a PDF to hand them.

## Publish Member View
Click **Member View** (Update the members' file on the Desk; supervisors and admins since v25.5). Wait for "Members' file published". If it asks "The members' file is open in this Excel ... Close it and publish?", Yes closes it and publishes. If it says the file couldn't be written, it is open on a plant PC: have it closed there and publish again (nothing was written). "NOT published. Something private would have gone into the members' file" names a calendar line or posting reason that mentions a kind of leave, sick or personal: reword it and publish again. Member View refuses to run from a dated backup or the frozen year-end file: open the live roster.

## Postings when someone moves unit
A person who moves into a unit (today, or on a date the admin entered ahead of time) is added to that unit's open postings for shifts from the move that can still be signed up for, in call order. Print those sheets again; the box after the move names them. If a planned move is cancelled before the day, they come off again.

## Asking for a void
**Void Entry** is an admin button. To correct a logged call-out or award:
1. Tell the admin the number (S-0012, P-0003) and what was wrong.
2. The admin voids it. The entry stays in the log, marked VOID.
3. Log the correct entry.

## The status line
The three lines above the roster show the pay period, the Kronos date and the log size. If it says "Integrity: n PROBLEMS" in red, tell the admin before logging.

## Words you will see
- **Call order:** grade code first, then the lowest Total Hours, then plant seniority. Everyone is called in that order.
- **Total Hours:** worked hours (Kronos) + refused hours whose pay period has started (a refusal for a shift in a later period counts from that period; v26.5) + accepted overtime not yet in Kronos (shown with `*`).
- **Charged:** refused or left-message hours x 1.75, added to Total Hours; grade 2 and below only. `>>` = above grade 2: called, never charged.
- `~` = an estimate, because the admin has entered a unit move for the person for a later date. `?` = left a message (the call-back pile). `S/S` = working a shift switch that shift.
- **NCN** = no cover needed. **PHA** = hazard-analysis coverage. **OPEN** = the line still needs cover. **POSTED** = a posting is out for it. **DAYS** = the Monday-to-Friday day crew.
- A sick call or a personal day shows to members as **Off**, never why.
- The calendar's words: OPEN (red) = nobody on it yet, POSTED (orange) = waiting for sign-ups, NO TAKERS = posted, nobody signed up, NCN = no coverage needed, a name = covered, S/S = shift switch, N short = below minimum staffing ((part) after it, v27.5 = a partial cover is counted as a share of a person, not a whole one). Help > Words on the roster has the full glossary.

## Close the file
- Saving is automatic after each entry, and again when you close. Just close the file.
- If a red "Saving is blocked" message appeared, close WITHOUT saving. Reopen the file. Check your last entry and redo it if it is missing. Tell the admin.
