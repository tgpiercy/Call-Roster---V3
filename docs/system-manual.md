# Call Out Roster: System Manual

This manual states the current rules of the overtime call-out system, for supervisors and the union committee.

Current as of v24.3.1 (Oct 2026). It replaces the rules in all older guides.

## 1. What the system is

- One Excel workbook, the master roster, holds everything.
- One person edits it at a time. A second person gets a read-only copy and a notice saying who has it open.
- Every call-out, posting result and adjustment is a row in the **OT Log**.
- Roster totals are formulas over the OT Log. Nobody types hours into the roster.
- Members see a separate file, published from the master. It holds values only.

### Words used here

| Word | Meaning |
|---|---|
| Unit | Phosphine, Derivatives, Cypure, Material Handler (Fumigation), QC Derivatives (QC Lab), QC Gases, Utilities, SBP |
| Code | The grade code in a person's unit column, for example D1.5 or P2. The number is the grade. |
| Home Unit | The one unit a person belongs to (roster column B). Used for sick calls and the leave top-up. |
| `~` | On the call-out pop-up and the printed call-out roster: the hours shown are an estimate, because the person has a unit move planned for a later date (section 15). |
| S/S | Shift switch |
| NCN | No Coverage Needed |
| PHA | Hazard-analysis coverage |
| The admin | A person whose Windows login is on the admin list |

## 2. Shifts and the rotation

| Item | Rule |
|---|---|
| Shift length | 12 hours |
| Days | 7 AM to 7 PM |
| Nights | 7 PM to 7 AM. A night shift is dated by the evening it starts. |
| Crews A, B, C, D | A 28-day rotation of days and nights. It repeats, so it works for any date. |
| Crews A/B and C/D | Work their pair's day shifts whenever either crew of the pair is on days. Never nights. |
| DAYS | Monday to Friday days. Off on plant holidays. |
| Which crew | The crew a person was on **that date**. The Crew History sheet holds each crew with the date it started, so a crew change does not move earlier time off or switches onto other shifts. The roster shows the newest crew. |
| Clock-change nights | The night starting the Saturday before the first Sunday in November is 13 hours. The night starting the Saturday before the second Sunday in March is 11 hours. Day shifts are not affected. |

## 3. Plant holidays

The system works out the same 11 holidays for any year: New Year's Day, Family Day, Good Friday, Victoria Day, Canada Day, Civic Holiday, Labour Day, Thanksgiving, Remembrance Day, Christmas Day, Boxing Day.

| Who | Rule on a holiday |
|---|---|
| DAYS workers | Off. They are free for overtime that day. |
| A, B, C, D, A/B, C/D | Scheduled as usual. |

- When a holiday falls on a weekend, the plant announces the weekday given off.
- The admin enters that date in "Observed on" on the Holidays sheet.
- Until it is entered, Needs Attention lists the holiday under HOLIDAYS TO CONFIRM.
- The floater is booked as vacation.
- The holiday name shows on its date on the Calendar.

## 4. Total Hours

**Total Hours = Worked + Refused + Provisional.**

| Part | What it is | Where it comes from |
|---|---|---|
| Worked | Hours worked this year | The Kronos import |
| Refused | Hours charged this year, plus adjustments | Hours Charged in the OT Log for this year's pay periods. Voided rows do not count. |
| Provisional | Accepted overtime that Kronos does not show yet | Accepted call-outs and awarded postings in the OT Log |

### Provisional accepted overtime

- An Accepted call-out or an Awarded posting counts from the moment it is logged.
- It counts even when the shift is in the future.
- The amount is the shift hours, not hours x 1.75.
- It stops counting when Kronos covers the shift. From then on the hours are in Worked.

### The data-through rule

| Rule | Detail |
|---|---|
| Data-through date | Kronos data runs through the day before the Kronos file date. |
| Day shift on date D | Covered when D is on or before the data-through date. |
| Night shift on date D | Covered when D + 1 is on or before the data-through date. |

- A total marked `*` includes provisional hours.
- The roster's status line shows the Kronos date and "data through".
- Call-outs and new postings warn when the Kronos data is more than 7 days old.

## 5. Call order

People are called in this order. Each step only breaks ties in the step before.

| Step | Rule |
|---|---|
| 1 | Grade code in that unit's column, lowest number first |
| 2 | Total Hours, lowest first |
| 3 | Plant seniority, earliest date first |
| 4 | Name, A to Z |

One rule decides the order everywhere. The call-out list, the posting list, the out-of-department box on the award pop-up, the printed roster and the members' file all use it. A Plant Seniority that isn't a date sorts last; IntegrityCheck reports it (and a unit column that mixes code letters) rather than change a cell.

## 6. Who is listed

A call-out list or posting list shows anyone with a code in that unit's column, at any grade, who is free for that date and shift.

A person is not free when one of these applies.

| Reason | Rule |
|---|---|
| Scheduled | Their crew works that date, by the rotation in section 2. |
| Rest rule, day overtime | Their crew comes off nights that morning. |
| Rest rule, night overtime | Their crew starts days the next morning. |
| Booked off | A time-off booking covers the date. |
| Shift switch | They are working a switch that day, or gave up a shift that day (not free for the other shift either). A switch worked the night before or the day after also counts for the rest rule; a shift GIVEN UP in a switch relieves the next day's rest rule, since they did not work it. |
| Sick window | They called in sick. They are off that shift and any shift starting within 24 hours after it ends, and inside that window they can neither work nor give up a shift in a switch. |

Old-calendar vacation lines that were imported are display only. They keep nobody off a list until the admin converts them into bookings.

## 7. Charging

Charged hours go into Refused, so they raise Total Hours.

| Rule | Detail |
|---|---|
| Rate | Shift hours x 1.75. A 12-hour shift charges 21 hours. |
| Who can be charged | Grade 2 and below only. Higher grades are called but never charged. |
| Out of department | A person awarded from another department is never charged. |
| Clock-change nights | Actual hours are used. A 13-hour night charges 22.75. An 11-hour night charges 19.25. |

| Event | Who is charged | Pay period |
|---|---|---|
| Call-out: Refused | That person | Call-out date |
| Call-out: Left Message | That person | Call-out date |
| Call-out: Accepted | Nobody. The hours count as provisional. | - |
| Call-out: part of the hours accepted | The shift is closed: nobody below is offered the rest and nobody is charged for it. Refusals logged before the partial keep their charge. (Since v25; before it the rest was offered down the list.) | - |
| Call-out: Unavailable (a reason is required) | Nobody | - |
| Posting: Submit Award | People ahead of the last awarded person in call order who did not sign up | Shift date |
| Posting: Award Out of Dept | Everyone eligible in the department who did not sign up | Shift date |
| Posting: No Takers | Everyone eligible on the list | Shift date |
| Posting: Cancel Request | Nobody | - |

## 8. Postings

| Rule | Detail |
|---|---|
| The list | Made when the posting is created. It holds everyone free for that shift, in call order. |
| Posted totals | Each person's Total Hours is frozen on the list at creation. Awards and charges use these posted totals, not later ones. |
| Sign-up | On the printed sheet. The sheet has 3 blank lines for out-of-department sign-ups. |
| Closing | Default is 14 days before the shift, at 8:00 AM. If that has passed, the day before the shift. The supervisor can type another date. A posting counts as closed from the day after its closing date; one test decides it everywhere (ready to award, still open, can still be signed up for). |
| Awarding early | The system warns when a posting is awarded before it closes. |
| Who is awarded | The supervisor marks who signed up and who is awarded. The system does not pick. |
| Signed up, ahead, not awarded | The system warns. These people are not charged. |
| More than one position | "Ahead" means ahead of the last awarded person in call order. |
| Electronic sign-ups | When switched on, members can sign up from their own file. The roster reads those requests when the posting is loaded for award and ticks the names; the supervisor confirms each one. The paper sheet still counts. A request sent after the closing time is shown but not ticked. |
| Partial sign-ups | Someone can sign up for part of the shift. Full-shift sign-ups are awarded first, in call order, even with more hours. A partial is awarded only when no full-shift sign-up is left for that position; the system refuses otherwise. Several partials can share a shift. |
| Hours nobody covers | They drop off. Nobody is charged for them. |
| Charging with a partial award | The same people as any award: ahead of the last awarded person and did not sign up at all. They are charged the posting's full hours x 1.75. A partial sign-up that is not awarded is not charged. |
| Moved unit or left | Someone who moved out of the unit, or left, before the shift is never charged on that posting, even though the frozen list still shows them. |
| Moved into the unit | Someone who moves into a unit, today or on a planned date, is added to that unit's open postings for shifts on or after the move that can still be signed up for, placed among the frozen names by call order. Nobody else's place or posted hours change. If a planned move is cancelled before the day, they come off those lists again. This is the only way a posted list changes after it is made. |
| Award email | A draft opens to the awarded person, with the unit's CC list. The supervisor checks it and sends it. |
| Reopen | A closed posting can be reopened only after its log rows are voided. Nothing is charged twice. A shift before the roster year is never reopened. The reopen names anyone it marks not chargeable (moved out, left, no code in the unit now). |

## 9. Shift switches and paybacks

- A switch is one record for one date. One person works another person's shift.
- The payback is its own record on its own date. It is linked to the first switch.
- A switch stays on the owed list until its payback is logged.
- A two-way swap logs both halves at once.
- Switches change no hours and charge nobody. They are scheduling only.

| Check | Result |
|---|---|
| Either person is booked off that day | The switch cannot be saved. |
| The person giving up the shift is already on a switch that day | The switch cannot be saved. |
| The person working the shift is already working a switch that day | The switch cannot be saved. |
| Either person is inside the 24 h window after a sick shift | The switch cannot be saved (they may neither work nor give up a shift). |
| The rotation looks wrong, for example the person is not on that shift | A warning. The supervisor can still save. |

Needs Attention lists every payback still owed.

## 10. Time off and privacy

| Type | Shows on the roster's calendar as | Shows to members as |
|---|---|---|
| Vacation | Vac | Vac |
| Personal day (unpaid) | Personal | Off |
| Training | Trn | Trn |
| Leave of absence, short-term disability, parental leave, WSIB, long-term disability, other leave | Leave | Leave |
| Sick | Sick | Off |
| Other | Off | Off |

Members see that a person is off on a day. They never see whether it was sick, personal or which kind of leave.

| Rule | Detail |
|---|---|
| Effect | A booking keeps the person off call lists for every date in its range. |
| Open-ended | A booking can have no end date. It runs until an end date is set. |
| Vacation and personal days | Booked on scheduled shifts only. |
| No scheduled shift in the range | The booking is refused. |
| Days off at the start or end of the range | The system offers to trim the booking to the scheduled shifts. The supervisor can also save it as typed. |
| Changes | A booking is never deleted. It is voided with a reason, or its end date is changed. |

### Leave privacy

- Every kind of leave shows only as "Leave" outside the Absences sheet.
- The Absences sheet is for the admin only. It is hidden in every saved copy.
- Before the members' file is published, every piece of free text going into it (calendar lines, posting reasons) is checked for a leave word or the words sick and personal. If one is found, nothing is published and the admin is told which line to reword.

### What members never see

- Phone numbers and email addresses
- The type of a leave, and leave notes
- Whether a day off was sick or personal (members see "Off")
- The date of any sick or personal day beyond that "Off" line
- Void notes
- Who logged an entry, and when
- Notes on log entries
- The Attendance sheet

### What members do see about sick time (since Oct 12, 2026)

Each member can look up a name on the Where I stand card and see that person's five sick-time COUNTS for the calendar year: sick days used, paid sick days left, occurrences used, personal days used, and one of four status phrases ("Paid sick available", "Paid sick only while the current occurrence continues", "Paid sick days used up", "No paid sick until <date>"). No dates of sick days, no reasons. This is the one deliberate exception to the rule above, and the publish refuses anything beyond those five figures.

## 11. Sick calls

1. The supervisor logs the sick shift with the Sick Call button.
2. The system books that shift as sick for that person.
3. The person is off call lists for that shift and any shift starting within 24 hours after it ends.
4. The roster's calendar shows the line as "Sick" and OPEN (members see "Off"). One line per sick call, on the shift that was called in, whether or not it was the person's scheduled shift.
5. With Save + Start Call-Out, Log Call-Out opens filled in, so the supervisor can fill the shift.

| Sick shift | Also off | Free again for |
|---|---|---|
| Monday days | Monday nights, Tuesday days | Tuesday nights |
| Monday nights | Tuesday days, Tuesday nights | Wednesday days |

- Each sick shift is its own sick call and its own row.
- The form shows the person's sick and personal day counts and whether this shift is paid. See section 13.
- Known limit: the shift-switch name lists use the sick date only, not the 24-hour window.

## 12. Leave and the return top-up

### Going on leave

- The Leave button books the absence. It shows as "Leave".
- The person is off call lists from the first day off.
- The person's Home Unit is recorded on the booking at that time.

### Returning

A person back from leave gets a top-up, so they do not return at the top of every list.

| Item | Rule |
|---|---|
| Top-up | The average Total Hours gained by peers between the first day off and the return date |
| Peers | People with a code in the returner's Home Unit, at grade 2 and below |
| Not counted as peers | The returner, anyone on leave during that time, anyone with no hours history before it |
| Leave across the yearly reset | The top-up counts only from the reset. |
| Return dated after the roster's last pay period | Refused until the year-end rollover has been run; then it is logged on the new year's roster. |
| Return dated in the future | Refused: the return is logged on or after the first day back. |
| Average of 0 hours | Still logged, as an LV- row with "no hours added", so the return is on record. |
| Rounding | To 0.1 hour |
| Before saving | The system shows the peers, the dates, the average, and the returner's hours before and after. |
| How it is recorded | One Adjustment row in the OT Log, numbered LV-####, with the calculation in its Note |
| Once only | A leave cannot be topped up twice. |
| Override | The admin can enter a different figure. A reason is required and is recorded. |
| Undo | Void Entry on the LV-#### number removes the hours. |

- Hours history began on Oct 9, 2026. For a leave that started earlier, the start figures are estimated, and the system says so.
- If no peers are found, nothing can be saved without an override and a reason.

## 13. Attendance: sick and personal days

**This is tracking only. It never changes Total Hours, call order, charging or the OT Log.** The roster records what the rule says. It does not pay anyone.

| Rule | Detail |
|---|---|
| Year | Calendar year, from Jan 1. This is not the pay-period year in section 16. |
| Paid sick days | 5 a calendar year, in up to 3 occurrences |
| A day | One scheduled shift. A partial shift counts as a day. |
| An occurrence | Consecutive scheduled shifts missed sick. Days off in between do not break it. |
| Unpaid sick day | The 6th sick day, or any day in a 4th occurrence |
| Personal days | 5 unpaid personal days a calendar year |
| Over 5 combined | When sick days plus personal days pass 5, the person is not entitled to paid sick days from that day. |
| How long | Until one year after the first day of their 2nd occurrence. If they had only one occurrence, one year after its first day. If they had none, one year after the day they passed 5. |
| After that | The counts restart and run to Dec 31. The normal Jan 1 cycle then resumes. |
| Across Jan 1 | A wait that has started carries over Jan 1. |
| Sick on an unscheduled shift | Counts as a day. |
| Leave | A leave is not a sick day and is not counted. |

| Rule | Detail |
|---|---|
| Start of counting | Counting started when v22.1 was installed, in October 2026. Earlier days were not entered. |
| 2026 figures | They cover only what was logged from then on. |
| First complete year | 2027 |
| Paid sick days left | 5 minus the sick days used, and 0 once the person is waiting (over 5 combined), once a 4th occurrence has started, or once the 3rd occurrence has ended (they have worked a scheduled shift after its last sick day). |
| Who sees the counts | The admin, on the Attendance sheet and in Needs Attention; and each member, as five counts on the Where I stand card for the name they pick (section 10). |
| What members see on the calendar | "Off" on the day. Never "Sick" or "Personal", no reasons. |
| Known limit | Occurrences are grouped by the crew rotation. Shift switches are not taken into account. |

## 14. The OT Log

| Rule | Detail |
|---|---|
| Append-only | Rows are added. They are never deleted, and their hours are never edited. |
| Each row records | The event, the person, their call position, their Total Hours then, the outcome, hours charged, and the pay period |
| Audit | Each row has an Entry ID, who logged it, and when. |
| Corrections | A wrong entry is voided, then entered again correctly. |
| A void | The row stays, marked VOID, with the reason, who voided it and when. It stops counting in every total. |
| Who can void | The admin, with Void Entry. A reason is required. |
| Roster cells | Hours are never typed into the roster. The totals are formulas over the log. |
| Refused by kind | Refused is one figure. For information it is also shown as call-out refusals, posting refusals and other (adjustments, opening balance), on the status bar for the selected person, in the members' file and in the Refusals Report. The parts always add up to Refused. |
| Name changes | Employee > Rename changes the name on the roster and in every table together: the log, postings, time off, switches, calendar lines, grade log, hours history and crew history. Hours do not change. Closed years keep the old name. |
| Grade changes | A change to a unit code on the roster is logged with the old code, the new code, who and when. |

| ID starts with | Meaning |
|---|---|
| S- | A call-out session |
| P- | A posting |
| ADJ- | An adjustment |
| LV- | A leave return top-up |
| NH- | A new hire's starting hours |
| UM- | A unit move top-up |
| SC- | A sick call (in the time-off records, never in the OT Log) |
| PM- | A planned unit move (on the Planned Moves sheet, never in the OT Log) |
| SS- | A shift switch (in the switch records) |
| OB- | An opening balance |
| MAN- | Values moved in from typed roster cells |
| # and a number | One row, by its Entry ID |

## 15. Adjustments

An adjustment adds or removes charged hours for one person. Only the admin can make one.

| Rule | Detail |
|---|---|
| Type | Required. One of: Leave return, Unit move, New hire, Correction, Other. |
| Where the type is kept | In the Reason column of the OT Log |
| Note | Required. It holds the reason, or the calculation when a rule worked out the hours. |
| Date | The admin gives the date it applies to. That decides the pay period. |
| Confirm | The system shows the hours, pay period, type and reason before saving. |
| Undo | Void Entry on its number. |
| New hire (NH-) | Starting hours = the home unit's average Total Hours on the start date, among grade 2 and below who are not on leave. The peers and their figures are in the Note. |
| Unit move (UM-) | For a full move only. Hours are raised to the highest Total Hours among grade 2 and below in the new unit on the move date. They are never lowered. An extra sign-off changes no hours. |
| Leaving | Employee > Remove moves the roster row to the Former sheet with the person's final figures. Their log rows stay. Bring back returns them to the roster. |

### Unit moves entered ahead of time (planned moves)

| Rule | Detail |
|---|---|
| Entering | Employee > Move to unit with a date after today. Nothing on the roster changes that day. One row goes on the Planned Moves sheet. |
| Lists before the day | For any shift on or after the date, the person is off the old unit's call-out and posting lists and on the new unit's, at the hours they would have after the top-up (an estimate, marked `~` on the call-out pop-up and the printed roster). The estimate is for call order only; a call-out's log row records their real total. A posting made in that window freezes the estimate on its list, as it freezes everyone's hours. |
| Open postings | Old-unit postings for shifts on or after the date mark the person "Exempt": never charged. New-unit postings for such shifts that can still be signed up for get the person added to their list (section 8). |
| On the day | The first admin to open the file on or after the date is asked whether to apply the due moves. Yes applies them (codes, Home Unit, grade log, the UM- top-up with hours as of the move date), with a backup first and one summary that carries the warnings, the postings marked exempt and the UM- number. No leaves them waiting; the question comes back at the next open, a Needs Attention line applies only that one move ("Apply it now?", then "Cancel the move?"), and ApplyPlannedMoves (Alt+F8) does them all at any time. |
| Changing | Cancel and enter again. A cancel puts the lists back, removes only that move's Exempt marks, takes them off the new unit's lists it had added them to, and undoes a crew change entered with it: before the date the never-effective crew row is deleted; on or after it the old crew is written back dated today, and the message says so. |
| One at a time | A second move for the same person is refused until the first is cancelled or applied. Rename carries a planned move; Remove cancels it. |

## 16. The yearly reset

| Rule | Detail |
|---|---|
| Year | The pay-period year, not the calendar year |
| Pay periods | Two weeks long, starting on a Sunday |
| 2026 | Began Dec 28, 2025 |
| 2027 | Begins Dec 27, 2026 |
| Reset | Totals go back to zero each year. |
| Carry-forward | None |
| Number of periods | 26 in most years. A year has 27 when one more period fits before the next Jan 1. The first is 2033. |

What the rollover does to the records:

| Item | At the rollover |
|---|---|
| Final totals | Each person's final Total, Worked, Refused and Provisional hours are written to the Year End sheet. |
| Frozen file | A read-only copy of the whole roster is saved as "Call Out Roster 2026 FINAL". It is the record of the year. |
| Old OT Log rows | Moved to a sheet named for the year ("OT Log 2026"). This is the only time rows leave the live log. The count and the hours are checked before the save. |
| Rows already logged for the new year | They stay, and count from the first day. A December posting for a January shift is one. |
| ID numbers | They carry on. S-, P- and entry numbers never restart. |
| Worked hours | Set to 0. They stay 0 until the first Kronos export of the new year is imported. |
| Old-year entries | Closed. They cannot be voided after the rollover. |
| Sick and personal day counts | Not touched. They run on the calendar year. |

The steps are in docs/rollover-one-pager.md.

## 17. The members' file

| Rule | Detail |
|---|---|
| What it is | A separate workbook for members, opened on plant PCs |
| Values only | The master works everything out and pushes values in. The members' file holds no rules and never reads the master. |
| When it updates | When the admin clicks Member View. Needs Attention reminds the admin when it is more than a day old. |
| Privacy check | Before publishing, the master checks every published value for emails, phone numbers, employee numbers and leave codes, and every free-text line for a leave word or the words sick and personal. If it finds one, nothing is published and the message names the line. |
| From a copy | Member View refuses to run from a daily backup, the frozen year-end file or a "Backup before" copy. A practice copy publishes a file named PRACTICE beside itself, never the real members' file. |

| Sheet | What it shows |
|---|---|
| Home | When the file was updated and the Kronos data date |
| Roster | Everyone's hours and each unit's call order. A member can find their own name. |
| Who's next | For a unit, date and shift in the next 14 days: who is free, in call order |
| Calendar | Who is off and what needs cover, by month and unit. Leave shows as "Leave"; sick and personal days as "Off". |
| Postings | Open postings with their closing dates and lists. With electronic sign-ups switched on, a Sign me up button. |
| My history | For any name: this year's entries, outcomes and charges, by pay period, with Refused split into call-out, posting and other |
| Where I stand | For the name picked with Find me: their call position in each unit, their open postings, and their five sick-time counts (section 10) |

## 18. Who can do what

| Who | Can do |
|---|---|
| Anyone who opens the roster | Sort and filter the roster, use Log Call-Out, look at the Calendar |
| Logins on the admin list | Everything else: postings, shift switches, time off, sick calls, leave, the printed call-out roster, publishing, Kronos import, adjustments and voids |

- Phone numbers and email addresses are hidden from people not on the admin list.
- The file is always saved in the restricted view, and so is every copy the system makes (pre-change backups, practice copies, the updater's backup).
- Every sheet is protected in both views. Admins type only into the roster's Home Unit, Crew, Phone, Plant Seniority, EMAIL and unit code columns and the Holidays sheet's "Observed on"; totals, pay-period cells and every other sheet are written by the buttons. Sheets can't be added, deleted or unhidden by hand.

## 19. Safeguards

| Safeguard | What it does |
|---|---|
| Automatic save | The file saves after each entry, and again when it is closed. |
| Daily backup | A dated copy is saved the first time an admin opens the live file each day. 30 days are kept. A backup, the frozen year-end file or a practice copy never takes that day's slot. |
| Error log | A button or pop-up that fails part-way writes a line to the admin-only Error Log sheet (and a text file beside the roster) and shows "This didn't finish ... Error Log line N" instead of a code error. |
| Smoke test | On a practice copy (MakePracticeCopy), SmokeTest runs 18 steps through the roster's own functions (add a person, post, award, call out, book off, sick call, switch, crew change, unit move, planned move, rename, publish, no takers / out of department / reopen, a leave going and returning, the Kronos heading rule, a planted leave word stopping the publish, void, remove, integrity) and checks every total comes back. It never saves. |
| One parser, one rule | Every date box reads through one parser; every rule (call order, who is free, closing, charging, the trim, the switch stops) lives in one module function the pop-ups call, with known answers in SelfTest. |
| Checked changes | A change to the roster's layout is compared before and after. If the totals differ, saving is blocked. |
| SystemCheck | Confirms the installed code is the right release and the data is in order. |
| IntegrityCheck | Compares every roster total with the OT Log, and looks for typed numbers, missing formulas and duplicate IDs. It only reads. It runs each time an admin opens the file, and the result is on the roster's status line. |
| Year-end checks | The rollover runs IntegrityCheck first and stops on any problem. It checks the frozen copy and the new roster before it saves. |
| SelfTest | Runs known-answer cases against the rules. It writes nothing. |
