# The Desk - one page (v27, Desk v2)

The Desk is the first tab you see when you open the roster. Three pills at the top say what today is, tiles say what is waiting, and the buttons sit in three bands by WHEN you reach for them. Every button opens the same pop-up the full roster uses; nothing is different underneath. (The layout is the approved mock-up `docs/desk-v2-mockup-2026-10.html`; before v27 the Desk had six bands in two columns.)

There are two Desks from the same screen. **A supervisor's Desk** (a login on the supervisor list, v25.5) has the jobs a supervisor may do: Someone called in sick, Fill a shift now, What needs cover, Post a shift, Award a posting, Update the members' file, Print the call-out roster, Employee history (v26), and the help row. **The admin Desk** (for a login on the Desk list) has those plus Book time off, Shift switch or payback, Leave, Undo my last log entry, the paybacks-owed tile and "Show everything".

## The pills (the header)
| Pill | Says |
|---|---|
| Mon Oct 19 - B on days, D on nights | Today and the crews on each shift |
| Pay period Oct 11 - Oct 24 | The current pay period |
| Kronos data through Fri Oct 9 | How current the worked hours are. **Red** when they are more than 7 days old: "Kronos data 10 days old - import it" (an admin) or "- ask Graeme to import it" (a supervisor) |
| NOT SAVED since 3:10 AM - File > Save now | Only when a save failed (v26.5.5); red |

## The tiles (TODAY)
Each tile is the count, the next thing, and what a click does. A grey tile means nothing to do; a coloured one has work in it.
| Tile | Means | Click it for |
|---|---|---|
| N shifts need cover, next Tue Oct 20 Nights, Phosphine | OPEN lines on the calendar in the next 14 days, every unit | "Open Needs Attention" (admin Desk) or "Open the Calendar" (supervisor Desk) |
| N postings ready to award, closed Sun Oct 18 8:00 AM | Postings whose closing time has passed and that nobody has awarded; the date is the one that has waited longest | "Award it" |
| N paybacks owed, 1 over 42 days (admin Desk) | Shift switches still waiting for their payback; the overdue count since v26 | "See the OWED list" |
| Members' file: updated today 7:12 AM | When the members' file was last published; navy when it is a day or more old or never published | "Update it" (both Desks since v25.5) |

## The bands
- **NOW** (orange): Someone called in sick (the Sick Call pop-up: book the sick shift and, if you want, start the call-out from it). Fill a shift now (the Log Call-Out pop-up: pick the unit, date and shift, load the list, call down it). What needs cover (the Calendar).
- **PLAN** (green): Post a shift (New Posting). Award a posting. Admin: Book time off. Shift switch or payback. Leave: going or returning.
- **FINISH** (navy): Update the members' file (Member View). Print the call-out roster. Employee history (report) (v26: one person, a date range, the History Report tab and a PDF; a supervisor gets the member-safe report, an admin may pick the full one). Admin: Reconcile with the paper roster (v27: the paper call order typed beside the roster's, with the reason per differing position, for the parallel run). Stats (quarterly report) (v27: unit by quarter plus the year to date, four groups, a PDF on request).
- **The grey row at the foot**: How do I ... (the Help tab inside the file: one heading and a paragraph per job; < Back returns here). Something went wrong (an email to Graeme with the newest Error Log line already in it - add what you were doing and send). Suggest an improvement (the same kind of email). Undo my last log entry (admin; v25.5: voids your newest OT Log entry of today through Void Entry's confirm and reason). Practice on a copy (v26.5.5: a copy with a red PRACTICE banner; nothing done in it counts). Under the row: Full roster >, Show everything > (admin), and the keys Ctrl+Shift+S sick call, Ctrl+Shift+C call-out, Ctrl+Shift+K calendar, Ctrl+Shift+D this Desk.

## The tabs
You see the **Desk** and the tabs every supervisor sees (OT Roster, Calendar, Postings, OT Log, Shift Switches ...); the admin Desk adds **Needs Attention**, and opens the roster and the Calendar at 120 % so they read easily. The admin-only tabs (Absences, Coverage, Holidays, the logs) are in the file but out of the way.

## If you need the full roster tools (admin Desk)
Click **Show everything (primary admin tools)** under the help row. The full bar and every tab appear for this session. **< Desk** on the bar brings the Desk back. Nothing is saved differently either way. On a supervisor's Desk, **Full roster >** goes to the roster and the Desk tab brings you back.

## If a red "This didn't finish" box appears
Click OK, then **Something went wrong** on the Desk: the email it opens has the details. Whatever was done before the error is still in the file.
