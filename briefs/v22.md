# Brief v22: sick calls and leave

**Target:** mid-November (before or with wave 2, Nov 15).
**Depends on:** v20 Hours History, v21 Home Unit + Adjustment subtypes + SelfTest.

## Scope
| # | Item |
|---|---|
| 1 | `TotalHoursAsOf(employee, date)` foundation |
| 2 | **Leave** button: Going / Returning, department-average top-up (`LV-####`) |
| 3 | **Sick Call** button (`SC-####`) |
| 4 | **Convert imported lines** to real bookings + "Imported, not booked" list |

## 1. TotalHoursAsOf(employee, d)
Everything in this update that needs "hours on a past date" goes through this one function.
- **Worked:** the Hours History snapshot with the latest data-through date ≤ d (snapshot data-through = import date - 1). Return the snapshot date used, so confirms can show it.
- **Plus** non-void Hours Charged rows for that employee with event date ≤ d in the same year (event date = Shift Date; Period Start for OB rows).
- **Plus** provisional accepted hours as defined in v21 *as of d*.
- Clipped to the year: nothing before the year start counts (reset to zero).
- **No snapshot on or before d** (someone already on leave before Hours History started): use the earliest snapshot and flag "start figure estimated from <date>" in the confirm. Allow an admin override value with a required reason.
- Flag any snapshot more than 7 days from d.

## 2. Leave (`frmLeave`, code-built, two modes)
**Going on leave**
- Person, type (PAR, STD, LOA, WSIB, LTD, OTH-leave), first day off, expected return (optional = open-ended), note.
- Creates the `tblAbsence` row (all leave types show as "Leave" outside the Absences sheet). Adds columns to `tblAbsence` if missing (verified upgrade): `HomeUnit` (copied from roster now, locked), `LeaveGroup` (LV id linking going/return).
- Off call lists from the first day off; shows in Needs Attention "Leave".

**Returning**
- Pick from people with an open leave booking; enter return date.
- Top-up:
  1. Window = first day off → return date, clipped to the year start (**default for leave spanning the reset**).
  2. Peers = the locked HomeUnit's people with a code **grade ≤ 2**, excluding the returner, anyone with a leave-type absence overlapping the window, and anyone with no Hours History before the window start.
  3. For each peer: `TotalHoursAsOf(return) - TotalHoursAsOf(firstDayOff)`.
  4. Average, rounded to 0.1.
- Confirm shows unit, peer count and names, dates, snapshot dates used, average, and the returner's before → after: "Derivatives, 11 operators, Mar 3 to Sep 14: average 412.6 h. Add 412.6 h to Smith, J (310 → 722.6)?"
- On OK: OT Log Adjustment, Subtype **Leave return**, Ref `LV-####`, Hours Charged = average, period = return date's period, Reason = the calculation in one line. Closes the absence (End = return - 1). Back on call lists from the return date.
- Void Entry `LV-####` removes the top-up (absence stays closed; reopening is a separate action in the form).
- Needs Attention: new **Returning from leave** list (expected return within the window, or open-ended) with a "Return now..." menu.

## 3. Sick Call (`frmSickCall`, code-built; LOG group, admin)
1. Pick person and date; Shift and Unit fill from crew rotation and Home Unit (editable). Hours default from `DstNightHours`/12; editable for partial.
2. Saves `tblAbsence` type **SIC** with a `Shift` value (add column, verified). Ref `SC-####`.
3. **Off call lists** for that shift and the next 24 hours. Default implementation: block any candidate shift that **starts before the sick shift's end + 24 h** (sick days on D → also off nights D and days D+1). (Question below.)
4. Adds an OPEN coverage line for date/shift/unit and marks the calendar dirty.
5. Opens Log Call-Out pre-filled (date, shift, unit, crew to call). An accepted person fills the line as now.
6. A second sick call for the next shift extends it (new SC row, same person).
- Calendar shows **"Sick"** (default) in the off column.

## 4. Convert imported lines
- Source rows: `tblCoverage` lines brought in by Import Old Cal (verify the source marker column) with a person-off type of vacation, dated today or later.
- Name matching: imported text like "Paul M" → roster by first name + last initial; show Match / Ambiguous (pick) / No match.
- Review form: tick rows; consecutive dates for the same person merge into one `tblAbsence` VAC range. Each converted coverage line keeps its display but records the new absence ID (so it isn't converted twice).
- Needs Attention: **Imported, not booked** list (next 56 days) until converted or dismissed.

## SelfTest additions
TotalHoursAsOf (snapshot choice, year clip, void excluded); leave average with a fixed peer set incl. an excluded on-leave peer and a grade-3 peer; sick 24 h window for days and nights; convert merges a 5-day run into one booking.

## Acceptance
- On a copy: a test leave return shows the peers and average; Void Entry LV- removes it; totals verify.
- A sick call opens the line and the call-out; that person is absent from the next shift's list and back after 24 h.
- Converting a P&D import removes those people from call lists on those dates.

## Questions (multiple choice, default first)
- Sick exclusion: **any shift starting within 24 h after the sick shift ends** / the sick shift plus the next shift only.
- Calendar label: **"Sick"** / masked as "Off".
- Leave types: **add WSIB and LTD codes** / use OTH for them.
