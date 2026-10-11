# Changelog

## v21 - call order accuracy (Oct 2026)
Brief: `briefs/v21.md`. Install notes: `INSTALL.md` in `dist/Callout_Roster_v21.zip`.

- **Provisional accepted OT.** An Accepted call-out or Awarded posting counts in Total Hours (G) from the moment it is logged, until Kronos covers the shift. The amount is the row's shift hours (not x 1.75), non-void rows in the roster year only. A day shift is covered once its date is on or before the Kronos data-through date; a night shift once the next day is. The import stores the data-through date (file date minus one) on every Hours History row and in the hidden name `KronosDataThrough`; column C on the roster (hidden) holds the provisional total per person and G = H + I + C. The call-out form marks such totals with `*`, the posting form says how many, the printed posting footer says so.
- **DST nights.** The night starting the Saturday before the first Sunday in November is 13 h; the Saturday before the second Sunday in March is 11 h (worked out by rule for any year). New Posting and Log Call-Out default Hours to the actual length on those nights when the shift is Nights and show a warning. Setting `DST_USE_ACTUAL_HOURS` (True).
- **Home Unit** column (B) on the roster with a drop-down of the eight units. `SetupHomeUnit` fills in anyone with a code in exactly one unit; SystemCheck and Needs Attention list the rest. Used by v22 and v24; nothing else reads it yet.
- **Adjustment types.** Add Adjustment asks for a type (Leave return / Unit move / New hire / Correction / Other), stored in the OT Log's Reason column.
- **Grade change log.** Changing a unit code on the roster writes a row to `tblGradeChange` (hidden admin sheet Grade Log; `ShowGradeLog`): timestamp, employee, unit, old code, new code, Windows login. Sorting, filtering and macro writes create no entries.
- **Daily backup.** The first admin open each day saves `Call Out Roster yyyy-mm-dd.xlsm` into `BACKUP_FOLDER` and deletes copies older than 30 days there. Never blocks opening; SystemCheck says when the folder isn't set or can't be reached.
- **SelfTest** (Alt+F8): 147 known-answer cases over rotation, holidays, DST, call order, charging, eligibility and provisional coverage, in memory only. The same table (`tests/cases.csv`) drives the Python mirror in `tests/`.
- **Printed call out roster.** The Call Out Roster button redraws the Blank Call Out Roster sheet from the data: title block with pay period and Kronos status, one section per unit in call order (#, Hours with `*` for provisional, Name [code], Crew, Phone), page setup done, then offers to print. It replaces the old macro that sorted the roster unit by unit and copied the chunks. Roster column A is hidden (empty); B shows Home Unit.
- **Log Call-Out form:** the rotation line is two lines high and wraps (it clipped once the DST note or the crew note was added); the Crew box fits "All not scheduled". Every form now carries a `@version` line and SystemCheck checks all five.
- **SystemCheck** expects v21 on 12 modules and also flags Total Hours without the provisional term, blank Home Units and the backup.
- **Roster writes in this release** (all disclosed): `RefreshRosterFormulas` rewrites C and G formulas (verified, save blocked on a mismatch); `SetupHomeUnit` writes column B and its B2 title (backup first, verified); `ApplyRosterLayout` writes the hidden C2 title "PROV" and, once Total includes the provisional term, rewrites C and G with the period formulas. No hours are ever typed into roster cells.
- Rules consolidated: the call-out charge is one function (`CalloutCharge`); the rotation part of eligibility is `RotationFreeReason`; out-of-department and cancel charging are `ComputeOutside` / `ComputeCancel`.

## v20
Baseline imported from the v20 master (first commit of this repo). See `docs/UPDATE-20.md`.
