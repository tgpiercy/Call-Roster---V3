# Brief v24: employee changes, crew history, partial OT

**Target:** live before **Jan 16, 2027** (no installs the week before Jan 23).
**Depends on:** v22 `TotalHoursAsOf`, v21 Home Unit + Grade Changes, v23 rollover done.

## Scope
| # | Item |
|---|---|
| 1 | **Employee** pop-up (`frmEmployee`, code-built, admin): Add, Move to unit, Rename, Remove |
| 2 | Crew history (effective-dated crew changes) |
| 3 | Partial OT on postings |

## 1. Employee pop-up
**Add employee** (`NH-####`)
- Name, phone, email, crew, seniority date, home unit, grade code, start date.
- Starting hours = home unit average Total Hours (grade ≤ 2, not on leave on the start date) via `TotalHoursAsOf`. Confirm shows peers and average.
- Adds the roster row with formulas and layout (no Refresh Formulas / ApplyRosterLayout needed), logs Adjustment Subtype **New hire**.

**Move to unit** (`UM-####`)
- Person, new unit, grade code, effective date, "crew changes too?" (feeds crew history).
- Clears the old unit code, sets the new one, updates Home Unit, logs a Grade Changes row.
- Target = highest Total Hours among grade ≤ 2 in the new unit (excluding the mover) on the effective date. If above the mover's total, log the difference as Adjustment Subtype **Unit move**; otherwise nothing (never lowered). Only for a full move, not an extra sign-off (the form says so).

**Rename** - replaces `RenameEmployee`; updates roster, log, postings lists, absences, switches, coverage text, Grade Changes snapshot, Hours History, together. Verified (counts per table before/after).

**Remove** (retired / left) - sets a `Status = Inactive` column (verified add) and an end date. Off every list, hidden from the roster view and member publish; log kept. No row deletion.

## 2. Crew history
- New table `tblCrewHistory` (Employee, Crew, EffectiveFrom). Seed one row per person from the current roster with EffectiveFrom = year start.
- `CrewOn(employee, date)` replaces every direct read of the roster crew column in rotation, eligibility, absences mapping, switches and calendar. Search for all reads.
- Changing crew on the roster by hand prompts for an effective date (default today) and appends a row; past bookings keep their shifts.

## 3. Partial OT on postings
- `tblPostingList` gets `SignupHours` (blank = full). Award pop-up: double-click a sign-up → Full / Partial (hours).
- **Award order (decided):** full-shift sign-ups first, in call order; a partial is awarded only if no full-shift sign-up exists. Multiple partials may be awarded in call order until the hours are covered; any uncovered remainder drops off, uncharged.
- Charging: people ahead of the last awarded person who **didn't sign up at all** are charged as today. Default: charge the **posting's full hours** × 1.75 (they refused the whole shift). People who signed up partial and weren't awarded: not charged.
- Calendar line shows "Joe D 8h" for partial awards (existing format).

## SelfTest additions
New-hire average; move up / move with higher own hours (no change); crew change mid-year keeps a past absence on its old shifts; partial vs full award ordering; partial remainder uncharged.

## Questions (multiple choice, default first)
- Charge for refusers when only a partial is awarded: **full posting hours** / only the hours that were awarded.
- Remove: **inactive flag, hidden** / move to a separate "Former" sheet.
