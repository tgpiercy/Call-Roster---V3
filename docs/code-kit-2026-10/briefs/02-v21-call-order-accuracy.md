# Brief v21: call order accuracy

**Target:** installed before **Sat Oct 31** (first DST night). Shadow run Oct 18-31 runs on it.
**Depends on:** v20 master (Hours History, Kronos data date).

## Scope
| # | Item | Modules (likely) |
|---|---|---|
| 1 | Provisional accepted OT in Total Hours | `ImportKronosData`, `modOTLog`, `modOvertimeEntry`, roster formulas |
| 2 | DST night hours | `modShiftCycle`, `frmOvertimeEntry`, `frmNewPosting`, `modPostings` |
| 3 | Home Unit column | `modOTLog` / roster layout, `modUI` |
| 4 | Adjustment subtypes | `modOTLog`, `modLogLayout` |
| 5 | Grade change log | `ThisWorkbook`, new table |
| 6 | SelfTest v1 | new `modSelfTest` |
| 7 | Daily backup on first open | `ThisWorkbook`, `modSystem` |

## 1. Provisional accepted OT
**Rule (decided):** an Accepted call-out row or an Awarded posting row (in-dept or out-of-dept) adds its hours to the person's Total Hours **as soon as it's logged, even for a future shift**, until a Kronos import covers that shift; then it drops out automatically (Kronos worked hours now include it).

- **Data-through date:** Kronos data runs through the **day before** the import date. `ImportKronosData` stores `KronosDataThrough = importDate - 1` in a workbook name (if v20 already stores the import date, derive from it).
- **Covered when:** shift end date ≤ KronosDataThrough. Days on D end D; **nights dated D end D+1**. (Question below.)
- **Hours:** the row's event Hours (partial acceptances use their own hours). Void rows never count.
- **Implementation preference:** keep it formula-driven so it recalculates on import with no macro step: a `Provisional` column on the roster (find by header; place it beside G) = SUMIFS over `tblOTLog` (Outcome ∈ Accepted/Awarded, Status ≠ VOID, Employee = row, ShiftEnd > KronosDataThrough, current year). G = H + refusals + Provisional. Add a helper column `Shift End` to `tblOTLog` only if SUMIFS can't express it; if you add one, do it through a verified upgrade (CLAUDE.md convention 3).
- **Postings** freeze totals including provisional (they already freeze G).
- **Status line:** "Provisional: N people, X h".

**Tests:** accept a call-out for tomorrow → that person's G rises immediately and they drop in order; import Kronos with data through the shift → G unchanged overall (provisional moves into H); void the acceptance → G drops back.

## 2. DST nights
- `DstNightHours(shiftDate, shift)` → 13 when a **night** dated D has D+1 = first Sunday in November; 11 when D+1 = second Sunday in March; else 12. Any year. Constant `DST_ENABLED = True`.
- New Posting and Log Call-Out default Hours to this value and show a yellow note: "DST night: 13 h (clocks go back)".
- Refusal charge = **actual hours** × 1.75 (13 → 22.75, 11 → 19.25). Charging must use the posting's/session's Hours, not a hard-coded 12. Search for any literal `12` used as shift hours and replace with the function or the row's hours.
- Next dates: night of **Sat Oct 31, 2026** = 13 h; night of **Sat Mar 13, 2027** = 11 h.

## 3. Home Unit column
- New roster column **Home Unit** (header lookup). Data validation list = the 8 units.
- Auto-fill when a person has exactly one unit code; leave blank (yellow) where they have several. SystemCheck lists blanks.
- Inserting a column inside A:S shifts formulas: do it through a verified upgrade macro `AddHomeUnitColumn` (backup → insert → recompute → compare every G/I/period value → block save on mismatch). Alternative if cleaner: repurpose hidden column C only if nothing references it (verify first).
- Used by v22 (leave, sick) and v24 (moves, new hires).

## 4. Adjustment subtypes
- Add Adjustment gets a required **Subtype**: Leave return, Unit move, New hire, Correction, Other.
- Store in a new `Subtype` column appended at the **end** of the Event group or table (verified upgrade; existing columns don't move). Existing Adjustment rows: blank.
- Void Entry and Member publish ignore it; reports will use it.

## 5. Grade change log
- New hidden sheet **Grade Changes** (`tblGradeChange`: Employee, Unit, OldCode, NewCode, ChangedAt, ChangedBy).
- Detect by **comparing to a snapshot keyed by name**, not by cell address, so sorting never logs a change. Snapshot refreshed after each logged change and on open.
- Prompt nothing; just log. Renames (RenameEmployee) update the snapshot key.

## 6. SelfTest v1 (`modSelfTest`, macro `SelfTest`)
- Runs on a hidden scratch sheet; **never writes to real tables**. If charging/eligibility logic isn't callable as pure functions, extract pure helpers first **without behaviour change** (e.g. `ChargeDecision(rank, lastAwardedRank, signedUp, grade, inDept, outcome)`), and have the live code call them.
- Report: a message + a sheet listing each case PASS/FAIL with expected vs actual. SystemCheck shows the last SelfTest result and date.

| Group | Cases (minimum) |
|---|---|
| Rotation | Sun Dec 28, 2025 = B days / A nights; Mon Oct 5, 2026 = D days / C nights; + 18 more dates **Graeme supplies from SHIFT2026** (put them in a table at the top of the module, marked TODO if missing) |
| Call order | Grade beats hours; hours beat seniority; seniority beats name; A/B code listed for A and B |
| Charging | Awarded; ahead + no sign-up (charged); ahead + signed up (not charged, warning); below (not); grade 2.5/3 (never); no takers; out of dept; cancel; 2 positions uses last awarded; call-out Refused / Left Message / Accepted |
| Eligibility | Rest rule day + night; A/B and C/D days only; DAYS Mon-Fri; DAYS free on Thanksgiving Oct 12, 2026; booked off; working a switch; gave up a shift |
| Holidays | All 11 for 2026 (Thanksgiving Oct 12, Good Friday Apr 3, Victoria Day May 18, Civic Aug 3, Labour Day Sep 7, Family Day Feb 16, Boxing Day Sat Dec 26 flagged weekend) |
| DST | Oct 31, 2026 night = 13; Mar 13, 2027 night = 11; Nov 1, 2026 day = 12 |
| Provisional | Future accepted counts; covered drops; void drops; night shift uses D+1 |

## 7. Daily backup
- On the first open of the day by an admin (date stored in CalMap or a hidden name): `SaveCopyAs` to `BACKUP_FOLDER\Call Out Roster yyyy-mm-dd.xlsm`. Setting `BACKUP_FOLDER` (default: a `Backups` folder beside the workbook), `BACKUP_KEEP_DAYS = 30`.
- Prune only files matching that exact name pattern older than 30 days. Never touch anything else.
- Failure is a status-line warning, never a blocker.

## Acceptance
- Compile clean; SystemCheck "all on v21"; SelfTest all green (except TODO rotation rows).
- Verified upgrades report "totals unchanged".
- Smoke: accept a call-out → order changes now; New Posting for night of Oct 31 → 13 h default; change someone's grade → one Grade Changes row; sort the roster → no rows.

## Questions (multiple choice, default first)
- Night shift covered by Kronos when data reaches: **the morning it ends (D+1)** / the evening it starts (D).
- Home Unit placement: **verified insert near the unit columns** / repurpose hidden C.
