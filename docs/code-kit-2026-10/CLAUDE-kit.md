# Call Out Roster: CLAUDE.md

Context for every coding session on this repo. Read this first, then the brief for the update you're building (`briefs/`).

## What this is
An Excel desktop `.xlsm` (VBA) overtime call-out system for Syensqo's Niagara Falls phosphine plant.
- **Master file:** `V2 - Call Out Roster.xlsm`. Main tab `OT Roster` (setting `ROSTER_SHEET`; code also finds it by headers).
- **Owner / maintainer:** Graeme (shift supervisor, admin). 8-10 supervisors use it, one at a time.
- **Units:** Phosphine, Derivatives, Cypure, Material Handler (Fumigation, called "Gases"), QC Lab (QC Derivatives), QC Gases, Utilities, SBP.
- **Plant terms:** NCN = No Coverage Needed. S/S or SS = shift switch. PHA = hazard-analysis coverage. 12-hour shifts, days 7 AM-7 PM, nights 7 PM-7 AM, nights dated by the evening they start.

**Nothing in this repo can be run.** There is no Excel here. Graeme imports the code, runs Debug > Compile, and tests by hand on a copy. So:
- Write compile-safe VBA (`Option Explicit` everywhere) and review it as if no compiler will check it.
- Prefer small, verifiable changes over rewrites.
- Anything that changes data runs a before/after check and blocks the save on a mismatch.

## Repo layout
```
modules/        *.bas  (standard modules, imported via File > Import)
paste-in/       *.txt  (code for ThisWorkbook and each UserForm, pasted via Ctrl+A)
briefs/         per-update build briefs (vNN-*.md)
docs/           user docs (system manual, supervisor guide, admin/install, rollover one-pager)
dist/           Callout_Roster_vNN.zip, one per release
CHANGELOG.md
```
**Seeding:** the repo starts from `Callout_Roster_v20.zip` once it's verified (see `briefs/00-v20-verify-and-merge.md`). If `modules/` is empty, stop and ask for the zip. Never rebuild a module from memory or from these docs.

## Workbook map (as of v20)
### Modules (~17; confirm against the zip)
| Module | Holds |
|---|---|
| `modOvertimeEntry` | **All settings** at the top: `ADMIN_USERS`, `REPORT_FOLDER`, `PUBLISH_FOLDER`, `KRONOS_PATH`, `ROSTER_SHEET`, `KRONOS_STALE_DAYS` (7), `POSTING_USE_GRADE_CUTOFFS` (False). Call-out logic. |
| `modOTLog` | OT Log writes, Void Entry (`S-####`, `P-####`, `#<EntryID>`), Add Adjustment, Refresh Roster Formulas, RenameEmployee |
| `modLogLayout` | OT Log column layout, `FormatOTLog`, `UpgradeOTLogLayout` |
| `modMigration` | Original `RunMigration` (keep: pattern for verified migrations) |
| `modPostings` | Postings, awards, emails (CC constants, `POSTING_CLOSE_TIME`, `POSTING_CLOSE_DAYS` 14, `POSTING_OUTSIDE_LINES` 3, `POSTING_PDF_SUBFOLDER`), `ReopenPosting`, Outlook + mailto fallback |
| `modShiftCycle` | 28-day rotation (two 28-letter strings; day 1 = Sun Dec 28, 2025, B days / A nights), `ScheduledShift`, eligibility |
| `modShiftSwitch` | Shift switches, `SwitchPickList`, `PersonFlag` (returns a code, not message text) |
| `modCalendar` | Calendar + Needs Attention drawing, THEME colours, `CalGroups`, change tracking (CalMap), Import Old Cal, `CalendarCheck` |
| `modHolidays` | 11 plant holidays for any year + Holidays sheet "Observed on" |
| `modSystem` | `SystemCheck` (module versions, `...1` duplicates, re-paste checks) |
| `modUI` | Command bar, status line, `ApplyRosterLayout` |
| `ImportKronosData` | Kronos import, Hours History snapshot, unmatched-names list |
| `Sort_Hours_Dept`, `Call_Out_Roster`, `OT_Posting`, `LoginFullAccess` | Sort buttons, printed roster, legacy posting print, admin |

### Forms
`frmOvertimeEntry` and `frmPostingEntry` have designer layouts (don't touch the designs). `frmNewPosting`, `frmShiftSwitch`, `frmTimeOff` are **blank in the designer and built entirely in code**. New forms should follow the code-built pattern so install is "Insert UserForm, set (Name), paste code".

### Sheets and tables
| Sheet | Notes |
|---|---|
| `OT Roster` | Rows 1-4 command bar + navy date block (period start / "to" / end); A:S frozen. G = Total Hours, H = Worked Hours (Kronos), I = Refused, J = Plant Seniority, K = EMAIL, L:S = department code columns (grade codes like D1.5, P2, A/B). Pay-period columns after S are **formulas over the OT Log**. Hidden C (dup hours). Shift/crew column: (verify letter). |
| `OT Log` (`tblOTLog`) | **Append-only, newest first.** Event: Ref ID, Type, Shift Date, Unit, Shift, Crew, Reason, Hours. Person: Call Order, Employee, Dept Code, Total Hrs Then, Outcome, Note, Hours Charged, Period Start. Audit (grouped): Entry ID, Logged By, Timestamp, Status, Void Note. |
| `Postings` (`tblPostings`, `tblPostingList`) | Posting header + frozen list with hours at posting time |
| `Absences` (`tblAbsence`) | One row per booking; start/end (blank end = open-ended), hours per shift, times note. Admin only, hidden in saved copies. |
| `Coverage` (`tblCoverage`) | One row per date + shift line |
| `Shift Switches` (`tblShiftSwitch`) | SS-#### with payback links |
| `Calendar`, `Needs Attention` | Drawn, never typed into. Needs Attention admin-only. |
| `Holidays` | Admin; Observed on column for weekend holidays |
| `Hours History` (hidden) | Dated snapshot of worked hours per import (from v20) |
| `CalMap` (hidden) | Calendar change counters |
| `Overtime Posting` | Output only, redrawn per posting |
| `Legacy Entry Log` (hidden) | Pre-migration history |

### Ref ID prefixes
`OB-` opening balance, `S-` call-out session, `P-` posting, `SS-` shift switch, `#n` = Entry ID. Planned: `LV-` leave return, `UM-` unit move, `NH-` new hire, `SC-` sick call (v22), `YE-` year-end (if needed).

## Canonical business rules
These win over anything older in `docs/` or the project notes. Each rule must live in **one** function that forms, printouts, Needs Attention and the member publish all call.

**Call order:** grade code (in that unit's column) → Total Hours G (lowest first) → Plant Seniority J (earliest first) → name.

**Total Hours** = Kronos worked hours (H) + non-void Hours Charged in the OT Log for the current year (+ provisional accepted OT from v21).

**Who's listed (call-outs and postings):** anyone with a code in the unit's column, **any grade**, who is free for that date and shift. Not free when:
| Reason | Rule |
|---|---|
| Scheduled | Crew on that shift per rotation. A/B and C/D work their pair's **day** shifts whenever either is on days, never nights. DAYS = Mon-Fri days, **off on plant holidays** (observed date). |
| Rest rule | Day OT: crew coming off nights that morning. Night OT: crew starting days next morning. A switch worked the night before / day after counts. |
| Booked off | Absences row covering the date (not imported calendar lines, which are display only). |
| Switch | Working a switch that day, or gave up a shift that day. |

**Charging (only grade ≤ 2, × 1.75, never out-of-department people):**
| Event | Charged | Period |
|---|---|---|
| Call-out | Refused or Left Message | callout date |
| Posting: Submit Award | Ahead of the **last** awarded person in call order **and didn't sign up** | **shift date** |
| Posting: Award Out of Dept | Everyone eligible in dept who didn't sign up | shift date |
| Posting: No Takers | Everyone eligible on the list | shift date |
| Posting: Cancel | Nobody | - |
| DST nights | **Actual hours** (fall night = 13 h → 22.75; spring night = 11 h → 19.25) | - |

Postings compare using the **posted** totals (frozen at creation). Awarded is chosen in the Award pop-up, not "first sign-up".

**Year:** totals reset to zero each year; no carry-forward. Pay periods are biweekly, Sunday start; 2026 periods start Dec 28, 2025; 2027 starts Dec 27, 2026.

**Privacy:** STD, parental, LOA show only as "Leave" everywhere outside the Absences sheet. Never publish phone, email, leave type, leave notes, void notes or Logged By.

**People changes (v22/v24):** every rule-driven hours change is an **Adjustment** row in the OT Log with the calculation in Reason, shown in a confirm before saving, undoable by Void Entry.

## Engineering conventions
1. **Versioning.** Every module and paste-in carries a version constant that `SystemCheck` reads (verify the exact constant name in `modSystem`). Bump each changed file and SystemCheck's expected version in the same update.
2. **Append-only log.** Never edit or delete OT Log rows. Corrections are voids (struck through, Status VOID, Void Note).
3. **Verified data changes.** Any macro that restructures data: save a dated backup first → change → compare every roster period, Refused (I) and Total (G) value and the log's total Hours Charged before vs after → on mismatch, block saving and tell the user to close without saving. Copy the pattern in `RunMigration` / `UpgradeOTLogLayout`.
4. **Read every row, whatever the view.** Use `ListObject.DataBodyRange` or header lookups, never last-visible-row. Clear filters before measuring. Handle `DataBodyRange Is Nothing` on empty tables.
5. **Find columns by header text**, not fixed letters, for anything new.
6. **Bulk work:** turn off ScreenUpdating, EnableEvents and calculation, and restore them in an error handler on every exit path.
7. **Calendar dirty flag:** macro writes to Absences, Coverage, Postings or Shift Switches with events off must mark the calendar out of date explicitly (verify the helper name in `modCalendar`).
8. **Late binding only** (Outlook, Scripting, etc.). No new references.
9. **Dates:** build with `DateSerial`; parse typed dates through the existing date parser (handles "Oct 12" and year rollover). Never `CDate` a locale-ambiguous string.
10. **Admin gating** via `ADMIN_USERS`; admin buttons hidden for others.
11. **Flat, list-shaped tables** (one header row, stable names, no merged cells) so they can become SharePoint Lists later.
12. **Mac/popup-blocked fallback:** menus fall back to a numbered list (existing helper).

### VBA compile traps to check in review
- `Dim a, b As Long` makes `a` a Variant. Declare each.
- Max 24 line continuations per statement; 64 KB per procedure. Split big builders.
- Use `Long` for row counts, not `Integer`.
- Duplicate procedure names across modules → "Ambiguous name". Search before adding a Public name.
- Code-built forms: control names must be unique; `WithEvents` handlers need class wrappers (follow existing pattern).
- `Application.Match`/`VLookup` return errors, not raise them: test with `IsError`.

## How every update ships
1. Start from the current master zip (in `modules/` + `paste-in/`).
2. Build to the brief. If the brief marks something **(verify)** or **Question**, check the code first; if it's a real decision, ask Graeme with **multiple-choice questions** (he prefers the pop-up), offering the brief's default first.
3. **Independent review** of every changed file by a separate agent that didn't write it: compile traps above, rule consistency with this file, filter/visible-row bugs, events/calc restored, privacy.
4. Add **SelfTest** cases for every rule touched (from v21 on).
5. Package `dist/Callout_Roster_vNN.zip` with the **full** module and paste-in set (it becomes the master), plus `INSTALL-vNN.md`: Remove/Import list, re-paste list, new forms, Compile, macros to run once, SystemCheck expectation, smoke test.
6. Update `CHANGELOG.md` and the affected `docs/`.

**Install freezes:** no updates Dec 20 - Jan 4 except the rollover; none in the week before Jan 23, 2027.

## Decisions already made (don't re-ask)
- Postings list any grade; only grade ≤ 2 charged.
- Totals reset yearly, no carry-forward.
- Stat holidays: DAYS off and free; A/B, C/D stay scheduled. Weekend holidays use the plant's announced weekday (Holidays sheet).
- Provisional accepted OT counts **as soon as it's logged, even for a future shift**, until Kronos covers it. Kronos data runs **through the day before** the import date.
- DST nights charge **actual** hours.
- Partial sign-ups: a full-shift sign-up beats a partial one even with more hours; a partial is awarded only if nobody takes the whole shift; uncovered remainder drops off, uncharged.
- Leave return: + home unit's average Total Hours gained between first day off and return (grade ≤ 2, not on leave).
- Full unit move: up to the highest Total Hours among grade ≤ 2 in the new unit; never lowered. Extra sign-offs don't trigger it.
- New hire: home unit's average (grade ≤ 2, not on leave).
- Sick call-in: opens a line, starts a call-out; off call lists for that shift and the next 24 h.
- Member workbook: plant PCs, desktop Excel, everyone's hours visible, master pushes values, member file has no rules.
- Platform: stay on the Excel desktop master through Q1 2027; keep tables platform-ready.

## Still open (use the default unless Graeme answers)
| Question | Default |
|---|---|
| Boxing Day 2026 observed weekday | Data entry on Holidays sheet, no code |
| Sick calls on the posted calendar | "Sick" |
| Leave spanning the year reset | Top-up counts only from the reset to the return date |
| Members see each other's My history | Yes |
| Member folder path | Existing `PUBLISH_FOLDER` |
| Who runs the rollover / where the frozen copy lives | Settings constants, blank until given |
| Home unit for multi-unit people | Graeme fills the column once |
| Does Kronos export an employee number | Needed before v25 |

## Roadmap
| Brief | Target | Theme |
|---|---|---|
| `00-v20-verify-and-merge` | Now | Confirm v20 has the Oct 9 fixes; name the master |
| `01-M1-member-workbook` | Build Oct 12-30, live Nov 15 | Member file + Publish |
| `02-v21-call-order-accuracy` | Before Oct 31 | Provisional OT, DST, Home Unit, SelfTest, daily backup |
| `03-v22-sick-and-leave` | Mid-Nov | Sick Call, Leave going/returning, convert imported lines |
| `04-v23-year-end` | Build by Nov 30, live Dec 27 | RolloverYear, Integrity check, docs |
| `05-v24-employee-changes` | Before Jan 16 | Employee pop-up, crew history, partial OT |
| `06-v25-hardening` | Feb-Mar 2027 | Employee ID key, protection, calendar speed |
| `07-v26-reporting` | Mar 2027+ | History report, staffing check, board print, payback aging |
