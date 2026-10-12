# Brief v25: hardening

**Target:** February - March 2027.
**Depends on:** v24 (structure stable). Needs from Graeme: does Kronos export an employee number?

## Scope
| # | Item |
|---|---|
| 1 | Employee ID as the key in every table |
| 2 | Protection: macros-only writes, very hidden sheets, VBA password |
| 3 | Calendar speed |

## 1. Employee ID
- Roster column **Emp ID** (Kronos number if exported; otherwise generated `E-0001` and Kronos matched by name once, then by ID).
- Add `EmpID` to `tblOTLog`, `tblPostingList`, `tblAbsence`, `tblCoverage` (where a person is named), `tblShiftSwitch` (both people), `tblCrewHistory`, `tblGradeChange`, Hours History.
- Migration macro `MigrateEmpIDs`: backup → fill by exact name match → list anything unmatched and stop for a manual map table → verify per-person totals and row counts per table unchanged → block save on mismatch.
- All lookups switch to EmpID; names remain for display. Rename then only touches the roster name and display text.
- Tables stay list-shaped (future SharePoint Lists).

## 2. Protection
- Protect the roster, OT Log and the person tables with `UserInterfaceOnly:=True`, re-applied in `Workbook_Open` (Excel drops it on close). Every macro write must still work; search for any write that unprotects/reprotects and remove the need.
- `Absences`, `Needs Attention`, `Hours History`, `Grade Changes`, `CalMap`, archive sheets → `xlVeryHidden`; admin views unhide temporarily.
- VBA project password (Graeme sets it; never stored in the repo).
- Make clear in the admin guide: this keeps honest people honest. The admin check is a Windows login and is easy to spoof; the master must live in a supervisor-only folder.
- After this, Refresh Formulas becomes a recovery tool only.

## 3. Calendar speed
1. Build the month in arrays and write values in one pass.
2. Group cells by style and format each group once (Union ranges or a pre-formatted day template copied per day).
3. Then partial redraw: when a change touches one date/unit, redraw only that day's block (CalMap tracks which dates changed).
- Measure before/after (status line shows draw time) and record it in CHANGELOG.

## Acceptance
- Integrity check OK before and after the ID migration; SelfTest green.
- Typing into a period cell is refused by Excel; every button still works.
- Full month redraw at least 3× faster; single-line change redraws in under a second.
