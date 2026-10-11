# M1 report-back (before any code)

Date: 2026-10-12. Baseline: tag `v21`.

## Answers to the five questions

1. **Today's publish.** `PublishMemberView` (modOTLog.bas:427-595) writes a values-only `.xlsx`: roster columns D, E, G, H, I, J, L:S plus the pay-period cells, and the OT Log with every column except Void Note (and Note unless `PUBLISH_REASONS`). **It still publishes "Logged By" and "Timestamp"**, which CLAUDE.md's privacy rule says members must not see. The Member View button's OnAction is `PublishMemberView` (modUI.bas:298), so that name stays and becomes a wrapper around `PublishMembers`.
2. **Calendar lines.** `SlotsFor(d, shift, groupName, slots)` (modCalendar.bas:569-627) returns `CalSlot` records: key, left text, right text, type code, open flag. The left text already masks leave (`TypeTag` turns LOA/STD/PAR into "Leave"); the times note ("7-3") and hours tag are included and are fine. **The type code is not masked** (it is STD/PAR/LOA), so the export maps those three to one public code before publishing. `groupName` may be a single unit name (`UnitInGroup` falls back to a unit match), so dCal can be built per unit.
3. **Who is free.** `LoadEligible(unit, "ALL", recs)` then `FilterNotFree(recs, n, d, shift, removed)`. `FilterNotFree` calls `LoadCache` every time (reads Absences, Coverage and the roster), and `SwitchWorkingOn`/`SwitchOffOn` read the whole switch table from the sheet on every call per person. For 8 units x 14 days x 2 shifts that is ~224 cache loads and tens of thousands of switch-table reads: slow (likely a minute or more). Proposal: a publish-time cache, `BeginSwitchBatch`/`EndSwitchBatch` in modShiftSwitch (one read, reused) and a `useCache` flag on `FilterNotFree` as `SlotsFor` already has. Same results, one read each.
4. **Postings.** `tblPostings` (PS_ID, status, posted, closes, date, unit, shift, crew, hours, positions, reason, awarded to/on/by, posted by) and `tblPostingList` (PL_ID, rank, employee, shift, code, hours posted, signed up, result, charged). `LoadPosting`/`OpenPostingChoices` already give open postings and their lists; export uses them.
5. **Needs Attention.** Sections are gathered into `NARow` arrays with `NAAdd` and drawn with `NASection`/`NARows` (modCalendar.bas:2594-2960); keys route double-clicks. A "MEMBERS FILE" section with key `M:` fits the existing pattern; double-click runs Member View.

## Changes to the plan
- The existing `.xlsx` publish leaks who logged each entry; M1 replaces it and the `.xlsx` is no longer written.
- Published calendar type codes are masked (LOA/STD/PAR -> "LV") so colours still work without revealing the leave type.
- Publish needs the batch cache above to run in seconds rather than minutes.
- Version stamp: master modules changed for M1 carry "v21 M1 (Oct ..)" so SystemCheck's v21 test still passes; `modPublish` joins the list (13 modules).

## Member file layouts (proposed)
- **Home:** title, "Updated <when> by <who> / Kronos data <date>", six big buttons, a Find me box (name picker) that opens Roster on that person.
- **Roster:** one table: Name, Crew, Home Unit, Total*, Worked, Refused, then per unit two narrow columns (Code, #). Buttons: each unit (sort by its # and hide blanks), A-Z, Show all, crew filter cells. Find me highlights the row. Without macros: table with AutoFilter, A-Z.
- **Who's next:** three picker cells (unit, date, Days/Nights) with data validation; a Show button fills the list from dNext in rank order. (With macros on, changing a picker refreshes automatically via Worksheet_Change.)
- **Calendar:** month and unit pickers; the master's grid shape (days row / nights row per date) drawn from dCal with the same colours; read-only.
- **Postings:** each open posting as a block: header line, then its list.
- **My history:** name picker; rows from dLog newest first; totals per pay period and for the year.
