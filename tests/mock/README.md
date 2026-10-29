# tests/mock - the alpha's mock plant (P9, Nov 1-14, 2026, on v29)

Every file here is made up: the names come from two word lists, phones are `905-555-01nn`, emails `first.last@example.test`,
Payroll #s `9` + 7 digits. Nothing in this folder is or resembles a plant record. The set is deterministic: `make_plant.py`
writes the same CSVs every run; `expected.py` the same `expected.csv` and `docs/alpha-test-plan.md`; `gen_mock_loader.py`
the same data block in `modMockData.bas`; `tests/test_mock.py` fails on any drift between them.

## How the pieces fit
```
make_plant.py  ->  people.csv, crew_history.csv, log_opening.csv, absences.csv, coverage.csv, switches.csv,
                   postings.csv, posting_list.csv, planned_moves.csv, restrictions.csv, hours_history.csv,
                   kronos_totals.csv, kronos_days.csv, idmarks.csv, expected_load.csv, roles.csv
plant.py       =   the engine: the roster's state and its readers, every rule from tests/rules.py (used by both scripts)
steps.csv      =   the plan's steps, hand-written (id, date, when, who, where, action, inputs, check, rules)
expected.py    ->  expected.csv (step id, what, figure) + docs/alpha-test-plan.md (the steps with EXPECT filled)
gen_mock_loader.py -> the MOCK DATA block of modMockData.bas (the CSVs as VBA constants: the loader needs no files)
make_kronos.py ->  kronos/*.xlsx (gitignored): the 17 mock Kronos report files, in the layout the import reads
modMockData.bas =  TEST-ONLY VBA: LoadMockData / ResetRosterData / MockSignupFile / MockRequestFile(FromBox)
```
Rebuild order after a change: `make_plant.py`, then `expected.py`, then `gen_mock_loader.py`, then `make_kronos.py`;
then `py -3.14 -m pytest tests/test_mock.py -q`.

## The plant
69 people over the 8 units: four rotating crews of 14 (Phosphine 4, Derivatives 5, Cypure 2, Material Handler 2,
QC Derivatives 1) on the 28-day rotation (day 1 = Sun Dec 28, 2025) and 13 on DAYS (QC Gases 3, Utilities 3, SBP 3,
QC Derivatives 2, Phosphine 1, Derivatives 1). Every unit has grades 1-3 (grade 3 listed, never charged). Two
cross-coded people (Marlow Yardley: Cypure G1 + Material Handler M2; Teal Copperfield: Derivatives D1 + QC Derivatives
Q2). An exact Total Hours tie broken by seniority (Merritt Vauxhall / Wren Fenwick at 194.5 on the D1.5 list, once Wren's Oct 30 night is in Kronos), one by
name (Devon Stonebridge / Finley Fairbanks at 203, the same seniority date, on P-0003's list) and one more by seniority
that the plan reaches on Day 12 (Arden Glassford / Jordan Hollins at 310.5 after the unit move). The frozen posting
lists and the historic rows' "Total Hrs Then" carry the Oct 30 figures (the mock has no earlier roster state): the
call orders are the same, the figures a little higher than a real Oct 20 list would have shown. `roles.csv` names the people the plan's steps lean on.

Opening Total Hours per person is designed (`TARGETS` in make_plant.py) and reached through the `OB-` opening-balance
row: `G = H + OB + the historic charges`. H is the weighted Kronos year to date (`kronos_totals.csv`, the multipliers of
`PREMIUM_CODES`); the loader writes the Oct 18 snapshot's lower figure and the Day-0 import brings H to the Oct 30 one.

## The columns
Every table CSV carries the ROSTER TABLE's own headers, in the table's column order, so the loader writes them straight
through (`WriteTable` with a type per column: text / date / stamp / number / Reason-with-apostrophe).

| file | columns | notes |
|---|---|---|
| people.csv | Payroll #, Home Unit, Name, Crew, Phone, Worked H, Seniority, Email, Phosphine, Derivatives, Cypure, Material Handler, QC Gases, QC Derivatives, Utilities, SBP | one roster row; the unit columns hold the grade code or blank; Worked H = the Oct 30 figure (after the Day-0 import) |
| crew_history.csv | Employee, Crew, Effective From, Logged By, Timestamp | one dated row per person (2025-12-28); Eden Larkspur has a second row: crew A until Oct 19, C from Oct 20 |
| log_opening.csv | the OT Log's 22 columns (Ref ID ... Payroll #) | oldest first, Entry ID 1..n: the 69 `OB-` rows, then the historic `S-0001..S-0006` sessions and `P-0001` (awarded) / `P-0002` (no takers) of periods 21-22 |
| absences.csv | Absence ID, Employee, Unit, Type, Start, End, Hours / Shift, Times, Note, Status, Logged By, Timestamp, Void Note, Home Unit, Leave Group, Shift | SC-0001/2 = two past sick occurrences of Blair Bellweather; AB-0001 vacation Nov 2-5; AB-0002 STD leave since Oct 5 (open-ended); AB-0003 personal day Nov 10; AB-0004 a closed September vacation |
| coverage.csv | Line ID, Date, Shift, Unit, Absence ID, Off / Need, Type, Covered By, Hours, Source, Status, Logged By, Timestamp, Note | the calendar lines: the vacation's two open lines, the personal day, the switch, one line per historic Accepted / Awarded row (Source = its ref) |
| switches.csv | SS ID, Shift Date, Shift, Unit, Gives Up, Off Crew, Works, Works Crew, Payback For, Paid Back By, Note, Status, Logged By, Timestamp, Void Note | SS-0001: Sep 12 (50 days by Nov 1 - over PAYBACK_WARN_DAYS 42), payback still owed |
| postings.csv | Posting ID, Status, Posted On, Closes On, Shift Date, Unit, Shift, Crew, Hours, Positions, Reason, Awarded To, Awarded On, Awarded By, Posted By | P-0001/2 closed history; P-0003 OPEN Tue Nov 3 Derivatives Days (closes Nov 2; two sign-up files); P-0004 OPEN Thu Nov 5 Material Handler Days (closes Sun Nov 1, no sign-up: Ray's exception); P-0005 OPEN Fri Nov 20 Phosphine Days (period 24; closes Nov 6) |
| posting_list.csv | Posting ID, Rank, Employee, Shift, Dept Code, Total Hrs (posted), Signed Up, Result, Hours Charged, Signup Hours | the frozen lists (call order at posting) |
| planned_moves.csv | ID, Employee, New Unit, Grade Code, Codes To Clear, New Crew, Old Crew, Effective From, Status, Entered By, Entered On, Closed On, Note | PM-0001 Blake Ingleby Cypure -> Material Handler M2 on Sun Nov 8 |
| restrictions.csv | Restriction ID, Employee, From, To, Reason, Status, Entered By, Entered At, Void Note | RS-0001 Ellis Everhart Nov 9-13 |
| hours_history.csv | Import ID, Imported At, Kronos File Date, Employee, Worked Hours, Imported By, Data Through, Year | the two snapshots before Day 0: KR-0021 (through Oct 3) and KR-0022 (through Oct 17) |
| kronos_totals.csv | Payroll #, Employee, Pay Code, Hours | the Day-0 cumulative's Totals block (Jan 11 - Oct 30); Bryn Hartwell carries LIEU TKN 8 h (not a roster code: named, never counted) |
| kronos_days.csv | Report, Date, Payroll #, Employee, Pay Code, Hours | the one-day reports: Report `day` = the morning after; `re` = the Nov 3 day re-pulled on Nov 6 (Harper Greenlaw 12 -> 10 OVERTIME); `cum` = Nov 14's rows, in the close-out cumulative only |
| idmarks.csv | Prefix, Number | the IdMark_ hidden names the loader sets (only the log-based prefixes read them) |
| expected_load.csv | Employee, G Total, I Refused, Period cells sum, H Worked, C Provisional | what the loader verifies (G / I / the period sum per person through RosterTotalsMap, and the log's charged total) before it saves |
| roles.csv | Role, Employee | who plays what |
| steps.csv | id, date, when, who, where, action, inputs, check, rules | the plan; inputs = `key=value;...` (people in an outcome list separated by `/`); rules = the coverage-matrix tags |
| expected.csv | Step, What, Figure | written by expected.py |

## What is deliberately NOT here
The real roster's Min Staff, Holidays, Rules and Settings sheets (the loader never touches them). Real people, real
hours, real logins beyond the five the plan names (piercyg, lanthierr, mihaychukc, sheltond, campeauc - already in the
code). The members' sign-up / request files for the scripted member steps: the testers make them on the members' file;
`MockSignupFile` / `MockRequestFromBox` are the fallback.
