# v24 report-back (before any code)

Brief: `docs/code-kit-2026-10/briefs/05-v24-employee-changes.md` (Employee pop-up, crew history, partial OT). Read against the code at tag `v23` on Oct 12, 2026. Nothing is built yet.

**Timing.** v24 depends on the rollover being done, and the install freeze runs Dec 20 - Jan 4. So it can be built now but installs **Jan 5 - 15, 2027**. By then the roster year starts Dec 27, 2026, which is the date the crew history is seeded from.

## What the code already gives us
| Brief item | What I found | Effect on the plan |
|---|---|---|
| `NH-` / `UM-` adjustments | `AdjustmentTypes` already lists "New hire" and "Unit move"; `NextRefID` takes any prefix and honours the ID high-water marks. | Nothing new for IDs. |
| Starting hours "via `TotalHoursAsOf`" | Exists (modLeave), with the batch cache, the peer test (`PeerEligible`, `HasLeaveOverlapping`) and `LeaveAverage` rounding. | New hire and unit move reuse them. Peers are found the way leave return finds them: everyone with a code in that unit's column, grade 2 and below. |
| "Roster row with formulas, no Refresh Formulas" | `WritePeriodRow`, `WriteRefusedRow`, `WriteTotalRow` write one row; `EnsureRosterFilter` regrows the filter when the roster grows. | Add = new row under the last one, formats copied from the row above, the three writers, then the filter. |
| Grade Changes row on a move | The grade log compares the roster to its name-keyed snapshot. | The move writes the codes and calls the same comparison; no second copy of the rule. |
| Crew on switches and posting lists | Shift switches already store each person's crew in the row; posting lists store it at posting time. | Past switches and postings are already safe from a crew change. |
| Partial hours on the calendar | `HrsTag` already draws "Joe D 8h" for hours under a full shift. | Reused, with one change (finding 9). |

## Findings that change the plan
1. **Rename is broken today, beyond the log.** `RenameEmployee` changes the roster, the OT Log and the grade snapshot only. The name also lives in: posting lists, the postings' "Awarded To" text, Absences, Shift Switches (two columns), Coverage, the Grade Log, Hours History, and (after v24) crew history. A rename today leaves those behind: the person's bookings stop blocking them on call lists, and `TotalHoursAsOf` loses their Hours History. v24's Rename fixes all of them in one verified pass (backup, counts per table before and after, every roster period / Refused / Total value unchanged, save blocked on a mismatch).
2. **Coverage holds short names as free text** ("Fred C", sometimes "Fred C / Sam K"). Two people can share a short name. Rename changes a coverage line only when the line is tied to that person (their absence, their log rows, their switch); any other line containing the old short name is listed in the confirm and left alone.
3. **Closed years.** `OT Log 2026` and `Year End` stay protected even in admin view, and the frozen FINAL file can't change. Decision 4 below.
4. **Crew is not only read in "direct reads".** There are eight read sites (LoadEligible, the calendar cache, the calendar fingerprint, two in modShiftSwitch, the out-of-department lookup, the publish, attendance), but the crew then travels as text inside each list record and is handed to `NotFreeReason` by about ten callers. Plan: `CrewOn(employee, date)` in a new module `modPeople`, and `NotFreeReason` / `EmpShiftOn` / the switch checks resolve the crew from name + date themselves, so the callers keep their signatures. `LoadEligible` gets an optional date so a posting for a future shift lists the crew each person will be on.
5. **Typing a crew on the roster.** A sort fires the same change event for the whole block, so the check must be by name. The crew history itself is the "was" value (no second snapshot): a crew cell that differs from the person's newest history row prompts for the effective date; Cancel puts the old crew back. A change made with macros off is picked up on the next admin open, dated that day. Column E always shows the newest crew.
6. **Dates before the seed.** `CrewOn` for a date earlier than a person's first row returns that first row's crew, so 2026 bookings draw exactly as they do today.
7. **Remove (decided Oct 12: move to a "Former" sheet, not a flag).** What that means in the code:
   - Remove becomes the one place a roster row is taken off the roster, so it follows the verified pattern: backup, record the person's Total / Worked / Refused / Provisional and their log total, copy the row to the `Former` sheet as values with a Left On date, delete the roster row, then check every remaining row's period, Refused and Total values and the log's total Hours Charged are unchanged. Save blocked on a mismatch.
   - `Former` (`tblFormer`, admin, hidden in saved copies, never published): name, Left On, crew, home unit, seniority, the eight codes, the four final figures, removed by, timestamp. Phone and email are not copied.
   - Their OT Log, absence, switch and posting rows stay. IntegrityCheck's "not on the roster" check treats names on `Former` as known, and the Kronos import stops listing them.
   - Every call list, picker and the member publish read the roster, so a former person drops off all of them with no other change. The filter re-apply I had planned is not needed.
   - It is not date-aware: once moved, the person is off lists for every date. Remove is run on or after their last day; until then, book them off.
   - **Bring back** moves the row back and rewrites its formulas, so the log hours reappear. Add refuses a name that is on `Former` and offers Bring back instead.
   - The rollover's Year End record reads roster rows only. Someone who leaves during 2027 would be missing from the 2027 finals, so `RolloverYear` also copies that year's `Former` rows into Year End. This changes modRollover after its first live run (Dec 27, 2026) and before its second.
8. **Column A stays empty.** No roster column is added.
9. **A partial award would over-count provisional hours.** Column C sums the log row's Hours for Awarded rows, and every posting row carries the posting's full hours today. An 8-hour award would add 12. Plan: the Awarded row carries the awarded hours; Refused rows keep the posting's hours.
10. **The award is chosen by hand.** The pop-up doesn't pick winners; the supervisor does, and `SubmitAward` warns. So "full first, partial only if no full" becomes a pure check (`AwardOrderProblem`) at submit, counted per position: with 2 positions and 1 full sign-up, a partial may take the second. Decision 5 below. Partial hours awarded can't exceed the posting's hours per position. The workbook tracks hours, not clock times, so it can't tell whether two partials overlap.
11. **Award pop-up.** Double-click already means "signed up". Plan: keep that, and add a code-built **PARTIAL HOURS** button (blank = full shift); the Signed column shows "Yes" or "8 h". The form's design is not touched.
12. **`tblPostingList` gets a 10th column** by a verified upgrade (the column right of the table must be empty or it stops). Old 9-column tables keep loading until the upgrade is run.
13. **Unit codes aren't dated.** Changing a code changes every list from that moment. So a unit move's effective date can be today or earlier (it dates the hours calculation and the adjustment); a move planned for later is entered on the day. A crew change can be future-dated.
14. **Which code is cleared on a move.** People can hold codes in several units. Plan: the form shows their codes, the Home Unit's code is ticked for clearing, others stay unless ticked.
15. **Open postings.** Someone who moves or leaves is still on the frozen list of open postings in the old unit and would be charged for not signing up. Decision 3 below.
16. **Kronos matches on the name.** Add and Rename say so: the name must be spelled the way Kronos spells it, or Worked hours stop updating for that person.
17. **Out of scope:** partial hours on call-outs and on out-of-department awards (both stay full shift); effective-dated unit codes.

## Build order
| # | Item | Notes |
|---|---|---|
| 1 | `modPeople`: crew history + `CrewOn` | `tblCrewHistory` on an admin sheet, seeded from the roster at the year start by a verified one-time macro (`SetupPeople`). Pure core takes an array so SelfTest plants rows in memory. All crew reads moved. Hand-edit prompt in ThisWorkbook. |
| 2 | `Former` sheet | `tblFormer`, the verified Remove and Bring back, IntegrityCheck / Kronos / rollover read it. |
| 3 | `frmEmployee` (code-built) | Add (`NH-`), Move to unit (`UM-`), Rename, Remove / Bring back. Each shows the calculation and asks before saving; hours changes are Adjustment rows, undone by Void Entry. The ADMIN "Rename Employee" button becomes "Employee"; the macro name `RenameEmployee` stays and opens the pop-up. |
| 4 | Partial OT | `SignupHours` column, PARTIAL HOURS button, award-order check, log hours, calendar text, award email. |
| 5 | Tests | Cases in `tests/cases.csv` + `rules.py` + SelfTest: new-hire average, move up, move with higher own hours, crew change keeps a past absence on its old shifts, partial vs full ordering, remainder uncharged, former person off a list from their date. |
| 6 | Reviews, zip, tag `v24` | Two independent read-only reviews, fixes, re-review. 18 modules. ThisWorkbook and `frmPostingEntry` re-pasted; one new form. Run once: `SetupPeople`, then `ApplyRosterLayout` (button rename). `Sort_Hours_Dept` changes too, although it carries no version stamp. |

## Decisions (Graeme, Oct 12, 2026)
| # | Question | Answer |
|---|---|---|
| 1 | Refusers when only a partial is awarded | **The posting's full hours x 1.75** (brief default). |
| 2 | Remove | **Move the row to a "Former" sheet** (not the brief's default flag). See finding 7. |
| 3 | Open postings when someone moves or leaves | **Never charged** on open postings for shifts on or after the date; the confirm lists the postings affected. |
| 4 | Rename and closed years | **Leave `OT Log 2026` / `Year End` under the old name**; the rename is noted on the Grade Log. |
| 5 | A partial marked awarded while a full-shift sign-up is free | **Blocked**, with a message to award the full sign-up or untick it if they withdrew. |
| 6 | Packaging | **One v24 install.** |

Waiting for a go-ahead before any code.
