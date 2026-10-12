# Brief M1: member workbook + Publish

**Target:** build Oct 12-30, test on plant PCs by Oct 30, live Nov 15 with wave 2.
**Depends on:** brief 00 (v20 master).
**Source design:** project doc `member-workbook-design.md`.

## Goal
A macro workbook for union members (`Call Out Roster - Members.xlsm`) that shows everyone's hours, call order, who's free next, the calendar, open postings and personal history. The master **pushes values** into it; the member file **never reads the master** and contains **no rules**.

## Deliverables
| File | What |
|---|---|
| `modules/modPublish.bas` (master) | `PublishMembers`, called by the existing **Member View** button (replaces the old `.xlsx` publish) |
| `member/modMember.bas` | `BuildMemberTemplate` + all member-side navigation |
| `member/ThisWorkbook.txt` | Member file open handler (go to Home, protect sheets, show stamp) |
| `modules/modCalendar.bas` | Needs Attention line "Member file last published ... (N days ago)", red when > 1 day |
| `modules/modOvertimeEntry.bas` | Settings: `MEMBER_TEMPLATE_PATH`, `MEMBER_FILE_NAME`, `PUBLISH_REASONS` (False). Reuse `PUBLISH_FOLDER`. |
| `modules/modUI.bas` | Button wiring only |

## Data contract (very hidden sheets in the member file)
All values, one header row each, list-shaped. Column names are the contract; don't rename later without bumping `dMeta.SchemaVersion`.

| Sheet | Columns |
|---|---|
| `dMeta` | SchemaVersion, PublishedAt, PublishedBy (display name only, no login), MasterVersion, KronosDataThrough, CurrentPeriodStart, CurrentPeriodEnd |
| `dRoster` | Employee, Crew, TotalHours, WorkedHours, RefusedHours, then per unit: `<Unit>Code`, `<Unit>Pos` (call position, blank if no code) |
| `dNext` | Unit, Date, Shift (Days/Nights), Pos, Employee, TotalHours, Code. 14 days from today × 2 shifts × 8 units. Uses the master's **own** eligibility + call-order functions. |
| `dCal` | Date, Shift, Unit, Line, OffOrNeed (leave as "Leave"), Coverage, Type. Last month → +3 months. |
| `dPost` | PostingID, Unit, ShiftDate, Shift, Hours, Positions, Reason, Closes, Status (open only) |
| `dPostList` | PostingID, Rank, Employee, PostedHours |
| `dLog` | ShiftDate, RefID, Type, Unit, Shift, Employee, Outcome, HoursCharged, PeriodStart. Non-void rows, current year. Note column only if `PUBLISH_REASONS`. |

**Never published:** phone, email, leave type, absence notes, void rows/notes, Logged By, Timestamp, Entry ID.

## PublishMembers
1. Admin-only. Build every data set into arrays from the master (no sheet-by-sheet copying of formatted ranges).
2. Open the template read-only, clear the `d*` sheets, write arrays in one `Range.Value` per sheet, stamp `dMeta`.
3. `SaveAs` to `PUBLISH_FOLDER\MEMBER_FILE_NAME` (xlOpenXMLWorkbookMacroEnabled), close.
4. If the save fails (file open on a plant PC): save `... - Members <yyyy-mm-dd hhmm>.xlsm` beside it, tell the user.
5. Record last-published time in the master (hidden name or CalMap) for Needs Attention.
6. **Privacy self-check before saving:** scan every published cell for anything matching a roster phone/email, or an absence type code other than "Leave"; abort with a message if found.
7. Never changes master data. Restore events/calc on every exit path.

## Member file sheets (built by BuildMemberTemplate)
| Sheet | Behaviour |
|---|---|
| Home | "Updated <PublishedAt> / Kronos data through <date>". Big buttons per view. **Find me** box. |
| Roster | Everyone with hours and each unit's call position. Buttons: one per unit (sort by that unit's Pos, hide people without the code), A-Z, Show all, crew filter. Find me highlights a row. |
| Who's next | Pick unit, date (next 14 days), Days/Nights → list from `dNext`, top first. |
| Calendar | Month + unit pickers; read-only grid from `dCal`, same colours as master THEME (copy the constants). |
| Postings | Open postings with closing date and each list. Note: sign-up is still on paper. |
| My history | Pick a name (default: everyone selectable) → rows from `dLog` + totals per pay period. |

- All sheets protected (`UserInterfaceOnly:=True`, re-applied on open).
- **Without Enable Content** the file must still open on Roster, sorted A-Z, with AutoFilter usable. So the template is saved with Roster active and a static A-Z copy of `dRoster` visible.
- No code in the member file computes eligibility, order or charges. It only sorts, filters and looks up.

## Acceptance tests
1. Publish with the master open; member file opens on a PC **with the master closed** and no access to the master folder.
2. Who's next for P&D, tomorrow, Days matches Log Call-Out's list and order exactly.
3. Roster sorted by Derivatives matches the master's Derivatives sort.
4. Search the published file for every roster phone number and email: zero hits. Calendar shows "Leave", never STD/PAR/LOA.
5. Publish while the member file is open elsewhere: dated copy saved, message shown.
6. Needs Attention shows the published line and turns red after a day.
7. Opened with macros disabled: Roster visible, A-Z, filters work.

## Questions (ask with multiple choice; default first)
- Member folder: **existing `PUBLISH_FOLDER`** / a new path.
- My history: **anyone can pick any name** / only the Windows user's own name.
