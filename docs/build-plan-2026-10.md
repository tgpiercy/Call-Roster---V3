# Call Out Roster: build plan (Oct 2026 - Q1 2027)

(Supplied by Graeme on 2026-10-12. Item numbers refer to deep-dive-review-2026-10-09.md, which has not been supplied.)

Turns the queue in `deep-dive-review-2026-10-09.md` into seven updates, plus the member workbook (moved up on Oct 9, see `member-workbook-design-2026-10-09.md`). Each update is small enough to install in one sitting and test the same day.

## How every update ships
1. Start of session: attach the current module zip (or the .xlsm). Nothing is built against guesses about the code.
2. Build, then an independent review pass on every changed module before it's packaged.
3. Package: zip of changed .bas / paste-in files plus a short install note.
4. Install on a copy: backup, Remove + Import, Debug > Compile.
5. Check: SystemCheck (versions, no `...1` duplicates), then SelfTest once it exists (from v21). Then the update's smoke test.
6. Go live, and update the project docs.

Install freezes: no updates Dec 20 - Jan 4 apart from the rollover, and none in the week before Jan 23.

## v20: Safety and stats - built Oct 9 (see UPDATE-20.md)
SystemCheck (3.3), holidays (1.2, 11 rules, Holidays sheet with Observed on), Shift Switch form fixes, Kronos age (1.4), Hours History (6.1), read-only notice + AutoSave off (1.5), Outlook fallback (3.5), verifies (1.6).

## M1: Member workbook (moved up from v26)
Target: build once v20 is verified (from Oct 12); test by Oct 30; live Nov 15 with wave 2. Design: member-workbook-design-2026-10-09.md.

| Part | What | Where |
|---|---|---|
| Template | `BuildMemberTemplate` makes `Call Out Roster - Members.xlsm` with Home, Roster, Who's next, Calendar, Postings and My history, plus very hidden data sheets. | new `modMember` + member `ThisWorkbook` |
| Publish | The Member View button now runs `PublishMembers`, any time: values only, pushed from the master; the members' file never reads the master. Replaces the old Member View publish. | new `modPublish`, `modUI` |
| Reminder | Needs Attention shows "Members file last published ..." when it's stale. | `modCalendar` |

Need: the members' folder path; whether My history shows everyone or only the viewer's own name.
Smoke test: publish, then open the members' file on a plant PC with the master closed; Who's next matches a call-out list; no phones, emails, leave types or void notes anywhere in it.

## v21: Call order accuracy - built Oct 11-12 (see briefs/v21.md, CHANGELOG)
1.3 provisional accepted OT; 1.7 DST; 6.6 Home Unit, adjustment types, grade change log; 3.4 SelfTest v1.

## v22: Sick calls and leave
Target: mid-November. Depends on Hours History (v20), Home Unit (v21).

| Item | What |
|---|---|
| 6.5 | Sick Call button: absence, off call lists for that shift + 24 h, OPEN calendar line, Log Call-Out pre-filled. |
| 6.2 | Leave button: Going (booking, home unit locked in) and Returning (department-average top-up with a confirm, `LV-####`). Needs Attention "Returning from leave" list. |
| 1.1 | Convert imported lines into real bookings, and an "Imported, not booked" list. |

Need: "Sick" or "Off" on the posted calendar.
Smoke test: a sick call opens the line and the call-out; a test leave return on a copy shows the people counted and the average, and Void Entry removes it.

## v23: Year-end
Build by Nov 30. Dry run on a copy the first week of December. Live Dec 27.
2.1 RolloverYear (frozen 2026 copy, 2027 pay periods from Dec 27, reset to zero, 2026 log archived, December-logged 2027 entries placed on the new roster); 2.2 Integrity check on open and on demand; 3.7 docs incl. a one-page rollover guide.
Need: who runs the rollover if Graeme is away; where the 2026 frozen copy is kept. Proposed default: leave spanning Jan 1 tops up only from the reset.

## v24: Employee changes
Target: first half of January, before Jan 16.
6.3 Move to unit (`UM-####`, up to the highest grade-2-and-below total, never lowered); 6.4 Add employee at the department average (`NH-####`), Rename, Remove (replaces RenameEmployee); 4 crew history (crew changes take effect from a date); 4 partial OT on postings.

## v25: Hardening
Target: February - March 2027. 3.1 Employee ID as the key across every table (verified migration); 3.2 sheet protection, very hidden sheets, VBA password; 3.6 calendar speed.
Need: whether Kronos exports an employee number.

## v26: Reporting
Target: March 2027 onward. Employee history report (PDF per person and date range); minimum staffing check; two-week board print; payback aging. (Member file and Who's next moved to M1.)

## At a glance
| Update | Target | Theme | Blocked on |
|---|---|---|---|
| v20 | Built Oct 9 | Safety and stats | done |
| M1 | Build after v20, live Nov 15 | Member workbook | members' folder path |
| v21 | Before Nov 1 | Call order accuracy | built Oct 12 |
| v22 | Mid-Nov | Sick calls and leave | v20, v21 |
| v23 | Build Nov 30, live Dec 27 | Year-end | who runs it |
| v24 | Early Jan, before Jan 16 | Employee changes | - |
| v25 | Feb - Mar | Hardening | Kronos employee number |
| v26 | Mar onward | Reporting | - |
