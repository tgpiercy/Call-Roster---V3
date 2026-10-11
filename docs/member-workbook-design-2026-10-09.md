# Member workbook: design (Oct 9, 2026)

A separate macro workbook for union members, built **after v20 is verified** and live before the Nov 15 union note. It replaces today's values-only Member View.

## Decisions (Oct 9)
| Question | Answer |
|---|---|
| Where members open it | **Plant PCs, desktop Excel**. Macros run after Enable Content. |
| What they see | **Everyone's hours and call order** (full transparency, like today's Member View) |
| Content | **Roster, Coverage calendar, Open postings, My history, Who's next by unit** |
| Refresh | **Publish button** in the master, any time. Needs Attention reminds when the published copy is more than a day old. |

## Principles
- **The master pushes; the member file never reads the master.** Members can't reach the master's folder, leave codes or phone numbers.
- **No rules in the member file.** Call order, eligibility ("who's free"), charging and the rotation are all worked out by the master and published as values. The member file only sorts, filters and looks things up, so it can never disagree with the master.
- **Harmless if tampered with.** It sits in the read-only member folder, so a member changing their copy changes nothing for anyone else.
- **Privacy.**
  - Not published: phone numbers, emails, leave types (calendar shows "Leave"), void notes, who logged an entry.
  - Log notes stay out unless `PUBLISH_REASONS` is True.

## How it's built and published
1. **Member template** (`Call Out Roster - Members TEMPLATE.xlsm`), kept in the supervisor folder. Built once:
   1. Start a blank workbook.
   2. Import `modMember.bas`.
   3. Paste its ThisWorkbook code.
   4. Run **BuildMemberTemplate**, which lays out every sheet and button.
   5. Save as `.xlsm`.
2. **Publish button** in the master (`PublishMembers`):
   1. Open the template.
   2. Clear its hidden data sheets and paste the current values.
   3. Stamp the time and version.
   4. Save it as `Call Out Roster - Members.xlsm` in the member folder and close it.

   It takes a few seconds and never changes the master.
3. If the member file is open on a plant PC when you publish, the save fails. The Publish button reports that and keeps a dated copy beside it, so nothing is lost.

## Published data (hidden, very hidden sheets in the member file)
| Sheet | Contents |
|---|---|
| `dMeta` | Published at, by whom, master version, Kronos data date, current pay period. |
| `dRoster` | Each person: name, crew, Total / Worked / Refused hours, and for each department: grade code + **call position** (worked out by the master: grade > hours > seniority). |
| `dNext` | **Who's free**: for each unit, each date in the next 14 days, Days and Nights: the people free that shift, in call order. Uses the master's own eligibility check (rotation, rest rule, holidays, booked off, switches). |
| `dCal` | Calendar lines for last month through the next 3 months: date, shift, unit, who's off / need (leave shown as "Leave"), coverage text, type. |
| `dPost` | Open postings (ID, unit, date, shift, hours, positions, reason, closing date) and each posting's list (rank, name, hours). |
| `dLog` | This year's non-void OT Log rows: date, Ref, type, unit, shift, employee, outcome, hours charged, pay period. |

## Member sheets
| Sheet | What it does |
|---|---|
| **Home** | "Updated Oct 21, 3:12 PM / Kronos data Oct 20". Big buttons to each view. **Find me** box. |
| **Roster** | Everyone, with hours and each department's call position. Buttons: one per department (sorted by that department's call order, people without the code hidden), A-Z, Show all. **Find me** jumps to and highlights a name. Filter by crew. |
| **Who's next** | Pick unit, date (next 14 days) and Days / Nights: the free people in call order, top first, with hours. |
| **Calendar** | Pick month and unit: the coverage calendar, read-only, same colours as the master. |
| **Postings** | Open postings with closing date and the list of who's on each. Sign-up is still on the paper sheet. |
| **My history** | Pick your name: your calls, outcomes and charges this year, with totals per pay period. |

All sheets are protected (macros only). Buttons are macro-driven; without Enable Content, the file still opens on Roster sorted A-Z with filters working.

## Master changes
- New module `modPublish` with `PublishMembers`. It replaces the old Member View `.xlsx` publish; the button keeps the name **Member View**.
- A Needs Attention line: "Member file last published Oct 20 (1 day ago)".
- Settings: the template path, the member folder (`PUBLISH_FOLDER`) and the file name.

## Build steps
1. Verify v20 (attach the zip): the three Oct 9 morning fixes are present, and SystemCheck is clean.
2. Build `modMember.bas` + the member ThisWorkbook code, and `modPublish.bas` in the master.
3. Review pass, package, install on a copy, then publish a test file and open it on a plant PC.
4. Live with wave 2 (Nov 15), with the union note.

## Still to confirm
- The member folder path. Today's publish goes to `...\Production Overtime Roster\Call Out Roster - Member View.xlsx`.
- Whether members should see each other's **My history**. The default is yes, consistent with full transparency.
