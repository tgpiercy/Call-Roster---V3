# Members' workbook - setup

Use this when the master is on v22 or later. Those releases already hold the publish code (`modPublish`), so Part A of the old M1 note is done. What is left is building the members' template once and publishing.

Files in this zip: `member/modMember.bas`, `member/ThisWorkbook.txt`. They work with a v22, v22.1 or v23 master.

About 15 minutes.

## 1. Build the members' template (once)
1. **File > New** blank workbook. **File > Save As** `Call Out Roster - Members TEMPLATE.xlsm` (type: Excel Macro-Enabled Workbook), **in the same folder as the master**.
2. Alt+F11 in that workbook. **File > Import File:** `member/modMember.bas`.
3. In its project tree open **ThisWorkbook**, Ctrl+A, paste the contents of `member/ThisWorkbook.txt`.
4. **Debug > Compile VBAProject.** No message expected.
5. Alt+F8 > **BuildMemberTemplate** > OK. It lays out Home, Roster, Who's next, Calendar, Postings and My history (M3.1: + the dWorked data sheet), with the buttons, and protects them.
6. **Save** (Ctrl+S) and close it. The template holds no data. That is correct.

## 2. Where the published file goes
For now Member View saves `Call Out Roster - Members.xlsm` **next to the master** (`PUBLISH_FOLDER = "."` at the top of `modOvertimeEntry`). Nothing to set.

At the network roll-out, change that one line to the members' folder. The comment above the setting shows the P: path.

## 3. Publish
1. In the master: Alt+F8 > **SystemCheck**. No "members' template not found" line. It shows "Members file: never published".
2. Click **Member View** on the roster. After a few seconds: "Members' file published ..." with the counts. Allow 10-40 seconds the first time.
3. Open `Call Out Roster - Members.xlsm` (next to the master), **with the master closed**. Click Enable Content.

## 4. Smoke test (members' file)
1. **Home** shows "Updated <today's time> | Kronos data <date>".
2. **Roster:** click a unit button. The list is in that unit's call order (1 = first). It matches the printed Call Out Roster for that unit. Click A - Z. Pick your name in Find me: your row is highlighted.
3. **Who's next:** pick a unit, tomorrow and Days. The list matches what Log Call-Out shows in the master for that unit and shift.
4. **Calendar:** pick this month and a unit. Off people and coverage lines appear. OPEN lines are red. Leave shows as "Leave"; a sick call or a personal day shows as "Off" (v24.3.1).
5. **Postings:** every open posting with its list.
6. **My history:** pick a name: that person's entries this year with **Call order** and a **Why** line on every charged row (M3), and the pay-period block with the charged hours, the posting refusals, the call-out refusals and Worked (Kronos) per period (M3.1; a dash until the roster has imported a Kronos export). **Print / PDF** saves "History <name> <date>.pdf" beside the file.
7. **Privacy:** Ctrl+F across the workbook (Options > Within: Workbook) for `@`, a phone number, `STD`, `PAR`, `LOA`, `WSIB`, `VOID`: nothing found.
7a. **M3 (v26):** on Home, **Larger type** toggles every sheet to 130 % and back; with a name picked in Find me the card shows **Am I next?** (the next shift they are listed for and their place, or "Not listed in the next 14 days"). Postings: a member who moved unit shows "-" and "(not on this list now)" on the old unit's list, with no Sign me up. Who's next: `~` before the hours of a planned mover's estimate, with its legend.
8. **Macros off:** close and reopen without enabling content. Roster shows A - Z with the filter arrows working.
9. Next day in the master: Needs Attention shows "Members file: last published ...". Double-click it: it publishes again.

## Troubleshooting
| Message | What to do |
|---|---|
| "The members' template wasn't found" | The template must be in the same folder as the master, with exactly the name in `MEMBER_TEMPLATE_NAME`. |
| "Check the drive, or PUBLISH_FOLDER" | The publish folder can't be reached. Only possible once `PUBLISH_FOLDER` points at a network folder. |
| "Is the file open on a plant PC?" | The publish saved a dated copy beside the members' file. Close the file there and publish again. |
| "NOT published. Something private would have gone in" | The message names the table and row. Fix that entry (an email or phone number typed into a note or reason) and publish again. |
| The members' views look stale after a publish | Open the file with macros enabled once. It redraws on open. |

## When a later release changes the members' code
The INSTALL note for that release says so. Open the TEMPLATE, Remove + Import `modMember`, re-paste `ThisWorkbook.txt` if it changed, Compile, Alt+F8 > BuildMemberTemplate, Save. Then click Member View once in the master. Since v24.4 the Install panel's `template` row does all of this.

**v26.5 = Members M3.1 (Oct 19, 2026):** the template must be rebuilt again (a new data sheet, dWorked = premium hours worked per pay period from the Kronos export); My history's pay-period block shows Charged hours, Posting refusals, Call-out refusals and Worked (Kronos).

**v26 = Members M3 (Oct 18, 2026):** the template must be rebuilt (the master refuses an older one: "the template's code is Members M2.1 and this roster needs Members M3"). The tables grew: dLog has Kind, Call Order and Why; dPostList has Exempt; dNext has EstHrs; dMeta's "Signup folder" is relative (".\Signups" beside the members' file). The `PUBLISH_REASONS` setting is gone from the master.
