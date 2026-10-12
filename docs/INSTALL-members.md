# Members' workbook - setup

Use this when the master is on v22 or later. Those releases already hold the publish code (`modPublish`), so Part A of the old M1 note is done. What is left is building the members' template once and publishing.

Files in this zip: `member/modMember.bas`, `member/ThisWorkbook.txt`. They work with a v22, v22.1 or v23 master.

About 15 minutes.

## 1. Build the members' template (once)
1. **File > New** blank workbook. **File > Save As** `Call Out Roster - Members TEMPLATE.xlsm` (type: Excel Macro-Enabled Workbook), **in the same folder as the master**.
2. Alt+F11 in that workbook. **File > Import File:** `member/modMember.bas`.
3. In its project tree open **ThisWorkbook**, Ctrl+A, paste the contents of `member/ThisWorkbook.txt`.
4. **Debug > Compile VBAProject.** No message expected.
5. Alt+F8 > **BuildMemberTemplate** > OK. It lays out Home, Roster, Who's next, Calendar, Postings and My history, with the buttons, and protects them.
6. **Save** (Ctrl+S) and close it. The template holds no data. That is correct.

## 2. Where the published file goes
The master publishes to `PUBLISH_FOLDER` (top of `modOvertimeEntry`). It ships pointing at the P: drive.

**While you are working on your desktop, P: is not there, so Member View stops with "Check the drive, or PUBLISH_FOLDER".** For local testing, change that one line to a folder that exists, for example:

```vb
Public Const PUBLISH_FOLDER As String = "C:\Users\Graeme\Desktop\Roster\Members\"
```

Create the folder first. Keep the trailing backslash. Each time you re-import `modOvertimeEntry` from a zip the line goes back to P:, so change it again, until the network paths are wired in at go-live.

## 3. Publish
1. In the master: Alt+F8 > **SystemCheck**. No "members' template not found" line. It shows "Members file: never published".
2. Click **Member View** on the roster. After a few seconds: "Members' file published ..." with the counts. Allow 10-40 seconds the first time.
3. Open `Call Out Roster - Members.xlsm` from the publish folder, **with the master closed**. Click Enable Content.

## 4. Smoke test (members' file)
1. **Home** shows "Updated <today's time> | Kronos data <date>".
2. **Roster:** click a unit button. The list is in that unit's call order (1 = first). It matches the printed Call Out Roster for that unit. Click A - Z. Pick your name in Find me: your row is highlighted.
3. **Who's next:** pick a unit, tomorrow and Days. The list matches what Log Call-Out shows in the master for that unit and shift.
4. **Calendar:** pick this month and a unit. Off people and coverage lines appear. OPEN lines are red. Leave shows as "Leave", sick as "Sick", personal days as "Personal".
5. **Postings:** every open posting with its list.
6. **My history:** pick a name: that person's entries this year and the pay-period totals.
7. **Privacy:** Ctrl+F across the workbook (Options > Within: Workbook) for `@`, a phone number, `STD`, `PAR`, `LOA`, `WSIB`, `VOID`: nothing found.
8. **Macros off:** close and reopen without enabling content. Roster shows A - Z with the filter arrows working.
9. Next day in the master: Needs Attention shows "Members file: last published ...". Double-click it: it publishes again.

## Troubleshooting
| Message | What to do |
|---|---|
| "The members' template wasn't found" | The template must be in the same folder as the master, with exactly the name in `MEMBER_TEMPLATE_NAME`. |
| "Check the drive, or PUBLISH_FOLDER" | The publish folder can't be reached. See step 2. |
| "Is the file open on a plant PC?" | The publish saved a dated copy beside the members' file. Close the file there and publish again. |
| "NOT published. Something private would have gone in" | The message names the table and row. Fix that entry (an email or phone number typed into a note or reason) and publish again. |
| The members' views look stale after a publish | Open the file with macros enabled once. It redraws on open. |

## When a later release changes the members' code
The INSTALL note for that release says so. Open the TEMPLATE, Remove + Import `modMember`, re-paste `ThisWorkbook.txt` if it changed, Compile, Save. Then click Member View once in the master.
