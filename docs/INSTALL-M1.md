# Call Out Roster M1 - members' workbook - install

Two parts: the master gets the Publish code; the members' template is built once from `member/`.
Needs v21 installed and checked first. Takes about 20 minutes.

**Back up first.**

## Part A: the master

1. In the VBA editor (Alt+F11), for EACH of these: right-click > **Remove** (answer **No** to export), then **File > Import File** from `modules/`:
   `modOvertimeEntry`, `modOTLog`, `modCalendar`, `modShiftSwitch`, `modSystem`
2. **File > Import File:** `modules/modPublish.bas` (new).
3. **Debug > Compile VBAProject.** No message expected.
4. Settings (modOvertimeEntry, near the top): `PUBLISH_FOLDER` is the members' folder (unchanged). `MEMBER_TEMPLATE_NAME` and `MEMBER_FILE_NAME` are the two file names; leave them unless you want different names.
5. Save.

## Part B: build the members' template (once)

1. **File > New** blank workbook. **File > Save As** `Call Out Roster - Members TEMPLATE.xlsm` (type: Excel Macro-Enabled Workbook) **in the same folder as the master**.
2. Alt+F11 in that workbook. **File > Import File:** `member/modMember.bas`.
3. In its project tree open **ThisWorkbook**, Ctrl+A, paste the contents of `member/ThisWorkbook.txt`.
4. **Debug > Compile VBAProject.** No message expected.
5. Alt+F8 > **BuildMemberTemplate** > OK. It lays out Home, Roster, Who's next, Calendar, Postings, My history and the hidden data sheets, with the buttons, and protects them.
6. **Save** (Ctrl+S) and close it. The template is empty of data; that's correct.

## Part C: publish

1. Back in the master: **Alt+F8 > SystemCheck**. Expect "Modules: 13 of 13 on v21" and a "Members file: never published" line, no template warning.
2. Click **Member View** on the roster (or Alt+F8 > PublishMemberView). After a few seconds: "Members' file published ..." with the counts. Expect 10-40 s the first time.
3. Open `Call Out Roster - Members.xlsm` from the members' folder **on a plant PC, with the master closed**. Enable Content.

## Smoke test (members' file)
1. **Home** shows "Updated <today's time> | Kronos data <date>".
2. **Roster:** click a unit button: the list is in that unit's call order (1 = first), people without the code are gone; it matches the printed Call Out Roster for that unit. Click A - Z. Pick your name in Find me: your row is highlighted.
3. **Who's next:** pick a unit, tomorrow and Days: the list matches what Log Call-Out shows in the master for that unit and shift (same people, same order).
4. **Calendar:** pick this month and a unit: off people and coverage lines appear; OPEN lines are red; leave shows as "Leave".
5. **Postings:** every open posting with its list.
6. **My history:** pick your name: your entries this year and the pay-period totals.
7. **Privacy:** Ctrl+F across the workbook (Within: Workbook, incl. hidden sheets via the Find dialog's Options) for `@`, your phone number, `STD`, `PAR`, `LOA`, `VOID`: nothing found. The very hidden sheets can be checked by unhiding them in the VBA editor (Properties > Visible) on a copy.
8. **Macros off:** close and reopen without enabling content: Roster shows A - Z with the filter arrows working.
9. Next day in the master: Needs Attention shows "Members file: last published ...". Double-click it: it publishes again.

## Troubleshooting
- "The members' template wasn't found": the template must be in the same folder as the master with exactly the name in `MEMBER_TEMPLATE_NAME`.
- "Is the file open on a plant PC?": the publish saved a dated copy beside the members' file. Close the file there and publish again.
- "NOT published. Something private would have gone in": the message names the table and row; fix that entry (an email or phone number typed into a note or reason) and publish again.
- The members' views look stale after a publish: open the file with macros enabled once; it redraws on open.
