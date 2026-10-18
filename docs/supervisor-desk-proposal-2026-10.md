# Proposal: a "Desk" view for the second admin

Written Oct 15, 2026 against v25. A proposal only: nothing here is built, and nothing changes for the primary admin. Implementation waits until the current version revisions are done.

Companion mock-up: `docs/supervisor-desk-mockup-2026-10.html` (open it in a browser).

## 1. What the admin view is today

Measured from `modules/modUI.bas` (BuildCommandBar), `modOvertimeEntry.ApplyViewCore` and the ten pop-ups in `paste-in/`.

| Item | Today |
|---|---|
| Where the controls live | One bar in row 1 of the OT Roster tab, 96 points tall, inside the frozen area |
| Buttons on that bar for an admin | 26, in 6 groups (SORT & FILTER 9, LOG 3, PEOPLE 3, POSTINGS 3, PRINT / PUBLISH 4, ADMIN 6) |
| Buttons a non-admin sees | 11 (the 9 sort buttons, Log Call-Out, Calendar) |
| Button text | Segoe UI 8.5 pt on a 24-point-tall shape; group labels 7.5 pt |
| Roster body | Segoe UI 9 pt; the status line is three lines of 9 pt in one cell |
| Columns on the roster | Names and details in A:S (19 columns), then 26 or 27 pay-period columns |
| Tabs an admin sees | About 24: the 8 everyday sheets plus the 16 that ApplyView shows to admins only (Absences, Coverage, Needs Attention, Holidays, Hours History, Grade Log, Attendance, Integrity, Year End, Crew History, Former, Planned Moves, Refusal Report, Error Log, Install Log, Overtime Posting) |
| Pop-ups | 10, from 412 to 600 points wide; labels 8 to 9 pt, lists 10 pt, the main action button 11 pt |
| Jobs that live only under Alt+F8 | SystemCheck, SelfTest, IntegrityCheck, MakePracticeCopy, SmokeTest, ShowErrorLog, ApplyPlannedMoves, RolloverYear, the one-time macros |

Everything on the bar is there because an admin might need it. For the primary admin that is right. For a second admin who does a handful of jobs a week and does not read small type, it is the wrong starting screen for four reasons.

1. The everyday jobs (sick call, call-out, post, award, time off) sit beside tools he should never touch (Void Entry, Add Adjustment, Employee, Import Kronos, Full Access, Convert imports), with nothing to say which is which.
2. Button labels are nouns from the system's vocabulary ("Load Posting", "Convert imports", "Member View"), not the job in his head ("someone called in sick").
3. The text is 8.5 pt on a 24-point button. The roster behind it is 9 pt across 45 columns.
4. There is no "what needs doing today" anywhere on the first screen. Needs Attention exists, but it is a tab among 24, or a small button on the Calendar.

The buttons, pop-ups and rules underneath are sound and tested. The proposal changes the front door, not the house.

## 2. Principles for the second admin's view

- **One screen, one question: "what do you need to do?"** Every everyday job is one large button with a plain-language label. Nothing else is on that screen.
- **Tell him what is waiting.** The screen opens with the counts that matter today (shifts needing cover, postings ready to award, paybacks owed, how old the Kronos data is, when the members' file was last published). Each count is a button that takes him there.
- **Bigger by default.** 12 pt labels, 14 pt headings, buttons 36 points tall, the roster and calendar at 120 percent zoom. Nothing under 10 pt on his screens.
- **Fewer tabs.** He sees five tabs. The other nineteen stay in the file, reachable only through buttons, exactly as a supervisor sees them today.
- **He can never be locked out.** A "Show everything" button on his screen gives the full master layout for the session. The primary admin's layout does not change at all.
- **No new rules.** The Desk calls the same macros the bar calls. Every gate (RequireAdmin, RequireWritable, the privacy screens) stays where it is.

## 3. The scheme: three tiers, one new sheet

### Tier 1: the Desk (new sheet, first tab, his landing page)

A new sheet named **Desk**, built by code like the Calendar is (shapes with OnAction, redrawn on open and after every pop-up closes). Layout, top to bottom:

| Band | Content |
|---|---|
| Header | "Call Out Roster", today's date, the pay period, "Kronos data through <date>" (red when more than 7 days old). 14 pt. |
| TODAY | Four tiles, 12 pt bold, each a button: "N shifts need cover" (Calendar / Needs Attention), "N postings ready to award" (Award Posting), "N paybacks owed" (Shift Switch, OWED list), "Members' file: published <when>" (Member View). A tile with nothing to do is grey; one with work is coloured. |
| RIGHT NOW | Three buttons: **Someone called in sick** (Sick Call), **Fill a shift now** (Log Call-Out), **What needs cover** (Calendar). |
| PLAN AHEAD | Four buttons: **Post a shift** (New Posting), **Award a posting** (Award Posting), **Book time off** (Book Time Off), **Shift switch or payback** (Shift Switch). |
| PEOPLE | One button: **Leave: going or returning** (Leave). |
| WHEN YOU ARE DONE | Two buttons: **Update the members' file** (Member View), **Print the call-out roster** (Call Out Roster). |
| HELP | **How do I ...** (opens the supervisor quick guide), **Something went wrong** (opens an email draft to the primary admin with the newest Error Log line already pasted in). |
| Footer | Two small links: "Full roster >" (OT Roster tab), "Show everything (primary admin tools) >". |

Fourteen buttons, grouped by when he would reach for them, every label a verb phrase. The mock-up shows it.

### Tier 2: the working sheets he keeps

| Tab | What changes for a Desk user |
|---|---|
| OT Roster | Opens at 120 percent zoom. The bar shows the 9 sort buttons, Log Call-Out, Calendar and one new **< Desk** button. The PEOPLE, POSTINGS, PRINT / PUBLISH and ADMIN groups are hidden (he reaches those jobs from the Desk). EMAIL and Payroll # columns hidden as for a supervisor; Phone stays. |
| Calendar | 120 percent zoom. Unchanged otherwise (its own buttons already read well). |
| Needs Attention | Unchanged. Reached from the TODAY tiles. |
| Postings | Unchanged. |
| OT Log | Visible, read-only as now. |

Every other sheet is very-hidden for him, the same list ApplyView already hides for non-admins. The pop-ups still write to them.

### Tier 3: the master layout (unchanged)

The primary admin's login keeps today's bar, tabs and Alt+F8 tools exactly as they are. "Show everything" on the Desk switches a Desk user to this layout for the session (ApplyView with the full flag), and a **< Desk** button on the full bar switches back. Nothing is saved differently: Workbook_BeforeSave already stores the restricted view.

### Who gets the Desk

Recommended: a new constant beside ADMIN_USERS, for example `DESK_USERS = "lanthierr"`. A login on both lists has every admin right and lands on the Desk. A login on ADMIN_USERS alone lands on the roster as today. The primary admin can open the Desk from the bar at any time to see what the other admin sees. This fits the version-constant pattern already used for admins and needs no settings sheet.

Alternative (not recommended): every admin lands on the Desk and the primary admin clicks through to the full view each time. One more click a day for the person who uses the file most.

### Font size, in three layers

| Layer | What | Cost |
|---|---|---|
| Windows display scaling | Settings > System > Display > Scale 125 percent on his PC. Makes every Excel pop-up and sheet larger today, no code. | None. Worth doing now. |
| Sheet zoom | Desk users open the roster and calendar at 120 percent (ActiveWindow.Zoom when the sheet activates). The Desk itself is drawn at 100 percent with 12 to 14 pt type. | Small: one line per sheet activate, inside the Desk mode. |
| Pop-up scale | The pop-ups lay themselves out in code with absolute points. A `UI_SCALE` constant (1.0 for the master, 1.2 for Desk users) multiplied into Left / Top / Width / Height and Font.Size in each form's layout routine would enlarge them without reflow. | Medium: ten forms, mechanical, but every one must be re-pasted and the layouts re-checked by eye on his screen. Proposed as a second step after the Desk lands, starting with the three he uses most (Sick Call, Callout Session, Award Overtime Posting). |

Independent of scale: the three smallest pieces of text in the pop-ups he uses (8 pt legend and status lines on Callout Session, the 8 pt "ACCEPTED PART OF THE HOURS" button) should go to 10 pt regardless. Those are one-line changes.

### What the TODAY tiles read

All four figures already exist; the Desk only displays them.

| Tile | Source today |
|---|---|
| Shifts needing cover | The open-line count the Calendar's "N need cover >" button shows (modCalendar, the Needs Attention builder) |
| Postings ready to award | The Needs Attention "ready to award" rows (closing date passed, not awarded) |
| Paybacks owed | The Shift Switch OWED list count |
| Members' file | The publish timestamp Needs Attention already checks ("more than a day old") |

Kronos age and the pay period come from the same calls that build the roster status line (SetRosterStatus).

## 4. Example

The mock-up file shows the Desk at the size it would appear on a 1920 x 1080 screen at 100 percent, with the current bar drawn above it at true scale for comparison. Text version of the Desk:

```
CALL OUT ROSTER                        Wed Oct 15, 2026   |   Pay period Oct 11 - Oct 24   |   Kronos data through Oct 13

TODAY
[ 3 shifts need cover ]   [ 2 postings ready to award ]   [ 1 payback owed ]   [ Members' file: updated yesterday ]

RIGHT NOW                              PLAN AHEAD
[ Someone called in sick      ]        [ Post a shift                ]
[ Fill a shift now (call-out) ]        [ Award a posting             ]
[ What needs cover (calendar) ]        [ Book time off               ]
                                       [ Shift switch or payback     ]
PEOPLE
[ Leave: going or returning   ]        WHEN YOU ARE DONE
                                       [ Update the members' file    ]
HELP                                   [ Print the call-out roster   ]
[ How do I ... ]  [ Something went wrong ]

Full roster >                                        Show everything (primary admin tools) >
```

A walk through his most common morning, on the Desk:

1. Opens the file. Lands on the Desk. TODAY says "1 shift needs cover" in orange.
2. Clicks **Someone called in sick**. The Sick Call pop-up opens as it does today. SAVE + START CALL-OUT.
3. Log Call-Out opens filled in. He calls down the list and submits.
4. Back on the Desk, the tile now reads "0 shifts need cover" in grey.
5. Clicks **Update the members' file**. Done.

He has not seen a tab name, a sort button, an ID prefix or an admin tool.

## 5. Size of the job

| Piece | Where | Notes |
|---|---|---|
| New module `modDesk` | modules/ | BuildDesk (draw), RefreshDesk (recount the tiles), the 14 OnAction wrappers, DeskUsers / IsDeskUser. About the size of modAttendance. |
| Desk mode in ApplyView | modOvertimeEntry.ApplyViewCore | A third state beside admin / non-admin: admin rights, supervisor visibility plus Desk, Needs Attention and Postings, the bar's group visibility, zoom on activate. |
| Open step | ThisWorkbook | Step 10 activates the Desk for Desk users instead of the roster. |
| One new button on the master bar | modUI.BuildCommandBar | "< Desk". ApplyRosterLayout rebuilds the bar, so it is a one-time macro on install. |
| Help buttons | modDesk | Quick guide: a PDF path constant beside PUBLISH_FOLDER. Email: the existing award-email pattern plus the newest Error Log line. |
| Checks | modSystem, modSelfTest | SystemCheck counts 25 modules; SelfTest cases for the view-mode decision per login and the tile counts on a known log. |
| Docs | docs/ | A one-page "Desk guide" for him; the supervisor quick guide gains a Desk section. |
| Not touched | Every rule module, every pop-up (until the scale step), the OT Log, the members' file. |

One version, one brief, one zip, the usual review. The pop-up scale step is its own later version.

## 6. Decisions (Graeme, Oct 15, 2026)

1. Who gets the Desk: a `DESK_USERS` list beside ADMIN_USERS. A login on both lists has every admin right and lands on the Desk; the primary admin keeps today's layout.
2. Pop-up size: the Desk, zoom and tabs ship first as one version. The `UI_SCALE` pass through the ten form layouts is its own later version, starting with Sick Call, Callout Session and Award.
3. Tabs for a Desk user: the five above (Desk, OT Roster, Calendar, Needs Attention, Postings).
4. The "Something went wrong" button: an email draft to the primary admin with the newest Error Log line, on the existing award-email pattern.

Build waits until the current version revisions are finished. When it starts, this document becomes the brief's source.
