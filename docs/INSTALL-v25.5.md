# Call Out Roster v25.5 - install (the Install panel)

The ALPHA BUILD (`docs/timeline-2026-10-18.md`): the v25.3.5 bundle (`briefs/v25.2.5.md` items 0-2, 6-9) folded into the pop-up release (`briefs/v25.5.md` items 1-7), one install. Alpha testing Nov 1-15; what it turns up is batched into v25.5.5.

Needs **v25.3 installed** and **Roster Updater.xlsm v4.1**. **The updater does not change.** The panel's swap replaces every module (all 25 stamps on v25.5; no new module), **re-pastes ThisWorkbook and EVERY FORM** (the call-out pop-up is rebuilt; the other eight gain `UI_SCALE`), and leaves the members' template alone. No rule changed; SelfTest's `viewmode` suite gains 6 cases: 581 cases, 68 suites. SmokeTest is 19 steps (the two v25.2 draw-compare steps went with the legacy writers).

**Back up first. Do it on a copy before the real roster.** The panel's group 3 runs `ApplyRosterLayout` once (the bar is rebuilt: the supervisor buttons under their new names, Undo last entry, Find a person) and asks you to confirm `SUPERVISOR_USERS` and `UI_SCALE`.

## What's new
- **Supervisors by login (bundle item 9).** `SUPERVISOR_USERS = "mihaychukc,sheltond,campeauc"` beside `ADMIN_USERS`. A login on it (or an admin) gets the supervisor's jobs - sick call, new posting, award to the list, print, and now **Member View (item 8)** - and lands on the supervisor Desk; anyone else is a plain user (log a call-out, the calendar, the sorts) and lands on the roster. `RequireSupervisor` refuses by name. The close-time "publish the members' file?" question asks supervisors too.
- **The calendar: click a line for its details, no hover boxes (item 0); Needs Attention without notes (item 1).** One click writes the line's details into the row under the title and onto the status bar; an empty cell brings the hint back; double-click for the options as before. The draw makes no comment shapes (the v25.2 numbers put the OPEN / POSTED boxes at 1.6 of the 2.5 s). The v25.2 legacy writers and SmokeTest steps 19-20 are gone (item 2).
- **The award pop-up greys No Takers / Award Out of Dept / Cancel Request for a supervisor (item 6).**
- **A Help tab inside the file (item 7).** Desk > How do I ... draws it (one heading and a paragraph per job, 11 pt, "Words you will see" near the end); < Back; the printable guide stays optional behind `GUIDE_PATH`.
- **The call-out session reworked (v25.5 item 1).** The person being called is the screen (name 16 pt, phone 24 pt, one line of what matters); one key per outcome - A accepted, R refused, L left message, U unavailable (the reason list; Enter records it), P part of the hours; K skip, B back, Enter submit, Esc close; "?" marks the left-message pile and **Call back** goes down it; the progress line with the start time; a four-line summary (the detail on the status bar and in the OT Log); type 9.5-11 pt; 500 x 600. Same rules, same seven designer controls.
- **`UI_SCALE` on every pop-up (item 2).** One constant (1 ships; 1.2 on a 1920 x 1080 screen) multiplies every pop-up's positions, sizes and point sizes.
- **Plain language (item 3):** names before IDs in the saved-messages, verb-phrase buttons, S/S explained in the call-out legend, a glossary on the Help tab and in the quick guide.
- **Shortcuts and what's-new (item 4):** Ctrl+Shift+S sick call, Ctrl+Shift+C call-out, Ctrl+Shift+K calendar, Ctrl+Shift+D the Desk, while the roster is the active workbook (listed under the Desk's HELP group); a three-line what's-new box once per login after an install (Alt+F8 > ShowWhatsNew brings it back).
- **Desk refinements (item 5):** the cover tile names the next uncovered shift; the header says which crews are on today; Suggest an improvement beside Something went wrong.
- **Undo my last log entry (item 6, admin):** the admin Desk, and Undo last entry on the bar, void this login's newest OT Log entry of today through Void Entry's own confirm and reason (log rows only: a P- / LV- / UM- / NH- entry's box says what stays and the next step).
- **Find a person (item 7):** a FIND group at the right end of the bar (supervisors and admins): part of a name selects the row; the status bar says "N of M in the <unit> call order today".
- Settings (modOvertimeEntry, under DESK_USERS): `SUPERVISOR_USERS`, `UI_SCALE`. modSystem: `WHATS_NEW`.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.5` folder.
3. Click the one blue button, read its box, click the next. Group 3: the bar rebuild, System check ("25 of 25 on v25.5", ThisWorkbook v25.5, every form v25.5), Self test (581 / 68), `SUPERVISOR_USERS`, `UI_SCALE`. Group 4: the practice copy + SmokeTest (19 steps), the call-out pop-up and its keys, the sick-call and award pop-ups at scale, the calendar's click details, the Help tab, the shortcuts and the what's-new box, the Desk's new lines, Find a person, a supervisor's login and a plain user's.

By hand instead: remove and import the 25 modules, re-paste ThisWorkbook and the nine forms from `paste-in/`, Debug > Compile, Alt+F8 > ApplyRosterLayout, then the group 3-4 rows yourself.

## Then
Send the SmokeTest box, a photo of the call-out pop-up on the plant PC (say whether anything is cut off or under 10 pt), and what the first supervisor says on seeing the Desk under their own login.
