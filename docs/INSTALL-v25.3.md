# Call Out Roster v25.3 - install (the Install panel)

The Desk: the landing screen for the second admin (`briefs/v25.3.md`, from the Oct 15 proposal). One install; what its testing turns up is batched into the next bundle.

Needs **v25.2 installed** and **Roster Updater.xlsm v4.1**. **The updater does not change.** The panel's swap replaces every module (one NEW module, `modDesk`; all 25 stamps on v25.3), **re-pastes ThisWorkbook** (a Desk user lands on the Desk; the view survives a save), and leaves the forms and the members' template alone. No rule changed; SelfTest gains one suite (`viewmode`, 7 cases): 575 cases, 68 suites.

**Back up first. Do it on a copy before the real roster.** The panel's group 3 runs `ApplyRosterLayout` once (the "< Desk" button on the bar) and asks you to set `DESK_USERS`.

## What it is
- A **Desk** tab, first in the file, for a login on both `ADMIN_USERS` and `DESK_USERS`: a header (today, the pay period, how old the Kronos data is), four TODAY tiles (shifts needing cover in the next 14 days, postings ready to award, paybacks owed, the members' file's age; grey when there is nothing to do) and fourteen plain-language buttons at 12 pt, each opening the usual pop-up. A supervisor's tabs plus the Desk and Needs Attention (the admin-only tabs go); the roster and the Calendar at 120 %. **Show everything** gives the full master layout for the session; **< Desk** on the bar (admins only) comes back. Your own login keeps today's layout.
- **Something went wrong** opens an email draft to `PRIMARY_ADMIN_EMAIL` with the newest Error Log line in it. **How do I ...** opens the quick guide once `GUIDE_PATH` is set (empty until go-live: it says to ask you).
- Settings (modOvertimeEntry, under ADMIN_USERS): `DESK_USERS` (set at install), `PRIMARY_ADMIN_NAME`, `PRIMARY_ADMIN_EMAIL`, `GUIDE_PATH`.
- Nothing changes for members or for the other supervisors; no pop-up changes (the larger pop-up text is the UI_SCALE version, later).
- `docs/desk-guide.md` is his one page; the supervisor quick guide has a Desk paragraph.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.3` folder.
3. Click the one blue button, read its box, click the next. Group 3: the bar rebuild, System check ("25 of 25 on v25.3", ThisWorkbook v25.3), Self test (575 / 68), `DESK_USERS`. Group 4: the practice copy + SmokeTest (21 steps), the Desk from your own login, save-and-reopen.

By hand instead: import `modDesk` (new), remove and import the other 24 modules, re-paste ThisWorkbook from `paste-in/ThisWorkbook.txt`, Debug > Compile, Alt+F8 > ApplyRosterLayout, then the group 3-4 rows yourself.

## Then
Send the SmokeTest box, a photo of the Desk from your login, and what your boss says when he first sees it.
