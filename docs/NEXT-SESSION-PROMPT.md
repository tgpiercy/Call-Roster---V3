# Paste this into a fresh session to carry on (written Oct 18, 2026)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/next-round-2026-10-17.md,
docs/ease-of-use-plan-2026-10.md, briefs/v25.2.5.md (the bundle list) and briefs/v25.5.md (the pop-ups).
Confirm the four checkers and pytest are green on main and the tree is clean (two docs about the Desk
proposal are committed; nothing should be untracked).

Where things stand: the real roster is on v25.2 (installed Oct 17; Calendar 4.6 -> 2.5 s). v25.3 (the two
Desks + supervisor rights) is tagged and its zip is built; its install is pending, including one hand-check
row that needs a practice copy opened under a real supervisor's login (the gates). Everything is pushed to
origin (https://github.com/tgpiercy/Call-Roster---V3).

The numbering rule (Oct 17): features x.y as one install; smaller changes and fixes batched into x.y.5;
fixes to a bundle count up from it; never a single-fix install unless the roster is broken.

Here is my v25.3 install report: <paste the SmokeTest box, the Desk photo, what the boss said, and the
supervisor-login row's result>.

Then, in this order, asking every decision as a multiple-choice question (recommended option first):
1. Fold anything the install found into briefs/v25.2.5.md (it is the v25.3.5 bundle now).
2. Build the bundle (v25.3.5) before Dec 20: click-details on the calendar (item 0), Needs Attention without
   hover notes, the legacy writers and SmokeTest steps 19-20 out, the three admin award buttons greyed for a
   supervisor, the Help sheet, plus the install's findings. One commit per item, the four checkers before
   review, read-only reviewers in isolation worktrees, re-review the fixes, tag, zip, push from my Terminal.
3. The rollover rehearsal script (late November): one page for me and my boss on a practice copy, from
   docs/rollover-one-pager.md against the current code.
4. v25.5 in January: answer the four decisions at the end of briefs/v25.5.md first, then the call-out
   session rework, then the rest of that brief.
Still open and dated: the December dry run (first week), the install freeze Dec 20 - Jan 4, the live
rollover Dec 27, no installs the week before Jan 23, 2027; go-live not set.
```

## Why this file exists
A fresh session starts with none of this conversation. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state (the last line of the session's final message should say it was updated).

## The members' side: alignment with this week's changes (Oct 18)
What the members' file is: values pushed by the master (`PublishMembers`: dMeta, dRoster, dNext, dCal, dPost, dPostList, dLog, dAttend), privacy-screened in code, with its own front page (Find me, the month grid, My postings, Sign me up, My history, sick-time counts). It carries no rule and never names who logged or awarded anything.
- **Unchanged by v25.1-v25.3:** every published table, the privacy screens, the front page. The Desk and the supervisor rights are master-side only.
- **One gap the rights change opens:** a supervisor can now post and award, but **publishing (Member View) is still admin-only** (`PublishMembers` gate, the bar's `ui_adm_Member`, the Desk's members' tile is information for a supervisor, and the close-time reminder `PublishReminderOnClose` runs only for admins through `IsLiveRoster`). So after a supervisor's posting or award, the members' file is behind until an admin publishes. **Decided Oct 18: supervisors may publish** (the publish is code with the leave-word screens; the only requirement is write access to the members' folder). It is item 8 of `briefs/v25.2.5.md` (the v25.3.5 bundle): `PublishMembers` on `RequireSupervisor`, the bar's Member View for everyone, the Desk's members' tile live for a supervisor, the close-time reminder for supervisors too.
- **The M3 follow-up** (deep review 4.5) is the members' own ease-of-use list: withdrawals and login mismatches reported at LOAD, the PC name in the sign-up file name, `Kind` in dLog, `Exempt` in dPostList, `EstHrs` with `~`, the month picker, the phone pattern; plus C2 of the plan (an "Am I next?" tile, large-type mode). One members' install (the template rebuilt), with v26 or after.
- **At go-live:** `SIGNUPS_ON` (the inbox the award pop-up reads), the members' folder and its rights, `PUBLISH_FOLDER` moved off ".".
