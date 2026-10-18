# Paste this into a fresh session to carry on (written Oct 18, 2026)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/next-round-2026-10-17.md,
docs/ease-of-use-plan-2026-10.md, briefs/v25.2.5.md (the bundle list) and briefs/v25.5.md (the pop-ups).
Confirm the four checkers and pytest are green on main and the tree is clean (two docs about the Desk
proposal are committed; nothing should be untracked).

Where things stand: the real roster is on v25.3 (installed Oct 18: the two Desks + supervisor rights; v25.2's
calendar speed-up 4.6 -> 2.5 s is in it). The hand-check row that needs a practice copy opened under a
supervisor's login (the gates) is ON HOLD by Graeme's choice. The supervisor logins are known
(briefs/v25.2.5.md item 9: mihaychukc, sheltond, campeauc) and the bundle makes supervisors a named list.
Everything is pushed to origin (https://github.com/tgpiercy/Call-Roster---V3).

The numbering rule (Oct 17): features x.y as one install; smaller changes and fixes batched into x.y.5;
fixes to a bundle count up from it; never a single-fix install unless the roster is broken.

Here is my v25.3 install report: <paste the SmokeTest box, the Desk photo, what the boss said, and the
supervisor-login row's result>.

The timeline (docs/timeline-2026-10-18.md): all development before the rollout. v25.5 = the ALPHA BUILD by
Oct 31 = the v25.3.5 bundle (briefs/v25.2.5.md items 0-9) folded in + the pop-up release (briefs/v25.5.md,
its four decisions taken by default unless I say otherwise). Alpha testing Nov 1-15, with the documentation
and the what-if scenario sheets written during it. v26 = the BETA BUILD by Nov 15 (the reports + M3 members
+ the alpha's findings), presented to management and the union that day. Revisions v26.5 as needed; beta
through December; training from Dec 1; the December dry run; the freeze Dec 20 - Jan 4 (the live rollover
Dec 27); go-live early January with v25.4 (the Payroll # key) last.

Then, in this order, asking every decision as a multiple-choice question (recommended option first):
1. Build v25.5 (the alpha build): the bundle's items first (click-details item 0, named supervisors item 9,
   publish for supervisors item 8, the Help sheet, the legacy writers and steps 19-20 out, Needs Attention
   without notes, the greyed award buttons), then the call-out session rework, UI_SCALE on every pop-up, the
   language pass, shortcuts, what's-new, the Desk refinements, undo, find. One commit per item, the four
   checkers before review, read-only reviewers in isolation worktrees, re-review the fixes, tag, zip, push
   from my Terminal. Ship by Oct 31.
2. Nov 1-14: v26 (the beta build) from docs/code-kit-2026-10/briefs/07-v26-reporting.md + deep review 4.6,
   M3 from deep review 4.5, plus the alpha's findings; and the documentation set in
   docs/timeline-2026-10-18.md (the what-if scenario sheets, the guides, the Nov 15 pack, the training plan).
3. After Nov 15: the revisions, the rollover rehearsal page for the December dry run, the freeze, go-live.
```

## Why this file exists
A fresh session starts with none of this conversation. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state (the last line of the session's final message should say it was updated).

## The members' side: alignment with this week's changes (Oct 18)
What the members' file is: values pushed by the master (`PublishMembers`: dMeta, dRoster, dNext, dCal, dPost, dPostList, dLog, dAttend), privacy-screened in code, with its own front page (Find me, the month grid, My postings, Sign me up, My history, sick-time counts). It carries no rule and never names who logged or awarded anything.
- **Unchanged by v25.1-v25.3:** every published table, the privacy screens, the front page. The Desk and the supervisor rights are master-side only.
- **One gap the rights change opens:** a supervisor can now post and award, but **publishing (Member View) is still admin-only** (`PublishMembers` gate, the bar's `ui_adm_Member`, the Desk's members' tile is information for a supervisor, and the close-time reminder `PublishReminderOnClose` runs only for admins through `IsLiveRoster`). So after a supervisor's posting or award, the members' file is behind until an admin publishes. **Decided Oct 18: supervisors may publish** (the publish is code with the leave-word screens; the only requirement is write access to the members' folder). It is item 8 of `briefs/v25.2.5.md` (the v25.3.5 bundle): `PublishMembers` on `RequireSupervisor`, the bar's Member View for everyone, the Desk's members' tile live for a supervisor, the close-time reminder for supervisors too.
- **The M3 follow-up** (deep review 4.5) is the members' own ease-of-use list: withdrawals and login mismatches reported at LOAD, the PC name in the sign-up file name, `Kind` in dLog, `Exempt` in dPostList, `EstHrs` with `~`, the month picker, the phone pattern; plus C2 of the plan (an "Am I next?" tile, large-type mode). One members' install (the template rebuilt), with v26 or after.
- **At go-live:** `SIGNUPS_ON` (the inbox the award pop-up reads), the members' folder and its rights, `PUBLISH_FOLDER` moved off ".".
