# Paste this into a fresh session to carry on (written Oct 18, 2026, evening: the v25.5 build)

```
Read CLAUDE.md in C:\Users\Graeme\Claude.code\callout-roster first, then docs/timeline-2026-10-18.md,
briefs/v25.2.5.md (the bundle's items 0-9) and briefs/v25.5.md (the pop-up release). Confirm the four
checkers (py -3.14 tests/vbacheck.py . / shadowcheck / undeclcheck / callcheck) and py -3.14 -m pytest
tests -q are green on main and the tree is clean (main is at 2bfa424 or later; everything is pushed to
origin).

Where things stand: the real roster is on v25.3 (installed Oct 18: the two Desks, supervisor rights,
the v25.2 calendar speed-up). The supervisor-login gate check is on hold. The supervisor logins are
mihaychukc, sheltond, campeauc (briefs/v25.2.5.md item 9).

The timeline (Graeme, Oct 18): all development before rollout. v25.5 = the ALPHA BUILD, to ship by
Oct 31 = the bundle's items folded in + the pop-up release. Alpha testing Nov 1-15. v26 = the BETA
BUILD by Nov 15 (the reports + M3 members + alpha findings), presented to management and the union
that day. The four decisions at the end of briefs/v25.5.md are taken by default (UI_SCALE for
everyone, the key letters as written, the four-line summary, the 500 x 600 form) unless I say
otherwise today.

Build v25.5 now, one commit per item, in this order:
 1. Named supervisors (bundle item 9): SUPERVISOR_USERS, RequireSupervisor = on the list or an
    admin and writable, IsSupervisor, ViewModeFor's fourth answer + the viewmode cases, the ui_sup_
    bar prefix, DeskOpenLanding for supervisors only.
 2. Supervisors may publish (item 8).
 3. Click-details on the calendar, no hover boxes (item 0); Needs Attention without notes (item 1).
 4. The legacy writers and SmokeTest steps 19-20 out (item 2); the three admin award buttons greyed
    for a supervisor (item 6); the Help sheet (item 7).
 5. The call-out session reworked (briefs/v25.5.md item 1).
 6. UI_SCALE on every pop-up (item 2), then the plain-language pass (3), shortcuts + what's-new (4),
    the Desk refinements (5), undo-last (6), find-a-person (7).
After each item: the four checkers. Before the tag: read-only reviewers in isolation worktrees (opus),
fix, re-review the fixes; then INSTALL-v25.5.md, briefs/steps-v25.5.txt, CHANGELOG, CLAUDE.md, every
stamp to v25.5, FORM_RELEASES / THISWB_RELEASE for whatever changed, tag v25.5 (annotated), build the
zip, and give me the push command for my Terminal panel. Ask every decision as a multiple-choice
question with the recommended option first; never git add -A. Update docs/NEXT-SESSION-PROMPT.md at
the end of the session.
```

## Why this file exists
A fresh session starts with none of the conversation that produced it. CLAUDE.md holds the rules and the map; this page holds the hand-off: what shipped, what is pending from Graeme, and the next work in order. Update it at the end of every session that changes the state.

## The members' side: alignment with the week's changes (Oct 18)
What the members' file is: values pushed by the master (`PublishMembers`: dMeta, dRoster, dNext, dCal, dPost, dPostList, dLog, dAttend), privacy-screened in code, with its own front page (Find me, the month grid, My postings, Sign me up, My history, sick-time counts). It carries no rule and never names who logged or awarded anything.
- **Unchanged by v25.1-v25.3:** every published table, the privacy screens, the front page. The Desk and the supervisor rights are master-side only.
- **The gap the rights change opened, and its answer:** a supervisor can post and award, but publishing was admin-only, so the members' file would sit behind until an admin published. **Decided Oct 18: supervisors may publish** - `briefs/v25.2.5.md` item 8 (`PublishMembers` on `RequireSupervisor`, the bar's Member View for supervisors, the Desk's members' tile live for a supervisor, the close-time reminder for supervisors too). In v25.5.
- **The M3 follow-up** (deep review 4.5) is the members' own ease-of-use list: withdrawals and login mismatches reported at LOAD, the PC name in the sign-up file name, `Kind` in dLog, `Exempt` in dPostList, `EstHrs` with `~`, the month picker, the phone pattern; plus an "Am I next?" tile and a large-type mode. In v26 (the beta build) as one members' install.
- **At go-live:** `SIGNUPS_ON` (the inbox the award pop-up reads), the members' folder and its rights, `PUBLISH_FOLDER` moved off ".".
