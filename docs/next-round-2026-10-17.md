# The next round, reviewed Oct 17, 2026

> **Superseded on Oct 18 by `docs/timeline-2026-10-18.md`:** the builds move up (v25.5 the alpha build by Oct 31 with the bundle folded in; v26 the beta build by Nov 15 with M3); the order below is kept for the record.

Where things stand: the real roster is on v25.1.1 (v25.1.2 measurement installed the same day); v25.2 (calendar speed) is tagged and waits for its install and the "after" numbers. The numbering rule from today: features x.y as one install, bundles x.y.5, no single-fix installs. Fixed dates: the December rollover dry run (first week of December, Graeme with his boss), the install freeze Dec 20 - Jan 4 except the rollover, the live rollover Dec 27, no installs the week before Jan 23, 2027. Go-live is not set.

## What is queued, by source
| Source | Item | State |
|---|---|---|
| v25.2 testing | the "after" numbers; anything steps 19-20 find | in flight |
| `briefs/v25.2.5.md` | the bundle: Needs Attention without hover notes (Graeme, Oct 17), the legacy writers + steps 19-20 out, C6 if cheap, v25.2 findings | collecting |
| `docs/supervisor-desk-proposal-2026-10.md` (Oct 15) | the Desk view for the second admin: 4 TODAY tiles + 14 verb-phrase buttons at 12 pt, 5 tabs, 120 % zoom, "Show everything"; `DESK_USERS`; the "Something went wrong" email draft; the pop-up `UI_SCALE` pass later | proposal approved in principle, waits for "the current version revisions" - which are now done |
| `docs/rollover-one-pager.md`, v23 `RolloverYear` | the December dry run on a copy with the boss, then the live rollover Dec 27 | nothing built since v23; a rehearsal guide and a practice copy are all it needs |
| kit brief 07 + deep review 4.6 | v26 reports: employee history (admin + member-safe), "why was I charged" (status bar + a Refusals Report column, D-v26), two-week board print, payback aging (highlight + oldest first), the IdMark-vs-archive integrity line, backup housekeeping (prune "Backup before ...", folder size on SystemCheck), `LastReason` on the Attendance sheet; minimum staffing waits for the minimums | not started |
| deep review 4.5 | M3 members' follow-up: the template rebuild (withdrawals / login mismatches at LOAD, `COMPUTERNAME` in the sign-up file name, `Kind` in dLog, `Exempt` in dPostList, `EstHrs` with `~`, `"@"` in WriteTable, dMeta, the month picker, the phone pattern) | not started; the members' file is in use, so this is its own install of the template |
| deep review 4.8 | the Payroll # as the key in every table (last before go-live; needs column A filled by a Kronos import and a go-live date) | not started |
| deep review 4.9 | go-live: folder topology, no VBA password (decided), `ADMIN_USERS`, P: paths, `Signups\` rights, where `Roster error log.txt` goes, `CC_*` / `ADMIN_USERS` on a settings sheet | needs the date |

## The order I recommend
1. **v25.2 on the real roster** (now): the SmokeTest box and the ShowCalDrawTimes lines; the numbers into the CHANGELOG. Anything found goes into v25.2.5.
2. **The Desk view** as the next feature (v25.3; the Payroll # key moves to v25.4 - it is "last before go-live" anyway). Reason: the boss is the second admin and sits with Graeme at the December dry run and the Dec 27 rollover; a screen he can read by then is worth more than any report. Brief written from the proposal (`briefs/desk.md`), report-back with the decisions already taken on Oct 15, then built as one install before Dec 20.
3. **The rollover rehearsal** (late November): a one-page rehearsal script for the two of them on a practice copy (make the copy, roll it, read Year End, check IntegrityCheck, close without saving), refreshed from the v23 one-pager against v25.2's code. No code unless the rehearsal finds something; that goes into v25.2.5.
4. **v25.2.5** (early December, before the freeze): the bundle as collected, plus the rehearsal's findings.
5. **v26 reports** (January, after the Jan 23 freeze week): the kit 07 list with 4.6's additions, one install. "Why was I charged" as decided (status bar + Refusals Report column, no PDF pack).
6. **M3 members' follow-up** (with v26 or right after): the template rebuild, one members' install.
7. **v25.4 Payroll # key, then go-live** (when the date is set): the key last, then the settings sheet and the paths.

## Open before any of it (ask, don't guess)
- Go-live date and the members' folder: still not set; items 6-7 wait on them.
- Minimum staffing numbers per unit and shift (v26 item 2): without them the item stays out.
- The Desk view's `DESK_USERS` logins and the boss's PC display scaling (125 % was suggested Oct 15).
