# Brief 00: verify v20 and name the master

**Why:** code was built in more than one session on Oct 9. Three fixes delivered that morning may not be in the v20 zip. Nothing else gets built until there's one master.

**Input:** `Callout_Roster_v20.zip` (required) and, if available, the separate Oct 9 morning files (`frmShiftSwitch.txt`, posting/award files).

## Steps
1. Unzip v20 into `modules/` and `paste-in/`. `git init`, commit as "v20 as delivered".
2. Check each item below in the code. Report present / missing with the procedure name and line.

| # | Fix | Where to look | Present when |
|---|---|---|---|
| 1 | Out-of-dept drop-down shows a **total hours** column, sorted **lowest first** (ties by seniority), heading shows how many qualify | `frmPostingEntry`, `modPostings` | List built with hours and sorted ascending |
| 2 | Printed posting blank lines ask for **name and initials only** (no hours) | `modPostings` posting draw | `POSTING_OUTSIDE_LINES` lines with Name / Initials only, note "awarded only if no one in the department signs up" |
| 3 | Two-way swap: payback date **blocked** unless the payer is free (not scheduled, rest rule, not booked off, not on a switch) and the person paid back is scheduled; same-day sister swap checks booked-off only | `frmShiftSwitch`, `modShiftSwitch` | Save disabled + red "NOT POSSIBLE - ..." hint |
| 4 | Holidays feed `ScheduledShift` **and** `SwitchPickList` | `modShiftCycle`, `modShiftSwitch`, `modHolidays` | DAYS not scheduled on an observed holiday in both paths |
| 5 | Form uses a `PersonFlag` **code**, not message text | `modShiftSwitch`, `frmShiftSwitch` | No string matching on message wording |
| 6 | Oct 9 Shift Switch form fixes: typed "Jan 5" in Dec → next year; unit change clears names; payback hint shows hard stops; lists rebuild only on date change; void clears form; Enter doesn't save | `frmShiftSwitch` | Each behaviour in code |

3. Merge anything missing **into the v20 files** (keep v20 version stamps, note it in CHANGELOG as "v20 master merge").
4. Independent review of every merged file (compile traps in CLAUDE.md).
5. Package `dist/Callout_Roster_v20-master.zip` (full set) + `INSTALL-v20.md` (copy the v20 install steps; add any extra re-paste caused by the merge).
6. Commit; tag `v20-master`.

## Acceptance
- A table of the six checks with evidence.
- One zip labelled master. SystemCheck in that zip expects every module on v20.

## Smoke test for Graeme (after install)
- SystemCheck: "all on v20"; only Boxing Day and Kronos age listed.
- Log Call-Out for **Mon Oct 12 Days**: free DAYS people listed; Calendar shows "Thanksgiving".
- Award pop-up out-of-dept box: lowest hours on top.
- Shift Switch two-way with an impossible payback: Save blocked.
