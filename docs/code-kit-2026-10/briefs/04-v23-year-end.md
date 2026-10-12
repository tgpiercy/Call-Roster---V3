# Brief v23: year-end rollover, integrity check, docs

**Build by Nov 30. Dry run on a copy the first week of December (run by the backup admin from the one-pager). Live Dec 27.** Install freeze Dec 20 - Jan 4 applies to everything except this.
**Depends on:** v21 SelfTest, v22.

## Scope
| # | Item |
|---|---|
| 1 | `IntegrityCheck` (build first; the rollover uses it before and after) |
| 2 | `RolloverYear` |
| 3 | Docs set + one-page rollover guide |

## 1. IntegrityCheck
Runs quietly on open (admin only, result on the status line: "Integrity: OK" / "Integrity: 3 issues") and from Alt+F8 with a full report sheet.

| Check | Fail when |
|---|---|
| Roster = log | Any person's period cell, I or G differs from what the log + H + provisional gives |
| IDs | Duplicate Entry ID; duplicate session/posting/switch/LV/SC IDs |
| Orphans | Log employee not on the roster; absence/switch employee not on the roster |
| Postings | Charges exist but posting not closed; closed as Award/No takers with zero charges where the list had eligible grade ≤ 2 people who didn't sign up |
| Calendar | Coverage line pointing at a void posting or switch, or a posting ID that doesn't exist |
| Paybacks | Owed switch older than `PAYBACK_WARN_DAYS` (default 42) |
| Typed values | Any roster period cell that isn't a formula |
| Kronos | Data older than `KRONOS_STALE_DAYS` |

## 2. RolloverYear
**Year definition:** pay-period year. Period 1 is the biweekly period containing Jan 1 (2026: Dec 28, 2025; 2027: **Dec 27, 2026**). Handle 26- or 27-period years generically.

Steps (one macro, admin only, refuses to run before the last period's start date unless forced on a copy):
1. Run IntegrityCheck; stop if it fails.
2. `SaveCopyAs` the frozen year file: `Call Out Roster 2026 FINAL.xlsm` in `ARCHIVE_FOLDER`. Record each person's final G, H, I.
3. Copy 2026 log rows (period start in 2026's year) to a new sheet **OT Log 2026** (values, same columns), then remove them from `tblOTLog`. This is the one sanctioned removal: verify row count and Hours Charged total in archive = removed.
4. Rebuild the roster pay-period columns for 2027 (headers start / to / end, current-period highlight, formulas over the log). Entries already logged with 2027 period starts (postings with 14-day closes, from mid-December) appear immediately.
5. Hours History: keep; add a Year column if missing. Grade Changes, Absences, Coverage, Postings, Switches are **not** reset.
6. Verify: new roster G = H(as imported) + 2027 log rows only; frozen copy totals = recorded finals. Block save on mismatch.
7. Status line + Needs Attention: "Rolled over to 2027 on <date> by <user>".

**Kronos worked hours across the boundary:** H comes from Kronos year-to-date. Between Dec 27 and the first January import, H still holds 2026 worked hours. Default: after rollover, H is **zeroed and shown "awaiting 2027 Kronos"** until an import with data through ≥ Jan 1; refusals Dec 27-31 count in 2027 per the pay-period year. (Question below.)

**Open postings across the boundary:** postings created in 2026 keep their frozen 2026 totals for award charging (existing rule: posted totals decide). Note it in the one-pager.

## 3. Docs (in `docs/`, also written to the project)
| Doc | Audience |
|---|---|
| `system-manual.md` | Current rules only (supersedes the old guides' rules) |
| `supervisor-quick-guide.md` | 1-2 pages: call-out, post, award (incl. out of dept), switch, time off, sick call, void request, close the file |
| `admin-install-guide.md` | Install pattern, SystemCheck, SelfTest, backups, ADMIN_USERS, trusted location |
| `rollover-one-pager.md` | Step-by-step for the backup admin, with what to do if verify fails (close without saving, restore backup) |
| `CHANGELOG.md` | v15 → v23 |

## SelfTest additions
Period-1 calculation for 2026, 2027, and a 27-period year; IntegrityCheck catches a planted typed value and a planted duplicate ID on the scratch sheet.

## Acceptance
- Dry run on a copy with today's date forced to Dec 27, 2026: integrity OK before and after, frozen totals match, December-logged January posting charges visible on the 2027 roster.
- Backup admin completes the dry run from the one-pager without help.

## Questions (multiple choice, default first)
- After Jan 1, what does the Kronos export's worked hours show: **2027 year-to-date only** / still cumulative 2026 / not sure (check first January export).
- Reset point for refusals: **pay-period year (Dec 27)** / calendar Jan 1 (would split period 1).
- Who runs the rollover and where `ARCHIVE_FOLDER` is (free text).
