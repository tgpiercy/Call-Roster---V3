# v23 report-back (before any code)

Brief: `briefs/v23.md` (kit brief 04: IntegrityCheck, RolloverYear, docs set). Read against the code at tag `v22.1` on Oct 12, 2026. Nothing is built yet.

## What the code already gives us
| Brief item | What I found | Effect on the plan |
|---|---|---|
| "Rebuild the pay-period columns for 2027" | The whole year hangs off **one typed cell**: roster `T2` (first period start). `U2:AS2` are `=RC[-1]+14`, row 4 is `=R2C+13`, every period cell is a SUMIFS on that column's start date, and the provisional formula, the leave top-up window, the members' publish and `PeriodStartFor` all read the same header. | The rollover's core step is writing `T2 = Dec 27, 2026`. Nothing else on the roster is rebuilt, which keeps the risky night small. |
| "Entries already logged with 2027 period starts appear immediately" | True for free: SUMIFS match on Period Start, so mid-December postings for January shifts show the moment `T2` changes. | No code. A SelfTest / dry-run check. |
| "Roster = log" check | `LogTotalOnRoster` and `IsOurFormula` / `IsBrokenFormula` exist (v20/v21). | IntegrityCheck reuses them; the per-cell compare is one pass over the log, not a SUMIFS per cell. |
| Paybacks owed | `OwedSwitches` exists. | Add `PAYBACK_WARN_DAYS = 42`. |
| Verified-change pattern | `UpgradeOTLogLayout`, `UpgradeAbsenceColumns`, `SetupHomeUnit`: backup, change, compare, `MarkNoSave`. | RolloverYear follows it exactly. |
| Year-end questions already answered | Reset point = pay-period year (Dec 27, 2026). Rollover run by Graeme with his boss. Everything local for now. | Not re-asked. `ARCHIVE_FOLDER = "."` (next to the roster), same convention as `BACKUP_FOLDER`. |

## Findings that change the plan
1. **Removing 2026 rows from the log would restart the ID numbering.** `NextRefID` and `NextEntryID` take the highest number found in `tblOTLog`. Empty the table and the next call-out is `S-0001`, Entry ID 1, colliding with the archived year. `LeaveToppedUp` (v22) also reads the live log, so a leave topped up in December would be offered for a second top-up in January. If rows are removed, both need fixing (ID high-water marks kept in the workbook; the check reads the archive too).
2. **Removal isn't needed for the totals.** Because of the header design above, 2026 rows simply stop matching any column once `T2` changes. The brief's step 3 ("the one sanctioned removal") is then a speed and tidiness choice, against CLAUDE.md's append-only rule. Decision 1 below.
3. **A 27-period year can't fit the roster as built.** Columns `T:AS` are 26 fixed columns (`C_PERIOD1`..`C_PERIODN`), used by sorts, filters, formats and the verify code. The first 27-period year is **2033** (2026-2032 and 2034-2035 are all 26; worked out from the Dec 28, 2025 anchor). Decision 2 below.
4. **Worked hours across the boundary.** After the rollover, column H still holds 2026 Kronos hours until the first January export. The brief's default stands: H is zeroed and the status line says "awaiting 2027 Kronos"; an import whose data-through date is before Jan 1, 2027 is refused for H. Two side effects to accept or fix:
   - **Hours History needs the Year column** (brief step 5) for a real reason: a snapshot taken Dec 27-31 would otherwise be read as 2027 hours by the leave top-up. `PickSnapshot` will use the Year column when present.
   - **OT worked Dec 27-31** sits in period 1 of 2027 but, if Kronos year-to-date restarts on Jan 1, never appears in 2027's H. Provisional accepted OT covers it until the January import, then it drops out. Small (five days), but real. Decision 3 below.
5. **`REPORT_FOLDER` has the year typed into it** (`...\Production Overtime Roster\2026\`). Posting PDFs and reports would keep landing in the 2026 folder. Plan: the folder name's year follows the roster year (`{YEAR}` in the setting), flagged in SystemCheck when the folder is missing.
6. **Late entries for the closed year.** Logging a Dec 20 call-out on Dec 28 already prompts "not one of this roster's columns". Kept as is; IntegrityCheck lists such rows as information, and the one-pager says the frozen file is the record.
7. **Members' history.** After the rollover the members' file shows 2027 rows only (it publishes "this year"). The frozen 2026 file is the record; noted in the one-pager. A members' 2026 archive is out of scope.
8. **Dry run with a forced date.** The macro refuses to run before the new year's first period starts (built stricter than the brief's "last period's start date": rolling mid-period would strand that period's entries). For the December dry run it allows an early run only when the workbook's file name contains `COPY` or `DRY RUN`, and says so in every message, so the live file can't be rolled early by accident.
9. **Docs.** `CHANGELOG.md` can only be exact from v20 (the first commit). v15-v19 get a summary paragraph taken from `docs/UPDATE-20.md`.

## Build order
| # | Item | Notes |
|---|---|---|
| 1 | Pure year functions + cases | `FirstPeriodStart(year)`, `RosterYearOf(start)`, `PeriodsInYear(year)` in modOvertimeEntry, mirrored in `tests/rules.py`; cases for 2026, 2027, 2033 (27) and 2034. |
| 2 | `modIntegrity` (new) | `IntegrityCheck` (report sheet, Alt+F8) and `IntegrityQuick` (on open, admin only, status line "Integrity: OK / n issues"). The eight checks in the brief; each returns rows `(check, severity, what, where)`. Pure parts (duplicate IDs, typed value, orphan names) take arrays so SelfTest can plant a bad row in memory, never on a sheet. |
| 3 | `modRollover` (new) | `RolloverYear`: guard (admin, writable, date, not already rolled) -> IntegrityCheck must pass -> frozen copy `Call Out Roster 2026 FINAL.xlsm` (read-only attribute set) and each person's final G / H / I recorded on a hidden `Year End` sheet -> log step per Decision 1 -> `T2` to the new year -> H zeroed, Kronos name reset -> Hours History Year column -> verify (new G = 2027 log rows + provisional; recorded finals = frozen copy; row counts) -> `MarkNoSave` on any mismatch -> status line + Needs Attention line. |
| 4 | Kronos import | Refuse H for a file dated before the new year's Jan 1 after a rollover; stamp the Year column. |
| 5 | Docs | `system-manual.md`, `supervisor-quick-guide.md`, `admin-install-guide.md`, `rollover-one-pager.md` (written for two people: Graeme drives, his boss reads the checks aloud), `CHANGELOG.md`. |
| 6 | Reviews, zip, tag `v23` | Two independent read-only reviews, fixes, re-review. `N_MODULES = 17`. ThisWorkbook re-paste (the on-open check). |

## Decisions needed
1. **2026 log rows at rollover:** archive to `OT Log 2026` and remove from the live log (brief), or keep every year in the one log.
2. **27-period years:** refuse to roll into one with a clear message (first is 2033), or build the 27th column now.
3. **OT worked Dec 27-31:** wait for the first January export before deciding, or build a carry-over now.
