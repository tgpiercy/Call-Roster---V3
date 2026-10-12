# Call Out Roster: progress review and plan (Oct 12, 2026)

Written while v24.1 is being installed and tested on Graeme's desktop. Answers from Graeme the same day: the system is **in testing only, not live**; next priorities are **v25 as written, reports for the union, and easier installs**; alternatives to compare: **code in an add-in file** and **SharePoint Lists for the data**.

## 1. Where we are
| Release | Theme | State |
|---|---|---|
| v20, v21, M1 | Safety, call-order accuracy, members' workbook | Installed on the desktop copy |
| v22, v22.1 | Sick calls, leave, attendance | Built; v22 install hit two compile errors (fixed), in progress |
| v23 | Integrity check, year-end rollover | Built, not installed |
| v24 | Employee pop-up, crew history, partial OT | Built, not installed |
| v24.1 + M2 | Refused by kind, attendance tab, part call-outs, sign-up inbox (off), members' workbook rebuilt | Built today, not installed |

About 22,000 lines of VBA in 20 checked modules and 9 forms, 320 known-answer cases, four static checkers, and an independent review of every release. Nothing has been compiled by Excel since v21.

## 2. What the review shows
**Working well**
- One rule in one place, mirrored in Python and checked by SelfTest. Every review round this caught real slips.
- Append-only log with verified changes: no release so far has had a path that changes hours without a log row.
- Decisions are written down before code, so builds don't drift.

**The three real risks**
1. **Five releases are stacked uninstalled, and nothing here compiles the code.** The reviews have been good at logic and have still missed compile errors twice. Every release adds to the pile that meets Excel for the first time in one sitting. This is the biggest risk in the project.
2. **Installs are long and manual.** v24.1 is about 40 minutes of Remove / Import / Ctrl+A paste across 20 modules, 4 code windows and the template. A missed paste is invisible until something breaks. This will not scale to 8 - 10 supervisors' releases a year.
3. **Runtime behaviour is checked by reading.** A merged cell that Excel refuses to unlock, a date that isn't "numeric": these were found by reviewers reasoning about Excel, not by running it. My history has been publishing empty since M1 and nobody saw it, because nothing runs the publish end to end.

**Smaller observations**
- Names are still the key in every table. v24's Rename now changes all of them together, so the day-to-day pain is gone; the remaining reason for an employee ID is a future move to another platform.
- Sign-ups, the refused split and sick-time counts all widened what the members' file shows. The privacy self-check in the publish is now doing real work and deserves its own test cases.
- The transition plan (Oct 9) has wave 1 live on Nov 1. With testing still on the desktop and five releases to install, that date needs a decision (section 6).

## 3. What was on the ticket for v25
From the code kit (`docs/code-kit-2026-10/briefs/06-v25-hardening.md`), target February - March 2027:
| # | Item | What it is |
|---|---|---|
| 1 | **Employee ID as the key** | A roster column `Emp ID` (generated `E-0001`, because Kronos exports names only), added to the OT Log, posting lists, absences, coverage, switches, crew history, grade log and Hours History. A verified migration fills it by name, stops on anything unmatched, and blocks the save on a mismatch. Lookups switch to the ID; names stay for display. |
| 2 | **Protection** | Macros-only writes (`UserInterfaceOnly`), admin sheets very hidden, a VBA project password. "Keeps honest people honest": the admin check is a Windows login. |
| 3 | **Calendar speed** | Build the month in arrays, format by group, then redraw only the day that changed. Target: full month 3x faster, one-line change under a second. |

And v26 (March 2027 on): employee history report, minimum staffing check, two-week board print, payback aging.

## 4. Proposals
Ordered by what removes the most risk first. Items 1 - 3 are the "easier, safer installs" priority.

### 4.1 One-click updater (new)
A macro `UpdateFromFolder`: point it at the unzipped release folder and it removes and imports every module, replaces the code behind ThisWorkbook and each form, adds a missing form, compiles, and runs SystemCheck. It needs Excel's "Trust access to the VBA project object model" on the admin's PC (SystemCheck already uses it when it is on).
- An install becomes: back up, run the updater, run the listed one-time macros. About 5 minutes, nothing to paste.
- The manual steps stay in every install note as the fallback.
- The updater cannot update itself mid-run, so it lives in its own small module that changes rarely.
- The members' template gets the same treatment from the master: `UpdateMemberTemplate` opens the template, swaps its code, rebuilds it, saves it.

### 4.2 SmokeTest macro (new)
One macro, run on a copy, that does what the install notes ask a person to do: creates a posting, signs people up, awards it, logs a call-out with a part acceptance, books time off, runs a sick call, adds and removes a test employee, publishes to a scratch template, voids everything it made, and finishes with IntegrityCheck and a before / after comparison of every total. It reports the first step that failed. It refuses to run unless the file name contains COPY or DRY RUN.
- This is the piece that would have caught the empty My history, and it catches runtime errors no review can.

### 4.3 Error log (new)
Today a runtime error gives a supervisor the VBA debug box, and the feedback that reaches me is a photo of it. A shared handler writes procedure, error, user and time to a hidden admin sheet and shows a plain message ("this didn't save; tell Graeme: line 14 of the Error Log"). Needs Attention shows unseen errors.

### 4.4 Try a real compile on this laptop: DROPPED (Graeme, Oct 12: not possible; Claude has no access to the work laptop)
Excel is installed on the machine this session runs on. It starts under automation and can edit a blank workbook, but it is an **unlicensed** install and its VBA project is not reachable until "Trust access to the VBA project object model" is ticked in its Trust Center (I can't and won't change that setting myself). If you tick it and it works, I can import every release into a blank workbook and compile it before you see a zip. If unlicensed Excel refuses, nothing is lost and 4.1 - 4.2 on your work PC carry the load.

### 4.5 v25, reshaped
Keep the three items, change the order and split the risky one:
| Release | Content | Why |
|---|---|---|
| **v25** | Protection, calendar speed, the updater / SmokeTest / error log if not already shipped as v24.2 | Low data risk, immediate benefit |
| **v25.1** | Employee ID migration, alone | The highest-risk change left in the roadmap: it touches every table. It should ship by itself, after a stable month of live use, with the SmokeTest available to prove it. |
**Update, later on Oct 12:** the Kronos report does carry an employee number (`Payroll #`, first column). The key in v25.1 is that number, not a generated ID; v24.2 starts capturing it into the free roster column A at each import, so by v25.1 most rows already have it and the migration only has to copy it into the other tables.

Additions to the ID work that the brief doesn't have: the Former sheet and Year End carry the ID; the sign-up file carries it (so a sign-up no longer depends on spelling); the members' file keys Find me on it internally while showing names.

### 4.6 v26, reports for the union
The brief's four items, plus what v24.1 opened up:
- **Employee history report**: one person, a date range, one PDF: every call with their place in call order, postings (listed, signed, awarded, charged), adjustments with the calculation, crew and unit changes. Admin version has Logged By; the member-safe version doesn't.
- **Refusals report**: add the pay-period range picker left out of v24.1.
- **Two-week board print**, **minimum staffing check**, **payback aging** as briefed.
- **Charge explanation**: from any Refused row, "why was I charged" in one sentence (who was awarded, their rank, your rank, the rule). The data is already in the log.

### 4.7 Smaller refinements worth queuing
- Privacy self-check gets known-answer cases (a planted phone number, email, leave code, sick date must each stop a publish).
- Publish also writes a read-only PDF pack (roster by unit, open postings, this month's calendar) beside the members' file, for anyone without Excel.
- A "Make a practice copy" button that saves a COPY-named file with its own template beside it, so testing never starts from the live file by hand.
- v24.2 as already planned: unit moves entered ahead of time.

## 5. Alternatives compared
### A. Code in an add-in file (.xlam) instead of in the roster workbook
| | One-click updater (4.1) | Add-in file |
|---|---|---|
| What an update is | Run one macro | Replace one file |
| Change to today's code | One new module | Every `ThisWorkbook` reference, every sheet event, every form launch has to work across two files; workbook events must be re-hooked from the add-in |
| Risk | Low | High: a refactor of all 22,000 lines with no compiler here |
| Needs on each PC | Trust access to the VBA project (admins only) | The add-in loaded for every supervisor, on every PC, at the right version |
| Failure mode | Updater fails: fall back to the manual steps | Add-in missing or stale: the roster opens with no code at all |
| Data and code versions | Always together in one file | Can drift apart (old add-in, new data layout) |

**Recommendation:** build the updater. It gets most of the benefit for a fraction of the risk, and keeps the property that the roster file is complete by itself. Revisit the add-in only if IT blocks VBA project access.

### B. SharePoint Lists for the data
| | Stay on tables in the workbook | Lists for the OT Log / switches / sign-ups |
|---|---|---|
| Editing at once | One editor at a time | Several |
| History | Our append-only log and daily backups | Version history on every row, built in |
| Rules | VBA, one place per rule, tested | Still in Excel (or rewritten in Power Apps); a list holds data, not rules |
| Excel reading / writing lists from VBA | n/a | Works but slow and fragile; every roster formula over the log has to become a refresh |
| Needs | Nothing new | IT: list creation on the site; licences if Power Apps is used; the employee ID (v25.1) first |
| What it solves that hurts today | n/a | Lock-outs at shift change; phones |

**Recommendation:** keep the Q2 2027 decision gate from the transition plan, and make the first list the **sign-up inbox**, not the OT Log. It is already a drop-box of independent rows, it has no formulas hanging off it, a wrong row can't change anyone's hours, and members could reach it from a phone. If that pilot is easy, the switches table is next. The OT Log moves last, if ever: the whole roster is formulas over it.

Both alternatives depend on the same preparation already in the roadmap: flat tables and the employee ID.

## 6. Proposed sequence
| When | What | Notes |
|---|---|---|
| **Now** | Install **v24.1's full set in one sitting** on the desktop copy | Nothing is live, so there is no reason to step through v22.1, v23, v24 separately: the v24.1 zip is every module. Run the one-time macros in order: `UpgradeAbsenceColumns` (if not done), `SetupPeople`, `ApplyRosterLayout`; rebuild the members' template; SystemCheck 20 of 20, SelfTest 320. Send me whatever Compile stops on. |
| Now, 5 min | Tick "Trust access to the VBA project object model" in the Excel on this laptop and tell me | Lets me test whether I can compile here (4.4) |
| After your feedback | **v24.2**: fixes from the install, planned unit moves, the one-click updater, SmokeTest, error log | The install kit ships before anything else is added |
| Before go-live | Decide the go-live dates again | The Oct 9 plan has wave 1 on Nov 1. Suggest: go live only after one clean SmokeTest on the desktop and one shadow pay period. |
| Dec, first week | Rollover dry run with your boss; live Dec 27 | Unchanged |
| Jan 5 - 15, 2027 | Any release held for the freeze | No installs the week before Jan 23 |
| Feb 2027 | **v25**: protection, calendar speed | |
| Feb - Mar 2027 | **v26**: employee history report first, then the rest | Each as its own small update |
| Mar 2027, after a stable month | **v25.1**: employee ID migration | Alone |
| Q2 2027 | Platform gate: sign-up inbox as a SharePoint List pilot | Only if IT allows lists on the site |

## 7. Decisions for Graeme
1. Install v24.1 in one sitting on the desktop copy (recommended), or step through each release.
2. v24.2 content: install kit + planned moves together (recommended), or the install kit first by itself.
3. Split the employee ID out of v25 into v25.1 (recommended), or keep v25 as written.
4. New go-live dates.
5. Whether to try the compile on this laptop.
