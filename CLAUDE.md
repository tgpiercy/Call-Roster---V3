# Call Out Roster: guide for Claude Code

## What this is
This repo holds the VBA source for the Call Out Roster workbook. The workbook handles overtime call-outs, postings, shift switches and the coverage calendar for the production units at the Syensqo Niagara Falls plant. Supervisors run it in desktop Excel on Windows. Graeme, a shift supervisor, owns it.

**This repo is the single master for the code.** Nothing is installed from anywhere else, and every release is built from a tagged commit.

## Layout
```
modules/     standard modules, one .bas per module (file name = VB_Name)
paste-in/    code-behind for ThisWorkbook and the forms (.txt, pasted with Ctrl+A)
briefs/      one brief per update (v21.md, v22.md ...): what to build
docs/        planning docs, install notes, manual, guides
tests/       Python mirrors of the rules with known-answer cases (from v21)
CHANGELOG.md
dist/        built zips (gitignored)
```

## Never commit data
- No workbook files: `*.xls`, `*.xlsx`, `*.xlsm`, `*.xlsb`.
- No Kronos exports, posting PDFs or plant CSVs.
- They hold names, phone numbers, emails and leave codes. `.gitignore` blocks them; don't override it.
- Test fixtures use made-up names only.

## This environment can't run Excel
Excel and VBA don't run here, so every change is checked by reading:
- **Read before changing.** Never assume a column letter, sheet name, table name or procedure name. Grep for it.
- **Independent review.** Every update gets a review by a subagent that hasn't seen the work being written. It checks:
  - `Option Explicit` is present and every variable is declared;
  - blocks are matched (`If`/`End If`, `For`/`Next`, `With`/`End With`, `Select`/`End Select`);
  - every procedure called from a button, a form or another module still exists with the same signature;
  - there are no duplicate Public names across modules;
  - no code path writes to the roster or OT Log outside the existing write paths.
- **Python mirrors.** Pure rules (rotation, holidays, DST, call order, charging, eligibility) are mirrored in `tests/` and checked against the same known answers as the workbook's SelfTest.

## File format (the VBA editor is picky)
- Use CRLF line endings for `.bas`, `.cls`, `.frm` and paste-in `.txt` files. `.gitattributes` enforces this.
- **Use ASCII only, in code and in strings.** The VBA editor is ANSI and mangles curly quotes, en/em dashes, arrows and the single ellipsis character. Write `->`, `-` and `...` instead.
- Every `.bas` file starts with `Attribute VB_Name = "..."`. Paste-in files have no Attribute lines.
- Every module keeps `Option Explicit`.

## Versioning and releases
- Each module has a version constant that SystemCheck (`modSystem`) reads. Follow the existing pattern exactly.
  - Bump it in every module you change.
  - Update SystemCheck's expected version for the release.
- One update means:
  - one brief;
  - one tag (e.g. `v21`);
  - one zip, `dist/Callout_Roster_v21.zip`, containing the changed files plus `INSTALL.md`;
  - one CHANGELOG entry.
- **Install freezes:** no releases Dec 20 - Jan 4 except the year-end rollover, and none in the week before Jan 23, 2027.

## Rules: the system's spec
Don't change any of these unless a brief says so.

**Call order.** Sort by grade code, then Total Hours (G), then plant seniority (J, earliest first), then name. There is one implementation, used by:
- the call-out form;
- the posting list;
- the printed roster;
- the member publish.

**Total Hours** is Kronos worked hours plus logged refusal charges, plus provisional accepted OT from v21. Totals reset each new year (rollover in v23). Pay periods are two weeks starting on Sundays; the first 2026 period starts Dec 28, 2025.

**Charging.** Only grade 2 and below is ever charged. Out-of-department people are never charged. A charge is hours x 1.75.

| Event | Charged |
|---|---|
| Call-out | Refused or Left Message. Goes in the pay period of the shift date. |
| Posting award | Anyone ahead of the last awarded person in call order who didn't sign up. Not: the awarded person, anyone who signed up but wasn't awarded, or anyone below. |
| No takers | Everyone eligible on the list. |
| Award out of dept | In-department people who didn't sign up. |
| Cancel | Nobody. |

Posting charges go in the shift date's pay period. Postings compare hours using the totals snapshotted when the posting was created.

**OT Log** is the source of truth and is append-only.
- A correction is a void (the row stays, struck through) plus a new entry.
- Roster pay-period cells are formulas over the log.
- Never write hours into roster cells directly.

**Rotation**
- 28-day 12-hour cycle. Day 1 is Sun Dec 28, 2025 (B days / A nights).
- Nights start at 7 PM and are dated by the evening they start.
- A/B and C/D work their pair's day shifts only.
- DAYS work Monday to Friday days.

**Eligibility.** A person is free for a shift when they are:
- not scheduled;
- not coming off nights that morning (for day OT);
- not starting days the next morning (for night OT);
- not booked off;
- not working, or having given up, a shift switch that day.

On plant holidays (the observed date), DAYS workers are off and free. A/B, C/D and the A-D crews are unchanged.

**Privacy**
- STD, parental and LOA show only as "Leave" everywhere outside the Absences sheet.
- Leave notes are never shown.
- The member publish leaves out phone numbers, emails, leave types, void notes and who logged each entry.

**Admin**
- Admin access is by Windows login, listed in `ADMIN_USERS`.
- Absences, Needs Attention and Holidays are admin-only and hidden in saved copies.

## Engineering rules
- **One place per rule.** If you find a second copy of call-order or eligibility logic, consolidate it instead of patching both.
- **Names are an interface.** Public procedure names that buttons call (OnAction), and macro names in install notes, don't change. If a rename is unavoidable:
  - ApplyRosterLayout must rebuild the buttons;
  - the install note must say so.
- **Columns.** Find them by header text wherever the code already does that.
- **No column inserts without an upgrade macro.** Never insert a column into the roster or OT Log except through a macro modelled on `UpgradeOTLogLayout`. It must:
  1. back up;
  2. move the data;
  3. verify that every roster period total and the log's total Hours Charged are unchanged;
  4. block saving if they aren't.
- **Bulk changes.** Any macro that changes data in bulk saves a backup first, verifies afterwards, and blocks saving on a mismatch.
- **Tables stay flat and list-shaped:** `tblOTLog`, `tblPostings`, `tblPostingList`, `tblAbsence`, `tblCoverage`, `tblShiftSwitch`. They're meant to map onto SharePoint Lists later.
- **Calendar change tracking.** Writes to Absences, Coverage, Postings or Shift Switches go through the existing change-marking path, so the calendar redraws.
- **Messages.** Every message tells a supervisor what to do next, in plain words.

## Install notes
Graeme installs by hand in the VBA editor. `INSTALL.md` in each zip follows the existing style, as short numbered steps:
1. Back up first.
2. **Remove** each changed module (answer No to export), then **File > Import** it.
3. Re-paste each changed form or ThisWorkbook: open its code, Ctrl+A, paste.
   - For a new form, say so: Insert > UserForm, set (Name), leave it blank, then paste the code.
4. **Debug > Compile VBAProject.** It should finish with no message.
5. Run any one-time macros, giving each Alt+F8 name.
6. Run **SystemCheck** and say what result to expect.
7. A short smoke test.

List only the files that changed.

## Working with Graeme
- Ask questions as multiple choice (AskUserQuestion), not in plain text.
- Start each update by reading the brief and every module it touches. Report anything that changes the plan, and wait for a go-ahead before writing code.
- If a brief conflicts with this file, ask.
