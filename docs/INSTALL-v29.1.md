# Call Out Roster v29.1 - install (the Install panel) - THE ALPHA'S MOCK FILE ONLY

v29 with the alpha's changes, all on the MOCK file only (the real roster is set aside under `Pre-alpha\` first -
docs/alpha-data-reset.md page 1 - and stays on v29). It supersedes v29.0.2 and v29.0.1, which were never installed. The
Oct 28 full review (docs/review-2026-10-28-full.md) found the alpha items and three save / backup HIGHs; Graeme chose to
fix them in this build (docs/v29.1-build-prompt.md). The CHANGELOG entry lists every item with its review number.

**The alpha's three settings (temporary; the release after the alpha puts them back):**
1. **`KRONOS_PATH`** = `C:\Users\piercyg\Desktop\Kronos mock\` (one backslash now) - the local folder on GRAEME's PC the
   mock Kronos reports go into; only Graeme imports. **Regenerate the reports from this build:** `py -3.14
   tests/mock/make_kronos.py` (the Oct 31 row moved to Wren Fenwick) and copy the Day-0 report into the folder.
2. **`REPORT_FOLDER`** = `.\Alpha reports\{YEAR}\` - every PDF the mock file makes (postings, the month calendar, the
   Board) lands in `Alpha reports\2026\` beside it, never in the real P: union folder. The folder is made on first use.
3. **Act as** (Graeme, Oct 28): Alt+F8 > `ActAs`, the password, pick a chair - `lanthierr` (Ray's Lite Desk),
   `mihaychukc`, `sheltond` or `campeauc` - then close and reopen the roster: it lands on that chair's Desk with that
   chair's rights, and every row it logs says that login. **The chair is honoured by YOUR Windows login only** (another
   admin opening the file is themselves). Alt+F8 > `ActAsClear` + reopen = you again; a save that fails is said.
   SystemCheck shows "ACTING AS ..." while a chair is set. It stops working after Nov 20, 2026. The members' file has
   the same for the admins' own logins only: Alt+F8 > `MemberActAs` (any login name, e.g. `tester2`), `MemberActAsClear`.
   **The password is `alpha2026`** (`CHAIR_PASSWORD` in modOvertimeEntry and `MEM_CHAIR_PASSWORD` in member/modMember.bas).
   The award email's CCs are blank until Nov 20 (no mock award drafts to the real superintendent and supervisors).

**The review's fixes (in the code for good - but NOTE: the real roster stays on v29 until its next install, so the
save / backup fixes below reach it only then; the Pre-alpha copy of the real file is frozen by its Read-only attribute,
runbook step 3, not by this build):** a partial acceptance closes a call-out
by the order the outcomes were LOGGED (a call-back who takes part of the hours after refusals were logged leaves those
refusals charged); one `DefaultClosesOn` for New Posting, Post all lines and Edit posting (inside 14 days a posting now
closes the day before the shift, never today); Fix an entry pre-fills a re-entry with the session's hours on offer and
picks a time-off type by its code; the calendar auto-pick counts a NO TAKERS line; the Lite exceptions card reads "with no
electronic sign-up: P-nnnn (closed - check the paper sheet)" and its button line says "Paper sheet first, then award or No takers"; a failed re-entry names the voided Ref and leaves the
feedback unresolved; the Kronos age reads from the report's data; the year-to-date import skips removed people; Bring
back restores Worked (H); the hourly backup is saved in the RESTRICTED view; Save As and a read-only copy's save are
refused; a backup-named copy (a daily / hourly backup, the FINAL, a "Backup before ..." copy) is frozen whether or not it is
read-only - once it carries this build's code; the PRE-ALPHA copy carries v29's and is frozen by its Read-only attribute.

Needs **v29 installed** and **Roster Updater.xlsm v4.2**. The updater does not change. The panel's swap replaces every
module (all 45 stamps move to v29.1), re-pastes **ThisWorkbook** and **four forms** (frmOvertimeEntry, frmNewPosting,
frmTimeOff, frmFixEntry), and rebuilds the members' template (the members' code changed: "Members M6.1"). No new module,
no new form, no run-once macro, no What's New box.

**Order on Oct 31 (Day 0 of docs/alpha-test-plan.md):** the real file set aside FIRST (page 1 of the reset runbook),
then a practice copy, this install on the copy, then on the file itself, then `LoadMockData` (the test-only module,
below; it now refuses the live-named file until the Pre-alpha copy exists beside it), then the Day-0 Kronos import.

## The panel
1. Open the roster (ADMIN view) from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v29.1` folder. Groups: **1. Prepare**, **2. Install code**, **3. Run once** (the members' template, System check, Self test), **4. Checks**.
3. Click the one blue button, read its box, click the next.
   - **System check:** "Modules: 45 of 45 on v29.1"; ThisWorkbook and the four forms on v29.1; the paths box names the Kronos mock folder and `Alpha reports\2026\` as reachable.
   - **Self test:** 1165 passed, 123 suites (the new suite `closeson` = the default closing date; `calloffer` and `actas` grew).
4. **Show the log** opens the roster's Install Log sheet.

By hand instead: remove and import all 45 stamped modules from `modules/`; paste `paste-in/ThisWorkbook.txt` and the four
forms' `.txt` over their code (Ctrl+A first); leave the six unstamped modules and the other forms alone; Debug > Compile;
in the members' TEMPLATE remove modMember, import member/modMember.bas, run `BuildMemberTemplate`, save; run the group 3
and 4 rows yourself.

## The test-only module (not installed by the panel)
`test-only/modMockData.bas` in the zip is the alpha's mock-plant loader. File > Import it by hand into the mock file
ONLY (never into the real roster, never into `modules/`); SystemCheck does not count it. Alt+F8 > `LoadMockData`
(type `MOCK` on the live-named file) does the rest - docs/alpha-data-reset.md page 1, step 8; the rehearsal on a
practice copy first is plan step D-1.1.

## Then
Tell me "v29.1 OK" with the Install Log's rows, the paths box's two lines and the ACTING AS line after a first
ActAs, or the box of the step that failed.
