# Call Out Roster v29.0.2 - install (the Install panel) - THE ALPHA'S MOCK FILE ONLY

v29 with two changes for the alpha of Nov 1-14, 2026, both on the MOCK file only (the real roster is set aside under
`Pre-alpha\` first - docs/alpha-data-reset.md page 1 - and stays on v29). It supersedes v29.0.1, which was never
installed.

1. **`KRONOS_PATH`** = `C:\Users\piercyg\Desktop\Kronos mock\` (the local folder the mock Kronos reports go into).
2. **Act as** (Graeme, Oct 28: "so I can do the testing myself without interference"): Alt+F8 > `ActAs`, type the
   password, pick a chair - `lanthierr` (Ray's Lite Desk), `mihaychukc`, `sheltond` or `campeauc` - then close and
   reopen the roster: it lands on that chair's Desk with that chair's rights, and every row it logs says that login.
   Alt+F8 > `ActAsClear` + reopen = you again. SystemCheck shows "ACTING AS ..." while a chair is set. Only your own
   admin login may set or clear it; it stops working after Nov 20, 2026 and the release after the alpha removes it.
   The members' file has the same: Alt+F8 > `MemberActAs` (any login name you like, e.g. `tester2`) so you can send a
   switch request as one member and Agree as another; `MemberActAsClear` puts the PC's login back. A freshly published
   members' file starts with no chair.
   **The password is `alpha2026`** (`CHAIR_PASSWORD` in modOvertimeEntry and `MEM_CHAIR_PASSWORD` in member/modMember.bas; change both to
   change it).

Needs **v29 installed** and **Roster Updater.xlsm v4.2**. The updater does not change. The panel's swap replaces every
module (all 45 stamps move to v29.0.2); ThisWorkbook and the forms are not touched; the members' template is rebuilt by
the panel's template step (the members' code changed; its layout and version stamp "Members M6" do not).

**Order on Oct 31 (Day 0 of docs/alpha-test-plan.md):** the real file set aside FIRST (page 1 of the reset runbook),
then a practice copy, this install on the copy, then on the file itself, then `LoadMockData` (the test-only module,
below), then the Day-0 Kronos import.

## The panel
1. Open the roster (ADMIN view) from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v29.0.2` folder. Groups: **1. Prepare**, **2. Install code**, **3. Run once** (the members' template, System check, Self test), **4. Checks**.
3. Click the one blue button, read its box, click the next.
   - **System check:** "Modules: 45 of 45 on v29.0.2"; the paths box names the Kronos mock folder as reachable.
   - **Self test:** 1140 passed, 122 suites (the new suite `actas` = the Act-as rule).
4. **Show the log** opens the roster's Install Log sheet.

By hand instead: remove and import all 45 stamped modules from `modules/`; leave the six unstamped modules, ThisWorkbook
and the forms alone; Debug > Compile; in the members' TEMPLATE remove modMember and import member/modMember.bas, save;
run the group 3 and 4 rows yourself.

## The test-only module (not installed by the panel)
`test-only/modMockData.bas` in the zip is the alpha's mock-plant loader. File > Import it by hand into the mock file
ONLY (never into the real roster, never into `modules/`); SystemCheck does not count it. Alt+F8 > `LoadMockData`
(type `MOCK` on the live-named file) does the rest - docs/alpha-data-reset.md page 1, step 8; the rehearsal on a
practice copy first is plan step D-1.1.

## Then
Tell me "v29.0.2 OK" with the Install Log's rows, the paths box's Kronos line and the ACTING AS line after a first
ActAs, or the box of the step that failed.
