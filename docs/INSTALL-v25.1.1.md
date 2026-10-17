# Call Out Roster v25.1.1 - install (the Install panel)

v25.1 plus one SmokeTest fix from Graeme's Oct 16 run: the full set, tagged, so the real roster takes one zip. **If v25.1 is already on a copy and you re-imported modSmokeTest by hand on Oct 16, that copy already runs this code; only its SystemCheck line will read "1 of 24 on v25.1.1" until the panel runs there.**

Needs **v25 or v25.1 installed** and **Roster Updater.xlsm v4.1** (the one v25 put there). **The updater does not change: nothing goes in by hand.** The panel's swap replaces every module (all 24 stamps move to v25.1.1); ThisWorkbook, the forms and the members' template are not touched.

**Back up first. Do it on a copy before the real roster.**

## What changed since v25 (`CHANGELOG.md` has the detail)
- v25.1 (Oct 16): the Calendar's unit list raised "Application-defined or object-defined error" - the Calendar and Needs Attention sheets set their tab colour on every call under v25's tab protection, and the hidden CalMap sheet was written bare; both now go through the unlock / protect pairs. SmokeTest step 4 expected v24.1's partial-acceptance rule; it asserts v25's (charged 21: the shift closes at the partial).
- v25.1.1 (Oct 17): SmokeTest step 9's second plan is dated the day after the newest crew change (step 7's, a week out) instead of tomorrow; the roster rightly refused the earlier date. The four plan / cancel refusals now put their reason in the red box. Graeme's Oct 16 copy passed all 18 steps on this code.
- No rule changed: SelfTest 568 / 67, the Python mirror 581, the four checkers at baseline.

## The panel
1. Open the roster (ADMIN view) from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.1.1` folder. Groups: **1. Prepare**, **2. Install code**, **3. Run once** (System check, Self test), **4. Checks** (practice copy + SmokeTest, the Calendar's lists).
3. Click the one blue button, read its box, click the next.
   - **System check:** "Modules: 24 of 24 on v25.1.1", the Protection line.
   - **Self test:** 568 passed, 67 suites.
   - **Smoke test** in the practice copy: SMOKE TEST PASSED, 18 steps. CLOSE THE COPY WITHOUT SAVING, then Done.
4. **Show the log** opens the roster's Install Log sheet.

By hand instead: remove and import all 24 stamped modules from `modules/`; leave the six unstamped modules, ThisWorkbook and the forms alone; Debug > Compile; run the group 3 and 4 rows yourself.

## Then
Tell me "v25.1.1 OK" with the Install Log's rows, or the box of the step that failed.
