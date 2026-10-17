# Call Out Roster v25.1 - install (the Install panel)

Two findings from Graeme's Oct 16 test of v25, fixed on top of v25. Install it the same way, through the panel; it is small.

Needs **v25 installed** on the roster and **Roster Updater.xlsm v4.1** on your PC (the one v25 put there). **The updater does not change: nothing goes in by hand.** Full module, paste-in, member and updater set in the zip, plus `install-steps.txt` at its root (the panel reads it). The panel's swap replaces every module (all 24 stamps move to v25.1); ThisWorkbook, the forms and the members' template are not touched (the folder's paste-ins are v25's: replacing them changes nothing).

**Back up first. Do it on the copy you tested v25 on before the real roster.** Close the practice copy where SmokeTest failed WITHOUT saving (saving is blocked there anyway) and make the next practice copy from the roster you install v25.1 on.

## What changed (brief: this note; `CHANGELOG.md` has the detail)
- **The Calendar's unit list raised "Application-defined or object-defined error" (Error Log: CalUnitChanged).** v25 keeps every sheet protected, the hidden CalMap sheet included, but the calendar code still wrote to that sheet bare (the month / unit / range lists, the "something changed" mark, the drawn-from stamps) while every other sheet is written between unlock and protect. Now every CalMap write does that too, and the draw unlocks the map for its duration. The "something changed" mark was the quiet one: a refused write there would have left the calendar never redrawing after a booking.
- **Tab colours.** The Calendar and Needs Attention sheets set their tab colour every time they are touched, and new sheets set theirs right after the tabs are relocked; v25 protects the workbook's tabs, and a tab colour is a change to them. One helper now sets a tab colour only when it differs, with the tabs unlocked for that moment, and never raises.
- **SmokeTest step 4 failed "S-0004 charged 21, expected 28".** The roster was right: v25's rule (decided Oct 14) closes the shift at a partial acceptance, so the person below is offered 0 h and charged nothing: 12 h refused x 1.75 = 21. The step still expected v24.1's rule (the 4 h left offered down the list). The step now asserts the v25 rule, and step 5 proves the calendar's "something changed" mark is really written.
- No rule changed: SelfTest stays 568 / 67; the Python mirror 581.

## The panel
1. Open the roster (ADMIN view) from its folder, then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.1` folder. The Install sheet fills: **1. Prepare** (backup, check), **2. Install code** (swap, compile, finish + reopen), **3. Run once** (System check, Self test), **4. Checks** (practice copy + SmokeTest, the Calendar's lists, the Error Log).
3. Click the one blue button, read its box, click the next. A FAILED step blocks the rest: **Retry** or **Skip this step...** with a note.
   - **System check:** **"Modules: 24 of 24 on v25.1"**, the Protection line, the forms line unchanged from v25.
   - **Self test:** 568 passed, 67 suites.
   - **Smoke test** in the practice copy: SMOKE TEST PASSED, 18 steps; step 4's line reads "charged 21 h (0 h offered below the partial)". CLOSE THE COPY WITHOUT SAVING, then click Done.
   - **The Calendar's lists:** the row says what to click; each pick redraws with no error box.
4. **Show the log** opens the roster's Install Log sheet.

By hand instead (if the panel can't run): remove and import all 24 stamped modules from `modules/` as the v24.3.2 note lists them, plus `modSmoke`; leave the six unstamped modules, ThisWorkbook and the forms alone; Debug > Compile; then run the group 3 and 4 rows yourself.

## Then
Tell me "v25.1 OK" with the Install Log's rows (a photo is enough), or the box of the step that failed. Two things from the Oct 16 run I still want: the Error Log's line 3 (its error number and Detail) and whether lines 1 and 2 are from the same open - Alt+F8 > ShowErrorLog in ADMIN view, a photo of the rows.
