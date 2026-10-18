# Call Out Roster v25.2 - install (the Install panel)

The calendar speed-up (`briefs/v25.2.md`), built on the v25.1.2 numbers from the real file (Calendar October 2026: 4.6 s, of which the hover notes 2.8; Needs Attention 1.4 s). One install; what its testing turns up is batched into v25.2.5.

Needs **v25.1.2 installed** and **Roster Updater.xlsm v4.1**. **The updater does not change.** The panel's swap replaces every module (all 24 stamps move to v25.2; the code changed in modCalendar and modSmokeTest); ThisWorkbook, the forms and the members' template are not touched. No rule changed: SelfTest 568 / 67 as before.

**Back up first. Do it on a copy before the real roster.**

## What changed
- **Hover notes only where action is needed.** OPEN, NO TAKERS and POSTED lines keep their hover box; a covered line or an absence has none - double-click it > **Details...** for the same text (that menu item was already there). 85 notes were 2.8 of the 4.6 s.
- **The page setup is done once**, not on every draw (0.6 s): only when the month's row count moved or the release changed.
- **The month is written in blocks**, one write per week for the texts and one for the click map, with the colours and fonts applied per class instead of per cell (0.7 s of cell-by-cell writes). What you see is the same grid, cell for cell: SmokeTest proves it (below).
- **Needs Attention** writes each section as one block the same way, and stops scanning the roster for every name's home unit.
- **SmokeTest has two new steps (19 and 20)** that draw the Calendar and Needs Attention both ways - the new writer and the old cell-by-cell one, kept for this one release as `DrawCalendarLegacy` / `DrawNeedsLegacy` - and compare every cell (text, fill, font, alignment, merge, note), every row height and every click-map key, failing on the first difference. They go again in the release after v25.2.
- **Expected**: the Calendar draw from 4.6 s to about 1 s on the same month; Needs Attention from 1.4 s to about 0.5 s. The two "Draw times after" rows read the real numbers off your file; they go in the CHANGELOG next to the before.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.2` folder.
3. Click the one blue button, read its box, click the next. Group 4: the practice copy + SmokeTest (20 steps, about two minutes), the two draw-time rows, a Print / PDF look.
   - **System check:** "Modules: 24 of 24 on v25.2". **Self test:** 568 passed, 67 suites.
   - **Smoke test:** SMOKE TEST PASSED, 20 steps; the lines of steps 19 and 20 carry both draw times (new and old).

By hand instead: remove and import `modCalendar` and `modSmokeTest` (SystemCheck then reads 2 of 24 on v25.2 until the panel runs), Debug > Compile, then the group 4 rows yourself.

## If the Calendar ever looks wrong
Alt+F8 > **DrawCalendarLegacy** redraws the month the old way (and DrawNeedsLegacy the Needs Attention tab); send me a photo of both. Nothing else to do: the next open draws it the new way again.

## Then
Send the SmokeTest box (steps 19 and 20 lines) and the two ShowCalDrawTimes lines.
