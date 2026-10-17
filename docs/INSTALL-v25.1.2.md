# Call Out Roster v25.1.2 - install (the Install panel)

Measurement only, before the v25.2 calendar speed-up (`briefs/v25.2.md`, item 0): the Calendar and Needs Attention draws time themselves, so the "before" numbers come off the real file. **Nothing drawn changes except the hint under the Calendar's title, which now ends "Drawn in N s."** No rule changed.

Needs **v25.1.1 installed** (the real roster, Oct 17) and **Roster Updater.xlsm v4.1**. **The updater does not change.** The panel's swap replaces every module (all 24 stamps move to v25.1.2; only modCalendar's code changed); ThisWorkbook, the forms and the members' template are not touched.

**Back up first** (the daily backup counts). It is small, but the usual way: on a copy first if you like, or straight on the real roster since nothing it draws or saves changes.

## What changed
- `DrawCalendar` and `DrawNeeds` read the clock between their phases and keep the breakdown in the hidden CalMap sheet. The Calendar's hint cell ends "Drawn in N s.". Alt+F8 > **ShowCalDrawTimes** (admin) shows both breakdown lines: Calendar = caches, wipe, header, weeks (slots / notes / cells), page, controls; Needs Attention = gather (slots), rows, tiles.
- Those two lines, from the busiest month and the longest Needs Attention range, are what I need back: they decide where the v25.2 work goes first and are the "before" the CHANGELOG quotes.

## The panel
1. Open the roster (ADMIN view), then File > Open `Roster Updater.xlsm` in the same Excel window. Click the roster so it is in front.
2. Alt+F8 > `'Roster Updater.xlsm'!InstallPanel`. Pick the unzipped `Callout_Roster_v25.1.2` folder.
3. Click the one blue button, read its box, click the next. Group 4 has the two draw-time rows: do what each says and put the line in the note (or photograph the box).
   - **System check:** "Modules: 24 of 24 on v25.1.2". **Self test:** 568 passed, 67 suites.

By hand instead: remove and import `modCalendar` only (SystemCheck will then read 1 of 24 on v25.1.2; fine for a copy), Debug > Compile, then the two draw-time rows yourself.

## Then
Send the two ShowCalDrawTimes lines (a photo of the box is enough). v25.2 is built from them.
