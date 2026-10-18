# Ease of use: the improvement plan (Oct 18, 2026)

Where things stand: v25.3 (the two Desks, supervisor rights) is tagged; the calendar draws in 2.5 s and goes to about 0.9 s with click-details in the next bundle; the roster's rules are stable and tested (575 known-answer cases, a 21-step smoke test). The system is correct. What is left to earn is that every person who touches it does their job with the fewest clicks, the least reading and no vocabulary of ours in the way. This plan is by user, then packaged by the numbering rule (features x.y, bundles x.y.5, no single-fix installs) around the fixed dates (dry run first week of December, freeze Dec 20 - Jan 4, rollover Dec 27, freeze the week before Jan 23).

## 1. Who uses it and what they do most
| User | How often | The jobs that matter | What gets in their way today |
|---|---|---|---|
| Supervisors (8-10, one at a time) | every shift | log a call-out (the big one: a phone list, outcomes one by one), sick call-ins, post and award overtime (new in v25.3), read the calendar | the call-out pop-up is dense (8.5-9 pt, many fields); outcomes are a click each; summaries are long boxes; system words ("Load Posting", "Convert imports", "NCN") |
| The second admin (the boss) | a few times a week | the same jobs plus time off, switches, leave, publish | reading small type; finding the job among 26 buttons - the Desk answers this from v25.3; the pop-ups are still small |
| The primary admin (Graeme) | daily | everything, plus installs, Kronos, integrity, the rollover | repetition (Kronos by hand from a PDF), knowing what changed between releases, chasing supervisors for what went wrong |
| Members (the union) | weekly | am I next, what is posted, sign me up, my history, my sick-time counts | a separate workbook on plant PCs only; sign-ups on paper until the inbox is switched on at go-live |

## 2. The opportunities, with size and risk
Size: S = under a day, M = a few days, L = a release of its own. Risk: data (touches rules or the log) / UI (looks only) / form (a pop-up re-paste, re-checked by eye).

### A. Supervisors and the boss (the roster)
| # | Opportunity | What changes | Size / risk | Earns |
|---|---|---|---|---|
| A1 | **Click-details on the calendar** (decided Oct 17) | one click on a line shows its details in the hint row and the status bar; no hover boxes | S / UI | the month draws in ~0.9 s; one gesture for details |
| A2 | **Pop-up scale pass (`UI_SCALE`)** - the proposal's step 2 | 12 pt labels, 11 pt lists, buttons 30+ pt tall in the three daily pop-ups first (Sick Call, Callout Session, Award), then the rest; the two 8 pt texts in Callout Session to 10 pt | M / form | the boss and anyone over forty reads them without leaning in |
| A3 | **The call-out session in fewer clicks** | the person being called shown big (name, phone, hours, "why they're next"); one key per outcome (A accepted, R refused, L left message, U unavailable) that also moves to the next name; a "call back later" pile; the elapsed time since the first call; the submit summary one short line with the detail on the status bar | M / form (the rules stay: `CalloutCharge`, `CalloutOffers` untouched) | the job a supervisor does under pressure at 3 AM gets faster and harder to fumble |
| A4 | **Plain-language audit of every pop-up and message** | the bar got verb phrases in v25.3; the pop-ups and boxes still say "Load Posting", "Convert imports", "Submit", "NCN", "S/S"; one pass with a glossary line where a plant term must stay | S-M / form (text only) | fewer "what does this mean" calls to Graeme |
| A5 | **Keyboard shortcuts** | `Application.OnKey` on open: Ctrl+Shift+S sick call, Ctrl+Shift+C call-out, Ctrl+Shift+K calendar, Ctrl+Shift+D the Desk; listed on the Desk's help | S / UI | the regulars never reach for the mouse |
| A6 | **"What's new" once per release** | the first open after an install shows the release's three lines for that user (a hidden name remembers the last version seen; the text from a WHATS_NEW constant) | S / UI | nobody discovers a change by surprise (v25.3's new rights, for instance) |
| A7 | **Help inside the file** | a `Help` sheet drawn from the quick guide (headings + one paragraph each), the Desk's "How do I ..." jumps to it; `GUIDE_PATH` becomes optional (a PDF for printing) | S / UI | help that is never "not on this PC yet" |
| A8 | **Desk refinements from first use** | the cover tile names the next uncovered shift ("Tue Oct 21 Nights, Phosphine") not only a count; a "today's shifts" line in the header (crews on days / nights); a "Suggest an improvement" button beside "Something went wrong" (the same email draft) | S / UI | the Desk answers "what, when, where" at a glance; a channel for the next round of this plan |
| A9 | **Undo the last thing I did** | a Desk button that voids the newest OT Log entry made by this login today (through the existing void path, with its confirm) | S / data (void only, no new write path) | the most common correction without finding the Entry ID |
| A10 | **Find a person** | a search box on the roster (a small shape + InputBox): jumps to the row and shows their place in the call order for the unit on the status bar | S / UI | faster than sorting and scrolling for one name |
| A11 | **Electronic sign-ups on** (`SIGNUPS_ON`, built in v24.1, dormant) | the members' file drops a request file per click; Award Posting pre-ticks them | S to switch on; needs the members' folder (go-live) | the paper sheet stops being typed into the award pop-up |
| A12 | **Daily Kronos import** (`ImportKronosData`, built; waits for the daily export) | one click finds the newest export in `KRONOS_PATH` and imports it; a Desk tile says "Kronos data through <date>" in red when old (already) and "Import now" when a newer file is there | S once the export exists / data (verified import) | the hours typed from a PDF by hand today stop being typed |

### B. The primary admin
| # | Opportunity | What changes | Size / risk |
|---|---|---|---|
| B1 | A morning tile on the admin Desk: integrity result, backup taken, members' file age, error rows unseen | reads what SystemCheck / IntegrityOnOpen already know | S / UI |
| B2 | The install panel asks for the "after" numbers itself | the ShowCalDrawTimes lines go into the Install Log row automatically | S / UI |
| B3 | Backup housekeeping (v26 item) | prune "Backup before ..." copies, folder size on SystemCheck | S / data (deletes copies: verified list first) |
| B4 | A usage count per button | a hidden sheet counts macro runs per login-less bucket (no names): which jobs are used, which never; drives A2-A4 priorities | S / UI (privacy: counts only) |

### C. Members (the members' workbook)
| # | Opportunity | What changes | Size / risk |
|---|---|---|---|
| C1 | M3 follow-up as listed in the deep review 4.5 (withdrawals at LOAD, the file name with the PC, dLog Kind, dPostList Exempt, EstHrs with ~, the month picker, the phone pattern) | the template rebuild, one members' install | M / UI |
| C2 | An "Am I next?" front tile and large-type mode | the front page says the next shift the member is listed for and their place; a zoom button | S / UI |
| C3 | Sign me up live (with A11) | the button writes the request file; the award pop-up reads it | S at go-live |
| C4 | Beyond Excel: a phone-readable view | Track 2 (Microsoft Forms / Power Apps / SharePoint lists), the Q2 2027 decision gate; the tables are already list-shaped for it | L / platform |

### D. Process (how the plan keeps earning)
| # | Opportunity |
|---|---|
| D1 | A 30-minute sit-with session with the boss on his first week of the Desk, and with two supervisors on the call-out pop-up; the notes become the A3 / A4 brief |
| D2 | The "Suggest an improvement" button (A8) and the usage counts (B4) feed each bundle |
| D3 | Every pop-up change ships with a hand-check row on the panel (the pop-up opened pre-filled from the panel, as v24.4 does) so the look is checked on the plant PC, not in the repo |

## 3. Packaged by the numbering rule
> **Re-dated Oct 18 (`docs/timeline-2026-10-18.md`):** the bundle and v25.5 together are the alpha build by Oct 31; v26 with M3 is the beta build by Nov 15; the table below keeps the contents, the dates are the timeline's.

| Release | When | Holds |
|---|---|---|
| **v25.3.5** (the bundle) | before Dec 20 | A1 click-details; Needs Attention without hover notes; the legacy writers and steps 19-20 out; the three admin award buttons greyed for a supervisor; A5 shortcuts; A6 what's-new; A7 the Help sheet; whatever v25.3's install finds |
| **v25.5** "the pop-ups" (a feature) | January, after the Jan 23 freeze week | A2 `UI_SCALE` on the three daily pop-ups then the rest; A3 the call-out session; A4 the language audit; A8 Desk refinements from D1; A9 undo; A10 find |
| **v26** reports (planned) | January-February | employee history, "why was I charged", the board print, payback aging, B3 housekeeping, B1 the morning tile |
| **M3** members (planned) | with v26 or after | C1, C2 |
| **go-live** | when the date is set | A11 sign-ups on, A12 the daily import, C3, the paths and folders, v25.4 the Payroll # key |
| B2, B4 | any bundle | small, no user-facing change |

## 4. What this plan does not do
No rule changes (charging, call order, who is free, the year, privacy) ride on an ease-of-use release; each stays its own decision with its own known-answer cases. Nothing here needs a new platform before the Q2 2027 gate. Every pop-up change is a form re-paste, so A2-A4 ship together and are checked on the plant PC through the panel's hand-check rows.

## 5. Decisions taken (Graeme, Oct 18, 2026)
1. The bundle takes A7 (the Help sheet) only; A5 (shortcuts) and A6 (what's new) go to v25.5.
2. v25.5 starts with A3, the call-out session rework, then the scale pass (A2) on the other pop-ups.
3. No sit-with first: the A3 brief is written from the code and the proposal (`briefs/v25.5.md`), corrected by Graeme, and adjusted after the first week.
