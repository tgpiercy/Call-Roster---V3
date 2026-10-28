# Brief: Admin Lite Desk (ships in v29; first written Oct 25, 2026, updated Oct 27)

> **Repo note (Oct 27, 2026, later the same day; Graeme's decision when v29 was pulled ahead):** v29 is built on **v28** (v28 is on the real roster since Oct 27), and v28.1 (the second refactor) comes AFTER v29, built on it. So wherever this brief says "built on v28.1" or "unchanged from v28.1", read **v28**. The text below is Graeme's brief as sent, unchanged.

> **Oct 27 update:**
> - This now ships in **v29**, the go-live kit, built on v28.1. It was briefed for "v28" before v28's scope changed.
> - Members' requests shipped in v28, so **the requests card is live from day one**.
> - Ray's login is already in code: `lanthierr` is on `ADMIN_USERS` and `DESK_USERS`. L1 is answered.
> - "Unchanged" below means unchanged from **v28.1**.

## Goal
Ray has full admin rights but opens to a simpler Desk. It shows only his jobs, laid out in a fixed order, with one action per card. Everything else is one click away on the full admin Desk. **Lite changes what's on screen, never what he's allowed to do.**

Mockup: the "Admin Lite Desk" design canvas, with two artboards: the Desk itself, and "What goes where".

## Decided (Oct 25, Graeme)
- **Ray is an admin** (`ADMIN_USERS`). He doesn't get a reduced-rights role.
- **Ray's decision jobs:**
  - member requests;
  - time off, leave and needs;
  - posting exceptions (no takers, cancel, award out of department);
  - corrections (void and re-enter, adjustments).
- **Ray's oversight jobs:** staffing and OT reports; people changes; backup for the daily admin check.
- **How it grows with him:** he always opens on the Lite Desk, and one button shows the full admin Desk.
- **It ships with the go-live kit (v29).** It's display-only. The supervisor Desk preview set the pattern in v27.1.
- **Admins approve requests.** There is no `APPROVER_USERS`, as built in v28.

## How it works
| Piece | What |
|---|---|
| `LITE_DESK_USERS` (new settings constant) | `lanthierr`. The login must also be on `ADMIN_USERS`. SystemCheck flags a login that's on this list and not an admin. |
| `DrawDesk` | Picks the layout: supervisor if not admin (or previewing supervisor); otherwise lite if the login is on the lite list and the session isn't showing the full Desk; otherwise full. |
| `gDeskShowFull` (session flag) | The header button **"Show full admin Desk"** reads **"Back to my Desk"** while the full Desk is showing. The flag resets on every open, so Ray always lands on Lite. It's never saved. |
| Admin preview | The v27.1 "See supervisor Desk" button becomes **"See Desk as: Supervisor / Admin Lite"**, so Graeme can see exactly what Ray sees. |
| Rights | Unchanged. Every button on the Lite Desk calls the same macro as its counterpart on the full Desk, and `IsAdmin` is untouched. |

## The layout: four bands, always in the same places
Cards keep their positions even when empty ("Nothing waiting", greyed), so Ray learns where things are.

**Header**
- "Call Out Roster - Desk";
- an **ADMIN LITE** pill;
- "Signed in as Ray - admin rights";
- the Show full admin Desk button.

**Trust strip.** Four items, each OK or red:
- hours through `<Kronos date>`, plus N provisional;
- today's backup;
- members' file updated `<time>`;
- system checks.

Any red item reads "call `<PRIMARY_ADMIN_NAME>`". This strip replaces the full Desk's system detail for Ray.

**Band 1: Needs your decision.** Oldest first, one verb per card.

| Card | Count comes from | Button opens |
|---|---|---|
| Member requests | `tblRequests` Waiting (plus "Waiting for partner" shown as a second number) | Needs Attention's MEMBERS' REQUESTS section, where a double-click books (the prefilled pop-up) and a right-click declines |
| Posting exceptions | Postings past closing with nobody signed up, closing within 24 h with nobody signed up, or awaiting an out-of-department award | The Award pop-up on that posting, where No takers, Cancel and Out of dept are offered |
| Corrections asked for | Unresolved "No" Feedback rows, plus open disputes | **Fix an entry** (below) |
| Returning from leave | Needs Attention's Returning list | Leave > Returning, pre-filled |

**Band 2: Book and change.** Four large buttons, each with a one-line hint:
- Book time off;
- Leave (going or coming back);
- Add a need;
- Change a person (new hire, unit move, rename).

**Band 3: Daily check.** Labelled "only when Graeme is off". Three numbered steps, each green when anyone has done it today:
1. Import Kronos;
2. Clear Needs Attention (shows its count, with an Open button);
3. Publish the members' file.

When all three are done, the band collapses to one green line naming who did them and when.

**Band 4: Staffing and OT.**
- Below minimum in the next 14 days: the count and the first shifts (from MinStaff).
- OT hours this pay period by unit, as bars.
- Buttons, **each shown only once the report it opens exists**:
  - Staffing outlook;
  - OT by unit;
  - Posting outcomes;
  - Look up a person (employee history).

## Fix an entry (new, guided; uses the existing void path)
1. Pick a Ref ID from the "No" receipts, or type one in.
2. A read-only view shows every row the Ref wrote and what it changed for each person, in the receipt's layout.
3. **Void.** A reason is required. This uses the existing `VoidRef`, with its backup, verify and NoSave.
4. **"Re-enter it now?"** opens the right pop-up, pre-filled from the voided rows.
5. The receipt shows the before and after, and marks the Feedback row resolved with the new Ref.

## Needs Attention for a Lite login
- Ray's sections come first:
  - members' requests;
  - open lines;
  - returning from leave;
  - minimum staffing;
  - disputes.
- DATA CHECK, unread errors and the full Feedback list are folded under **"More (system)"**. They're still one click away.

## Learning aids
- **A "?" on every card** shows a one-line explanation from a Help table. The same text goes on his cheat sheet.
- **First-week strip** for a Lite login, ticked on the practice copy too:
  1. book time off;
  2. approve a member's request;
  3. resolve a posting exception;
  4. fix an entry;
  5. run the daily check.
- **Paper:** a one-page Admin Lite cheat sheet laid out in the same four bands. Training guide 08 is reorganised around the bands. Ray trains on the v29 practice copy in the Dec 7-11 rehearsal week.

## Tests
- **modSnapshot** (test-only): the full admin Desk and the supervisor Desk are identical to v28.1. A new snapshot of the Lite Desk is taken for the fixed test data.
- **SmokeTest:**
  - a Lite login lands on Lite, can go to full and back, and lands on Lite again after reopening;
  - an admin not on the lite list sees no change;
  - a supervisor sees no change;
  - "See Desk as: Admin Lite" from Graeme's login matches Ray's layout.
- **SelfTest:** each card's count matches its source; an empty card stays in place and greyed.
- **SystemCheck:** a lite login that isn't an admin is flagged.

## Still open (the default applies if not answered)
| # | Question | Default |
|---|---|---|
| L1 | Ray's Windows login | **Answered:** `lanthierr` |
| L2 | Does the daily check band show every day, or only when Graeme hasn't done it? | Every day, collapsed to one line when done |
| L3 | Does Graeme want the Lite Desk for himself on busy days? | No; he can view it through See Desk as |
| L4 | Reports not built by v29 (posting outcomes, OT by unit) | Their buttons stay hidden until built |
