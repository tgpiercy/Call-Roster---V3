# Brief v26: reporting

**Target:** March 2027 onward. Build each item as its own small update.

| # | Item | Notes |
|---|---|---|
| 1 | **Employee history report** | One person + date range → PDF: every call (place in call order, outcome, charge), postings (signed up / awarded / charged), adjustments with reasons, who logged it and when. Admin version includes Logged By; member-safe version (from M1 data) doesn't. Answers most union questions in one click. |
| 2 | **Minimum staffing check** | Setting table `tblMinStaff` (Unit, Shift, Minimum). Flag a unit/shift in Needs Attention and on the calendar date when scheduled - booked off - open lines < minimum. |
| 3 | **Two-week board print** | Calendar print for today → +13 days, one page per calendar group, landscape. |
| 4 | **Payback aging** | Needs Attention highlights owed switches older than `PAYBACK_WARN_DAYS`; sorted oldest first. |

## Not in scope here
- Who's next (now in M1).
- Microsoft Forms sign-ups and the Power Apps pilot (Track 2, decision gate Q2 2027).

## Acceptance
- History report for a person matches their OT Log rows and the member file's My history.
- Staffing flag appears for a planted under-staffed shift on a copy.
