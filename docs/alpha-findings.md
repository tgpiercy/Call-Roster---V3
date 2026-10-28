# Alpha findings (Nov 1-14, 2026, on v29)

The one log of what the alpha turns up. One row per finding, added at triage (no code changes at triage). Current as of v29 on the real roster (Oct 28, 2026; Graeme pulled the v29 install ahead of the alpha - "we will alpha on V29").

## How a finding gets here
- **Sources:** a plan step id from `docs/alpha-test-plan.md` (the P9 mock testing plan: a figure that differs from its EXPECT), a Ref ID from the OT Log, a receipt's "Was this right?" answer of No / Not sure (`tblFeedback`), a Reconcile difference, an error box's saved line (`Roster error log.txt`), SystemCheck's first line when something looked wrong.
- **Triage (the P1 prompt):** find the code path and quote file:line; say whether the roster or paper was right against CLAUDE.md's canonical rules (or "rule missing"); give a severity and a release; for an S2, write the void-and-re-enter the admin does now.
- **Paper governs during the alpha.** The roster's figure is not corrected in Kronos or on paper; the roster is.

## Severity key
| Severity | Means | Fixed now by | Code fix in |
|---|---|---|---|
| **S1** | The roster is broken: a save fails, a sheet will not load, a rule charges or lists the wrong person with no way round it | A **v29.0.x hotfix** through a practice copy (the one install allowed during the alpha). Stop and put the hotfix choice to Graeme as multiple choice | v29.0.x |
| **S2** | Wrong hours, charge or call order, with a way round it | The admin voids and re-enters (Fix an entry on the Lite Desk, or `VoidRef` then the entry again); the row says exactly what | **v29.5** (or v28.1's slot) |
| **S3** | Confusing, slow or cosmetic | Nothing now | **v29.5** or later |

## Findings
| Date | Ref ID | Who | What happened | Severity | Decision | Release |
|---|---|---|---|---|---|---|
| | | | | | | |

## Close-out (Nov 16-17)
The cumulative Kronos report lands -> Import Kronos -> IntegrityCheck -> ClosePeriod for Nov 1-14 on v29 -> the Reconcile sheet's differences -> the Stats report -> the Feedback rows -> every row above triaged. The pack's evidence is captured on the release that ran the alpha (v29), before anything else goes on.

After the close-out, every S2 / S3 row becomes an item of the next release's report-back (v29.5, or v28.1's slot if Graeme folds them in) with its brief.
