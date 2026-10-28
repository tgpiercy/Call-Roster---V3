# Alpha findings (Nov 1-14, 2026, on v28)

The one log of what the alpha turns up. One row per finding, added at triage (no code changes at triage). Current as of v28 on the real roster (Oct 27, 2026; v27.5 inside).

## How a finding gets here
- **Sources:** a Ref ID from the OT Log, a receipt's "Was this right?" answer of No / Not sure (`tblFeedback`), a Reconcile difference, an error box's saved line (`Roster error log.txt`), SystemCheck's first line when something looked wrong.
- **Triage (the P1 prompt):** find the code path and quote file:line; say whether the roster or paper was right against CLAUDE.md's canonical rules (or "rule missing"); give a severity and a release; for an S2, write the void-and-re-enter the admin does now.
- **Paper governs during the alpha.** The roster's figure is not corrected in Kronos or on paper; the roster is.

## Severity key
| Severity | Means | Fixed now by | Code fix in |
|---|---|---|---|
| **S1** | The roster is broken: a save fails, a sheet will not load, a rule charges or lists the wrong person with no way round it | A **v28.0.x hotfix** through a practice copy (the one install allowed during the alpha). Stop and put the hotfix choice to Graeme as multiple choice | v28.0.x |
| **S2** | Wrong hours, charge or call order, with a way round it | The admin voids and re-enters (`VoidRef`, then the entry again); the row says exactly what | **v29** |
| **S3** | Confusing, slow or cosmetic | Nothing now | **v29** or later (v29.5) |

## Findings
| Date | Ref ID | Who | What happened | Severity | Decision | Release |
|---|---|---|---|---|---|---|
| | | | | | | |

## Close-out (Nov 16-17)
The cumulative Kronos report lands -> Import Kronos -> IntegrityCheck -> ClosePeriod for Nov 1-14 on v28 -> the Reconcile sheet's differences -> the Stats report -> the Feedback rows -> every row above triaged. The pack's evidence is captured on the release that ran the alpha (v28), before v28.1 goes on.

After the close-out, every S2 / S3 row marked "v29" becomes a v29 item in the part-B report-back (Nov 18) and `briefs/v29.md`.
