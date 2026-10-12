# Call Out Roster: code repo

VBA source for `V2 - Call Out Roster.xlsm`. Context for Claude Code is in `CLAUDE.md`; each update's spec is in `briefs/`.

## Set up (once)
1. Create a folder (or private GitHub repo) and copy this kit into it.
2. Unzip `Callout_Roster_v20.zip` into `modules/` and `paste-in/`.
3. Open Claude Code in the folder.

## Kick-off prompt for each build session
> Read CLAUDE.md, then briefs/<brief>.md. Check every (verify) item against the code first and tell me what you found. Ask me any open questions as multiple choice. Then build, run an independent review of every changed file, and package dist/Callout_Roster_vNN.zip with INSTALL-vNN.md.

## Order
`00` verify v20 → `01` M1 member workbook → `02` v21 → `03` v22 → `04` v23 → `05` v24 → `06` v25 → `07` v26
(M1 and v21 can run in parallel sessions only if they don't touch the same modules; `modOvertimeEntry` and `modUI` overlap, so do them one after the other.)
