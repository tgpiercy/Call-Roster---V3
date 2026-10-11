# Call Out Roster: transition review and plan (Oct 9, 2026)

Two tracks:
1. **Go-live:** moving the plant off the old process onto the new workbook. The old process is the paper posting, the old roster/Refusal Report, and the 4-5 department vacation calendar workbooks.
2. **Platform:** where this lives long-term. Power Apps is out for now but will be explored later.

Decided Oct 9: plan for both tracks, and all shift supervisors run the workbook at go-live.

Read with `deep-dive-review-2026-10-09.md` (what to fix), `build-plan-2026-10.md` (v20-v26) and `member-workbook-design.md` (the members' file).

**Changed Oct 9 (later):** the member workbook moves up from v26. It's built right after v20 is verified and goes live with wave 2 on Nov 15. This plan adds the people, process and cutover steps around those builds.

---

## 1. Current status

### What's built
| Area | State |
|---|---|
| **OT Log** (source of truth) | Append-only log with voids, Ref IDs, period start, who/when. Roster totals are formulas over the log plus Kronos worked hours. |
| **Call-outs** | Pop-up with Days/Nights and auto crew. The list is everyone with the unit code who's free that shift (rest rule, booked off, switches, holidays from v20). Accepted people go onto the calendar. |
| **Postings** | Created from the roster or calendar, with the closing date defaulting to 14 days before the shift. The printed sheet has the list plus 3 blank out-of-department lines. The Award pop-up offers Award / Out of dept / No takers / Cancel, with rank-based charging and emails CC'd by unit. |
| **Shift switches** | Owed list with paybacks, two-way and sister swaps, hard stops. Name lists are filtered by unit and availability, and the payback date is validated. |
| **Coverage calendar** | One sheet for all units, with privacy-masked leave, hover notes, double-click menus and change-tracked redraw. |
| **Needs Attention** | Open lines, postings, paybacks owed, leave to watch, holidays to confirm (v20). |
| **v20** (built Oct 9, other session) | Stat holidays, Kronos age, Hours History, SystemCheck, read-only notice, AutoSave off, Outlook fallback. |
| **Docs** | 10 project docs. The early ones partly describe older rules (see deep-dive 3.7). |

### Not done yet / open
| Item | Status |
|---|---|
| **Single master copy** | **Risk.** Code has been built in more than one session on Oct 9. Three fixes from this morning (out-of-dept hours column sorted lowest first, name-and-initials blank lines, two-way payback availability check) were delivered before v20 was built. Confirm v20 includes them before installing (Phase 0). |
| **Testing** | Every update is checked in code and compiled by you; nothing is run before it reaches you. SelfTest (v21) is the first automated check. |
| **Imported calendar lines** | Display only, so imported vacations don't keep people off call lists (deep-dive 1.1). |
| **Other units' calendars** | Only the Phosphine/Derivatives calendar is confirmed imported. Cypure, Fumigation, QC and Utilities/SBP still need importing or re-booking. |
| **Decisions owed** | Boxing Day 2026 weekday; DST charging (scheduled vs actual hours); home unit per person; who runs the rollover; "Sick" vs "Off" on posted calendars; members on desktop or phones. |
| **Continuity** | Only you can run updates or fix problems. |

---

## 2. Platform review (Power Apps excluded for now)

### What any platform must do
These are what make the system work, and what a replacement has to reproduce:
- Call order (grade > total hours > seniority), applied identically everywhere.
- Rotation and eligibility: the 28-day cycle, rest rule, A/B and C/D, DAYS, holidays, switches, booked off.
- Charging rules: the 1.75x rank-based refusal on awards, no takers, out of department, cancel.
- An append-only audit log with voids.
- Printed postings and calendars, plus award emails.
- Admin-only leave detail, with a public "Leave" mask.
- Usable by about 8-10 supervisors, around the clock, one or two at a time.

### Options
| Option | Licence / IT ask | Fit | Verdict |
|---|---|---|---|
| **A. Excel desktop .xlsm master** (current), hardened | None | Runs all the logic today. One editor at a time; depends on macros staying allowed. | **Recommended now.** |
| **B. Excel Online + Office Scripts** | Office Scripts must be enabled for the tenant | Browser use and co-authoring, but every VBA routine and all 5 pop-ups would be rewritten in TypeScript. Scripts can't show forms the same way. | No. A full rewrite for less than A. |
| **C. SharePoint Lists as the data store, Excel as the front end** | List creation on your site (no Power Apps) | Lists give version history and per-row permissions. VBA writing to lists is fragile, and the logic still lives in Excel. | Not now. Becomes the data layer if Power Apps is approved (Track 2). |
| **D. Microsoft Forms for sign-ups** | Forms is usually on for everyone | Members could sign up for postings by QR/link instead of paper. Getting responses into the master needs Power Automate or a manual import. | Later, as an add-on to A. |
| **E. Access front end** | Access installed on every supervisor PC | Better integrity, but still a single file, and likely not deployed. | No. |
| **F. Custom web app** | IT project, hosting, support | Best possible fit; unrealistic under current restrictions. | No. |
| **G. Power Apps + SharePoint Lists** (future) | About 8-10 licence requests (only people who log need one; viewers don't) | The natural long-term home: multi-user, phones, audit history. | Keep the path open (Track 2). |

### What could force a move
| Risk | Signal | Mitigation in place / needed |
|---|---|---|
| **Macro policy tightens** | IT blocks VBA, or files from SharePoint/email open with macros disabled | Open the master from a fixed supervisor location (mapped drive or synced folder), never from email or browser downloads. Ask IT to make that folder a **Trusted Location**. |
| **Outlook automation ends** (new Outlook) | Award email fails | v20 mailto/clipboard fallback. |
| **Contention** | Supervisors locked out at shift change | v20 read-only notice + the "close after use" habit (section 4). |
| **Key-person dependency** | You're away or move on | Backup admin trained; rollover guide; docs (Phase 5, Jan). |

### Recommendation
Stay on the Excel desktop master through 2027 Q1. Make it **platform-ready**, so a later move is a data migration rather than a rebuild:
- **Employee ID** as the key on every table (v25).
- Every table stays flat and list-shaped (`tblOTLog`, `tblPostings`, `tblPostingList`, `tblAbsence`, `tblCoverage`, `tblShiftSwitch`). They map one-to-one to SharePoint Lists.
- Rules stay in one place per rule (call order, eligibility, charging), and SelfTest documents them as known-answer cases. Those become the acceptance tests for any future platform.

---

## 3. Where the files live

| File | Location | Who |
|---|---|---|
| **Master `.xlsm`** | One supervisor-only folder: a mapped network drive or a synced SharePoint library. **Not** the `Union\...\Production Overtime Roster` folder, because the master holds leave codes and phone numbers. | Supervisors (edit), Ray (read) |
| **Backups** | A daily dated copy, plus version history if it's on SharePoint. Keep 30 days. Done by hand until an automatic "first open of the day" copy is added (proposed for v21). | Admins |
| **Member workbook** `Call Out Roster - Members.xlsm` (values only, pushed by the master's Publish button) | The union/read-only folder. Path to confirm. | Members, union |
| **Posting and calendar PDFs** | The existing report folder. | Everyone |
| **Code master** | One zip per version (`Callout_Roster_vNN.zip`) in the project plus a supervisor folder. Nothing is installed from anywhere else. | You + backup admin |

**Needs your confirmation:** where the master sits today. If it's under the Union folder now, move it before go-live.

---

## 4. Ways of working once all supervisors are on it

**One editor at a time:**
- Open the file, do the job, **close it**.
- Never leave it open on a control-room PC.
- If it opens read-only, the banner says who has it.
- At shift change, the outgoing supervisor closes it before handover.

**Who does what:**
| Task | Who |
|---|---|
| Call-outs | Any supervisor on shift |
| New postings | Any supervisor |
| Awards | The unit's supervisor, or the shift supervisor on the closing morning (8 AM) |
| Book time off, needs, leave | Admins (all supervisors, per the go-live decision) |
| Shift switches | The supervisor who receives the request |
| Void Entry, Add Adjustment, rename | Graeme + backup admin only, by agreement (the code doesn't enforce it) |
| Install updates, rollover | Graeme + backup admin only |

**The paper posting stays:** the system prints it, operators initial it, and the awarding supervisor records sign-ups in the Award pop-up.

**Corrections:** never type into the roster or log. Use Void Entry, then redo it. Any mistake gets a note in the void reason.

---

## 5. Stepwise go-live plan

Dates follow the pay periods (Sunday starts): Oct 18, **Nov 1**, Nov 15, **Nov 29**, Dec 13, **Dec 27**.

### Phase 0: one master, installed and checked (now to Oct 17)
1. **Merge the code.** Before installing v20, confirm it contains this morning's three fixes:
   - Out-of-dept drop-down: hours column, sorted lowest first.
   - Posting blank lines: name and initials only.
   - Two-way swap: payback date blocked unless the payer is free.
   If any are missing, have them merged into v20 in one session with the v20 zip attached, then label that zip **the master**.
2. Back up, install v20, compile, and run **SystemCheck** (expect all modules on v20). Import Kronos once.
3. Confirm the master's location (section 3) and set up the daily backup.
4. Add every supervisor's Windows login to `ADMIN_USERS`.
5. Run the v20 smoke test, including Thanksgiving (Oct 12) DAYS eligibility.
6. Answer the open decisions: Boxing Day weekday when announced, DST charging rule, Sick vs Off.

**Exit:** SystemCheck clean, smoke test passed, one master zip named.

### Phase 0b: member workbook build (after v20 is verified, Oct 12-30)
See `member-workbook-design.md`.
1. Build `BuildMemberTemplate` (creates the members' file) and `modPublish` (`PublishMembers`) in the master, behind the existing **Member View** button.
2. Test on plant PCs: open the members' file, sort the roster, check Who's next against a call-out list, find yourself in My history.
3. Privacy check: no phones, emails, leave types, void notes or logged-by in the published file.
4. Show it to Ray (union) before Nov 15.

**Exit:** a published file passes the privacy check and matches the master's call order.

### Phase 1: shadow run, Phosphine & Derivatives (Oct 18 - Oct 31, one pay period)
The **old process stays authoritative**. You run the workbook alongside it for every P&D call-out and posting.
1. Each call-out: compare the workbook's list and order with the old roster. Note every difference and its cause (data, rule, or bug).
2. Each posting: print from the workbook, but post the old one if they differ.
3. **Twice a week:** compare roster totals with Kronos + refusals. The Integrity check (v23) automates this later.
4. Re-book every upcoming P&D vacation imported from the old calendar through **Book Time Off** (deep-dive 1.1).
5. Install **v21** before Oct 31: DST night defaults, provisional accepted OT, Home Unit column, SelfTest.

**Exit:** two weeks with no unexplained differences in call order or charges, SelfTest green, P&D bookings re-entered.

### Phase 2: go-live wave 1, Phosphine & Derivatives (Nov 1)
1. The workbook becomes **authoritative** for P&D: call-outs, postings, awards, switches, time off.
2. The P&D vacation calendar workbook becomes a **read-only archive**. The Calendar tab and its PDF replace it.
3. Old-roster refusal entries stop for P&D. The OT Log is the record.
4. Print the P&D calendar PDF weekly for the board.
5. Rollback: if a blocking bug appears, go back to the old process for that unit and log manually. Re-enter the entries once it's fixed.

**Exit:** two weeks live with no rollback.

### Phase 3: train the other supervisors (Nov 1 - Nov 14)
1. **Supervisor quick guide** (1-2 pages, from the deep-dive 3.7 docs work): call-out, post, award (including out of dept), switch, time off, void request, "close the file".
2. A **30-minute walkthrough** with each of Chris, David, Claude, Ian, Nicole and Mark, on a **practice copy** (clearly named `PRACTICE - ...`, kept in a separate folder).
3. Each one does five tasks themselves on the practice copy: a call-out, a posting, an award, a two-way switch and a time-off booking.
4. Pick the **backup admin** (rollover, installs, voids). Recommended: one shift supervisor who's comfortable with Excel.

**Exit:** all six have done the five tasks; backup admin named.

### Phase 4: other units' data ready (Nov 1 - Nov 14, in parallel)
For each of Cypure, Fumigation (Material Handler), QC Lab, QC Gases, Utilities and SBP:
1. Import the unit's calendar (Oct-Dec and onward) with **Import Old Cal**, then re-book the time off in it through Book Time Off.
2. Check each person's crew, unit codes (grade) and email on the roster with that unit's supervisor.
3. Check the CC addresses for the unit's award emails.

**Exit:** each unit supervisor signs off their unit's people and bookings.

### Phase 5: go-live wave 2, remaining units (Nov 15)
1. All units go authoritative on the workbook. The unit calendars become read-only archives.
2. **Union communication** (before Nov 15). A short note to the local covering:
   - postings are printed from the system and list everyone signed off and free;
   - out-of-department sign-ups go on the blank lines;
   - call order, the coverage calendar, open postings and personal history are in the **member workbook**;
   - where to find it.
3. **Member workbook live.** Publish from the master any time with the button; Needs Attention reminds you when it's more than a day old.

### Phase 6: full operation and retirement (Nov 29)
1. Retire the old roster workbook and the Refusal Report. Archive them read-only with the date.
2. All supervisors log their own call-outs and switches.
3. Review after two weeks: issues list, guide updates, any rule questions for the union.

### Phase 7: year-end (Dec)
| When | Step |
|---|---|
| By Nov 30 | **v23** built: RolloverYear, Integrity check, rollover one-pager. |
| First week of Dec | **Dry run** on a copy, done by the **backup admin** following the one-pager, with you watching. |
| Dec 20 - Jan 4 | **Install freeze** (rollover only). |
| Dec 27 | **Rollover live.** Totals reset to zero on Jan 1 per the policy. The frozen 2026 copy is saved, and the Integrity check runs before and after. |

### Phase 8: cover for your late-January absence (Jan)
1. v24 (employee changes) live before Jan 16, then **no installs** the week before.
2. The backup admin runs the system alone for one full week before you leave.
3. The contact sheet in the guide covers who to call and how to roll back (restore yesterday's backup, re-enter the day from the paper postings and call notes).

---

## 6. Track 2: platform readiness and a later Power Apps look

Doesn't block go-live; runs alongside it.

| When | Step |
|---|---|
| Nov | Ask IT four things in writing: (1) licence path for Power Apps/Automate for ~10 named supervisors; (2) can we create SharePoint Lists on our site; (3) is VBA staying supported, and can our folder be a Trusted Location; (4) is the new Outlook coming, and when. |
| Q1 2027 | **v25:** Employee ID on every table, protection, and calendar speed-up. The tables stay list-shaped. |
| Q1 2027 | **SelfTest as the spec.** The known-answer cases (rotation, call order, charging, eligibility, holidays) become the acceptance tests any new platform must pass. |
| Q2 2027 (if licences are possible) | **Pilot:** the OT Log and Shift Switches as SharePoint Lists, with a small Power App for logging a call-out and a switch on a phone. Excel keeps the roster, postings and calendar and reads the lists. **Gate:** two weeks with SelfTest passing on both before anything else moves. |
| After the pilot | Move postings, then the calendar. Excel stays as the print/report engine as long as it's useful. |

**Decision gate (Q2 2027):** go to the pilot only if licences are approved for every person who logs, IT supports Lists on the site, and the pilot doesn't add work for supervisors on shift. Otherwise stay on Excel, which by then is hardened (v25) and documented.

---

## 7. Risks

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| Diverged code across sessions | Medium (happened Oct 9) | Wrong or old code installed | One master zip per version; SystemCheck; attach the zip at the start of every build session. |
| Untested code reaches live use | Medium | Wrong call or charge | Shadow run; SelfTest from v21; install on a copy first. |
| Imported vacations not re-booked | High until Phase 1/4 | Someone on vacation called or charged | Re-book during Phases 1 and 4; the convert tool in v22. |
| Lock-outs at shift change | Medium | Delayed call-out | Close-after-use rule; read-only banner names the holder. |
| Macro blocking by IT | Low-Medium | System unusable | Trusted Location; the IT questions in Track 2. |
| Union disputes the call order | Low-Medium | Grievances | Member workbook transparency; the union note before Nov 15; employee history report (v26). |
| Leave details exposed | Low | Privacy breach | Master in a supervisor-only folder; members get the values-only file, which never reads the master. |
| Year-end mistakes | Medium | Wrong 2027 start | Dry run by the backup admin; frozen copy; integrity check before and after. |
| Only you can maintain it | High today | Stalls while you're away | Backup admin; docs; freeze windows. |

---

## 8. Dates at a glance

| Date | Milestone |
|---|---|
| Oct 9-17 | Phase 0: master merged, v20 installed, admins added |
| Oct 12-30 | Phase 0b: member workbook built and tested |
| Oct 18-31 | Phase 1: P&D shadow run; v21 by Oct 31 |
| **Nov 1** | **Wave 1: P&D live** |
| Nov 1-14 | Training; other units' data |
| **Nov 15** | **Wave 2: all units live**; member workbook live; union note sent |
| **Nov 29** | Old roster and Refusal Report retired |
| Nov 30 | v23 built |
| Early Dec | Rollover dry run (backup admin) |
| Dec 20 - Jan 4 | Install freeze |
| **Dec 27** | **Rollover live** |
| Before Jan 16 | v24 live; backup admin solo week |
| Q1 2027 | v25 platform-readiness |
| Q2 2027 | Power Apps decision gate |

## 9. Needed from you
1. Where the master file lives today.
2. Who the backup admin is.
3. The Windows logins of all supervisors (for `ADMIN_USERS`).
4. Whether v20 already has the three Oct 9 morning fixes. Attach the v20 zip at the start of the next build session.
5. The open decisions from section 1.
6. The member folder path, and whether My history shows everyone's history or only the viewer's.
