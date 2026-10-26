# v28 report-back (Oct 25, 2026)

v28 = the January 2027 go-live release. Graeme decided Oct 25 (multiple choice): **build now** (the pre-release rule
waived for v28: v27.5 is tagged and zipped but not yet installed; any v27.5 fix is re-based under v28); **scope =
members' requests + the Payroll # as the key**; the "phone-readable page" is **cancelled**; the go-live facts (the date,
the master and members' folders, the final logins, the digest wording) are **placeholders set at install**. NOT in v28:
the email digest (item 39, still waits for the wording) and moving the go-live settings to the Settings sheet (they stay
constants set at install, as every release so far).

This page is step 2 of "How every update ships": what the code does today, the design, the decisions put to Graeme
(recorded at the end), then `briefs/v28.md`.

## 1. Members' requests - what exists today
- **The sign-up inbox is the model.** The members' file (`member/modMember.bas`: `MemberSignUp` / `MemberWithdraw` ->
  `SendSignup`) writes one small text file per click into `Signups\` beside the members' file (`SignupFolderPath`,
  `SIGNUP_FOLDER` ".\Signups"): `posting=`, `employee=`, `hours=`, `action=`, `login=`, `stamp=` lines; the master's
  `ReadSignups` (modSignups.bas:250) reads `P-nnnn__*.txt`, `ReduceSignups` keeps the newest action per person,
  `ApplySignups` pre-ticks the award pop-up, `ArchiveSignups` moves the files once the posting is closed. v27 added the
  suspect / login-mix flags.
- **The members' file has no UserForms** (modMember + ThisWorkbook only) and holds no rules; its pages are drawn from
  the published `d*` tables; Home is "me first" (M4) with the remembered name, tiles and six buttons.
- **Time off and switches are ADMIN jobs** (`ShowTimeOffForm` modCalendar.bas:3565, `ShowShiftSwitchForm`
  modShiftSwitch.bas:67; `RequireAdmin`). The pop-ups take their values from their own boxes; `frmNewPosting.Prefill`
  (modCalendar.bas:3489) is the one prefill pattern.
- **Needs Attention** (`modNeeds`: `NAGatherWork` / `NAGatherAdmin`, `NASection`, keys like `N:` / `U:` / `F:`) lists
  what waits for a supervisor or an admin; a double-click opens the right thing.
- **The publish** (`modPublish.PublishMembers`) writes the `d*` tables through `WriteTable`, every free text screened by
  `LeaveWordsIn`, `MEMBER_NEEDS` "Members M4" pins the template the tables are laid out for.

### Design (R)
- **Members' side (Members M5, one template rebuild):** a `Requests` page from Home ("ASK FOR TIME OFF" / "ASK FOR A
  SWITCH" buttons): typed cells - Kind (Time off / Switch), From, To, Shift (switch), Partner (switch: a roster name from
  a list), Note - and **Send**, which writes `REQ__<login>__<stamp>.txt` into `Requests\` beside the members' file
  (`REQUEST_FOLDER` ".\Requests", a sibling of `Signups\`; the same shape as a sign-up file, `kind=` / `from=` / `to=` /
  `shift=` / `partner=` / `note=` lines; the typed note screened for leave words in the FILE before it is written, like
  the Reason box). **My requests** on the same page: the member's rows from the published `dRequests` (Kind, From, To,
  Status Waiting / Booked / Declined / Withdrawn, the admin's answer line, decided on) + **Withdraw** (a second file,
  `action=withdraw`). Only "Time off" is ever published as a kind (no vacation / personal / lieu word leaves the master).
- **Master side (`modRequests`, a NEW module - N_MODULES 42):** `ReadRequests` reads the folder into `tblRequests` on
  the admin-only `Requests` sheet (Request ID `RQ-nnnn` from `NextRefID`'s marks, Kind, Employee, Login, From, To,
  Shift, Partner, Note, Status, Received, Decided By, Decided At, Answer, Ref = the `AB-` / `SS-` id it became) - on
  every open (Workbook_Open step) and from the Desk; the files are archived like sign-ups once decided. **Needs
  Attention** gets a section "MEMBERS' REQUESTS - waiting (double-click: book it, or decline with a reason)": the
  double-click opens the time-off pop-up / the switch pop-up **prefilled** (`frmTimeOff.Prefill` /
  `frmShiftSwitch.Prefill` - two form re-pastes) with the request's id carried; a save marks the request Booked with the
  Ref and a receipt; **Decline** (a menu item on the row) takes a reason and marks it Declined. A Desk tile "Requests
  waiting (N)" on the admin Desk. The publish adds `dRequests` (Employee, Kind "Time off" / "Switch", From, To, Shift,
  Partner, Status, Answer, Decided) screened by `LeaveWordsIn` / `PrivateDataIn`; `MEMBER_NEEDS` "Members M5".
- **Rules:** a request writes NOTHING on the roster by itself - the admin's save does, through the existing pop-ups and
  their rules (the 16 h stop, the sick window, bookings); a request past its date is marked Expired at open; a member may
  have several requests waiting; the Note is free text and is never published (the admin's Answer line is, screened).

## 2. The Payroll # as the key - what exists today
- Column A (`C_PAYROLL`) holds the Payroll # (8 digits, Kronos' Global ID), filled by the Kronos import, admin view
  only, **never published** (a personal identifier). `PayrollText` reads it; `modFormer` carries it (`FM_PAYROLL`); the
  three Kronos tables carry it; the import matches by number THEN name (`RosterNameCore`: a clash is named).
- **Every other table names people by the roster's spelling of the name** and `RenameEverywhere`
  (modPeople.bas, `NameColumns`: 15 name columns in 13 tables) rewrites them all together; the roster's pay-period
  cells are SUMIFS over the OT Log **by name**; `LoadEligible` / the call order / the publish / the members' Find me
  all work by name.

### Design (K) - "carry the key everywhere; rename and match by it; the name stays the display"
- Every people table that lacks one gets a **Payroll # column** (added on first use, as `FM_PAYROLL` was): OT Log (an
  `UpgradeOTLogLayout`-style verified layout step - the log is append-only, the new column is filled, nothing else
  moves), Posting lists, Absences, Shift Switches (two), Grade Log, Hours History, Crew History, Planned Moves,
  Restrictions, Evidence, Requests. A run-once `AssignPayrollKeys` (admin; the Install panel's group 3) fills them from
  the roster (name -> Payroll #) and the Former sheet, verified by `RosterTotalsMap` before / after; rows nobody
  matches are listed and left blank.
- **Writers** stamp the key on every new row (the one `EmpRec` / `PayrollOf(name)` lookup); **readers** match by key
  when both sides have one, by name when either lacks it (a new hire before their first Kronos import). The SUMIFS stay
  by name in v28 (a rename rewrites the names as today, now BY KEY, so a row whose name was misspelt by hand is found);
  `NameColumns` becomes the display-rewrite list. The Kronos import fills column A as today; **`FillPayrollKeys`** runs
  after it and after Add / Edit, propagating a newly known number to that person's rows.
- **The members' file never sees the number.** Find me stays by name; the sign-up and request files carry the name and
  the login as today.

## 3. Decisions put to Graeme (Oct 25)
1. Who acts on a member's request: **admins only** (time off and switches are admin jobs) / supervisors too.
2. A switch request: **the partner's name only, the admin confirms with both as today** / the partner must also click
   "Agree" in the file before the admin sees it.
3. The Payroll # key's reach in v28: **carry the key on every table + rename and match by it, the roster formulas stay
   by name** / the full key (the formulas and every lookup by number too) / only the tables that lack it, no rename
   change.
4. A new hire's Payroll #: **typed in the Employee pop-up at Add** (the import then confirms it) / left blank until the
   first Kronos import fills it.

## 4. Graeme's answers (Oct 25, 2026)
| Question | Answer |
|---|---|
| Who acts on a request | **Admins only** |
| A switch request's partner | **The partner clicks Agree in the file first** - the request waits in the members' file until the named partner, on their own PC, clicks Agree; only then does it reach the admin (a second file carrying the partner's login) |
| The Payroll # key's reach | **Carry it everywhere; rename and match by it; the roster's SUMIFS stay by name; the members' file never sees the number** |
| A new hire's Payroll # | **Typed at Add** in the Employee pop-up (8 digits, checked, may be blank); the import confirms or fills it |

Earlier the same day: build now (the pre-release rule waived), scope = these two items, the phone page cancelled, the
go-live facts as placeholders. The brief is `briefs/v28.md`.
