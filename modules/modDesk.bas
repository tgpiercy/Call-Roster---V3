Attribute VB_Name = "modDesk"
Option Explicit

' ====================================================================
'  DESK (v25.3): the landing screen - the admin Desk and the supervisor Desk
'
'  One sheet, "Desk", drawn by code like the Calendar: a header (today, the
'  pay period, how old the Kronos data is), TODAY tiles (shifts needing
'  cover, postings ready to award, paybacks owed, the members' file's age) and
'  plain-language buttons grouped by when he reaches for them. Every button
'  calls the macro the roster's bar calls; no rule lives here.
'
'  Two Desks from one drawing (Graeme, Oct 17, 2026):
'   - the ADMIN Desk: a login on BOTH ADMIN_USERS and DESK_USERS (settings) has
'     every admin right and lands on it (14 buttons); "Show everything" gives
'     the full master layout for the session; "< Desk" on the bar comes back.
'     The primary admin keeps today's layout and looks at it through that button.
'   - the SUPERVISOR Desk: a login on SUPERVISOR_USERS lands on it (SUPERVISOR_DESK;
'     v25.5 - until then every non-admin did): the jobs a supervisor may do - sick
'     call, call-out, calendar, new posting, award posting, print the call-out
'     roster, publish the members' file (v25.5) - plus the help buttons. The admin-only
'     jobs are not on it (and the roster's gates refuse them anyway). Anyone on
'     neither list is a plain user and lands on the roster.
'  ViewModeFor is the one pure decision (SelfTest suite "viewmode"). The admin
'  Desk's view is ApplyViewCore's desk mode; the supervisor Desk is the
'  ordinary view with the Desk sheet in front.
'
'  v27 (Desk v2, the approved mock-up docs/desk-v2-mockup-2026-10.html): four
'  bands - TODAY's tiles, then NOW | PLAN | FINISH side by side, one colour per
'  band (orange, green, navy; grey for help) so colour says WHEN; the header's
'  facts as three pills (the Kronos pill red with "import it" when old); every
'  tile names the next thing and the verb a click takes; Help, the two emails,
'  Undo and Practice in one grey footer row. Same buttons, same macros, same
'  rights: a UI-only redraw.
'
'  v29 (briefs/admin-lite.md; Graeme, Oct 27, 2026): a THIRD drawing, the ADMIN
'  LITE Desk, for an admin on LITE_DESK_USERS (Ray): drawn on the FULL admin view
'  (not desk mode - his call), four bands in fixed places - NEEDS YOUR DECISION
'  (requests, posting exceptions, corrections, returning from leave), BOOK AND
'  CHANGE, DAILY CHECK (three steps, green when done today, one line when all
'  three are), STAFFING AND OT (below minimum, OT by unit as bars) - over a trust
'  strip; "Show full admin Desk" draws the full Desk on the same sheet for the
'  session (gDeskShowFull, reset on open); "See Desk as: Supervisor / Admin Lite"
'  lets any admin look at either. Every Lite button runs the macro its full-Desk
'  twin runs; IsAdmin is untouched. The figures come from modLite; DrawLiteDesk
'  is its own procedure (DrawDesk sits at the 300-line limit).
' ====================================================================

Public Const DESK_VERSION As String = "Desk v28 (Oct 25)"
Public Const DESK_SHEET As String = "Desk"
Public Const DESK_AHEAD_DAYS As Long = 14           ' the "shifts need cover" tile looks this far ahead
Public Const HELP_SHEET As String = "Help"           ' v25.5: the Help sheet (ShowHelp / DrawHelp below)
Private Const HELP_COL_W As Double = 118            ' the Help sheet's column B: one wide wrapped column (characters)

Public gDeskMode As Boolean                         ' the Desk view is on (set by ApplyViewCore)
Private mZoomChecked As String                      ' sheets whose left-behind 120 % was put back this session (review)

Private Const DK_FONT As String = "Segoe UI"
Private Const DK_FONT_B As String = "Segoe UI Semibold"
Private Const BTN_H As Double = 36
Private Const BTN_W As Double = 270                 ' v27 (Desk v2): one band per column, three columns side by side
Private Const TILE_H As Double = 76                 ' v27: three lines - the count, the next thing, the verb a click takes
Private Const GAP As Double = 8
Private Const COLGAP As Double = 12                 ' v27: between the three band columns
Private Const PILL_H As Double = 20                 ' v27: the header's three pills
Private mGuideCounted As Boolean                 ' v27 (item 27): the first days strip counted this session
Private mPreview As Boolean                      ' v27.1: an admin is looking at the supervisor Desk (DeskPreview; this session only - gone on reopen)
Private mPreviewLite As Boolean                  ' v29: an admin is looking at the Admin Lite Desk (DeskPreviewLite; this session only)
Public gDeskShowFull As Boolean                  ' v29: a lite login asked for the full admin Desk (DeskShowFull; this session only - he lands on Lite again)

' ---- colours (the mock-up's) ----
Private Function dkNavy() As Long
    dkNavy = RGB(30, 56, 100)
End Function
Private Function dkInk() As Long
    dkInk = RGB(55, 65, 81)
End Function
Private Function dkMuted() As Long
    dkMuted = RGB(120, 130, 140)
End Function
Private Function dkGrey() As Long
    dkGrey = RGB(236, 239, 243)
End Function
Private Function dkOrange() As Long
    dkOrange = RGB(230, 81, 0)
End Function
Private Function dkGreen() As Long
    dkGreen = RGB(0, 128, 70)
End Function
Private Function dkRed() As Long
    dkRed = RGB(176, 0, 32)
End Function

' ====================================================================
'  WHO
' ====================================================================
' The one pure decision: "lite" (v29) = on ADMIN_USERS and LITE_DESK_USERS (admin rights, opens on the Admin Lite Desk
' drawn on the full admin view - Graeme, Oct 27, 2026; it beats "desk"), "desk" = on both ADMIN_USERS and DESK_USERS
' (admin rights, lands on the Desk in desk mode), "admin" = on ADMIN_USERS only (today's layout), "supervisor" (v25.5) = on
' SUPERVISOR_USERS and not an admin (the supervisor's jobs, the supervisor Desk), "user" = none of them (a desk-only or
' lite-only login has no admin rights; a plain user logs call-outs and looks at the calendar). "lite" is NEVER given to a
' login that is not an admin: IsSupervisor treats every answer but "user" as a supervisor. Lists are comma separated, any
' case, spaces ignored. Mirrored in tests/rules.py (suite viewmode). liteList is Optional so the older call sites compile.
Public Function ViewModeFor(ByVal login As String, ByVal adminList As String, ByVal deskList As String, ByVal supList As String, _
                            Optional ByVal liteList As String = "") As String
    Dim u As String: u = LCase$(Trim$(login))
    ViewModeFor = "user"
    If u = "" Then Exit Function
    Dim inA As Boolean, inD As Boolean, inS As Boolean, inL As Boolean
    inA = (InStr(1, "," & LCase$(Replace(adminList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    inD = (InStr(1, "," & LCase$(Replace(deskList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    inS = (InStr(1, "," & LCase$(Replace(supList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    inL = (InStr(1, "," & LCase$(Replace(liteList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    If inA And inL Then
        ViewModeFor = "lite"
    ElseIf inA And inD Then
        ViewModeFor = "desk"
    ElseIf inA Then
        ViewModeFor = "admin"
    ElseIf inS Then
        ViewModeFor = "supervisor"
    End If
End Function

' v29: the logins on the Lite list that are NOT admins, as "a, b" in the list's order (lower case, spaces dropped);
' "" when every one is an admin or the list is empty. SystemCheck's warn line. Pure; mirrored in tests/rules.py
' (suite litelogin).
Public Function LiteLoginProblems(ByVal liteList As String, ByVal adminList As String) As String
    Dim admins As String: admins = "," & LCase$(Replace(adminList, " ", "")) & ","
    Dim parts() As String, i As Long, u As String, out As String
    parts = Split(LCase$(Replace(liteList, " ", "")), ",")
    For i = LBound(parts) To UBound(parts)
        u = parts(i)
        If u <> "" Then
            If InStr(1, admins, "," & u & ",", vbBinaryCompare) = 0 Then out = out & IIf(out = "", "", ", ") & u
        End If
    Next i
    LiteLoginProblems = out
End Function

Public Function IsDeskUser() As Boolean
    IsDeskUser = (ViewModeFor(CurrentUser(), ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS, LITE_DESK_USERS) = "desk")
End Function

' v29: this login opens on the Admin Lite Desk (an admin on LITE_DESK_USERS). Not a desk-mode user: the Lite Desk draws on
' the full admin view (Graeme, Oct 27, 2026), so IsDeskUser is False for a lite login and the open applies the admin view.
Public Function IsLiteUser() As Boolean
    IsLiteUser = (ViewModeFor(CurrentUser(), ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS, LITE_DESK_USERS) = "lite")
End Function

' v29: the Lite layout is what shows - a lite login, or an admin previewing it (See Desk as: Admin Lite), and the full Desk
' was not asked for. Needs Attention draws in Ray's order on the same test (review MED-9); DrawDesk adds asAdmin.
Public Function LiteShowing() As Boolean
    LiteShowing = (IsLiteUser() Or mPreviewLite) And Not gDeskShowFull
End Function

' ====================================================================
'  THE SHEET
' ====================================================================
Private Function DeskSheet() As Worksheet
    If Not SheetExists(DESK_SHEET) Then
        Dim ws As Worksheet
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))     ' the first tab
        ws.Name = DESK_SHEET
        RelockStructure
    End If
    Set DeskSheet = ThisWorkbook.Worksheets(DESK_SHEET)
    SetTabColor DeskSheet, dkNavy()
End Function

' The four tiles' figures, from the roster's own functions (nothing is worked out here)
Public Sub DeskCounts(ByRef nCover As Long, ByRef nReady As Long, ByRef nOwed As Long, ByRef pubAt As Date)
    nCover = OpenLinesAhead(DESK_AHEAD_DAYS)
    nReady = ReadyToAwardCount()
    nOwed = 0
    Dim ids As Variant, v As Variant
    v = OwedSwitches(ids)
    If IsArray(v) Then nOwed = UBound(v, 1) - LBound(v, 1) + 1
    pubAt = MembersPublishedAt()
End Sub

' Draws (or redraws) the whole Desk: cheap, about twenty shapes
Public Sub DrawDesk()
    Dim ws As Worksheet: Set ws = DeskSheet()
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Failed
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "dk_" Then ws.Shapes(i).Delete
    Next i
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
    ws.Cells.Interior.Color = vbWhite
    With ws.Cells.Font
        .Name = DK_FONT
        .Size = 10
    End With
    ws.Columns(1).ColumnWidth = 2
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 28

    ' ---- header (v27, Desk v2: docs/desk-v2-mockup-2026-10.html): the title, then three pills - today's crews |
    '      the pay period | the Kronos data's age (red with "import it" when old) - and a red save warning when one is up ----
    Dim nCover As Long, nReady As Long, nOwed As Long, pubAt As Date
    DeskCounts nCover, nReady, nOwed, pubAt
    Dim ps As Date: ps = PeriodStartFor(Date)
    Dim dt As Date: dt = KronosDataThrough()
    Dim kOld As Boolean: kOld = (dt = 0 Or KronosIsStale())      ' the status line's one rule (review)
    Dim kTxt As String
    If dt = 0 Then
        kTxt = "Kronos data: none imported yet"
    ElseIf kOld Then
        kTxt = "Kronos data " & (Int(Date) - Int(dt)) & " days old"
    Else
        kTxt = "Kronos data through " & Format$(dt, "ddd mmm d")
    End If
    ' v27.1: the admin layout unless an admin is previewing the supervisor Desk (DeskPreview) - the three choices below read it
    Dim asAdmin As Boolean: asAdmin = gAdminUnlocked And Not mPreview
    ' the remedy belongs to whoever may run it: an admin imports, a supervisor asks (B5)
    If kOld Then kTxt = kTxt & IIf(asAdmin, " - import it", " - ask " & PRIMARY_ADMIN_NAME & " to import it")
    Dim kMac As String                             ' v27 (item 2): a report in KRONOS_PATH the roster has not read = one click imports it (admin)
    If asAdmin Then
        Dim nNewK As Long: nNewK = KronosNewFileCount()
        If nNewK > 0 Then
            kTxt = nNewK & " new Kronos report" & IIf(nNewK = 1, "", "s") & " - click to import"
            kOld = True: kMac = "modDesk.DeskKronos"     ' (review) the import, then the Desk redrawn
        End If
    End If
    Dim nsTxt As String                            ' v26.5.5 (A1): a failed save stays on the header until a save goes through
    If NoSave() Then
        nsTxt = "SAVING IS BLOCKED - close WITHOUT saving, reopen, log again"
    ElseIf gSaveFailedAt > 0 And Not ThisWorkbook.Saved Then
        nsTxt = "NOT SAVED since " & Format$(gSaveFailedAt, "h:mm AM/PM") & " - File > Save now"
    End If
    With ws.Cells(1, 2)
        .Value = "CALL OUT ROSTER"
        .Font.Name = DK_FONT_B: .Font.Size = 18: .Font.Color = dkNavy()
        .VerticalAlignment = xlCenter
    End With

    ' ---- positions: the pills on row 2, everything else hangs off the top-left of row 3 ----
    Dim x0 As Double: x0 = ws.Cells(1, 2).Left
    Dim y0 As Double: y0 = ws.Cells(3, 1).Top + 6
    Dim yy As Double, xx As Double
    Dim bandW As Double: bandW = 3 * BTN_W + 2 * COLGAP     ' the three columns' span: the tiles and the footer share it

    xx = x0: yy = ws.Cells(2, 1).Top + 3
    ' v25.5 (A8): today's shifts - the crews on days and nights - in the header
    If nsTxt <> "" Then xx = DeskPill(ws, "dk_pill_save", nsTxt, xx, yy, True)          ' first: it must never run off the right edge (review)
    xx = DeskPill(ws, "dk_pill_today", Format$(Date, "ddd mmm d") & " - " & DaysCrew(Date) & " on days, " & NightsCrew(Date) & " on nights", xx, yy, False)
    xx = DeskPill(ws, "dk_pill_period", "Pay period " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d"), xx, yy, False)
    xx = DeskPill(ws, "dk_pill_kronos", kTxt, xx, yy, kOld, kMac)
    ' v27.1: the preview says so, after the other pills (nothing else moves)
    If mPreview And gAdminUnlocked Then xx = DeskPill(ws, "dk_pill_preview", "PREVIEW: a supervisor's Desk", xx, yy, True, "modDesk.DeskPreview")
    ' v29 (item 3b): the Admin Lite Desk - a lite login, or an admin previewing it, unless the full Desk was asked for
    Dim lite As Boolean: lite = asAdmin And LiteShowing()
    ws.Rows(3).RowHeight = ws.StandardHeight: ws.Rows(4).RowHeight = ws.StandardHeight   ' (review LOW-8) the Lite rows' height does not linger on the other drawings
    If mPreviewLite And gAdminUnlocked Then xx = DeskPill(ws, "dk_pill_preview", "PREVIEW: the Admin Lite Desk", xx, yy, True, "modDesk.DeskPreviewLite")
    If asAdmin And (IsLiteUser() Or mPreviewLite) And gDeskShowFull Then xx = DeskPill(ws, "dk_pill_lite", "<  Back to my Desk", xx, yy, False, "modDesk.DeskShowFull")
    If lite Then
        yy = DrawLiteDesk(ws, x0, bandW)
        GoTo Footer
    End If

    ' ---- the layout: the admin Desk (every job) or the supervisor Desk (a supervisor's jobs) ----
    Dim sup As Boolean: sup = Not asAdmin

    ' ---- TODAY: the tiles, each the count, the next thing, the verb. v27 (item 2): an admin's sit in rows of three
    '      under the three band columns (cover, award, owed / members' file, this morning); a supervisor's in one row ----
    DeskLabel ws, "dk_lbl_today", "TODAY", x0, y0
    yy = y0 + 16
    Dim nTiles As Long: nTiles = IIf(sup, 4, IIf(REQUESTS_ON, 7, 6))         ' v27 (item 3): + Since you were last here; v28: + members' requests (admin)
    Dim k As Long, tw As Double, tileX As Double, tileY As Double
    Dim nextTxt As String                          ' v25.5 (A8): the next uncovered shift under the count
    If nCover > 0 Then nextTxt = NextOpenLine(DESK_AHEAD_DAYS)
    TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
    DeskTile ws, "dk_tile_cover", nCover, IIf(nCover = 1, "shift needs cover", "shifts need cover") & IIf(nextTxt <> "", ", next " & nextTxt, ""), _
             IIf(sup, "modDesk.DeskCalendar", "modDesk.DeskNeeds"), tileX, tileY, tw, IIf(nCover > 0, dkOrange(), 0), _
             IIf(nCover = 0, "", IIf(sup, "Open the Calendar", "Open Needs Attention"))   ' Needs Attention is admin: a supervisor's tile opens the Calendar; no verb on a grey tile
    k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
    Dim oldest As Date: If nReady > 0 Then oldest = ReadyToAwardOldest()      ' v27: the posting that has waited longest
    DeskTile ws, "dk_tile_award", nReady, IIf(nReady = 1, "posting ready to award", "postings ready to award") & _
             IIf(oldest > 0, ", closed " & Format$(oldest, "ddd mmm d") & " " & POSTING_CLOSE_TIME, ""), "modDesk.DeskAward", _
             tileX, tileY, tw, IIf(nReady > 0, dkOrange(), 0), IIf(nReady = 0, "", IIf(nReady = 1, "Award it", "Award them"))
    If Not sup Then
        k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
        Dim nOver As Long: nOver = OwedOverdueCount()      ' v26 (payback aging): owed longer than PAYBACK_WARN_DAYS
        DeskTile ws, "dk_tile_owed", nOwed, IIf(nOwed = 1, "payback owed", "paybacks owed") & _
                 IIf(nOver > 0, ", " & nOver & " over " & PaybackWarnDays() & " days", ""), "modDesk.DeskSwitch", _
                 tileX, tileY, tw, IIf(nOwed > 0, dkGreen(), 0), IIf(nOwed = 0, "", "See the OWED list")
    End If
    k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
    Dim pubTxt As String, pubCol As Long
    If pubAt = 0 Then
        pubTxt = "never published": pubCol = dkNavy()
    ElseIf Int(Date) - Int(pubAt) >= MembersStaleDays() Then
        pubTxt = "updated " & Format$(pubAt, "ddd mmm d"): pubCol = dkNavy()
    Else
        pubTxt = "updated today " & Format$(pubAt, "h:mm AM/PM"): pubCol = 0
    End If
    ' v25.5 (bundle item 8): the tile is live for a supervisor too - they publish (coloured = work to do)
    DeskTileText ws, "dk_tile_pub", "Members' file", pubTxt, "modDesk.DeskPublish", tileX, tileY, tw, pubCol, _
                 IIf(pubCol <> 0, "Update it", "Click to update again")
    If Not sup Then                                  ' v27 (item 2): the admin's morning check - what the open already knows
        k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
        Dim mTxt As String, mCol As Long, mVerb As String
        MorningCheck mTxt, mCol, mVerb
        DeskTileText ws, "dk_tile_morning", "This morning", mTxt, "modDesk.DeskMorning", tileX, tileY, tw, mCol, mVerb
    End If
    ' v27 (item 3): what other logins logged since this login's previous open
    k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
    Dim prevAt As Date: prevAt = PrevOpenAt()
    Dim nSince As Long: nSince = ActivitySinceCount(prevAt, True)
    DeskTile ws, "dk_tile_since", nSince, IIf(prevAt = 0, "first open on record for you", _
             IIf(nSince = 1, "entry", "entries") & " by others since you were last here (" & Format$(prevAt, "ddd h:mm AM/PM") & ")"), _
             "modDesk.DeskSince", tileX, tileY, tw, IIf(nSince > 0, dkNavy(), 0), IIf(nSince = 0, "", "See what changed")
    If Not sup And REQUESTS_ON Then                  ' v28 (item 1): members' requests waiting for an admin (time off, switches)
        k = k + 1: TilePos k, sup, nTiles, x0, yy, bandW, tileX, tileY, tw
        Dim nReq As Long: nReq = RequestsWaitingCount()
        DeskTile ws, "dk_tile_requests", nReq, IIf(nReq = 1, "member's request waiting (time off or a switch)", "members' requests waiting (time off, switches)"), _
                 "modDesk.DeskRequests", tileX, tileY, tw, IIf(nReq > 0, dkGreen(), 0), IIf(nReq = 0, "", "Open Needs Attention")
    End If

    ' ---- the three bands side by side: NOW (orange) | PLAN (green) | FINISH (navy) - colour says WHEN, not which pop-up ----
    yy = yy + TileRows(sup, nTiles) * (TILE_H + GAP) - GAP + 18
    Dim yN As Double, yP As Double, yF As Double
    Dim xN As Double, xP As Double, xF As Double
    xN = x0: xP = x0 + BTN_W + COLGAP: xF = x0 + 2 * (BTN_W + COLGAP)
    yN = yy: yP = yy: yF = yy
    DeskLabel ws, "dk_lbl_now", "NOW", xN, yN
    DeskLabel ws, "dk_lbl_plan", "PLAN", xP, yP
    DeskLabel ws, "dk_lbl_finish", "FINISH", xF, yF
    yN = yN + 16: yP = yP + 16: yF = yF + 16
    ' NOW
    DeskBtn ws, "dk_btn_sick", "Someone called in sick", "modDesk.DeskSick", xN, yN, dkOrange()
    yN = yN + BTN_H + GAP
    DeskBtn ws, "dk_btn_callout", "Fill a shift now (call-out)", "modDesk.DeskCallout", xN, yN, dkOrange()
    yN = yN + BTN_H + GAP
    DeskBtn ws, "dk_btn_cal", "What needs cover (calendar)", "modDesk.DeskCalendar", xN, yN, dkOrange()
    yN = yN + BTN_H + GAP
    ' PLAN
    DeskBtn ws, "dk_btn_post", "Post a shift", "modDesk.DeskPost", xP, yP, dkGreen()
    yP = yP + BTN_H + GAP
    DeskBtn ws, "dk_btn_award", "Award a posting", "modDesk.DeskAward", xP, yP, dkGreen()
    yP = yP + BTN_H + GAP
    DeskBtn ws, "dk_btn_ledger", "Paybacks owed (ledger)", "modDesk.DeskLedger", xP, yP, dkGreen()   ' v27 (item 25): by person and age, a statement per person
    yP = yP + BTN_H + GAP
    If Not sup Then
        DeskBtn ws, "dk_btn_timeoff", "Book time off", "modDesk.DeskTimeOff", xP, yP, dkGreen()
        yP = yP + BTN_H + GAP
        DeskBtn ws, "dk_btn_switch", "Shift switch or payback", "modDesk.DeskSwitch", xP, yP, dkGreen()
        yP = yP + BTN_H + GAP
        DeskBtn ws, "dk_btn_leave", "Leave: going or returning", "modDesk.DeskLeave", xP, yP, dkGreen()   ' v27: planned, like time off
        yP = yP + BTN_H + GAP
        DeskBtn ws, "dk_btn_tables", "Holidays / minimum staffing", "modDesk.DeskTables", xP, yP, dkGreen()   ' v27 (item 33): the two pop-ups
        yP = yP + BTN_H + GAP
    End If
    ' FINISH
    DeskBtn ws, "dk_btn_publish", "Update the members' file", "modDesk.DeskPublish", xF, yF, dkNavy()   ' v25.5: on the supervisor Desk too (bundle item 8)
    yF = yF + BTN_H + GAP
    DeskBtn ws, "dk_btn_print", "Print the call-out roster", "modDesk.DeskPrint", xF, yF, dkNavy()
    yF = yF + BTN_H + GAP
    DeskBtn ws, "dk_btn_history", "Employee history (report)", "modDesk.DeskHistory", xF, yF, dkNavy()   ' v26: one person, a date range, a PDF
    yF = yF + BTN_H + GAP
    DeskBtn ws, "dk_btn_handover", "Shift handover (last 12 h)", "modDesk.DeskHandover", xF, yF, dkNavy()   ' v27 (item 25): one printable page
    yF = yF + BTN_H + GAP
    If Not sup Then
        DeskBtn ws, "dk_btn_reconcile", "Reconcile with the paper roster", "modDesk.DeskReconcile", xF, yF, dkNavy()   ' v27 (item 4): the parallel run's sheet
        yF = yF + BTN_H + GAP
        DeskBtn ws, "dk_btn_stats", "Stats (quarterly report)", "modDesk.DeskStats", xF, yF, dkNavy()   ' v27 (item 26)
        yF = yF + BTN_H + GAP
        DeskBtn ws, "dk_btn_rules", "The rules (sheet)", "modRules.DrawRules", xF, yF, dkNavy()   ' v27 (item 22): the Rules sheet, Basis typed by an admin
        yF = yF + BTN_H + GAP
        DeskBtn ws, "dk_btn_close", "Close the pay period", "modEvidence.ClosePeriod", xF, yF, dkNavy()   ' v27 (item 23): the CSV + PDF + hash
        yF = yF + BTN_H + GAP
        DeskBtn ws, "dk_btn_withdraw", "Withdraw an award / edit a posting", "modDesk.DeskPostingJobs", xF, yF, dkNavy()   ' v27 (item 24)
        yF = yF + BTN_H + GAP
    End If

    ' ---- the footer row: help, the two emails, Undo (admin), Practice - grey, reached for rarely ----
    yy = yN: If yP > yy Then yy = yP
    If yF > yy Then yy = yF
    yy = yy + 10
Footer:                                              ' v29: the Lite drawing joins here (sup is False: an admin's footer)
    Dim nFoot As Long: nFoot = IIf(sup, 5, 6)           ' v27 (item 20, B14): + Larger type
    Dim fw As Double: fw = (bandW - (nFoot - 1) * GAP) / nFoot
    xx = x0
    DeskBtn ws, "dk_btn_guide", "How do I ...", "modDesk.DeskGuide", xx, yy, dkGrey(), fw, dkInk(), 10
    xx = xx + fw + GAP
    DeskBtn ws, "dk_btn_report", "Something went wrong", "modDesk.DeskReport", xx, yy, dkGrey(), fw, dkInk(), 10
    xx = xx + fw + GAP
    ' v25.5 (A8): a suggestion goes the same way as a fault, with its own subject
    DeskBtn ws, "dk_btn_suggest", "Suggest an improvement", "modDesk.DeskSuggest", xx, yy, dkGrey(), fw, dkInk(), 10
    xx = xx + fw + GAP
    If Not sup Then
        DeskBtn ws, "dk_btn_undo", "Undo my last log entry", "modDesk.DeskUndo", xx, yy, dkGrey(), fw, dkInk(), 10   ' v25.5 (A9): admin only (void); "log entry": it undoes OT Log rows only (review)
        xx = xx + fw + GAP
    End If
    ' v26.5.5 (5.4): a practice copy for anyone who may log - nothing done in it counts
    DeskBtn ws, "dk_btn_practice", "Practice on a copy", "modDesk.DeskPractice", xx, yy, dkGrey(), fw, dkInk(), 10
    xx = xx + fw + GAP
    ' v27 (item 20, B14): this login's pop-ups at 120 % (a hidden name per login; the next pop-up shows it)
    DeskBtn ws, "dk_btn_larger", IIf(UiScale() > 1.05, "Normal type", "Larger type"), "modDesk.DeskLargerType", xx, yy, dkGrey(), fw, dkInk(), 10

    ' ---- the links, then the shortcut keys (v25.5, A5) ----
    yy = yy + BTN_H + 10
    DeskLink ws, "dk_lnk_roster", "Full roster  >", "modDesk.DeskRoster", x0, yy, 140
    ' v29: not on the Lite Desk (it already sits on the full view) - unless desk mode is on (< Desk or Ctrl+Shift+D put it there: review MED-3)
    If Not sup And (Not lite Or gDeskMode) Then DeskLink ws, "dk_lnk_all", "Show everything (primary admin tools)  >", "modDesk.DeskShowAll", x0 + 150, yy, 260
    ' v27.1: an admin can look at the supervisor Desk and come back; v29: or at the Admin Lite Desk ("See Desk as"; a lite
    ' login has no Admin Lite link - it is his own Desk)
    If (mPreview Or mPreviewLite) And gAdminUnlocked Then
        DeskLink ws, "dk_lnk_preview", "<  Back to my Desk", "modDesk.DeskPreviewOff", IIf(mPreview, x0 + 150, x0 + 420), yy, 200
    ElseIf asAdmin Then
        DeskLink ws, "dk_lnk_preview", "See Desk as:  Supervisor  >", "modDesk.DeskPreview", x0 + 420, yy, 190
        If Not IsLiteUser() Then DeskLink ws, "dk_lnk_previewlite", "Admin Lite  >", "modDesk.DeskPreviewLite", x0 + 615, yy, 110
    End If
    yy = yy + 26
    DeskLabel ws, "dk_lbl_keys", "Ctrl+Shift+S sick call   Ctrl+Shift+C call-out   Ctrl+Shift+K calendar   Ctrl+Shift+D this Desk", x0, yy
    ws.Shapes("dk_lbl_keys").Width = 440
    ' v27 (item 27, 5.5): the one legend, the calendar's words and the log's marks
    yy = yy + 18
    DeskLabel ws, "dk_lbl_legend1", "Calendar:  " & CalLegendText(), x0, yy
    yy = yy + 16
    DeskLabel ws, "dk_lbl_legend2", "Log:  " & OutcomeLegendText(), x0, yy
    On Error Resume Next
    With ws.Shapes("dk_lbl_legend1")                ' (review) wrapped inside the band, two lines each
        .TextFrame2.WordWrap = msoTrue: .Width = bandW: .Height = 26
    End With
    With ws.Shapes("dk_lbl_legend2")
        .TextFrame2.WordWrap = msoTrue: .Width = bandW: .Height = 26: .Top = .Top + 12
    End With
    yy = yy + 24
    On Error GoTo Failed
    If GuideShows(lite) Then                        ' v27 (item 27): the first days strip; v29: a lite login's own five steps
        yy = yy + 22
        DeskLabel ws, "dk_lbl_guide", GuideLine(lite), x0, yy
        On Error Resume Next
        ws.Shapes("dk_lbl_guide").Width = 640
        ws.Shapes("dk_lbl_guide").TextFrame2.TextRange.Font.Fill.ForeColor.RGB = dkNavy()
        On Error GoTo Failed
        DeskLink ws, "dk_lnk_guidehide", "Hide this", "modDesk.DeskGuideHide", x0 + 650, yy - 4, 90
    End If
    With ws.Cells(1, 9)
        .Value = DESK_VERSION
        .Font.Size = 7: .Font.Color = RGB(176, 190, 197)
        .HorizontalAlignment = xlRight
    End With

    If ActiveSheet Is ws Then
        On Error Resume Next
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .Zoom = 100
        End With
        On Error GoTo Failed
    End If
    GoTo Done
Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    Resume AfterFail
AfterFail:
    On Error Resume Next
    LogError "DrawDesk", en, em
    UnlockSheet ws
    ws.Cells(3, 2).Value = "The Desk couldn't finish drawing: " & em & "  (" & DESK_VERSION & ")"
    ws.Cells(3, 2).Font.Color = dkRed()
Done:
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
End Sub

' ---- shape helpers ----
Private Sub DeskLabel(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal l As Double, ByVal t As Double)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, 260, 14)
    s.Name = nm
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    On Error Resume Next                            ' never raises
    s.DrawingObject.PrintObject = False
    With s.TextFrame2
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = txt
        .TextRange.Font.Name = DK_FONT_B
        .TextRange.Font.Size = 9
        .TextRange.Font.Fill.ForeColor.RGB = dkMuted()
    End With
End Sub

Private Function DeskShape(ByVal ws As Worksheet, ByVal nm As String, ByVal mac As String, ByVal l As Double, ByVal t As Double, _
                           ByVal w As Double, ByVal h As Double, ByVal fill As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    s.Name = nm
    s.Fill.ForeColor.RGB = fill
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    If mac <> "" Then s.OnAction = mac
    On Error Resume Next                            ' never raises
    s.DrawingObject.PrintObject = False
    s.Adjustments.Item(1) = 0.18
    With s.TextFrame2
        .WordWrap = msoTrue
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 2: .MarginBottom = 2
    End With
    Set DeskShape = s
End Function

' a verb-phrase button, 36 points tall (12 pt; the footer's are 10 pt)
Private Sub DeskBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                    ByVal l As Double, ByVal t As Double, ByVal fill As Long, _
                    Optional ByVal w As Double = BTN_W, Optional ByVal fontCol As Long = vbWhite, Optional ByVal fontSize As Double = 12)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, BTN_H, fill)
    On Error Resume Next                            ' never raises
    With s.TextFrame2.TextRange
        .Text = cap
        .ParagraphFormat.Alignment = msoAlignLeft
        .Font.Name = DK_FONT_B
        .Font.Size = fontSize
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = fontCol
    End With
End Sub

' v27 (Desk v2): a header pill - one fact, grey; red with white text when it is a warning. Returns the next x.
Private Function DeskPill(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, ByVal l As Double, ByVal t As Double, _
                          ByVal warn As Boolean, Optional ByVal mac As String = "") As Double     ' v27 (item 2): a pill may run a macro
    Dim w As Double: w = Len(txt) * 5.4 + 22                 ' Segoe UI Semibold 10 pt, about 5.4 pt a character
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, PILL_H, IIf(warn, dkRed(), dkGrey()))
    On Error Resume Next                            ' never raises
    s.Adjustments.Item(1) = 0.5
    With s.TextFrame2
        .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        With .TextRange
            .Text = txt
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = DK_FONT_B
            .Font.Size = 10
            .Font.Bold = msoFalse
            .Font.Fill.ForeColor.RGB = IIf(warn, vbWhite, dkInk())
        End With
    End With
    DeskPill = l + w + GAP
End Function

' v27 (item 2): where TODAY tile k (0-based) sits - an admin's in rows of three under the band columns, a
' supervisor's all in one row across the bands' span
Private Sub TilePos(ByVal k As Long, ByVal sup As Boolean, ByVal nTiles As Long, ByVal x0 As Double, ByVal y0 As Double, _
                    ByVal bandW As Double, ByRef x As Double, ByRef y As Double, ByRef w As Double)
    If sup Then
        w = (bandW - (nTiles - 1) * GAP) / nTiles
        x = x0 + k * (w + GAP)
        y = y0
    Else
        w = BTN_W
        x = x0 + (k Mod 3) * (BTN_W + COLGAP)
        y = y0 + (k \ 3) * (TILE_H + GAP)
    End If
End Sub

' how many rows of tiles TilePos laid out
Private Function TileRows(ByVal sup As Boolean, ByVal nTiles As Long) As Long
    If sup Then TileRows = 1 Else TileRows = (nTiles + 2) \ 3
End Function

' ====================================================================
'  v29 (briefs/admin-lite.md): THE ADMIN LITE DESK'S DRAWING
' ====================================================================
' Called by DrawDesk after the shared header pills (row 2); draws rows 3-4 (the Lite pills, the trust strip) and the
' four bands from row 5 down, every shape named dk_lite_* / dk_q_* so DrawDesk's redraw deletes them; returns the y the
' footer starts at. Every figure comes from modLite or the roster's own functions; every button runs the macro its
' full-Desk twin runs. Cards keep their places when empty (grey, "Nothing waiting") so the eye learns where things are.
Private Function DrawLiteDesk(ByVal ws As Worksheet, ByVal x0 As Double, ByVal bandW As Double) As Double
    Dim xx As Double, yy As Double, k As Long
    Dim tw As Double: tw = (bandW - 3 * GAP) / 4            ' four cards across the bands' span

    ' ---- row 3: ADMIN LITE | Signed in as Ray - admin rights | Show full admin Desk ----
    ws.Rows(3).RowHeight = 26: ws.Rows(4).RowHeight = 26
    xx = x0: yy = ws.Cells(3, 1).Top + 3
    xx = DeskPill(ws, "dk_lite_pill", "ADMIN LITE", xx, yy, False)
    xx = DeskPill(ws, "dk_lite_who", "Signed in as " & IIf(IsLiteUser(), LITE_FIRST_NAME, CurrentUser()) & " - admin rights", xx, yy, False)
    xx = DeskPill(ws, "dk_lite_full", "Show full admin Desk  >", xx, yy, False, "modDesk.DeskShowFull")

    ' ---- row 4: the trust strip - four facts, each OK (grey) or red with "call <admin>" ----
    xx = x0: yy = ws.Cells(4, 1).Top + 3
    Dim call_ As String: call_ = " - call " & PRIMARY_ADMIN_NAME
    Dim dt As Date: dt = KronosDataThrough()
    Dim bad As Boolean: bad = (dt = 0 Or KronosIsStale())
    Dim nProv As Long: nProv = ProvisionalCount()
    xx = DeskPill(ws, "dk_lite_t1", IIf(dt = 0, "Hours: no Kronos data yet", "Hours through " & Format$(dt, "ddd mmm d")) & _
                  IIf(nProv > 0, " + " & nProv & IIf(nProv = 1, " person", " people") & " provisional", "") & IIf(bad, call_, ""), xx, yy, bad)
    Dim bk As Long: bk = BackupTodayState()
    xx = DeskPill(ws, "dk_lite_t2", Choose(bk + 2, "Backup: not in this copy", "NO backup today" & call_, "Backup today: yes"), xx, yy, (bk = 0))
    Dim pubAt As Date: pubAt = MembersPublishedAt()
    bad = (pubAt = 0 Or Int(Date) - Int(pubAt) >= MembersStaleDays())
    xx = DeskPill(ws, "dk_lite_t3", IIf(pubAt = 0, "Members' file: never published", "Members' file updated " & Format$(pubAt, "ddd h:mm AM/PM")) & IIf(bad, call_, ""), xx, yy, bad)
    Dim nErr As Long: nErr = UnseenErrorCount()
    bad = IntegrityHasProblems() Or nErr > 0
    xx = DeskPill(ws, "dk_lite_t4", IIf(bad, "System checks: " & IIf(IntegrityHasProblems(), "an integrity PROBLEM", nErr & " new error" & IIf(nErr = 1, "", "s")) & call_, "System checks: OK"), xx, yy, bad)

    ' ---- band 1: NEEDS YOUR DECISION - four cards, one verb each ----
    yy = ws.Cells(5, 1).Top + 6
    DeskLabel ws, "dk_lite_lbl1", "NEEDS YOUR DECISION  -  oldest first, one decision per card", x0, yy
    ws.Shapes("dk_lite_lbl1").Width = 420
    DeskLink ws, "dk_q_decide", "?", "modDesk.DeskLiteHelp", x0 + bandW - 22, yy - 4, 22
    yy = yy + 16
    Dim n As Long, n2 As Long, txt As String, id As String, what As String
    n = RequestsWaitingCount(): n2 = RequestsPartnerCount()
    If n = 0 Then
        txt = IIf(n2 > 0, "no request waiting for you; " & n2 & " waiting for a partner's Agree", "Nothing waiting - members' requests")
    Else
        txt = IIf(n = 1, "member's request waiting", "members' requests waiting") & IIf(n2 > 0, "; " & n2 & " waiting for a partner", "")
    End If
    DeskTile ws, "dk_lite_req", n, txt, "modDesk.DeskRequests", x0, yy, tw, IIf(n > 0, dkGreen(), 0), IIf(n = 0, "", "Decide in Needs Attention")
    n = PostingExceptionCount(id, what)
    DeskTile ws, "dk_lite_post", n, IIf(n = 0, "Nothing waiting - posting exceptions", IIf(n = 1, "posting", "postings") & " with no member sign-up" & IIf(id <> "", ": " & id & " (" & what & ")", "")), _
             "modDesk.DeskLiteAward", x0 + (tw + GAP), yy, tw, IIf(n > 0, dkOrange(), 0), IIf(n = 0, "", "No takers / Cancel / Out of dept")
    n = UnresolvedNoCount(): n2 = 0: txt = OpenDisputesText(n2)
    txt = IIf(n > 0, n & " receipt " & IIf(n = 1, "answer", "answers") & " of No / Not sure", "") & IIf(n > 0 And n2 > 0, "; ", "") & _
          IIf(n2 > 0, n2 & " open dispute" & IIf(n2 = 1, "", "s"), "")
    DeskTile ws, "dk_lite_fix", n + n2, IIf(n + n2 = 0, "Nothing waiting - corrections", txt), _
             "modDesk.DeskFixEntry", x0 + 2 * (tw + GAP), yy, tw, IIf(n + n2 > 0, dkRed(), 0), IIf(n + n2 = 0, "", "Fix an entry")
    n = LeaveReturningCount(DESK_AHEAD_DAYS, txt)
    DeskTile ws, "dk_lite_leave", n, IIf(n = 0, "Nothing waiting - returning from leave", "returning from leave in " & DESK_AHEAD_DAYS & " days" & IIf(txt <> "", ": " & txt, "")), _
             "modDesk.DeskLeave", x0 + 3 * (tw + GAP), yy, tw, IIf(n > 0, dkGreen(), 0), IIf(n = 0, "", "Leave > Returning")
    yy = yy + TILE_H + 18

    ' ---- band 2: BOOK AND CHANGE - four large buttons with a one-line hint each ----
    DeskLabel ws, "dk_lite_lbl2", "BOOK AND CHANGE", x0, yy
    DeskLink ws, "dk_q_book", "?", "modDesk.DeskLiteHelp", x0 + bandW - 22, yy - 4, 22
    yy = yy + 16
    DeskBtn ws, "dk_lite_timeoff", "Book time off", "modDesk.DeskTimeOff", x0, yy, dkGreen(), tw
    DeskBtn ws, "dk_lite_leavebtn", "Leave: going or returning", "modDesk.DeskLeave", x0 + (tw + GAP), yy, dkGreen(), tw
    DeskBtn ws, "dk_lite_need", "Add a need", "modDesk.DeskAddNeed", x0 + 2 * (tw + GAP), yy, dkGreen(), tw
    DeskBtn ws, "dk_lite_person", "Change a person", "modDesk.DeskPerson", x0 + 3 * (tw + GAP), yy, dkGreen(), tw
    yy = yy + BTN_H + 3
    DeskLabel ws, "dk_lite_h1", "vacation, personal, a booked sick day", x0, yy
    DeskLabel ws, "dk_lite_h2", "STD, LOA, parental; the return top-up", x0 + (tw + GAP), yy
    DeskLabel ws, "dk_lite_h3", "an uncovered shift on the calendar", x0 + 2 * (tw + GAP), yy
    DeskLabel ws, "dk_lite_h4", "new hire, unit move, rename, remove", x0 + 3 * (tw + GAP), yy
    For k = 1 To 4
        ws.Shapes("dk_lite_h" & k).Width = tw
    Next k
    yy = yy + 14 + 14

    ' ---- band 3: DAILY CHECK - three numbered steps, green when done today; one green line when all three are ----
    DeskLabel ws, "dk_lite_lbl3", "DAILY CHECK  -  only when " & PRIMARY_ADMIN_NAME & " is off", x0, yy
    ws.Shapes("dk_lite_lbl3").Width = 420
    DeskLink ws, "dk_q_daily", "?", "modDesk.DeskLiteHelp", x0 + bandW - 22, yy - 4, 22
    yy = yy + 16
    Dim dl As String: dl = DailyCheckLine()
    If dl <> "" Then
        DeskBtn ws, "dk_lite_daily", dl, "", x0, yy, dkGreen(), bandW
        GuideTick 5, True                            ' the strip's fifth step: the daily check done
    Else
        Dim tw3 As Double: tw3 = (bandW - 2 * GAP) / 3
        Dim st As String, nCover As Long: nCover = OpenLinesAhead(DESK_AHEAD_DAYS)
        st = DailyStepText("kronos")
        DeskBtn ws, "dk_lite_d1", "1  Import Kronos" & IIf(st <> "", " - done " & st, ""), "modDesk.DeskKronos", x0, yy, IIf(st <> "", dkGreen(), dkOrange()), tw3
        st = DailyStepText("needs")
        DeskBtn ws, "dk_lite_d2", "2  Clear Needs Attention" & IIf(nCover > 0, " (" & nCover & " to cover)", "") & IIf(st <> "", " - done " & st, ""), _
                "modDesk.DeskNeeds", x0 + tw3 + GAP, yy, IIf(st <> "", dkGreen(), dkOrange()), tw3
        st = DailyStepText("publish")
        DeskBtn ws, "dk_lite_d3", "3  Publish the members' file" & IIf(st <> "", " - done " & st, ""), "modDesk.DeskPublish", x0 + 2 * (tw3 + GAP), yy, IIf(st <> "", dkGreen(), dkOrange()), tw3
    End If
    yy = yy + BTN_H + 18

    ' ---- band 4: STAFFING AND OT - below minimum ahead, then this period's OT by unit as bars ----
    DeskLabel ws, "dk_lite_lbl4", "STAFFING AND OT", x0, yy
    DeskLink ws, "dk_q_staff", "?", "modDesk.DeskLiteHelp", x0 + bandW - 22, yy - 4, 22
    yy = yy + 16
    StaffSweepBegin Date                             ' (review MED-2) fresh minimums, names and absences for this read
    Dim sl As Variant: sl = StaffShortList(Date, LITE_STAFF_DAYS, "")
    n = 0: txt = ""
    If IsArray(sl) Then
        n = UBound(sl, 1)
        For k = 1 To n
            If k <= 3 Then txt = txt & IIf(txt = "", "", "; ") & sl(k, 1) & " " & Format$(CDate(sl(k, 2)), "ddd d") & " " & sl(k, 3)
        Next k
    End If
    DeskTile ws, "dk_lite_staff", n, IIf(n = 0, "Nothing short - minimum staffing, next " & LITE_STAFF_DAYS & " days", _
             "below minimum in the next " & LITE_STAFF_DAYS & " days: " & txt & IIf(n > 3, " ...", "")), "modDesk.DeskLiteStaff", x0, yy, BTN_W, IIf(n > 0, dkOrange(), 0), IIf(n = 0, "", "See the Calendar")
    Dim bx As Double: bx = x0 + BTN_W + COLGAP
    DeskLabel ws, "dk_lite_bars", "OT hours charged this pay period, by unit", bx, yy
    ws.Shapes("dk_lite_bars").Width = 300
    Dim uv As Variant: uv = UnitChargedThisPeriod()
    If IsArray(uv) Then
        Dim mx As Double, by As Double, bw As Double, nCells As Long, s As Shape
        For k = 1 To UBound(uv, 1)
            If uv(k, 2) > mx Then mx = uv(k, 2)
        Next k
        bw = bandW - BTN_W - COLGAP - 170: nCells = 20
        by = yy + 16
        For k = 1 To UBound(uv, 1)
            DeskLabel ws, "dk_lite_u" & k, uv(k, 1), bx, by
            ws.Shapes("dk_lite_u" & k).Width = 110
            Dim nc As Long: nc = BarCellsCore(CDbl(uv(k, 2)), mx, nCells)
            If nc > 0 Then Set s = DeskShape(ws, "dk_lite_bar" & k, "", bx + 112, by + 1, nc * bw / nCells, 10, dkNavy())
            DeskLabel ws, "dk_lite_uh" & k, Format$(uv(k, 2), "0") & " h", bx + 116 + nc * bw / nCells, by
            by = by + 14
        Next k
        If by > yy + TILE_H Then yy = by - TILE_H
    Else
        DeskLabel ws, "dk_lite_nobars", "nothing charged yet this period", bx, yy + 16
    End If
    yy = yy + TILE_H + GAP
    ' the reports: a button shows only once its report exists (brief rule L4) - Staffing outlook and Look up a person today
    DeskBtn ws, "dk_lite_outlook", "Staffing outlook (calendar)", "modDesk.DeskCalendar", x0, yy, dkNavy(), tw
    DeskBtn ws, "dk_lite_person2", "Look up a person (history)", "modDesk.DeskHistory", x0 + (tw + GAP), yy, dkNavy(), tw
    yy = yy + BTN_H + 10
    DrawLiteDesk = yy
End Function

' the "?" on each Lite band: Application.Caller is the clicked shape's name (dk_q_<band>)
Public Sub DeskLiteHelp()
    On Error GoTo Oops
    Dim who As String
    On Error Resume Next
    who = CStr(Application.Caller)
    On Error GoTo Oops
    MsgBox LiteHelpText(Mid$(who, 6)), vbInformation, "Admin Lite Desk - what this band means"
    Exit Sub
Oops:
    ReportError "DeskLiteHelp"
End Sub

' one plain line per card; the same words go on the Admin Lite cheat sheet (docs/admin-lite-cheat-sheet.md)
Public Function LiteHelpText(ByVal key As String) As String
    Select Case LCase$(key)
        Case "decide"
            LiteHelpText = "Member requests: time off or a switch a member asked for from the members' file - book it (double-click) or decline it (right-click) in Needs Attention." & vbNewLine & _
                           "Posting exceptions: an open posting nobody signed up for, closed or closing within 24 h - the Award pop-up offers No takers, Cancel or Award Out of Dept." & vbNewLine & _
                           "Corrections asked for: a receipt answered No or Not sure, or an open dispute - Fix an entry voids the entry (a reason is required) and re-enters it." & vbNewLine & _
                           "Returning from leave: a leave that ends in the next " & DESK_AHEAD_DAYS & " days, or an open-ended one - Leave > Returning books the return and the top-up."
        Case "book"
            LiteHelpText = "Book time off: vacation, personal days, a sick day booked ahead (the calendar line and the cover come with it)." & vbNewLine & _
                           "Leave: STD, LOA, parental - going, and the return with its top-up." & vbNewLine & _
                           "Add a need: an uncovered shift on the calendar with no posting yet - right-click the day." & vbNewLine & _
                           "Change a person: a new hire, a unit move, a rename, a removal - the Employee pop-up."
        Case "daily"
            LiteHelpText = "Only when " & PRIMARY_ADMIN_NAME & " is off: 1 Import Kronos (the morning report), 2 Clear Needs Attention (every line looked at), 3 Publish the members' file. " & _
                           "A step turns green once anyone has done it today; when all three are done the band is one green line naming who and when."
        Case "staff"
            LiteHelpText = "Below minimum: the shifts in the next " & LITE_STAFF_DAYS & " days where the people on, less those off, plus the covers, fall under the Min Staff table - the Calendar shows the day." & vbNewLine & _
                           "OT by unit: the hours charged in the OT Log this pay period, per unit (display only). Staffing outlook opens the Calendar; Look up a person the employee history."
        Case Else
            LiteHelpText = "The Admin Lite Desk shows your jobs in four fixed bands; Show full admin Desk has everything else. Nothing here changes what you may do."
    End Select
End Function

' v27 (item 2): the admin's "This morning" tile - the integrity result of this open, today's backup, the Error Log
' rows nobody has read; red for a PROBLEM or a missing backup, orange for unread errors, grey when all is well
Private Sub MorningCheck(ByRef txt As String, ByRef col As Long, ByRef verb As String)
    Dim iTxt As String: iTxt = IntegrityStatusText()
    If iTxt = "" Then iTxt = "Integrity: not run"
    iTxt = Replace(iTxt, "Integrity: ", "Integrity ")
    Dim bk As Long: bk = BackupTodayState()
    Dim nErr As Long: nErr = UnseenErrorCount()
    txt = iTxt & "; " & Choose(bk + 2, "no backups in this copy", "NO backup today", "backup taken") & "; " & _
          IIf(nErr = 0, "no new errors", nErr & " new error" & IIf(nErr = 1, "", "s"))
    If IntegrityHasProblems() Then
        col = dkRed(): verb = "Open the Integrity sheet"
    ElseIf bk = 0 Then
        col = dkRed(): verb = "Why no backup?"
    ElseIf nErr > 0 Then
        col = dkOrange(): verb = "Open the Error Log"
    Else
        col = 0: verb = ""
    End If
End Sub

' v27 (item 2): the morning tile's click - the first thing that needs looking at
Public Sub DeskMorning()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If IntegrityHasProblems() Then IntegrityCheck: Exit Sub
    If BackupTodayState() = 0 Then
        Tell "Today's backup copy is not in:" & vbNewLine & BackupFolderPath() & vbNewLine & vbNewLine & _
             "It is taken on the first writable open of the day; the folder may have been unreachable (the open said so). " & _
             "Close and reopen the roster once the folder is back.", vbExclamation, "Backup"
        Exit Sub
    End If
    If UnseenErrorCount() > 0 Then ShowErrorLog: Exit Sub
    IntegrityCheck
    Exit Sub
Oops:
    ReportError "DeskMorning"
End Sub

' a TODAY tile: the figure big, the next thing under it, then the verb a click takes (v27); grey when there is nothing to do
Private Sub DeskTile(ByVal ws As Worksheet, ByVal nm As String, ByVal n As Long, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal col As Long, ByVal verb As String)
    Dim active As Boolean: active = (col <> 0)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, TILE_H, IIf(active, col, dkGrey()))
    Dim numTxt As String: numTxt = CStr(n)
    On Error Resume Next                            ' never raises
    With s.TextFrame2.TextRange
        .Text = numTxt & vbCr & cap & IIf(verb <> "", vbCr & verb, "")
        .ParagraphFormat.Alignment = msoAlignLeft
        .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0
        .Font.Name = DK_FONT_B
        .Font.Size = 11
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = IIf(active, vbWhite, dkMuted())
        .Characters(1, Len(numTxt)).Font.Size = 20
        If verb <> "" Then .Characters(Len(numTxt) + Len(cap) + 3, Len(verb)).Font.Size = 10
    End With
End Sub

' the members'-file tile: a caption and a text instead of a figure, then the verb (v27); mac "" = information only
Private Sub DeskTileText(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal txt As String, ByVal mac As String, _
                         ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal col As Long, ByVal verb As String)
    Dim active As Boolean: active = (col <> 0)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, TILE_H, IIf(active, col, dkGrey()))
    On Error Resume Next                            ' never raises
    With s.TextFrame2.TextRange
        .Text = cap & vbCr & txt & IIf(verb <> "", vbCr & verb, "")
        .ParagraphFormat.Alignment = msoAlignLeft
        .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0
        .Font.Name = DK_FONT_B
        .Font.Size = 11
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = IIf(active, vbWhite, dkMuted())
        .Characters(1, Len(cap)).Font.Size = 13
        If verb <> "" Then .Characters(Len(cap) + Len(txt) + 3, Len(verb)).Font.Size = 10
    End With
End Sub

' a small flat link
Private Sub DeskLink(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal w As Double)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, 22, vbWhite)
    On Error Resume Next                            ' never raises
    With s.TextFrame2.TextRange
        .Text = cap
        .ParagraphFormat.Alignment = msoAlignLeft
        .Font.Name = DK_FONT_B
        .Font.Size = 10
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = dkNavy()
    End With
End Sub

' ====================================================================
'  SHOWING THE DESK; THE VIEW
' ====================================================================
' The Desk drawn and brought to the front (the open step for a Desk user; ShowDesk after a view switch).
' Activating it fires SheetActivate -> DeskSheetActivated -> DrawDesk, so the draw is done here only when
' the event can't do it (the Desk already in front, or events off) - never twice (review).
Public Sub ShowDeskSheet()
    If Not (gAdminUnlocked Or IsSupervisor()) Then Exit Sub   ' v25.5 review: a plain user has no Desk (Alt+F8 too)
    Dim ws As Worksheet: Set ws = DeskSheet()
    SetSheetVisible ws, xlSheetVisible
    If ActiveSheet Is ws Or Not Application.EnableEvents Then
        DrawDesk
    Else
        ws.Activate
    End If
    On Error Resume Next                            ' never raises
    ws.Range("A1").Select
End Sub

' ThisWorkbook's open step 11 (v28; 10 before): an admin Desk user and (SUPERVISOR_DESK) every supervisor on SUPERVISOR_USERS land on
' the Desk (v25.5: a plain user lands on the roster), the primary admin on the roster; a Desk that can't be made or
' brought to the front leaves the user on the roster (DrawDesk never raises: it writes its own red line on the
' sheet), and a 120 % left behind is put back
Public Sub DeskOpenLanding()
    On Error Resume Next
    gDeskShowFull = False                            ' v29: a lite login lands on the Lite Desk every open
    If gDeskMode Or IsLiteUser() Or (SUPERVISOR_DESK And Not gAdminUnlocked And IsSupervisor()) Then      ' v29: + a lite login (on the full admin view)
        ShowDeskSheet
        If Err.Number <> 0 Then
            LogError "DeskOpenLanding", Err.Number, Err.Description
            Err.Clear
            RosterSheet().Activate
        End If
    Else
        RosterSheet().Activate
        DeskZoomReset
    End If
End Sub

' A Desk user's 120 % is saved in the file with the sheet: anyone else who lands on that sheet gets 100 % back
' (review: Excel keeps each sheet's zoom in the file; activating an already-active sheet fires no event)
Public Sub DeskZoomReset()
    On Error Resume Next                            ' never raises
    If gDeskMode Then Exit Sub
    Select Case ActiveSheet.Name
        Case ROSTER_SHEET, CAL_SHEET
            If ActiveWindow.Zoom = 120 Then ActiveWindow.Zoom = 100
    End Select
End Sub

' "< Desk" on the master bar, or Alt+F8: the Desk view for the session (admins only)
Public Sub ShowDesk()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    ApplyView True, True
    ShowDeskSheet
    Exit Sub
Oops:
    ReportError "ShowDesk"
End Sub

' "Show everything": the full master layout for the session; "< Desk" on the bar comes back
Public Sub DeskShowAll()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    ApplyView True, False
    RosterSheet().Activate
    Exit Sub
Oops:
    ReportError "DeskShowAll"
End Sub

' v27.1 (Graeme, Oct 22): an admin sees the Desk a supervisor on SUPERVISOR_USERS lands on, and comes back with the same
' link. This session only (a module flag: gone when the file reopens). The buttons work as normal - they are the
' supervisor's jobs, which an admin may do anyway; the red PREVIEW pill says which Desk it is. Nothing a supervisor sees
' changes. Turning it ON needs an admin; turning it off never asks.
Public Sub DeskPreview()
    On Error GoTo Oops
    If Not mPreview Then
        If Not RequireAdmin() Then Exit Sub
        If Not gAdminUnlocked Then
            MsgBox "The preview is for the admin Desk: open the Desk in the admin view first.", vbInformation, "Desk"
            Exit Sub
        End If
    End If
    mPreview = Not mPreview
    mPreviewLite = False                             ' v29: one preview at a time
    DrawDesk
    If ActiveSheet.Name <> DESK_SHEET Then ShowDeskSheet       ' (review V3) run from Alt+F8 elsewhere: show the Desk it changed
    Exit Sub
Oops:
    ReportError "DeskPreview"
End Sub

' v29: "See Desk as: Admin Lite" - an admin looks at the Desk a lite login lands on (and comes back with the same link).
' This session only. The buttons work as normal: they are admin jobs. For a lite login itself the link just redraws Lite.
Public Sub DeskPreviewLite()
    On Error GoTo Oops
    If Not mPreviewLite Then
        If Not RequireAdmin() Then Exit Sub
        If Not gAdminUnlocked Then
            MsgBox "The preview is for the admin Desk: open the Desk in the admin view first.", vbInformation, "Desk"
            Exit Sub
        End If
    End If
    mPreviewLite = Not mPreviewLite
    mPreview = False
    gDeskShowFull = False
    DrawDesk
    If ActiveSheet.Name <> DESK_SHEET Then ShowDeskSheet
    Exit Sub
Oops:
    ReportError "DeskPreviewLite"
End Sub

' v29: "< Back to my Desk" while either preview is on
Public Sub DeskPreviewOff()
    On Error GoTo Oops
    mPreview = False: mPreviewLite = False: gDeskShowFull = False
    DrawDesk
    If ActiveSheet.Name <> DESK_SHEET Then ShowDeskSheet
    Exit Sub
Oops:
    ReportError "DeskPreviewOff"
End Sub

' v29: a lite login's "Show full admin Desk" / "< Back to my Desk" - the full admin drawing on the same sheet, this session
' only (the next open lands on Lite again). Rights are unchanged either way.
Public Sub DeskShowFull()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    gDeskShowFull = Not gDeskShowFull
    DrawDesk
    If ActiveSheet.Name <> DESK_SHEET Then ShowDeskSheet
    Exit Sub
Oops:
    ReportError "DeskShowFull"
End Sub

' v29: the Posting exceptions card - the Award pop-up on the exception that has waited longest (as CalMenuAward loads a
' posting: the pop-up's list, pre-picked), where No takers / Cancel / Award Out of Dept are offered
Public Sub DeskLiteAward()
    On Error GoTo Oops
    GuideTick 3, True                                ' the Lite strip: a posting exception looked at
    Dim id As String, what As String
    If PostingExceptionCount(id, what) = 0 Or id = "" Then ShowPostingForm: Exit Sub
    If Not RequireSupervisor() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim ids As Variant, ch As Variant, i As Long, idx As Long
    idx = -1
    ch = OpenPostingChoices(ids)
    If IsArray(ids) Then
        For i = LBound(ids) To UBound(ids)
            If StrComp(CStr(ids(i)), id, vbTextCompare) = 0 Then idx = i - LBound(ids)
        Next i
    End If
    Load frmPostingEntry
    If idx >= 0 Then
        On Error Resume Next
        frmPostingEntry.Controls("cmbPosting").ListIndex = idx
        On Error GoTo Oops
    End If
    frmPostingEntry.Show
    RefreshCalendar
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskLiteAward"
End Sub

' v29: the Corrections card - Fix an entry (item 4: the guided card, modLite.ShowFixEntry); the Desk redrawn after
Public Sub DeskFixEntry()
    On Error GoTo Oops
    GuideTick 4, True                                ' the Lite strip: a correction looked at
    ShowFixEntry
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskFixEntry"
End Sub

' v29: "Add a need" - the Calendar on today; a need is added from a day's right-click menu (CalMenuAddNeed), so the
' status bar says where
Public Sub DeskAddNeed()
    On Error GoTo Oops
    CalShowDate Date
    ShowBarText "Add a need: right-click the day > Add a need (no posting yet)..."
    Exit Sub
Oops:
    ReportError "DeskAddNeed"
End Sub

' v29: "Change a person" - the Employee pop-up (add / move / rename / remove; admin)
Public Sub DeskPerson()
    On Error GoTo Oops
    ShowEmployeeForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskPerson"
End Sub

' v29: the below-minimum card - the Calendar on the first short day (today when none)
Public Sub DeskLiteStaff()
    On Error GoTo Oops
    StaffSweepBegin Date
    Dim sl As Variant: sl = StaffShortList(Date, LITE_STAFF_DAYS, "")
    If IsArray(sl) Then
        CalShowDate CDate(sl(1, 2)), CStr(sl(1, 1))
    Else
        CalShowDate Date
    End If
    Exit Sub
Oops:
    ReportError "DeskLiteStaff"
End Sub

' Called from ThisWorkbook's SheetActivate: the Desk is refreshed every time anyone looks at it (its figures
' must never be stale); an admin Desk user sees the roster and the Calendar at 120 %; anyone else landing on
' those sheets at a left-behind 120 % gets 100 % back ONCE per sheet per session (re-review: a user who picks
' 120 % himself afterwards is left alone).
Public Sub DeskSheetActivated(ByVal Sh As Object)
    On Error Resume Next                            ' never raises
    Select Case Sh.Name
        Case DESK_SHEET
            DrawDesk
            If ActiveWindow.Zoom <> 100 Then ActiveWindow.Zoom = 100
        Case ROSTER_SHEET, CAL_SHEET
            If gDeskMode Then
                If ActiveWindow.Zoom <> 120 Then ActiveWindow.Zoom = 120
            ElseIf InStr(mZoomChecked, "|" & Sh.Name & "|") = 0 Then
                mZoomChecked = mZoomChecked & "|" & Sh.Name & "|"
                If ActiveWindow.Zoom = 120 Then ActiveWindow.Zoom = 100
            End If
    End Select
End Sub

' After a pop-up the Desk opened returns, its tiles are recounted when the Desk is what's on screen
Private Sub DeskRefreshIfShown()
    On Error Resume Next                            ' never raises
    If StrComp(ActiveSheet.Name, DESK_SHEET, vbTextCompare) = 0 Then DrawDesk
End Sub

' ====================================================================
'  THE BUTTONS (each one the macro the bar calls; framed entry points)
' ====================================================================
Public Sub DeskSick()
    On Error GoTo Oops
    ShowSickCallForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskSick"
End Sub

Public Sub DeskCallout()
    On Error GoTo Oops
    GuideTick 2                                     ' v27 (item 27): the first days strip
    ShowOvertimeForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskCallout"
End Sub

Public Sub DeskCalendar()
    On Error GoTo Oops
    ' v27 (item 15, B2): the cover tile lands on the next uncovered shift's month and group (modCalendar.NextOpenSlot);
    ' with nothing to cover, the Calendar as it was
    GuideTick 3                                     ' v27 (item 27): the first days strip
    Dim d As Date, sh As String, unit As String
    If NextOpenSlot(DESK_AHEAD_DAYS, d, sh, unit) Then CalShowDate d, unit: Exit Sub
    ShowCalendar
    Exit Sub
Oops:
    ReportError "DeskCalendar"
End Sub

Public Sub DeskNeeds()
    On Error GoTo Oops
    ShowNeedsAttention
    Exit Sub
Oops:
    ReportError "DeskNeeds"
End Sub

' v28 (item 1): the requests tile - the folder read again, then Needs Attention (its MEMBERS' REQUESTS section)
Public Sub DeskRequests()
    On Error GoTo Oops
    GuideTick 2, True                                ' v29: the Lite strip - a member's request looked at
    If REQUESTS_ON And IsAdmin() And IsLiveRoster() Then ReadRequests False
    ShowNeedsAttention
    Exit Sub
Oops:
    ReportError "DeskRequests"
End Sub

Public Sub DeskPost()
    On Error GoTo Oops
    ShowNewPostingForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskPost"
End Sub

Public Sub DeskAward()
    On Error GoTo Oops
    ShowPostingForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskAward"
End Sub

Public Sub DeskTimeOff()
    On Error GoTo Oops
    GuideTick 1, True                                ' v29: the Lite strip - time off booked
    ShowTimeOffForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskTimeOff"
End Sub

Public Sub DeskSwitch()
    On Error GoTo Oops
    ShowShiftSwitchForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskSwitch"
End Sub

Public Sub DeskLeave()
    On Error GoTo Oops
    ShowLeaveForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskLeave"
End Sub

Public Sub DeskPublish()
    On Error GoTo Oops
    PublishMemberView
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskPublish"
End Sub

Public Sub DeskPrint()
    On Error GoTo Oops
    Call_Out
    Exit Sub
Oops:
    ReportError "DeskPrint"
End Sub

' v26.5.5 (5.4): the Practice button - a throw-away copy with a red banner (modSystem.MakePracticeCopy)
Public Sub DeskPractice()
    On Error GoTo Oops
    GuideTick 1                                     ' v27 (item 27): the first days strip
    MakePracticeCopy
    Exit Sub
Oops:
    ReportError "DeskPractice"
End Sub

Public Sub DeskRoster()
    On Error GoTo Oops
    RosterSheet().Activate
    Exit Sub
Oops:
    ReportError "DeskRoster"
End Sub

' v26: the employee history report (modReports.EmployeeHistory; its own error frame)
Public Sub DeskHistory()
    On Error GoTo Oops
    EmployeeHistory
    Exit Sub
Oops:
    ReportError "DeskHistory"
End Sub

' v27 (item 4): the Reconcile sheet - the paper call order beside the roster's (admin)
Public Sub DeskReconcile()
    On Error GoTo Oops
    ShowReconcile
    Exit Sub
Oops:
    ReportError "DeskReconcile"
End Sub

' v27 (item 26): the quarterly Stats report (admin)
' v27 (item 2, review): the Kronos pill's click - the import, then the Desk redrawn with the new figures
Public Sub DeskKronos()
    On Error GoTo Oops
    SyncKronosData
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskKronos"
End Sub

' v27 (item 33): the Holidays and Min Staff pop-ups behind one button (admin)
Public Sub DeskTables()
    On Error GoTo Oops
    Select Case MsgBox("Yes = the Holidays (the weekday off for a weekend holiday)." & vbNewLine & "No = minimum staffing (per unit and shift, and the unit's note).", vbYesNoCancel + vbQuestion, "Holidays / minimum staffing")
        Case vbYes: ShowHolidayEditor
        Case vbNo: ShowMinStaffEditor
    End Select
    Exit Sub
Oops:
    ReportError "DeskTables"
End Sub

' v27 (item 3): the Since-you-were-last-here list
Public Sub DeskSince()
    On Error GoTo Oops
    SinceLastHere
    Exit Sub
Oops:
    ReportError "DeskSince"
End Sub

' v27 (item 25): the shift handover page and the payback ledger (supervisors and admins)
Public Sub DeskHandover()
    On Error GoTo Oops
    ShiftHandover
    Exit Sub
Oops:
    ReportError "DeskHandover"
End Sub

Public Sub DeskLedger()
    On Error GoTo Oops
    SwitchLedger
    Exit Sub
Oops:
    ReportError "DeskLedger"
End Sub

Public Sub DeskStats()
    On Error GoTo Oops
    StatsReport
    Exit Sub
Oops:
    ReportError "DeskStats"
End Sub

' v25.5 (A9) "Undo my last log entry": the admin Desk; voids this login's newest entry of today through Void Entry's path
Public Sub DeskUndo()
    On Error GoTo Oops
    UndoLast
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskUndo"
End Sub

' v25.5 (A8) "Suggest an improvement": the same email draft as a fault report, with its own subject and no error line
Public Sub DeskSuggest()
    On Error GoTo Oops
    Dim body As String
    body = "A suggestion for the Call Out Roster." & vbNewLine & vbNewLine & _
           "Who: " & CurrentUser() & vbNewLine & "When: " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & vbNewLine & _
           "File: " & ThisWorkbook.Name & "  (" & CORE_VERSION & ")" & vbNewLine & vbNewLine & _
           "What would help: " & vbNewLine & vbNewLine & "Where (which pop-up or tab): " & vbNewLine
    If Not OpenMailTo(PRIMARY_ADMIN_EMAIL, "Call Out Roster: a suggestion", body) Then
        MsgBox "No mail app could be opened. Tell " & PRIMARY_ADMIN_NAME & " yourself, or send this:" & _
               vbNewLine & vbNewLine & Left$(body, 850), vbExclamation, "Suggest an improvement"
    End If
    Exit Sub
Oops:
    ReportError "DeskSuggest"
End Sub

' ====================================================================
'  HELP (v25.5, bundle item 7, Graeme Oct 18): a read-only Help sheet inside the file, drawn from the quick
'  guide - one heading and one paragraph per job at 11 pt - so help is never "not on this PC yet". The
'  Desk's "How do I ..." opens it; "< Back" returns to the Desk (or the roster); the printable guide
'  (GUIDE_PATH, a PDF) stays optional behind its own button. Drawn on first use and again after an install
'  (the version in D1 differs). Protected like every sheet; nothing is typed into it. Holds no rule.
' ====================================================================
' (HELP_SHEET / HELP_COL_W are declared at the top of the module - a Const after a procedure does not compile)

' "How do I ...": the Help sheet (the name stays: the Desk button calls it)
Public Sub DeskGuide()
    On Error GoTo Oops
    GuideTick 4                                     ' v27 (item 27): the first days strip
    ShowHelp
    Exit Sub
Oops:
    ReportError "DeskGuide"
End Sub

Private Function HelpSheet() As Worksheet
    If Not SheetExists(HELP_SHEET) Then
        Dim ws As Worksheet
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HELP_SHEET
        RelockStructure
    End If
    Set HelpSheet = ThisWorkbook.Worksheets(HELP_SHEET)
    SetTabColor HelpSheet, dkMuted()
End Function

' The sheet drawn (when missing or from an older release) and brought to the front (Alt+F8 too: framed)
Public Sub ShowHelp()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = HelpSheet()
    If SafeStr(ws.Cells(1, 4).Value) <> DESK_VERSION Then DrawHelp ws
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    On Error Resume Next                            ' never raises
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        If .Zoom <> 100 Then .Zoom = 100
        .ScrollRow = 1
    End With
    ws.Range("B1").Select
    Exit Sub
Oops:
    ReportError "ShowHelp"
End Sub

' "< Back": the Desk when this login has one, else the roster
Public Sub HelpBack()
    On Error GoTo Oops
    If SheetExists(DESK_SHEET) And (gAdminUnlocked Or IsSupervisor()) Then
        If ThisWorkbook.Worksheets(DESK_SHEET).Visible = xlSheetVisible Then ShowDeskSheet: Exit Sub
    End If
    RosterSheet().Activate                              ' re-review: never leaves anyone on the Help sheet
    Exit Sub
Oops:
    ReportError "HelpBack"
End Sub

' the printable guide (GUIDE_PATH, a PDF on the plant PC; the button shows only when it is set)
Public Sub DeskGuidePdf()
    On Error GoTo Oops
    If Trim$(GUIDE_PATH) = "" Then
        MsgBox "The printable guide isn't on this PC yet. Ask " & PRIMARY_ADMIN_NAME & " for it; the Help tab has the same jobs.", _
               vbInformation, "How do I ..."
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=GUIDE_PATH
    If Err.Number <> 0 Then
        On Error GoTo Oops
        MsgBox "The guide couldn't be opened from" & vbNewLine & GUIDE_PATH & vbNewLine & vbNewLine & _
               "Ask " & PRIMARY_ADMIN_NAME & " for it.", vbExclamation, "How do I ..."
    End If
    Exit Sub
Oops:
    ReportError "DeskGuidePdf"
End Sub

' Draws (or redraws) the Help sheet: a title band, the two buttons, then the sections. Never raises: a
' draw that fails writes its own red line (like DrawDesk) and the sheet is protected again whatever happened.
Public Sub DrawHelp(ByVal ws As Worksheet)
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Failed
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "hp_" Then ws.Shapes(i).Delete
    Next i
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
    ws.Cells.Interior.Color = vbWhite
    With ws.Cells.Font
        .Name = DK_FONT
        .Size = 11
        .Color = dkInk()
    End With
    ws.Columns(1).ColumnWidth = 2
    ws.Columns(2).ColumnWidth = HELP_COL_W
    ws.Columns(3).ColumnWidth = 2

    ' ---- row 1: the title band; row 2: what this is; row 3: the buttons ----
    ws.Rows(1).RowHeight = 34
    ws.Range("A1:C1").Interior.Color = dkNavy()
    With ws.Cells(1, 2)
        .Value = "HOW DO I ..."
        .Font.Name = DK_FONT_B: .Font.Size = 18: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(1, 4)                                  ' the version this sheet was drawn from (ShowHelp redraws on a change)
        .NumberFormat = "@"
        .Value = DESK_VERSION
        .Font.Size = 7: .Font.Color = RGB(176, 190, 197)
    End With
    ws.Rows(2).RowHeight = 20
    With ws.Cells(2, 2)
        .Value = "Each job in the roster, in plain words: what to click and what happens. Scroll down, or Ctrl+F for a word. " & _
                 "Nothing on this tab can be typed into."
        .Font.Size = 10: .Font.Color = dkMuted()
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 30
    Dim x0 As Double: x0 = ws.Cells(3, 2).Left
    Dim y0 As Double: y0 = ws.Cells(3, 2).Top + 3
    DeskBtn ws, "hp_btn_back", "<  Back", "modDesk.HelpBack", x0, y0, dkNavy(), 110
    If Trim$(GUIDE_PATH) <> "" Then DeskBtn ws, "hp_btn_pdf", "Open the printable guide (PDF)", "modDesk.DeskGuidePdf", x0 + 110 + GAP, y0, dkGrey(), 230, dkInk()
    ws.Rows(4).RowHeight = 10

    ' ---- the sections ----
    Dim h() As String, n As Long
    HelpSections h, n
    Dim r As Long: r = 5
    For i = 1 To n
        With ws.Cells(r, 2)
            .Value = h(1, i)
            .Font.Name = DK_FONT_B: .Font.Size = 12: .Font.Color = dkNavy()
            .VerticalAlignment = xlBottom
        End With
        ws.Rows(r).RowHeight = 24
        With ws.Cells(r + 1, 2)
            .NumberFormat = "@"
            .Value = h(2, i)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
        ws.Rows(r + 1).AutoFit
        ws.Rows(r + 1).RowHeight = ws.Rows(r + 1).RowHeight + 2
        ws.Rows(r + 2).RowHeight = 8
        r = r + 3
    Next i
    With ws.Cells(r, 2)
        .Value = "Call Out Roster  -  " & DESK_VERSION & ".  Something not here? Ask " & PRIMARY_ADMIN_NAME & ", or click Something went wrong on the Desk."
        .Font.Size = 9: .Font.Color = dkMuted()
    End With
    GoTo Done
Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    Resume AfterFail
AfterFail:
    On Error Resume Next
    LogError "DrawHelp", en, em
    UnlockSheet ws
    ws.Cells(1, 4).Value = ""                            ' so the next "How do I ..." draws again
    ws.Cells(2, 2).Value = "The Help tab couldn't finish drawing: " & em & "  (" & DESK_VERSION & ")"
    ws.Cells(2, 2).Font.Color = dkRed()
Done:
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
End Sub

Private Sub HelpAdd(ByRef h() As String, ByRef n As Long, ByVal head As String, ByVal body As String)
    n = n + 1
    ReDim Preserve h(1 To 2, 1 To n)
    h(1, n) = head
    h(2, n) = body
End Sub

' The text, one heading and one paragraph per job (from docs/supervisor-quick-guide.md; keep the two in step)
Private Sub HelpSections(ByRef h() As String, ByRef n As Long)
    HelpAdd h, n, "Open the file", _
        "Open the roster from its usual folder and click Enable Content if Excel asks. A notice that someone has it open for " & _
        "editing means you can look but not log: close it and come back when they are out. Every sheet is protected: the " & _
        "buttons do the writing, nothing is typed into cells. Close the file as soon as you are done so the next supervisor can get in."
    HelpAdd h, n, "The Desk", _
        "A supervisor on the list, and an admin on the Desk list, open on the Desk tab (the primary admin opens on the roster and reaches it with < Desk). " & _
        "The three pills at the top say which crews are on today, the pay period, and how old the Kronos worked hours are (red when they " & _
        "need an import). The TODAY tiles say what is waiting - shifts needing cover, postings ready to award, paybacks owed (admin), the members' file's age - " & _
        "each with the next thing and what a click does. Under them the jobs sit in three bands by WHEN you reach for them: NOW (orange: " & _
        "Someone called in sick, Fill a shift now, What needs cover), PLAN (green: Post a shift, Award a posting, and for an admin time " & _
        "off, switches and leave) and FINISH (navy: Update the members' file, Print the call-out roster, Employee history, and for an admin Reconcile with the paper roster " & _
        "and Stats). The grey row " & _
        "at the foot is help: How do I ..., Something went wrong, Suggest an improvement, Practice on a copy (and Undo for an admin). " & _
        "Full roster > goes to the roster; the Desk tab brings it back. Anyone not on the supervisor list opens on the roster and can log " & _
        "call-outs and look at the calendar."
    HelpAdd h, n, "The Admin Lite Desk", _
        "An admin on the Lite list (" & LITE_FIRST_NAME & ") opens on a simpler Desk with only his jobs, in four bands that never move: NEEDS YOUR DECISION " & _
        "(members' requests, posting exceptions - open postings nobody signed up for -, corrections asked for through the receipts, people returning from leave), " & _
        "BOOK AND CHANGE (time off, leave, add a need, change a person), DAILY CHECK (only when " & PRIMARY_ADMIN_NAME & " is off: Import Kronos, Clear Needs " & _
        "Attention, Publish the members' file - each green once anyone has done it today) and STAFFING AND OT (below minimum in the next two weeks, this " & _
        "period's OT hours by unit). The strip under the title says whether the hours, today's backup, the members' file and the system checks are fine; a red " & _
        "item says to call " & PRIMARY_ADMIN_NAME & ". A ? on each band explains its cards. Show full admin Desk > shows everything else for the session; " & _
        "< Back to my Desk returns. It changes what is on screen, never what the login may do. Any admin can look at it with See Desk as: Admin Lite."
    HelpAdd h, n, "Log a call-out (fill a shift now)", _
        "Click Fill a shift now (or Log Call-Out on the roster's bar). Pick the Department, Shift and Date, type the Hours and " & _
        "click LOAD LIST: the list is in call order and holds only the people free for that shift. The card shows who to call " & _
        "and their phone; press A (accepted), R (refused, charged x1.75), L (left message, charged x1.75) or U (unavailable, " & _
        "not charged: pick the reason, Enter records it) - or click the button - and the next person comes up. K skips for " & _
        "now (Call back), B goes back one, P is part of the hours (that closes the shift). When the shift is filled press Enter " & _
        "or click SUBMIT ALL ENTRIES and note the session number (S-0012); the box ends with what happens to the line next. " & _
        "Refused and Left Message are charged for grade 2 and below; >> means above grade 2 (called, never charged); * means " & _
        "the total includes accepted overtime not yet in Kronos; ~ means an estimate because a unit move is planned for the " & _
        "person; [max N h] (since v26.5.5) means the 16 h rule leaves them less than the offer - OT they accepted, an award " & _
        "or a switch touches this shift - so A is refused and P up to that many hours is allowed. Before 7 AM the date box " & _
        "is yesterday: the night under way is dated by the evening it started. The grey line at the bottom says how current " & _
        "the worked hours are and when the members' file was last updated."
    HelpAdd h, n, "Someone takes only part of the hours", _
        "Select them and press P (Part of the hours), then enter the hours. That closes the shift: nobody else is called, " & _
        "nobody below is charged, and the pop-up refuses a further Accepted (a refusal or a message already logged can still be " & _
        "corrected to a refusal, a message or unavailable). Anyone who refused before the partial keeps their " & _
        "charge. Click SUBMIT ALL ENTRIES."
    HelpAdd h, n, "Someone called in sick", _
        "Click Someone called in sick (Sick Call on the bar). Pick who called, check the Shift date, Shift, Unit to cover and " & _
        "Hours, and read the line under the Note box: it gives their sick-day counts and whether this shift is a paid sick day. " & _
        "SAVE + START CALL-OUT books the sick shift and opens the call-out filled in (with the hours); Save, no call-out books it " & _
        "without a call-out. Either way a box reads the booking back first (who, the date and shift, the unit, until when they " & _
        "are off the lists): Yes books it. The person is off every call list for that shift and any shift starting within 24 " & _
        "hours after it ends. Log each sick shift as its own sick call. Members see it as Off, never why. Wrong person? Only an " & _
        "admin can void it: ask " & PRIMARY_ADMIN_NAME & ". Before 7 AM the date box is yesterday (the night under way)."
    HelpAdd h, n, "Post a shift", _
        "Click Post a shift (New Posting on the bar). Fill in Unit, Shift date, Shift, Reason, Hours and Positions; Crew and " & _
        "Posting closes fill in. Leave Print it and Save PDF ticked, click CREATE POSTING, then put up the printed sheet. Sign-up " & _
        "is on the paper sheet or, from the members' file, Sign me up (since v26.5): Award Posting's LOAD ticks those and names " & _
        "late ones and withdrawals; your tick is what counts. The Reason must not say what kind of leave anyone is on: a leave word is refused. Load Posting (admin) " & _
        "shows, prints or saves a posting again."
    HelpAdd h, n, "Award a posting", _
        "Click Award a posting (Award Posting on the bar), pick the posting and click LOAD. Double-click each name that signed " & _
        "up. Someone who signed up for only part of the shift: select the name, click PARTIAL HOURS and enter the hours. Select " & _
        "who gets it and click AWARD (on / off) - anyone who signed up for the whole shift comes before a partial sign-up. Read " & _
        "the charge summary, click SUBMIT AWARD, then send the email draft that opens. People ahead of the last awarded person " & _
        "who did not sign up are charged the posting's full hours x 1.75. NO TAKERS, AWARD OUT OF DEPT and CANCEL REQUEST are " & _
        "an admin's buttons (greyed for a supervisor): ask the admin when nobody took it, when someone from another department " & _
        "signed a blank line, or when the overtime is no longer needed. Awarded wrong: ask the admin to void and reopen it."
    HelpAdd h, n, "Print the call-out roster", _
        "Click Print the call-out roster (Call Out Roster on the bar). It is redrawn in call order from the roster; answer Yes " & _
        "to print."
    HelpAdd h, n, "Practice on a copy", _
        "Click Practice on a copy (the grey row at the foot of the Desk). A separate file opens with the roster as it is right now, a red " & _
        "PRACTICE COPY banner on every sheet and COPY in its title: nothing you do in it counts, reaches the roster or the " & _
        "members' file. Log a call-out, award a posting, book a sick call - whatever you want to try. Close it when you are " & _
        "done; copies older than a day are deleted with the roster's daily backup (the first open of the day). (v26.5.5)"
    HelpAdd h, n, "Update the members' file", _
        "Click Update the members' file (Member View on the bar) after you post, award or book, so the members see it. Wait for " & _
        "Members' file published. If it asks to close the members' file open in this Excel, Yes closes it and publishes. If it " & _
        "says the file couldn't be written it is open on a plant PC: have it closed there and publish again (nothing was " & _
        "written). NOT published - something private: a calendar line or a posting reason names a kind of leave, sick or " & _
        "personal; change the words (Change the text of this line / Change the Reason of posting on the calendar) and publish " & _
        "again. You are also asked when you close the file with changes not yet published."
    HelpAdd h, n, "The calendar", _
        "Click What needs cover (Calendar on the bar). Pick the month and the unit at the top. OPEN means the line still needs " & _
        "cover; POSTED means a posting is out; a name means covered; NCN means no cover needed. Click a line: its details show " & _
        "in the row under the title. Double-click a line for its options (an admin: set coverage, change the text, " & _
        "post it). Print / PDF prints the month or saves it as a PDF as the members' copy (a sick call or a personal day reads Off)."
    HelpAdd h, n, "Needs Attention (admin)", _
        "From the calendar's Needs attention > button or the tab. It lists open lines, postings ready to award, paybacks owed, " & _
        "leaves to watch, planned unit moves and other items; double-click a row for its options. A due planned move asks " & _
        "Apply it now?: Yes applies that one move, No asks whether to cancel it."
    HelpAdd h, n, "Book time off (admin)", _
        "Click Book time off. Pick Employee, Unit and Type (VAC vacation, PER personal day, TRN training; for a leave use the " & _
        "Leave button), type From and To (last day). Vacation and personal days go on scheduled shifts only: a range that " & _
        "starts or ends on days off asks to be trimmed. Open-ended for no return date yet; Hours / shift for part of a shift. " & _
        "SAVE. To fix a booking select it under BOOKED TIME OFF and click Change end date or Void this booking."
    HelpAdd h, n, "Shift switch and payback (admin)", _
        "Click Shift switch or payback. Type the Shift date, pick the Unit, then Gives up shift and Works the shift. Read the hint " & _
        "lines: NOT ELIGIBLE stops the save, check is a warning. SAVE SWITCH. It stays on the OWED list until paid back: click " & _
        "the switch in the list, type the payback date and save. Tick log the payback now when both dates are known. A wrong " & _
        "switch: select it and click Void the selected switch."
    HelpAdd h, n, "Leave: going and returning (admin)", _
        "Click Leave. Going on leave: pick Employee and Type, type the first day off and the expected return (or leave it " & _
        "blank), SAVE LEAVE. Returning: pick who is back, type the first day back, click Work out the top-up, read the peers, " & _
        "the average and the hours before and after, then SAVE TOP-UP. A return date in the future, or after the roster's last " & _
        "pay period, is refused."
    HelpAdd h, n, "Changing a person's details (admin)", _
        "Nothing on the roster is typed into. Double-click a person's row (or click it and press Employee): the pop-up opens on Edit " & _
        "with their home unit, crew, phone, email, seniority date and unit codes. Change what you need and SAVE CHANGES: the box lists " & _
        "every change first; code changes go in the Grade Log, a crew change is dated and goes in the crew history, and nobody's hours " & _
        "move. Worked hours come from the Kronos import. A full move to another unit is Move to unit (it tops up the hours); a new " & _
        "person is Add; someone leaving is Remove."
    HelpAdd h, n, "Words on the roster (glossary)", _
        "Call-out: phoning down the list for a shift today or tomorrow (Fill a shift now). Session S-nnnn: one call-out's rows. " & _
        "Posting P-nnnn: a shift put up for sign-ups (the old paper posting sheet); Award: picking who works it; No takers: nobody signed; " & _
        "Out of department: awarded to someone with a code in another unit. Total Hours: Kronos worked hours + Refused + provisional " & _
        "OT; the lowest is called first within a grade. Refused: the hours charged for a refusal or a message left, x 1.75, grade 2 and " & _
        "below only. Provisional (*): OT accepted but not in Kronos yet. ~: the hours a planned unit move will give. Exempt: never charged " & _
        "on that posting (moved, left, restricted). NCN: no coverage needed. S/S: a shift switch; a payback is the shift owed back. " & _
        "PHA: hazard-analysis coverage. The 16 h rule: no more than 16 h in a row with 11 h rest before the next work. " & _
        "Calendar: " & CalLegendText() & ".  Log: " & OutcomeLegendText() & "."
    HelpAdd h, n, "Find a person", _
        "Supervisors and admins: click Find a person on the roster's bar and type part of the name: the roster jumps to the row (a unit filter that " & _
        "hides it is cleared) and the status bar at the bottom says where the person stands in their home unit's call order " & _
        "today, for example 4 of 19 in the Phosphine call order. Click any roster row for the person's refused hours by kind."
    HelpAdd h, n, "Minimum staffing (admin)", _
        "The Min Staff sheet (admin) holds the minimum operators per unit and shift; blank = no minimum on that shift (an admin changes it, " & _
        "and the unit's note the call-out pop-up shows, through the Desk's Holidays / minimum staffing). For today and " & _
        "the next 13 days the calendar's date cell shows a red N short flag when any unit's day or night shift is below it, whatever unit group the " & _
        "calendar shows (click the date for which), and Needs Attention lists each short shift under BELOW MINIMUM STAFFING, every unit " & _
        "(double-click to open the calendar on that date). Staffed " & _
        "= the people on the crew that is on, holding a code in the unit, minus those booked off or sick for that shift or off in a " & _
        "switch, plus every coverage line a named person covers (a switch's own S/S line counts the one working it); an OPEN, " & _
        "NO TAKERS, POSTED or NCN line adds nobody. A partial cover (Joe D 6h) is 6/12 of a person: the shares on the shift are added " & _
        "and rounded down, and the date's details and Needs Attention say (part) when a fraction was dropped. It only reads: fix a short shift the usual way " & _
        "(a call-out, a posting, a switch)."
    HelpAdd h, n, "The two-week board", _
        "Board on the Calendar's top row (supervisors and admins): the Board tab is drawn for today and the next 13 days, one page " & _
        "per unit group, one row per date with the crews on and the Days and Nights lines as the calendar shows them, OPEN lines " & _
        "in red; a PDF goes beside the month calendars, then it asks whether to print. Nothing is changed: it is a picture of the " & _
        "calendar for the board in the control room."
    HelpAdd h, n, "Since you were last here", _
        "The last TODAY tile counts what OTHER logins logged since you last opened the roster - call-out sessions, awards and " & _
        "adjustments (one per Ref ID), time off and sick calls, switches. Click it for the list on the Handover tab (no logins, " & _
        "leave shown as Leave); your own entries are left out. The time is kept per login in the roster file and moves on " & _
        "when the file is saved, so after a look-only visit the tile still counts from your last SAVED session."
    HelpAdd h, n, "Shift handover", _
        "Shift handover (last 12 h) in the Desk's FINISH band draws one page of everything logged in the last 12 hours - the OT Log rows " & _
        "(call-outs, awards, adjustments, with a void shown in Status), the time off, sick calls and leave booked (leave shows as Leave), " & _
        "and the shift switches entered - and offers to print it. It changes nothing; it is drawn again from the data each time."
    HelpAdd h, n, "Paybacks owed (the ledger)", _
        "Paybacks owed (ledger) in the Desk's PLAN band lists every shift switch still waiting for its payback: whoever gave up a " & _
        "shift owes the one who worked it, oldest first, the ones over the warning age in amber. Type part of a name for that " & _
        "person's statement (what they owe and what they are owed) and save it as a PDF to hand them."
    HelpAdd h, n, "Employee history (report)", _
        "Employee history on the bar's FIND group or the Desk: part of a name, then a From and To date (the boxes open on the whole roster " & _
        "year; the member-safe report takes the pay periods starting in the range, the full one the shift dates), and the History Report tab lists every OT Log entry for that person in the range, newest first: call-outs with " & _
        "their place in the call order, postings, adjustments, the hours charged and a Why line on every charged row (for example Charged 21 h: 3rd " & _
        "called, Refused, 12 h offered x 1.75), then per pay period the charged hours, the posting refusals, the call-out refusals and the " & _
        "premium hours worked from the Kronos export (v26.5), and the total. An admin is asked for the full " & _
        "report (on the History Report (admin) tab: who logged each entry and when, a Why on every row, void rows struck through) or the " & _
        "member-safe one (the same rows and columns the members' My history shows); a supervisor gets the member-safe one. Save a PDF as well? puts it in the " & _
        "reports folder (beside the roster until go-live). Click any OT Log row for its Why on the status bar."
    HelpAdd h, n, "Reconcile with the paper roster (admin)", _
        "While paper runs alongside the roster, Reconcile (the admin Desk's FINISH band) puts the paper call order beside the roster's: pick the unit and the day the paper " & _
        "was good for, type the paper's names top to bottom in column B, click Compare. The roster's order for that unit is drawn beside them (Total Hours as of now; the crews and a " & _
        "planned move as of that day) and every position that differs says why - where the roster puts the paper's person and which key decided it (grade code, Total Hours, seniority). " & _
        "Type a Decision (Paper wrong / Rule missing / Data error / Agreed) and a Note per line; Save PDF files the page beside the roster for the Nov 15 pack. Nothing here changes data."
    HelpAdd h, n, "Stats (quarterly report, admin)", _
        "Stats (the admin Desk's FINISH band) draws the Stats tab: unit by quarter plus the year to date, in four groups - call-outs (sessions, filled, calls per fill, the outcomes, " & _
        "hours offered / accepted / charged), postings (posted, awarded, no takers, cancelled, sign-ups, days to award), sick days (shifts, the rate per scheduled shift from the rotation, nights, " & _
        "weekends, holidays) and fairness today (the spread of Total Hours and of worked premium hours per unit, the refusal rate). A PDF beside the roster on request. Figures are honest from the first full quarter after go-live."
    HelpAdd h, n, "Keyboard shortcuts", _
        "While the roster is the active workbook: Ctrl+Shift+S opens Sick Call, Ctrl+Shift+C the call-out, Ctrl+Shift+K the " & _
        "calendar, Ctrl+Shift+D the Desk. In the call-out pop-up the letters A R L U P record the outcome, K skips, B goes back, " & _
        "Enter submits, Esc closes. Alt+F8 > ShowWhatsNew brings back the what's-new box that shows once after an install."
    HelpAdd h, n, "Typing a date", _
        "Every date box takes Oct 12, 12 Oct, Oct 12, 2026 or 10/12 (month then day) and shows the full date it understood. With " & _
        "no year, a date more than three months back means next year and one more than nine months ahead means last year. A " & _
        "time typed after the date is ignored. The box says when it can't read what you typed."
    HelpAdd h, n, "Asking for a void", _
        "Void Entry is an admin button. To correct a logged call-out or award, tell the admin the number (S-0012, P-0003) and " & _
        "what was wrong: the admin voids it (the entry stays in the log, marked VOID) and you log the correct entry. An admin " & _
        "who just logged something wrong clicks Undo my last log entry (the admin Desk, or Undo last entry on the bar): it offers " & _
        "to void their own newest OT Log entry of today, with the same confirm and reason. It undoes log entries only: a sick " & _
        "call, time off, a switch or a new posting is taken back in its own pop-up."
    HelpAdd h, n, "If a red box appears", _
        "This didn't finish means something stopped part-way. The box names what you were doing and an Error Log line. Nothing " & _
        "you did before it is lost. Try once more; if it fails again, click Something went wrong on the Desk (supervisors and admins; anyone else tells " & _
        PRIMARY_ADMIN_NAME & " the line number): an email to " & _
        PRIMARY_ADMIN_NAME & " opens with the error line in it (or a box to photograph). Saving is blocked in red means close " & _
        "WITHOUT saving, reopen, check your last entry, and tell the admin."
    HelpAdd h, n, "Close the file", _
        "Saving is automatic after each entry and again when you close. Just close the file. The members' file question on " & _
        "close: Yes publishes before closing."
    HelpAdd h, n, "Words you will see", _
        "Call order: grade code first, then the lowest Total Hours, then plant seniority - the order everyone is called in. " & _
        "Total Hours: worked (Kronos) + the hours charged in the OT Log this year whose pay period has started (refusals, left " & _
        "messages, adjustments; a refusal for a shift in a later pay period counts from that period) + accepted " & _
        "overtime not yet in Kronos (*). Charged: refused or left-message " & _
        "hours x 1.75 added to Total Hours, grade 2 and below only (>> = above grade 2, called but never charged). ~ = an " & _
        "estimate because a unit move is planned for the person. ? = left a message. S/S = working a shift switch that " & _
        "shift. NCN = no cover needed. PHA = hazard-analysis coverage. OPEN = the line still needs cover; POSTED = a " & _
        "posting is out for it. DAYS = the Monday-to-Friday day crew. A sick call or a personal day is Off to the " & _
        "members, never why."
End Sub

' v27 (item 20, B14): toggles this login's pop-up scale (UiScale_<login>) and redraws the Desk
Public Sub DeskLargerType()
    On Error GoTo Oops
    Dim larger As Boolean: larger = (UiScale() <= 1.05)
    SetUiScaleForMe larger
    Dim roomTxt As String                            ' (review L9) the call-out pop-up (615 pt high) may have no room for it
    If larger And UiScale(615) <= 1 Then roomTxt = vbNewLine & vbNewLine & "The call-out pop-up has no room for it on this screen and stays at its size; the others grow."
    Tell IIf(larger, "Your pop-ups open at " & CLng(UI_LARGER_SCALE * 100) & " % (this login, this roster; kept when the file is saved - a read-only copy forgets it). Click again for normal type." & roomTxt, "Your pop-ups open at the normal size again."), vbInformation, "Larger type"
    DrawDesk
    Exit Sub
Oops:
    ReportError "DeskLargerType"
End Sub

' v27 (item 24): the two admin posting jobs behind one button
Public Sub DeskPostingJobs()
    On Error GoTo Oops
    Select Case MsgBox("Yes = withdraw an award (its Awarded rows voided, the line opens again)." & vbNewLine & "No = edit an open posting (hours, date, positions).", vbYesNoCancel + vbQuestion, "Posting jobs")
        Case vbYes: WithdrawAward
        Case vbNo: EditPosting
    End Select
    Exit Sub
Oops:
    ReportError "DeskPostingJobs"
End Sub

' "Something went wrong": an email draft to the primary admin with the newest Error Log line in it
Public Sub DeskReport()
    On Error GoTo Oops
    Dim body As String
    ' the prompt comes BEFORE the error line, and the line is cut, so a long error never pushes the prompt out (review)
    body = "Something went wrong on the Call Out Roster." & vbNewLine & vbNewLine & _
           "Who: " & CurrentUser() & vbNewLine & "When: " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & vbNewLine & _
           "File: " & ThisWorkbook.Name & "  (" & CORE_VERSION & ")" & vbNewLine & vbNewLine & _
           "What I was doing: " & vbNewLine & vbNewLine & _
           "Newest line in the Error Log:" & vbNewLine & Left$(NewestErrorLine(), 350) & vbNewLine
    If Not OpenMailTo(PRIMARY_ADMIN_EMAIL, "Call Out Roster: something went wrong", body) Then
        MsgBox "No mail app could be opened. Send this to " & PRIMARY_ADMIN_NAME & " yourself (a photo is fine):" & _
               vbNewLine & vbNewLine & Left$(body, 850), vbExclamation, "Something went wrong"
    End If
    Exit Sub
Oops:
    ReportError "DeskReport"
End Sub

' ====================================================================
'  v27 (item 27, 5.5): THE FIRST DAYS STRIP
' ====================================================================
' For a login's first three Desk sessions, a line under the links: the four things to try, each ticked once done
' (practice copy, a call-out, the Calendar, How do I ...), and "Hide this". Kept per login in hidden names
' (GuideOpens_<login> counted once a session, GuideDone_<login> = the steps done or "hidden"); a saved workbook
' stays marked saved. Never raises.
Private Function GuideGet(ByVal prefix As String) As String
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(LoginNameKey(prefix))
    If nmObj Is Nothing Then Exit Function
    GuideGet = Replace(Replace(Mid$(nmObj.RefersTo, 2), """", ""), "'", "")
End Function

Private Sub GuideSet(ByVal prefix As String, ByVal v As String)
    On Error Resume Next                            ' never raises
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    Dim key As String: key = LoginNameKey(prefix)
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:="=""" & v & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & v & """"
    End If
    If wasSaved Then ThisWorkbook.Saved = True
End Sub

' a step done (1 practice copy, 2 a call-out, 3 the Calendar, 4 How do I ...); v29: lite = the Lite Desk's own strip
' (1 book time off, 2 a member's request, 3 a posting exception, 4 fix an entry, 5 the daily check), its own hidden names
Private Sub GuideTick(ByVal k As Long, Optional ByVal lite As Boolean = False)
    If lite And Not IsLiteUser() Then Exit Sub        ' an admin previewing Lite ticks nothing
    Dim pfx As String: pfx = IIf(lite, "GuideLite", "Guide")
    Dim s As String: s = GuideGet(pfx & "Done")
    If InStr(s, "hidden") > 0 Or InStr(s, "|" & k & "|") > 0 Then Exit Sub
    GuideSet pfx & "Done", IIf(s = "", "|", s) & k & "|"
End Sub

' True while the strip shows: the first three sessions, not hidden. The session is counted on its first Desk draw.
' v29: the Lite strip counts its own sessions (a lite login's three ordinary sessions are long used up).
Private Function GuideShows(Optional ByVal lite As Boolean = False) As Boolean
    On Error Resume Next                            ' never raises
    Dim pfx As String: pfx = IIf(lite, "GuideLite", "Guide")
    If lite And Not IsLiteUser() Then Exit Function   ' the preview shows no strip
    If InStr(GuideGet(pfx & "Done"), "hidden") > 0 Then Exit Function
    If Not mGuideCounted Then
        mGuideCounted = True
        GuideSet pfx & "Opens", CStr(Val(GuideGet(pfx & "Opens")) + 1)
    End If
    GuideShows = (Val(GuideGet(pfx & "Opens")) <= 3)
End Function

Private Function GuideLine(Optional ByVal lite As Boolean = False) As String
    Dim s As String: s = GuideGet(IIf(lite, "GuideLiteDone", "GuideDone"))
    Dim names As Variant
    If lite Then
        names = Array("Book time off", "Approve a member's request", "Resolve a posting exception", "Fix an entry", "Run the daily check")
    Else
        names = Array("Practice on a copy", "Fill a shift now (a call-out)", "The Calendar", "How do I ...")
    End If
    Dim k As Long, t As String
    For k = 1 To UBound(names) + 1
        t = t & IIf(InStr(s, "|" & k & "|") > 0, "[x] ", "[  ] ") & k & " " & names(k - 1) & "     "
    Next k
    GuideLine = IIf(lite, "Your first week:     ", "Your first days:     ") & t
End Function

' the strip's "Hide this" (v29: the Lite strip when the Lite Desk is showing)
Public Sub DeskGuideHide()
    On Error GoTo Oops
    GuideSet IIf(IsLiteUser() And Not gDeskShowFull, "GuideLiteDone", "GuideDone"), "hidden"
    DrawDesk
    Exit Sub
Oops:
    ReportError "DeskGuideHide"
End Sub
