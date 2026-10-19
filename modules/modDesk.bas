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
' ====================================================================

Public Const DESK_VERSION As String = "Desk v26 (Oct 18)"
Public Const DESK_SHEET As String = "Desk"
Public Const DESK_AHEAD_DAYS As Long = 14           ' the "shifts need cover" tile looks this far ahead
Public Const HELP_SHEET As String = "Help"           ' v25.5: the Help sheet (ShowHelp / DrawHelp below)
Private Const HELP_COL_W As Double = 118            ' the Help sheet's column B: one wide wrapped column (characters)

Public gDeskMode As Boolean                         ' the Desk view is on (set by ApplyViewCore)
Private mZoomChecked As String                      ' sheets whose left-behind 120 % was put back this session (review)

Private Const DK_FONT As String = "Segoe UI"
Private Const DK_FONT_B As String = "Segoe UI Semibold"
Private Const BTN_H As Double = 36
Private Const BTN_W As Double = 230
Private Const TILE_H As Double = 62                 ' v25.5: room for the cover tile's third line (the next uncovered shift)
Private Const TILE_W As Double = 205
Private Const GAP As Double = 8
Private Const COL2 As Double = 470                  ' the right column's left edge, points from the sheet's left margin

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
Private Function dkSwitch() As Long
    dkSwitch = RGB(180, 100, 0)
End Function
Private Function dkBlue() As Long
    dkBlue = RGB(21, 101, 192)
End Function
Private Function dkLeave() As Long
    dkLeave = RGB(106, 27, 154)
End Function
Private Function dkRed() As Long
    dkRed = RGB(176, 0, 32)
End Function

' ====================================================================
'  WHO
' ====================================================================
' The one pure decision: "desk" = on both ADMIN_USERS and DESK_USERS (admin rights, lands on the Desk), "admin" = on
' ADMIN_USERS only (today's layout), "supervisor" (v25.5) = on SUPERVISOR_USERS and not an admin (the supervisor's
' jobs, the supervisor Desk), "user" = none of them (a desk-only login has no admin rights; a plain user logs
' call-outs and looks at the calendar). Lists are comma separated, any case, spaces ignored. Mirrored in
' tests/rules.py (suite viewmode).
Public Function ViewModeFor(ByVal login As String, ByVal adminList As String, ByVal deskList As String, ByVal supList As String) As String
    Dim u As String: u = LCase$(Trim$(login))
    ViewModeFor = "user"
    If u = "" Then Exit Function
    Dim inA As Boolean, inD As Boolean, inS As Boolean
    inA = (InStr(1, "," & LCase$(Replace(adminList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    inD = (InStr(1, "," & LCase$(Replace(deskList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    inS = (InStr(1, "," & LCase$(Replace(supList, " ", "")) & ",", "," & u & ",", vbBinaryCompare) > 0)
    If inA And inD Then
        ViewModeFor = "desk"
    ElseIf inA Then
        ViewModeFor = "admin"
    ElseIf inS Then
        ViewModeFor = "supervisor"
    End If
End Function

Public Function IsDeskUser() As Boolean
    IsDeskUser = (ViewModeFor(CurrentUser(), ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS) = "desk")
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
    ws.Rows(2).RowHeight = 24

    ' ---- header: the title, then today | the pay period | the Kronos data's age ----
    Dim nCover As Long, nReady As Long, nOwed As Long, pubAt As Date
    DeskCounts nCover, nReady, nOwed, pubAt
    Dim ps As Date: ps = PeriodStartFor(Date)
    Dim dt As Date: dt = KronosDataThrough()
    Dim kTxt As String
    If dt = 0 Then
        kTxt = "Kronos data: none imported yet"
    Else
        kTxt = "Kronos data through " & Format$(dt, "ddd mmm d")
    End If
    Dim kOld As Boolean: kOld = (dt = 0 Or KronosIsStale())      ' the status line's one rule (review)
    With ws.Cells(1, 2)
        .Value = "CALL OUT ROSTER"
        .Font.Name = DK_FONT_B: .Font.Size = 18: .Font.Color = dkNavy()
        .VerticalAlignment = xlCenter
    End With
    Dim info As String
    ' v25.5 (A8): today's shifts - the crews on days and nights - in the header
    info = Format$(Date, "dddd mmm d, yyyy") & "    |    Today: " & DaysCrew(Date) & " on days, " & NightsCrew(Date) & " on nights" & _
           "    |    Pay period " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d") & "    |    " & kTxt
    With ws.Cells(2, 2)
        .NumberFormat = "@"
        .Value = info
        .Font.Size = 11: .Font.Color = dkInk()
        .VerticalAlignment = xlCenter
        If kOld Then .Characters(InStr(info, kTxt), Len(kTxt)).Font.Color = dkRed()
    End With

    ' ---- positions: everything hangs off the top-left of row 3 ----
    Dim x0 As Double: x0 = ws.Cells(1, 2).Left
    Dim y0 As Double: y0 = ws.Cells(3, 1).Top + 6
    Dim yy As Double

    ' ---- the layout: the admin Desk (every job) or the supervisor Desk (a supervisor's jobs) ----
    Dim sup As Boolean: sup = Not gAdminUnlocked
    Dim tileX As Double: tileX = x0

    ' ---- TODAY: the tiles ----
    DeskLabel ws, "dk_lbl_today", "TODAY", x0, y0
    yy = y0 + 16
    Dim nextTxt As String                          ' v25.5 (A8): the next uncovered shift under the count
    If nCover > 0 Then nextTxt = NextOpenLine(DESK_AHEAD_DAYS)
    DeskTile ws, "dk_tile_cover", nCover, IIf(nCover = 1, "shift needs cover", "shifts need cover") & IIf(nextTxt <> "", vbCr & nextTxt, ""), _
             IIf(sup, "modDesk.DeskCalendar", "modDesk.DeskNeeds"), tileX, yy, IIf(nCover > 0, dkOrange(), 0)   ' Needs Attention is admin: a supervisor's tile opens the Calendar
    tileX = tileX + TILE_W + GAP
    DeskTile ws, "dk_tile_award", nReady, IIf(nReady = 1, "posting ready to award", "postings ready to award"), "modDesk.DeskAward", _
             tileX, yy, IIf(nReady > 0, dkOrange(), 0)
    tileX = tileX + TILE_W + GAP
    If Not sup Then
        Dim nOver As Long: nOver = OwedOverdueCount()      ' v26 (payback aging): owed longer than PAYBACK_WARN_DAYS
        DeskTile ws, "dk_tile_owed", nOwed, IIf(nOwed = 1, "payback owed", "paybacks owed") & _
                 IIf(nOver > 0, " (" & nOver & " overdue)", ""), "modDesk.DeskSwitch", _
                 tileX, yy, IIf(nOwed > 0, dkSwitch(), 0)
        tileX = tileX + TILE_W + GAP
    End If
    Dim pubTxt As String, pubCol As Long
    If pubAt = 0 Then
        pubTxt = "never published": pubCol = dkBlue()
    ElseIf Int(Date) - Int(pubAt) >= MEMBERS_STALE_DAYS Then
        pubTxt = "updated " & Format$(pubAt, "ddd mmm d"): pubCol = dkBlue()
    Else
        pubTxt = "updated today " & Format$(pubAt, "h:mm AM/PM"): pubCol = 0
    End If
    ' v25.5 (bundle item 8): the tile is live for a supervisor too - they publish (coloured = work to do)
    DeskTileText ws, "dk_tile_pub", "Members' file", pubTxt, "modDesk.DeskPublish", tileX, yy, pubCol

    ' ---- RIGHT NOW (left column) ----
    yy = yy + TILE_H + 18
    Dim yL As Double: yL = yy
    DeskLabel ws, "dk_lbl_now", "RIGHT NOW", x0, yL
    yL = yL + 16
    DeskBtn ws, "dk_btn_sick", "Someone called in sick", "modDesk.DeskSick", x0, yL, dkOrange()
    yL = yL + BTN_H + GAP
    DeskBtn ws, "dk_btn_callout", "Fill a shift now (call-out)", "modDesk.DeskCallout", x0, yL, dkGreen()
    yL = yL + BTN_H + GAP
    DeskBtn ws, "dk_btn_cal", "What needs cover (calendar)", "modDesk.DeskCalendar", x0, yL, dkNavy()
    yL = yL + BTN_H + 18
    If Not sup Then
        DeskLabel ws, "dk_lbl_people", "PEOPLE", x0, yL
        yL = yL + 16
        DeskBtn ws, "dk_btn_leave", "Leave: going or returning", "modDesk.DeskLeave", x0, yL, dkLeave()
        yL = yL + BTN_H + GAP
        DeskBtn ws, "dk_btn_undo", "Undo my last log entry", "modDesk.DeskUndo", x0, yL, dkGrey(), BTN_W, dkInk()   ' v25.5 (A9): admin only (void); "log entry": it undoes OT Log rows only (review)
        yL = yL + BTN_H + 18
    End If
    DeskLabel ws, "dk_lbl_help", "HELP", x0, yL
    yL = yL + 16
    DeskBtn ws, "dk_btn_guide", "How do I ...", "modDesk.DeskGuide", x0, yL, dkGrey(), (BTN_W - GAP) / 2, dkInk()
    DeskBtn ws, "dk_btn_report", "Something went wrong", "modDesk.DeskReport", x0 + (BTN_W + GAP) / 2, yL, dkGrey(), (BTN_W - GAP) / 2, dkInk()
    ' v25.5 (A8): a suggestion goes the same way as a fault, with its own subject
    yL = yL + BTN_H + GAP
    DeskBtn ws, "dk_btn_suggest", "Suggest an improvement", "modDesk.DeskSuggest", x0, yL, dkGrey(), BTN_W, dkInk()
    ' v25.5 (A5): the shortcut keys, under the help buttons
    yL = yL + BTN_H + 6
    DeskLabel ws, "dk_lbl_keys", "Ctrl+Shift+S sick call   Ctrl+Shift+C call-out   Ctrl+Shift+K calendar   Ctrl+Shift+D this Desk", x0, yL
    ws.Shapes("dk_lbl_keys").Width = 440
    yL = yL + 14 - BTN_H                            ' the footer adds BTN_H below the taller column

    ' ---- PLAN AHEAD / WHEN YOU ARE DONE (right column) ----
    Dim yR As Double: yR = yy
    Dim xR As Double: xR = x0 + COL2
    DeskLabel ws, "dk_lbl_plan", "PLAN AHEAD", xR, yR
    yR = yR + 16
    DeskBtn ws, "dk_btn_post", "Post a shift", "modDesk.DeskPost", xR, yR, dkGreen()
    yR = yR + BTN_H + GAP
    DeskBtn ws, "dk_btn_award", "Award a posting", "modDesk.DeskAward", xR, yR, dkGreen()
    yR = yR + BTN_H + GAP
    If Not sup Then
        DeskBtn ws, "dk_btn_timeoff", "Book time off", "modDesk.DeskTimeOff", xR, yR, dkNavy()
        yR = yR + BTN_H + GAP
        DeskBtn ws, "dk_btn_switch", "Shift switch or payback", "modDesk.DeskSwitch", xR, yR, dkSwitch()
        yR = yR + BTN_H + GAP
    End If
    yR = yR + 10
    DeskLabel ws, "dk_lbl_done", "WHEN YOU ARE DONE", xR, yR
    yR = yR + 16
    DeskBtn ws, "dk_btn_publish", "Update the members' file", "modDesk.DeskPublish", xR, yR, dkBlue()   ' v25.5: on the supervisor Desk too (bundle item 8)
    yR = yR + BTN_H + GAP
    DeskBtn ws, "dk_btn_print", "Print the call-out roster", "modDesk.DeskPrint", xR, yR, dkNavy()
    yR = yR + BTN_H + GAP
    DeskBtn ws, "dk_btn_history", "Employee history (report)", "modDesk.DeskHistory", xR, yR, dkNavy()   ' v26: one person, a date range, a PDF

    ' ---- footer links ----
    Dim yF As Double: yF = IIf(yL > yR, yL, yR) + BTN_H + 22
    DeskLink ws, "dk_lnk_roster", "Full roster  >", "modDesk.DeskRoster", x0, yF, 140
    If Not sup Then DeskLink ws, "dk_lnk_all", "Show everything (primary admin tools)  >", "modDesk.DeskShowAll", xR, yF, 260
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
    On Error Resume Next
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
    On Error Resume Next
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

' a 12 pt verb-phrase button, 36 points tall
Private Sub DeskBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                    ByVal l As Double, ByVal t As Double, ByVal fill As Long, _
                    Optional ByVal w As Double = BTN_W, Optional ByVal fontCol As Long = vbWhite)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, BTN_H, fill)
    On Error Resume Next
    With s.TextFrame2.TextRange
        .Text = cap
        .ParagraphFormat.Alignment = msoAlignLeft
        .Font.Name = DK_FONT_B
        .Font.Size = 12
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = fontCol
    End With
End Sub

' a TODAY tile: the figure big, the caption under it; grey when there is nothing to do
Private Sub DeskTile(ByVal ws As Worksheet, ByVal nm As String, ByVal n As Long, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal col As Long)
    Dim active As Boolean: active = (col <> 0)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, TILE_W, TILE_H, IIf(active, col, dkGrey()))
    Dim numTxt As String: numTxt = CStr(n)
    On Error Resume Next
    With s.TextFrame2.TextRange
        .Text = numTxt & vbCr & cap
        .ParagraphFormat.Alignment = msoAlignLeft
        .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0
        .Font.Name = DK_FONT_B
        .Font.Size = 11
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = IIf(active, vbWhite, dkMuted())
        .Characters(1, Len(numTxt)).Font.Size = 20
    End With
End Sub

' the members'-file tile: a caption and a text instead of a figure (mac "" = information only)
Private Sub DeskTileText(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal txt As String, ByVal mac As String, _
                         ByVal l As Double, ByVal t As Double, ByVal col As Long)
    Dim active As Boolean: active = (col <> 0)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, TILE_W, TILE_H, IIf(active, col, dkGrey()))
    On Error Resume Next
    With s.TextFrame2.TextRange
        .Text = cap & vbCr & txt
        .ParagraphFormat.Alignment = msoAlignLeft
        .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0
        .Font.Name = DK_FONT_B
        .Font.Size = 11
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.RGB = IIf(active, vbWhite, dkMuted())
        .Characters(1, Len(cap)).Font.Size = 13
    End With
End Sub

' a small flat link
Private Sub DeskLink(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal w As Double)
    Dim s As Shape: Set s = DeskShape(ws, nm, mac, l, t, w, 22, vbWhite)
    On Error Resume Next
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
    On Error Resume Next
    ws.Range("A1").Select
End Sub

' ThisWorkbook's open step 10: an admin Desk user and (SUPERVISOR_DESK) every supervisor on SUPERVISOR_USERS land on
' the Desk (v25.5: a plain user lands on the roster), the primary admin on the roster; a Desk that can't be made or
' brought to the front leaves the user on the roster (DrawDesk never raises: it writes its own red line on the
' sheet), and a 120 % left behind is put back
Public Sub DeskOpenLanding()
    On Error Resume Next
    If gDeskMode Or (SUPERVISOR_DESK And Not gAdminUnlocked And IsSupervisor()) Then
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
    On Error Resume Next
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

' Called from ThisWorkbook's SheetActivate: the Desk is refreshed every time anyone looks at it (its figures
' must never be stale); an admin Desk user sees the roster and the Calendar at 120 %; anyone else landing on
' those sheets at a left-behind 120 % gets 100 % back ONCE per sheet per session (re-review: a user who picks
' 120 % himself afterwards is left alone).
Public Sub DeskSheetActivated(ByVal Sh As Object)
    On Error Resume Next
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
    On Error Resume Next
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
    ShowOvertimeForm
    DeskRefreshIfShown
    Exit Sub
Oops:
    ReportError "DeskCallout"
End Sub

Public Sub DeskCalendar()
    On Error GoTo Oops
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
    On Error Resume Next
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
        "A supervisor on the list, and an admin on the Desk list, open on the Desk tab (the primary admin opens on the roster and reaches it with < Desk): the tiles say what is waiting today (shifts needing cover, postings ready " & _
        "to award, the members' file's age) and the buttons are the jobs - Someone called in sick, Fill a shift now, What needs " & _
        "cover, Post a shift, Award a posting, Update the members' file, Print the call-out roster. Full roster > at the bottom " & _
        "goes to the roster; the Desk tab brings it back. Anyone not on the supervisor list opens on the roster and can log " & _
        "call-outs and look at the calendar."
    HelpAdd h, n, "Log a call-out (fill a shift now)", _
        "Click Fill a shift now (or Log Call-Out on the roster's bar). Pick the Department, Shift and Date, type the Hours and " & _
        "click LOAD LIST: the list is in call order and holds only the people free for that shift. The card shows who to call " & _
        "and their phone; press A (accepted), R (refused), L (left message) or U (unavailable: pick the reason, Enter records " & _
        "it) - or click the button - and the next person comes up. K skips for now, B goes back one, P is part of the hours, " & _
        "Call back goes down the left-message pile. When the shift is filled press Enter or click SUBMIT ALL ENTRIES and note " & _
        "the session number (S-0012). Refused and Left Message are charged for grade 2 " & _
        "and below; >> means above grade 2 (called, never charged); * means the total includes accepted overtime not yet in " & _
        "Kronos; ~ means an estimate because a unit move is planned for the person."
    HelpAdd h, n, "Someone takes only part of the hours", _
        "Select them and press P (Part of the hours), then enter the hours. That closes the shift: nobody else is called, " & _
        "nobody below is charged, and the pop-up refuses a further Accepted (a refusal or a message already logged can still be " & _
        "corrected to a refusal, a message or unavailable). Anyone who refused before the partial keeps their " & _
        "charge. Click SUBMIT ALL ENTRIES."
    HelpAdd h, n, "Someone called in sick", _
        "Click Someone called in sick (Sick Call on the bar). Pick who called, check the Shift date, Shift, Unit to cover and " & _
        "Hours, and read the line under the Note box: it gives their sick-day counts and whether this shift is a paid sick day. " & _
        "SAVE + START CALL-OUT books the sick shift and opens the call-out filled in; Save, no call-out books it without a call-out. " & _
        "The person is off every call list for that shift and any shift starting within 24 hours after it ends. Log each sick " & _
        "shift as its own sick call. Members see it as Off, never why."
    HelpAdd h, n, "Post a shift", _
        "Click Post a shift (New Posting on the bar). Fill in Unit, Shift date, Shift, Reason, Hours and Positions; Crew and " & _
        "Posting closes fill in. Leave Print it and Save PDF ticked, click CREATE POSTING, then put up the printed sheet. Sign-up " & _
        "is on paper. The Reason must not say what kind of leave anyone is on: a leave word is refused. Load Posting (admin) " & _
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
        "post it). Print / PDF prints the month or saves it as a PDF."
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
    HelpAdd h, n, "Find a person", _
        "Supervisors and admins: click Find a person on the roster's bar and type part of the name: the roster jumps to the row (a unit filter that " & _
        "hides it is cleared) and the status bar at the bottom says where the person stands in their home unit's call order " & _
        "today, for example 4 of 19 in the Phosphine call order. Click any roster row for the person's refused hours by kind."
    HelpAdd h, n, "Minimum staffing (admin)", _
        "The Min Staff sheet (admin) holds the minimum operators per unit and shift; blank = no minimum on that shift. For today and " & _
        "the next 13 days the calendar's date cell shows a red N short flag when any unit's day or night shift is below it, whatever unit group the " & _
        "calendar shows (click the date for which), and Needs Attention lists each short shift under BELOW MINIMUM STAFFING, every unit " & _
        "(double-click to open the calendar on that date). Staffed " & _
        "= the people on the crew that is on, holding a code in the unit, minus those booked off or sick for that shift or off in a " & _
        "switch, plus every coverage line a named person covers (a switch's own S/S line counts the one working it); an OPEN, " & _
        "NO TAKERS, POSTED or NCN line adds nobody. It only reads: fix a short shift the usual way (a call-out, a posting, a switch)."
    HelpAdd h, n, "The two-week board", _
        "Board on the Calendar's top row (supervisors and admins): the Board tab is drawn for today and the next 13 days, one page " & _
        "per unit group, one row per date with the crews on and the Days and Nights lines as the calendar shows them, OPEN lines " & _
        "in red; a PDF goes beside the month calendars, then it asks whether to print. Nothing is changed: it is a picture of the " & _
        "calendar for the board in the control room."
    HelpAdd h, n, "Employee history (report)", _
        "Employee history on the bar's FIND group or the Desk: part of a name, then a From and To date (the boxes open on the whole roster " & _
        "year; the member-safe report takes the pay periods starting in the range, the full one the shift dates), and the History Report tab lists every OT Log entry for that person in the range, newest first: call-outs with " & _
        "their place in the call order, postings, adjustments, the hours charged and a Why line on every charged row (for example Charged 21 h: 3rd " & _
        "called, Refused, 12 h offered x 1.75), then the charged hours per pay period and the total. An admin is asked for the full " & _
        "report (on the History Report (admin) tab: who logged each entry and when, a Why on every row, void rows struck through) or the " & _
        "member-safe one (the same rows and columns the members' My history shows); a supervisor gets the member-safe one. Save a PDF as well? puts it in the " & _
        "reports folder (beside the roster until go-live). Click any OT Log row for its Why on the status bar."
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
        "Total Hours: worked (Kronos) + the hours charged in the OT Log this year (refusals, left messages, adjustments) + accepted " & _
        "overtime not yet in Kronos (*). Charged: refused or left-message " & _
        "hours x 1.75 added to Total Hours, grade 2 and below only (>> = above grade 2, called but never charged). ~ = an " & _
        "estimate because a unit move is planned for the person. ? = left a message. S/S = working a shift switch that " & _
        "shift. NCN = no cover needed. PHA = hazard-analysis coverage. OPEN = the line still needs cover; POSTED = a " & _
        "posting is out for it. DAYS = the Monday-to-Friday day crew. A sick call or a personal day is Off to the " & _
        "members, never why."
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
