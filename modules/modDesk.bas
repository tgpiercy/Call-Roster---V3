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
'     roster - plus the help buttons. The admin-only
'     jobs are not on it (and the roster's gates refuse them anyway). Anyone on
'     neither list is a plain user and lands on the roster.
'  ViewModeFor is the one pure decision (SelfTest suite "viewmode"). The admin
'  Desk's view is ApplyViewCore's desk mode; the supervisor Desk is the
'  ordinary view with the Desk sheet in front.
' ====================================================================

Public Const DESK_VERSION As String = "Desk v25.3 (Oct 17)"
Public Const DESK_SHEET As String = "Desk"
Public Const DESK_AHEAD_DAYS As Long = 14           ' the "shifts need cover" tile looks this far ahead

Public gDeskMode As Boolean                         ' the Desk view is on (set by ApplyViewCore)
Private mZoomChecked As String                      ' sheets whose left-behind 120 % was put back this session (review)

Private Const DK_FONT As String = "Segoe UI"
Private Const DK_FONT_B As String = "Segoe UI Semibold"
Private Const BTN_H As Double = 36
Private Const BTN_W As Double = 230
Private Const TILE_H As Double = 58
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
    info = Format$(Date, "dddd mmm d, yyyy") & "    |    Pay period " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d") & "    |    " & kTxt
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
    DeskTile ws, "dk_tile_cover", nCover, IIf(nCover = 1, "shift needs cover", "shifts need cover"), _
             IIf(sup, "modDesk.DeskCalendar", "modDesk.DeskNeeds"), tileX, yy, IIf(nCover > 0, dkOrange(), 0)   ' Needs Attention is admin: a supervisor's tile opens the Calendar
    tileX = tileX + TILE_W + GAP
    DeskTile ws, "dk_tile_award", nReady, IIf(nReady = 1, "posting ready to award", "postings ready to award"), "modDesk.DeskAward", _
             tileX, yy, IIf(nReady > 0, dkOrange(), 0)
    tileX = tileX + TILE_W + GAP
    If Not sup Then
        DeskTile ws, "dk_tile_owed", nOwed, IIf(nOwed = 1, "payback owed", "paybacks owed"), "modDesk.DeskSwitch", _
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
    If sup Then pubCol = 0                          ' information only for a supervisor: never coloured (coloured = work to do)
    DeskTileText ws, "dk_tile_pub", "Members' file", pubTxt, IIf(sup, "", "modDesk.DeskPublish"), tileX, yy, pubCol

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
        yL = yL + BTN_H + 18
    End If
    DeskLabel ws, "dk_lbl_help", "HELP", x0, yL
    yL = yL + 16
    DeskBtn ws, "dk_btn_guide", "How do I ...", "modDesk.DeskGuide", x0, yL, dkGrey(), (BTN_W - GAP) / 2, dkInk()
    DeskBtn ws, "dk_btn_report", "Something went wrong", "modDesk.DeskReport", x0 + (BTN_W + GAP) / 2, yL, dkGrey(), (BTN_W - GAP) / 2, dkInk()

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
    If Not sup Then
        DeskBtn ws, "dk_btn_publish", "Update the members' file", "modDesk.DeskPublish", xR, yR, dkBlue()
        yR = yR + BTN_H + GAP
    End If
    DeskBtn ws, "dk_btn_print", "Print the call-out roster", "modDesk.DeskPrint", xR, yR, dkNavy()

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

' "How do I ...": the quick guide (GUIDE_PATH, a PDF on the plant PC; empty until go-live)
Public Sub DeskGuide()
    On Error GoTo Oops
    If Trim$(GUIDE_PATH) = "" Then
        MsgBox "The quick guide isn't on this PC yet. Ask " & PRIMARY_ADMIN_NAME & " for it, or for the job you're stuck on.", _
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
    ReportError "DeskGuide"
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
