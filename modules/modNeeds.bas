Attribute VB_Name = "modNeeds"
Option Explicit
' v27.1 (the refactor, Oct 2026): the Needs Attention tab (DrawNeeds and its steps), from modCalendar; moved unchanged (tests/movecheck.py proves it)
Public Const NEEDS_VERSION As String = "Needs Attention v28 (Oct 25)"

' one row on the Needs Attention tab
Private Type NARow
    Cell(0 To 7) As String
    Key As String
    SortKey As Double
    Accent As Long
    StatusCol As Long
    StatusFill As Long
    WhenCol As Long
End Type
' ====================================================================
'  NEEDS ATTENTION TAB  (admin only - double-click any row for options)
' ====================================================================
Private Function NeedsSheet() As Worksheet
    If Not SheetExists(NA_SHEET) Then
        Dim ws As Worksheet
        Dim cs As Worksheet: Set cs = CalendarSheet()  ' v25.1 review: before the unlock (its SetTabColor relocks the tabs)
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=cs)
        ws.Name = NA_SHEET
        RelockStructure
    End If
    Set NeedsSheet = ThisWorkbook.Worksheets(NA_SHEET)
    SetTabColor NeedsSheet, thOpen()                ' v25.1
End Function

Public Sub ShowNeedsAttention()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    CalendarSheet
    Dim ws As Worksheet: Set ws = NeedsSheet()
    SetSheetVisible ws, xlSheetVisible
    ws.Activate                          ' the tab-switch check runs here (ThisWorkbook)
    If NAStale() Then DrawNeeds
    LiteMark "needs"                     ' v29: the Lite Desk's daily check - Needs Attention was looked at
    Exit Sub
Oops:
    ReportError "ShowNeedsAttention"
End Sub

Private Function NAGroup() As String
    NAGroup = SafeStr(MapSheet().Range("Z5").Value)
    If NAGroup = "" Then NAGroup = "All units"
End Function

Private Function NADays() As Long
    NADays = CLng(Val(MapSheet().Range("Z6").Value))
    If NADays <= 0 Then NADays = 14
End Function

Private Function NARanges() As Variant
    NARanges = Array(7, 14, 28, 56)
End Function

Public Sub NAUnitChanged()
    On Error GoTo Oops
    Dim dd As Object
    On Error Resume Next
    Set dd = NeedsSheet().DropDowns("na_unit")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z5", dd.List(dd.ListIndex)
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NAUnitChanged"
End Sub

Public Sub NARangeChanged()
    On Error GoTo Oops
    Dim dd As Object
    On Error Resume Next
    Set dd = NeedsSheet().DropDowns("na_range")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z6", NARanges()(dd.ListIndex - 1)
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NARangeChanged"
End Sub

Public Sub NARefresh()
    On Error GoTo Oops
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NARefresh"
End Sub

Public Sub NAGoCalendar()
    On Error GoTo Oops
    ShowCalendar
    Exit Sub
Oops:
    ReportError "NAGoCalendar"
End Sub

' "Today", "Tomorrow", "in 5 days", "3 days ago"
Private Function WhenText(ByVal d As Date) As String
    Dim n As Long: n = Int(d) - Int(Date)
    Select Case n
        Case 0: WhenText = "Today"
        Case 1: WhenText = "Tomorrow"
        Case -1: WhenText = "Yesterday"
        Case Is > 1: WhenText = "in " & n & " days"
        Case Else: WhenText = -n & " days ago"
    End Select
End Function

Private Function WhenColor(ByVal d As Date) As Long
    Select Case Int(d) - Int(Date)
        Case Is <= 0: WhenColor = thOpen()
        Case 1: WhenColor = thPosted()
        Case Else: WhenColor = thMuted()
    End Select
End Function

' tblPostingList snapshot (for "n signed up")
Private Function PostListData() As Variant
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(PL_TABLE)
        If Not lo Is Nothing Then
            PostListData = TableData(lo)
            Exit Function
        End If
    Next ws
End Function

Private Function SignedCount(ByVal plV As Variant, ByVal id As String) As Long
    If Not IsArray(plV) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(plV, 1)
        If StrComp(SafeStr(plV(i, PL_ID)), id, vbTextCompare) = 0 Then
            If UCase$(SafeStr(plV(i, PL_SIGNED))) = "YES" Then SignedCount = SignedCount + 1
        End If
    Next i
End Function


Private Sub NAAdd(ByRef recs() As NARow, ByRef n As Long, ByVal c0 As String, ByVal c1 As String, ByVal c2 As String, _
                  ByVal c3 As String, ByVal c4 As String, ByVal c5 As String, ByVal c6 As String, ByVal c7 As String, _
                  ByVal keyTxt As String, ByVal sortKey As Double, ByVal accent As Long, ByVal stCol As Long, _
                  ByVal stFill As Long, ByVal whenCol As Long)
    n = n + 1
    If n > UBound(recs) Then ReDim Preserve recs(1 To n + 50)
    With recs(n)
        .Cell(0) = c0: .Cell(1) = c1: .Cell(2) = c2: .Cell(3) = c3
        .Cell(4) = c4: .Cell(5) = c5: .Cell(6) = c6: .Cell(7) = c7
        .Key = keyTxt
        .SortKey = sortKey
        .Accent = accent
        .StatusCol = stCol
        .StatusFill = stFill
        .WhenCol = whenCol
    End With
End Sub

Private Sub NASort(ByRef recs() As NARow, ByVal n As Long)
    Dim i As Long, j As Long, t As NARow
    For i = 2 To n
        t = recs(i)
        j = i - 1
        Do While j >= 1
            If recs(j).SortKey <= t.SortKey Then Exit Do
            recs(j + 1) = recs(j)
            j = j - 1
        Loop
        recs(j + 1) = t
    Next i
End Sub

Public Sub DrawNeeds()
    Dim ws As Worksheet: Set ws = NeedsSheet()
    Dim grp As String: grp = NAGroup()
    Dim nDays As Long: nDays = NADays()
    Dim drawnFrom As String: drawnFrom = SafeStr(MapSheet().Range("Z7").Value)
    Dim t0 As Double, tA As Double, tB As Double, tC As Double    ' v25.1.2: draw-time measurement (CalMap Z15)
    t0 = Timer
    mTimeSlots = 0
    mDrawing = True
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim calcWas As Long: calcWas = Application.Calculation
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    On Error GoTo Failed
    UnlockSheet ws
    LoadCache
    LoadPostCache
    Dim plV As Variant: plV = PostListData()

    ' ---- gather ----
    Dim pr() As NARow, nP As Long, lr() As NARow, nL As Long, sr() As NARow, nS As Long, br() As NARow, nB As Long
    ReDim pr(1 To 50): ReDim lr(1 To 50): ReDim sr(1 To 50): ReDim br(1 To 50)
    Dim shortRows() As NARow, nShort As Long          ' v26: below minimum staffing
    ReDim shortRows(1 To 50)
    Dim nReady As Long                                ' v27.1: the lists the steps fill (and two counts) live here
    Dim hr() As NARow, nH As Long, hLook As Long
    Dim ir() As NARow, nI As Long, mr() As NARow, nM As Long, er() As NARow, nE As Long
    Dim ur() As NARow, nU As Long, dcR() As NARow, nDc As Long, vr() As NARow, nV As Long
    Dim tr() As NARow, nT As Long
    Dim rq() As NARow, nRq As Long                    ' v28 (item 1): members' requests waiting for an admin
    NAGatherWork grp, nDays, plV, pr, nP, nReady, lr, nL, shortRows, nShort, sr, nS, br, nB, hr, nH, hLook
    NAGatherAdmin grp, ir, nI, mr, nM, er, nE, ur, nU, dcR, nDc, vr, nV, tr, nT, rq, nRq

    tA = Timer                                      ' gathered
    NADrawHead ws, grp, nDays                       ' v27.1: the sheet wiped, the columns, rows 1-6

    Dim r As Long: r = 7
    Dim rowP As Long, rowL As Long, rowS As Long, rowB As Long
    rowL = r
    r = NASection(ws, r, "OPEN LINES  -  need cover in the next " & nDays & " days", nL, thOpen(), _
                  Array("Date", "Shift", "Crew", "Unit", "Who's off / need", "Type", "Status", "When"))
    r = NARows(ws, r, lr, nL, "Nothing open - every line in the window is covered, posted or NCN.")
    If nShort > 0 Or MinStaffCount() > 0 Then       ' v26: only once the Min Staff sheet holds a minimum
        r = NASection(ws, r, "BELOW MINIMUM STAFFING  -  every unit, next " & IIf(nDays < BOARD_DAYS, nDays, BOARD_DAYS) & " days (double-click for the calendar on that date)", nShort, thOpen(), _
                      Array("Shift date", "Shift", "Crew", "Unit", "Staffed", "Type", "Status", "When"))
        r = NARows(ws, r, shortRows, nShort, "Every unit with a minimum is at or above it.")
    End If
    rowP = r
    r = NASection(ws, r, "POSTINGS  -  open postings, oldest closing first", nP, thPosted(), _
                  Array("Shift date", "Shift", "Crew", "Unit", "Posting", "Type", "Status", "Closes"))
    r = NARows(ws, r, pr, nP, "No open postings.")
    rowS = r
    r = NASection(ws, r, "SHIFT SWITCH PAYBACKS OWED  -  oldest first; over " & PaybackWarnDays() & " days in amber", nS, TypeColor("SS"), _
                  Array("Switch date", "Shift", "Crew", "Unit", "Owed", "Type", "Status", "When"))
    r = NARows(ws, r, sr, nS, "No paybacks owed.")
    rowB = r
    r = NASection(ws, r, "LEAVE  -  returning soon or open-ended", nB, TypeColor("LOA"), _
                  Array("Started", "", "Crew", "Unit", "Who", "Type", "Status", "When"))
    r = NARows(ws, r, br, nB, "Nobody returning from leave in the window, and nothing open-ended.")
    If nH > 0 Then
        r = NASection(ws, r, "HOLIDAYS TO CONFIRM  -  weekend holiday, weekday off not entered (next " & hLook & " days)", _
                      nH, RGB(176, 0, 32), Array("Holiday date", "", "Crew", "Unit", "Holiday", "Type", "Status", "When"))
        r = NARows(ws, r, hr, nH, "")
    End If
    If nU > 0 Then
        r = NASection(ws, r, "HOME UNIT NOT SET  -  double-click a name: the Employee pop-up opens on Edit, pick the home unit there", _
                      nU, thMuted(), Array("", "", "Crew", "", "Who", "Type", "Status", ""))
        r = NARows(ws, r, ur, nU, "")
    End If
    If nDc > 0 Then                                 ' v27 (item 25)
        r = NASection(ws, r, "DATA CHECK  -  roster details that look wrong (double-click: the roster row; Employee > Edit fixes it)", _
                      nDc, thMuted(), Array("", "", "Crew", "", "Who and what", "Type", "Status", ""))
        r = NARows(ws, r, dcR, nDc, "")
    End If
    If nRq > 0 Then                                   ' v28 (item 1): members' requests waiting for an admin
        r = NASection(ws, r, "MEMBERS' REQUESTS  -  waiting for an admin (double-click: book it through the pop-up; right-click: decline with a reason)", _
                      nRq, thPosted(), Array("From", "", "Crew", "Unit", "Who and what", "Kind", "Status", "Received"))
        r = NARows(ws, r, rq, nRq, "")
    End If
    If nE > 0 Then
        r = NASection(ws, r, "ERRORS  -  something didn't finish (double-click to open the Error Log)", _
                      nE, RGB(176, 0, 32), Array("Today", "", "", "", "What", "Type", "Status", ""))
        r = NARows(ws, r, er, nE, "")
    End If
    If nM > 0 Then
        r = NASection(ws, r, "MEMBERS FILE  -  what the union sees is out of date (double-click to publish)", _
                      nM, thPosted(), Array("Today", "", "", "", "Members file", "Type", "Status", ""))
        r = NARows(ws, r, mr, nM, "")
    End If
    If nI > 0 Then
        r = NASection(ws, r, "IMPORTED, NOT BOOKED  -  old-calendar vacations that don't keep people off call lists yet (double-click to convert)", _
                      nI, TypeColor("VAC"), Array("Date", "Shift", "", "Unit", "As imported", "Type", "Status", "When"))
        r = NARows(ws, r, ir, nI, "")
    End If
    If nV > 0 Then
        r = NASection(ws, r, "PLANNED UNIT MOVES  -  applied when an admin opens the file on or after the date (double-click to cancel)", _
                      nV, thMuted(), Array("Effective", "", "New crew", "New unit", "Who", "Type", "Status", "When"))
        r = NARows(ws, r, vr, nV, "")
    End If
    If nT > 0 Then
        r = NASection(ws, r, "ATTENDANCE  -  sick / personal days at or over the limit (double-click for the Attendance sheet)", _
                      nT, TypeColor("SIC"), Array("Last sick", "", "Crew", "Home unit", "Who", "Type", "Status", "Until"))
        r = NARows(ws, r, tr, nT, "")
    End If

    tB = Timer                                      ' rows drawn
    ' ---- summary tiles (double-click one to jump to its list) ----
    NATile ws, 2, 3, nL, "OPEN LINES", thOpen(), rowL
    NATile ws, 4, 5, nReady, "READY TO AWARD", thPosted(), rowP
    NATile ws, 6, 6, nP, "OPEN POSTINGS", TypeColor("VAC"), rowP
    NATile ws, 7, 7, nS, "PAYBACKS OWED", TypeColor("SS"), rowS
    NATile ws, 8, 9, nB, "LEAVE TO WATCH", TypeColor("LOA"), rowB

    BuildNAControls ws, grp, nDays

    On Error Resume Next
    If ActiveSheet Is ws Then
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .FreezePanes = False
            .ScrollRow = 1
            .SplitColumn = 0
            .SplitRow = 2
            .FreezePanes = True
        End With
    End If
    On Error GoTo Failed
    MapWrite "Z9", drawnFrom                        ' v25.1: through the pair
    MapWrite "Z11", "'" & DrawContext()
    tC = Timer                                      ' v25.1.2: tiles drawn
    MapWrite "Z15", "'" & "Needs Attention " & grp & ", " & nDays & " days (" & (nP + nL + nS + nB + nH + nU + nE + nM + nI + nV + nT) & " rows): drawn in " & _
        Format$(Secs(t0, tC), "0.0") & " s = gather " & Format$(Secs(t0, tA), "0.0") & " (slots " & Format$(mTimeSlots, "0.0") & _
        "), rows " & Format$(Secs(tA, tB), "0.0") & ", tiles " & Format$(Secs(tB, tC), "0.0") & "  [" & CAL_VERSION & "]"
    GoTo Done

Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description & " (error " & Err.Number & ")"
    Resume AfterFail
AfterFail:
    On Error Resume Next
    LogError "DrawNeeds", en, em          ' v24.3.1
    UnlockSheet ws
    ws.Cells(3, 2).Value = "Couldn't finish: " & em
    ws.Cells(3, 2).Font.Color = thOpen()
    BuildNAControls ws, grp, nDays
    Application.ScreenUpdating = True
    MsgBox "The Needs Attention tab couldn't finish drawing:" & vbNewLine & vbNewLine & em & vbNewLine & vbNewLine & _
           "Send a screenshot of this message.  (" & CAL_VERSION & ")", vbExclamation, "Needs Attention"
Done:
    On Error Resume Next
    ProtectSheet ws
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
    mDrawing = False
End Sub

' v27.1: DrawNeeds' steps. No handler of their own: an error rises to DrawNeeds' Failed, which restores
' everything. Called only while that handler is active. The lists and counts are DrawNeeds' own (ByRef).
Private Sub NAGatherWork(ByVal grp As String, ByVal nDays As Long, ByRef plV As Variant, _
                         ByRef pr() As NARow, ByRef nP As Long, ByRef nReady As Long, _
                         ByRef lr() As NARow, ByRef nL As Long, ByRef shortRows() As NARow, ByRef nShort As Long, _
                         ByRef sr() As NARow, ByRef nS As Long, ByRef br() As NARow, ByRef nB As Long, _
                         ByRef hr() As NARow, ByRef nH As Long, ByRef hLook As Long)
    Dim i As Long, j As Long, m As Long, cl As Date, sd As Date, d As Date, sh As Variant, rt As String
    Dim tier As Long, st As String, stC As Long, stF As Long, id As String, txt As String

    ' A. postings still open
    If IsArray(mPostV) Then
        For i = 1 To UBound(mPostV, 1)
            If SafeStr(mPostV(i, PS_STATUS)) = "Open" And UnitInGroup(SafeStr(mPostV(i, PS_UNIT)), grp) Then
                id = SafeStr(mPostV(i, PS_ID))
                cl = AsDate(mPostV(i, PS_CLOSES))
                sd = AsDate(mPostV(i, PS_DATE))
                If cl > 0 And PostingClosed(cl) Then          ' v25 review: the one closing test
                    tier = 0: st = "READY TO AWARD": stC = thOpen(): stF = thOpenFill(): nReady = nReady + 1
                ElseIf cl > 0 And Int(cl) = Date Then
                    tier = 1: st = "Closes today " & POSTING_CLOSE_TIME: stC = thPosted(): stF = thPostedFill()
                ElseIf cl > 0 Then
                    tier = 2: st = "Open until " & Format$(cl, "ddd mmm d"): stC = thMuted(): stF = -1
                Else
                    tier = 3: st = "Open (no closing date)": stC = thMuted(): stF = -1
                End If
                txt = id & "  " & SafeStr(mPostV(i, PS_REASON)) & "   (" & SafeStr(mPostV(i, PS_HOURS)) & "h, " & _
                      SafeStr(mPostV(i, PS_POS)) & " pos, " & SignedCount(plV, id) & " signed up)"
                If Not CovHasSourceCached(id) Then
                    txt = txt & "   - NOT ON CALENDAR"
                    If tier > 0 Then stC = thOpen()
                End If
                NAAdd pr, nP, Format$(sd, "ddd mmm d"), SafeStr(mPostV(i, PS_SHIFT)), SafeStr(mPostV(i, PS_CREW)), _
                      SafeStr(mPostV(i, PS_UNIT)), txt, TypeLong(ReasonType(SafeStr(mPostV(i, PS_REASON)))), st, _
                      IIf(cl > 0, WhenText(cl), ""), "P:" & id, tier * 1000000# + CDbl(Int(cl)), thPosted(), stC, stF, _
                      IIf(cl > 0, WhenColor(cl), thMuted())
            End If
        Next i
        NASort pr, nP
    End If

    ' B. open lines in the window
    Dim slots() As CalSlot, unit As String
    For i = 0 To nDays - 1
        d = Date + i
        For Each sh In Array("Days", "Nights")
            m = SlotsFor(d, CStr(sh), grp, slots, True)
            For j = 1 To m
                rt = UCase$(slots(j).RightText)
                If slots(j).IsOpen Or rt Like "NO TAKERS*" Then
                    unit = SlotUnit(slots(j).Key)
                    NAAdd lr, nL, Format$(d, "ddd mmm d"), CStr(sh), IIf(sh = "Days", DaysCrew(d), NightsCrew(d)), unit, _
                          slots(j).LeftText, TypeLong(slots(j).TypeCode), IIf(rt Like "NO TAKERS*", "NO TAKERS", "OPEN"), _
                          WhenText(d), slots(j).Key & "|" & CLng(Int(d)) & "|" & sh, 0, thOpen(), thOpen(), thOpenFill(), WhenColor(d)
                End If
            Next j
        Next sh
    Next i

    ' B2. below minimum staffing (v26): today .. today + nDays - 1, the group's units that have a minimum
    StaffSweepBegin Date
    Dim ms As Variant: ms = StaffShortList(Date, IIf(nDays < BOARD_DAYS, nDays, BOARD_DAYS), "")    ' the brief's 14 days, whatever the box's window; v26.5: every unit, whatever group is picked
    If IsArray(ms) Then
        For i = 1 To UBound(ms, 1)
            d = CDate(CDbl(ms(i, 2)))
            NAAdd shortRows, nShort, Format$(d, "ddd mmm d"), CStr(ms(i, 3)), IIf(CStr(ms(i, 3)) = "Days", DaysCrew(d), NightsCrew(d)), CStr(ms(i, 1)), _
                  CStr(ms(i, 4)), "Staffing", CStr(ms(i, 5)), WhenText(d), "N:" & CStr(ms(i, 1)) & "|" & CLng(Int(d)) & "|" & CStr(ms(i, 3)), _
                  CDbl(Int(d)), thOpen(), thOpen(), thOpenFill(), WhenColor(d)
        Next i
        NASort shortRows, nShort
    End If

    ' C. shift switch paybacks still owed (v26: oldest first; one owed longer than PAYBACK_WARN_DAYS
    '    gets its age in the status, in red on amber)
    Dim ow As Variant, owIds As Variant, swD As Date, swSh As String, swU As String, swOff As String, swWorks As String
    Dim owDays As Long
    On Error Resume Next
    ow = OwedSwitches(owIds)
    On Error GoTo 0
    If IsArray(owIds) Then
        For i = LBound(owIds) To UBound(owIds)
            id = CStr(owIds(i))
            If SwitchInfo(id, swD, swSh, swU, swOff, swWorks) Then
                If UnitInGroup(swU, grp) Then
                    owDays = CLng(Int(Date) - Int(swD))
                    st = OwedStatusText(owDays, PaybackWarnDays())
                    If owDays > PaybackWarnDays() Then
                        stC = thOpen(): stF = RGB(255, 236, 179)
                    Else
                        stC = TypeColor("SS"): stF = Tint(TypeColor("SS"), 0.85)
                    End If
                    NAAdd sr, nS, Format$(swD, "ddd mmm d"), swSh, IIf(swSh = "Days", DaysCrew(swD), NightsCrew(swD)), swU, _
                          swOff & " owes " & swWorks & " a shift   (" & id & ")", "Shift switch", st, _
                          WhenText(swD), "S:" & id, CDbl(Int(swD)), TypeColor("SS"), stC, stF, thMuted()
                End If
            End If
        Next i
        NASort sr, nS
    End If

    ' D. leave returning soon, and anything open-ended
    Dim code As String, eD As Date, nOpenEnded As Long, empNm As String
    If IsArray(mAbs) Then
        For i = 1 To UBound(mAbs, 1)
            If UCase$(SafeStr(mAbs(i, AB_STATUS))) <> "VOID" And IsDate(mAbs(i, AB_START)) Then
                If UnitInGroup(SafeStr(mAbs(i, AB_UNIT)), grp) Then
                    code = TypeCode(SafeStr(mAbs(i, AB_TYPE)))
                    sd = CDate(mAbs(i, AB_START))
                    st = ""
                    If Not IsDate(mAbs(i, AB_END)) Then
                        If Int(sd) <= Date + nDays - 1 Then
                            st = "Open-ended - no return date": eD = sd: nOpenEnded = nOpenEnded + 1
                        End If
                    ElseIf IsLeaveCode(code) Then
                        eD = CDate(mAbs(i, AB_END)) + 1
                        If Int(eD) >= Date And Int(eD) <= Date + nDays - 1 Then st = "Returns " & Format$(eD, "ddd mmm d")
                    End If
                    If st <> "" Then
                        empNm = SafeStr(mAbs(i, AB_EMP))
                        NAAdd br, nB, Format$(sd, "ddd mmm d"), "", EmpCrewOf(empNm), SafeStr(mAbs(i, AB_UNIT)), _
                              CalName(empNm), TypeLong(code), st, WhenText(eD), "B:" & SafeStr(mAbs(i, AB_ID)), _
                              IIf(Left$(st, 4) = "Open", 0, 1000000#) + CDbl(Int(eD)), TypeColor(code), TypeColor(code), _
                              Tint(TypeColor(code), 0.85), thMuted()
                    End If
                End If
            End If
        Next i
        NASort br, nB
    End If

    ' E. weekend holidays still waiting for the weekday the plant gives off (v20)
    Dim hv As Variant, hk As Variant, hd As Date
    ReDim hr(1 To 20)
    hLook = nDays: If hLook < HOL_CONFIRM_DAYS Then hLook = HOL_CONFIRM_DAYS
    hv = HolidaysToConfirm(Date, Date + hLook - 1)
    If IsArray(hv) Then
        For Each hk In hv
            hd = DateSerial(CLng(Left$(hk, 4)), CLng(Mid$(hk, 6, 2)), CLng(Mid$(hk, 9, 2)))
            NAAdd hr, nH, Format$(hd, "ddd mmm d"), "", "DAYS", "All units", Mid$(hk, 12), "Holiday", _
                  "Set the weekday off", WhenText(hd), "H:" & Format$(hd, "yyyy-mm-dd"), CDbl(hd), _
                  RGB(176, 0, 32), RGB(176, 0, 32), Tint(RGB(176, 0, 32), 0.88), thMuted()
        Next hk
    End If
End Sub

Private Sub NAGatherAdmin(ByVal grp As String, ByRef ir() As NARow, ByRef nI As Long, _
                          ByRef mr() As NARow, ByRef nM As Long, ByRef er() As NARow, ByRef nE As Long, _
                          ByRef ur() As NARow, ByRef nU As Long, ByRef dcR() As NARow, ByRef nDc As Long, _
                          ByRef vr() As NARow, ByRef nV As Long, ByRef tr() As NARow, ByRef nT As Long, _
                          ByRef rq() As NARow, ByRef nRq As Long)
    Dim i As Long
    ' R. members' requests waiting for an admin (v28, item 1): time off and switches asked for from the members' file
    ReDim rq(1 To 20)
    Dim rqv As Variant, rEmp As String, rFrom As Date, rHome As String
    If REQUESTS_ON Then rqv = RequestRows()
    If IsArray(rqv) Then
        For i = 1 To UBound(rqv, 1)
            If SafeStr(rqv(i, RQ_STATUS)) = RQ_ST_WAITING Then
                rEmp = SafeStr(rqv(i, RQ_EMP))
                rHome = HomeUnitCached(rEmp)
                If rHome = "" Or UnitInGroup(rHome, grp) Then
                    rFrom = Dt2(rqv(i, RQ_FROM))
                    NAAdd rq, nRq, Format$(rFrom, "ddd mmm d"), "", EmpCrewOf(rEmp), rHome, _
                          CalName(rEmp) & "   " & RequestLineText(rqv, i), SafeStr(rqv(i, RQ_KIND)), RQ_ST_WAITING, _
                          Format$(Dt2(rqv(i, RQ_RECEIVED)), "ddd mmm d"), "R:" & SafeStr(rqv(i, RQ_ID)), CDbl(Int(rFrom)), _
                          thPosted(), thOpen(), thOpenFill(), WhenColor(rFrom)
                End If
            End If
        Next i
        NASort rq, nRq
    End If
    ' H. imported vacation lines not yet booked (v22) - they don't keep people off call lists
    Dim iv As Variant
    ReDim ir(1 To 20)
    On Error Resume Next
    iv = ImportCandidates(Date, Date + IMPORT_WATCH_DAYS - 1)
    On Error GoTo 0
    If IsArray(iv) Then
        For i = 1 To UBound(iv, 1)
            If UnitInGroup(SafeStr(iv(i, 4)), grp) Then
                NAAdd ir, nI, Format$(Dt2(iv(i, 2)), "ddd mmm d"), SafeStr(iv(i, 3)), "", SafeStr(iv(i, 4)), _
                      SafeStr(iv(i, 5)) & IIf(SafeStr(iv(i, 6)) = "", "   (no roster match)", IIf(SafeStr(iv(i, 6)) = "?", "   (several matches)", "   = " & SafeStr(iv(i, 6)))), _
                      "Imported", "Not booked", WhenText(Dt2(iv(i, 2))), "I:" & SafeStr(iv(i, 1)), CDbl(Int(Dt2(iv(i, 2)))), _
                      TypeColor("VAC"), thOpen(), thOpenFill(), WhenColor(Dt2(iv(i, 2)))
            End If
        Next i
        NASort ir, nI
    End If

    ' G. members' file not published lately (M1)
    Dim pubAt As Date, pubTxt As String
    ReDim mr(1 To 2)
    pubAt = MembersPublishedAt()
    If pubAt = 0 Then
        pubTxt = "never published - click Member View on the roster"
    ElseIf Int(Date) - Int(pubAt) >= MembersStaleDays() Then
        pubTxt = "last published " & Format$(pubAt, "ddd mmm d, h:mm AM/PM") & " (" & (Int(Date) - Int(pubAt)) & " days ago)"
    End If
    If pubTxt <> "" Then
        NAAdd mr, nM, Format$(Date, "ddd mmm d"), "", "", "All units", "Members file: " & pubTxt, "Publish", _
              "Publish again", "", "M:publish", 0, thPosted(), thOpen(), thOpenFill(), thMuted()
    End If

    ' E. errors nobody has looked at yet (v24.2)
    Dim nErrLines As Long
    ReDim er(1 To 2)
    nErrLines = UnseenErrorCount()
    If nErrLines > 0 Then
        NAAdd er, nE, Format$(Date, "ddd mmm d"), "", "", "All units", nErrLines & " new line(s) in the Error Log", "Error", _
              "Open the log", "", "X:errors", 0, RGB(176, 0, 32), thOpen(), thOpenFill(), thMuted()
    End If
    ' v27 (item 24): call-out rows logged as Exception (a plant upset, nobody charged) in the last 30 days - for the superintendent
    Dim xl As Variant: xl = LogRowsNow()
    If IsArray(xl) Then
        Dim xi As Long, nX As Long
        For xi = 1 To UBound(xl, 1)
            If UCase$(Trim$(SafeStr(xl(xi, L_OUTCOME)))) = "EXCEPTION" And UCase$(Trim$(SafeStr(xl(xi, L_STATUS)))) <> "VOID" And IsDate(xl(xi, L_DATE)) Then
                If Int(CDate(xl(xi, L_DATE))) >= Int(Date) - 30 And nX < 12 Then
                    nX = nX + 1
                    NAAdd er, nE, Format$(CDate(xl(xi, L_DATE)), "ddd mmm d"), SafeStr(xl(xi, L_SHIFTLBL)), "", SafeStr(xl(xi, L_DEPT)), _
                          "Exception (plant upset): " & SafeStr(xl(xi, L_REF)) & "  " & CalName(SafeStr(xl(xi, L_EMP))) & "  -  " & Left$(SafeStr(xl(xi, L_TASK)), 60), "Exception", _
                          "Review", "", "Q:exception", 0, RGB(176, 0, 32), thOpen(), thOpenFill(), thMuted()
                End If
            End If
        Next xi
    End If
    ' v27 (item 23): the open disputes, one line each (an admin settles one through CloseDispute)
    Dim nDp As Long, dpTxt As String: dpTxt = OpenDisputesText(nDp)
    If nDp > 0 Then
        Dim dpLine As Variant
        For Each dpLine In Split(dpTxt, vbNewLine)
            NAAdd er, nE, Format$(Date, "ddd mmm d"), "", "", "All units", "Disputed: " & CStr(dpLine), "Dispute", _
                  "Alt+F8 > CloseDispute", "", "Q:dispute", 0, RGB(176, 0, 32), thOpen(), thOpenFill(), thMuted()
        Next dpLine
    End If
    ' v27 (item 21): the active OT restrictions, one line each (information: an admin ends one through EndRestriction)
    Dim nRs As Long, rsTxt As String: rsTxt = ActiveRestrictionsText(nRs)
    If nRs > 0 Then
        Dim rsLine As Variant
        For Each rsLine In Split(rsTxt, vbNewLine)
            NAAdd er, nE, Format$(Date, "ddd mmm d"), "", "", "All units", "OT restricted: " & CStr(rsLine), "Restriction", _
                  "Alt+F8 > EndRestriction", "", "Q:restriction", 0, RGB(176, 0, 32), thOpen(), thOpenFill(), thMuted()
        Next rsLine
    End If
    ' v27 (item 9): receipt answers ("Was this right?") nobody has read yet (modReceipt)
    Dim nFb As Long: nFb = UnseenFeedbackCount()
    If nFb > 0 Then
        NAAdd er, nE, Format$(Date, "ddd mmm d"), "", "", "All units", nFb & " receipt answer(s) not yet read", "Feedback", _
              "Open the list", "", "F:feedback", 0, thPosted(), thOpen(), thOpenFill(), thMuted()
    End If

    ' F. people with no Home Unit (v21) - Graeme fills these in on the roster
    Dim uv As Variant, uk As Variant
    ReDim ur(1 To 20)
    On Error Resume Next
    uv = BlankHomeUnits()
    On Error GoTo 0
    If IsArray(uv) Then
        For Each uk In uv
            NAAdd ur, nU, "", "", EmpCrewOf(CStr(uk)), "", CStr(uk), "Roster", "Home Unit blank", "", _
                  "U:" & CStr(uk), 0, thMuted(), thOpen(), thOpenFill(), thMuted()
        Next uk
    End If

    ' D2. v27 (item 25, 1.4): the DATA CHECK - roster details that look wrong (modPeople.DataCheckLines)
    Dim dcV As Variant, dcL As Variant, dcP As Variant
    ReDim dcR(1 To 20)
    On Error Resume Next
    dcV = DataCheckLines()
    On Error GoTo 0
    If IsArray(dcV) Then
        For Each dcL In dcV
            dcP = Split(CStr(dcL) & vbTab, vbTab)
            NAAdd dcR, nDc, "", "", EmpCrewOf(CStr(dcP(0))), "", CStr(dcP(0)) & ":  " & CStr(dcP(1)), "Roster", "Data check", "", _
                  "U:" & CStr(dcP(0)), 0, thMuted(), thPosted(), -1, thMuted()
        Next dcL
    End If

    ' V. unit moves entered ahead of time (v24.3): waiting for their date, or due and not applied
    Dim pmv As Variant, pmD As Date, pmDue As Boolean, pmWho As String
    ReDim vr(1 To 10)
    On Error Resume Next
    pmv = PlannedRows()
    On Error GoTo 0
    If IsArray(pmv) Then
        For i = 1 To UBound(pmv, 1)
            pmD = Dt2(pmv(i, PM_FROM))
            pmDue = (Int(pmD) <= Date)                  ' a date that can't be read (0) shows as a problem too
            pmWho = SafeStr(pmv(i, PM_EMP))
            NAAdd vr, nV, IIf(pmD = 0, "?", Format$(pmD, "ddd mmm d")), "", SafeStr(pmv(i, PM_CREW)), SafeStr(pmv(i, PM_UNIT)), _
                  CalName(pmWho) & "   " & HomeUnitCached(pmWho) & " -> " & SafeStr(pmv(i, PM_UNIT)) & " (" & SafeStr(pmv(i, PM_CODE)) & ")   " & _
                  SafeStr(pmv(i, PM_ID)), "Unit move", IIf(pmD = 0, "Date can't be read", IIf(pmDue, "DUE - not applied", "Planned")), _
                  IIf(pmD = 0, "", WhenText(pmD)), "V:" & SafeStr(pmv(i, PM_ID)), CDbl(Int(pmD)), thMuted(), _
                  IIf(pmDue, thOpen(), thMuted()), IIf(pmDue, thOpenFill(), -1), WhenColor(pmD)
        Next i
        NASort vr, nV
    End If

    ' T. attendance (v22.1): sick / personal days at or over the limit, by home unit
    Dim rn As Variant, ast As AttendanceStatus, lvl As Long, hu As String, nmT As String
    ReDim tr(1 To 20)
    On Error Resume Next
    rn = RosterNames()
    If IsArray(rn) Then
        BeginAttendanceBatch
        For i = LBound(rn) To UBound(rn)
            nmT = CStr(rn(i))
            hu = HomeUnitCached(nmT)
            If hu = "" Or UnitInGroup(hu, grp) Then
                If AttendanceFor(nmT, Date, ast) Then
                    lvl = AttendanceLevel(ast)
                    If lvl > 0 Then
                        NAAdd tr, nT, IIf(ast.LastSick > 0, Format$(ast.LastSick, "ddd mmm d"), ""), "", EmpCrewOf(nmT), hu, _
                              CalName(nmT) & "   " & ast.SickUsed & " sick / " & ast.PersonalUsed & " personal, " & ast.Occurrences & " occ.", _
                              "Attendance", AttendanceStatusText(ast), _
                              IIf(ast.WaitUntil > 0, "until " & Format$(ast.WaitUntil, "mmm d, yyyy"), ""), _
                              "T:" & nmT, CDbl(-lvl * 1000 - ast.Combined), TypeColor("SIC"), _
                              IIf(lvl = 2, thOpen(), thPosted()), IIf(lvl = 2, thOpenFill(), thPostedFill()), thMuted()
                    End If
                End If
            End If
        Next i
        EndAttendanceBatch
        NASort tr, nT
    End If
    On Error GoTo 0
End Sub

Private Sub NADrawHead(ByVal ws As Worksheet, ByVal grp As String, ByVal nDays As Long)
    ' ---- draw ----
    ws.Rows("1:3000").Delete
    With ws.Cells.Font
        .Name = CAL_FONT
        .Size = 9
    End With
    ws.Cells.Interior.Color = vbWhite
    ws.Columns(1).ColumnWidth = 1.5
    ws.Columns(2).ColumnWidth = 13
    ws.Columns(3).ColumnWidth = 8
    ws.Columns(4).ColumnWidth = 7
    ws.Columns(5).ColumnWidth = 17
    ws.Columns(6).ColumnWidth = 46
    ws.Columns(7).ColumnWidth = 16
    ws.Columns(8).ColumnWidth = 26
    ws.Columns(9).ColumnWidth = 13
    ws.Columns(10).ColumnWidth = 2
    ws.Range(ws.Columns(NA_KEY), ws.Columns(NA_KEY + 12)).Hidden = True

    ws.Rows(1).RowHeight = 38
    ws.Range("A1:J1").Interior.Color = thBand()
    With ws.Range("B1:E1")
        .Merge
        .Value = "NEEDS ATTENTION"
        .Font.Name = CAL_FONT_B: .Font.Size = 15: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 22
    With ws.Range("B2:I2")
        .Merge
        .Value = grp & "   " & ChrW(183) & "   " & Format$(Date, "ddd mmm d") & " to " & Format$(Date + nDays - 1, "ddd mmm d") & _
                 "   " & ChrW(183) & "   updated " & Format$(Now, "h:mm AM/PM") & "   " & ChrW(183) & "   double-click any row for options" & _
                 IIf(RolloverNote() <> "", "   " & ChrW(183) & "   " & RolloverNote(), "")
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = thMuted()
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 8
    ws.Rows(4).RowHeight = 30
    ws.Rows(5).RowHeight = 16
    ws.Rows(6).RowHeight = 12
End Sub


Private Function EmpCrewOf(ByVal empName As String) As String
    If mShiftMap Is Nothing Then Exit Function
    If mShiftMap.Exists(Trim$(empName)) Then EmpCrewOf = CrewOn(empName, Date, CStr(mShiftMap(Trim$(empName))))   ' v24: today's crew
End Function

' coloured bar + column headings; returns the first row for data
Private Function NASection(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal n As Long, _
                           ByVal accent As Long, ByVal heads As Variant) As Long
    ws.Rows(r).RowHeight = 24
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = accent
    With ws.Cells(r, 2)
        .Value = title & "     " & n
        .Font.Name = CAL_FONT_B: .Font.Size = 10.5: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    Dim i As Long
    ws.Rows(r + 1).RowHeight = 18
    For i = 0 To 7
        With ws.Cells(r + 1, 2 + i)
            .Value = heads(i)
            .Font.Name = CAL_FONT_B: .Font.Size = 8: .Font.Color = thMuted()
            .Interior.Color = RGB(245, 247, 248)
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i
    ws.Cells(r + 1, 1).Interior.Color = RGB(245, 247, 248)
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, 9)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = thGrid()
    End With
    NASection = r + 2
End Function

' the rows of one section (zebra, status pill); returns the next free row
Private Function NARows(ByVal ws As Worksheet, ByVal r As Long, ByRef recs() As NARow, ByVal n As Long, _
                        ByVal emptyMsg As String) As Long
    Dim i As Long, c As Long
    If n = 0 Then
        ws.Rows(r).RowHeight = 20
        With ws.Cells(r, 2)
            .Value = emptyMsg
            .Font.Italic = True: .Font.Color = thGreen()
            .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
        ws.Rows(r + 1).RowHeight = 14
        NARows = r + 2
        Exit Function
    End If
    ' v25.2 (item 4): the section's rows as ONE array write and its formats by class (Union), as the Calendar's
    ' weeks (item 3); the row-by-row writer it replaced was kept one release for SmokeTest's compare step (gone in v25.5)
    Dim vals() As Variant: ReDim vals(1 To n, 1 To 8)
    Dim keyV() As Variant: ReDim keyV(1 To n, 1 To 1)
    Dim rgAlt As Range, whenMuted As Range, whenBold As Range
    Dim accents As Object: Set accents = CreateObject("Scripting.Dictionary")     ' accent colour -> column A cells
    Dim stCols As Object: Set stCols = CreateObject("Scripting.Dictionary")       ' status font colour -> column H cells
    Dim stFills As Object: Set stFills = CreateObject("Scripting.Dictionary")     ' status fill -> column H cells
    Dim whenCols As Object: Set whenCols = CreateObject("Scripting.Dictionary")   ' "when" font colour -> column I cells
    Dim rowAt As Long, key As Variant
    For i = 1 To n
        rowAt = r + i - 1
        For c = 0 To 7
            vals(i, c + 1) = "'" & recs(i).Cell(c)
        Next c
        keyV(i, 1) = recs(i).Key
        If i Mod 2 = 0 Then AddToRange rgAlt, ws.Range(ws.Cells(rowAt, 2), ws.Cells(rowAt, 9))
        DictAddRange accents, recs(i).Accent, ws.Cells(rowAt, 1)
        DictAddRange stCols, recs(i).StatusCol, ws.Cells(rowAt, 8)
        If recs(i).StatusFill >= 0 Then DictAddRange stFills, recs(i).StatusFill, ws.Cells(rowAt, 8)
        DictAddRange whenCols, recs(i).WhenCol, ws.Cells(rowAt, 9)
        If recs(i).WhenCol = thMuted() Then AddToRange whenMuted, ws.Cells(rowAt, 9) Else AddToRange whenBold, ws.Cells(rowAt, 9)
    Next i
    Dim blk As Range: Set blk = ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 9))
    blk.Value = vals
    ws.Range(ws.Cells(r, NA_KEY), ws.Cells(r + n - 1, NA_KEY)).Value = keyV
    ws.Range(ws.Rows(r), ws.Rows(r + n - 1)).RowHeight = 20
    With blk
        .VerticalAlignment = xlCenter: .IndentLevel = 1: .ShrinkToFit = True
    End With
    If Not rgAlt Is Nothing Then rgAlt.Interior.Color = RGB(250, 250, 251)
    For Each key In accents.Keys
        accents(key).Interior.Color = CLng(key)
    Next key
    ws.Range(ws.Cells(r, 6), ws.Cells(r + n - 1, 6)).Font.Name = CAL_FONT_B
    ws.Range(ws.Cells(r, 8), ws.Cells(r + n - 1, 8)).Font.Name = CAL_FONT_B
    For Each key In stCols.Keys
        stCols(key).Font.Color = CLng(key)
    Next key
    For Each key In stFills.Keys
        stFills(key).Interior.Color = CLng(key)
    Next key
    For Each key In whenCols.Keys
        whenCols(key).Font.Color = CLng(key)
    Next key
    If Not whenMuted Is Nothing Then whenMuted.Font.Name = CAL_FONT
    If Not whenBold Is Nothing Then whenBold.Font.Name = CAL_FONT_B
    ' each row's bottom edge: the block's inside lines plus its bottom edge
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 9))
        If n > 1 Then
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = thGrid()
            End With
        End If
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = thGrid()
        End With
    End With
    ' v25.5 (bundle item 1, Graeme Oct 17): no hover notes - the row already shows what the note carried
    r = r + n
    ws.Rows(r).RowHeight = 14
    NARows = r + 1
End Function

' v25.2 (item 4): a range per key grown by Union (the key is a colour)
Private Sub DictAddRange(ByVal dict As Object, ByVal key As Long, ByVal part As Range)
    If dict.Exists(key) Then Set dict(key) = Union(dict(key), part) Else dict.Add key, part
End Sub

' summary tile over columns c1..c2, rows 4-5
Private Sub NATile(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal n As Long, ByVal cap As String, _
                   ByVal col As Long, ByVal jumpRow As Long)
    With ws.Range(ws.Cells(4, c1), ws.Cells(4, c2))
        .Merge
        .Value = n
        .Font.Name = CAL_FONT_B: .Font.Size = 20: .Font.Color = IIf(n > 0, col, RGB(176, 190, 197))
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlBottom: .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(5, c1), ws.Cells(5, c2))
        .Merge
        .Value = cap
        .Font.Name = CAL_FONT_B: .Font.Size = 7.5: .Font.Color = thMuted()
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(4, c1), ws.Cells(5, c2))
        .Interior.Color = Tint(col, 0.9)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = col
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = vbWhite
        End With
    End With
    ' jump key for this tile (hidden columns to the right of the key column)
    ws.Cells(4, NA_KEY + c1).Value = "J:" & jumpRow
    ws.Cells(5, NA_KEY + c1).Value = "J:" & jumpRow
End Sub

Private Sub BuildNAControls(ByVal ws As Worksheet, ByVal grp As String, ByVal nDays As Long)
    On Error Resume Next                            ' never raises
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "na_" Then ws.Shapes(i).Delete
    Next i
    CalDrop ws, "na_unit", ws.Range("F1")
    CalDrop ws, "na_range", ws.Range("G1")
    CalBtn ws, "na_refresh", "Refresh", "modNeeds.NARefresh", ws.Range("H1"), thNav()
    CalBtn ws, "na_cal", "Calendar", "modNeeds.NAGoCalendar", ws.Range("I1"), thGreen()

    Dim dd As Object, g As Variant, idx As Long, rg As Variant
    Set dd = Nothing
    Set dd = ws.DropDowns("na_unit")
    If Not dd Is Nothing Then
        g = CalGroupNames()
        For i = 0 To UBound(g)
            dd.AddItem g(i)
            If StrComp(g(i), grp, vbTextCompare) = 0 Then idx = i + 1
        Next i
        If idx = 0 Then idx = UBound(g) + 1
        dd.ListIndex = idx
        dd.OnAction = "modNeeds.NAUnitChanged"
    End If
    Set dd = Nothing
    Set dd = ws.DropDowns("na_range")
    If Not dd Is Nothing Then
        rg = NARanges()
        idx = 2
        For i = 0 To UBound(rg)
            dd.AddItem "Next " & rg(i) & " days"
            If rg(i) = nDays Then idx = i + 1
        Next i
        dd.ListIndex = idx
        dd.OnAction = "modNeeds.NARangeChanged"
    End If
End Sub

