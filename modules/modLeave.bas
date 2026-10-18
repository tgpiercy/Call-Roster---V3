Attribute VB_Name = "modLeave"
Option Explicit

' ====================================================================
'  SICK CALLS, LEAVE AND IMPORTED LINES (v22)
'
'  Pure rules (mirrored in tests/rules.py and checked by SelfTest):
'   SickBlocks        a sick call keeps the person off every shift that
'                     starts within 24 h after the sick shift ends
'   HoursAsOfCore     a person's Total Hours on a past date: the Hours
'                     History snapshot with the latest data-through <= d,
'                     plus non-void charges with shift date <= d, plus
'                     provisional accepted OT as of that snapshot
'   PeerEligible      who counts as a leave-top-up peer
'   LeaveAverage      the top-up: average of the peers' gains
'   MergeRuns         consecutive imported dates -> one booking
'
'  Sheet-facing: TotalHoursAsOf, ComputeLeaveReturn, SaveLeaveGoing,
'  SaveLeaveReturn, SaveSickCall, ImportCandidates / ConvertImports.
'  Hours changes are Adjustment rows in the OT Log (Ref LV-####, Reason
'  "Leave return", Note = the calculation), undone with Void Entry.
' ====================================================================

Public Const LV_VERSION As String = "Leave v25.3 (Oct 17)"
Public Const PEER_MAX_GRADE As Double = 2        ' leave top-up peers: grade 2 and below
Public Const SNAPSHOT_WARN_DAYS As Long = 7      ' flag a snapshot this far from the date wanted
Public Const IMPORT_WATCH_DAYS As Long = 56      ' Needs Attention: imported lines not yet booked

' compact arrays the pure core works on
'   hh(i, 1..3): employee, data-through date, worked hours
'   lg(i, 1..6): employee, shift date, outcome, status, shift hours, hours charged, (7) is night
Private Const HH_EMP As Long = 1
Private Const HH_THROUGH As Long = 2
Private Const HH_HOURS As Long = 3
Private Const LG_EMP As Long = 1
Private Const LG_DATE As Long = 2
Private Const LG_OUTCOME As Long = 3
Private Const LG_STATUS As Long = 4
Private Const LG_HOURS As Long = 5
Private Const LG_CHARGED As Long = 6
Private Const LG_NIGHT As Long = 7

' what the Returning pop-up shows and saves
Public Type LeaveCalc
    AbsId As String
    Emp As String
    HomeUnit As String
    FirstOff As Date
    ReturnDate As Date
    WindowFrom As Date              ' first day off, clipped to the year start
    NPeers As Long
    PeerNames As String
    Average As Double
    Before As Double
    After As Double
    SnapshotText As String
    Warnings As String
    Summary As String
End Type

' batch cache for the leave maths (many TotalHoursAsOf calls)
Private mHH As Variant
Private mLG As Variant
Private mYearFrom As Date
Private mYearTo As Date
Private mBatch As Boolean


' ====================================================================
'  PURE RULES
' ====================================================================
' Days 7 AM - 7 PM, nights 7 PM - 7 AM next morning (dated by the evening).
' Blocked: the sick shift itself and any shift starting within 24 h after it ends.
Public Function SickBlocks(ByVal sickDate As Date, ByVal sickNight As Boolean, ByVal d As Date, ByVal night As Boolean) As Boolean
    Dim sickStart As Double, sickEnd As Double, candStart As Double
    sickStart = Int(sickDate) + IIf(sickNight, 19 / 24, 7 / 24)
    sickEnd = sickStart + 0.5
    candStart = Int(d) + IIf(night, 19 / 24, 7 / 24)
    SickBlocks = (candStart >= sickStart - 0.0001 And candStart < sickEnd + 1 - 0.0001)
End Function

' The Hours History snapshot to use for emp as of d: the latest data-through
' <= d within the roster year. None in the year on or before d: the earliest
' one in the year, with estimated = True. Snapshots dated before the year
' start are last year's figures and are ignored (worked = 0 at the reset:
' True, hours 0, estimated False when such older snapshots exist).
' False = no snapshot at all.
Public Function PickSnapshot(ByRef hh As Variant, ByVal emp As String, ByVal d As Date, ByVal yearFrom As Date, _
                             ByRef through As Date, ByRef hours As Double, ByRef estimated As Boolean) As Boolean
    through = 0: hours = 0: estimated = False
    If Not IsArray(hh) Then Exit Function
    Dim i As Long, t As Date, bestLE As Date, bestAny As Date, hLE As Double, hAny As Double
    Dim anyFound As Boolean, leFound As Boolean, olderFound As Boolean
    bestAny = DateSerial(2999, 12, 31)
    For i = 1 To UBound(hh, 1)
        If StrComp(Trim$(SafeStr(hh(i, HH_EMP))), Trim$(emp), vbTextCompare) = 0 And IsDate(hh(i, HH_THROUGH)) Then
            t = Int(CDate(hh(i, HH_THROUGH)))
            If t < Int(yearFrom) Then
                olderFound = True
            Else
                If t <= Int(d) Then
                    If Not leFound Or t > bestLE Then bestLE = t: hLE = NumOr0(hh(i, HH_HOURS)): leFound = True
                End If
                If t < bestAny Then bestAny = t: hAny = NumOr0(hh(i, HH_HOURS)): anyFound = True
            End If
        End If
    Next i
    If leFound Then
        through = bestLE: hours = hLE
    ElseIf olderFound And Int(d) < Int(yearFrom) + 14 Then
        ' v23: in the year's first pay period, before any import of the new year, nothing has been worked yet
        through = 0: hours = 0: estimated = False
    ElseIf anyFound Then
        through = bestAny: hours = hAny: estimated = True
    ElseIf olderFound Then
        through = 0: hours = 0: estimated = False       ' the reset: nothing worked yet this year
    Else
        Exit Function
    End If
    PickSnapshot = True
End Function

' Non-void Hours Charged for emp with shift date in [yearFrom, min(d, yearTo)]
Public Function ChargedAsOf(ByRef lg As Variant, ByVal emp As String, ByVal d As Date, _
                            ByVal yearFrom As Date, ByVal yearTo As Date) As Double
    If Not IsArray(lg) Then Exit Function
    Dim i As Long, sd As Date
    For i = 1 To UBound(lg, 1)
        If StrComp(Trim$(SafeStr(lg(i, LG_EMP))), Trim$(emp), vbTextCompare) = 0 And UCase$(SafeStr(lg(i, LG_STATUS))) <> "VOID" Then
            If IsDate(lg(i, LG_DATE)) Then
                sd = Int(CDate(lg(i, LG_DATE)))
                If sd >= Int(yearFrom) And sd <= Int(yearTo) And sd <= Int(d) Then ChargedAsOf = ChargedAsOf + NumOr0(lg(i, LG_CHARGED))
            End If
        End If
    Next i
End Function

' Provisional accepted OT as of d, given the snapshot's data-through:
' Accepted / Awarded non-void rows with shift date <= d (in the year) whose
' shift had not ended by `through` (v21 rule, ProvisionalCovered).
Public Function ProvisionalAsOf(ByRef lg As Variant, ByVal emp As String, ByVal d As Date, ByVal through As Date, _
                                ByVal yearFrom As Date, ByVal yearTo As Date) As Double
    If Not IsArray(lg) Then Exit Function
    Dim i As Long, sd As Date, o As String
    For i = 1 To UBound(lg, 1)
        If StrComp(Trim$(SafeStr(lg(i, LG_EMP))), Trim$(emp), vbTextCompare) = 0 And UCase$(SafeStr(lg(i, LG_STATUS))) <> "VOID" Then
            o = SafeStr(lg(i, LG_OUTCOME))
            If (o = "Accepted" Or o = "Awarded") And IsDate(lg(i, LG_DATE)) Then
                sd = Int(CDate(lg(i, LG_DATE)))
                If sd >= Int(yearFrom) And sd <= Int(yearTo) And sd <= Int(d) Then
                    If Not ProvisionalCovered(sd, CBool(lg(i, LG_NIGHT)), through) Then ProvisionalAsOf = ProvisionalAsOf + NumOr0(lg(i, LG_HOURS))
                End If
            End If
        End If
    Next i
End Function

' Total Hours for emp as of d (0 before the year start = the reset)
Public Function HoursAsOfCore(ByRef hh As Variant, ByRef lg As Variant, ByVal emp As String, ByVal d As Date, _
                              ByVal yearFrom As Date, ByVal yearTo As Date, _
                              ByRef through As Date, ByRef estimated As Boolean) As Double
    through = 0: estimated = False
    If Int(d) < Int(yearFrom) Then Exit Function
    Dim worked As Double
    If Not PickSnapshot(hh, emp, d, yearFrom, through, worked, estimated) Then estimated = True
    HoursAsOfCore = worked + ChargedAsOf(lg, emp, d, yearFrom, yearTo) + ProvisionalAsOf(lg, emp, d, through, yearFrom, yearTo)
End Function

' A leave top-up peer: grade 2 and below, not the returner, no leave of their
' own overlapping the window, and Hours History before the window start.
Public Function PeerEligible(ByVal grade As Double, ByVal isReturner As Boolean, ByVal leaveOverlaps As Boolean, _
                             ByVal historyBefore As Boolean) As Boolean
    PeerEligible = (grade <= PEER_MAX_GRADE) And Not isReturner And Not leaveOverlaps And historyBefore
End Function

' Average of the peers' gains, rounded to 0.1 half away from zero (VBA's
' Round is banker's rounding, which the Python mirror doesn't share).
' 0 when there are no peers.
Public Function LeaveAverage(ByRef deltas() As Double, ByVal n As Long) As Double
    If n <= 0 Then Exit Function
    Dim i As Long, s As Double, x As Double
    For i = 1 To n
        s = s + deltas(i)
    Next i
    x = s / n
    LeaveAverage = Sgn(x) * Int(Abs(x) * 10 + 0.5 + 0.000001) / 10
End Function

' Consecutive dates (sorted, unique) -> runs. Returns the count; runs(k, 1) = first, runs(k, 2) = last.
Public Function MergeRuns(ByRef dates() As Date, ByVal n As Long, ByRef runs() As Date) As Long
    If n <= 0 Then Exit Function
    ReDim runs(1 To n, 1 To 2)
    Dim i As Long, k As Long
    k = 1: runs(1, 1) = Int(dates(1)): runs(1, 2) = Int(dates(1))
    For i = 2 To n
        If Int(dates(i)) = runs(k, 2) + 1 Then
            runs(k, 2) = Int(dates(i))
        ElseIf Int(dates(i)) > runs(k, 2) Then
            k = k + 1: runs(k, 1) = Int(dates(i)): runs(k, 2) = Int(dates(i))
        End If
    Next i
    MergeRuns = k
End Function


' ====================================================================
'  TOTAL HOURS AS OF A DATE (sheet-facing)
' ====================================================================
' Loads Hours History and the OT Log once for a run of calls
Public Sub BeginHoursBatch()
    mHH = CompactHH()
    mLG = CompactLog()
    Dim ws As Worksheet: Set ws = RosterSheet()
    mYearFrom = RosterYearFrom()
    mYearTo = RosterYearTo()
    mBatch = True
End Sub

Public Sub EndHoursBatch()
    mBatch = False
    mHH = Empty: mLG = Empty
End Sub

' Total Hours on date d, the snapshot date used, and a note for the confirm
Public Function TotalHoursAsOf(ByVal emp As String, ByVal d As Date, Optional ByRef snapThrough As Date, _
                               Optional ByRef note As String) As Double
    Dim own As Boolean
    If Not mBatch Then BeginHoursBatch: own = True
    Dim est As Boolean
    TotalHoursAsOf = HoursAsOfCore(mHH, mLG, emp, d, mYearFrom, mYearTo, snapThrough, est)
    note = ""
    If snapThrough = 0 Then
        If HasAnyHistory(emp) Then
            note = "worked hours taken as 0 for " & emp & " (the year reset, before this year's first Kronos import)"
        Else
            note = "no Hours History for " & emp & " - worked hours taken as 0"
        End If
    ElseIf est Then
        note = "start figure estimated from the " & Format$(snapThrough, "mmm d") & " snapshot (none on or before " & Format$(d, "mmm d") & ")"
    ElseIf Abs(Int(d) - Int(snapThrough)) > SNAPSHOT_WARN_DAYS Then
        note = "snapshot is " & Abs(Int(d) - Int(snapThrough)) & " days from " & Format$(d, "mmm d")
    End If
    If own Then EndHoursBatch
End Function

' Hours History -> (employee, data-through, hours); v20 rows: file date - 1
Private Function CompactHH() As Variant
    Dim lo As ListObject: Set lo = HoursHistoryTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim out() As Variant: ReDim out(1 To UBound(v, 1), 1 To 3)
    Dim i As Long, nc As Long: nc = UBound(v, 2)
    Dim yrNow As Long, yrFrom As Date
    On Error Resume Next
    yrNow = RosterYear(): yrFrom = RosterYearFrom()
    On Error GoTo 0
    For i = 1 To UBound(v, 1)
        out(i, HH_EMP) = SafeStr(v(i, 4))
        If nc >= 7 Then
            If IsDate(v(i, 7)) Then out(i, HH_THROUGH) = Int(CDate(v(i, 7)))
        End If
        If IsEmpty(out(i, HH_THROUGH)) And IsDate(v(i, 3)) Then out(i, HH_THROUGH) = Int(CDate(v(i, 3))) - 1
        If nc >= 8 And yrNow > 0 And Not IsEmpty(out(i, HH_THROUGH)) Then       ' v23: the Year column
            If IsNumeric(v(i, 8)) And Not IsEmpty(v(i, 8)) Then
                out(i, HH_THROUGH) = SnapshotThroughForYear(CDate(out(i, HH_THROUGH)), CLng(v(i, 8)), yrNow, yrFrom)
            End If
        End If
        out(i, HH_HOURS) = NumOr0(v(i, 5))
    Next i
    CompactHH = out
End Function

' PURE (SelfTest "snapyear"): a snapshot stamped with an EARLIER roster year
' holds last year's hours even when its data-through date falls inside this
' year's first period (an import in the last days of December, before the
' rollover). It is treated as the day before the year starts, so
' PickSnapshot reads it as last year's. No stamp (0) = judge by date as before.
Public Function SnapshotThroughForYear(ByVal through As Date, ByVal snapYr As Long, ByVal curYr As Long, ByVal yrFrom As Date) As Date
    SnapshotThroughForYear = through
    If snapYr <= 0 Or snapYr >= curYr Then Exit Function
    If Int(through) >= Int(yrFrom) Then SnapshotThroughForYear = Int(yrFrom) - 1
End Function

' OT Log -> (employee, shift date, outcome, status, hours, charged, is night)
Private Function CompactLog() As Variant
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim out() As Variant: ReDim out(1 To UBound(v, 1), 1 To 7)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        out(i, LG_EMP) = SafeStr(v(i, L_EMP))
        out(i, LG_DATE) = v(i, L_DATE)
        out(i, LG_OUTCOME) = SafeStr(v(i, L_OUTCOME))
        out(i, LG_STATUS) = SafeStr(v(i, L_STATUS))
        out(i, LG_HOURS) = NumOr0(v(i, L_BASEHRS))
        out(i, LG_CHARGED) = NumOr0(v(i, L_CHARGED))
        out(i, LG_NIGHT) = IsNightShift(SafeStr(v(i, L_SHIFTLBL)))
    Next i
    CompactLog = out
End Function

' True when emp has a Hours History snapshot with data-through < d (in the year)
Private Function HasHistoryBefore(ByVal emp As String, ByVal d As Date) As Boolean
    If Not IsArray(mHH) Then Exit Function
    Dim i As Long, t As Date
    For i = 1 To UBound(mHH, 1)
        If StrComp(Trim$(SafeStr(mHH(i, HH_EMP))), Trim$(emp), vbTextCompare) = 0 And IsDate(mHH(i, HH_THROUGH)) Then
            t = Int(CDate(mHH(i, HH_THROUGH)))
            If t < Int(d) And t >= Int(mYearFrom) Then HasHistoryBefore = True: Exit Function
        End If
    Next i
End Function

' True when emp has any Hours History snapshot at all (any year)
Private Function HasAnyHistory(ByVal emp As String) As Boolean
    If Not IsArray(mHH) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(mHH, 1)
        If StrComp(Trim$(SafeStr(mHH(i, HH_EMP))), Trim$(emp), vbTextCompare) = 0 And IsDate(mHH(i, HH_THROUGH)) Then HasAnyHistory = True: Exit Function
    Next i
End Function

' Earliest data-through date in the year (0 = no Hours History yet)
Private Function EarliestThrough() As Date
    If Not IsArray(mHH) Then Exit Function
    Dim i As Long, t As Date, best As Date
    For i = 1 To UBound(mHH, 1)
        If IsDate(mHH(i, HH_THROUGH)) Then
            t = Int(CDate(mHH(i, HH_THROUGH)))
            If t >= Int(mYearFrom) Then
                If best = 0 Or t < best Then best = t
            End If
        End If
    Next i
    EarliestThrough = best
End Function


' ====================================================================
'  LEAVE
' ====================================================================
Public Sub ShowLeaveForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not AbsenceColumnsCurrent() Then
        MsgBox "Run Alt+F8 > UpgradeAbsenceColumns first (the Absences table needs its v22 columns).", vbExclamation
        Exit Sub
    End If
    frmLeave.Show
    Exit Sub
Oops:
    ReportError "ShowLeaveForm"
End Sub

' Going on leave: the booking, with the home unit locked in and a leave group id
Public Function SaveLeaveGoing(ByVal emp As String, ByVal typeText As String, ByVal firstOff As Date, _
                               ByVal expectedReturn As Variant, ByVal note As String, ByRef homeUnit As String) As String
    homeUnit = HomeUnitOf(emp)
    Dim unit As String: unit = homeUnit
    If unit = "" Then
        Dim recs() As EmpRec, d As Variant, n As Long
        For Each d In DeptList()
            n = LoadEligible(CStr(d), "ALL", recs)
            Dim i As Long
            For i = 1 To n
                If StrComp(recs(i).EmpName, emp, vbTextCompare) = 0 Then unit = CStr(d): Exit For
            Next i
            If unit <> "" Then Exit For
        Next d
    End If
    Dim lo As ListObject: Set lo = AbsTable(True)
    If lo Is Nothing Then Exit Function
    Dim grp As String: grp = NextGroupId(lo)
    Dim absEnd As Variant
    If IsDate(expectedReturn) Then absEnd = CDate(expectedReturn) - 1 Else absEnd = Empty
    SaveLeaveGoing = SaveAbsence(emp, unit, typeText, firstOff, absEnd, Empty, "", note, homeUnit, grp, "", "AB")
    If SaveLeaveGoing <> "" Then MarkCalDirty
End Function

' Next LV-#### across the Absences table's Leave Group column and the OT Log's Ref IDs
Private Function NextGroupId(ByVal lo As ListObject) As String
    Dim mx As Long, i As Long, s As String
    s = NextRefID("LV")                              ' OT Log side
    mx = CLng(Val(Mid$(s, 4))) - 1
    If Not lo.DataBodyRange Is Nothing And lo.ListColumns.Count >= AB_GROUP Then
        Dim v As Variant: v = RangeTo2D(lo.ListColumns(AB_GROUP).DataBodyRange)
        For i = 1 To UBound(v, 1)
            s = UCase$(SafeStr(v(i, 1)))
            If Left$(s, 3) = "LV-" Then
                If CLng(Val(Mid$(s, 4))) > mx Then mx = CLng(Val(Mid$(s, 4)))
            End If
        Next i
    End If
    NextGroupId = "LV-" & Format$(mx + 1, "0000")
End Function

Public Function HomeUnitOf(ByVal emp As String) As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), Trim$(emp), vbTextCompare) = 0 Then
            HomeUnitOf = Trim$(SafeStr(ws.Cells(r, C_HOME).Value))
            Exit Function
        End If
    Next r
End Function

' Works out the top-up for one open leave booking. False = can't (message in lc.Warnings).
Public Function ComputeLeaveReturn(ByVal absId As String, ByVal returnDate As Date, ByRef lc As LeaveCalc) As Boolean
    Dim blank As LeaveCalc: lc = blank                       ' a fresh result every time (no stale warnings)
    Dim emp As String, unit As String, code As String, dStart As Date, dEnd As Variant, home As String, grp As String, st As String
    If Not AbsenceInfo(absId, emp, unit, code, dStart, dEnd, home, grp, st) Then lc.Warnings = absId & " not found": Exit Function
    If st = "VOID" Then lc.Warnings = absId & " is void": Exit Function
    If Not IsLeaveCode(code) Then lc.Warnings = absId & " isn't a leave booking": Exit Function
    If LeaveToppedUp(grp) Then lc.Warnings = absId & " has already been topped up (" & grp & " in the OT Log). Void that entry first if it was wrong.": Exit Function
    If Int(returnDate) <= Int(dStart) Then lc.Warnings = "The return date must be after the first day off (" & Format$(dStart, "mmm d, yyyy") & ")": Exit Function
    If Int(returnDate) > Date Then lc.Warnings = "The return date is after today: log the return on or after the first day back.": Exit Function    ' v25 (D-C7), in the rule itself (review)
    If home = "" Then home = HomeUnitOf(emp)
    If home = "" Then lc.Warnings = emp & " has no Home Unit (roster column B) and none was locked in on the booking - fill it in first": Exit Function
    lc.AbsId = absId: lc.Emp = emp: lc.HomeUnit = home: lc.FirstOff = Int(dStart): lc.ReturnDate = Int(returnDate)

    On Error GoTo Failed
    BeginHoursBatch
    If Int(returnDate) < Int(mYearFrom) Then     ' v23: that year is closed and its rows are archived
        lc.Warnings = "The return date is in a year that has been rolled over and closed. If hours are still owed, use Add Adjustment (type Leave return)."
        EndHoursBatch
        Exit Function
    End If
    If Int(returnDate) > Int(mYearTo) Then       ' v24.3.1 (review L2): logged before the rollover, the LV row would carry this year's average into the new year
        lc.Warnings = "The return date is after the last day of this roster's year (" & Format$(mYearTo, "mmm d, yyyy") & "). " & _
                      "Run the year-end rollover first, then log the return."
        EndHoursBatch
        Exit Function
    End If
    lc.WindowFrom = lc.FirstOff
    If lc.WindowFrom < mYearFrom Then lc.WindowFrom = mYearFrom: lc.Warnings = lc.Warnings & "Leave spans the year reset: the window starts at " & Format$(mYearFrom, "mmm d, yyyy") & ". "
    Dim recs() As EmpRec, n As Long, i As Long, nP As Long, deltas() As Double, names As String
    n = LoadEligible(home, "ALL", recs)
    ReDim deltas(1 To IIf(n > 0, n, 1))
    RefreshCalCache
    ' Hours History started Oct 9, 2026: a leave that began before the first
    ' snapshot has no peer with history before the window. Then every peer's
    ' start figure is estimated from the earliest snapshot, and the confirm says so.
    Dim relaxHistory As Boolean, firstSnap As Date, atReset As Boolean
    ' v23: a window that opens in the year's first pay period starts from the reset, where worked hours are 0 (as PickSnapshot does)
    atReset = (Int(lc.WindowFrom) < Int(mYearFrom) + 14)
    firstSnap = EarliestThrough()
    If Not atReset And firstSnap > 0 And lc.WindowFrom <= firstSnap Then
        relaxHistory = True
        lc.Warnings = lc.Warnings & "Hours History starts " & Format$(firstSnap, "mmm d") & ", after the first day off: every start figure is ESTIMATED from that snapshot. "
    End If
    Dim snapA As Date, snapB As Date, noteA As String, noteB As String, a As Double, b As Double, snaps As String, flagged As String
    For i = 1 To n
        If PeerEligible(recs(i).Grade, StrComp(recs(i).EmpName, emp, vbTextCompare) = 0, _
                        HasLeaveOverlapping(recs(i).EmpName, lc.WindowFrom, lc.ReturnDate), _
                        atReset Or relaxHistory Or HasHistoryBefore(recs(i).EmpName, lc.WindowFrom)) Then
            a = TotalHoursAsOf(recs(i).EmpName, lc.WindowFrom, snapA, noteA)
            b = TotalHoursAsOf(recs(i).EmpName, lc.ReturnDate, snapB, noteB)
            nP = nP + 1
            deltas(nP) = b - a
            names = names & IIf(names = "", "", ", ") & CalName(recs(i).EmpName)
            If snaps = "" Then
                snaps = "from " & IIf(snapA = 0, "the year reset (worked 0)", "the " & Format$(snapA, "mmm d") & " snapshot") & " to " & _
                        IIf(snapB = 0, "the year reset (worked 0)", "the " & Format$(snapB, "mmm d") & " snapshot") & " (first peer; others may differ)"
            End If
            If noteA <> "" And Not relaxHistory And InStr(flagged, "|start|") = 0 Then flagged = flagged & "|start|": lc.Warnings = lc.Warnings & "Start: " & noteA & ". "
            If noteB <> "" And InStr(flagged, "|end|") = 0 Then flagged = flagged & "|end|": lc.Warnings = lc.Warnings & "Return: " & noteB & ". "
        End If
    Next i
    lc.NPeers = nP
    lc.PeerNames = names
    lc.Average = LeaveAverage(deltas, nP)
    lc.SnapshotText = snaps
    Dim noteR As String, snapR As Date
    lc.Before = TotalHoursAsOf(emp, lc.ReturnDate, snapR, noteR)
    If noteR <> "" Then lc.Warnings = lc.Warnings & emp & ": " & noteR & ". "
    lc.After = lc.Before + lc.Average
    EndHoursBatch
    If nP = 0 Then lc.Warnings = lc.Warnings & "No peers found in " & home & " (grade " & PEER_MAX_GRADE & " and below, not on leave, with Hours History). "
    lc.Summary = home & ", " & nP & " operator" & IIf(nP = 1, "", "s") & ", " & Format$(lc.WindowFrom, "mmm d") & " to " & _
                 Format$(lc.ReturnDate, "mmm d") & ": average " & Format$(lc.Average, "0.0") & " h (" & snaps & "). " & _
                 "Add " & Format$(lc.Average, "0.0") & " h to " & emp & " (" & Format$(lc.Before, "0.0") & " -> " & Format$(lc.After, "0.0") & ")"
    ComputeLeaveReturn = True
    Exit Function
Failed:
    lc.Warnings = "The calculation stopped: " & Err.Description
    On Error Resume Next
    EndHoursBatch
End Function

' Writes the top-up (Adjustment, Ref LV-####) and closes the absence. hoursToAdd
' may be an admin override (then overrideReason is required and goes in the Note).
Public Function SaveLeaveReturn(ByRef lc As LeaveCalc, ByVal hoursToAdd As Double, ByVal overrideReason As String) As String
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim emp As String, unit As String, code As String, dStart As Date, dEnd As Variant, home As String, grp As String, st As String
    If Not AbsenceInfo(lc.AbsId, emp, unit, code, dStart, dEnd, home, grp, st) Then Exit Function
    If grp = "" Then grp = NextGroupId(AbsTable())
    Dim ps As Date: ps = PeriodStartFor(lc.ReturnDate)
    If Not ConfirmPeriod(ps) Then Exit Function
    Dim calc As String
    calc = "Leave return " & Format$(lc.FirstOff, "mmm d") & " -> " & Format$(lc.ReturnDate, "mmm d, yyyy") & ": " & lc.HomeUnit & _
           " average of " & lc.NPeers & " peer(s) [" & lc.PeerNames & "] " & Format$(lc.Average, "0.0") & " h " & lc.SnapshotText & _
           IIf(hoursToAdd <> lc.Average, " - OVERRIDDEN to " & Format$(hoursToAdd, "0.0") & " h: " & overrideReason, "")
    ' v25 (L5): the row goes in even at 0 h - it records that the return was logged (Needs Attention and
    ' the leave pop-up read LeaveToppedUp from the log), and it is what Void Entry undoes
    If hoursToAdd = 0 Then calc = calc & " - no hours added"
    Dim data() As Variant: ReDim data(1 To 1, 1 To L_NCOLS)
    data(1, L_REF) = grp
    data(1, L_TYPE) = "Adjustment"
    data(1, L_DATE) = lc.ReturnDate
    data(1, L_DEPT) = lc.HomeUnit
    data(1, L_TASK) = "Leave return"
    data(1, L_EMP) = emp
    data(1, L_EMPHRS) = lc.Before
    data(1, L_REASON) = calc
    data(1, L_CHARGED) = hoursToAdd
    data(1, L_PERIOD) = ps
    data(1, L_USER) = CurrentUser()
    data(1, L_STAMP) = Now
    If Not AppendLogRows(data) Then Exit Function
    SetAbsenceEndTo lc.AbsId, lc.ReturnDate - 1
    If grp <> "" Then SetAbsenceGroup lc.AbsId, grp
    MarkCalDirty
    SaveIfWritable
    SaveLeaveReturn = grp
End Function

Private Sub SetAbsenceGroup(ByVal absId As String, ByVal grp As String)
    Dim lo As ListObject: Set lo = AbsTable()
    If lo Is Nothing Then Exit Sub
    If lo.ListColumns.Count < AB_GROUP Then Exit Sub
    Dim v As Variant: v = RangeTo2D(lo.ListColumns(AB_ID).DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, 1)), absId, vbTextCompare) = 0 Then
            UnlockSheet lo.Parent
            lo.DataBodyRange.Cells(i, AB_GROUP).Value = grp
            ProtectSheet lo.Parent
            Exit Sub
        End If
    Next i
End Sub


' ====================================================================
'  SICK CALL
' ====================================================================
Public Sub ShowSickCallForm()
    On Error GoTo Oops
    If Not RequireSupervisor() Then Exit Sub     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not AbsenceColumnsCurrent() Then
        MsgBox "Run Alt+F8 > UpgradeAbsenceColumns first (the Absences table needs its v22 columns).", vbExclamation
        Exit Sub
    End If
    frmSickCall.Show
    Exit Sub
Oops:
    ReportError "ShowSickCallForm"
End Sub

' Books the sick shift (Absence ID SC-####, type SIC) and makes sure the
' calendar shows an OPEN line for it. Returns the SC id.
Public Function SaveSickCall(ByVal emp As String, ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, _
                             ByVal hrs As Variant, ByVal note As String) As String
    Dim id As String
    If HasSickRow(emp, d, shiftLbl) Then Exit Function   ' v25 (L7): never a second row for the same shift, whoever calls
    id = SaveAbsence(emp, unit, "SIC", d, d, hrs, "", note, HomeUnitOf(emp), "", shiftLbl, "SC", True)
    If id = "" Then Exit Function
    ' v24.3.1 (review C1): the absence row carries the shift (AB_SHIFT) and SlotsFor draws a sick
    ' call on that shift whether or not it is the scheduled one, so the extra coverage row this
    ' used to add for an OT shift drew the same sick call twice ("Name (Sick)" twice, both open).
    RefreshCalCache
    MarkCalDirty
    SaveIfWritable
    SaveSickCall = id
End Function


' ====================================================================
'  IMPORTED CALENDAR LINES -> REAL BOOKINGS
' ====================================================================
Public Sub ShowConvertForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not AbsenceColumnsCurrent() Then
        MsgBox "Run Alt+F8 > UpgradeAbsenceColumns first (the Absences table needs its v22 columns).", vbExclamation
        Exit Sub
    End If
    frmConvert.Show
    Exit Sub
Oops:
    ReportError "ShowConvertForm"
End Sub

' Imported vacation lines dated today or later that aren't booked yet.
' 2-D (rows x 6): line id, date, shift, unit, text as imported, matched roster name ("" / "?" ambiguous)
Public Function ImportCandidates(ByVal fromDate As Date, ByVal toDate As Date) As Variant
    Dim v As Variant: v = CovTableData()
    If Not IsArray(v) Then Exit Function
    Dim names As Variant: names = RosterNames()
    Dim out() As Variant: ReDim out(1 To UBound(v, 1), 1 To 6)
    Dim i As Long, n As Long, d As Date
    For i = 1 To UBound(v, 1)
        If Left$(SafeStr(v(i, CV_SOURCE)), 6) = "Import" And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" And _
           UCase$(SafeStr(v(i, CV_TYPE))) = "VAC" And SafeStr(v(i, CV_ABS)) = "" And IsDate(v(i, CV_DATE)) Then
            d = Int(CDate(v(i, CV_DATE)))
            If d >= Int(fromDate) And d <= Int(toDate) Then
                n = n + 1
                out(n, 1) = SafeStr(v(i, CV_ID))
                out(n, 2) = d
                out(n, 3) = SafeStr(v(i, CV_SHIFT))
                out(n, 4) = SafeStr(v(i, CV_UNIT))
                out(n, 5) = SafeStr(v(i, CV_NEED))
                out(n, 6) = MatchImportName(SafeStr(v(i, CV_NEED)), names)
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    Dim res() As Variant: ReDim res(1 To n, 1 To 6)
    Dim j As Long, c As Long
    For j = 1 To n
        For c = 1 To 6
            res(j, c) = out(j, c)
        Next c
    Next j
    ImportCandidates = res
End Function

Private Function CovTableData() As Variant
    Dim lo As ListObject: Set lo = CovTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    CovTableData = RangeTo2D(lo.DataBodyRange)
End Function

' "Paul M" (first name + last initial, as the old calendars wrote it) against
' roster names "Last, First": one match -> the roster name; several -> "?";
' none -> "". Extra words ("3rd", "8h", "vac") are ignored.
Public Function MatchImportName(ByVal txt As String, ByVal rosterList As Variant) As String
    If Not IsArray(rosterList) Then Exit Function
    Dim parts As Variant: parts = Split(Trim$(txt), " ")
    If UBound(parts) < 0 Then Exit Function
    Dim first As String, ini As String, i As Long
    first = UCase$(Trim$(CStr(parts(0))))
    If UBound(parts) >= 1 Then ini = UCase$(Left$(Trim$(CStr(parts(1))), 1))
    If first = "" Then Exit Function
    Dim hits As Long, hit As String, nm As String, p As Long, rFirst As String, rLast As String
    For i = LBound(rosterList) To UBound(rosterList)
        nm = CStr(rosterList(i))
        p = InStr(nm, ",")
        If p > 0 Then
            rLast = UCase$(Trim$(Left$(nm, p - 1)))
            rFirst = UCase$(Trim$(Mid$(nm, p + 1)))
        Else
            p = InStrRev(nm, " ")
            rFirst = UCase$(Trim$(Left$(nm, p)))
            rLast = UCase$(Trim$(Mid$(nm, p + 1)))
        End If
        If Left$(rFirst, Len(first)) = first Then
            If ini = "" Or Left$(rLast, 1) = ini Then hits = hits + 1: hit = nm
        End If
    Next i
    If hits = 1 Then
        MatchImportName = hit
    ElseIf hits > 1 Then
        MatchImportName = "?"
    End If
End Function

' Converts the given lines (ids, names decided) into VAC bookings: consecutive
' dates for one person merge into one booking; every line gets the Absence ID.
' Returns the number of bookings made.
Public Function ConvertImports(ByRef lineIds() As String, ByRef empNames() As String, ByVal n As Long, ByRef report As String) As Long
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim v As Variant: v = CovTableData()
    If Not IsArray(v) Or n = 0 Then Exit Function
    Dim lo As ListObject: Set lo = CovTable(False)
    ' group by person
    Dim i As Long, j As Long, done() As Boolean: ReDim done(1 To n)
    Dim made As Long
    For i = 1 To n
        If Not done(i) And empNames(i) <> "" Then
            Dim dates() As Date, ids() As String, k As Long, r As Long, shiftTxt As String
            ReDim dates(1 To n): ReDim ids(1 To n): k = 0: shiftTxt = ""
            For j = i To n
                If Not done(j) And StrComp(empNames(j), empNames(i), vbTextCompare) = 0 Then
                    r = RowOfId(v, lineIds(j))
                    If r > 0 Then
                        k = k + 1: dates(k) = Int(CDate(v(r, CV_DATE))): ids(k) = lineIds(j)
                        If InStr(shiftTxt, SafeStr(v(r, CV_SHIFT))) = 0 Then shiftTxt = shiftTxt & IIf(shiftTxt = "", "", "/") & SafeStr(v(r, CV_SHIFT))
                    End If
                    done(j) = True
                End If
            Next j
            If k > 0 Then
                SortDates dates, ids, k
                Dim runs() As Date, nr As Long, q As Long, absId As String, unit As String
                nr = MergeRuns(dates, k, runs)
                r = RowOfId(v, ids(1)): unit = SafeStr(v(r, CV_UNIT))
                For q = 1 To nr
                    absId = SaveAbsence(empNames(i), unit, "VAC", runs(q, 1), runs(q, 2), Empty, "", "Converted from the imported calendar", _
                                        HomeUnitOf(empNames(i)), "", IIf(InStr(shiftTxt, "/") > 0, "Both", shiftTxt), "AB", True, True)
                    If absId <> "" Then
                        made = made + 1
                        report = report & "  " & empNames(i) & ": " & Format$(runs(q, 1), "mmm d") & _
                                 IIf(runs(q, 2) <> runs(q, 1), " - " & Format$(runs(q, 2), "mmm d"), "") & " (" & absId & ")" & vbNewLine
                        For j = 1 To k
                            If dates(j) >= runs(q, 1) And dates(j) <= runs(q, 2) Then MarkLineBooked lo, v, ids(j), absId
                        Next j
                    Else
                        report = report & "  " & empNames(i) & ": " & Format$(runs(q, 1), "mmm d") & _
                                 IIf(runs(q, 2) <> runs(q, 1), " - " & Format$(runs(q, 2), "mmm d"), "") & _
                                 " SKIPPED - already has time off in that range (book or dismiss by hand)" & vbNewLine
                    End If
                Next q
            End If
        End If
    Next i
    If made > 0 Then
        ResetCalCacheAll
        MarkCalDirty
        SaveIfWritable
    End If
    ConvertImports = made
End Function

' "Dismiss": the line stays on the calendar but drops off the Needs Attention list
Public Sub DismissImports(ByRef lineIds() As String, ByVal n As Long)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim v As Variant: v = CovTableData()
    Dim lo As ListObject: Set lo = CovTable(False)
    If Not IsArray(v) Or lo Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To n
        MarkLineBooked lo, v, lineIds(i), "n/a"
    Next i
    MarkCalDirty
    SaveIfWritable
End Sub

Private Sub MarkLineBooked(ByVal lo As ListObject, ByRef v As Variant, ByVal lineId As String, ByVal absId As String)
    Dim r As Long: r = RowOfId(v, lineId)
    If r = 0 Then Exit Sub
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(r, CV_ABS).Value = absId
    ProtectSheet lo.Parent
End Sub

Private Function RowOfId(ByRef v As Variant, ByVal id As String) As Long
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), id, vbTextCompare) = 0 Then RowOfId = i: Exit Function
    Next i
End Function

Private Sub SortDates(ByRef dates() As Date, ByRef ids() As String, ByVal n As Long)
    Dim i As Long, j As Long, td As Date, ti As String
    For i = 2 To n
        td = dates(i): ti = ids(i): j = i - 1
        Do While j >= 1
            If dates(j) <= td Then Exit Do
            dates(j + 1) = dates(j): ids(j + 1) = ids(j): j = j - 1
        Loop
        dates(j + 1) = td: ids(j + 1) = ti
    Next i
End Sub

' Needs Attention: imported vacation lines in the next IMPORT_WATCH_DAYS days not yet booked
Public Function ImportedNotBookedCount() As Long
    Dim v As Variant: v = ImportCandidates(Date, Date + IMPORT_WATCH_DAYS - 1)
    If IsArray(v) Then ImportedNotBookedCount = UBound(v, 1)
End Function
