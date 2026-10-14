Attribute VB_Name = "modAttendance"
Option Explicit

' ====================================================================
'  ATTENDANCE (v22.1): sick and personal day counters
'
'  The plant rule (Graeme, Oct 12, 2026):
'   - 5 PAID sick days a calendar year, in up to 3 occurrences, 5 days in
'     all. A day is one scheduled 12-hour shift (a partial shift still
'     counts). An occurrence is a run of consecutive scheduled shifts
'     missed sick; days off between them don't break the run.
'   - 5 UNPAID personal days a calendar year.
'   - When sick + personal days pass 5, the person is not entitled to
'     paid sick days from that day until one year after the first day
'     of their 2nd occurrence (the only one, if just one; the crossing
'     day itself if none). On that anniversary the counters restart and
'     run to Dec 31; from the next Jan 1 the normal year resumes.
'
'  One rule function: AttendanceCore (pure, mirrored in tests/rules.py,
'  checked by SelfTest suite "attend"). Everything that shows a count -
'  the Sick Call pop-up, the Attendance sheet, Needs Attention - calls
'  AttendanceFor, which gathers the person's SIC and PER rows from the
'  Absences table and calls the core. Nothing here changes Total Hours,
'  call order, charging or the OT Log.
'
'  Privacy: counts and reasons are admin-only (the Attendance sheet is
'  hidden in saved copies and never published).
' ====================================================================

Public Const AT_VERSION As String = "Attendance v24.4 (Oct 14)"
Public Const AT_SHEET As String = "Attendance"

Public Const PAID_SICK_DAYS As Long = 5
Public Const PAID_SICK_OCC As Long = 3
Public Const PERSONAL_DAYS As Long = 5
Public Const COMBINED_LIMIT As Long = 5
Public Const MAX_OCC As Long = 12          ' occurrences kept apart per window; any beyond are folded into the last

Public Type AttendanceStatus
    WindowStart As Date        ' Jan 1, or the day the counters restarted after a wait
    SickUsed As Long           ' sick days in the window
    PaidUsed As Long           ' of which paid
    PaidLeft As Long
    Occurrences As Long
    OccDates As String         ' first day of each occurrence, "Jan 5, Mar 2"
    PersonalUsed As Long
    PersonalLeft As Long
    Combined As Long
    Entitled As Boolean        ' False while waiting for the anniversary
    WaitUntil As Date          ' when paid sick comes back (0 = not waiting)
    CrossedOn As Date          ' the day sick + personal passed the limit
    ResetOn As Date            ' the anniversary that restarted the counters (0 = none this window)
    LastSick As Date           ' latest sick shift on or before the as-of date
    LastPaid As Boolean
    LastReason As String       ' "paid sick day 2 of 5, occurrence 1 of 3" / "unpaid - ..."
    ' v24.1: each occurrence in the window (1 .. Occurrences, at most MAX_OCC)
    OccFirst(1 To MAX_OCC) As Date
    OccLast(1 To MAX_OCC) As Date
    OccDays(1 To MAX_OCC) As Long
    OccPaid(1 To MAX_OCC) As Long
End Type

Private Const EV_SICK As Long = 1
Private Const EV_PER As Long = 2

' running state of one AttendanceCore pass (module-level so the helpers can share it)
Private Type AtState
    NextJan1 As Date
    WindowStart As Date
    Waiting As Boolean
    WaitUntil As Date
    CrossedOn As Date
    ResetOn As Date
    Used As Long
    Paid As Long
    Occ As Long
    Occ1 As Date
    Occ2 As Date
    OccTxt As String
    Per As Long
    HaveLast As Boolean
    LastD As Date
    LastN As Boolean
    OccFirst(1 To MAX_OCC) As Date
    OccLast(1 To MAX_OCC) As Date
    OccDays(1 To MAX_OCC) As Long
    OccPaid(1 To MAX_OCC) As Long
End Type
Private mS As AtState

' the Attendance sheet's view (v24.1)
Private mAttSort As String          ' "" / "AZ" = A - Z, "DEPT" = by Home Unit, "WORST" = worst first
Private mAttAll As Boolean          ' True = everyone, False = only people with something logged

Private mAbsV As Variant          ' Absences table values while a batch is open
Private mBatch As Boolean


' ====================================================================
'  THE RULE
' ====================================================================

' sickD/sickN: each sick shift (date, night?); perD: each personal day
' (one entry per scheduled shift inside a personal booking). Events after
' asOf are ignored. Every earlier row is replayed from Jan 1 of its year:
' a wait set one year restarts the next year part-way, which changes the
' year after that, so no fixed look-back is safe (the table is small).
Public Sub AttendanceCore(ByVal crewCode As String, ByRef sickD() As Date, ByRef sickN() As Boolean, ByVal nSick As Long, _
                          ByRef perD() As Date, ByVal nPer As Long, ByVal asOf As Date, ByRef st As AttendanceStatus, _
                          Optional ByVal whoName As String = "")
    ' whoName (v24): when given, each day's crew comes from the crew history, with
    ' crewCode as the answer for someone who has none. Left out (SelfTest, the
    ' Python mirror): crewCode for every day.
    Dim blank As AttendanceStatus: st = blank
    Dim blankS As AtState: mS = blankS
    Dim n As Long: n = nSick + nPer
    Dim evD() As Date, evN() As Boolean, evK() As Long, i As Long, j As Long, m As Long
    ReDim evD(1 To n + 1): ReDim evN(1 To n + 1): ReDim evK(1 To n + 1)
    For i = 1 To nSick
        If Int(sickD(i)) <= Int(asOf) Then
            m = m + 1: evD(m) = Int(sickD(i)): evN(m) = sickN(i): evK(m) = EV_SICK
        End If
    Next i
    For i = 1 To nPer
        If Int(perD(i)) <= Int(asOf) Then
            m = m + 1: evD(m) = Int(perD(i)): evN(m) = False: evK(m) = EV_PER
        End If
    Next i
    ' insertion sort by shift position (days before nights, sick before personal on a tie)
    Dim tD As Date, tN As Boolean, tK As Long
    For i = 2 To m
        tD = evD(i): tN = evN(i): tK = evK(i): j = i - 1
        Do While j >= 1
            If EvPos(evD(j), evN(j), evK(j)) <= EvPos(tD, tN, tK) Then Exit Do
            evD(j + 1) = evD(j): evN(j + 1) = evN(j): evK(j + 1) = evK(j)
            j = j - 1
        Loop
        evD(j + 1) = tD: evN(j + 1) = tN: evK(j + 1) = tK
    Next i

    Dim y0 As Long: y0 = Year(asOf)
    If m > 0 Then
        If Year(evD(1)) < y0 Then y0 = Year(evD(1))
    End If
    mS.WindowStart = DateSerial(y0, 1, 1)
    mS.NextJan1 = DateSerial(y0 + 1, 1, 1)
    Dim paid As Boolean, why As String, sameOcc As Boolean, oc As Long
    For i = 1 To m
        If evD(i) >= mS.WindowStart Then
            AdvanceTo evD(i)
            If evK(i) = EV_SICK Then
                mS.Used = mS.Used + 1
                sameOcc = False
                If mS.HaveLast Then sameOcc = Not SchedBetween(crewCode, mS.LastD, mS.LastN, evD(i), evN(i), whoName)
                If Not sameOcc Then
                    mS.Occ = mS.Occ + 1
                    If mS.Occ = 1 Then mS.Occ1 = evD(i)
                    If mS.Occ = 2 Then mS.Occ2 = evD(i)
                    mS.OccTxt = mS.OccTxt & IIf(mS.OccTxt = "", "", ", ") & Format$(evD(i), "mmm d")
                End If
                oc = mS.Occ
                If oc > MAX_OCC Then oc = MAX_OCC
                If mS.OccDays(oc) = 0 Then mS.OccFirst(oc) = evD(i)
                mS.OccLast(oc) = evD(i)
                mS.OccDays(oc) = mS.OccDays(oc) + 1
                mS.HaveLast = True: mS.LastD = evD(i): mS.LastN = evN(i)
                If mS.Waiting Then
                    paid = False
                    why = "unpaid - not entitled until " & Format$(mS.WaitUntil, "mmm d, yyyy") & _
                          " (over " & COMBINED_LIMIT & " sick + personal days on " & Format$(mS.CrossedOn, "mmm d, yyyy") & ")"
                ElseIf mS.Occ > PAID_SICK_OCC Then
                    paid = False
                    why = "unpaid - occurrence " & mS.Occ & " (only " & PAID_SICK_OCC & " occurrences are paid)"
                ElseIf mS.Used > PAID_SICK_DAYS Then
                    paid = False
                    why = "unpaid - sick day " & mS.Used & " (only " & PAID_SICK_DAYS & " are paid)"
                Else
                    paid = True
                    why = "paid sick day " & mS.Used & " of " & PAID_SICK_DAYS & ", occurrence " & mS.Occ & " of " & PAID_SICK_OCC
                End If
                If Not mS.Waiting Then
                    If mS.Used + mS.Per > COMBINED_LIMIT Then
                        StartWait evD(i)
                        paid = False
                        why = "unpaid - sick + personal day " & (mS.Used + mS.Per) & " passes " & COMBINED_LIMIT & _
                              ": not entitled until " & Format$(mS.WaitUntil, "mmm d, yyyy")
                    End If
                End If
                If paid Then mS.Paid = mS.Paid + 1: mS.OccPaid(oc) = mS.OccPaid(oc) + 1
                st.LastSick = evD(i): st.LastPaid = paid: st.LastReason = why
            Else
                mS.Per = mS.Per + 1
                If Not mS.Waiting Then
                    If mS.Used + mS.Per > COMBINED_LIMIT Then StartWait evD(i)
                End If
            End If
        End If
    Next i
    AdvanceTo Int(asOf)

    st.WindowStart = mS.WindowStart
    st.SickUsed = mS.Used
    st.PaidUsed = mS.Paid
    st.PaidLeft = PAID_SICK_DAYS - mS.Used              ' days still payable inside the current occurrences
    If st.PaidLeft < 0 Then st.PaidLeft = 0
    ' v24.3.1 (review L1, decided D-L4): nothing is payable once a 4th occurrence has started, or once
    ' the 3rd has ENDED - a scheduled shift of theirs after its last sick shift and before the as-of
    ' date's NIGHT (the far end of the strictly-between test is the night of as-of: a day shift on
    ' as-of counts as worked, the night on as-of does not, since it starts at 7 pm and a sick call
    ' for it may still come). The status phrase "Paid sick days used up" follows from PaidLeft = 0;
    ' "only while the current occurrence continues" is now said only while it really does.
    If mS.Occ > PAID_SICK_OCC Then
        st.PaidLeft = 0
    ElseIf mS.Occ = PAID_SICK_OCC And mS.HaveLast Then
        If SchedBetween(crewCode, mS.LastD, mS.LastN, Int(asOf), True, whoName) Then st.PaidLeft = 0
    End If
    st.Occurrences = mS.Occ
    st.OccDates = mS.OccTxt
    st.PersonalUsed = mS.Per
    st.PersonalLeft = PERSONAL_DAYS - mS.Per
    If st.PersonalLeft < 0 Then st.PersonalLeft = 0
    st.Combined = mS.Used + mS.Per
    st.Entitled = Not mS.Waiting
    If mS.Waiting Then st.PaidLeft = 0
    st.WaitUntil = mS.WaitUntil
    st.CrossedOn = mS.CrossedOn
    st.ResetOn = mS.ResetOn
    For oc = 1 To MAX_OCC
        st.OccFirst(oc) = mS.OccFirst(oc): st.OccLast(oc) = mS.OccLast(oc)
        st.OccDays(oc) = mS.OccDays(oc): st.OccPaid(oc) = mS.OccPaid(oc)
    Next oc
End Sub

' One occurrence as the Attendance sheet shows it: "Jan 5 - 6 (2 days, 2 paid)",
' "Mar 2 (1 day, unpaid)", "Jan 30 - Feb 2 (3 days, 1 paid)". "" when there is none.
Public Function OccurrenceText(ByRef st As AttendanceStatus, ByVal k As Long) As String
    If k < 1 Or k > MAX_OCC Then Exit Function
    If st.OccDays(k) = 0 Then Exit Function
    Dim s As String
    If st.OccLast(k) = st.OccFirst(k) Then
        s = Format$(st.OccFirst(k), "mmm d")
    ElseIf Month(st.OccLast(k)) = Month(st.OccFirst(k)) And Year(st.OccLast(k)) = Year(st.OccFirst(k)) Then
        s = Format$(st.OccFirst(k), "mmm d") & " - " & Day(st.OccLast(k))
    Else
        s = Format$(st.OccFirst(k), "mmm d") & " - " & Format$(st.OccLast(k), "mmm d")
    End If
    s = s & " (" & st.OccDays(k) & IIf(st.OccDays(k) = 1, " day, ", " days, ")
    If st.OccPaid(k) = 0 Then
        s = s & "unpaid)"
    ElseIf st.OccDays(k) = 1 Then
        s = s & "paid)"
    Else
        s = s & st.OccPaid(k) & " paid)"
    End If
    OccurrenceText = s
End Function

' Sort position of an event: days before nights on a date, sick before personal
Private Function EvPos(ByVal d As Date, ByVal night As Boolean, ByVal kind As Long) As Double
    EvPos = CDbl(Int(d)) * 4 + IIf(night, 2, 0) + (kind - 1)
End Function

' Applies every Jan 1 and every wait-anniversary on or before d, in date order
Private Sub AdvanceTo(ByVal d As Date)
    Dim nxt As Date, kind As Long
    Do
        nxt = 0: kind = 0
        If mS.NextJan1 <= d Then nxt = mS.NextJan1: kind = 1
        If mS.Waiting Then
            If mS.WaitUntil <= d Then
                If kind = 0 Or mS.WaitUntil < nxt Then nxt = mS.WaitUntil: kind = 2
            End If
        End If
        If kind = 0 Then Exit Do
        mS.Used = 0: mS.Paid = 0: mS.Occ = 0: mS.Occ1 = 0: mS.Occ2 = 0: mS.OccTxt = "": mS.Per = 0: mS.HaveLast = False
        Dim oc As Long
        For oc = 1 To MAX_OCC
            mS.OccFirst(oc) = 0: mS.OccLast(oc) = 0: mS.OccDays(oc) = 0: mS.OccPaid(oc) = 0
        Next oc
        mS.WindowStart = nxt
        If kind = 1 Then
            mS.NextJan1 = DateSerial(Year(nxt) + 1, 1, 1)
            mS.ResetOn = 0
        Else
            mS.Waiting = False
            mS.ResetOn = mS.WaitUntil
            mS.WaitUntil = 0
            mS.CrossedOn = 0
        End If
    Loop
End Sub

' Sick + personal passed the limit on d
Private Sub StartWait(ByVal d As Date)
    mS.Waiting = True
    mS.CrossedOn = d
    If mS.Occ >= 2 Then
        mS.WaitUntil = AddYear(mS.Occ2)
    ElseIf mS.Occ >= 1 Then
        mS.WaitUntil = AddYear(mS.Occ1)
    Else
        mS.WaitUntil = AddYear(d)
    End If
End Sub

Public Function AddYear(ByVal d As Date) As Date
    AddYear = DateAdd("yyyy", 1, d)            ' Feb 29 -> Feb 28
End Function

' True when a scheduled shift of the crew lies strictly between two shifts
Public Function SchedBetween(ByVal crewCode As String, ByVal d1 As Date, ByVal n1 As Boolean, _
                             ByVal d2 As Date, ByVal n2 As Boolean, Optional ByVal whoName As String = "") As Boolean
    Dim p1 As Double, p2 As Double, p As Double, x As Date, s As String
    p1 = EvPos(d1, n1, 1): p2 = EvPos(d2, n2, 1)
    If p2 <= p1 Then Exit Function
    If Int(d2) - Int(d1) > 60 Then SchedBetween = True: Exit Function
    For x = Int(d1) To Int(d2)
        If whoName <> "" Then
            s = CrewShiftOn(CrewOn(whoName, x, crewCode), x)       ' v24: the crew they were on that day
        Else
            s = CrewShiftOn(crewCode, x)
        End If
        If s = "?" Then s = "Days"
        If s <> "" Then
            p = EvPos(x, (s = "Nights"), 1)
            If p > p1 And p < p2 Then SchedBetween = True: Exit Function
        End If
    Next x
End Function


' ====================================================================
'  LIVE DATA
' ====================================================================
Public Sub BeginAttendanceBatch()
    LoadAbs
    mBatch = True
End Sub

Public Sub EndAttendanceBatch()
    mBatch = False
    mAbsV = Empty
End Sub

Private Sub LoadAbs()
    mAbsV = AbsTableData()               ' Empty when there is no Absences table yet
End Sub

' Status of emp as of asOf. addDate / addShift add one more sick shift
' first (the Sick Call pop-up previews the shift being logged).
Public Function AttendanceFor(ByVal emp As String, ByVal asOf As Date, ByRef st As AttendanceStatus, _
                              Optional ByVal addDate As Date = 0, Optional ByVal addShift As String = "") As Boolean
    If Not mBatch Then LoadAbs
    Dim sd() As Date, sn() As Boolean, pd() As Date, nS As Long, nP As Long
    ReDim sd(1 To 1): ReDim sn(1 To 1): ReDim pd(1 To 1)
    Dim r As Long, code As String, e1 As Date, e2 As Date, x As Date, hasShiftCol As Boolean
    If IsArray(mAbsV) Then
        hasShiftCol = (UBound(mAbsV, 2) >= AB_SHIFT)
        For r = 1 To UBound(mAbsV, 1)
            If StrComp(Trim$(SafeStr(mAbsV(r, AB_EMP))), Trim$(emp), vbTextCompare) = 0 And _
               UCase$(SafeStr(mAbsV(r, AB_STATUS))) <> "VOID" And IsDate(mAbsV(r, AB_START)) Then
                code = UCase$(SafeStr(mAbsV(r, AB_TYPE)))
                If code = "SIC" Then
                    nS = nS + 1
                    If nS > UBound(sd) Then ReDim Preserve sd(1 To nS + 20): ReDim Preserve sn(1 To nS + 20)
                    sd(nS) = Int(CDate(mAbsV(r, AB_START)))
                    If hasShiftCol Then sn(nS) = IsNightShift(SafeStr(mAbsV(r, AB_SHIFT))) Else sn(nS) = False
                    If addDate > 0 Then                      ' already saved: preview it as is (no Shift column = any shift that day)
                        If sd(nS) = Int(addDate) And (Not hasShiftCol Or sn(nS) = IsNightShift(addShift)) Then addDate = 0
                    End If
                ElseIf code = "PER" Then
                    e1 = Int(CDate(mAbsV(r, AB_START)))
                    If IsDate(mAbsV(r, AB_END)) Then e2 = Int(CDate(mAbsV(r, AB_END))) Else e2 = Int(asOf)
                    If e2 > Int(asOf) Then e2 = Int(asOf)
                    If e2 - e1 > 400 Then e2 = e1 + 400
                    For x = e1 To e2
                        If EmpShiftOn(emp, x) <> "" Then
                            nP = nP + 1
                            If nP > UBound(pd) Then ReDim Preserve pd(1 To nP + 20)
                            pd(nP) = x
                        End If
                    Next x
                End If
            End If
        Next r
    End If
    If addDate > 0 Then
        nS = nS + 1
        If nS > UBound(sd) Then ReDim Preserve sd(1 To nS + 1): ReDim Preserve sn(1 To nS + 1)
        sd(nS) = Int(addDate): sn(nS) = IsNightShift(addShift)
    End If
    AttendanceCore EmpCrewCode(emp), sd, sn, nS, pd, nP, asOf, st, emp
    AttendanceFor = True
End Function

' 0 = fine, 1 = at a limit (5 sick, 3 occurrences, 5 personal or 5 combined), 2 = over the limit / waiting
Public Function AttendanceLevel(ByRef st As AttendanceStatus) As Long
    If Not st.Entitled Then
        AttendanceLevel = 2
    ElseIf st.SickUsed >= PAID_SICK_DAYS Or st.Occurrences >= PAID_SICK_OCC Or _
           st.PersonalUsed >= PERSONAL_DAYS Or st.Combined >= COMBINED_LIMIT Then
        AttendanceLevel = 1
    End If
End Function

Public Function AttendanceStatusText(ByRef st As AttendanceStatus) As String
    If Not st.Entitled Then
        AttendanceStatusText = "NOT ENTITLED to paid sick"
    ElseIf AttendanceLevel(st) = 1 Then
        AttendanceStatusText = "At the limit"
    Else
        AttendanceStatusText = "OK"
    End If
End Function

' One line for the Sick Call pop-up and the Attendance sheet
Public Function AttendanceLine(ByRef st As AttendanceStatus) As String
    Dim s As String
    If Month(st.WindowStart) = 1 And Day(st.WindowStart) = 1 Then
        s = Format$(st.WindowStart, "yyyy") & ": "
    Else
        s = "Since " & Format$(st.WindowStart, "mmm d, yyyy") & ": "
    End If
    s = s & st.SickUsed & " sick day" & IIf(st.SickUsed = 1, "", "s") & " (" & st.PaidUsed & " paid of " & PAID_SICK_DAYS & ") in " & _
        st.Occurrences & " of " & PAID_SICK_OCC & " occurrences; personal " & st.PersonalUsed & " of " & PERSONAL_DAYS & "."
    If Not st.Entitled Then
        s = s & " NOT entitled to paid sick until " & Format$(st.WaitUntil, "mmm d, yyyy") & _
            " (over " & COMBINED_LIMIT & " on " & Format$(st.CrossedOn, "mmm d, yyyy") & ")."
    End If
    AttendanceLine = s
End Function

' Text for the Sick Call pop-up: the counters, then what this shift would be
Public Function SickCallPreview(ByVal emp As String, ByVal d As Date, ByVal shiftLbl As String) As String
    Dim st As AttendanceStatus
    If Not AttendanceFor(emp, d, st, d, shiftLbl) Then Exit Function
    SickCallPreview = AttendanceLine(st) & vbLf & "This shift: " & UCase$(Left$(st.LastReason, 1)) & Mid$(st.LastReason, 2)
End Function

' The outcome the rule gives the sick shift being saved (goes in the absence Note)
Public Function SickOutcome(ByVal emp As String, ByVal d As Date, ByVal shiftLbl As String) As String
    Dim st As AttendanceStatus
    If AttendanceFor(emp, d, st, d, shiftLbl) Then SickOutcome = UCase$(Left$(st.LastReason, 1)) & Mid$(st.LastReason, 2)
End Function


' ====================================================================
'  PAST SICK DAYS (Book Time Off > "SIC - Sick day (past)")
'  One SC row per scheduled shift in the range, oldest first so each
'  row's note says what the rule made of it. No OPEN line, no call-out.
' ====================================================================
' v25 (L6): nUnsched / nAlready get the days skipped (no scheduled shift / already booked sick), so the
' pop-up can say why a range made fewer rows than days.
Public Function SaveSickPast(ByVal emp As String, ByVal unit As String, ByVal d1 As Date, ByVal d2 As Date, _
                             ByVal note As String, Optional ByVal hrs As Variant = Empty, _
                             Optional ByRef nUnsched As Long, Optional ByRef nAlready As Long) As Long
    nUnsched = 0: nAlready = 0
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If d2 < d1 Then Exit Function
    If d1 < RosterYearFrom() Then Exit Function       ' v25: never a sick row before this roster year (the pop-up says why)
    If Not AbsenceColumnsCurrent() Then
        MsgBox "The Absences table needs its v22 columns first: run Alt+F8 > UpgradeAbsenceColumns.", vbExclamation
        Exit Function
    End If
    If d2 - d1 > 120 Then d2 = d1 + 120
    Dim x As Date, sh As String, id As String, made As Long, outcome As String
    For x = Int(d1) To Int(d2)
        sh = EmpShiftOn(emp, x)
        If sh = "" Then
            nUnsched = nUnsched + 1
        ElseIf HasSickRow(emp, x, sh) Then                        ' never a second row for the same shift
            nAlready = nAlready + 1
        Else
            outcome = SickOutcome(emp, x, sh)
            id = SaveAbsence(emp, unit, "SIC", x, x, hrs, "", "Past sick day" & IIf(note = "", "", " - " & note) & _
                             IIf(outcome = "", "", " - " & outcome), HomeUnitOf(emp), "", sh, "SC", True)
            If id <> "" Then made = made + 1
        End If
    Next x
    If made > 0 Then
        RefreshCalCache
        MarkCalDirty
        SaveIfWritable
    End If
    SaveSickPast = made
End Function


' True when emp already has a live SIC row for that date and shift
Public Function HasSickRow(ByVal emp As String, ByVal d As Date, ByVal shiftLbl As String) As Boolean
    Dim v As Variant: v = AbsTableData()
    If Not IsArray(v) Then Exit Function
    Dim r As Long, hasShiftCol As Boolean: hasShiftCol = (UBound(v, 2) >= AB_SHIFT)
    For r = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(r, AB_TYPE))) = "SIC" And UCase$(SafeStr(v(r, AB_STATUS))) <> "VOID" And IsDate(v(r, AB_START)) Then
            If StrComp(Trim$(SafeStr(v(r, AB_EMP))), Trim$(emp), vbTextCompare) = 0 And Int(CDate(v(r, AB_START))) = Int(d) Then
                If Not hasShiftCol Then HasSickRow = True: Exit Function
                If IsNightShift(SafeStr(v(r, AB_SHIFT))) = IsNightShift(shiftLbl) Then HasSickRow = True: Exit Function
            End If
        End If
    Next r
End Function


' ====================================================================
'  ATTENDANCE SHEET (admin, drawn, hidden in saved copies)
' ====================================================================
Private Function AttendanceSheet(ByVal create As Boolean) As Worksheet
    If SheetExists(AT_SHEET) Then
        Set AttendanceSheet = ThisWorkbook.Worksheets(AT_SHEET)
    ElseIf create Then
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = AT_SHEET
        RelockStructure
        ws.Tab.Color = RGB(230, 81, 0)
        Set AttendanceSheet = ws
    End If
End Function

' The sort buttons on the sheet
Public Sub AttendanceSortAZ()
    On Error GoTo Oops
    mAttSort = "AZ"
    ShowAttendance
    Exit Sub
Oops:
    ReportError "AttendanceSortAZ"
End Sub

Public Sub AttendanceSortDept()
    On Error GoTo Oops
    mAttSort = "DEPT"
    ShowAttendance
    Exit Sub
Oops:
    ReportError "AttendanceSortDept"
End Sub

Public Sub AttendanceSortWorst()
    On Error GoTo Oops
    mAttSort = "WORST"
    ShowAttendance
    Exit Sub
Oops:
    ReportError "AttendanceSortWorst"
End Sub

Public Sub AttendanceToggleAll()
    On Error GoTo Oops
    mAttAll = Not mAttAll
    ShowAttendance
    Exit Sub
Oops:
    ReportError "AttendanceToggleAll"
End Sub

Private Sub AttBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                   ByVal l As Double, ByVal w As Double, ByVal isOn As Boolean)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, ws.Rows(3).Top + 3, w, 20)
    s.Name = nm
    s.Fill.ForeColor.RGB = IIf(isOn, RGB(230, 81, 0), RGB(236, 239, 243))
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.OnAction = mac
    With s.TextFrame2
        .MarginLeft = 1: .MarginRight = 1: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = cap
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = "Segoe UI"
            .Font.Size = 8.5
            .Font.Bold = IIf(isOn, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = IIf(isOn, vbWhite, RGB(51, 51, 51))
        End With
    End With
End Sub

' One row per person (v24.1): Employee, Crew, Home Unit, one column per
' occurrence (its dates, days and paid days), then the counters and status.
' Default: A - Z, only people with a sick or personal day in their current
' window (or waiting for paid sick to come back). The buttons in row 3 sort
' by Home Unit or worst first, and show everyone. A filter sits on the header.
Public Sub ShowAttendance()
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Failed
    Set ws = AttendanceSheet(True)
    If ws Is Nothing Then Err.Raise vbObjectError + 701, , "the sheet couldn't be created (is the workbook structure protected?)"
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    Dim i As Long, j As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    If mAttSort = "" Then mAttSort = "AZ"

    ' ---- work everyone out ----
    Dim names As Variant: names = RosterNames()
    Dim nAll As Long, n As Long, st As AttendanceStatus, nOcc As Long, lvl As Long
    If IsArray(names) Then nAll = UBound(names) - LBound(names) + 1
    Dim sts() As AttendanceStatus, who() As String, crewOf() As String, unitOf() As String, sortKey() As String, ord() As Long
    ReDim sts(1 To nAll + 1): ReDim who(1 To nAll + 1): ReDim crewOf(1 To nAll + 1): ReDim unitOf(1 To nAll + 1)
    ReDim sortKey(1 To nAll + 1): ReDim ord(1 To nAll + 1)
    nOcc = PAID_SICK_OCC
    If nAll > 0 Then
        BeginAttendanceBatch
        For i = 1 To nAll
            AttendanceFor CStr(names(LBound(names) + i - 1)), Date, st
            If mAttAll Or st.SickUsed + st.PersonalUsed > 0 Or Not st.Entitled Then
                n = n + 1
                sts(n) = st
                who(n) = CStr(names(LBound(names) + i - 1))
                crewOf(n) = EmpCrewCode(who(n), Date)             ' today's crew (column E can hold one that starts later)
                unitOf(n) = HomeUnitOf(who(n))
                If st.Occurrences > nOcc Then nOcc = st.Occurrences
                lvl = AttendanceLevel(st)
                Select Case mAttSort
                    Case "DEPT": sortKey(n) = LCase$(IIf(unitOf(n) = "", "~", unitOf(n))) & "|" & LCase$(who(n))
                    Case "WORST": sortKey(n) = CStr(9 - lvl) & Format$(999 - st.Combined, "000") & LCase$(who(n))
                    Case Else: sortKey(n) = LCase$(who(n))
                End Select
                ord(n) = n
            End If
        Next i
        EndAttendanceBatch
    End If
    If nOcc > MAX_OCC Then nOcc = MAX_OCC
    Dim t As Long
    For i = 2 To n
        t = ord(i): j = i - 1
        Do While j >= 1
            If sortKey(ord(j)) <= sortKey(t) Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = t
    Next i

    ' ---- title, buttons, header ----
    Dim nc As Long: nc = 3 + nOcc + 6
    Dim cSick As Long: cSick = 3 + nOcc + 1
    ws.Range("A1").Value = "ATTENDANCE  -  sick and personal days (calendar year; " & PAID_SICK_DAYS & " paid sick in " & _
                           PAID_SICK_OCC & " occurrences, " & PERSONAL_DAYS & " personal; over " & COMBINED_LIMIT & _
                           " = no paid sick until a year after the 2nd occurrence)"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = IIf(mAttAll, "Everyone on the roster", "People with a sick or personal day in their current window") & _
                           " (" & n & " of " & nAll & ").  Occurrences are for the current window: they restart on Jan 1 and when paid sick comes back.  " & _
                           "Drawn " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & ". Admin only - never published."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    ws.Rows(3).RowHeight = 26
    AttBtn ws, "at_az", "A - Z", "modAttendance.AttendanceSortAZ", 4, 60, (mAttSort = "AZ")
    AttBtn ws, "at_dept", "By department", "modAttendance.AttendanceSortDept", 68, 96, (mAttSort = "DEPT")
    AttBtn ws, "at_worst", "Worst first", "modAttendance.AttendanceSortWorst", 168, 80, (mAttSort = "WORST")
    AttBtn ws, "at_all", IIf(mAttAll, "Only with absences", "Show everyone"), "modAttendance.AttendanceToggleAll", 262, 116, False
    AttBtn ws, "at_redraw", "Redraw", "modAttendance.ShowAttendance", 382, 60, False

    Dim heads() As Variant: ReDim heads(1 To nc)
    heads(1) = "Employee": heads(2) = "Crew": heads(3) = "Home Unit"
    For j = 1 To nOcc
        heads(3 + j) = "Occurrence " & j
    Next j
    heads(cSick) = "Sick days": heads(cSick + 1) = "Paid": heads(cSick + 2) = "Personal"
    heads(cSick + 3) = "Sick + personal": heads(cSick + 4) = "Status": heads(cSick + 5) = "Until / restarted"
    ws.Range("A4").Resize(1, nc).Value = heads
    With ws.Range("A4").Resize(1, nc)
        .Font.Bold = True: .Interior.Color = RGB(230, 81, 0): .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(4).RowHeight = 20

    ' ---- rows ----
    If n > 0 Then
        Dim out() As Variant: ReDim out(1 To n, 1 To nc)
        Dim k As Long, band As Boolean, lastUnit As String
        For i = 1 To n
            k = ord(i)
            out(i, 1) = who(k): out(i, 2) = crewOf(k): out(i, 3) = unitOf(k)
            For j = 1 To nOcc
                out(i, 3 + j) = OccurrenceText(sts(k), j)
            Next j
            out(i, cSick) = sts(k).SickUsed
            out(i, cSick + 1) = sts(k).PaidUsed
            out(i, cSick + 2) = sts(k).PersonalUsed
            out(i, cSick + 3) = sts(k).Combined
            out(i, cSick + 4) = AttendanceStatusText(sts(k))
            If sts(k).WaitUntil > 0 Then
                out(i, cSick + 5) = "until " & Format$(sts(k).WaitUntil, "mmm d, yyyy")
            ElseIf sts(k).ResetOn > 0 Then
                out(i, cSick + 5) = "restarted " & Format$(sts(k).ResetOn, "mmm d, yyyy")
            Else
                out(i, cSick + 5) = ""
            End If
        Next i
        ws.Range("A5").Resize(n, 3 + nOcc).NumberFormat = "@"
        ws.Range(ws.Cells(5, cSick + 4), ws.Cells(4 + n, nc)).NumberFormat = "@"
        ws.Range("A5").Resize(n, nc).Value = out
        For i = 1 To n
            Select Case Left$(CStr(out(i, cSick + 4)), 3)
                Case "NOT": ws.Cells(4 + i, cSick + 4).Interior.Color = RGB(255, 205, 210)
                Case "At ": ws.Cells(4 + i, cSick + 4).Interior.Color = RGB(255, 236, 179)
            End Select
            If mAttSort = "DEPT" Then              ' a light band every other unit (no heading rows: the filter keeps working)
                If i = 1 Or CStr(out(i, 3)) <> lastUnit Then band = Not band
                lastUnit = CStr(out(i, 3))
                If band Then ws.Range(ws.Cells(4 + i, 1), ws.Cells(4 + i, 3)).Interior.Color = RGB(245, 247, 250)
            End If
        Next i
        ws.Range("A4").Resize(n + 1, nc).AutoFilter
    Else
        ws.Range("A5").Value = IIf(nAll = 0, "No names on the roster.", "Nobody has a sick or personal day logged. Click Show everyone for the full list.")
    End If
    ws.Columns(1).ColumnWidth = 26
    ws.Columns(2).ColumnWidth = 7
    ws.Columns(3).ColumnWidth = 16
    For j = 1 To nOcc
        ws.Columns(3 + j).ColumnWidth = 27
    Next j
    ws.Range(ws.Columns(cSick), ws.Columns(cSick + 3)).ColumnWidth = 10
    ws.Columns(cSick + 4).ColumnWidth = 24
    ws.Columns(cSick + 5).ColumnWidth = 22
    ws.Range(ws.Cells(5, cSick), ws.Cells(4 + IIf(n > 0, n, 1), cSick + 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(5, 2), ws.Cells(4 + IIf(n > 0, n, 1), 2)).HorizontalAlignment = xlCenter
    ws.Range("A4").Resize(IIf(n > 0, n, 1) + 1, nc).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(4 + IIf(n > 0, n, 1), nc)).Address
        .PrintTitleRows = "$4:$4"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo Failed
    ProtectSheet ws
    ws.Activate
    ReleaseStatusBar                        ' v24.1: the roster's status-bar line must not sit over this tab
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("B5").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    LogError "ShowAttendance", Err.Number, em
    On Error Resume Next
    EndAttendanceBatch                          ' never leave the stale snapshot behind
    If Not ws Is Nothing Then ProtectSheet ws
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    MsgBox "The Attendance sheet couldn't be drawn: " & em, vbExclamation
End Sub
