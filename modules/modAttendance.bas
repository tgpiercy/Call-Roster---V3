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

Public Const AT_VERSION As String = "Attendance v22.1 (Oct 14)"
Public Const AT_SHEET As String = "Attendance"

Public Const PAID_SICK_DAYS As Long = 5
Public Const PAID_SICK_OCC As Long = 3
Public Const PERSONAL_DAYS As Long = 5
Public Const COMBINED_LIMIT As Long = 5

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
End Type
Private mS As AtState

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
                          ByRef perD() As Date, ByVal nPer As Long, ByVal asOf As Date, ByRef st As AttendanceStatus)
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
    Dim paid As Boolean, why As String, sameOcc As Boolean
    For i = 1 To m
        If evD(i) >= mS.WindowStart Then
            AdvanceTo evD(i)
            If evK(i) = EV_SICK Then
                mS.Used = mS.Used + 1
                sameOcc = False
                If mS.HaveLast Then sameOcc = Not SchedBetween(crewCode, mS.LastD, mS.LastN, evD(i), evN(i))
                If Not sameOcc Then
                    mS.Occ = mS.Occ + 1
                    If mS.Occ = 1 Then mS.Occ1 = evD(i)
                    If mS.Occ = 2 Then mS.Occ2 = evD(i)
                    mS.OccTxt = mS.OccTxt & IIf(mS.OccTxt = "", "", ", ") & Format$(evD(i), "mmm d")
                End If
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
                If paid Then mS.Paid = mS.Paid + 1
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
End Sub

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
                             ByVal d2 As Date, ByVal n2 As Boolean) As Boolean
    Dim p1 As Double, p2 As Double, p As Double, x As Date, s As String
    p1 = EvPos(d1, n1, 1): p2 = EvPos(d2, n2, 1)
    If p2 <= p1 Then Exit Function
    If Int(d2) - Int(d1) > 60 Then SchedBetween = True: Exit Function
    For x = Int(d1) To Int(d2)
        s = CrewShiftOn(crewCode, x)
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
    AttendanceCore EmpCrewCode(emp), sd, sn, nS, pd, nP, asOf, st
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
Public Function SaveSickPast(ByVal emp As String, ByVal unit As String, ByVal d1 As Date, ByVal d2 As Date, _
                             ByVal note As String, Optional ByVal hrs As Variant = Empty) As Long
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If d2 < d1 Then Exit Function
    If Not AbsenceColumnsCurrent() Then
        MsgBox "The Absences table needs its v22 columns first: run Alt+F8 > UpgradeAbsenceColumns.", vbExclamation
        Exit Function
    End If
    If d2 - d1 > 120 Then d2 = d1 + 120
    Dim x As Date, sh As String, id As String, made As Long, outcome As String
    For x = Int(d1) To Int(d2)
        sh = EmpShiftOn(emp, x)
        If sh <> "" And Not HasSickRow(emp, x, sh) Then          ' never a second row for the same shift
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
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = AT_SHEET
        ws.Tab.Color = RGB(230, 81, 0)
        Set AttendanceSheet = ws
    End If
End Function

Public Sub ShowAttendance()
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Failed
    Set ws = AttendanceSheet(True)
    If ws Is Nothing Then Err.Raise vbObjectError + 701, , "the sheet couldn't be created (is the workbook structure protected?)"
    ws.Visible = xlSheetVisible
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells.Clear
    Dim names As Variant: names = RosterNames()
    Dim heads As Variant
    heads = Array("Employee", "Crew", "Home Unit", "Sick days", "Paid", "Occurrences", "First days", "Personal", _
                  "Sick + personal", "Status", "Until / restarted", "Last sick", "Window")
    Const NC As Long = 13
    ws.Range("A1").Value = "ATTENDANCE  -  sick and personal days (calendar year; " & PAID_SICK_DAYS & " paid sick in " & _
                           PAID_SICK_OCC & " occurrences, " & PERSONAL_DAYS & " personal; over " & COMBINED_LIMIT & _
                           " = no paid sick until a year after the 2nd occurrence)"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = "Drawn " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & " from the Absences table. Admin only - never published. Click Attendance again to redraw."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    ws.Range("A4").Resize(1, NC).Value = heads
    With ws.Range("A4").Resize(1, NC)
        .Font.Bold = True: .Interior.Color = RGB(230, 81, 0): .Font.Color = vbWhite
    End With
    Dim n As Long, i As Long, st As AttendanceStatus, lvl As Long
    If IsArray(names) Then n = UBound(names) - LBound(names) + 1
    If n > 0 Then
        Dim v() As Variant, k() As Double, ord() As Long
        ReDim v(1 To n, 1 To NC): ReDim k(1 To n): ReDim ord(1 To n)
        BeginAttendanceBatch
        For i = 1 To n
            AttendanceFor CStr(names(LBound(names) + i - 1)), Date, st
            lvl = AttendanceLevel(st)
            v(i, 1) = CStr(names(LBound(names) + i - 1))
            v(i, 2) = EmpCrewCode(v(i, 1))
            v(i, 3) = HomeUnitOf(v(i, 1))
            v(i, 4) = st.SickUsed
            v(i, 5) = st.PaidUsed
            v(i, 6) = st.Occurrences
            v(i, 7) = st.OccDates
            v(i, 8) = st.PersonalUsed
            v(i, 9) = st.Combined
            v(i, 10) = AttendanceStatusText(st)
            If st.WaitUntil > 0 Then
                v(i, 11) = "until " & Format$(st.WaitUntil, "mmm d, yyyy")
            ElseIf st.ResetOn > 0 Then
                v(i, 11) = "restarted " & Format$(st.ResetOn, "mmm d, yyyy")
            Else
                v(i, 11) = ""
            End If
            If st.LastSick > 0 Then v(i, 12) = Format$(st.LastSick, "mmm d, yyyy") Else v(i, 12) = ""
            v(i, 13) = IIf(Month(st.WindowStart) = 1 And Day(st.WindowStart) = 1, Format$(st.WindowStart, "yyyy"), "from " & Format$(st.WindowStart, "mmm d, yyyy"))
            k(i) = -lvl * 1000 - st.Combined        ' worst first, then most days
            ord(i) = i
        Next i
        EndAttendanceBatch
        ' sort the index by k, then name (insertion sort)
        Dim j As Long, t As Long
        For i = 2 To n
            t = ord(i): j = i - 1
            Do While j >= 1
                If k(ord(j)) < k(t) Then Exit Do
                If k(ord(j)) = k(t) Then
                    If StrComp(v(ord(j), 1), v(t, 1), vbTextCompare) <= 0 Then Exit Do
                End If
                ord(j + 1) = ord(j): j = j - 1
            Loop
            ord(j + 1) = t
        Next i
        Dim out() As Variant: ReDim out(1 To n, 1 To NC)
        For i = 1 To n
            For j = 1 To NC
                out(i, j) = v(ord(i), j)
            Next j
        Next i
        ws.Range("A5").Resize(n, NC).NumberFormat = "@"
        ws.Range("A5").Resize(n, NC).Value = out
        For i = 1 To n
            Select Case Left$(CStr(out(i, 10)), 3)
                Case "NOT": ws.Cells(4 + i, 10).Interior.Color = RGB(255, 205, 210)
                Case "At ": ws.Cells(4 + i, 10).Interior.Color = RGB(255, 236, 179)
            End Select
        Next i
    Else
        ws.Range("A5").Value = "No names on the roster."
    End If
    Dim w As Variant: w = Array(26, 6, 14, 9, 6, 11, 22, 9, 13, 24, 22, 14, 18)
    For i = 1 To NC
        ws.Columns(i).ColumnWidth = w(i - 1)
    Next i
    ws.Range("D5").Resize(IIf(n > 0, n, 1), 6).HorizontalAlignment = xlCenter
    ws.Range("A4").Resize(IIf(n > 0, n, 1) + 1, NC).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
    ProtectSheet ws
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("A5").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    On Error Resume Next
    EndAttendanceBatch                          ' never leave the stale snapshot behind
    If Not ws Is Nothing Then ProtectSheet ws
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    MsgBox "The Attendance sheet couldn't be drawn: " & em, vbExclamation
End Sub
