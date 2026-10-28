Attribute VB_Name = "modStats"
Option Explicit

' ====================================================================
'  STATS (v27, item 26 - briefs/v27.md): the quarterly report
'
'  One admin sheet "Stats": unit x quarter + year to date, drawn like the
'  Refusals Report, a PDF beside the master on request. The counting rules
'  Graeme signed Oct 21 (the report-back): a session = one S- ref; filled =
'  any Accepted row in it; a call = one person row; the sick rate's
'  denominator = scheduled shifts from the rotation (today's crews, the
'  quarter's days up to yesterday); the fairness spread per unit is over
'  TOTAL HOURS (G, the call order's figure), with the premium hours (H) as a
'  second line; honest from the first full quarter after go-live. Four
'  groups: call-outs, postings, sick days, fairness. Every figure comes from
'  the log, the Postings tables, the Absences table and the roster as they
'  are; nothing here changes data and nothing is published (the unit-level
'  cut for the members' file, StatsUnitRows, is wired in with Members M4).
'  The pure classifiers (StatsQuarter, StatsBucket, TimeOfDayBand,
'  SpreadCore) are mirrored in tests/rules.py (suites squarter, sbucket,
'  sband, sspread); the sums are plain loops over those.
' ====================================================================

Public Const STATS_VERSION As String = "Stats v29.0.1 (Oct 28)"      ' stamped with the release at release prep (v27)
Public Const STATS_SHEET As String = "Stats"
Private mAcc As Object                      ' "section|unit|metric|q" -> amount (q 1..4; 5 = the year to date)
Private mUnits As Object                    ' unit -> True (the roster's units, in DeptList order, + "All units")
Private mJan1 As Date                       ' the roster year's Jan 1, worked out once per report (review N1: QuarterNo runs tens of thousands of times)

' ====================================================================
'  THE PURE RULES (mirrored in tests/rules.py)
' ====================================================================
' A date's quarter, "Q1" .. "Q4": the calendar quarter, except that the roster year's first days (its first period starts
' in late December, before jan1) belong to Q1 (review M1; suite squarter)
Public Function StatsQuarter(ByVal d As Date, ByVal jan1 As Date) As String
    If d < jan1 Then StatsQuarter = "Q1" Else StatsQuarter = "Q" & ((Month(d) - 1) \ 3 + 1)
End Function

' What an OT Log row counts as (suite sbucket): void rows "void"; a Callout row by its Outcome - accepted / refused /
' message / unavailable / other; a Posting row - awarded / p-refused (charged, did not sign up) / signedup / other;
' an Adjustment "adjustment"; anything else "other"
Public Function StatsBucket(ByVal typ As String, ByVal outcome As String, ByVal status As String) As String
    If UCase$(Trim$(status)) = "VOID" Then StatsBucket = "void": Exit Function
    Dim t As String: t = LCase$(Trim$(typ))
    Dim o As String: o = LCase$(Trim$(outcome))
    Select Case t
        Case "callout"
            Select Case o
                Case "accepted": StatsBucket = "accepted"
                Case "refused": StatsBucket = "refused"
                Case "left message": StatsBucket = "message"
                Case "unavailable": StatsBucket = "unavailable"
                Case "exception": StatsBucket = "exception"      ' v27 (item 24, review L9): the plant-upset outcome
                Case Else: StatsBucket = "other"
            End Select
        Case "posting"
            Select Case o
                Case "awarded": StatsBucket = "awarded"
                Case "refused": StatsBucket = "p-refused"
                Case "signed up": StatsBucket = "signedup"
                Case Else: StatsBucket = "other"
            End Select
        Case "adjustment": StatsBucket = "adjustment"
        Case Else: StatsBucket = "other"
    End Select
End Function

' The time-of-day band a call-out session was logged in, from its hour 0..23 (suite sband)
Public Function TimeOfDayBand(ByVal hr As Long) As String
    Select Case hr
        Case 0 To 5: TimeOfDayBand = "night 0-6"
        Case 6 To 11: TimeOfDayBand = "morning 6-12"
        Case 12 To 17: TimeOfDayBand = "afternoon 12-18"
        Case Else: TimeOfDayBand = "evening 18-24"
    End Select
End Function

' "h;h;h" -> "min|max|mean|sd|gap" to one decimal (sd = population SD; gap = max - min); "-|-|-|-|-" for no values
' (suite sspread)
Public Function SpreadCore(ByVal spec As String) As String
    Dim parts As Variant, i As Long, n As Long, x As Double, mn As Double, mx As Double, sum As Double, sq As Double
    If Trim$(spec) = "" Then SpreadCore = "-|-|-|-|-": Exit Function
    parts = Split(spec, ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            x = Val(Trim$(CStr(parts(i))))
            n = n + 1
            If n = 1 Then mn = x: mx = x
            If x < mn Then mn = x
            If x > mx Then mx = x
            sum = sum + x
        End If
    Next i
    If n = 0 Then SpreadCore = "-|-|-|-|-": Exit Function
    Dim mean As Double: mean = sum / n
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            x = Val(Trim$(CStr(parts(i))))
            sq = sq + (x - mean) * (x - mean)
        End If
    Next i
    SpreadCore = OneDec(mn) & "|" & OneDec(mx) & "|" & OneDec(mean) & "|" & OneDec(Sqr(sq / n)) & "|" & OneDec(mx - mn)
End Function

Private Function OneDec(ByVal x As Double) As String
    OneDec = Replace(Format$(x, "0.0"), ",", ".")
End Function

' ====================================================================
'  THE REPORT (admin)
' ====================================================================
Public Sub StatsReport()
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating       ' captured before the handler (convention 6)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set mAcc = CreateObject("Scripting.Dictionary"): mAcc.CompareMode = 1
    Set mUnits = CreateObject("Scripting.Dictionary"): mUnits.CompareMode = 1
    Dim depts As Variant: depts = DeptList()
    Dim i As Long
    For i = LBound(depts) To UBound(depts)
        mUnits(CStr(depts(i))) = True
    Next i
    mUnits("All units") = True
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    mJan1 = DateSerial(RosterYear(), 1, 1)
    Dim yearTo As Date: yearTo = RosterYearTo()
    CountCallouts yearFrom, yearTo
    CountPostings yearFrom, yearTo
    CountSick yearFrom, yearTo
    Dim fair As Object: Set fair = FairnessRows(yearFrom, yearTo)
    DrawStats yearFrom, yearTo, fair
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(STATS_SHEET)
    ws.Activate
    If Ask("The Stats sheet is drawn for " & RosterYear() & " (unit x quarter + year to date)." & vbNewLine & vbNewLine & _
           "Save a PDF beside the roster as well?", vbYesNo + vbQuestion, "Stats", vbNo) = vbYes Then
        ReportPdf ws, "Stats " & RosterYear(), "Stats", True
    End If
    Exit Sub
Oops:
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    ReportError "StatsReport"
End Sub

Private Sub Acc(ByVal section As String, ByVal unit As String, ByVal metric As String, ByVal q As Long, ByVal amount As Double)
    If q < 1 Or q > 4 Then Exit Sub
    Dim u As String: u = unit
    If u = "" Or Not mUnits.Exists(u) Then u = "(other)"
    AccOne section & "|" & u & "|" & metric & "|" & q, amount
    AccOne section & "|" & u & "|" & metric & "|5", amount
    If u <> "All units" Then
        AccOne section & "|All units|" & metric & "|" & q, amount
        AccOne section & "|All units|" & metric & "|5", amount
    End If
End Sub

Private Sub AccOne(ByVal key As String, ByVal amount As Double)
    If mAcc.Exists(key) Then mAcc(key) = mAcc(key) + amount Else mAcc.Add key, amount
End Sub

Private Function HasOther() As Boolean
    Dim k As Variant
    For Each k In mAcc.Keys
        If InStr(CStr(k), "|(other)|") > 0 Then HasOther = True: Exit Function
    Next k
End Function

Private Function Got(ByVal section As String, ByVal unit As String, ByVal metric As String, ByVal q As Long) As Double
    Dim key As String: key = section & "|" & unit & "|" & metric & "|" & q
    If mAcc.Exists(key) Then Got = mAcc(key)
End Function

Private Function QuarterNo(ByVal d As Date) As Long
    If mJan1 = 0 Then mJan1 = DateSerial(RosterYear(), 1, 1)
    QuarterNo = CLng(Mid$(StatsQuarter(d, mJan1), 2))
End Function

' ---- A. call-outs: one session = one S- ref ----
Private Sub CountCallouts(ByVal yearFrom As Date, ByVal yearTo As Date)
    Dim lg As Variant: lg = LogRowsNow()
    If Not IsArray(lg) Then Exit Sub
    Dim sessUnit As Object: Set sessUnit = CreateObject("Scripting.Dictionary"): sessUnit.CompareMode = 1
    Dim sessQ As Object: Set sessQ = CreateObject("Scripting.Dictionary"): sessQ.CompareMode = 1
    Dim sessFilled As Object: Set sessFilled = CreateObject("Scripting.Dictionary"): sessFilled.CompareMode = 1
    Dim sessMinOrder As Object: Set sessMinOrder = CreateObject("Scripting.Dictionary"): sessMinOrder.CompareMode = 1   ' the first person actually called
    Dim sessAccOrder As Object: Set sessAccOrder = CreateObject("Scripting.Dictionary"): sessAccOrder.CompareMode = 1   ' the lowest call order that accepted
    Dim sessOffered As Object: Set sessOffered = CreateObject("Scripting.Dictionary"): sessOffered.CompareMode = 1
    Dim sessHour As Object: Set sessHour = CreateObject("Scripting.Dictionary"): sessHour.CompareMode = 1
    Dim sessWeekend As Object: Set sessWeekend = CreateObject("Scripting.Dictionary"): sessWeekend.CompareMode = 1
    Dim i As Long, ref As String, b As String, unit As String, q As Long, d As Date
    For i = 1 To UBound(lg, 1)
        If LCase$(Trim$(SafeStr(lg(i, L_TYPE)))) = "callout" And IsDate(lg(i, L_DATE)) Then
            b = StatsBucket(SafeStr(lg(i, L_TYPE)), SafeStr(lg(i, L_OUTCOME)), SafeStr(lg(i, L_STATUS)))
            If b <> "void" Then
                d = CDate(lg(i, L_DATE))
                If d >= yearFrom And d <= yearTo Then
                    ref = Trim$(SafeStr(lg(i, L_REF))): unit = Trim$(SafeStr(lg(i, L_DEPT))): q = QuarterNo(d)
                    If Not sessUnit.Exists(ref) Then
                        sessUnit.Add ref, unit: sessQ.Add ref, q
                        sessFilled.Add ref, False: sessMinOrder.Add ref, 0#: sessAccOrder.Add ref, 0#
                        sessOffered.Add ref, NumOr0(lg(i, L_BASEHRS))
                        If IsDate(lg(i, L_STAMP)) Then sessHour.Add ref, Hour(CDate(lg(i, L_STAMP))) Else sessHour.Add ref, -1
                        sessWeekend.Add ref, (Weekday(d, vbMonday) >= 6)
                    End If
                    ' (review M2) the hours on offer: the most any row shows, and the "of M h" a capped or partial row's Note names
                    If OfferedOnRow(lg(i, L_BASEHRS), SafeStr(lg(i, L_REASON))) > sessOffered(ref) Then sessOffered(ref) = OfferedOnRow(lg(i, L_BASEHRS), SafeStr(lg(i, L_REASON)))
                    If NumOr0(lg(i, L_ORDER)) > 0 Then
                        If sessMinOrder(ref) = 0 Or NumOr0(lg(i, L_ORDER)) < sessMinOrder(ref) Then sessMinOrder(ref) = NumOr0(lg(i, L_ORDER))
                    End If
                    Acc "A", unit, "calls", q, 1
                    Acc "A", unit, b, q, 1
                    Acc "A", unit, "charged", q, NumOr0(lg(i, L_CHARGED))
                    If b = "accepted" Then
                        sessFilled(ref) = True
                        Acc "A", unit, "accepted-h", q, NumOr0(lg(i, L_BASEHRS))
                        If sessAccOrder(ref) = 0 Or NumOr0(lg(i, L_ORDER)) < sessAccOrder(ref) Then sessAccOrder(ref) = NumOr0(lg(i, L_ORDER))
                    End If
                End If
            End If
        End If
    Next i
    Dim k As Variant
    For Each k In sessUnit.Keys
        unit = sessUnit(k): q = sessQ(k)
        Acc "A", unit, "sessions", q, 1
        Acc "A", unit, "offered-h", q, sessOffered(k)
        If sessFilled(k) Then Acc "A", unit, "filled", q, 1
        If sessFilled(k) And sessAccOrder(k) > 0 And sessAccOrder(k) = sessMinOrder(k) Then Acc "A", unit, "first", q, 1   ' (review) the first person called took it
        If sessWeekend(k) Then Acc "A", unit, "weekend", q, 1
        If sessHour(k) >= 0 Then Acc "A", unit, "band " & TimeOfDayBand(CLng(sessHour(k))), q, 1
    Next k
End Sub

' The hours a call-out row says were on offer: the row's Hours, or the M in a Note that reads "Took N of M h" /
' "Offered N of M h" (a partial acceptance or a 16 h cap wrote the smaller N into Hours)
Private Function OfferedOnRow(ByVal hrs As Variant, ByVal note As String) As Double
    OfferedOnRow = NumOr0(hrs)
    Dim p As Long: p = InStr(1, note, " of ", vbTextCompare)
    If p = 0 Then Exit Function
    Dim rest As String: rest = Mid$(note, p + 4)
    Dim q As Long: q = InStr(rest, " h")
    If q = 0 Then Exit Function
    Dim m As Double: m = Val(Left$(rest, q - 1))
    If m > OfferedOnRow Then OfferedOnRow = m
End Function

' ---- B. postings: the Postings table, its list, and the Posting rows of the log ----
Private Sub CountPostings(ByVal yearFrom As Date, ByVal yearTo As Date)
    Dim v As Variant: v = TableByName(PT_TABLE)
    Dim postUnit As Object: Set postUnit = CreateObject("Scripting.Dictionary"): postUnit.CompareMode = 1
    Dim postQ As Object: Set postQ = CreateObject("Scripting.Dictionary"): postQ.CompareMode = 1
    Dim i As Long, id As String, st As String, unit As String, q As Long, d As Date
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            id = Trim$(SafeStr(v(i, PS_ID)))
            If id <> "" And IsDate(v(i, PS_DATE)) Then
                d = CDate(v(i, PS_DATE))
                If d >= yearFrom And d <= yearTo Then
                    unit = Trim$(SafeStr(v(i, PS_UNIT))): q = QuarterNo(d): st = LCase$(Trim$(SafeStr(v(i, PS_STATUS))))
                    postUnit(id) = unit: postQ(id) = q
                    Acc "B", unit, "posted", q, 1
                    Select Case st
                        Case "awarded"
                            Acc "B", unit, "awarded", q, 1
                            If IsDate(v(i, PS_AWARDON)) And IsDate(v(i, PS_POSTED)) Then
                                Acc "B", unit, "award-days", q, CDbl(Int(CDate(v(i, PS_AWARDON)))) - CDbl(Int(CDate(v(i, PS_POSTED))))
                                Acc "B", unit, "award-n", q, 1
                            End If
                        Case "no takers": Acc "B", unit, "notakers", q, 1
                        Case "cancelled": Acc "B", unit, "cancelled", q, 1
                        Case "open": Acc "B", unit, "open", q, 1
                    End Select
                End If
            End If
        Next i
    End If
    v = TableByName(PL_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            id = Trim$(SafeStr(v(i, PL_ID)))
            If postUnit.Exists(id) Then
                If UCase$(Trim$(SafeStr(v(i, PL_SIGNED)))) = "YES" Then Acc "B", CStr(postUnit(id)), "signups", CLng(postQ(id)), 1
            End If
        Next i
    End If
    Dim lg As Variant: lg = LogRowsNow()
    If Not IsArray(lg) Then Exit Sub
    Dim b As String
    For i = 1 To UBound(lg, 1)
        If LCase$(Trim$(SafeStr(lg(i, L_TYPE)))) = "posting" And IsDate(lg(i, L_DATE)) Then
            b = StatsBucket(SafeStr(lg(i, L_TYPE)), SafeStr(lg(i, L_OUTCOME)), SafeStr(lg(i, L_STATUS)))
            If b <> "void" Then
                d = CDate(lg(i, L_DATE))
                If d >= yearFrom And d <= yearTo Then
                    unit = Trim$(SafeStr(lg(i, L_DEPT))): q = QuarterNo(d)
                    Acc "B", unit, "charged", q, NumOr0(lg(i, L_CHARGED))
                    If b = "p-refused" Then Acc "B", unit, "p-refused", q, 1
                    If b = "awarded" And LCase$(Left$(Trim$(SafeStr(lg(i, L_REASON))), 17)) = "out of department" Then Acc "B", unit, "outside", q, 1
                End If
            End If
        End If
    Next i
End Sub

' ---- C. sick days: SIC bookings, and the rotation's scheduled shifts as the denominator ----
Private Sub CountSick(ByVal yearFrom As Date, ByVal yearTo As Date)
    Dim v As Variant: v = AbsTableData()
    Dim i As Long, unit As String, d As Date
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If UCase$(Trim$(SafeStr(v(i, AB_TYPE)))) = "SIC" And UCase$(Trim$(SafeStr(v(i, AB_STATUS)))) <> "VOID" And IsDate(v(i, AB_START)) Then
                d = CDate(v(i, AB_START))
                If d >= yearFrom And d <= yearTo Then
                    unit = Trim$(SafeStr(v(i, AB_UNIT)))
                    If unit = "" Then unit = Trim$(SafeStr(v(i, AB_HOME)))
                    Acc "C", unit, "sick", QuarterNo(d), 1
                    If LCase$(Trim$(SafeStr(v(i, AB_SHIFT)))) = "nights" Then Acc "C", unit, "nights", QuarterNo(d), 1
                    If Weekday(d, vbMonday) >= 6 Then Acc "C", unit, "weekend", QuarterNo(d), 1
                    If DaysHolidayOff(d - 1) Or DaysHolidayOff(d) Or DaysHolidayOff(d + 1) Then Acc "C", unit, "holiday", QuarterNo(d), 1
                End If
            End If
        Next i
    End If
    ' the denominator: for each unit's people (a code in its column today), the rotation's shifts on each day of the
    ' year up to yesterday (today's crew: column E - the report says so)
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    Dim rv As Variant: rv = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, C_LAST)).Value
    Dim depts As Variant: depts = DeptList()
    Dim u As Long, dc As Long, r As Long, crew As String, lastDay As Date, sh As String, q As Long
    Dim personDone As Object: Set personDone = CreateObject("Scripting.Dictionary"): personDone.CompareMode = 1   ' (review M5): All units once per person
    lastDay = Date - 1
    If lastDay > yearTo Then lastDay = yearTo
    For u = LBound(depts) To UBound(depts)
        dc = DeptCol(CStr(depts(u)))
        If dc > 0 Then
            For r = 1 To UBound(rv, 1)
                If Trim$(SafeStr(rv(r, C_NAME))) <> "" And Trim$(SafeStr(rv(r, dc))) <> "" Then
                    crew = Trim$(SafeStr(rv(r, C_SHIFT)))
                    If crew <> "" Then
                        For d = yearFrom To lastDay
                            sh = CrewShiftOn(crew, d)
                            If sh = "Days" Or sh = "Nights" Then
                                q = QuarterNo(d)
                                AccOne "C|" & CStr(depts(u)) & "|scheduled|" & q, 1
                                AccOne "C|" & CStr(depts(u)) & "|scheduled|5", 1
                                If Not personDone.Exists(Trim$(SafeStr(rv(r, C_NAME)))) Then
                                    AccOne "C|All units|scheduled|" & q, 1
                                    AccOne "C|All units|scheduled|5", 1
                                End If
                            End If
                        Next d
                        personDone(Trim$(SafeStr(rv(r, C_NAME)))) = True
                    End If
                End If
            Next r
        End If
    Next u
End Sub

' ---- D. fairness: today's figures per unit, grade <= CHARGE_MAX_GRADE; the refusal rate from the log ----
' unit -> "total spread|premium spread|refusal rate|people"
Private Function FairnessRows(ByVal yearFrom As Date, ByVal yearTo As Date) As Object
    Dim out As Object: Set out = CreateObject("Scripting.Dictionary"): out.CompareMode = 1
    Set FairnessRows = out
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim rv As Variant: rv = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, C_LAST)).Value
    Dim lg As Variant: lg = LogRowsNow()
    ' per person this year: offers (a call-out row accepted / refused / left message; a posting row awarded / charged)
    ' and refusals (refused / left message / charged on a posting) - unavailable and signed-up rows are neither
    Dim offers As Object: Set offers = CreateObject("Scripting.Dictionary"): offers.CompareMode = 1
    Dim refusals As Object: Set refusals = CreateObject("Scripting.Dictionary"): refusals.CompareMode = 1
    Dim i As Long, b As String, nm As String
    If IsArray(lg) Then
        For i = 1 To UBound(lg, 1)
            If IsDate(lg(i, L_DATE)) Then
                If CDate(lg(i, L_DATE)) >= yearFrom And CDate(lg(i, L_DATE)) <= yearTo Then
                    b = StatsBucket(SafeStr(lg(i, L_TYPE)), SafeStr(lg(i, L_OUTCOME)), SafeStr(lg(i, L_STATUS)))
                    nm = Trim$(SafeStr(lg(i, L_EMP)))
                    If nm <> "" Then
                        If b = "accepted" Or b = "refused" Or b = "message" Or b = "awarded" Or b = "p-refused" Then offers(nm) = NumOr0(offers(nm)) + 1
                        If b = "refused" Or b = "message" Or b = "p-refused" Then refusals(nm) = NumOr0(refusals(nm)) + 1
                    End If
                End If
            End If
        Next i
    End If
    Dim depts As Variant: depts = DeptList()
    Dim u As Long, dc As Long, r As Long, tot As String, prem As String, nOff As Double, nRef As Double, nPeople As Long, rate As String
    For u = LBound(depts) To UBound(depts)
        dc = DeptCol(CStr(depts(u)))
        tot = "": prem = "": nOff = 0: nRef = 0: nPeople = 0
        If dc > 0 Then
            For r = 1 To UBound(rv, 1)
                nm = Trim$(SafeStr(rv(r, C_NAME)))
                If nm <> "" And Trim$(SafeStr(rv(r, dc))) <> "" Then
                    If DeptCodeValue(Trim$(SafeStr(rv(r, dc)))) <= CHARGE_MAX_GRADE Then
                        nPeople = nPeople + 1
                        tot = tot & IIf(tot = "", "", ";") & Trim$(Str$(NumOr0(rv(r, C_TOTAL))))
                        prem = prem & IIf(prem = "", "", ";") & Trim$(Str$(NumOr0(rv(r, C_WORKED))))
                        nOff = nOff + NumOr0(offers(nm)): nRef = nRef + NumOr0(refusals(nm))
                    End If
                End If
            Next r
        End If
        ' (review H1) a block, not IIf: 0 / 0 inside an IIf raises even when the other branch is taken
        If nOff > 0 Then rate = OneDec(100 * nRef / nOff) & "%" Else rate = "-"
        out(CStr(depts(u))) = SpreadCore(tot) & "|" & SpreadCore(prem) & "|" & rate & "|" & nPeople
    Next u
End Function

' ---- the sheet ----
Private Sub DrawStats(ByVal yearFrom As Date, ByVal yearTo As Date, ByVal fair As Object)
    Dim ws As Worksheet
    If Not SheetExists(STATS_SHEET) Then
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = STATS_SHEET
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)
    End If
    Set ws = ThisWorkbook.Worksheets(STATS_SHEET)
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "STATS  -  " & RosterYear() & "  -  unit x quarter + year to date"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Color = RGB(30, 56, 100)
    ws.Range("A2").Value = "From the OT Log (non-void rows), the Postings tables, the Absences table and the roster as they are, " & _
                           "quarters by the shift date (" & Format$(yearFrom, "mmm d, yyyy") & " - " & Format$(yearTo, "mmm d, yyyy") & "). " & _
                           "Honest from the first full quarter after go-live: earlier quarters hold what was logged, not what happened. Drawn " & _
                           Format$(Now, "mmm d, yyyy h:mm AM/PM") & ".  " & TrustStrip()
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    With ws.Range("A2:L2")                                   ' (review M4) one wrapped line across the page, not a cut-off cell
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
    End With
    ws.Rows(2).RowHeight = 48
    Dim r As Long: r = 4
    Dim depts As Variant: depts = DeptList()
    Dim units() As String, nU As Long, u As Long
    ReDim units(1 To UBound(depts) - LBound(depts) + 2)
    For u = LBound(depts) To UBound(depts)
        nU = nU + 1: units(nU) = CStr(depts(u))
    Next u
    nU = nU + 1: units(nU) = "All units"
    If HasOther() Then
        ReDim Preserve units(1 To nU + 1)
        nU = nU + 1: units(nU) = "(other)"             ' (review) a unit name the roster does not know, so the unit rows add up to All units
    End If

    ' A
    r = SectionTitle(ws, r, "A. Call-outs (a session = one S- ref; filled = anyone accepted; a call = one person asked)")
    For u = 1 To nU
        r = UnitHeader(ws, r, units(u))
        r = CountRow(ws, r, "Sessions", "A", units(u), "sessions")
        r = CountRow(ws, r, "Filled (anyone accepted)", "A", units(u), "filled")
        r = RateRow(ws, r, "Fill rate", "A", units(u), "filled", "sessions", True)
        r = RateRow(ws, r, "Calls per filled session", "A", units(u), "calls", "filled", False)
        r = CountRow(ws, r, "Calls: accepted", "A", units(u), "accepted")
        r = CountRow(ws, r, "Calls: refused", "A", units(u), "refused")
        r = CountRow(ws, r, "Calls: left message", "A", units(u), "message")
        r = CountRow(ws, r, "Calls: unavailable", "A", units(u), "unavailable")
        r = CountRow(ws, r, "Calls: exception (plant upset)", "A", units(u), "exception")     ' v27 (item 24, review L9)
        r = CountRow(ws, r, "Hours offered (per session)", "A", units(u), "offered-h")
        r = CountRow(ws, r, "Hours accepted", "A", units(u), "accepted-h")
        r = CountRow(ws, r, "Hours charged (refusals x 1.75)", "A", units(u), "charged")
        r = RateRow(ws, r, "Filled by the first person called", "A", units(u), "first", "filled", True)
        r = RateRow(ws, r, "Sessions on a weekend", "A", units(u), "weekend", "sessions", True)
        If units(u) = "All units" Then
            r = CountRow(ws, r, "Logged: night 0-6", "A", units(u), "band night 0-6")
            r = CountRow(ws, r, "Logged: morning 6-12", "A", units(u), "band morning 6-12")
            r = CountRow(ws, r, "Logged: afternoon 12-18", "A", units(u), "band afternoon 12-18")
            r = CountRow(ws, r, "Logged: evening 18-24", "A", units(u), "band evening 18-24")
        End If
    Next u
    ' B
    r = SectionTitle(ws, r, "B. Postings (by the shift date; charges = the posting rows of the log)")
    For u = 1 To nU
        r = UnitHeader(ws, r, units(u))
        r = CountRow(ws, r, "Posted", "B", units(u), "posted")
        r = CountRow(ws, r, "Awarded", "B", units(u), "awarded")
        r = CountRow(ws, r, "Awarded out of department", "B", units(u), "outside")
        r = CountRow(ws, r, "No takers", "B", units(u), "notakers")
        r = CountRow(ws, r, "Cancelled", "B", units(u), "cancelled")
        r = CountRow(ws, r, "Still open", "B", units(u), "open")
        r = RateRow(ws, r, "Sign-ups per posting", "B", units(u), "signups", "posted", False)
        r = RateRow(ws, r, "Days from posting to award", "B", units(u), "award-days", "award-n", False)
        r = CountRow(ws, r, "People charged (did not sign up)", "B", units(u), "p-refused")
        r = CountRow(ws, r, "Hours charged", "B", units(u), "charged")
    Next u
    ' C
    r = SectionTitle(ws, r, "C. Sick days (sick calls booked; the rate per scheduled shift from the rotation, today's crews, days up to yesterday)")
    For u = 1 To nU
        r = UnitHeader(ws, r, units(u))
        r = CountRow(ws, r, "Sick shifts", "C", units(u), "sick")
        r = CountRow(ws, r, "Scheduled shifts (rotation)", "C", units(u), "scheduled")
        r = RateRow(ws, r, "Sick rate per scheduled shift", "C", units(u), "sick", "scheduled", True)
        r = RateRow(ws, r, "On nights", "C", units(u), "nights", "sick", True)
        r = RateRow(ws, r, "On a weekend", "C", units(u), "weekend", "sick", True)
        r = RateRow(ws, r, "On or beside a plant holiday", "C", units(u), "holiday", "sick", True)
    Next u
    ' D
    r = SectionTitle(ws, r, "D. Fairness today (grade " & CHARGE_MAX_GRADE & " and below with a code in the unit): the spread of Total Hours, then of premium hours worked; the refusal rate this year = refused + left message + charged on a posting, over every call-out and posting row that asked them")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Value = Array("Unit", "People", "Total min", "Total max", "Total mean", "Total SD", "Top-to-bottom gap", _
                                                           "Worked min", "Worked max", "Worked mean", "Worked SD", "Refusal rate")
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 12))
    r = r + 1
    Dim parts As Variant
    For u = LBound(depts) To UBound(depts)
        If fair.Exists(CStr(depts(u))) Then
            parts = Split(CStr(fair(CStr(depts(u)))), "|")      ' 5 total | 5 premium | rate | people
            ws.Cells(r, 1).Value = CStr(depts(u))
            ws.Cells(r, 2).Value = Val(parts(11))
            ws.Cells(r, 3).Value = parts(0): ws.Cells(r, 4).Value = parts(1): ws.Cells(r, 5).Value = parts(2): ws.Cells(r, 6).Value = parts(3): ws.Cells(r, 7).Value = parts(4)
            ws.Cells(r, 8).Value = parts(5): ws.Cells(r, 9).Value = parts(6): ws.Cells(r, 10).Value = parts(7): ws.Cells(r, 11).Value = parts(8)
            ws.Cells(r, 12).Value = parts(10)
            r = r + 1
        End If
    Next u
    ws.Columns(1).ColumnWidth = 44
    ws.Range(ws.Columns(2), ws.Columns(12)).ColumnWidth = 11
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 12)).Address
        .Orientation = xlPortrait: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = "Stats " & RosterYear() & " - Page &P of &N"
    End With
    On Error GoTo 0
    ProtectSheet ws
End Sub

Private Function SectionTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String) As Long
    ws.Cells(r + 1, 1).Value = txt
    ws.Cells(r + 1, 1).Font.Bold = True: ws.Cells(r + 1, 1).Font.Size = 12: ws.Cells(r + 1, 1).Font.Color = RGB(30, 56, 100)
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, 12))      ' (review N2) a long title wraps across the page instead of being cut off in the PDF
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
    End With
    ws.Rows(r + 1).RowHeight = IIf(Len(txt) > 110, 48, 24)
    SectionTitle = r + 2
End Function

Private Function UnitHeader(ByVal ws As Worksheet, ByVal r As Long, ByVal unit As String) As Long
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Value = Array(unit, "Q1", "Q2", "Q3", "Q4", "Year to date")
    HeaderStyle ws.Range(ws.Cells(r, 1), ws.Cells(r, 6))
    UnitHeader = r + 1
End Function

Private Sub HeaderStyle(ByVal rng As Range)
    rng.Font.Bold = True: rng.Font.Color = vbWhite: rng.Interior.Color = RGB(30, 56, 100)
    rng.HorizontalAlignment = xlCenter
    rng.Cells(1, 1).HorizontalAlignment = xlLeft
End Sub

Private Function CountRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal section As String, ByVal unit As String, ByVal metric As String) As Long
    ws.Cells(r, 1).Value = label
    Dim q As Long
    For q = 1 To 5
        ws.Cells(r, 1 + q).Value = Got(section, unit, metric, q)
    Next q
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).NumberFormat = "General"
    CountRow = r + 1
End Function

' a ratio row: percent (num / den x 100, one decimal) or a plain ratio (one decimal); "-" when den = 0
Private Function RateRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByVal section As String, ByVal unit As String, _
                         ByVal num As String, ByVal den As String, ByVal percent As Boolean) As Long
    ws.Cells(r, 1).Value = label
    Dim q As Long, a As Double, b As Double
    For q = 1 To 5
        a = Got(section, unit, num, q): b = Got(section, unit, den, q)
        ws.Cells(r, 1 + q).NumberFormat = "@"
        If b = 0 Then
            ws.Cells(r, 1 + q).Value = "-"
        ElseIf percent Then
            ws.Cells(r, 1 + q).Value = OneDec(100 * a / b) & "%"
        Else
            ws.Cells(r, 1 + q).Value = OneDec(a / b)
        End If
    Next q
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Font.Color = RGB(90, 100, 110)
    RateRow = r + 1
End Function

' A table's rows by name, from any sheet (Empty when none)
Private Function TableByName(ByVal tableName As String) As Variant
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then Exit For
    Next ws
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    TableByName = RangeTo2D(lo.DataBodyRange)
End Function

' The unit-level cut for the members' file (Members M4 publishes it as dStats): Unit, Quarter, Metric, Value for the
' call-out, posting and sick-day COUNTS of sections A-C (no names, no fairness). Empty until StatsReport has run.
Public Function StatsUnitRows() As Variant
    If mAcc Is Nothing Then Exit Function
    Dim out() As Variant: ReDim out(1 To mAcc.Count + 1, 1 To 4)
    out(1, 1) = "Unit": out(1, 2) = "Quarter": out(1, 3) = "Metric": out(1, 4) = "Value"
    Dim k As Variant, p As Variant, n As Long
    n = 1
    For Each k In mAcc.Keys
        p = Split(CStr(k), "|")
        If CStr(p(0)) <> "D" And CStr(p(1)) <> "(other)" Then
            n = n + 1
            out(n, 1) = p(1): out(n, 2) = IIf(CStr(p(3)) = "5", "Year to date", "Q" & p(3)): out(n, 3) = p(0) & " " & p(2): out(n, 4) = mAcc(k)
        End If
    Next k
    If n < UBound(out, 1) Then
        Dim trimmed() As Variant: ReDim trimmed(1 To n, 1 To 4)
        Dim a As Long, b As Long
        For a = 1 To n
            For b = 1 To 4
                trimmed(a, b) = out(a, b)
            Next b
        Next a
        StatsUnitRows = trimmed
    Else
        StatsUnitRows = out
    End If
End Function
