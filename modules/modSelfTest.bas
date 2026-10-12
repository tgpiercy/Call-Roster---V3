Attribute VB_Name = "modSelfTest"
Option Explicit

' ====================================================================
'  SELF TEST v1 (v21)
'  Alt+F8 > SelfTest runs known-answer cases against the pure rules:
'  rotation, holidays, DST nights, call order, charging, eligibility
'  and the provisional-OT date logic. It reports pass / fail per case.
'
'  It works on in-memory data only. It never writes to the roster, the
'  OT Log or any table, and never saves. Eligibility cases that need an
'  absence or a shift switch feed those rows to the rule from memory
'  (SetTestFixtures / SetTestSwitches / SetTestHolidays) and clear them
'  again at the end, even if a case errors.
'
'  The cases are the SAME table as tests/cases.csv in the repo (the
'  Python mirror reads it directly). Don't edit the block between the
'  CASES markers by hand - run tests/gen_selftest_cases.py.
' ====================================================================

Public Const ST_VERSION As String = "Self test v22 (Oct 13)"

' provcount cases assume this roster year (26 periods from Dec 28, 2025)
Private Const ST_YEAR_FROM As Date = #12/28/2025#
Private Const ST_YEAR_TO As Date = #12/26/2026#
Private Const ST_NAME As String = "Test Person"

Private mCases() As String
Private mN As Long


' ---------------- entry ----------------
Public Sub SelfTest()
    mN = 0
    LoadCases
    If mN = 0 Then MsgBox "No test cases found in modSelfTest.", vbExclamation: Exit Sub

    Dim i As Long, f() As String, actual As String, expected As String
    Dim nPass As Long, nFail As Long, fails As String, suite As String
    Dim bySuite As Object: Set bySuite = CreateObject("Scripting.Dictionary")
    On Error GoTo Finally
    For i = 1 To mN
        f = Split(mCases(i), ",")
        If UBound(f) <> 7 Then
            actual = "BAD ROW (" & (UBound(f) + 1) & " fields)"
            expected = ""
            suite = "?"
            ReDim Preserve f(0 To 7)              ' so f(1) below is safe to read
        Else
            suite = f(0)
            expected = f(7)
            actual = RunCase(f)
        End If
        If StrComp(Trim$(actual), Trim$(expected), vbTextCompare) = 0 Then
            nPass = nPass + 1
        Else
            nFail = nFail + 1
            If nFail <= 15 Then fails = fails & "  " & suite & " / " & f(1) & ": got """ & actual & """, expected """ & expected & """" & vbNewLine
            Debug.Print "SelfTest FAIL " & mCases(i) & " -> " & actual
        End If
        bySuite(suite) = bySuite(suite) & IIf(StrComp(Trim$(actual), Trim$(expected), vbTextCompare) = 0, ".", "F")
    Next i
    If nFail > 15 Then fails = fails & "  ...and " & (nFail - 15) & " more (see the Immediate window)" & vbNewLine

Finally:
    Dim em As String
    If Err.Number <> 0 Then em = "Stopped by an error: " & Err.Description & vbNewLine
    On Error Resume Next
    ClearTestFixtures
    ClearTestSwitches
    ClearTestHolidays
    Dim k As Variant, summary As String
    For Each k In bySuite.Keys
        summary = summary & "  " & k & ": " & Len(bySuite(k)) - Len(Replace(bySuite(k), "F", "")) & " failed of " & Len(bySuite(k)) & vbNewLine
    Next k
    Dim msg As String
    msg = em & "SelfTest: " & nPass & " passed, " & nFail & " failed of " & mN & " cases." & vbNewLine & vbNewLine & summary
    If fails <> "" Then msg = msg & vbNewLine & "Failed:" & vbNewLine & fails
    msg = msg & vbNewLine & "Nothing was written or saved. (" & ST_VERSION & ")"
    MsgBox msg, IIf(nFail = 0 And em = "", vbInformation, vbExclamation), "Self test"
End Sub


' ---------------- one case ----------------
' f(0) suite, f(1) case, f(2..6) in1..in5, f(7) expected. Returns the actual answer.
Private Function RunCase(f() As String) As String
    On Error GoTo Oops
    Dim i1 As String, i2 As String, i3 As String, i4 As String, i5 As String
    i1 = f(2): i2 = f(3): i3 = f(4): i4 = f(5): i5 = f(6)
    Select Case f(0)
        Case "rotation"
            RunCase = DaysCrew(ParseD(i1)) & "/" & NightsCrew(ParseD(i1))
        Case "crewshift"
            RunCase = CrewShiftOn(i1, ParseD(i2))
            If RunCase = "" Then RunCase = "off"
        Case "avail"
            RunCase = AvailableCrew(ParseD(i1), (i2 = "N"))
        Case "holiday"
            RunCase = FmtD(HolidayDate(i1, CLng(i2)))
        Case "holoff"
            RunCase = CaseHolOff(i1, i2)
        Case "dst"
            RunCase = FmtD(IIf(i2 = "fall", DstFallBack(CLng(i1)), DstSpringForward(CLng(i1))))
        Case "dsthours"
            RunCase = NumText(DstNightHours(ParseD(i1)))
        Case "callorder"
            RunCase = CaseCallOrder(i1)
        Case "callout"
            RunCase = NumText(CalloutCharge(i1, CDbl(i2), CDbl(i3)))
        Case "posting"
            RunCase = CasePosting(i1, CDbl(i2), i3)
        Case "eligibility"
            RunCase = CaseEligibility(i1, ParseD(i2), i3, i4)
        Case "prov"
            RunCase = IIf(ProvisionalCovered(ParseD(i1), (i2 = "N"), ParseD(i3)), "covered", "counts")
        Case "provcount"
            RunCase = IIf(ProvisionalCounts(i1, i2, ParseD(i3), (i4 = "N"), ParseD(i5), ST_YEAR_FROM, ST_YEAR_TO), "counts", "no")
        Case "sick"
            RunCase = IIf(SickBlocks(ParseD(i1), (i2 = "N"), ParseD(i3), (i4 = "N")), "blocked", "free")
        Case "hoursasof"
            RunCase = CaseHoursAsOf(i1, i2, ParseD(i3))
        Case "leavepeer"
            RunCase = IIf(PeerEligible(CDbl(i1), (i2 = "Y"), (i3 = "Y"), (i4 = "Y")), "peer", "excluded")
        Case "leaveavg"
            RunCase = CaseLeaveAvg(i1)
        Case "merge"
            RunCase = CaseMerge(i1)
        Case Else
            RunCase = "UNKNOWN SUITE"
    End Select
    Exit Function
Oops:
    RunCase = "ERROR: " & Err.Description
End Function

' DAYS workers off on in1? in2 = "Holiday|year|observed" fixture or ""
Private Function CaseHolOff(ByVal dTxt As String, ByVal fixture As String) As String
    Dim rows As Variant
    If fixture <> "" Then
        Dim p As Variant: p = Split(fixture, "|")
        Dim r(1 To 1, 1 To 6) As Variant         ' Year, Holiday, Date, Day, Observed on, Note
        r(1, 1) = CLng(p(1))
        r(1, 2) = CStr(p(0))
        r(1, 3) = HolidayDate(CStr(p(0)), CLng(p(1)))
        r(1, 4) = Format$(r(1, 3), "ddd")
        r(1, 5) = ParseD(CStr(p(2)))
        r(1, 6) = ""
        rows = r
    Else
        rows = Empty
    End If
    SetTestHolidays rows
    CaseHolOff = IIf(DaysHolidayOff(ParseD(dTxt)), "off", "on")
    ClearTestHolidays
End Function

' "Name:grade:hours:seniority;..." -> names in call order
Private Function CaseCallOrder(ByVal spec As String) As String
    Dim recsTxt As Variant: recsTxt = Split(spec, ";")
    Dim n As Long: n = UBound(recsTxt) + 1
    Dim recs() As EmpRec: ReDim recs(1 To n)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(recsTxt(i - 1), ":")
        recs(i).EmpName = CStr(p(0))
        recs(i).Grade = CDbl(p(1))
        recs(i).TotalHrs = CDbl(p(2))
        If CStr(p(3)) = "" Then
            recs(i).Seniority = SeniorityKey(Empty)
        Else
            recs(i).Seniority = SeniorityKey(ParseD(CStr(p(3))))
        End If
    Next i
    SortEmpRecs recs, n
    Dim s As String
    For i = 1 To n
        s = s & IIf(i = 1, "", "/") & recs(i).EmpName
    Next i
    CaseCallOrder = s
End Function

' mode (award / notakers / outside / cancel), hours, "rank:grade:signed:awarded;..."
Private Function CasePosting(ByVal mode As String, ByVal hours As Double, ByVal spec As String) As String
    Dim items As Variant: items = Split(spec, ";")
    Dim n As Long: n = UBound(items) + 1
    Dim lst() As PListRec: ReDim lst(1 To n)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(items(i - 1), ":")
        lst(i).Rank = CLng(p(0))
        lst(i).Grade = CDbl(p(1))
        lst(i).SignedUp = (CStr(p(2)) = "Y")
        lst(i).Awarded = (CStr(p(3)) = "Y")
    Next i
    Dim pr As PostRec: pr.Hours = hours
    Dim maxRank As Long, maxHrs As Double, nAw As Long, nCh As Long, tot As Double
    Select Case mode
        Case "award": ComputeAward pr, lst, n, maxRank, maxHrs, nAw, nCh, tot
        Case "notakers": ComputeNoTakers pr, lst, n, nCh, tot
        Case "outside": ComputeOutside pr, lst, n, nCh, tot
        Case "cancel": ComputeCancel lst, n
        Case Else: CasePosting = "UNKNOWN MODE": Exit Function
    End Select
    ' rank order
    Dim j As Long, t As PListRec
    For i = 2 To n
        t = lst(i): j = i - 1
        Do While j >= 1
            If lst(j).Rank <= t.Rank Then Exit Do
            lst(j + 1) = lst(j): j = j - 1
        Loop
        lst(j + 1) = t
    Next i
    Dim s As String
    For i = 1 To n
        s = s & IIf(i = 1, "", ";") & lst(i).Rank & "=" & NumText(lst(i).Charged)
    Next i
    CasePosting = s
End Function

' crew code, date, shift, fixture ("", "off:VAC", "ss", "ssoff", "ssrest") -> reason or "free"
Private Function CaseEligibility(ByVal code As String, ByVal d As Date, ByVal shiftLbl As String, ByVal fixture As String) As String
    Dim absRows As Variant, swRows As Variant
    absRows = Empty: swRows = Empty
    If Left$(fixture, 4) = "off:" Then
        Dim a(1 To 1, 1 To AB_N) As Variant
        a(1, AB_ID) = "AB-9999"
        a(1, AB_EMP) = ST_NAME
        a(1, AB_UNIT) = "Phosphine"
        a(1, AB_TYPE) = Mid$(fixture, 5)
        a(1, AB_START) = d
        a(1, AB_END) = d
        a(1, AB_STATUS) = "Active"
        absRows = a
    ElseIf fixture = "ss" Or fixture = "ssoff" Or fixture = "ssrest" Then
        Dim s(1 To 1, 1 To SW_N) As Variant
        s(1, SW_ID) = "SS-9999"
        s(1, SW_UNIT) = "Phosphine"
        s(1, SW_STATUS) = "Active"
        If fixture = "ss" Then
            s(1, SW_DATE) = d: s(1, SW_SHIFT) = shiftLbl: s(1, SW_WORKS) = ST_NAME: s(1, SW_OFF) = "Someone Else"
        ElseIf fixture = "ssoff" Then
            s(1, SW_DATE) = d: s(1, SW_SHIFT) = shiftLbl: s(1, SW_WORKS) = "Someone Else": s(1, SW_OFF) = ST_NAME
        Else
            ' the adjoining shift: nights before a day OT, days after a night OT
            If IsNightShift(shiftLbl) Then
                s(1, SW_DATE) = d + 1: s(1, SW_SHIFT) = "Days"
            Else
                s(1, SW_DATE) = d - 1: s(1, SW_SHIFT) = "Nights"
            End If
            s(1, SW_WORKS) = ST_NAME: s(1, SW_OFF) = "Someone Else"
        End If
        swRows = s
    ElseIf fixture <> "" Then
        CaseEligibility = "UNKNOWN FIXTURE": Exit Function
    End If
    SetTestFixtures absRows
    SetTestSwitches swRows
    CaseEligibility = NotFreeReason(ST_NAME, code, d, shiftLbl)
    ClearTestFixtures
    ClearTestSwitches
    If CaseEligibility = "" Then CaseEligibility = "free"
End Function


' hh "through:hours;..." and log "date:outcome:status:hours:charged:D|N;..." for one
' person -> "total|through|est" (v22: the pure core modLeave.HoursAsOfCore)
Private Function CaseHoursAsOf(ByVal hhSpec As String, ByVal lgSpec As String, ByVal asOf As Date) As String
    Dim hh As Variant, lg As Variant, parts As Variant, p As Variant, i As Long, n As Long
    If hhSpec <> "" Then
        parts = Split(hhSpec, ";")
        n = UBound(parts) + 1
        Dim h() As Variant: ReDim h(1 To n, 1 To 3)
        For i = 1 To n
            p = Split(parts(i - 1), ":")
            h(i, 1) = ST_NAME: h(i, 2) = ParseD(CStr(p(0))): h(i, 3) = CDbl(p(1))
        Next i
        hh = h
    End If
    If lgSpec <> "" Then
        parts = Split(lgSpec, ";")
        n = UBound(parts) + 1
        Dim g() As Variant: ReDim g(1 To n, 1 To 7)
        For i = 1 To n
            p = Split(parts(i - 1), ":")
            g(i, 1) = ST_NAME: g(i, 2) = ParseD(CStr(p(0))): g(i, 3) = CStr(p(1)): g(i, 4) = CStr(p(2))
            g(i, 5) = CDbl(p(3)): g(i, 6) = CDbl(p(4)): g(i, 7) = (CStr(p(5)) = "N")
        Next i
        lg = g
    End If
    Dim through As Date, est As Boolean, tot As Double
    tot = HoursAsOfCore(hh, lg, ST_NAME, asOf, ST_YEAR_FROM, ST_YEAR_TO, through, est)
    CaseHoursAsOf = NumText(tot) & "|" & IIf(through = 0, "none", FmtD(through)) & "|" & IIf(est, "yes", "no")
End Function

' "grade:returner:onleave:hist:delta;..." -> the average of the eligible peers' deltas
Private Function CaseLeaveAvg(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim deltas() As Double: ReDim deltas(1 To UBound(parts) + 1)
    Dim i As Long, n As Long, p As Variant
    For i = 0 To UBound(parts)
        p = Split(parts(i), ":")
        If PeerEligible(CDbl(p(0)), (CStr(p(1)) = "Y"), (CStr(p(2)) = "Y"), (CStr(p(3)) = "Y")) Then
            n = n + 1: deltas(n) = CDbl(p(4))
        End If
    Next i
    CaseLeaveAvg = NumText(LeaveAverage(deltas, n))
End Function

' "date;date;..." -> "start..end;..."
Private Function CaseMerge(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim dates() As Date, ids() As String, n As Long, i As Long
    n = UBound(parts) + 1
    ReDim dates(1 To n)
    For i = 1 To n
        dates(i) = ParseD(CStr(parts(i - 1)))
    Next i
    ' MergeRuns expects sorted unique dates (ConvertImports sorts first): sort here too
    Dim j As Long, t As Date
    For i = 2 To n
        t = dates(i): j = i - 1
        Do While j >= 1
            If dates(j) <= t Then Exit Do
            dates(j + 1) = dates(j): j = j - 1
        Loop
        dates(j + 1) = t
    Next i
    Dim runs() As Date, k As Long, s As String
    k = MergeRuns(dates, n, runs)
    For i = 1 To k
        s = s & IIf(i = 1, "", ";") & FmtD(runs(i, 1)) & ".." & FmtD(runs(i, 2))
    Next i
    CaseMerge = s
End Function

' ---------------- helpers ----------------
Private Function ParseD(ByVal s As String) As Date         ' "yyyy-mm-dd" -> date
    ParseD = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
End Function

Private Function FmtD(ByVal d As Date) As String         ' date -> "yyyy-mm-dd"
    FmtD = Format$(d, "yyyy-mm-dd")
End Function

Private Function NumText(ByVal x As Double) As String ' 21 -> "21", 22.75 -> "22.75"
    If x = Int(x) Then
        NumText = CStr(CLng(x))
    Else
        NumText = Replace(CStr(x), ",", ".")
    End If
End Function

Private Sub AddCase(ByVal row As String)
    mN = mN + 1
    If mN = 1 Then ReDim mCases(1 To 256)
    If mN > UBound(mCases) Then ReDim Preserve mCases(1 To UBound(mCases) + 256)
    mCases(mN) = row
End Sub


' ---------------- the cases ----------------
' Fields: suite,case,in1,in2,in3,in4,in5,expected  (no commas inside a field)
Private Sub LoadCases()
' >>> CASES (generated from tests/cases.csv by tests/gen_selftest_cases.py - do not edit by hand)
    AddCase "rotation,2025-12-28,2025-12-28,,,,,B/A"
    AddCase "rotation,2025-12-29,2025-12-29,,,,,D/C"
    AddCase "rotation,2025-12-31,2025-12-31,,,,,A/B"
    AddCase "rotation,2026-01-01,2026-01-01,,,,,A/B"
    AddCase "rotation,2026-01-24,2026-01-24,,,,,B/A"
    AddCase "rotation,2026-01-25,2026-01-25,,,,,B/A"
    AddCase "rotation,2026-02-14,2026-02-14,,,,,D/C"
    AddCase "rotation,2026-03-07,2026-03-07,,,,,A/B"
    AddCase "rotation,2026-05-18,2026-05-18,,,,,D/C"
    AddCase "rotation,2026-07-01,2026-07-01,,,,,B/A"
    AddCase "rotation,2026-09-07,2026-09-07,,,,,D/C"
    AddCase "rotation,2026-10-12,2026-10-12,,,,,B/A"
    AddCase "rotation,2026-10-31,2026-10-31,,,,,B/A"
    AddCase "rotation,2026-11-01,2026-11-01,,,,,B/A"
    AddCase "rotation,2026-12-25,2026-12-25,,,,,B/A"
    AddCase "rotation,2026-12-26,2026-12-26,,,,,B/A"
    AddCase "rotation,2026-12-31,2026-12-31,,,,,A/B"
    AddCase "rotation,2027-01-01,2027-01-01,,,,,C/D"
    AddCase "rotation,2027-03-13,2027-03-13,,,,,D/C"
    AddCase "rotation,2027-07-01,2027-07-01,,,,,B/A"
    AddCase "rotation,2027-12-31,2027-12-31,,,,,C/D"
    AddCase "rotation,2028-02-29,2028-02-29,,,,,B/A"
    AddCase "crewshift,A-2026-10-12,A,2026-10-12,,,,Nights"
    AddCase "crewshift,B-2026-10-12,B,2026-10-12,,,,Days"
    AddCase "crewshift,C-2026-10-12,C,2026-10-12,,,,off"
    AddCase "crewshift,D-2026-10-12,D,2026-10-12,,,,off"
    AddCase "crewshift,A-B-2026-10-12,A/B,2026-10-12,,,,Days"
    AddCase "crewshift,C-D-2026-10-12,C/D,2026-10-12,,,,off"
    AddCase "crewshift,DAYS-2026-10-12,DAYS,2026-10-12,,,,off"
    AddCase "crewshift,A-2026-10-10,A,2026-10-10,,,,off"
    AddCase "crewshift,B-2026-10-10,B,2026-10-10,,,,off"
    AddCase "crewshift,C-2026-10-10,C,2026-10-10,,,,Days"
    AddCase "crewshift,D-2026-10-10,D,2026-10-10,,,,Nights"
    AddCase "crewshift,A-B-2026-10-10,A/B,2026-10-10,,,,off"
    AddCase "crewshift,C-D-2026-10-10,C/D,2026-10-10,,,,Days"
    AddCase "crewshift,DAYS-2026-10-10,DAYS,2026-10-10,,,,off"
    AddCase "crewshift,DAYS-2026-10-13,DAYS,2026-10-13,,,,Days"
    AddCase "crewshift,A-2026-10-31,A,2026-10-31,,,,Nights"
    AddCase "crewshift,B-2026-10-31,B,2026-10-31,,,,Days"
    AddCase "crewshift,C-2026-10-31,C,2026-10-31,,,,off"
    AddCase "crewshift,D-2026-10-31,D,2026-10-31,,,,off"
    AddCase "crewshift,A-B-2026-10-31,A/B,2026-10-31,,,,Days"
    AddCase "crewshift,C-D-2026-10-31,C/D,2026-10-31,,,,off"
    AddCase "crewshift,DAYS-2026-10-31,DAYS,2026-10-31,,,,off"
    AddCase "avail,2026-10-30-D,2026-10-30,D,,,,C"
    AddCase "avail,2026-10-30-N,2026-10-30,N,,,,C/D"
    AddCase "avail,2026-10-31-D,2026-10-31,D,,,,C/D"
    AddCase "avail,2026-10-31-N,2026-10-31,N,,,,C/D"
    AddCase "holiday,newyear-2026,New Year's Day,2026,,,,2026-01-01"
    AddCase "holiday,family-2026,Family Day,2026,,,,2026-02-16"
    AddCase "holiday,goodfri-2026,Good Friday,2026,,,,2026-04-03"
    AddCase "holiday,victoria-2026,Victoria Day,2026,,,,2026-05-18"
    AddCase "holiday,canada-2026,Canada Day,2026,,,,2026-07-01"
    AddCase "holiday,civic-2026,Civic Holiday,2026,,,,2026-08-03"
    AddCase "holiday,labour-2026,Labour Day,2026,,,,2026-09-07"
    AddCase "holiday,thanks-2026,Thanksgiving,2026,,,,2026-10-12"
    AddCase "holiday,remember-2026,Remembrance Day,2026,,,,2026-11-11"
    AddCase "holiday,xmas-2026,Christmas Day,2026,,,,2026-12-25"
    AddCase "holiday,boxing-2026,Boxing Day,2026,,,,2026-12-26"
    AddCase "holiday,newyear-2027,New Year's Day,2027,,,,2027-01-01"
    AddCase "holiday,family-2027,Family Day,2027,,,,2027-02-15"
    AddCase "holiday,goodfri-2027,Good Friday,2027,,,,2027-03-26"
    AddCase "holiday,victoria-2027,Victoria Day,2027,,,,2027-05-24"
    AddCase "holiday,canada-2027,Canada Day,2027,,,,2027-07-01"
    AddCase "holiday,civic-2027,Civic Holiday,2027,,,,2027-08-02"
    AddCase "holiday,labour-2027,Labour Day,2027,,,,2027-09-06"
    AddCase "holiday,thanks-2027,Thanksgiving,2027,,,,2027-10-11"
    AddCase "holiday,remember-2027,Remembrance Day,2027,,,,2027-11-11"
    AddCase "holiday,xmas-2027,Christmas Day,2027,,,,2027-12-25"
    AddCase "holiday,boxing-2027,Boxing Day,2027,,,,2027-12-26"
    AddCase "holoff,thanks-on-day,2026-10-12,,,,,off"
    AddCase "holoff,thanks-next-day,2026-10-13,,,,,on"
    AddCase "holoff,boxing-sat-no-observed,2026-12-26,,,,,on"
    AddCase "holoff,boxing-mon-no-observed,2026-12-28,,,,,on"
    AddCase "holoff,boxing-observed-mon,2026-12-28,Boxing Day|2026|2026-12-28,,,,off"
    AddCase "holoff,boxing-observed-sat-itself,2026-12-26,Boxing Day|2026|2026-12-28,,,,on"
    AddCase "holoff,boxing-observed-other-day,2026-12-29,Boxing Day|2026|2026-12-28,,,,on"
    AddCase "dst,fall-2026,2026,fall,,,,2026-10-31"
    AddCase "dst,spring-2026,2026,spring,,,,2026-03-07"
    AddCase "dst,fall-2027,2027,fall,,,,2027-11-06"
    AddCase "dst,spring-2027,2027,spring,,,,2027-03-13"
    AddCase "dst,fall-2028,2028,fall,,,,2028-11-04"
    AddCase "dst,spring-2028,2028,spring,,,,2028-03-11"
    AddCase "dsthours,fall-2026,2026-10-31,,,,,13"
    AddCase "dsthours,day-after-fall-2026,2026-11-01,,,,,12"
    AddCase "dsthours,day-before-fall-2026,2026-10-30,,,,,12"
    AddCase "dsthours,spring-2027,2027-03-13,,,,,11"
    AddCase "dsthours,spring-2028,2028-03-11,,,,,11"
    AddCase "dsthours,ordinary,2026-10-10,,,,,12"
    AddCase "callorder,ties,Ann:1:100:2010-01-01;Bob:1:100:2008-05-05;Cal:1:80:2015-01-01;Dee:2:10:2000-01-01;Eve:1:100:2008-05-05;Fay:1.5:0:2001-01-01,,,,,Cal/Bob/Eve/Ann/Fay/Dee"
    AddCase "callorder,blank-seniority-last,Gus:1:50:;Hal:1:50:2020-01-01,,,,,Hal/Gus"
    AddCase "callorder,grade-before-hours,Ivy:2:0:2001-01-01;Jon:1:500:2001-01-01,,,,,Jon/Ivy"
    AddCase "callorder,name-case,bea:1:10:2001-01-01;Abe:1:10:2001-01-01,,,,,Abe/bea"
    AddCase "callout,refused-g1,Refused,1,12,,,21"
    AddCase "callout,leftmsg-g2,Left Message,2,12,,,21"
    AddCase "callout,accepted,Accepted,1,12,,,0"
    AddCase "callout,unavailable,Unavailable,1,12,,,0"
    AddCase "callout,refused-g3,Refused,3,12,,,0"
    AddCase "callout,refused-g2.5,Refused,2.5,12,,,0"
    AddCase "callout,refused-dst13,Refused,1,13,,,22.75"
    AddCase "callout,leftmsg-dst11,Left Message,1.5,11,,,19.25"
    AddCase "posting,awarded,award,12,1:1:N:N;2:1:N:N;3:1:Y:Y;4:1:Y:N;5:1:N:N,,,1=21;2=21;3=0;4=0;5=0"
    AddCase "posting,ahead-signed-up,award,12,1:1:Y:N;2:1:N:N;3:1:Y:Y,,,1=0;2=21;3=0"
    AddCase "posting,ahead-grade3,award,12,1:3:N:N;2:2:N:N;3:1:Y:Y,,,1=0;2=21;3=0"
    AddCase "posting,below,award,12,1:1:Y:Y;2:1:N:N,,,1=0;2=0"
    AddCase "posting,two-positions,award,12,1:1:N:N;2:1:Y:Y;3:1:N:N;4:1:Y:N;5:1:Y:Y;6:1:N:N,,,1=21;2=0;3=21;4=0;5=0;6=0"
    AddCase "posting,no-takers,notakers,12,1:1:N:N;2:2.5:N:N;3:1.5:N:N,,,1=21;2=0;3=21"
    AddCase "posting,no-takers-dst,notakers,13,1:1:N:N,,,1=22.75"
    AddCase "posting,out-of-dept,outside,12,1:1:N:N;2:1:Y:N;3:3:N:N,,,1=21;2=0;3=0"
    AddCase "posting,cancel,cancel,12,1:1:N:N;2:1:Y:Y,,,1=0;2=0"
    AddCase "eligibility,on-days,B,2026-10-31,Days,,,scheduled"
    AddCase "eligibility,on-nights,A,2026-10-31,Nights,,,scheduled"
    AddCase "eligibility,on-nights-day-ot,A,2026-10-31,Days,,,scheduled"
    AddCase "eligibility,off-days,C,2026-10-31,Days,,,free"
    AddCase "eligibility,off-nights,D,2026-10-31,Nights,,,free"
    AddCase "eligibility,rest-off-nights,D,2026-10-30,Days,,,rest"
    AddCase "eligibility,rest-days-next-morning,D,2026-11-01,Nights,,,rest"
    AddCase "eligibility,rest-other-direction-ok,D,2026-11-01,Days,,,free"
    AddCase "eligibility,pair-on-days,A/B,2026-10-31,Days,,,scheduled"
    AddCase "eligibility,pair-off,C/D,2026-10-31,Days,,,free"
    AddCase "eligibility,pair-night-ot,C/D,2026-10-31,Nights,,,free"
    AddCase "eligibility,days-weekday,DAYS,2026-10-30,Days,,,scheduled"
    AddCase "eligibility,days-weekday-night-ot,DAYS,2026-10-30,Nights,,,scheduled"
    AddCase "eligibility,days-weekend,DAYS,2026-10-31,Days,,,free"
    AddCase "eligibility,days-holiday,DAYS,2026-10-12,Days,,,free"
    AddCase "eligibility,days-night-before-holiday,DAYS,2026-10-11,Nights,,,free"
    AddCase "eligibility,days-night-before-workday,DAYS,2026-11-01,Nights,,,rest"
    AddCase "eligibility,booked-off-vac,C,2026-10-31,Days,off:VAC,,Vac"
    AddCase "eligibility,booked-off-std-shows-leave,C,2026-10-31,Days,off:STD,,Leave"
    AddCase "eligibility,booked-off-beats-scheduled,B,2026-10-31,Days,off:TRN,,Trn"
    AddCase "eligibility,on-a-switch,C,2026-10-31,Days,ss,,S/S"
    AddCase "eligibility,gave-up-a-shift,C,2026-10-31,Days,ssoff,,S/S off"
    AddCase "eligibility,switch-rest,C,2026-10-31,Days,ssrest,,rest"
    AddCase "prov,day-covered-same-day,2026-10-10,D,2026-10-10,,,covered"
    AddCase "prov,day-not-covered,2026-10-10,D,2026-10-09,,,counts"
    AddCase "prov,night-not-covered-same-day,2026-10-10,N,2026-10-10,,,counts"
    AddCase "prov,night-covered-next-day,2026-10-10,N,2026-10-11,,,covered"
    AddCase "prov,future-day,2026-11-20,D,2026-10-09,,,counts"
    AddCase "provcount,accepted-future,Accepted,Active,2026-10-20,D,2026-10-09,counts"
    AddCase "provcount,awarded-future-night,Awarded,Active,2026-10-31,N,2026-10-30,counts"
    AddCase "provcount,accepted-void,Accepted,VOID,2026-10-20,D,2026-10-09,no"
    AddCase "provcount,refused-never,Refused,Active,2026-10-20,D,2026-10-09,no"
    AddCase "provcount,accepted-covered,Accepted,Active,2026-10-05,D,2026-10-09,no"
    AddCase "provcount,night-ends-on-through,Accepted,Active,2026-10-09,N,2026-10-09,counts"
    AddCase "provcount,awarded-last-year,Awarded,Active,2025-12-20,D,2026-10-09,no"
    AddCase "provcount,accepted-next-year,Accepted,Active,2026-12-27,D,2026-10-09,no"
    AddCase "provcount,accepted-last-day-of-year,Accepted,Active,2026-12-26,N,2026-10-09,counts"
    AddCase "sick,days-same-shift,2026-10-20,D,2026-10-20,D,,blocked"
    AddCase "sick,days-nights-same-day,2026-10-20,D,2026-10-20,N,,blocked"
    AddCase "sick,days-next-day-days,2026-10-20,D,2026-10-21,D,,blocked"
    AddCase "sick,days-next-day-nights-free,2026-10-20,D,2026-10-21,N,,free"
    AddCase "sick,days-day-before-free,2026-10-20,D,2026-10-19,N,,free"
    AddCase "sick,nights-same-shift,2026-10-20,N,2026-10-20,N,,blocked"
    AddCase "sick,nights-next-day-days,2026-10-20,N,2026-10-21,D,,blocked"
    AddCase "sick,nights-next-day-nights,2026-10-20,N,2026-10-21,N,,blocked"
    AddCase "sick,nights-two-days-on-free,2026-10-20,N,2026-10-22,D,,free"
    AddCase "sick,nights-same-day-days-free,2026-10-20,N,2026-10-20,D,,free"
    AddCase "hoursasof,snapshot-before-and-provisional,2026-10-05:100;2026-10-12:118,2026-10-08:Refused:Active:12:21:D;2026-10-10:Accepted:Active:12:0:D;2026-10-20:Accepted:Active:12:0:D,2026-10-11,,,133|2026-10-05|no"
    AddCase "hoursasof,later-snapshot-covers-accepted,2026-10-05:100;2026-10-12:118,2026-10-08:Refused:Active:12:21:D;2026-10-10:Accepted:Active:12:0:D;2026-10-20:Accepted:Active:12:0:D,2026-10-15,,,139|2026-10-12|no"
    AddCase "hoursasof,before-any-snapshot-estimated,2026-10-05:100;2026-10-12:118,2026-10-08:Refused:Active:12:21:D,2026-10-01,,,100|2026-10-05|yes"
    AddCase "hoursasof,void-excluded,2026-10-05:100,2026-10-08:Refused:VOID:12:21:D;2026-10-09:Refused:Active:12:21:D,2026-10-11,,,121|2026-10-05|no"
    AddCase "hoursasof,night-not-covered-on-through,2026-10-10:50,2026-10-10:Accepted:Active:12:0:N,2026-10-11,,,62|2026-10-10|no"
    AddCase "hoursasof,night-covered-day-after,2026-10-11:50,2026-10-10:Accepted:Active:12:0:N,2026-10-12,,,50|2026-10-11|no"
    AddCase "hoursasof,before-year-is-zero,2026-10-05:100,2026-10-08:Refused:Active:12:21:D,2025-12-20,,,0|none|no"
    AddCase "hoursasof,no-history-at-all,,2026-10-08:Refused:Active:12:21:D,2026-10-11,,,21|none|yes"
    AddCase "leavepeer,grade1,1,N,N,Y,,peer"
    AddCase "leavepeer,grade2,2,N,N,Y,,peer"
    AddCase "leavepeer,grade2.5,2.5,N,N,Y,,excluded"
    AddCase "leavepeer,grade3,3,N,N,Y,,excluded"
    AddCase "leavepeer,returner,1,Y,N,Y,,excluded"
    AddCase "leavepeer,on-leave,1,N,Y,Y,,excluded"
    AddCase "leavepeer,no-history,1,N,N,N,,excluded"
    AddCase "leaveavg,mixed-set,1:N:N:Y:412.6;1:N:Y:Y:900;3:N:N:Y:50;2:N:N:N:10;1.5:N:N:Y:380;1:Y:N:Y:0,,,,,396.3"
    AddCase "leaveavg,no-peers,3:N:N:Y:50;1:Y:N:Y:0,,,,,0"
    AddCase "leaveavg,rounding,1:N:N:Y:100;1:N:N:Y:101;1:N:N:Y:101,,,,,100.7"
    AddCase "merge,five-day-run,2026-11-02;2026-11-03;2026-11-04;2026-11-05;2026-11-06,,,,,2026-11-02..2026-11-06"
    AddCase "merge,two-runs,2026-11-02;2026-11-03;2026-11-05;2026-11-06;2026-11-07,,,,,2026-11-02..2026-11-03;2026-11-05..2026-11-07"
    AddCase "merge,unsorted-with-duplicate,2026-11-06;2026-11-05;2026-11-05;2026-11-09,,,,,2026-11-05..2026-11-06;2026-11-09..2026-11-09"
' <<< CASES
End Sub
