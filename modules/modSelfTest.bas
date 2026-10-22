Attribute VB_Name = "modSelfTest"
Option Explicit

' ====================================================================
'  SELF TEST v1 (v21)
'  Alt+F8 > SelfTest runs known-answer cases against the pure rules:
'  rotation, holidays, DST nights, call order, charging, eligibility
'  the provisional-OT date logic, leave top-ups, attendance counters and
'  the year-end pieces (pay-period year, ID numbering, integrity helpers).
'  It reports pass / fail per case.
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

Public Const ST_VERSION As String = "Self test v26.5.5 (Oct 21)"

' provcount cases assume this roster year (26 periods from Dec 28, 2025)
Private Const ST_YEAR_FROM As Date = #12/28/2025#
Private Const ST_YEAR_TO As Date = #12/26/2026#
Private Const ST_NAME As String = "Test Person"

Private mCases() As String
Private mN As Long
Private mHits As Long               ' the rename helpers count their replacements here


' ---------------- entry ----------------
' v24.4: the run itself, shared by SelfTest (Alt+F8) and SelfTestOK (the Install panel's step). True = every case passed.
Private Function SelfTestRun() As Boolean
    Dim i As Long, f() As String, actual As String, expected As String
    Dim nPass As Long, nFail As Long, fails As String, suite As String
    Dim bySuite As Object: Set bySuite = CreateObject("Scripting.Dictionary")
    On Error GoTo Finally                      ' v24.3.1: from the first line (LoadCases was outside it)
    mN = 0
    LoadCases
    If mN = 0 Then MsgBox "No test cases found in modSelfTest.", vbExclamation, "Call Out Roster": Exit Function

    SetTestCrewHistory Empty                ' v24: no case reads the Crew History sheet
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
        If StrComp(Trim$(actual), Trim$(expected), vbBinaryCompare) = 0 Then     ' v25 (T13): case counts
            nPass = nPass + 1
        Else
            nFail = nFail + 1
            If nFail <= 15 Then fails = fails & "  " & suite & " / " & f(1) & ": got """ & actual & """, expected """ & expected & """" & vbNewLine
            Debug.Print "SelfTest FAIL " & mCases(i) & " -> " & actual
        End If
        bySuite(suite) = bySuite(suite) & IIf(StrComp(Trim$(actual), Trim$(expected), vbBinaryCompare) = 0, ".", "F")
    Next i
    If nFail > 15 Then fails = fails & "  ...and " & (nFail - 15) & " more (see the Immediate window)" & vbNewLine

Finally:
    Dim em As String
    If Err.Number <> 0 Then
        em = "Stopped by an error: " & Err.Description & vbNewLine
        LogError "SelfTest", Err.Number, Err.Description     ' v24.3.1 (LogError clears Err: em is read first)
    End If
    On Error Resume Next
    ClearTestFixtures
    ClearTestSwitches
    ClearTestHolidays
    ClearTestCrewHistory
    Dim k As Variant, summary As String, nSuites As Long
    For Each k In bySuite.Keys          ' only the suites with a failure are listed (the box holds about 1,000 characters)
        nSuites = nSuites + 1
        If InStr(bySuite(k), "F") > 0 Then
            summary = summary & "  " & k & ": " & Len(bySuite(k)) - Len(Replace(bySuite(k), "F", "")) & " failed of " & Len(bySuite(k)) & vbNewLine
        End If
    Next k
    If summary = "" Then summary = "  all " & nSuites & " suites passed" & vbNewLine
    Dim msg As String
    msg = em & "SelfTest: " & nPass & " passed, " & nFail & " failed of " & mN & " cases." & vbNewLine & vbNewLine & summary
    If fails <> "" Then msg = msg & vbNewLine & "Failed:" & vbNewLine & fails
    msg = msg & vbNewLine & "Nothing was written or saved. (" & ST_VERSION & ")"
    MsgBox msg, IIf(nFail = 0 And em = "", vbInformation, vbExclamation), "Self test"
    SelfTestRun = (nFail = 0 And em = "" And mN > 0)
End Function

Public Sub SelfTest()
    On Error GoTo Oops
    SelfTestRun
    Exit Sub
Oops:
    ReportError "SelfTest"
End Sub

' v24.4: the same run as a result, for the Install panel
Public Function SelfTestOK() As Boolean
    On Error GoTo Oops
    SelfTestOK = SelfTestRun()
    Exit Function
Oops:
    ReportError "SelfTestOK"
End Function


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
        Case "typeddate"                           ' v25: text, today -> the date, or the hint word
            Dim tdD As Date, tdW As String
            tdW = TypedDateCore(Replace(i1, "_", " "), ParseD(i2), tdD)
            RunCase = IIf(tdW = "", FmtD(tdD), tdW)
        Case "trimkind"                            ' v25: start, end, first shift, last shift -> which ends are days off
            RunCase = TrimKind(ParseD(i1), ParseD(i2), ParseD(i3), ParseD(i4))
            If RunCase = "" Then RunCase = "none"
        Case "partialhrs"                          ' v25: typed hours, posting hours -> whole / the hours / bad
            Dim phH As Double, phW As String
            phW = PartialHoursValue(i1, CDbl(i2), phH)
            RunCase = IIf(phW <> "", "bad", IIf(phH = 0, "whole", NumText(phH)))
        Case "postclosed"                          ' v25 (P6): closing date, as-of -> open / closed
            RunCase = IIf(PostingClosed(ParseD(i1), ParseD(i2)), "closed", "open")
        Case "crewsoff"                            ' v25
            RunCase = CrewsOff(ParseD(i1))
        Case "shifthours"                          ' v25
            RunCase = NumText(DefaultShiftHours(ParseD(i1), (i2 = "night"), USUAL_SHIFT_HOURS))
        Case "shiftdate"                           ' v26.5.5: the date, the clock time hh:mm -> the date a shift logged then belongs to
            RunCase = FmtD(DefaultShiftDate(ParseD(i1) + TimeSerial(CLng(Left$(i2, 2)), CLng(Mid$(i2, 4, 2)), 0)))
        Case "accepted"                            ' v25
            RunCase = NumText(AcceptedHours(CDbl(i1), CDbl(i2)))
        Case "callorder"
            RunCase = CaseCallOrder(i1)
        Case "listinsert"                          ' v24.3.1: frozen list in its order, newcomer -> index 1..n+1
            RunCase = CaseListInsert(i1, i2)
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
        Case "attend"
            RunCase = CaseAttend(i1, i2, i3, i4)
        Case "year"
            RunCase = FmtD(FirstPeriodStart(ParseD(i1), CLng(i2))) & "|" & PeriodsInYear(ParseD(i1), CLng(i2)) & "|" & _
                      RosterYearOf(FirstPeriodStart(ParseD(i1), CLng(i2)))
        Case "nextid"
            RunCase = CStr(NextNumber(CLng(i1), CLng(i2)))
        Case "cell"
            RunCase = CaseCell(i1, i2, i3)
        Case "dups"
            RunCase = CaseDups(i1)
        Case "archive"
            RunCase = IIf(GoesToArchive(CaseValue(i1), ParseD(i2)), "archive", "keep")
        Case "dryname"
            RunCase = IIf(IsDryRunName(i1), "yes", "no")
        Case "backupname"                          ' v24.3.1: a backup, FINAL or pre-change copy by its name
            RunCase = IIf(IsBackupCopyName(i1), "yes", "no")
        Case "viewmode"                            ' v25.3: login, ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS (v25.5; ; for ,) -> desk / admin / supervisor / user
            RunCase = ViewModeFor(i1, Replace(i2, ";", ","), Replace(i3, ";", ","), Replace(i4, ";", ","))
        Case "backupdate"                          ' v25: file name -> the date a daily backup carries / none
            Dim bdD As Date: bdD = BackupFileDate(i1)
            RunCase = IIf(bdD = 0, "none", FmtD(bdD))
        Case "deptcode"                            ' v25: grade code (_ = space) -> its numeric part (999 = none)
            RunCase = NumText(DeptCodeValue(Replace(i1, "_", " ")))
        Case "hourstext"                           ' v25: hours -> the text every message shows (no trailing point)
            RunCase = HoursText(Val(i1))
        Case "payback"                             ' v26: days owed, warn days -> the Needs Attention status
            RunCase = OwedStatusText(CLng(Val(i1)), CLng(Val(i2)))
        Case "idmark"                              ' v26 (I8): Ref IDs joined by | -> highest number per prefix, or none
            RunCase = RefPrefixMaxText(i1)
            If RunCase = "" Then RunCase = "none"
        Case "refthrough"                          ' v26.5: period start, the current period's start -> counts / waits
            RunCase = IIf(RefusedCounts(ParseD(i1), ParseD(i2)), "counts", "waits")
        Case "minstaff"                            ' v26: scheduled, off, covered, minimum -> ok / SHORT staffed/minimum
            RunCase = StaffShort(CLng(Val(i1)), CLng(Val(i2)), CLng(Val(i3)), CLng(Val(i4)))
            If RunCase = "" Then RunCase = "ok"
        Case "practice"                            ' v26.5.5: a file name -> the date and time a practice copy carries (yyyy-mm-dd hh:mm) / none
            Dim prD As Date: prD = PracticeCopyDate(i1)
            RunCase = IIf(prD = 0, "none", Format$(prD, "yyyy-mm-dd hh:mm"))
        Case "nextstep"                            ' v26.5.5: accepted, partial hours, left messages, admin Y/N, admin name -> the last line (; for ,)
            RunCase = Replace(NextStepText(CLng(Val(i1)), CDbl(Val(i2)), CLng(Val(i3)), (i4 = "Y"), i5), ",", ";")
        Case "fatblock"                            ' v26.5.5: hours on the shift before the touching one, the touching shift, the next, the one after -> block|hours to next
            Dim fbW As Double, fbN As Double
            FatigueBlockCore CDbl(Val(i1)), CDbl(Val(i2)), CDbl(Val(i3)), CDbl(Val(i4)), fbW, fbN
            RunCase = NumText(fbW) & "|" & NumText(fbN)
        Case "postcap"                             ' v26.5.5: posting hours, the 16 h max, max known Y/N -> the hours charged on
            RunCase = NumText(CappedPostingHours(CDbl(Val(i1)), CDbl(Val(i2)), (i3 = "Y")))
        Case "fatigue"                             ' v26.5.5: worked before, hours offered, hours to the next work -> ok / 16 h rule | the max
            RunCase = FatigueProblem(CDbl(Val(i1)), CDbl(Val(i2)), CDbl(Val(i3)))
            If RunCase = "" Then RunCase = "ok"
            RunCase = RunCase & "|" & NumText(FatigueMaxHours(CDbl(Val(i1)), CDbl(Val(i3))))
        Case "chargewhy"                           ' v26: type, outcome, call order, hours|charged, code -> the sentence
            Dim hc As Variant: hc = Split(i4 & "|", "|")
            RunCase = Replace(ChargeWhy(i1, i2, CLng(Val(i3)), Val(hc(0)), Val(hc(1)), i5), ",", ";")   ' ; for , (no commas in a case)
        Case "prechange"                           ' v26: file name, roster file name -> the date a pre-change copy carries / none
            Dim pcD As Date: pcD = PreChangeCopyDate(i1, i2)
            RunCase = IIf(pcD = 0, "none", FmtD(pcD))
        Case "typetag"                             ' v25: absence code -> calendar tag / what members get
            RunCase = TypeTag(i1) & "/" & PublicType(i1)
        Case "planid"                              ' v25: an Exempt note -> the planned move's ID it carries / none
            RunCase = PlanIdIn(i1)
            If RunCase = "" Then RunCase = "none"
        Case "namekey"                             ' v25: a Kronos or roster name (_ = space) -> the match key (, shown as /)
            RunCase = Replace(NameKey(Replace(i1, "_", " ")), ",", "/")
        Case "barlist"                             ' v25: in|trim, list (_ = space), item -> yes / no, or the cleaned list
            If i1 = "in" Then
                RunCase = IIf(InBarList(Replace(i2, "_", " "), Replace(i3, "_", " ")), "yes", "no")
            Else
                RunCase = TrimBarList(Replace(i2, "_", " "))
                If RunCase = "" Then RunCase = "empty"
            End If
        Case "matchname"                           ' v25: old-calendar text (_ = space), roster 'Last/First|...' -> name / ? / none
            Dim mnList As Variant: mnList = Split(Replace(Replace(i2, "/", ", "), "_", " "), "|")
            RunCase = Replace(MatchImportName(Replace(i1, "_", " "), mnList), ", ", "/")
            If RunCase = "" Then RunCase = "none"
        Case "attendstatus"                        ' v25: entitled Y|N, paid left, occurrences, wait-until -> the phrase (, as .)
            Dim asSt As AttendanceStatus
            asSt.Entitled = (i1 = "Y"): asSt.PaidLeft = CLng(i2): asSt.Occurrences = CLng(i3)
            If i4 <> "" Then asSt.WaitUntil = ParseD(i4)
            RunCase = Replace(AttendPublicStatus(asSt), ",", ".")
        Case "totalsdiff"                          ' v25: before, after as 'name=G|I|periods;...', skip -> the lines (; between, / for ', ')
            RunCase = CaseTotalsDiff(i1, i2, i3)
        Case "snapyear"
            RunCase = FmtD(SnapshotThroughForYear(ParseD(i1), CLng(i2), CLng(i3), ParseD(i4)))
        Case "signups"
            RunCase = CaseSignups(i1, i2)
        Case "withdrawn"                           ' v26 (M3): 'name:S|W:hours:yyyy-mm-dd-hh:login;...' -> name=login[!];... / none
            RunCase = CaseWithdrawn(i1)
        Case "sgsuspect"                           ' v27 (item 14): the file's stamp, its time on disk (yyyy-mm-dd-hh-nn; 0 = unknown) -> suspect / ok
            RunCase = IIf(SignupSuspect(ParseStampMin(i1), ParseStampMin(i2)), "suspect", "ok")
        Case "sgmix"                               ' v27 (item 14): 'name:S|W:hours:yyyy-mm-dd-hh:login;...' -> the reduced sign-ups, name[!] (! = another login for that name) / none
            RunCase = CaseSignupMix(i1)
        Case "calloffer"
            RunCase = CaseCallOffer(CDbl(i1), i2)
        Case "attendocc"
            RunCase = CaseAttend(i1, i2, i3, i4, True)
        Case "refsplit"
            RunCase = CaseRefSplit(i1)
        Case "awardorder"
            RunCase = CaseAwardOrder(CDbl(i1), CLng(i2), i3, False)
        Case "awardcover"
            RunCase = CaseAwardOrder(CDbl(i1), CLng(i2), i3, True)
        Case "newhire"
            RunCase = CaseNewHire(i1)
        Case "unitmove"
            RunCase = CaseUnitMove(CDbl(i1), i2)
        Case "plancode"                            ' v24.3: plans, unit asked, date, roster code -> code / none
            RunCase = PlannedCodeCore(CasePlans(i1), ST_NAME, i2, ParseD(i3), i4)
            If RunCase = "" Then RunCase = "none"
        Case "planhours"                           ' own total, others 'grade:mover:total;...' -> listed Total
            RunCase = CasePlanHours(CDbl(i1), i2)
        Case "renamelist"
            RunCase = RenameInList(i1, i2, i3, mHits)
        Case "nametoken"
            RunCase = ReplaceNameToken(i1, i2, i3, mHits)
        Case "crewon"
            RunCase = CrewOnCore(CaseCrewHist(i1), ST_NAME, ParseD(i2), i3)
            If RunCase = "" Then RunCase = "none"
        Case "crewsched"
            RunCase = CrewShiftOn(CrewOnCore(CaseCrewHist(i1), ST_NAME, ParseD(i2), ""), ParseD(i2))
            If RunCase = "" Then RunCase = "off"
        Case "eligcrew"
            RunCase = CaseEligCrew(i1, ParseD(i2), i3)
        Case "privacy"                             ' v24.2: one published value -> ok / email / phone / number
            RunCase = PrivateValueKind(CaseValue(i1))
            If RunCase = "" Then RunCase = "ok"
        Case "attendpriv"                          ' count|status, value -> ok / stop
            If i1 = "count" Then
                RunCase = IIf(AttendCountOk(CaseValue(i2)), "ok", "stop")
            Else
                RunCase = IIf(AttendStatusOk(Mid$(i2, 3)), "ok", "stop")
            End If
        Case "leavewords"                          ' v24.3.1: published free text -> the offending word / ok
            RunCase = LeaveWordsIn(i1)
            If RunCase = "" Then RunCase = "ok"
        Case "paytext"                             ' a cell value -> the Payroll # as text / blank
            RunCase = PayrollText(CaseValue(i1))
            If RunCase = "" Then RunCase = "blank"
        Case "krhdr"                               ' v24.3.1: heading cell, wanted heading -> yes / no
            RunCase = IIf(IsHeaderMatch(i1, i2), "yes", "no")
        Case "payhdr"                              ' header text -> yes / no
            RunCase = IIf(IsPayrollHeader(i1), "yes", "no")
        Case "payroll"                             ' roster number, in Kronos Y|N, names agree Y|N -> name / number / mismatch
            RunCase = PayrollDecision(i1, (i2 = "Y"), (i3 = "Y"))
        Case "kmult"                               ' v27: a Kronos pay code -> its multiplier / none (the settings' table)
            Dim km As Double: km = PremiumMultiplier(i1, PREMIUM_CODES, PREMIUM_MULTIPLIERS)
            RunCase = IIf(km < 0, "none", NumText(km))
        Case "kweight"                             ' v27: "code=hours;code=hours" -> the weighted whole-hour total | the codes not counted
            RunCase = WeightedTotal(i1, PREMIUM_CODES, PREMIUM_MULTIPLIERS)
        Case "kperiod"                             ' v27: the report's Time Period line -> yyyy-mm-dd|yyyy-mm-dd or the problem
            Dim kf As Date, kt As Date, kp As String
            kp = TimePeriodCore(i1, kf, kt)
            RunCase = IIf(kp <> "", kp, Format$(kf, "yyyy-mm-dd") & "|" & Format$(kt, "yyyy-mm-dd"))
        Case "kkind"                               ' v27: from, to, the roster year's first and last day -> day / cumulative / range / outside
            RunCase = ReportKindCore(ParseD(i1), ParseD(i2), ParseD(i3), ParseD(i4))
        Case "firstref"                            ' v27 (item 8): a block of text -> its first Ref ID, or "" as "-"
            RunCase = FirstRefIn(i1)
            If RunCase = "" Then RunCase = "-"
        Case "reconcile"                           ' v27: the paper's names (;), the roster's names (;) -> positions that differ|not listed|not on paper
            RunCase = ReconcileDiffCore(i1, i2)
        Case "squarter"                            ' v27 (Stats): a date, the roster year's Jan 1 -> Q1 .. Q4 (late December before it = Q1)
            RunCase = StatsQuarter(ParseD(i1), ParseD(i2))
        Case "sbucket"                             ' v27 (Stats): a log row's Type, Outcome, Status -> what it counts as
            RunCase = StatsBucket(i1, i2, i3)
        Case "sband"                               ' v27 (Stats): an hour 0..23 -> the time-of-day band
            RunCase = TimeOfDayBand(CLng(Val(i1)))
        Case "sspread"                             ' v27 (Stats): "h;h;h" -> min|max|mean|sd|gap
            RunCase = SpreadCore(i1)
        Case "bkeep"                               ' v27 (backups): daily|hourly, the file's date, today, period start Y/N -> keep / drop
            RunCase = IIf(BackupKeepCore(i1, ParseD(i2), ParseD(i3), (i4 = "Y")), "keep", "drop")
        Case "bhourly"                             ' v27 (backups): a file name -> the hourly backup's day and hour / none
            Dim bh As Date: bh = HourlyBackupDate(i1)
            RunCase = IIf(bh = 0, "none", Format$(bh, "yyyy-mm-dd hh"))
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

' "name:grade:hours:seniority;..." -> recs(1..n) in the order given. Returns n.
Private Function CasePeople(ByVal spec As String, ByRef recs() As EmpRec) As Long
    Dim recsTxt As Variant: recsTxt = Split(spec, ";")
    Dim n As Long: n = UBound(recsTxt) + 1
    ReDim recs(1 To n)
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
    CasePeople = n
End Function

' a frozen list (kept in the order given) and one newcomer -> where they go
Private Function CaseListInsert(ByVal frozenSpec As String, ByVal newSpec As String) As String
    Dim frozen() As EmpRec, one() As EmpRec, n As Long
    n = CasePeople(frozenSpec, frozen)
    CasePeople newSpec, one
    CaseListInsert = CStr(ListInsertIndex(frozen, n, one(1)))
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

' mode (award / notakers / outside / cancel), hours, "rank:grade:signed:awarded[:X];..." (X = exempt)
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
        If UBound(p) >= 4 Then
            If CStr(p(4)) = "X" Then lst(i).ExemptNote = EXEMPT_TAG & "test"      ' v24: moved unit or left
        End If
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
    ElseIf fixture = "ss" Or fixture = "ssoff" Or fixture = "ssrest" Or fixture = "ssgaveup" Then
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
            If fixture = "ssgaveup" Then                 ' v25 (D-C3): they GAVE UP the adjoining shift
                s(1, SW_WORKS) = "Someone Else": s(1, SW_OFF) = ST_NAME
            Else
                s(1, SW_WORKS) = ST_NAME: s(1, SW_OFF) = "Someone Else"
            End If
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

' hours, positions, "rank:signup:awarded;..." (signup = N not signed up, F whole shift, or the hours)
' -> the award-order problem (or ok); with cover: "covered|uncovered"
Private Function CaseAwardOrder(ByVal hours As Double, ByVal positions As Long, ByVal spec As String, ByVal cover As Boolean) As String
    Dim items As Variant: items = Split(spec, ";")
    Dim n As Long: n = UBound(items) + 1
    Dim lst() As PListRec: ReDim lst(1 To n)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(items(i - 1), ":")
        lst(i).Rank = CLng(p(0))
        lst(i).Grade = 1
        lst(i).SignedUp = (CStr(p(1)) <> "N")
        If CStr(p(1)) <> "N" And CStr(p(1)) <> "F" Then lst(i).SignupHrs = CDbl(p(1))
        lst(i).Awarded = (CStr(p(2)) = "Y")
    Next i
    Dim pr As PostRec: pr.Hours = hours: pr.Positions = positions
    If cover Then
        Dim covered As Double, uncovered As Double
        AwardCoverage pr, lst, n, covered, uncovered
        CaseAwardOrder = NumText(covered) & "|" & NumText(uncovered)
    Else
        CaseAwardOrder = AwardOrderProblem(pr, lst, n)
        If CaseAwardOrder = "" Then CaseAwardOrder = "ok"
    End If
End Function

' "name:S|W:hours:yyyy-mm-dd-hh;...", closing "yyyy-mm-dd-hh" (or "") -> "name=F|hours[L];..." / none
Private Function CaseSignups(ByVal spec As String, ByVal closeTxt As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim rows() As Variant: ReDim rows(1 To n, 1 To SG_N)
    Dim i As Long, p As Variant, closesAt As Date, red() As Variant, m As Long
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        rows(i, SG_EMP) = CStr(p(0))
        rows(i, SG_ACTION) = IIf(CStr(p(1)) = "S", "signup", "withdraw")
        rows(i, SG_HOURS) = CDbl(p(2))
        rows(i, SG_STAMP) = ParseD(CStr(p(3))) + TimeSerial(CLng(Mid$(CStr(p(3)), 12, 2)), 0, 0)
    Next i
    If closeTxt <> "" Then closesAt = ParseD(closeTxt) + TimeSerial(CLng(Mid$(closeTxt, 12, 2)), 0, 0)
    Dim v As Variant: v = rows
    m = ReduceSignups(v, closesAt, red)
    For i = 1 To m
        CaseSignups = CaseSignups & IIf(i = 1, "", ";") & CStr(red(i, SR_EMP)) & "=" & _
                      IIf(NumOr0(red(i, SR_HOURS)) = 0, "F", NumText(NumOr0(red(i, SR_HOURS)))) & IIf(CBool(red(i, SR_LATE)), "L", "")
    Next i
    If CaseSignups = "" Then CaseSignups = "none"
End Function

' v26 (M3): "name:S|W:hours:yyyy-mm-dd-hh:login;..." -> "name=login" per newest-withdraw, "!" when a sign-up
' before it came from another login; "none"
Private Function CaseWithdrawn(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim rows() As Variant: ReDim rows(1 To n, 1 To SG_N)
    Dim i As Long, p As Variant, wd() As Variant, m As Long
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        rows(i, SG_EMP) = CStr(p(0))
        rows(i, SG_ACTION) = IIf(CStr(p(1)) = "S", "signup", "withdraw")
        rows(i, SG_HOURS) = CDbl(p(2))
        rows(i, SG_STAMP) = ParseD(CStr(p(3))) + TimeSerial(CLng(Mid$(CStr(p(3)), 12, 2)), 0, 0)
        If UBound(p) >= 4 Then rows(i, SG_LOGIN) = CStr(p(4)) Else rows(i, SG_LOGIN) = ""
    Next i
    Dim v As Variant: v = rows
    m = WithdrawnSignups(v, wd)
    For i = 1 To m
        CaseWithdrawn = CaseWithdrawn & IIf(i = 1, "", ";") & CStr(wd(i, WD_EMP)) & "=" & CStr(wd(i, WD_LOGIN)) & IIf(CBool(wd(i, WD_MISMATCH)), "!", "")
    Next i
    If CaseWithdrawn = "" Then CaseWithdrawn = "none"
End Function

' v27 (item 14): "yyyy-mm-dd-hh-nn" -> the date and time; "0" / "" -> 0
Private Function ParseStampMin(ByVal s As String) As Date
    s = Trim$(s)
    If Len(s) < 16 Then Exit Function
    ParseStampMin = ParseD(Left$(s, 10)) + TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), 0)
End Function

' v27 (item 14): "name:S|W:hours:yyyy-mm-dd-hh:login;..." -> "name[!];..." for the reduced sign-ups ("!" = files from
' more than one login for that name); "none"
Private Function CaseSignupMix(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim rows() As Variant: ReDim rows(1 To n, 1 To SG_N)
    Dim i As Long, p As Variant, red() As Variant, m As Long
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        rows(i, SG_EMP) = CStr(p(0))
        rows(i, SG_ACTION) = IIf(CStr(p(1)) = "S", "signup", "withdraw")
        rows(i, SG_HOURS) = CDbl(p(2))
        rows(i, SG_STAMP) = ParseD(CStr(p(3))) + TimeSerial(CLng(Mid$(CStr(p(3)), 12, 2)), 0, 0)
        If UBound(p) >= 4 Then rows(i, SG_LOGIN) = CStr(p(4)) Else rows(i, SG_LOGIN) = ""
    Next i
    Dim v As Variant: v = rows
    m = ReduceSignups(v, 0, red)
    For i = 1 To m
        CaseSignupMix = CaseSignupMix & IIf(i = 1, "", ";") & CStr(red(i, SR_EMP)) & IIf(CBool(red(i, SR_LOGINMIX)), "!", "")
    Next i
    If CaseSignupMix = "" Then CaseSignupMix = "none"
End Function

' hours, "outcome:taken;..." in call order (everyone grade 1) -> "offered=charged;..."
Private Function CaseCallOffer(ByVal hours As Double, ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim oc() As String, taken() As Double, offers() As Double
    ReDim oc(1 To n): ReDim taken(1 To n)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        oc(i) = CStr(p(0)): taken(i) = CDbl(p(1))
    Next i
    CalloutOffers hours, oc, taken, n, offers
    For i = 1 To n
        CaseCallOffer = CaseCallOffer & IIf(i = 1, "", ";") & NumText(offers(i)) & "=" & NumText(CalloutCharge(oc(i), 1, offers(i)))
    Next i
End Function

' "type:hours:V|A:Y|N;..." (V = void, Y = period inside the year) -> "callout|posting|other"
Private Function CaseRefSplit(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim lg() As Variant: ReDim lg(1 To n, 1 To L_NCOLS)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        lg(i, L_EMP) = ST_NAME
        lg(i, L_TYPE) = CStr(p(0))
        lg(i, L_CHARGED) = CDbl(p(1))
        lg(i, L_STATUS) = IIf(CStr(p(2)) = "V", "VOID", "")
        lg(i, L_PERIOD) = IIf(CStr(p(3)) = "Y", ST_YEAR_FROM + 14, ST_YEAR_FROM - 14)
    Next i
    Dim v As Variant: v = lg
    Dim m As Object: Set m = SplitTotals(v, CLng(ST_YEAR_FROM), CLng(ST_YEAR_TO) - 13)
    Dim c As Double, ps As Double, o As Double
    RefusedSplitOf m, ST_NAME, c, ps, o
    CaseRefSplit = NumText(c) & "|" & NumText(ps) & "|" & NumText(o)
End Function

' "grade:onleave:total;..." -> a new hire's starting hours
Private Function CaseNewHire(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim grades() As Double, onLeave() As Boolean, totals() As Double
    ReDim grades(1 To n): ReDim onLeave(1 To n): ReDim totals(1 To n)
    Dim i As Long, p As Variant, nP As Long
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        grades(i) = CDbl(p(0)): onLeave(i) = (CStr(p(1)) = "Y"): totals(i) = CDbl(p(2))
    Next i
    CaseNewHire = NumText(NewHireAverage(grades, onLeave, totals, n, nP))
End Function

' own total, "grade:mover:total;..." -> hours added on a unit move
Private Function CaseUnitMove(ByVal own As Double, ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim grades() As Double, isMover() As Boolean, totals() As Double
    ReDim grades(1 To n): ReDim isMover(1 To n): ReDim totals(1 To n)
    Dim i As Long, p As Variant, found As Boolean, top As Double
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        grades(i) = CDbl(p(0)): isMover(i) = (CStr(p(1)) = "Y"): totals(i) = CDbl(p(2))
    Next i
    top = HighestPeerTotal(grades, isMover, totals, n, found)
    If found Then CaseUnitMove = NumText(UnitMoveTopUp(own, top)) Else CaseUnitMove = "0"
End Function

' "unit:code:clear|clear:yyyy-mm-dd:status[:other];..." -> rows in the tblPlannedMoves
' layout for ST_NAME ("other" = someone else's move, "lower" = the name in lower case). Empty when blank.
Private Function CasePlans(ByVal spec As String) As Variant
    If spec = "" Then Exit Function
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim h() As Variant: ReDim h(1 To n, 1 To PM_N)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        h(i, PM_ID) = "PM-" & Format$(i, "0000")
        h(i, PM_EMP) = ST_NAME
        If UBound(p) >= 5 Then
            If CStr(p(5)) = "other" Then h(i, PM_EMP) = "Someone Else"
            If CStr(p(5)) = "lower" Then h(i, PM_EMP) = LCase$(ST_NAME)
        End If
        h(i, PM_UNIT) = CStr(p(0))
        h(i, PM_CODE) = CStr(p(1))
        h(i, PM_CLEAR) = CStr(p(2))
        h(i, PM_FROM) = ParseD(CStr(p(3)))
        h(i, PM_STATUS) = CStr(p(4))
    Next i
    CasePlans = h
End Function

' own total, "grade:mover:total;..." -> the Total a planned mover is listed at
Private Function CasePlanHours(ByVal own As Double, ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim grades() As Double, isMover() As Boolean, totals() As Double
    ReDim grades(1 To n): ReDim isMover(1 To n): ReDim totals(1 To n)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        grades(i) = CDbl(p(0)): isMover(i) = (CStr(p(1)) = "Y"): totals(i) = CDbl(p(2))
    Next i
    CasePlanHours = NumText(PlannedListHours(own, grades, isMover, totals, n))
End Function

' "crew:yyyy-mm-dd;..." -> rows in the tblCrewHistory layout for ST_NAME (Empty when blank)
Private Function CaseCrewHist(ByVal spec As String) As Variant
    If spec = "" Then Exit Function
    Dim parts As Variant: parts = Split(spec, ";")
    Dim n As Long: n = UBound(parts) + 1
    Dim h() As Variant: ReDim h(1 To n, 1 To CH_N)
    Dim i As Long, p As Variant
    For i = 1 To n
        p = Split(parts(i - 1), ":")
        h(i, CH_EMP) = ST_NAME
        h(i, CH_CREW) = CStr(p(0))
        h(i, CH_FROM) = ParseD(CStr(p(1)))
    Next i
    CaseCrewHist = h
End Function

' history, date, shift -> the rotation part of "who is free" with the crew
' taken from the history on each of the three days
Private Function CaseEligCrew(ByVal spec As String, ByVal d As Date, ByVal shiftLbl As String) As String
    Dim h As Variant: h = CaseCrewHist(spec)
    CaseEligCrew = RotationFreeReason3(CrewOnCore(h, ST_NAME, d - 1, ""), CrewOnCore(h, ST_NAME, d, ""), _
                                       CrewOnCore(h, ST_NAME, d + 1, ""), d, shiftLbl)
    If CaseEligCrew = "" Then CaseEligCrew = "free"
End Function

' ---------------- helpers ----------------
' v25: TotalsDifference on two packed maps 'name=12|21|21;name2=...' -> its report squeezed to one line
' ("  k: was Total 12, Refused 21, periods 21, now ..." lines become "k: was Total 12/Refused 21/periods 21/now ...;")
Private Function CaseTotalsDiff(ByVal beforeSpec As String, ByVal afterSpec As String, ByVal skipName As String) As String
    Dim s As String
    s = TotalsDifference(CaseMap(beforeSpec), CaseMap(afterSpec), skipName)
    s = Replace(s, vbNewLine & "  ", ";")
    s = Replace(s, vbNewLine, "")
    s = Replace(Trim$(s), ", ", "/")
    CaseTotalsDiff = IIf(s = "", "same", s)
End Function

Private Function CaseMap(ByVal spec As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim part As Variant, p As Long
    For Each part In Split(spec, ";")
        p = InStr(part, "=")
        If p > 1 Then d(Trim$(Left$(part, p - 1))) = Mid$(part, p + 1)
    Next part
    Set CaseMap = d
End Function

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

' crew, sick "date:D|N;...", personal "date;...", as-of -> used/paid/occ/per/ENT|WAIT date/PAID|UNPAID|-
Private Function CaseAttend(ByVal crew As String, ByVal sickSpec As String, ByVal perSpec As String, ByVal asOfTxt As String, _
                            Optional ByVal occOnly As Boolean = False) As String
    Dim sd() As Date, sn() As Boolean, pd() As Date, nS As Long, nP As Long, parts As Variant, q As Variant, i As Long
    ReDim sd(1 To 1): ReDim sn(1 To 1): ReDim pd(1 To 1)
    If sickSpec <> "" Then
        parts = Split(sickSpec, ";")
        nS = UBound(parts) + 1
        ReDim sd(1 To nS): ReDim sn(1 To nS)
        For i = 1 To nS
            q = Split(parts(i - 1), ":")
            sd(i) = ParseD(CStr(q(0))): sn(i) = (CStr(q(1)) = "N")
        Next i
    End If
    If perSpec <> "" Then
        parts = Split(perSpec, ";")
        nP = UBound(parts) + 1
        ReDim pd(1 To nP)
        For i = 1 To nP
            pd(i) = ParseD(CStr(parts(i - 1)))
        Next i
    End If
    Dim st As AttendanceStatus
    AttendanceCore crew, sd, sn, nS, pd, nP, ParseD(asOfTxt), st
    If occOnly Then                             ' v24.1: first..last:days:paid for each occurrence
        Dim oc As Long
        For oc = 1 To MAX_OCC
            If st.OccDays(oc) > 0 Then
                CaseAttend = CaseAttend & IIf(CaseAttend = "", "", ";") & FmtD(st.OccFirst(oc)) & ".." & FmtD(st.OccLast(oc)) & _
                             ":" & st.OccDays(oc) & ":" & st.OccPaid(oc)
            End If
        Next oc
        If CaseAttend = "" Then CaseAttend = "none"
        Exit Function
    End If
    CaseAttend = st.SickUsed & "/" & st.PaidUsed & "/" & st.PaidLeft & "/" & st.Occurrences & "/" & st.PersonalUsed & "/" & _
                 IIf(st.Entitled, "ENT", "WAIT " & FmtD(st.WaitUntil)) & "/" & _
                 IIf(st.LastSick = 0, "-", IIf(st.LastPaid, "PAID", "UNPAID"))
End Function

' Value specs shared with tests/test_cases.py: empty, err, d:yyyy-mm-dd, n:number, s:text
Private Function CaseValue(ByVal spec As String) As Variant
    If spec = "empty" Then
        CaseValue = Empty
    ElseIf spec = "err" Then
        CaseValue = CVErr(xlErrNA)
    ElseIf Left$(spec, 2) = "d:" Then
        CaseValue = ParseD(Mid$(spec, 3))
    ElseIf Left$(spec, 2) = "n:" Then
        CaseValue = Val(Mid$(spec, 3))
    Else
        CaseValue = Mid$(spec, 3)
    End If
End Function

' F|V (has a formula), LOG|OTHER|"" (which formula), value spec -> ok / the problem code.
' The planted typed number and the planted missing formula live here, in memory only.
Private Function CaseCell(ByVal hasF As String, ByVal which As String, ByVal valSpec As String) As String
    Dim txt As String
    If which = "LOG" Then
        txt = "=SUMIFS(" & LOG_TABLE & "[x])"
    ElseIf which = "OTHER" Then
        txt = "=SUM(A1:A2)"
    End If
    CaseCell = PeriodCellCode((hasF = "F"), txt, CaseValue(valSpec))
    If CaseCell = "" Then CaseCell = "ok"
End Function

' "a;b;a" -> "a x2" / none (the planted duplicate ID)
Private Function CaseDups(ByVal spec As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim keys() As String, n As Long, i As Long
    n = UBound(parts) + 1
    If n < 1 Then CaseDups = "none": Exit Function
    ReDim keys(1 To n)
    For i = 1 To n
        keys(i) = CStr(parts(i - 1))
    Next i
    CaseDups = DuplicatesIn(keys, n)
    If CaseDups = "" Then CaseDups = "none"
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
    LoadCases1
    LoadCases2
    LoadCases3
    LoadCases4
    LoadCases5
End Sub

Private Sub LoadCases1()
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
    AddCase "listinsert,fewest-hours-goes-first,Ann:1:50:2010-01-01;Bob:1:80:2010-01-01,New:1:10:2015-01-01,,,,1"
    AddCase "listinsert,between-two,Ann:1:50:2010-01-01;Bob:1:80:2010-01-01,New:1:60:2015-01-01,,,,2"
    AddCase "listinsert,most-hours-goes-last,Ann:1:50:2010-01-01;Bob:1:80:2010-01-01,New:1:100:2001-01-01,,,,3"
    AddCase "listinsert,grade-before-hours,Ann:1:500:2010-01-01;Bob:2:0:2010-01-01,New:1.5:999:2015-01-01,,,,2"
    AddCase "listinsert,tie-earlier-seniority-first,Ann:1:50:2010-01-01,New:1:50:2005-01-01,,,,1"
    AddCase "listinsert,tie-later-seniority-after,Ann:1:50:2010-01-01,New:1:50:2015-01-01,,,,2"
    AddCase "listinsert,tie-blank-seniority-after,Ann:1:50:2010-01-01,New:1:50:,,,,2"
    AddCase "listinsert,full-tie-name-decides-before,Bea:1:50:2010-01-01,Abe:1:50:2010-01-01,,,,1"
    AddCase "listinsert,full-tie-name-decides-after,Abe:1:50:2010-01-01,Bea:1:50:2010-01-01,,,,2"
    AddCase "listinsert,frozen-order-is-kept,Bob:1:80:2010-01-01;Ann:1:50:2010-01-01,New:1:60:2015-01-01,,,,1"
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
    AddCase "hoursasof,last-year-snapshot-ignored,2025-12-20:900,2026-01-03:Refused:Active:12:21:D,2026-01-05,,,21|none|no"
    AddCase "hoursasof,last-year-then-this-year,2025-12-20:900;2026-01-10:30,2026-01-03:Refused:Active:12:21:D,2026-01-12,,,51|2026-01-10|no"
    AddCase "leaveavg,half-rounds-up,1:N:N:Y:100.2;1:N:N:Y:100.3,,,,,100.3"
    AddCase "leaveavg,negative-half-away,1:N:N:Y:-100.2;1:N:N:Y:-100.3,,,,,-100.3"
    AddCase "attend,none,A,,,2026-03-01,,0/0/5/0/0/ENT/-"
    AddCase "attend,one-day,A,2026-01-05:N,,2026-03-01,,1/1/4/1/0/ENT/PAID"
    AddCase "attend,run-across-days-off,A,2026-01-05:N;2026-01-06:N;2026-01-09:D,,2026-03-01,,3/3/2/1/0/ENT/PAID"
    AddCase "attend,night-then-day-one-occurrence,A,2026-01-06:N;2026-01-09:D,,2026-03-01,,2/2/3/1/0/ENT/PAID"
    AddCase "attend,two-occurrences,A,2026-01-05:N;2026-01-10:D,,2026-03-01,,2/2/3/2/0/ENT/PAID"
    AddCase "attend,sixth-day-unpaid-and-wait,A,2026-01-05:N;2026-01-06:N;2026-01-09:D;2026-01-10:D;2026-01-11:D;2026-01-14:N,,2026-03-01,,6/5/0/1/0/WAIT 2027-01-05/UNPAID"
    AddCase "attend,fourth-occurrence-unpaid,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-20:D,,2026-03-01,,4/3/0/4/0/ENT/UNPAID"
    AddCase "attend,personal-pushes-over,A,2026-01-14:N;2026-01-15:N,2026-01-05;2026-01-06;2026-01-09;2026-01-10,2026-03-01,,2/1/0/1/4/WAIT 2027-01-14/UNPAID"
End Sub

Private Sub LoadCases2()
    AddCase "attend,wait-from-second-occurrence,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-19:D;2026-01-20:D;2026-01-23:N,,2026-03-01,,6/5/0/3/0/WAIT 2027-01-10/UNPAID"
    AddCase "attend,still-waiting-after-jan-1,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-19:D;2026-01-20:D;2026-01-23:N,,2027-01-05,,0/0/0/0/0/WAIT 2027-01-10/UNPAID"
    AddCase "attend,reset-at-anniversary,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-19:D;2026-01-20:D;2026-01-23:N;2027-01-12:D,,2027-01-15,,1/1/4/1/0/ENT/PAID"
    AddCase "attend,carry-over-from-last-year,A,2025-12-22:D;2025-12-23:D;2025-12-26:N;2025-12-27:N;2025-12-28:N;2025-12-31:D,,2026-06-01,,0/0/0/0/0/WAIT 2026-12-22/UNPAID"
    AddCase "attend,old-wait-already-served,A,2024-05-01:D;2024-05-02:D;2024-05-03:D;2024-05-06:D;2024-05-07:D;2024-05-08:D,,2026-03-01,,0/0/5/0/0/ENT/UNPAID"
    AddCase "attend,wait-restarts-following-year,A,2024-05-03:D;2024-05-08:N;2024-05-09:N;2024-05-13:D;2024-05-14:D;2024-05-17:N;2025-02-03:N;2025-03-03:N;2025-09-01:D;2025-09-02:D;2025-09-05:N,,2025-10-01,,3/3/2/1/0/ENT/PAID"
    AddCase "attend,waiting-into-next-year,A,2024-05-03:D;2024-05-08:N;2024-05-09:N;2024-05-13:D;2024-05-14:D;2024-05-17:N;2025-02-03:N;2025-03-03:N;2025-09-01:D;2025-09-02:D;2025-09-05:N,,2025-04-01,,2/0/0/2/0/WAIT 2025-05-08/UNPAID"
    AddCase "attend,after-as-of-ignored,A,2026-03-02:N,,2026-02-01,,0/0/5/0/0/ENT/-"
    AddCase "attend,personal-only-over,A,,2026-01-05;2026-01-06;2026-01-09;2026-01-10;2026-01-11;2026-01-14,2026-03-01,,0/0/0/0/6/WAIT 2027-01-14/-"
    AddCase "attend,unsorted-input,A,2026-01-10:D;2026-01-05:N,,2026-03-01,,2/2/3/2/0/ENT/PAID"
    AddCase "attend,third-occurrence-still-open,A,2026-01-05:N;2026-01-10:D;2026-01-15:N,,2026-01-15,,3/3/2/3/0/ENT/PAID"
    AddCase "attend,third-occurrence-ended,A,2026-01-05:N;2026-01-10:D;2026-01-15:N,,2026-01-25,,3/3/0/3/0/ENT/PAID"
    AddCase "attend,third-occurrence-run-still-open,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-16:N,,2026-01-16,,4/4/1/3/0/ENT/PAID"
    AddCase "attend,fourth-occurrence-nothing-left,A,2026-01-05:N;2026-01-10:D;2026-01-15:N;2026-01-20:D,,2026-01-21,,4/3/0/4/0/ENT/UNPAID"
    AddCase "attend,third-occurrence-night-on-asof-still-open,A,2026-01-05:N;2026-01-10:D;2026-01-14:N,,2026-01-15,,3/3/2/3/0/ENT/PAID"
    AddCase "attend,third-occurrence-ended-the-day-after,A,2026-01-05:N;2026-01-10:D;2026-01-14:N,,2026-01-16,,3/3/0/3/0/ENT/PAID"
    AddCase "year,2026-from-its-own-start,2025-12-28,2026,,,,2025-12-28|26|2026"
    AddCase "year,2027-starts-dec-27,2025-12-28,2027,,,,2026-12-27|26|2027"
    AddCase "year,2028-starts-dec-26,2025-12-28,2028,,,,2027-12-26|26|2028"
    AddCase "year,2029-after-leap-year,2025-12-28,2029,,,,2028-12-24|26|2029"
    AddCase "year,2033-has-27-periods,2025-12-28,2033,,,,2032-12-19|27|2033"
    AddCase "year,2034-starts-on-jan-1,2025-12-28,2034,,,,2034-01-01|26|2034"
    AddCase "year,2026-from-a-later-anchor,2026-12-27,2026,,,,2025-12-28|26|2026"
    AddCase "year,2033-from-2034-anchor,2034-01-01,2033,,,,2032-12-19|27|2033"
    AddCase "nextid,mark-above-log,12,480,,,,481"
    AddCase "nextid,log-above-mark,480,12,,,,481"
    AddCase "nextid,no-mark,7,0,,,,8"
    AddCase "nextid,empty-log-after-rollover,0,1312,,,,1313"
    AddCase "nextid,nothing-yet,0,0,,,,1"
    AddCase "cell,log-formula,F,LOG,n:21,,,ok"
    AddCase "cell,other-formula,F,OTHER,n:21,,,notlog"
    AddCase "cell,planted-typed-number,V,,n:21,,,typed"
    AddCase "cell,typed-zero,V,,n:0,,,typed"
    AddCase "cell,missing-formula,V,,empty,,,empty"
    AddCase "cell,formula-stored-as-text,V,,s:=SUMIFS(tblOTLog[x]),,,textformula"
    AddCase "cell,typed-text,V,,s:n/a,,,text"
    AddCase "cell,error-value,V,,err,,,error"
    AddCase "dups,planted-duplicate-entry-id,101;102;103;102;104,,,,,102 x2"
    AddCase "dups,none,S-0001;S-0002;P-0001,,,,,none"
    AddCase "dups,case-and-spaces,ss-0004; SS-0004 ;SS-0005,,,,,ss-0004 x2"
    AddCase "dups,two-keys-three-times,a;b;a;b;b;c,,,,,a x2; b x3"
    AddCase "dups,blanks-ignored,;;x;;,,,,,none"
    AddCase "archive,last-period-of-old-year,d:2026-12-13,2026-12-27,,,,archive"
    AddCase "archive,first-period-of-new-year,d:2026-12-27,2026-12-27,,,,keep"
    AddCase "archive,january-posting-logged-in-december,d:2027-01-10,2026-12-27,,,,keep"
    AddCase "archive,blank-stays,empty,2026-12-27,,,,keep"
    AddCase "archive,text-stays,s:n/a,2026-12-27,,,,keep"
    AddCase "archive,serial-number-old,n:46000,2026-12-27,,,,archive"
    AddCase "archive,serial-number-new,n:46400,2026-12-27,,,,keep"
    AddCase "archive,zero-stays,n:0,2026-12-27,,,,keep"
    AddCase "dryname,live-master,V2 - Call Out Roster.xlsm,,,,,no"
    AddCase "dryname,windows-copy,V2 - Call Out Roster - Copy.xlsm,,,,,yes"
    AddCase "dryname,copy-of,Copy of Call Out Roster.xlsm,,,,,yes"
    AddCase "dryname,dry-run,Call Out Roster DRY RUN.xlsm,,,,,yes"
    AddCase "dryname,lower-case,roster dry-run dec.xlsm,,,,,yes"
    AddCase "dryname,final,Call Out Roster 2026 FINAL.xlsm,,,,,no"
    AddCase "backupname,live-master,V2 - Call Out Roster.xlsm,,,,,no"
    AddCase "backupname,daily-backup,Call Out Roster 2026-10-12.xlsm,,,,,yes"
    AddCase "backupname,final,Call Out Roster 2026 FINAL.xlsm,,,,,yes"
    AddCase "backupname,pre-change,Backup before planned moves 2026-10-13 091500 - V2 - Call Out Roster.xlsm,,,,,yes"
    AddCase "backupname,practice-copy-is-not-one,V2 - Call Out Roster COPY 2026-10-13 0915.xlsm,,,,,no"
    AddCase "backupname,lower-case-final,roster 2026 final.xlsm,,,,,yes"
    AddCase "snapyear,no-stamp-judged-by-date,2026-12-28,0,2027,2026-12-27,,2026-12-28"
    AddCase "snapyear,last-year-stamp-inside-new-year,2026-12-28,2026,2027,2026-12-27,,2026-12-26"
    AddCase "snapyear,last-year-stamp-before-new-year,2026-12-20,2026,2027,2026-12-27,,2026-12-20"
    AddCase "snapyear,this-year-stamp,2027-01-05,2027,2027,2026-12-27,,2027-01-05"
    AddCase "snapyear,on-the-first-day,2026-12-27,2026,2027,2026-12-27,,2026-12-26"
    AddCase "hoursasof,reset-before-first-import,2025-12-20:900;2026-01-10:30,2026-01-03:Refused:Active:12:21:D,2026-01-05,,,21|none|no"
    AddCase "hoursasof,after-first-period-estimated,2025-12-20:900;2026-02-10:30,2026-01-03:Refused:Active:12:21:D,2026-01-20,,,51|2026-02-10|yes"
    AddCase "archive,date-typed-as-text-stays,s:12/13/2026,2026-12-27,,,,keep"
    AddCase "crewon,before-first-row,C:2026-12-27;A:2027-02-01,2026-10-31,,,,C"
    AddCase "crewon,old-crew-day-before,C:2026-12-27;A:2027-02-01,2027-01-31,,,,C"
    AddCase "crewon,new-crew-on-the-day,C:2026-12-27;A:2027-02-01,2027-02-01,,,,A"
    AddCase "crewon,new-crew-after,C:2026-12-27;A:2027-02-01,2027-03-01,,,,A"
    AddCase "crewon,same-date-later-row-wins,C:2026-12-27;A:2027-02-01;B:2027-02-01,2027-02-01,,,,B"
    AddCase "crewon,rows-out-of-order,A:2027-02-01;C:2026-12-27,2027-01-15,,,,C"
    AddCase "crewon,typed-lower-case,c/d:2026-12-27,2027-01-15,,,,C/D"
    AddCase "crewon,no-rows-uses-roster,,2027-01-15,DAYS,,,DAYS"
    AddCase "crewon,no-rows-no-roster,,2027-01-15,,,,none"
    AddCase "crewsched,past-absence-old-crew-days,C:2026-12-27;A:2027-02-01,2027-01-30,,,,Days"
    AddCase "crewsched,past-date-new-crew-day-is-off,C:2026-12-27;A:2027-02-01,2027-01-28,,,,off"
    AddCase "crewsched,after-change-new-crew-days,C:2026-12-27;A:2027-02-01,2027-02-05,,,,Days"
    AddCase "crewsched,after-change-old-crew-day-is-off,C:2026-12-27;A:2027-02-01,2027-02-08,,,,off"
    AddCase "eligcrew,night-ot-new-crew-days-next-morning,B:2026-12-27;A:2027-02-05,2027-02-04,Nights,,,rest"
    AddCase "eligcrew,day-ot-old-crew-off-nights,C:2026-12-27;B:2027-02-04,2027-02-04,Days,,,rest"
    AddCase "eligcrew,scheduled-on-new-crew,C:2026-12-27;B:2027-02-04,2027-02-05,Nights,,,scheduled"
    AddCase "eligcrew,free-before-change,B:2026-12-27;A:2027-02-05,2027-02-03,Days,,,free"
    AddCase "posting,exempt-ahead-not-charged,award,12,1:1:N:N:X;2:1:N:N;3:1:Y:Y,,,1=0;2=21;3=0"
    AddCase "posting,exempt-no-takers,notakers,12,1:1:N:N:X;2:1:N:N,,,1=0;2=21"
    AddCase "posting,exempt-out-of-dept,outside,12,1:1:N:N:X;2:1:N:N,,,1=0;2=21"
    AddCase "newhire,average,1:N:100;2:N:50;1.5:N:30,,,,,60"
    AddCase "newhire,grade-3-not-counted,1:N:100;3:N:0,,,,,100"
    AddCase "newhire,on-leave-not-counted,1:N:100;1:Y:10,,,,,100"
    AddCase "newhire,rounds-to-a-tenth,1:N:10;1:N:10;1:N:11,,,,,10.3"
    AddCase "newhire,no-peers,3:N:50,,,,,0"
    AddCase "unitmove,move-up,40,1:N:100;2:N:60,,,,60"
    AddCase "unitmove,own-hours-higher-no-change,120,1:N:100,,,,0"
    AddCase "unitmove,mover-not-their-own-target,40,1:Y:500;1:N:55.25,,,,15.25"
    AddCase "unitmove,grade-3-not-a-target,40,3:N:900;2:N:50,,,,10"
    AddCase "unitmove,nobody-to-match,40,3:N:900,,,,0"
    AddCase "renamelist,one-of-two,Smith Joe; Jones Al,Jones Al,Jones Alan,,,Smith Joe; Jones Alan"
    AddCase "renamelist,out-of-dept-kept,Jones Al (out of dept),Jones Al,Jones Alan,,,Jones Alan (out of dept)"
    AddCase "renamelist,longer-name-untouched,Jones Alfred; Smith Joe,Jones Al,Jones Alan,,,Jones Alfred; Smith Joe"
    AddCase "nametoken,with-hours,Fred C 8h / Sam K,Fred C,Fred D,,,Fred D 8h / Sam K"
    AddCase "nametoken,inside-a-longer-name,Alfred C / Fred Carter,Fred C,Fred D,,,Alfred C / Fred Carter"
    AddCase "nametoken,switch-line,Fred C S/S,Fred C,Freddy C,,,Freddy C S/S"
    AddCase "nametoken,twice,Fred C; Fred C 4h,Fred C,Fred D,,,Fred D; Fred D 4h"
    AddCase "awardorder,partial-only,12,1,1:N:N;2:8:Y,,,ok"
    AddCase "awardorder,full-signup-waiting,12,1,1:F:N;2:8:Y,,,full-first"
    AddCase "awardorder,full-with-more-hours-still-first,12,1,1:8:Y;2:F:N,,,full-first"
    AddCase "awardorder,full-awarded-partial-signed-only,12,1,1:F:Y;2:8:N,,,ok"
    AddCase "awardorder,position-filled-by-full,12,1,1:F:Y;2:8:Y,,,filled"
    AddCase "awardorder,second-position-to-a-partial,12,2,1:F:Y;2:8:Y,,,ok"
    AddCase "awardorder,two-partials-cover-the-shift,12,1,1:8:Y;2:4:Y,,,ok"
    AddCase "awardorder,two-partials-too-many-hours,12,1,1:8:Y;2:6:Y,,,over-hours"
    AddCase "awardorder,no-partial-awarded,12,1,1:F:Y;2:F:N,,,ok"
    AddCase "awardcover,remainder-drops-off,12,1,1:8:Y;2:N:N,,,8|4"
    AddCase "awardcover,full-plus-partial,12,2,1:F:Y;2:8:Y,,,20|4"
    AddCase "awardcover,dst-night-partial,13,1,1:6.5:Y,,,6.5|6.5"
    AddCase "awardcover,full-shift,12,1,1:F:Y,,,12|0"
    AddCase "posting,partial-award-remainder-uncharged,award,12,1:1:N:N;2:1:Y:Y;3:1:N:N,,,1=21;2=0;3=0"
    AddCase "awardorder,filled-comes-before-full-first,12,1,1:F:Y;2:F:N;3:8:Y,,,filled"
    AddCase "attendocc,none,A,,,2026-03-01,,none"
    AddCase "attendocc,two-occurrences,A,2026-01-05:N;2026-01-06:N;2026-03-02:D,,2026-04-01,,2026-01-05..2026-01-06:2:2;2026-03-02..2026-03-02:1:1"
    AddCase "attendocc,fourth-occurrence-unpaid,A,2026-01-05:N;2026-02-02:N;2026-03-02:N;2026-03-30:N,,2026-04-15,,2026-01-05..2026-01-05:1:1;2026-02-02..2026-02-02:1:1;2026-03-02..2026-03-02:1:1;2026-03-30..2026-03-30:1:0"
    AddCase "refsplit,mixed,Callout:21:A:Y;Posting:21:A:Y;Posting:22.75:A:Y;Adjustment:12.3:A:Y,,,,,21|43.75|12.3"
    AddCase "refsplit,void-not-counted,Callout:21:V:Y;Callout:21:A:Y,,,,,21|0|0"
    AddCase "refsplit,outside-the-year,Posting:21:A:N;Posting:21:A:Y,,,,,0|21|0"
    AddCase "refsplit,negative-adjustment,Callout:21:A:Y;Adjustment:-21:A:Y,,,,,21|0|-21"
    AddCase "refsplit,opening-balance-is-other,Opening Balance:40:A:Y,,,,,0|0|40"
    AddCase "calloffer,partial-then-refusal,12,Refused:0;Accepted:8;Refused:0;Left Message:0,,,,12=21;12=0;0=0;0=0"
    AddCase "calloffer,whole-shift-changes-nothing,12,Accepted:0;Refused:0,,,,12=0;12=21"
    AddCase "calloffer,two-partials,12,Accepted:8;Accepted:2;Refused:0,,,,12=0;0=0;0=0"
    AddCase "calloffer,partial-covering-the-rest,12,Accepted:8;Accepted:4;Refused:0,,,,12=0;0=0;0=0"
    AddCase "calloffer,unavailable-not-charged,13,Accepted:6.5;Unavailable:0;Refused:0,,,,13=0;0=0;0=0"
    AddCase "signups,one-whole-shift,Ann:S:0:2027-01-10-09,2027-01-12-08,,,,Ann=F"
    AddCase "signups,partial-hours,Ann:S:8:2027-01-10-09,2027-01-12-08,,,,Ann=8"
    AddCase "signups,withdrawn,Ann:S:0:2027-01-10-09;Ann:W:0:2027-01-10-11,2027-01-12-08,,,,none"
    AddCase "signups,signed-again-after-withdrawing,Ann:S:0:2027-01-10-09;Ann:W:0:2027-01-10-11;Ann:S:8:2027-01-11-07,2027-01-12-08,,,,Ann=8"
    AddCase "signups,newest-wins-whatever-the-file-order,Ann:S:8:2027-01-11-07;Ann:S:0:2027-01-10-09,2027-01-12-08,,,,Ann=8"
    AddCase "signups,after-closing-is-late,Ann:S:0:2027-01-10-09;Bob:S:0:2027-01-12-09,2027-01-12-08,,,,Ann=F;Bob=FL"
    AddCase "signups,three-people,Ann:S:0:2027-01-10-09;Bob:S:4:2027-01-10-10;Cy:W:0:2027-01-10-11,2027-01-12-08,,,,Ann=F;Bob=4"
    AddCase "privacy,plain-name,s:Smith Pat,,,,,ok"
    AddCase "privacy,planted-email,s:pat.smith@example.com,,,,,email"
    AddCase "privacy,planted-phone-dashes,s:905-555-0142,,,,,phone"
    AddCase "privacy,planted-phone-brackets,s:(905) 555 0142,,,,,phone"
    AddCase "privacy,planted-phone-seven-digits,s:call 5550142,,,,,phone"
    AddCase "privacy,payroll-number-as-text,s:10045678,,,,,phone"
    AddCase "privacy,payroll-number-as-number,n:10045678,,,,,number"
    AddCase "privacy,hours-figure,n:1398.25,,,,,ok"
    AddCase "privacy,times-are-fine,s:0700-1900,,,,,ok"
    AddCase "privacy,posting-id-is-fine,s:P-0012,,,,,ok"
    AddCase "privacy,iso-date-text-is-fine,s:2027-01-12 note,,,,,ok"
    AddCase "privacy,a-date-is-not-a-number,d:2027-01-12,,,,,ok"
    AddCase "privacy,blank,empty,,,,,ok"
    AddCase "attendpriv,count-number,count,n:3,,,,ok"
    AddCase "attendpriv,count-blank,count,empty,,,,ok"
    AddCase "attendpriv,planted-sick-day-date,count,d:2027-02-03,,,,stop"
    AddCase "attendpriv,planted-date-as-text,count,s:Feb 3,,,,stop"
    AddCase "attendpriv,status-available,status,s:Paid sick available,,,,ok"
    AddCase "attendpriv,status-until,status,s:No paid sick until Mar 4. 2028,,,,stop"
    AddCase "attendpriv,status-reason-leaks,status,s:Off sick Feb 3 with flu,,,,stop"
    AddCase "attendpriv,status-blank,status,s:,,,,stop"
    AddCase "leavewords,plain-leave,Smith P (Leave 8h),,,,,ok"
    AddCase "leavewords,planted-std,Smith P (STD),,,,,STD"
    AddCase "leavewords,planted-parental-lowercase,Smith P (par leave),,,,,PAR"
    AddCase "leavewords,planted-wsib,Jones A (WSIB) days,,,,,WSIB"
    AddCase "leavewords,vacation-is-fine,Smith P (VAC),,,,,ok"
    AddCase "leavewords,std-inside-a-word,Standard shift cover,,,,,ok"
    AddCase "leavewords,ltd-company-name,Acme Ltd contractor,,,,,LTD"
    AddCase "leavewords,parental-word,covering parental,,,,,PARENTAL"
    AddCase "leavewords,leave-of-absence,on leave of absence until June,,,,,LEAVE OF ABSENCE"
    AddCase "leavewords,sick-tag,Smith P (Sick),,,,,SICK"
    AddCase "leavewords,personal-tag,Smith P (Personal 8h),,,,,PERSONAL"
    AddCase "leavewords,off-is-fine,Smith P (Off 8h),,,,,ok"
    AddCase "leavewords,empty,,,,,,ok"
    AddCase "paytext,number,n:10045678,,,,,10045678"
    AddCase "paytext,number-with-decimals,n:10045678.0,,,,,10045678"
    AddCase "paytext,text-keeps-leading-zero,s:00456789,,,,,00456789"
    AddCase "paytext,text-trimmed,s: 10045678 ,,,,,10045678"
    AddCase "paytext,blank,empty,,,,,blank"
    AddCase "paytext,zero,n:0,,,,,blank"
    AddCase "paytext,error-cell,err,,,,,blank"
    AddCase "payhdr,as-reported,Payroll #,,,,,yes"
    AddCase "payhdr,upper-no-space,PAYROLL#,,,,,yes"
    AddCase "payhdr,number-spelled,Payroll Number,,,,,yes"
    AddCase "payhdr,no-dot,Payroll No.,,,,,yes"
    AddCase "payhdr,employee-name,Employee Name,,,,,no"
    AddCase "payhdr,grand-total,Grand Total,,,,,no"
    AddCase "payhdr,payroll-hours-is-not-it,Payroll Hours,,,,,no"
    AddCase "krhdr,name-as-reported,Employee Name,Employee Name,,,,yes"
    AddCase "krhdr,name-upper,EMPLOYEE NAME,Employee Name,,,,yes"
    AddCase "krhdr,name-double-space,Employee  Name,Employee Name,,,,yes"
    AddCase "krhdr,name-alone-is-not-it,Name,Employee Name,,,,no"
    AddCase "krhdr,total-as-reported,Grand Total,Grand Total,,,,yes"
    AddCase "krhdr,total-lower,Grand total,Grand Total,,,,yes"
    AddCase "krhdr,total-joined,GrandTotal,Grand Total,,,,yes"
    AddCase "krhdr,total-colon,Grand Total:,Grand Total,,,,yes"
    AddCase "krhdr,total-hours-is-not-it,Total Hours,Grand Total,,,,no"
    AddCase "krhdr,empty-cell,,Grand Total,,,,no"
End Sub

Private Sub LoadCases3()
    AddCase "payroll,no-number-yet,,N,N,,,name"
    AddCase "payroll,number-not-in-kronos,10045678,N,N,,,name"
    AddCase "payroll,number-and-name-agree,10045678,Y,Y,,,number"
    AddCase "payroll,number-under-another-name,10045678,Y,N,,,mismatch"
    AddCase "plancode,before-date-old-unit-stays,Derivatives:D2:Phosphine:2027-02-01:Planned,Phosphine,2027-01-31,P2,,P2"
    AddCase "plancode,before-date-new-unit-blank,Derivatives:D2:Phosphine:2027-02-01:Planned,Derivatives,2027-01-31,,,none"
    AddCase "plancode,on-the-date-old-unit-cleared,Derivatives:D2:Phosphine:2027-02-01:Planned,Phosphine,2027-02-01,P2,,none"
    AddCase "plancode,on-the-date-new-unit-code,Derivatives:D2:Phosphine:2027-02-01:Planned,Derivatives,2027-02-01,,,D2"
    AddCase "plancode,after-the-date-new-unit-code,Derivatives:D2:Phosphine:2027-02-01:Planned,Derivatives,2027-03-15,,,D2"
    AddCase "plancode,unit-not-in-the-move,Derivatives:D2:Phosphine:2027-02-01:Planned,Cypure,2027-03-15,C3,,C3"
    AddCase "plancode,extra-code-kept-nothing-cleared,Derivatives:D2::2027-02-01:Planned,Phosphine,2027-03-15,P2,,P2"
    AddCase "plancode,two-units-cleared,Derivatives:D2:Phosphine|Cypure:2027-02-01:Planned,Cypure,2027-02-01,C3,,none"
    AddCase "plancode,new-code-replaces-old-code,Derivatives:D2:Phosphine:2027-02-01:Planned,Derivatives,2027-02-01,D3,,D2"
    AddCase "plancode,old-code-until-the-date,Derivatives:D2:Phosphine:2027-02-01:Planned,Derivatives,2027-01-31,D3,,D3"
    AddCase "plancode,cancelled-is-ignored,Derivatives:D2:Phosphine:2027-02-01:Cancelled,Phosphine,2027-03-15,P2,,P2"
    AddCase "plancode,applied-is-ignored,Derivatives:D2:Phosphine:2027-02-01:Applied,Derivatives,2027-03-15,,,none"
    AddCase "plancode,someone-elses-move,Derivatives:D2:Phosphine:2027-02-01:Planned:other,Phosphine,2027-03-15,P2,,P2"
    AddCase "plancode,unit-name-any-case,Derivatives:D2:Phosphine:2027-02-01:Planned,phosphine,2027-02-01,P2,,none"
    AddCase "plancode,no-plans,,Phosphine,2027-02-01,P2,,P2"
    AddCase "plancode,employee-name-any-case,Derivatives:D2:Phosphine:2027-02-01:Planned:lower,Phosphine,2027-02-01,P2,,none"
    AddCase "plancode,status-any-case,Derivatives:D2:Phosphine:2027-02-01:planned,Phosphine,2027-02-01,P2,,none"
    AddCase "plancode,clear-list-typed-with-spaces,Derivatives:D2:Phosphine | Cypure:2027-02-01:Planned,Cypure,2027-02-01,C3,,none"
    AddCase "plancode,two-moves-only-the-first-is-due,Derivatives:D2:Phosphine:2027-02-01:Planned;Cypure:C2:Derivatives:2027-04-01:Planned,Derivatives,2027-03-01,,,D2"
    AddCase "plancode,two-moves-the-later-one-counts,Derivatives:D2:Phosphine:2027-02-01:Planned;Cypure:C2:Derivatives:2027-04-01:Planned,Derivatives,2027-04-01,D2,,none"
    AddCase "plancode,same-date-later-row-wins,Derivatives:D2:Phosphine:2027-02-01:Planned;Cypure:C2:Phosphine:2027-02-01:Planned,Cypure,2027-02-01,,,C2"
    AddCase "plancode,new-unit-also-ticked-to-clear,Derivatives:D2:Derivatives|Phosphine:2027-02-01:Planned,Derivatives,2027-02-01,D3,,D2"
    AddCase "planhours,below-the-highest-peer,40,1:N:100;2:N:60,,,,100"
    AddCase "planhours,already-above-keeps-own,120,1:N:100,,,,120"
    AddCase "planhours,own-entry-is-not-a-peer,10,1:Y:10;1:N:55.25,,,,55.25"
    AddCase "planhours,another-mover-is-not-a-peer,10,1:Y:10;1:Y:900;2:N:60,,,,60"
    AddCase "planhours,grade-3-is-not-a-peer,40,3:N:900;2:N:50,,,,50"
    AddCase "planhours,nobody-to-match-keeps-own,40,3:N:900,,,,40"
    AddCase "crewsoff,anchor-day1,2025-12-28,,,,,C/D"
    AddCase "crewsoff,day2,2025-12-29,,,,,A/B"
    AddCase "crewsoff,oct14-2026,2026-10-14,,,,,A/B"
    AddCase "shifthours,fall-night,2026-10-31,night,,,,13"
    AddCase "shifthours,spring-night,2026-03-07,night,,,,11"
    AddCase "shifthours,fall-day,2026-10-31,day,,,,12"
    AddCase "shifthours,plain-night,2026-06-01,night,,,,12"
    AddCase "accepted,part,12,8,,,,8"
    AddCase "accepted,whole,12,12,,,,12"
    AddCase "accepted,zero,12,0,,,,12"
    AddCase "accepted,over,12,15,,,,12"
    AddCase "typeddate,jan-in-december,Jan_5,2026-12-15,,,,2027-01-05"
    AddCase "typeddate,recent-past-stays,Oct_12,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,old-rolls-over,Sep_1,2026-12-15,,,,2027-09-01"
    AddCase "typeddate,year-typed,Oct_12_2026,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,year-with-comma,Oct_12_2026,2027-03-01,,,,2026-10-12"
    AddCase "typeddate,iso,2026-10-12,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,time-after,Oct_12_7pm,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,clock-after,Oct_12_7:00,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,ahead-in-march,Dec_1,2026-03-01,,,,2026-12-01"
    AddCase "typeddate,junk,junk,2026-12-15,,,,not a date"
    AddCase "typeddate,blank,,2026-12-15,,,,blank"
    AddCase "typeddate,decades-back,Mar_3_2011,2026-12-15,,,,check the year"
    AddCase "typeddate,too-far-ahead,2028-06-01,2026-12-15,,,,check the year"
    AddCase "typeddate,december-in-early-january-is-last-december,Dec_20,2027-01-03,,,,2026-12-20"
    AddCase "typeddate,july-in-october-is-next-july,Jul_1,2026-10-14,,,,2027-07-01"
    AddCase "typeddate,feb-29-in-a-non-leap-year,Feb_29,2026-10-14,,,,not a date"
    AddCase "typeddate,feb-29-leap-year,Feb_29_2028,2027-12-01,,,,2028-02-29"
    AddCase "typeddate,sep-31,Sep_31,2026-10-14,,,,not a date"
    AddCase "typeddate,am-pm-with-a-space,Jan_5_7_pm,2026-12-15,,,,2027-01-05"
    AddCase "typeddate,slash-month-day,10/12,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,slash-with-year,12/10/2026,2026-12-15,,,,2026-12-10"
    AddCase "typeddate,full-month-name,October_12,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,day-then-month,12_Oct,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,two-digit-year,Oct_12_26,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,month-point,Oct._12,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,number-only,12,2026-12-15,,,,not a date"
    AddCase "typeddate,feb-29-ahead-of-a-leap-february,Feb_29,2027-10-14,,,,2028-02-29"
    AddCase "typeddate,clock-then-pm,Oct_12_7:00_pm,2026-12-15,,,,2026-10-12"
    AddCase "typeddate,year-clock-pm,Oct_12_2026_7:00_pm,2026-12-15,,,,2026-10-12"
    AddCase "eligibility,rotation-rest-day-ot,C,2026-10-07,Days,,,rest"
    AddCase "eligibility,gave-up-night-relieves-rest,C,2026-10-07,Days,ssgaveup,,free"
    AddCase "eligibility,rotation-rest-night-ot,A,2026-10-06,Nights,,,rest"
    AddCase "eligibility,gave-up-day-relieves-rest,A,2026-10-06,Nights,ssgaveup,,free"
    AddCase "eligibility,gave-up-elsewhere-no-relief,C,2026-10-31,Days,ssgaveup,,free"
    AddCase "trimkind,exact,2026-10-05,2026-10-09,2026-10-05,2026-10-09,,none"
    AddCase "trimkind,days-off-at-start,2026-10-03,2026-10-09,2026-10-05,2026-10-09,,start"
    AddCase "trimkind,days-off-at-end,2026-10-05,2026-10-11,2026-10-05,2026-10-09,,end"
    AddCase "trimkind,both-ends,2026-10-03,2026-10-11,2026-10-05,2026-10-09,,both"
    AddCase "trimkind,one-shift,2026-10-03,2026-10-11,2026-10-07,2026-10-07,,both"
    AddCase "partialhrs,blank-is-whole,,12,,,,whole"
    AddCase "partialhrs,eight,8,12,,,,8"
    AddCase "partialhrs,half-hours,7.5,12,,,,7.5"
    AddCase "partialhrs,the-whole-shift,12,12,,,,whole"
    AddCase "partialhrs,more-than-the-shift,15,12,,,,whole"
    AddCase "partialhrs,words,abc,12,,,,bad"
    AddCase "partialhrs,zero,0,12,,,,bad"
    AddCase "partialhrs,negative,-3,12,,,,bad"
    AddCase "calloffer,refusals-before-the-partial-keep-their-charge,12,Refused:0;Left Message:0;Accepted:4;Refused:0,,,,12=21;12=21;12=0;0=0"
    AddCase "postclosed,day-before,2026-10-20,2026-10-19,,,,open"
    AddCase "postclosed,closing-day-still-open,2026-10-20,2026-10-20,,,,open"
    AddCase "postclosed,day-after,2026-10-20,2026-10-21,,,,closed"
    AddCase "backupdate,daily-backup,Call Out Roster 2026-10-12.xlsm,,,,,2026-10-12"
    AddCase "backupdate,live-master,V2 - Call Out Roster.xlsm,,,,,none"
    AddCase "backupdate,final-is-not-dated,Call Out Roster 2026 FINAL.xlsm,,,,,none"
    AddCase "backupdate,wrong-extension,Call Out Roster 2026-10-12.xlsx,,,,,none"
    AddCase "backupdate,letters-in-the-date,Call Out Roster 2026-1O-12.xlsm,,,,,none"
    AddCase "backupdate,month-13-rolls-over-as-dateserial-does,Call Out Roster 2026-13-12.xlsm,,,,,2027-01-12"
    AddCase "deptcode,d1,D1,,,,,1"
    AddCase "deptcode,p1-5,P1.5,,,,,1.5"
    AddCase "deptcode,g3,G3,,,,,3"
    AddCase "deptcode,blank,,,,,,999"
    AddCase "deptcode,letter-only,D,,,,,999"
    AddCase "deptcode,bare-number,2,,,,,2"
    AddCase "deptcode,trailing-space,P2_,,,,,2"
    AddCase "deptcode,two-points-not-numeric,D1.5.2,,,,,999"
    AddCase "deptcode,stops-at-the-first-gap,D1-2,,,,,1"
    AddCase "hourstext,whole,8,,,,,8"
    AddCase "hourstext,quarter,22.75,,,,,22.75"
    AddCase "hourstext,zero,0,,,,,0"
    AddCase "hourstext,third-rounds-to-hundredths,21.333,,,,,21.33"
    AddCase "hourstext,negative-quarter,-0.25,,,,,-0.25"
    AddCase "hourstext,twelve-and-a-half,12.5,,,,,12.5"
    AddCase "typetag,vacation,VAC,,,,,Vac/VAC"
    AddCase "typetag,std-is-leave-and-lv,STD,,,,,Leave/LV"
    AddCase "typetag,wsib-is-leave-and-lv,WSIB,,,,,Leave/LV"
    AddCase "typetag,ltd,LTD,,,,,Leave/LV"
    AddCase "typetag,olv,OLV,,,,,Leave/LV"
    AddCase "typetag,par,PAR,,,,,Leave/LV"
    AddCase "typetag,loa,LOA,,,,,Leave/LV"
    AddCase "typetag,sick-call-members-see-off,SIC,,,,,Sick/OFF"
    AddCase "typetag,personal-day-members-see-off,PER,,,,,Personal/OFF"
    AddCase "typetag,shift-switch,SS,,,,,S/S/SS"
    AddCase "typetag,training,TRN,,,,,Trn/TRN"
    AddCase "typetag,support,SUP,,,,,Support/SUP"
    AddCase "typetag,lower-case-code,std,,,,,Leave/LV"
    AddCase "typetag,unknown-code-is-off,XYZ,,,,,Off/XYZ"
    AddCase "planid,moves-mark,Exempt - moves Feb 1 (PM-0007),,,,,PM-0007"
    AddCase "planid,cancelled-mark,Exempt - move cancelled (PM-0012),,,,,PM-0012"
    AddCase "planid,no-id,Exempt - left Feb 1,,,,,none"
    AddCase "planid,lower-case-tag,exempt - moves Feb 1 (pm-0003),,,,,pm-0003"
    AddCase "planid,unclosed-bracket,Exempt - moves Feb 1 (PM-0007,,,,,none"
    AddCase "namekey,kronos-form,Dobbin_Stephen,,,,,dobbinstephen"
    AddCase "namekey,roster-form,Dobbin_Stephen_A.,,,,,dobbinstephena"
    AddCase "namekey,trailing-space-and-case,_SMITH_Paul_,,,,,smithpaul"
    AddCase "barlist,in-plain,in,Phosphine|Cypure,Cypure,,,yes"
    AddCase "barlist,in-trimmed-any-case,in,Phosphine_|_Cypure_|,cypure,,,yes"
    AddCase "barlist,in-not-there,in,Phosphine|Cypure,Gases,,,no"
    AddCase "barlist,in-blank-item-never,in,Phosphine||Cypure,,,,no"
    AddCase "barlist,in-partial-is-not-a-match,in,Phosphine|Cypure,Cyp,,,no"
    AddCase "barlist,trim-spaces-and-blanks,trim,Phosphine_|_Cypure_|,,,,Phosphine|Cypure"
    AddCase "barlist,trim-already-clean,trim,Phosphine|Cypure,,,,Phosphine|Cypure"
    AddCase "barlist,trim-only-bars,trim,_|_|,,,,empty"
    AddCase "matchname,one-match,Paul_M,Mason/Paul|Jones/Peter,,,,Mason/Paul"
    AddCase "matchname,first-name-only-one-match,Paul,Mason/Paul|Jones/Peter,,,,Mason/Paul"
    AddCase "matchname,two-pauls-ambiguous,Paul,Mason/Paul|Price/Paul,,,,?"
    AddCase "matchname,initial-picks-one,Paul_P,Mason/Paul|Price/Paul,,,,Price/Paul"
    AddCase "matchname,no-match,Fred,Mason/Paul|Jones/Peter,,,,none"
    AddCase "matchname,extra-words-ignored,Paul_M_8h_vac,Mason/Paul|Jones/Peter,,,,Mason/Paul"
    AddCase "matchname,first-last-roster-form,Paul_M,Paul_Mason|Peter_Jones,,,,Paul Mason"
    AddCase "matchname,first-is-a-prefix,Pa,Mason/Paul|Jones/Peter,,,,Mason/Paul"
    AddCase "attendstatus,available,Y,5,0,,,Paid sick available"
    AddCase "attendstatus,used-up,Y,0,2,,,Paid sick days used up"
    AddCase "attendstatus,third-occurrence-running,Y,2,3,,,Paid sick only while the current occurrence continues"
    AddCase "attendstatus,waiting,N,0,3,2027-03-04,,No paid sick until Mar 4. 2027"
    AddCase "attendstatus,waiting-beats-used-up,N,3,1,2027-01-15,,No paid sick until Jan 15. 2027"
    AddCase "attendstatus,used-up-beats-occurrences,Y,0,3,,,Paid sick days used up"
    AddCase "totalsdiff,same,Smith=12|21|21;Jones=8|0|0,Smith=12|21|21;Jones=8|0|0,,,,same"
    AddCase "totalsdiff,one-changed,Smith=12|21|21;Jones=8|0|0,Smith=13|21|21;Jones=8|0|0,,,,Smith: was Total 12/Refused 21/periods 21/now Total 13/Refused 21/periods 21"
    AddCase "totalsdiff,the-person-the-change-is-about-is-skipped,Smith=12|21|21;Jones=8|0|0,Smith=13|21|21;Jones=8|0|0,smith,,,same"
    AddCase "totalsdiff,removed,Smith=12|21|21;Jones=8|0|0,Jones=8|0|0,,,,Smith: no longer on the roster"
    AddCase "totalsdiff,removed-but-skipped,Smith=12|21|21;Jones=8|0|0,Jones=8|0|0,Smith,,,same"
    AddCase "totalsdiff,appeared,Jones=8|0|0,Jones=8|0|0;New=0|0|0,,,,New: appeared on the roster"
    AddCase "totalsdiff,quarter-hours-shown-plain,Smith=12.75|21|21,Smith=12.5|21|21,,,,Smith: was Total 12.75/Refused 21/periods 21/now Total 12.5/Refused 21/periods 21"
    AddCase "totalsdiff,two-lines,Smith=12|21|21;Jones=8|0|0,Smith=13|21|21,,,,Smith: was Total 12/Refused 21/periods 21/now Total 13/Refused 21/periods 21;Jones: no longer on the roster"
    AddCase "totalsdiff,capped-at-ten-lines,A=1|0|0;B=1|0|0;C=1|0|0;D=1|0|0;E=1|0|0;F=1|0|0;G=1|0|0;H=1|0|0;I=1|0|0;J=1|0|0;K=1|0|0;L=1|0|0,,,,,A: no longer on the roster;B: no longer on the roster;C: no longer on the roster;D: no longer on the roster;E: no longer on the roster;F: no longer on the roster;G: no longer on the roster;H: no longer on the roster;I: no longer on the roster;J: no longer on the roster;...and 2 more"
    AddCase "viewmode,both-lists,boss,ann;boss,boss,,,desk"
    AddCase "viewmode,admin-only,ann,ann;boss,boss,,,admin"
    AddCase "viewmode,desk-only-is-a-user,boss,ann,boss,,,user"
    AddCase "viewmode,neither,sam,ann;boss,boss,,,user"
    AddCase "viewmode,case-and-spaces,BOSS,ann ;Boss,boss ,,,desk"
    AddCase "viewmode,empty-login,,ann;boss,boss,,,user"
    AddCase "viewmode,empty-lists,boss,,,,,user"
    AddCase "viewmode,supervisor-on-list,sue,ann;boss,boss,sue;tom,,supervisor"
    AddCase "viewmode,supervisor-case-and-spaces,SUE,ann;boss,boss, Sue ;tom,,supervisor"
    AddCase "viewmode,supervisor-who-is-admin,ann,ann;boss,boss,ann;sue,,admin"
    AddCase "viewmode,supervisor-who-is-desk-admin,boss,ann;boss,boss,boss;sue,,desk"
    AddCase "viewmode,not-a-supervisor,sam,ann;boss,boss,sue;tom,,user"
    AddCase "viewmode,empty-supervisor-list,sue,ann;boss,boss,,,user"
    AddCase "payback,within-limit,10,42,,,,Payback owed"
    AddCase "payback,at-the-limit,42,42,,,,Payback owed"
    AddCase "payback,one-day-over,43,42,,,,Payback owed 43 days"
    AddCase "payback,long-overdue,120,42,,,,Payback owed 120 days"
    AddCase "payback,today,0,42,,,,Payback owed"
    AddCase "payback,limit-zero-flags-one-day,1,0,,,,Payback owed 1 day"
    AddCase "idmark,three-prefixes,S-0042|S-0431|P-0017|LV-0003,,,,,LV=3|P=17|S=431"
    AddCase "idmark,junk-skipped,S-0042|junk|S-x|-0012|S-,,,,,S=42"
    AddCase "idmark,lower-case-and-spaces,s-0007| S-0009 ,,,,,S=9"
    AddCase "idmark,entry-style-number-not-a-ref,#12|12,,,,,none"
    AddCase "idmark,long-prefix-ignored,TOOLONGPRE-0001|ADJ-0100,,,,,ADJ=100"
    AddCase "idmark,empty,,,,,,none"
    AddCase "idmark,digit-in-prefix-ignored,P2-0005|P-0004,,,,,P=4"
    AddCase "prechange,six-digit-time,Backup before planned moves 2026-10-13 091500 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,2026-10-13"
    AddCase "prechange,four-digit-time,Backup before Home Unit 2026-10-12 0915 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,2026-10-12"
    AddCase "prechange,another-rosters-copy,Backup before planned moves 2026-10-13 091500 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster COPY 2026-10-13 0915.xlsm,,,,none"
    AddCase "prechange,daily-backup-is-not-one,Call Out Roster 2026-10-12.xlsm,V2 - Call Out Roster.xlsm,,,,none"
    AddCase "prechange,no-time,Backup before migration 2026-10-13 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,none"
    AddCase "prechange,no-what,Backup before 2026-10-13 0915 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,none"
End Sub

Private Sub LoadCases4()
    AddCase "prechange,lower-case-head,backup before rename 2026-09-30 120000 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,2026-09-30"
    AddCase "prechange,letters-in-the-date,Backup before rename 2026-O9-30 120000 - V2 - Call Out Roster.xlsm,V2 - Call Out Roster.xlsm,,,,none"
    AddCase "prechange,no-roster-name,Backup before rename 2026-09-30 120000 - V2 - Call Out Roster.xlsm,,,,,none"
    AddCase "chargewhy,callout-refused-charged,Callout,Refused,3,12|21,P2,Charged 21 h: 3rd on the call list; Refused; 12 h offered x 1.75 (grade P2)"
    AddCase "chargewhy,callout-left-message-first,Callout,Left Message,1,12|21,D1.5,Charged 21 h: 1st on the call list; Left Message; 12 h offered x 1.75 (grade D1.5)"
    AddCase "chargewhy,callout-accepted,Callout,Accepted,2,12|0,P2,Not charged: 2nd on the call list; accepted 12 h"
    AddCase "chargewhy,callout-grade-3-not-charged,Callout,Refused,11,12|0,G3,Not charged: 11th on the call list; Refused; grade G3 is above 2"
    AddCase "chargewhy,callout-unavailable,Callout,Unavailable,22,12|0,P1,Not charged: 22nd on the call list; Unavailable"
    AddCase "chargewhy,callout-no-order-no-code,Callout,Refused,0,12|21,,Charged 21 h: Refused; 12 h offered x 1.75 (no code)"
    AddCase "chargewhy,callout-partial-offer,Callout,Refused,4,8|14,P2,Charged 14 h: 4th on the call list; Refused; 8 h offered x 1.75 (grade P2)"
    AddCase "chargewhy,callout-no-outcome,Callout,,5,12|0,P2,Not charged: 5th on the call list; no outcome"
    AddCase "chargewhy,posting-charged,Posting,Refused,3,12|21,P2,Charged 21 h: #3 in call order; didn't sign up; posting hours 12 x 1.75 (grade P2)"
    AddCase "chargewhy,posting-awarded,Posting,Awarded,5,12|0,P2,Not charged: #5 in call order; awarded 12 h"
    AddCase "chargewhy,posting-signed-up,Posting,Signed Up,4,12|0,P2,Not charged: #4 in call order; signed up; not awarded"
    AddCase "chargewhy,posting-grade-3,Posting,Refused,2,12|0,P3,Not charged: #2 in call order; didn't sign up; grade P3 is above 2"
    AddCase "chargewhy,posting-out-of-dept-award-no-rank,Posting,Awarded,0,12|0,OD,Not charged: awarded 12 h"
    AddCase "chargewhy,adjustment-leave-return,Adjustment,Leave return,0,0|40,,Adjustment: +40 h (Leave return)"
    AddCase "chargewhy,adjustment-negative,Adjustment,Correction,0,0|-8,,Adjustment: -8 h (Correction)"
    AddCase "chargewhy,blank-type,,,0,0|0,,Entry: +0 h"
    AddCase "chargewhy,ordinal-13th,Callout,Refused,13,8|14,P2,Charged 14 h: 13th on the call list; Refused; 8 h offered x 1.75 (grade P2)"
    AddCase "chargewhy,ordinal-101st,Callout,Unavailable,101,8|0,P2,Not charged: 101st on the call list; Unavailable"
    AddCase "chargewhy,ordinal-112th,Callout,Unavailable,112,8|0,P2,Not charged: 112th on the call list; Unavailable"
    AddCase "chargewhy,ordinal-23rd,Callout,Unavailable,23,8|0,P2,Not charged: 23rd on the call list; Unavailable"
    AddCase "minstaff,at-minimum,4,0,0,4,,ok"
    AddCase "minstaff,one-short,4,1,0,4,,SHORT 3/4"
    AddCase "minstaff,covered-back-up,4,1,1,4,,ok"
    AddCase "minstaff,extra-cover-over,3,0,2,4,,ok"
    AddCase "minstaff,no-minimum-never-short,0,0,0,0,,ok"
    AddCase "minstaff,nobody-scheduled,0,0,0,1,,SHORT 0/1"
    AddCase "minstaff,all-off,2,2,0,2,,SHORT 0/2"
    AddCase "minstaff,negative-minimum-is-none,1,1,0,-1,,ok"
    AddCase "withdrawn,withdrawn-same-login,Ann:S:0:2027-01-10-09:jdoe;Ann:W:0:2027-01-10-11:jdoe,,,,,Ann=jdoe"
    AddCase "withdrawn,withdrawn-other-login,Ann:S:0:2027-01-10-09:jdoe;Ann:W:0:2027-01-10-11:psmith,,,,,Ann=psmith!"
    AddCase "withdrawn,signed-again-after,Ann:S:0:2027-01-10-09:jdoe;Ann:W:0:2027-01-10-11:jdoe;Ann:S:8:2027-01-11-07:jdoe,,,,,none"
    AddCase "withdrawn,withdraw-with-no-sign-up,Ann:W:0:2027-01-10-11:jdoe,,,,,Ann=jdoe"
    AddCase "withdrawn,two-people-one-withdrawn,Ann:S:0:2027-01-10-09:jdoe;Bob:S:4:2027-01-10-10:bobk;Bob:W:0:2027-01-10-12:bobk,,,,,Bob=bobk"
    AddCase "withdrawn,earlier-sign-up-other-login-whatever-the-file-order,Ann:W:0:2027-01-10-11:jdoe;Ann:S:0:2027-01-10-09:psmith,,,,,Ann=jdoe!"
    AddCase "withdrawn,nobody,Ann:S:0:2027-01-10-09:jdoe,,,,,none"
    AddCase "privacy,dashed-local-number,s:555-0142,,,,,phone"
    AddCase "privacy,dashed-local-number-in-text,s:call ext 555-0142 after 4,,,,,phone"
    AddCase "privacy,times-without-the-leading-zero-are-refused,s:700-1900,,,,,phone"
    AddCase "privacy,digits-before-the-dash-group-are-fine,s:12345-6789,,,,,ok"
    AddCase "privacy,date-then-local-number,s:2027-01-12 555-0142,,,,,phone"
    AddCase "privacy,posting-hours-range,s:3-4 h,,,,,ok"
    AddCase "refthrough,current-period,2026-10-11,2026-10-11,,,,counts"
    AddCase "refthrough,last-period,2026-09-27,2026-10-11,,,,counts"
    AddCase "refthrough,next-period-waits,2026-10-25,2026-10-11,,,,waits"
    AddCase "refthrough,two-periods-out-waits,2026-11-08,2026-10-11,,,,waits"
    AddCase "refthrough,year-start,2025-12-28,2026-10-11,,,,counts"
    AddCase "fatigue,free,0,12,99,,,ok|16"
    AddCase "fatigue,night-then-day,12,12,99,,,16 h rule|4"
    AddCase "fatigue,night-then-4h,12,4,99,,,ok|4"
    AddCase "fatigue,12-plus-4-partial-accepted,16,4,99,,,16 h rule|0"
    AddCase "fatigue,12-plus-8-refused-counts-nothing,12,12,99,,,16 h rule|4"
    AddCase "fatigue,day-then-night,12,12,99,,,16 h rule|4"
    AddCase "fatigue,switch-given-up-is-rest,0,12,99,,,ok|16"
    AddCase "fatigue,next-shift-in-12,0,12,12,,,16 h rule|1"
    AddCase "fatigue,next-shift-in-12-one-hour,0,1,12,,,ok|1"
    AddCase "fatigue,next-shift-in-24,0,12,24,,,ok|13"
    AddCase "fatigue,gap-exactly-11-after-12,0,12,23,,,ok|12"
    AddCase "fatigue,gap-10-after-12,0,12,22,,,16 h rule|11"
    AddCase "fatigue,dst-13h-night-then-day,13,12,99,,,16 h rule|3"
    AddCase "fatigue,24-h-block,24,1,99,,,16 h rule|0"
    AddCase "fatigue,half-hour-fits,15.5,0.5,99,,,ok|0.5"
    AddCase "shiftdate,midnight-is-last-night,2026-10-20,00:00,,,,2026-10-19"
    AddCase "shiftdate,3am-is-last-night,2026-10-20,03:10,,,,2026-10-19"
    AddCase "shiftdate,659-is-last-night,2026-10-20,06:59,,,,2026-10-19"
    AddCase "shiftdate,7am-is-today,2026-10-20,07:00,,,,2026-10-20"
    AddCase "shiftdate,7pm-is-today,2026-10-20,19:00,,,,2026-10-20"
    AddCase "shiftdate,2359-is-today,2026-10-20,23:59,,,,2026-10-20"
    AddCase "shiftdate,jan-1-before-7,2027-01-01,02:00,,,,2026-12-31"
    AddCase "nextstep,nobody-with-messages-supervisor,0,0,2,N,Graeme,The line stays OPEN on the calendar. Next: call back later (Fill a shift now shows ? beside them); post the shift; or ask Graeme to award out of department."
    AddCase "nextstep,nobody-no-messages-admin,0,0,0,Y,Graeme,The line stays OPEN on the calendar. Next: post the shift; or award out of department."
    AddCase "nextstep,nobody-no-messages-supervisor,0,0,0,N,Graeme,The line stays OPEN on the calendar. Next: post the shift; or ask Graeme to award out of department."
    AddCase "nextstep,partial-8,1,8,0,N,Graeme,The shift is closed at 8 h; the rest is dropped; nobody below is charged."
    AddCase "nextstep,partial-4.5,1,4.5,1,Y,Graeme,The shift is closed at 4.5 h; the rest is dropped; nobody below is charged."
    AddCase "nextstep,filled,1,0,1,N,Graeme,The line is covered on the calendar."
    AddCase "nextstep,two-accepted,2,0,0,Y,Graeme,The line is covered on the calendar."
    AddCase "practice,plain,V2 - Call Out Roster COPY 2026-10-21 0314.xlsm,,,,,2026-10-21 03:14"
    AddCase "practice,lower-case-copy,V2 - Call Out Roster copy 2026-10-21 2359.xlsm,,,,,2026-10-21 23:59"
    AddCase "practice,no-copy-word,V2 - Call Out Roster 2026-10-21 0314.xlsm,,,,,none"
    AddCase "practice,daily-backup-name,Call Out Roster backup 2026-10-21.xlsm,,,,,none"
    AddCase "practice,bad-month,V2 - Call Out Roster COPY 2026-13-21 0314.xlsm,,,,,none"
    AddCase "practice,no-time,V2 - Call Out Roster COPY 2026-10-21.xlsm,,,,,none"
    AddCase "fatblock,nothing-around,0,0,0,0,,0|99"
    AddCase "fatblock,touching-full-night,0,12,0,0,,12|99"
    AddCase "fatblock,touching-full-joins-the-one-before,12,12,0,0,,24|99"
    AddCase "fatblock,touching-partial-joins-too,12,4,0,0,,16|99"
    AddCase "fatblock,before-but-not-touching-counts-nothing,12,0,0,0,,0|99"
    AddCase "fatblock,next-shift-worked,0,0,12,0,,0|12"
    AddCase "fatblock,next-but-one-worked,0,0,0,12,,0|24"
    AddCase "fatblock,next-partial-still-12,0,0,4,12,,0|12"
    AddCase "fatblock,dst-13-night-touching,0,13,0,0,,13|99"
    AddCase "postcap,no-max-known,12,0,N,,,12"
    AddCase "postcap,max-above-posting,12,16,Y,,,12"
    AddCase "postcap,max-4,12,4,Y,,,4"
    AddCase "postcap,max-0-not-charged,12,0,Y,,,0"
    AddCase "postcap,max-equal,12,12,Y,,,12"
    AddCase "kmult,overtime,OVERTIME,,,,,1.75"
    AddCase "kmult,ca-prefix,CA OVERTIME,,,,,1.75"
    AddCase "kmult,any-case,ca overtime,,,,,1.75"
    AddCase "kmult,ot-dt-split,CA OT - DT Split,,,,,2"
    AddCase "kmult,lieu-ot,CA LIEU OT,,,,,1.5"
    AddCase "kmult,lieu-dt,CA LIEU DT,,,,,2"
    AddCase "kmult,lieu-ot-dt-split,CA LIEU OT - DT Split,,,,,1.5"
    AddCase "kmult,lieu-st,CA LIEU ST,,,,,1"
    AddCase "kmult,callin-ot,CA CALLIN OT,,,,,1.75"
    AddCase "kmult,min-4,CA MIN 4,,,,,1.5"
    AddCase "kmult,statwk,CA STATWK 2.0,,,,,2"
    AddCase "kmult,dbltime-not-counted,CA DBLTIME,,,,,none"
    AddCase "kmult,unapproved-not-counted,CA Uanpproved OT,,,,,none"
    AddCase "kmult,lieu-taken-not-counted,CA LIEU TKN,,,,,none"
    AddCase "kmult,reg-not-counted,CA REG,,,,,none"
    AddCase "kmult,blank,,,,,,none"
    AddCase "kweight,one-code,CA OVERTIME=4,,,,,7"
    AddCase "kweight,two-codes,CA OVERTIME=4;CA OT - DT Split=2,,,,,11"
    AddCase "kweight,rounds-half-up,CA OVERTIME=0.3,,,,,1"
    AddCase "kweight,rounds-down,CA OVERTIME=0.2,,,,,0"
    AddCase "kweight,other-code-named,CA OVERTIME=2;CA DBLTIME=12,,,,,4|CA DBLTIME"
    AddCase "kweight,lieu-st-plain,CA LIEU ST=10;CA LIEU TKN=3,,,,,10|CA LIEU TKN"
    AddCase "kweight,empty,,,,,,0"
    AddCase "kweight,stat-and-min,CA STATWK 2.0=12;CA MIN 4=4,,,,,30"
    AddCase "kperiod,year-to-date,Time Period : 01/11/2026 - 10/17/2026,,,,,2026-01-11|2026-10-17"
    AddCase "kperiod,one-day,10/19/2026 - 10/19/2026,,,,,2026-10-19|2026-10-19"
    AddCase "kperiod,label-only,Time Period :,,,,,no dates"
    AddCase "kperiod,end-first,10/20/2026 - 10/19/2026,,,,,end before start"
    AddCase "kperiod,bad-day,2/30/2026 - 3/1/2026,,,,,no dates"
    AddCase "kperiod,one-date,Executed on : 10/21/2026 11:56 AM,,,,,no dates"
    AddCase "kkind,snapshot,2026-10-19,2026-10-19,2025-12-28,2026-12-26,,day"
    AddCase "kkind,year-to-date-jan-11,2026-01-11,2026-10-17,2025-12-28,2026-12-26,,cumulative"
    AddCase "kkind,year-to-date-dec-28,2025-12-28,2026-10-17,2025-12-28,2026-12-26,,cumulative"
    AddCase "kkind,starts-late,2026-01-12,2026-10-17,2025-12-28,2026-12-26,,range"
    AddCase "kkind,one-week,2026-10-13,2026-10-19,2025-12-28,2026-12-26,,range"
    AddCase "kkind,last-year,2025-11-01,2025-11-30,2025-12-28,2026-12-26,,outside"
    AddCase "kkind,next-year-day,2027-01-05,2027-01-05,2025-12-28,2026-12-26,,outside"
    AddCase "reconcile,same,Ann B;Bob C;Cal D,Ann B;Bob C;Cal D,,,,-|-|-"
    AddCase "reconcile,two-swapped,Ann B;Bob C;Cal D,Ann B;Cal D;Bob C,,,,2;3|-|-"
    AddCase "reconcile,case-and-spaces,ann b; BOB C,Ann B;Bob C,,,,-|-|-"
    AddCase "reconcile,paper-has-stranger,Ann B;Zed Q;Cal D,Ann B;Bob C;Cal D,,,,2|Zed Q|Bob C"
    AddCase "reconcile,paper-shorter,Ann B,Ann B;Bob C,,,,2|-|Bob C"
    AddCase "reconcile,roster-shorter,Ann B;Bob C,Ann B,,,,2|Bob C|-"
    AddCase "reconcile,blank-paper-line,Ann B;;Cal D,Ann B;Bob C;Cal D,,,,2|-|Bob C"
    AddCase "reconcile,empty-paper,,Ann B;Bob C,,,,1;2|-|Ann B;Bob C"
    AddCase "reconcile,both-empty,,,,,,-|-|-"
    AddCase "squarter,jan,2026-01-11,2026-01-01,,,,Q1"
    AddCase "squarter,mar-31,2026-03-31,2026-01-01,,,,Q1"
    AddCase "squarter,apr-1,2026-04-01,2026-01-01,,,,Q2"
    AddCase "squarter,sep,2026-09-15,2026-01-01,,,,Q3"
    AddCase "squarter,dec,2026-12-26,2026-01-01,,,,Q4"
    AddCase "sbucket,void-first,Callout,Accepted,VOID,,,void"
    AddCase "sbucket,callout-accepted,Callout,Accepted,Active,,,accepted"
    AddCase "sbucket,callout-refused,Callout,Refused,,,,refused"
    AddCase "sbucket,callout-message,Callout,Left Message,,,,message"
    AddCase "sbucket,callout-unavailable,Callout,Unavailable,,,,unavailable"
    AddCase "sbucket,callout-other,Callout,Not offered,,,,other"
    AddCase "sbucket,posting-awarded,Posting,Awarded,,,,awarded"
    AddCase "sbucket,posting-refused,Posting,Refused,,,,p-refused"
    AddCase "sbucket,posting-signed,Posting,Signed Up,,,,signedup"
    AddCase "sbucket,adjustment,Adjustment,,,,,adjustment"
    AddCase "sbucket,unknown-type,Kronos,,,,,other"
    AddCase "sband,midnight,0,,,,,night 0-6"
    AddCase "sband,five,5,,,,,night 0-6"
    AddCase "sband,six,6,,,,,morning 6-12"
    AddCase "sband,noon,12,,,,,afternoon 12-18"
    AddCase "sband,evening,19,,,,,evening 18-24"
    AddCase "sband,late,23,,,,,evening 18-24"
    AddCase "sspread,three,120;140;160,,,,,120.0|160.0|140.0|16.3|40.0"
    AddCase "sspread,one,100,,,,,100.0|100.0|100.0|0.0|0.0"
    AddCase "sspread,none,,,,,,-|-|-|-|-"
    AddCase "sspread,two-equal,50;50,,,,,50.0|50.0|50.0|0.0|0.0"
    AddCase "kperiod,digits-glued,110/17/2026 - 1/2/3/2026,,,,,no dates"
    AddCase "kkind,starts-last-year,2025-12-01,2026-10-17,2025-12-28,2026-12-26,,outside"
    AddCase "kweight,blank-hours,CA OVERTIME=,,,,,0"
    AddCase "squarter,dec-28-is-q1,2025-12-28,2026-01-01,,,,Q1"
    AddCase "squarter,dec-31-is-q1,2025-12-31,2026-01-01,,,,Q1"
    AddCase "sspread,quarter-hours,22.25;19.25,,,,,19.3|22.3|20.8|1.5|3.0"
    AddCase "bkeep,daily-fresh,daily,2026-10-01,2026-10-21,N,,keep"
    AddCase "bkeep,daily-30-days,daily,2026-09-21,2026-10-21,N,,keep"
    AddCase "bkeep,daily-31-days,daily,2026-09-20,2026-10-21,N,,drop"
    AddCase "bkeep,period-start-old,daily,2026-01-11,2026-10-21,Y,,keep"
    AddCase "bkeep,period-start-400-plus,daily,2025-09-01,2026-10-21,Y,,drop"
    AddCase "bkeep,hourly-fresh,hourly,2026-10-20,2026-10-21,N,,keep"
    AddCase "bkeep,hourly-7-days,hourly,2026-10-14,2026-10-21,N,,keep"
    AddCase "bkeep,hourly-8-days,hourly,2026-10-13,2026-10-21,N,,drop"
    AddCase "bhourly,hourly-name,Call Out Roster backup 2026-10-21 14.xlsm,,,,,2026-10-21 14"
    AddCase "bhourly,daily-is-not-hourly,Call Out Roster 2026-10-21.xlsm,,,,,none"
    AddCase "bhourly,bad-hour,Call Out Roster backup 2026-10-21 24.xlsm,,,,,none"
    AddCase "bhourly,pre-change-copy,Backup before kronos 2026-10-21 141500 - V2 - Call Out Roster.xlsm,,,,,none"
    AddCase "backupname,hourly-backup,Call Out Roster backup 2026-10-21 14.xlsm,,,,,yes"
    AddCase "bhourly,feb-29-leap,Call Out Roster backup 2028-02-29 06.xlsm,,,,,2028-02-29 06"
    AddCase "bhourly,feb-30-rejected,Call Out Roster backup 2026-02-30 06.xlsm,,,,,none"
    AddCase "firstref,session-row,  S-0012  Refused  Ann Bee  12 h  charged 21,,,,,S-0012"
    AddCase "firstref,adjustment,Adjustment ADJ-0003 added,,,,,ADJ-0003"
    AddCase "firstref,posting-award,P-0041 awarded to Bob Cee,,,,,P-0041"
    AddCase "firstref,second-wins-when-first-is-a-date,on 2026-10-22: SS-0007 saved,,,,,SS-0007"
    AddCase "firstref,leave-return,Leave return LV-0002 saved,,,,,LV-0002"
    AddCase "firstref,renamed-test-name,ZZ Test-1 took the shift,,,,,-"
    AddCase "firstref,lieu-code,LIEU OT - DT Split 4 h,,,,,-"
    AddCase "firstref,glued-letter,aS-12 is not a ref,,,,,-"
    AddCase "firstref,glued-digit,S-0012a is not a ref either,,,,,-"
End Sub

Private Sub LoadCases5()
    AddCase "firstref,five-letters,ABCDE-12 is too long a prefix,,,,,-"
    AddCase "firstref,lowercase,s-0012 lower case,,,,,-"
    AddCase "firstref,dash-only,nothing - here,,,,,-"
    AddCase "firstref,empty,,,,,,-"
    AddCase "firstref,entry-id-no-dash,#17 by piercyg,,,,,-"
    AddCase "firstref,at-the-end,voided NH-0009,,,,,NH-0009"
    AddCase "firstref,in-brackets,(see P-7),,,,,P-7"
    AddCase "sgsuspect,written-at-once,2027-01-10-09-00,2027-01-10-09-00,,,,ok"
    AddCase "sgsuspect,written-5-min-later,2027-01-10-09-00,2027-01-10-09-05,,,,ok"
    AddCase "sgsuspect,written-10-min-later-exactly,2027-01-10-09-00,2027-01-10-09-10,,,,ok"
    AddCase "sgsuspect,written-11-min-later,2027-01-10-09-00,2027-01-10-09-11,,,,suspect"
    AddCase "sgsuspect,written-a-day-later,2027-01-10-09-00,2027-01-11-09-00,,,,suspect"
    AddCase "sgsuspect,stamp-after-file-time,2027-01-10-09-30,2027-01-10-09-00,,,,ok"
    AddCase "sgsuspect,file-time-unknown,2027-01-10-09-00,0,,,,ok"
    AddCase "sgsuspect,stamp-unknown,0,2027-01-10-09-00,,,,ok"
    AddCase "sgmix,one-login,Ann:S:0:2027-01-10-09:jdoe,,,,,Ann"
    AddCase "sgmix,two-logins-same-name,Ann:S:0:2027-01-10-09:jdoe;Ann:S:8:2027-01-10-11:psmith,,,,,Ann!"
    AddCase "sgmix,other-login-earlier-withdraw,Ann:S:0:2027-01-10-09:psmith;Ann:W:0:2027-01-10-10:jdoe;Ann:S:0:2027-01-10-11:psmith,,,,,Ann!"
    AddCase "sgmix,two-names-two-logins,Ann:S:0:2027-01-10-09:jdoe;Bob:S:0:2027-01-10-09:psmith,,,,,Ann;Bob"
    AddCase "sgmix,blank-login-ignored,Ann:S:0:2027-01-10-09:;Ann:S:0:2027-01-10-11:jdoe,,,,,Ann"
    AddCase "sgmix,newest-is-withdraw,Ann:S:0:2027-01-10-09:jdoe;Ann:W:0:2027-01-10-11:psmith,,,,,none"
    AddCase "sgmix,case-of-login-ignored,Ann:S:0:2027-01-10-09:JDoe;Ann:S:0:2027-01-10-11:jdoe,,,,,Ann"
' <<< CASES
End Sub
