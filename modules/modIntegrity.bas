Attribute VB_Name = "modIntegrity"
Option Explicit

' ====================================================================
'  INTEGRITY CHECK (v23)
'
'  Reads the roster, the OT Log and the other tables and reports
'  anything that doesn't add up. It only READS - it never changes data.
'
'    Alt+F8 > IntegrityCheck   full report on the Integrity sheet (admin)
'    on open (admin)           quiet run; result on the roster status
'                              line: "Integrity: OK" / "n PROBLEMS"
'    RolloverYear              runs it first and stops on any PROBLEM
'
'  Three levels:
'    PROBLEM  totals can't be trusted until it is fixed (blocks the
'             year-end rollover): roster <> log, a typed number or a
'             missing formula in a pay-period cell, a duplicate ID,
'             an ID counter below an archived year's highest (v26).
'    Check    worth a look, nothing is wrong with the hours: people
'             no longer on the roster, a posting whose list and log
'             disagree, a calendar line pointing at nothing, an old
'             payback, old Kronos data.
'    Note     information only.
'
'  The pure parts (PeriodCellCode, DuplicatesIn) take plain values so
'  SelfTest can hand them a planted bad cell / duplicate ID in memory.
' ====================================================================

Public Const IC_VERSION As String = "Integrity v26.5.5 (Oct 21)"
Public Const IC_SHEET As String = "Integrity"

Private Const SEV_ERR As String = "PROBLEM"
Private Const SEV_WARN As String = "Check"
Private Const SEV_INFO As String = "Note"
Private Const MAX_ISSUES As Long = 400
Private Const TOL As Double = 0.001

Private mIss() As String          ' (1 To 4, 1 To n): level, check, what, where / what to do
Private mN As Long
Private mErr As Long
Private mWarn As Long
Private mRanAt As Date            ' last run this session (0 = not run)
Private mLastErr As Long
Private mLastWarn As Long

' data shared by the checks (loaded once per run)
Private mLog As Variant           ' OT Log rows, or Empty
Private mNames() As String        ' roster names by row (index 1 = R_FIRST), "" for a blank row
Private mRows As Long
Private mVals As Variant          ' roster values A..last period, rows R_FIRST..last
Private mLastPer As Long
Private mYearFrom As Date
Private mLastStart As Date
Private mYearTo As Date
Private mOnRoster As Object       ' LCase(name) -> True


' ====================================================================
'  PURE HELPERS (SelfTest "cell" and "dups")
' ====================================================================

' What is wrong with one pay-period cell: "" = fine, else a short code
'   notlog       a formula that doesn't read the OT Log
'   error        an error value
'   empty        nothing there (the formula is missing)
'   textformula  the formula was stored as text
'   text         typed text
'   typed        a typed number
Public Function PeriodCellCode(ByVal hasFormula As Boolean, ByVal formulaText As String, ByVal v As Variant) As String
    If hasFormula Then
        If InStr(1, formulaText, LOG_TABLE, vbTextCompare) = 0 Then PeriodCellCode = "notlog"
        Exit Function
    End If
    If IsError(v) Then PeriodCellCode = "error": Exit Function
    If IsEmpty(v) Then PeriodCellCode = "empty": Exit Function
    If VarType(v) = vbString Then
        If Left$(CStr(v), 1) = "=" Then PeriodCellCode = "textformula" Else PeriodCellCode = "text"
        Exit Function
    End If
    PeriodCellCode = "typed"
End Function

Private Function CellCodeWords(ByVal code As String) As String
    Select Case code
        Case "notlog": CellCodeWords = "a formula that doesn't read the OT Log"
        Case "error": CellCodeWords = "an error value"
        Case "empty": CellCodeWords = "empty (the formula is missing)"
        Case "textformula": CellCodeWords = "a formula stored as text"
        Case "text": CellCodeWords = "typed text"
        Case "typed": CellCodeWords = "a typed number"
        Case Else: CellCodeWords = code
    End Select
End Function

' Keys that occur more than once: "S-0004 x2; P-0010 x3" in first-seen
' order ("" = none). Blank keys are ignored; case and outer spaces don't count.
Public Function DuplicatesIn(ByRef keys() As String, ByVal n As Long) As String
    If n < 2 Then Exit Function
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim i As Long, k As String, kk As Variant, s As String
    For i = 1 To n
        k = Trim$(keys(i))
        If k <> "" Then
            If d.Exists(k) Then d(k) = d(k) + 1 Else d.Add k, 1
        End If
    Next i
    For Each kk In d.Keys
        If d(kk) > 1 Then s = s & IIf(s = "", "", "; ") & CStr(kk) & " x" & d(kk)
    Next kk
    DuplicatesIn = s
End Function


' ====================================================================
'  RUN
' ====================================================================
Private Sub AddIssue(ByVal sev As String, ByVal chk As String, ByVal what As String, ByVal whereTxt As String)
    If sev = SEV_ERR Then mErr = mErr + 1
    If sev = SEV_WARN Then mWarn = mWarn + 1
    If mN >= MAX_ISSUES Then Exit Sub
    mN = mN + 1
    If mN > UBound(mIss, 2) Then ReDim Preserve mIss(1 To 4, 1 To mN + 50)
    mIss(1, mN) = sev: mIss(2, mN) = chk: mIss(3, mN) = what: mIss(4, mN) = whereTxt
End Sub

' Runs every check. Returns the number of lines found; nErr / nWarn are
' the PROBLEM and Check counts. Never changes anything.
Public Function RunIntegrity(ByRef nErr As Long, ByRef nWarn As Long) As Long
    mN = 0: mErr = 0: mWarn = 0
    ReDim mIss(1 To 4, 1 To 50)
    Dim k As Long
    If LoadData() Then
        For k = 1 To 13
            RunCheck k
        Next k
    End If
    mLog = Empty: mVals = Empty: Set mOnRoster = Nothing
    nErr = mErr: nWarn = mWarn
    mRanAt = Now: mLastErr = mErr: mLastWarn = mWarn
    RunIntegrity = mN
End Function

Private Function CheckName(ByVal k As Long) As String
    Select Case k
        Case 1: CheckName = "Roster = log"
        Case 2: CheckName = "Typed values"
        Case 3: CheckName = "IDs"
        Case 4: CheckName = "Not on roster"
        Case 5: CheckName = "Postings"
        Case 6: CheckName = "Calendar"
        Case 7: CheckName = "Paybacks"
        Case 8: CheckName = "Kronos"
        Case 9: CheckName = "Log types"
        Case 10: CheckName = "Sign-ups"
        Case 11: CheckName = "Planned moves"
        Case 12: CheckName = "Coverage = log"
        Case 13: CheckName = "Settings"
    End Select
End Function

' One check; a check that can't run is itself a PROBLEM (the state is unknown)
Private Sub RunCheck(ByVal k As Long)
    On Error GoTo Oops
    Select Case k
        Case 1: CheckRosterVsLog
        Case 2: CheckTypedValues
        Case 3: CheckIds
        Case 4: CheckOrphans
        Case 5: CheckPostings
        Case 6: CheckCalendarLinks
        Case 7: CheckPaybacks
        Case 8: CheckKronos
        Case 9: CheckLogTypes
        Case 10: CheckSignups
        Case 11: CheckPlannedMoves
        Case 12: CheckCoverageVsLog
        Case 13: CheckSettings
    End Select
    Exit Sub
Oops:
    AddIssue SEV_ERR, CheckName(k), "This check could not run: " & Err.Description, "Send this line to the admin"
End Sub

Private Function LoadData() As Boolean
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then
        AddIssue SEV_ERR, "Roster = log", "The OT Log table wasn't found", "Run SystemCheck"
        Exit Function
    End If
    mLog = Empty
    If Not lo.DataBodyRange Is Nothing Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then mLog = RangeTo2D(lo.DataBodyRange)
    End If
    mLastPer = PeriodLastCol()
    mYearFrom = RosterYearFrom()
    mLastStart = Int(CDate(ws.Cells(PStartRow(), mLastPer).Value))
    mYearTo = RosterYearTo()
    Set mOnRoster = CreateObject("Scripting.Dictionary")
    mOnRoster.CompareMode = 1
    mRows = 0
    If lastRow >= R_FIRST Then
        mRows = lastRow - R_FIRST + 1
        mVals = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, mLastPer)).Value2
        ReDim mNames(1 To mRows)
        Dim i As Long
        For i = 1 To mRows
            mNames(i) = Trim$(SafeStr(mVals(i, C_NAME)))
            If mNames(i) <> "" Then
                If mOnRoster.Exists(mNames(i)) Then
                    AddIssue SEV_ERR, "Roster = log", mNames(i) & " is on the roster twice (row " & (R_FIRST + i - 1) & ")", _
                             "Both rows show the same hours and the call order lists the person twice; remove or rename one"   ' v26.5.5 (A14): a PROBLEM
                Else
                    mOnRoster.Add mNames(i), True
                End If
            End If
        Next i
    End If
    LoadData = True
    Exit Function
Oops:
    AddIssue SEV_ERR, "Roster = log", "The roster or the OT Log couldn't be read: " & Err.Description, "Run SystemCheck"
End Function

Private Function LogRows() As Long
    If IsArray(mLog) Then LogRows = UBound(mLog, 1)
End Function

Private Function IsVoidRow(ByVal i As Long) As Boolean
    IsVoidRow = (UCase$(Trim$(SafeStr(mLog(i, L_STATUS)))) = "VOID")
End Function

' Period Start of a log row as a whole number (0 = not a date)
Private Function RowPeriod(ByVal i As Long) As Long
    Dim v As Variant: v = mLog(i, L_PERIOD)
    If IsError(v) Then Exit Function
    If IsDate(v) Then
        RowPeriod = CLng(Int(CDbl(CDate(v))))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        RowPeriod = CLng(Int(CDbl(v)))
    End If
End Function

' ---- 1. roster = log ----
Private Sub CheckRosterVsLog()
    If mRows = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    sums.CompareMode = 1
    Dim prov As Object: Set prov = CreateObject("Scripting.Dictionary")
    prov.CompareMode = 1
    Dim through As Date: through = KronosDataThrough()
    Dim i As Long, key As String, emp As String, nOut As Long, nTxt As Long
    For i = 1 To LogRows()
        emp = Trim$(SafeStr(mLog(i, L_EMP)))
        If VarType(mLog(i, L_PERIOD)) = vbString Then
            If Trim$(CStr(mLog(i, L_PERIOD))) <> "" Then nTxt = nTxt + 1
        End If
        If emp <> "" And Not IsVoidRow(i) Then
            key = emp & "|" & RowPeriod(i)
            sums(key) = NumOr0(sums(key)) + NumOr0(mLog(i, L_CHARGED))
            If RowPeriod(i) < CLng(mYearFrom) Or RowPeriod(i) > CLng(mLastStart) Then nOut = nOut + 1
            If IsDate(mLog(i, L_DATE)) Then
                If ProvisionalCounts(SafeStr(mLog(i, L_OUTCOME)), SafeStr(mLog(i, L_STATUS)), CDate(mLog(i, L_DATE)), _
                                     IsNightShift(SafeStr(mLog(i, L_SHIFTLBL))), through, mYearFrom, mYearTo) Then
                    prov(emp) = NumOr0(prov(emp)) + NumOr0(mLog(i, L_BASEHRS))
                End If
            End If
        End If
    Next i
    ' period start per column
    Dim per() As Long: ReDim per(C_PERIOD1 To mLastPer)
    Dim c As Long, hv As Variant, sRow As Long: sRow = PStartRow()
    For c = C_PERIOD1 To mLastPer
        hv = ws.Cells(sRow, c).Value2
        If IsNumeric(hv) And Not IsEmpty(hv) Then per(c) = CLng(Int(CDbl(hv)))
    Next c
    Dim provLive As Boolean: provLive = ProvisionalLive()
    Dim nOldForm As Long                                 ' v26.5: rows whose Refused formula is the pre-v26.5 kind
    Dim thr As Date: thr = RefusedThrough()              ' v26.5: the STORED cap, as the formulas read it (the check changes nothing)
    Dim r As Long, expSum As Double, expCell As Double, act As Variant, bad As String, nBad As Long
    Dim vI As Variant, vG As Variant, vH As Variant, vC As Variant
    For r = 1 To mRows
        If mNames(r) <> "" Then
            expSum = 0: bad = "": nBad = 0
            ' v26.5: I counts the periods up to the current one - unless the row still carries the pre-v26.5 =SUM(...)
            ' formula (the first open after the install, before RefreshRosterFormulas): then the old rule is the
            ' expected figure and one Check line says to run the refresh
            Dim oldForm As Boolean: oldForm = OldRefusedFormula(ws.Cells(R_FIRST + r - 1, C_REFUSED))
            If oldForm Then nOldForm = nOldForm + 1
            For c = C_PERIOD1 To mLastPer
                expCell = NumOr0(sums(mNames(r) & "|" & per(c)))
                If oldForm Or RefusedCounts(CDate(per(c)), thr) Then expSum = expSum + expCell
                act = mVals(r, c)
                If IsError(act) Then
                    nBad = nBad + 1
                    If nBad <= 3 Then bad = bad & IIf(bad = "", "", ", ") & Format$(CDate(per(c)), "mmm d") & " shows an error"
                ElseIf Abs(NumOr0(act) - expCell) > TOL Then
                    nBad = nBad + 1
                    If nBad <= 3 Then bad = bad & IIf(bad = "", "", ", ") & Format$(CDate(per(c)), "mmm d") & " shows " & _
                                            NumOr0(act) & ", the log gives " & expCell
                End If
            Next c
            If nBad > 0 Then
                AddIssue SEV_ERR, "Roster = log", mNames(r) & ": " & bad & IIf(nBad > 3, " ... and " & (nBad - 3) & " more periods", ""), _
                         "Row " & (R_FIRST + r - 1) & ". Run RefreshRosterFormulas; if it still differs, check the name spelling in the OT Log"
            End If
            vI = mVals(r, C_REFUSED): vG = mVals(r, C_TOTAL): vH = mVals(r, C_WORKED): vC = mVals(r, C_PROV)
            If IsError(vI) Or IsError(vG) Or IsError(vH) Or IsError(vC) Then
                AddIssue SEV_ERR, "Roster = log", mNames(r) & ": Total, Worked, Refused or Provisional shows an error", _
                         "Row " & (R_FIRST + r - 1) & ". Run RefreshRosterFormulas"
            Else
                If Abs(NumOr0(vI) - expSum) > TOL Then
                    AddIssue SEV_ERR, "Roster = log", mNames(r) & ": Refused (I) shows " & NumOr0(vI) & ", the log gives " & expSum, _
                             "Row " & (R_FIRST + r - 1) & ". Run RefreshRosterFormulas"
                End If
                If provLive Then
                    If Abs(NumOr0(vC) - NumOr0(prov(mNames(r)))) > TOL Then
                        AddIssue SEV_ERR, "Roster = log", mNames(r) & ": provisional accepted OT (C) shows " & NumOr0(vC) & _
                                 ", the log gives " & NumOr0(prov(mNames(r))), "Row " & (R_FIRST + r - 1) & ". Run RefreshRosterFormulas"
                    End If
                End If
                If Abs(NumOr0(vG) - (NumOr0(vH) + NumOr0(vI) + IIf(provLive, NumOr0(vC), 0))) > TOL Then
                    AddIssue SEV_ERR, "Roster = log", mNames(r) & ": Total (G) shows " & NumOr0(vG) & " but Worked + Refused" & _
                             IIf(provLive, " + Provisional", "") & " is " & (NumOr0(vH) + NumOr0(vI) + IIf(provLive, NumOr0(vC), 0)), _
                             "Row " & (R_FIRST + r - 1) & ". Run RefreshRosterFormulas"
                End If
            End If
        End If
    Next r
    If nTxt > 0 Then
        AddIssue SEV_WARN, "Roster = log", nTxt & " OT Log row(s) have a Period Start stored as text, not as a date", _
                 "They may not count in the roster and will not move at the year-end rollover. Tell the admin before the rollover"
    End If
    If nOut > 0 Then
        AddIssue SEV_INFO, "Roster = log", nOut & " log row(s) belong to a pay period that isn't on this roster", _
                 "Expected for entries logged for another year; they don't count in these totals"
    End If
    If nOldForm > 0 Then                                 ' v26.5
        AddIssue SEV_WARN, "Roster = log", nOldForm & " roster row(s) still carry the pre-v26.5 Refused formula (every period summed)", _
                 "Run RefreshRosterFormulas once (Refused then counts the pay periods up to the current one)"
    End If
End Sub

' ---- 2. typed values / missing formulas in the pay-period cells ----
Private Sub CheckTypedValues()
    If mRows = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long, c As Long, rowRng As Range, hf As Variant, fm As Variant, code As String, bad As String, nBad As Long
    For r = 1 To mRows
        If mNames(r) <> "" Then
            Set rowRng = ws.Range(ws.Cells(R_FIRST + r - 1, C_PERIOD1), ws.Cells(R_FIRST + r - 1, mLastPer))
            hf = rowRng.HasFormula                    ' True = every cell, False = none, Null = mixed
            fm = rowRng.Formula
            bad = "": nBad = 0
            For c = C_PERIOD1 To mLastPer
                If VarType(hf) = vbBoolean Then
                    code = PeriodCellCode(CBool(hf), CStr(fm(1, c - C_PERIOD1 + 1)), mVals(r, c))
                Else
                    code = PeriodCellCode(ws.Cells(R_FIRST + r - 1, c).HasFormula, CStr(fm(1, c - C_PERIOD1 + 1)), mVals(r, c))
                End If
                If code <> "" Then
                    nBad = nBad + 1
                    If nBad <= 3 Then bad = bad & IIf(bad = "", "", ", ") & ColLetter(c) & (R_FIRST + r - 1) & " holds " & CellCodeWords(code)
                End If
            Next c
            If nBad > 0 Then
                AddIssue SEV_ERR, "Typed values", mNames(r) & ": " & bad & IIf(nBad > 3, " ... and " & (nBad - 3) & " more cells", ""), _
                         "Pay-period cells must be the log formula. Run RefreshRosterFormulas (it moves typed numbers into the OT Log)"
            End If
            ' v25 (D-E1/E2): a seniority that isn't a date sorts last in the call order but not under the sort buttons
            ' (v25 review: a real date arrives as a Double through .Value2, so IsNumeric passes it; text fails both)
            If VarType(mVals(r, C_SENIORITY)) = vbString Then
                If Trim$(CStr(mVals(r, C_SENIORITY))) <> "" Then
                    AddIssue SEV_WARN, "Typed values", mNames(r) & ": Plant Seniority '" & SafeStr(mVals(r, C_SENIORITY)) & "' is text, not a date", _
                             "Type it as a date (text sorts one way under the sort buttons and another in the call order)"
                End If
            ElseIf Not IsEmpty(mVals(r, C_SENIORITY)) And Not IsDate(mVals(r, C_SENIORITY)) And Not IsNumeric(mVals(r, C_SENIORITY)) Then
                AddIssue SEV_WARN, "Typed values", mNames(r) & ": Plant Seniority '" & SafeStr(mVals(r, C_SENIORITY)) & "' isn't a date", _
                         "Type it as a date (the call order puts a non-date last; the sort buttons don't)"
            End If
        End If
    Next r
    ' v25 (D-E1/E2): two code letters in one unit column (D1 beside P1) sort one way by the buttons and another in the call order
    Dim letters As String, firstL As String, cc As Long, nm As String, cd As String
    For cc = UnitColFirst() To UnitColLast()
        letters = "": firstL = ""
        For r = 1 To mRows
            cd = UCase$(Trim$(SafeStr(mVals(r, cc))))
            If cd <> "" And mNames(r) <> "" Then
                If firstL = "" Then firstL = Left$(cd, 1)
                If Left$(cd, 1) <> firstL And InStr(letters, Left$(cd, 1)) = 0 Then letters = letters & Left$(cd, 1)
            End If
        Next r
        If letters <> "" Then
            AddIssue SEV_WARN, "Typed values", "Column " & ColLetter(cc) & " mixes code letters (" & firstL & " with " & letters & ")", _
                     "One letter per unit column keeps the sort buttons and the call order in step"
        End If
    Next cc
End Sub

' ---- 3. IDs ----
Private Sub CheckIds()
    Dim keys() As String, n As Long, i As Long, s As String
    ' v24.2: the same Payroll # on two roster rows (the names are shown, never the number)
    Dim seenPay As Object, pt As String
    Set seenPay = CreateObject("Scripting.Dictionary")
    For i = 1 To mRows
        If mNames(i) <> "" Then
            pt = PayrollText(mVals(i, C_PAYROLL))
            If pt <> "" Then
                If seenPay.Exists(pt) Then
                    AddIssue SEV_WARN, "IDs", mNames(i) & " and " & CStr(seenPay(pt)) & " have the same Payroll # in roster column A", _
                             "Correct column A for one of them (admin view); the next Kronos import fills a blank"
                Else
                    seenPay.Add pt, mNames(i)
                End If
            End If
        End If
    Next i
    ' Entry IDs
    n = LogRows()
    If n > 0 Then
        ReDim keys(1 To n)
        For i = 1 To n
            keys(i) = SafeStr(mLog(i, L_ENTRY))
        Next i
        s = DuplicatesIn(keys, n)
        If s <> "" Then AddIssue SEV_ERR, "IDs", "Entry ID used more than once in the OT Log: " & s, "Send this line to the admin before logging anything else"
        ' a leave topped up twice, and one Ref ID covering two different events
        Dim lv() As String, nLv As Long, sig As Object, ref As String, sg As String, clash As String
        ReDim lv(1 To n)
        Set sig = CreateObject("Scripting.Dictionary")
        sig.CompareMode = 1
        For i = 1 To n
            ref = UCase$(Trim$(SafeStr(mLog(i, L_REF))))
            If Not IsVoidRow(i) Then
                If Left$(ref, 3) = "LV-" Then nLv = nLv + 1: lv(nLv) = ref
                If Left$(ref, 2) = "S-" Or Left$(ref, 2) = "P-" Then
                    sg = SafeStr(mLog(i, L_TYPE)) & "|" & SafeStr(mLog(i, L_DATE)) & "|" & SafeStr(mLog(i, L_DEPT)) & "|" & SafeStr(mLog(i, L_SHIFTLBL))
                    If Not sig.Exists(ref) Then
                        sig.Add ref, sg
                    ElseIf sig(ref) <> sg And sig(ref) <> "*" Then
                        sig(ref) = "*"
                        clash = clash & IIf(clash = "", "", ", ") & ref
                    End If
                End If
            End If
        Next i
        s = DuplicatesIn(lv, nLv)
        If s <> "" Then AddIssue SEV_ERR, "IDs", "Leave return logged more than once: " & s, "Void the extra Adjustment (Void Entry)"
        If clash <> "" Then
            AddIssue SEV_WARN, "IDs", "One Ref ID covers rows with different dates, units or shifts: " & clash, _
                     "Usually an entry corrected by hand; check those rows in the OT Log"
        End If
    End If
    TableIdCheck PT_TABLE, PS_ID, "posting"
    TableIdCheck SS_TABLE, SW_ID, "shift switch"
    TableIdCheck AB_TABLE, AB_ID, "absence / sick call"
    TableIdCheck CV_TABLE, CV_ID, "calendar line"
    ' v26 (I8): the IdMark_ counters (written at the rollover, read by NextRefID / NextEntryID) against
    ' the archived years' highest numbers: a counter below an archive means the next ID of that prefix
    ' could repeat an archived one (the live log alone can't show it once the old rows have moved)
    Dim mx As Object: Set mx = CreateObject("Scripting.Dictionary")
    mx.CompareMode = 1
    Dim mxE As Long, arcList As String, kk As Variant
    ArchiveMaxima mx, mxE, arcList
    For Each kk In mx.Keys
        If IdMark(CStr(kk)) < CLng(mx(kk)) Then
            AddIssue SEV_ERR, "IDs", "The ID counter for " & CStr(kk) & "- is " & Format$(IdMark(CStr(kk)), "0000") & _
                     " but an archived year (" & arcList & ") holds " & CStr(kk) & "-" & Format$(CLng(mx(kk)), "0000") & _
                     ": the next " & CStr(kk) & "- ID could repeat an archived one", "Run RepairIdMarks (admin, Alt+F8), then IntegrityCheck again"
        End If
    Next kk
    If mxE > 0 And IdMark("ENTRY") < mxE Then
        AddIssue SEV_ERR, "IDs", "The Entry ID counter is " & IdMark("ENTRY") & " but an archived year (" & arcList & ") holds #" & mxE & _
                 ": the next Entry ID could repeat an archived one", "Run RepairIdMarks (admin, Alt+F8), then IntegrityCheck again"
    End If
End Sub

Private Sub TableIdCheck(ByVal tableName As String, ByVal idCol As Long, ByVal what As String)
    Dim v As Variant: v = TableValues(tableName)
    If Not IsArray(v) Then Exit Sub
    Dim keys() As String, i As Long, s As String
    ReDim keys(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        keys(i) = SafeStr(v(i, idCol))
    Next i
    s = DuplicatesIn(keys, UBound(v, 1))
    If s <> "" Then AddIssue SEV_ERR, "IDs", "The same " & what & " ID is on two rows: " & s, "Send this line to the admin"
End Sub

' Any table in the workbook as a 2-D array (Empty when missing or empty)
Private Function TableValues(ByVal tableName As String) As Variant
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
    TableValues = RangeTo2D(lo.DataBodyRange)
End Function

' ---- 10. electronic sign-up files left in the inbox for a closed posting (v24.1) ----
Private Sub CheckSignups()
    Dim txt As String, n As Long
    n = StaleSignups(SignupStaleDays(), txt)
    If n > 0 Then
        AddIssue SEV_WARN, "Sign-ups", n & " sign-up file(s) over " & SignupStaleDays() & " days old are still in the inbox for postings that aren't open (" & txt & ")", _
                 "Move them into the Done folder inside " & SignupFolderPath()
    End If
End Sub

' ---- 9. charged rows whose Type the refused split doesn't know (v24.1) ----
' Call-out + Posting + Other = Refused always holds, because Other is "everything
' else". A charge on a row with an unexpected Type would land in Other unseen.
Private Sub CheckLogTypes()
    Dim i As Long, typ As String, seen As Object, kk As Variant
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For i = 1 To LogRows()
        If Not IsVoidRow(i) And NumOr0(mLog(i, L_CHARGED)) <> 0 Then
            If RowPeriod(i) >= CLng(mYearFrom) And RowPeriod(i) <= CLng(mLastStart) Then
                typ = Trim$(SafeStr(mLog(i, L_TYPE)))
                Select Case LCase$(typ)
                    Case "callout", "posting", "adjustment", "opening balance", "manual entry"
                    Case Else: seen(IIf(typ = "", "(blank)", typ)) = NumOr0(seen(IIf(typ = "", "(blank)", typ))) + 1
                End Select
            End If
        End If
    Next i
    For Each kk In seen.Keys
        AddIssue SEV_WARN, "Log types", seen(kk) & " charged OT Log row(s) this year have the Type '" & CStr(kk) & "'", _
                 "They count in Refused under 'other'. Send " & PRIMARY_ADMIN_NAME & " this line if you didn't expect it"
    Next kk
    ' call-out + posting + other must equal the roster's Refused (I) for every person
    Dim splitMap As Object, splitAll As Object, useMap As Object, c As Double, p As Double, o As Double, r As Long, nBad As Long
    ' v26.5: Refused (I) counts the periods up to the current one, so the split is capped the same way (SplitNow's
    ' rule); a row still carrying the pre-v26.5 =SUM formula is compared with the whole year
    Dim capStart As Long: capStart = CLng(mLastStart)
    Dim thrI As Long: thrI = CLng(Int(RefusedThrough()))   ' the stored cap, as the formulas read it
    If thrI > 0 And thrI < capStart Then capStart = thrI
    Set splitMap = SplitTotals(mLog, CLng(mYearFrom), capStart)
    Set splitAll = SplitTotals(mLog, CLng(mYearFrom), CLng(mLastStart))
    Dim wsR As Worksheet: Set wsR = RosterSheet()
    For r = 1 To mRows
        If mNames(r) <> "" And Not IsError(mVals(r, C_REFUSED)) Then
            If OldRefusedFormula(wsR.Cells(R_FIRST + r - 1, C_REFUSED)) Then Set useMap = splitAll Else Set useMap = splitMap
            RefusedSplitOf useMap, mNames(r), c, p, o
            If Abs(c + p + o - NumOr0(mVals(r, C_REFUSED))) > TOL Then
                nBad = nBad + 1
                If nBad <= 10 Then
                    AddIssue SEV_WARN, "Log types", mNames(r) & ": call-out " & c & " + posting " & p & " + other " & o & _
                             " = " & (c + p + o) & ", but Refused on the roster is " & NumOr0(mVals(r, C_REFUSED)), _
                             "A log row's Period Start is not one of the roster's pay-period dates. The status bar, the report and the members' file show the higher figure"
                End If
            End If
        End If
    Next r
    If nBad > 10 Then AddIssue SEV_WARN, "Log types", "... and " & (nBad - 10) & " more people whose refused split doesn't add up", ""
End Sub

' ---- 4. people in the data who aren't on the roster ----
Private Sub CheckOrphans()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Dim hrs As Object: Set hrs = CreateObject("Scripting.Dictionary")
    hrs.CompareMode = 1
    Dim i As Long, emp As String, kk As Variant, nShown As Long
    Dim former As Object: Set former = FormerNameSet()       ' v24: moved to the Former sheet = known, their log rows stay
    For i = 1 To LogRows()
        emp = Trim$(SafeStr(mLog(i, L_EMP)))
        If emp <> "" And Not IsVoidRow(i) Then
            If RowPeriod(i) >= CLng(mYearFrom) And RowPeriod(i) <= CLng(mLastStart) Then
                If Not mOnRoster.Exists(emp) And Not former.Exists(emp) Then
                    seen(emp) = NumOr0(seen(emp)) + 1
                    hrs(emp) = NumOr0(hrs(emp)) + NumOr0(mLog(i, L_CHARGED))
                End If
            End If
        End If
    Next i
    For Each kk In seen.Keys
        nShown = nShown + 1
        If nShown <= 15 Then
            AddIssue SEV_WARN, "Not on roster", CStr(kk) & " has " & seen(kk) & " OT Log row(s) this year (" & hrs(kk) & _
                     " h charged) but isn't on the roster", "If they left, put the roster row back and use Employee > Remove. If the name was typed over by hand, type the old name back and use Employee > Rename" & ByHandHint(RosterSheet().Name)
        End If
    Next kk
    If nShown > 15 Then AddIssue SEV_WARN, "Not on roster", "... and " & (nShown - 15) & " more names in the OT Log", ""

    Dim v As Variant, e2 As Date
    v = TableValues(AB_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            emp = Trim$(SafeStr(v(i, AB_EMP)))
            If emp <> "" And UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" Then
                If Not mOnRoster.Exists(emp) Then
                    If IsDate(v(i, AB_END)) Then e2 = CDate(v(i, AB_END)) Else e2 = DateSerial(2999, 12, 31)
                    If e2 >= Date Then
                        If former.Exists(emp) Then
                            AddIssue SEV_WARN, "Not on roster", "Booking " & SafeStr(v(i, AB_ID)) & " is for " & emp & ", who has left (Former sheet)", _
                                     "Void the booking in Book Time Off so the calendar stops showing it"
                        Else
                            AddIssue SEV_WARN, "Not on roster", "Booking " & SafeStr(v(i, AB_ID)) & " is for " & emp & ", who isn't on the roster", _
                                     "Void it in Book Time Off, or fix the name"
                        End If
                    End If
                End If
            End If
        Next i
    End If
    v = TableValues(SS_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" And IsDate(v(i, SW_DATE)) Then
                If CDate(v(i, SW_DATE)) >= Date Then
                    OrphanSwitchName SafeStr(v(i, SW_ID)), Trim$(SafeStr(v(i, SW_OFF)))
                    OrphanSwitchName SafeStr(v(i, SW_ID)), Trim$(SafeStr(v(i, SW_WORKS)))
                End If
            End If
        Next i
    End If
    ' v26.5.5 (A8): the Worked Periods rows (the Kronos export's per-period hours) name people too; a name
    ' that is on neither the roster nor the Former sheet is a Kronos spelling or a rename that missed them
    v = TableValues(WP_TABLE)
    If IsArray(v) Then
        Dim wpSeen As Object: Set wpSeen = CreateObject("Scripting.Dictionary")
        wpSeen.CompareMode = 1
        For i = 1 To UBound(v, 1)
            emp = Trim$(SafeStr(v(i, 2)))
            If emp <> "" Then
                If Not mOnRoster.Exists(emp) And Not former.Exists(emp) And Not wpSeen.Exists(emp) Then
                    wpSeen.Add emp, True
                    If wpSeen.Count <= 10 Then
                        AddIssue SEV_WARN, "Not on roster", "Worked Periods has rows for " & emp & ", who isn't on the roster or the Former sheet", _
                                 "Kronos spells the name differently (Rename matches it), or they left: the rows are rebuilt at the next import"
                    End If
                End If
            End If
        Next i
        If wpSeen.Count > 10 Then AddIssue SEV_WARN, "Not on roster", "... and " & (wpSeen.Count - 10) & " more names on Worked Periods", ""
    End If
    ' v27: the Kronos ledger, snapshot and change log name people too (the roster's spelling, written by the import)
    OrphanTableNames WD_TABLE, 2, "Worked Days", former
    OrphanTableNames KT_TABLE, 2, "Kronos Totals", former
    OrphanTableNames WC_TABLE, 3, "Worked Changes", former
End Sub

' v27: names in one table's column that are on neither the roster nor the Former sheet (a rename that missed them)
Private Sub OrphanTableNames(ByVal tableName As String, ByVal col As Long, ByVal label As String, ByVal former As Object)
    Dim v As Variant: v = TableValues(tableName)
    If Not IsArray(v) Then Exit Sub
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Dim i As Long, emp As String
    For i = 1 To UBound(v, 1)
        emp = Trim$(SafeStr(v(i, col)))
        If emp <> "" Then
            If Not mOnRoster.Exists(emp) And Not former.Exists(emp) And Not seen.Exists(emp) Then
                seen.Add emp, True
                If seen.Count <= 10 Then AddIssue SEV_WARN, "Not on roster", label & " has rows for " & emp & ", who isn't on the roster or the Former sheet", _
                                                 "A rename that missed the table, or they left: Rename matches it; the rows stay as history"
            End If
        End If
    Next i
    If seen.Count > 10 Then AddIssue SEV_WARN, "Not on roster", "... and " & (seen.Count - 10) & " more names on " & label, ""
End Sub

Private Sub OrphanSwitchName(ByVal id As String, ByVal emp As String)
    If emp = "" Then Exit Sub
    If mOnRoster.Exists(emp) Then Exit Sub
    AddIssue SEV_WARN, "Not on roster", "Switch " & id & " names " & emp & ", who isn't on the roster", "Void the switch or fix the name"
End Sub

' ---- 5. postings: the posting's own list against the log ----
Private Sub CheckPostings()
    Dim pv As Variant: pv = TableValues(PT_TABLE)
    Dim lv As Variant: lv = TableValues(PL_TABLE)
    Dim listSum As Object: Set listSum = CreateObject("Scripting.Dictionary")
    listSum.CompareMode = 1
    Dim logSum As Object: Set logSum = CreateObject("Scripting.Dictionary")
    logSum.CompareMode = 1
    Dim logAny As Object: Set logAny = CreateObject("Scripting.Dictionary")
    logAny.CompareMode = 1
    Dim i As Long, id As String, st As String
    ' v24.3.1 (review P1): list rows whose posting row is gone (a header deleted by hand, or a
    ' creation that failed between the list and the header). NextPostingID skips their ID now.
    Dim hdr As Object: Set hdr = CreateObject("Scripting.Dictionary")
    hdr.CompareMode = 1
    Dim orphan As Object: Set orphan = CreateObject("Scripting.Dictionary")
    orphan.CompareMode = 1
    If IsArray(pv) Then
        For i = 1 To UBound(pv, 1)
            id = Trim$(SafeStr(pv(i, PS_ID)))
            If id <> "" Then hdr(id) = True
        Next i
    End If
    If IsArray(lv) Then
        For i = 1 To UBound(lv, 1)
            id = Trim$(SafeStr(lv(i, PL_ID)))
            If id <> "" Then
                listSum(id) = NumOr0(listSum(id)) + NumOr0(lv(i, PL_CHARGED))
                If Not hdr.Exists(id) Then orphan(id) = NumOr0(orphan(id)) + 1
            End If
        Next i
    End If
    Dim kk As Variant
    For Each kk In orphan.Keys
        AddIssue SEV_WARN, "Postings", CStr(kk) & ": " & orphan(kk) & " list row(s) but no posting row", _
                 "Its header was deleted by hand or its creation stopped part-way. Nothing reuses the ID; send " & PRIMARY_ADMIN_NAME & " this line"
    Next kk
    If Not IsArray(pv) Then Exit Sub
    For i = 1 To LogRows()
        id = Trim$(SafeStr(mLog(i, L_REF)))
        If Left$(UCase$(id), 2) = "P-" And Not IsVoidRow(i) Then
            logSum(id) = NumOr0(logSum(id)) + NumOr0(mLog(i, L_CHARGED))
            logAny(id) = True
        End If
    Next i
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And IsDate(pv(i, PS_DATE)) Then
            ' this roster year only: last year's rows are in the archive
            If CDate(pv(i, PS_DATE)) >= mYearFrom And CDate(pv(i, PS_DATE)) <= mYearTo Then
                st = Trim$(SafeStr(pv(i, PS_STATUS)))
                If StrComp(st, "Open", vbTextCompare) = 0 Then
                    If logAny.Exists(id) Then
                        AddIssue SEV_WARN, "Postings", id & " is still Open but has rows in the OT Log", _
                                 "Load the posting and finish the award, or void its log rows"
                    End If
                ElseIf Abs(NumOr0(listSum(id)) - NumOr0(logSum(id))) > TOL Then
                    AddIssue SEV_WARN, "Postings", id & " (" & st & "): its list says " & NumOr0(listSum(id)) & " h charged, the OT Log has " & _
                             NumOr0(logSum(id)), "Expected if a charge was voided afterwards; otherwise reopen the posting and award it again"
                End If
            End If
        End If
    Next i
End Sub

' ---- 6. calendar lines that point at a posting or switch that isn't there ----
Private Sub CheckCalendarLinks()
    Dim cv As Variant: cv = TableValues(CV_TABLE)
    If Not IsArray(cv) Then Exit Sub
    Dim posts As Object: Set posts = CreateObject("Scripting.Dictionary")
    posts.CompareMode = 1
    Dim sw As Object: Set sw = CreateObject("Scripting.Dictionary")
    sw.CompareMode = 1
    Dim v As Variant, i As Long, src As String
    v = TableValues(PT_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Trim$(SafeStr(v(i, PS_ID))) <> "" Then posts(Trim$(SafeStr(v(i, PS_ID)))) = True
        Next i
    End If
    v = TableValues(SS_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Trim$(SafeStr(v(i, SW_ID))) <> "" Then sw(Trim$(SafeStr(v(i, SW_ID)))) = (UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID")
        Next i
    End If
    For i = 1 To UBound(cv, 1)
        If UCase$(SafeStr(cv(i, CV_STATUS))) <> "VOID" And IsDate(cv(i, CV_DATE)) Then
            If CDate(cv(i, CV_DATE)) >= mYearFrom Then
                src = UCase$(Trim$(SafeStr(cv(i, CV_SOURCE))))
                If Left$(src, 2) = "P-" Then
                    If Not posts.Exists(src) Then
                        AddIssue SEV_WARN, "Calendar", "Line " & SafeStr(cv(i, CV_ID)) & " on " & Format$(CDate(cv(i, CV_DATE)), "mmm d") & _
                                 " points at posting " & src & ", which doesn't exist", "Open the line on the calendar and clear or re-post it"
                    End If
                ElseIf Left$(src, 3) = "SS-" Then
                    If Not sw.Exists(src) Then
                        AddIssue SEV_WARN, "Calendar", "Line " & SafeStr(cv(i, CV_ID)) & " on " & Format$(CDate(cv(i, CV_DATE)), "mmm d") & _
                                 " points at switch " & src & ", which doesn't exist", "Open the line on the calendar and clear it"
                    ElseIf Not CBool(sw(src)) Then
                        AddIssue SEV_WARN, "Calendar", "Line " & SafeStr(cv(i, CV_ID)) & " on " & Format$(CDate(cv(i, CV_DATE)), "mmm d") & _
                                 " points at switch " & src & ", which is void", "Open the line on the calendar and clear it"
                    End If
                End If
            End If
        End If
    Next i
End Sub

' ---- 12. the calendar agrees with the log (v27, item 10, A5 b) ----
' Every non-void Accepted / Awarded row dated in the year needs a live coverage line whose Source is its ref and
' whose Covered By names the person: a PROBLEM for a shift from today on (the call order and the staffing count
' read the calendar), a Check for a past shift (review M3: history from before the auto-pick - a Cancel at the old
' question, a line voided with its absence - must not block the rollover or the drill). The other way, a live line
' with someone NAMED on it from an S- or P- source needs a live log row with that ref (Check; review M2: a POSTED
' line waits for its award, NCN and NO TAKERS are outcomes with no person, and a released line keeps its source
' with nobody on it).
Private Sub CheckCoverageVsLog()
    Dim cv As Variant: cv = TableValues(CV_TABLE)
    Dim i As Long, src As String, by As String, nm As String, ref As String, sd As Date
    ' the live lines by source: "|name|name|" of everyone named on them
    Dim lines As Object: Set lines = CreateObject("Scripting.Dictionary")
    lines.CompareMode = 1
    If IsArray(cv) Then
        For i = 1 To UBound(cv, 1)
            If UCase$(SafeStr(cv(i, CV_STATUS))) <> "VOID" Then
                src = UCase$(Trim$(SafeStr(cv(i, CV_SOURCE))))
                If src <> "" Then
                    by = SafeStr(cv(i, CV_BY))
                    If lines.Exists(src) Then lines(src) = lines(src) & by & "|" Else lines.Add src, "|" & by & "|"
                End If
            End If
        Next i
    End If
    ' the live Accepted / Awarded rows by ref
    Dim rows As Object: Set rows = CreateObject("Scripting.Dictionary")
    rows.CompareMode = 1
    Dim nBad As Long, sev As String, named As String
    If IsArray(mLog) Then
        For i = 1 To UBound(mLog, 1)
            If UCase$(Trim$(SafeStr(mLog(i, L_STATUS)))) <> "VOID" And IsDate(mLog(i, L_DATE)) Then
                ref = UCase$(Trim$(SafeStr(mLog(i, L_REF))))
                If ref <> "" Then rows(ref) = True
                sd = CDate(mLog(i, L_DATE))
                If sd >= mYearFrom And sd <= mYearTo Then
                    Select Case UCase$(Trim$(SafeStr(mLog(i, L_OUTCOME))))
                        Case "ACCEPTED", "AWARDED"
                            nm = Trim$(SafeStr(mLog(i, L_EMP)))
                            If nm <> "" And (Left$(ref, 2) = "S-" Or Left$(ref, 2) = "P-") Then
                                sev = IIf(Int(sd) >= Int(Date), SEV_ERR, SEV_WARN)
                                If Not lines.Exists(ref) Then
                                    nBad = nBad + 1
                                    AddIssue sev, "Coverage = log", ref & " " & Format$(sd, "mmm d") & " " & SafeStr(mLog(i, L_SHIFTLBL)) & ": " & nm & " is " & _
                                             LCase$(SafeStr(mLog(i, L_OUTCOME))) & " in the log but no live calendar line carries " & ref, _
                                             IIf(sev = SEV_ERR, "Calendar > that shift's line > Set coverage (or void the row if it is wrong)", "A past shift: the calendar's history is short by this line; nothing counts from it now")
                                ElseIf InStr(1, lines(ref), CalName(nm), vbTextCompare) = 0 And InStr(1, lines(ref), nm, vbTextCompare) = 0 Then
                                    nBad = nBad + 1
                                    named = Trim$(Replace(Mid$(lines(ref), 2, Len(lines(ref)) - 2), "|", " / "))
                                    If Replace(named, "/", "") = "" Or Trim$(named) = "" Then named = "nobody"
                                    AddIssue sev, "Coverage = log", ref & " " & Format$(sd, "mmm d") & " " & SafeStr(mLog(i, L_SHIFTLBL)) & ": " & nm & " is " & _
                                             LCase$(SafeStr(mLog(i, L_OUTCOME))) & " in the log but the calendar line for " & ref & " names " & named, _
                                             IIf(sev = SEV_ERR, "Calendar > that line > Set coverage", "A past shift: nothing counts from it now")
                                End If
                            End If
                    End Select
                End If
            End If
        Next i
    End If
    ' the other way: a line that says it came from a session or an award nobody logged
    If IsArray(cv) Then
        For i = 1 To UBound(cv, 1)
            If UCase$(SafeStr(cv(i, CV_STATUS))) <> "VOID" And IsDate(cv(i, CV_DATE)) Then
                If CDate(cv(i, CV_DATE)) >= mYearFrom Then
                    src = UCase$(Trim$(SafeStr(cv(i, CV_SOURCE))))
                    by = Trim$(SafeStr(cv(i, CV_BY)))
                    ' (review M2) only a line with a PERSON named on it: a POSTED line waits for its award, NCN / NO TAKERS are
                    ' outcomes without a person, and a voided session's released line keeps its source with nobody on it
                    If (Left$(src, 2) = "S-" Or Left$(src, 2) = "P-") And by <> "" And Not (UCase$(by) Like "POSTED*") _
                       And Not (UCase$(by) Like "NCN*") And Not (UCase$(by) Like "NO TAKERS*") Then
                        If Not rows.Exists(src) Then
                            AddIssue SEV_WARN, "Coverage = log", "Line " & SafeStr(cv(i, CV_ID)) & " on " & Format$(CDate(cv(i, CV_DATE)), "mmm d") & " (" & by & ") came from " & src & _
                                     ", which has no live OT Log row", "Open the line on the calendar: clear it, or log the session again"
                        End If
                    End If
                End If
            End If
        Next i
    End If
    If nBad = 0 Then AddIssue SEV_INFO, "Coverage = log", "Every accepted call-out and award this year has its calendar line", ""
    ' (review) the variable sev is reused per row above; nothing else to restore
End Sub

' ---- 13. the Settings sheet (v27, item 22) ----
Private Sub CheckSettings()
    Dim bad As String, txt As String
    txt = SettingsOverrideText(bad)
    If bad <> "" Then AddIssue SEV_WARN, "Settings", bad, "Alt+F8 > SetSetting writes a row; a row that is not a setting is ignored - remove it through UnlockForRepair"
    If txt <> "Settings overrides: none" Then AddIssue SEV_INFO, "Settings", txt, ""
End Sub

' ---- 7. paybacks owed for a long time ----
Private Sub CheckPaybacks()
    Dim txt As String, n As Long
    n = OwedOlderThan(PaybackWarnDays(), txt)
    If n > 0 Then
        AddIssue SEV_WARN, "Paybacks", n & " shift-switch payback(s) owed for more than " & PaybackWarnDays() & " days: " & txt, _
                 "See SHIFT SWITCH PAYBACKS OWED on Needs Attention"
    End If
End Sub

' ---- 11. planned unit moves (v24.3) ----
' A waiting move for someone who isn't on the roster, to a unit that doesn't
' exist, with no readable date, or more than 3 days past its date.
Private Sub CheckPlannedMoves()
    Dim pv As Variant: pv = PlannedRows()
    Dim i As Long, id As String, emp As String, n As Long
    Dim perEmp As Object: Set perEmp = CreateObject("Scripting.Dictionary")
    perEmp.CompareMode = 1
    If IsArray(pv) Then                               ' (v25 re-review: the cancelled-move check below runs either way)
    For i = 1 To UBound(pv, 1)
        n = n + 1
        id = Trim$(SafeStr(pv(i, PM_ID)))
        emp = Trim$(SafeStr(pv(i, PM_EMP)))
        If Not mOnRoster.Exists(emp) Then
            AddIssue SEV_WARN, "Planned moves", id & ": " & emp & " isn't on the roster", "Cancel it: Needs Attention > Planned unit moves"
        End If
        If DeptCol(Trim$(SafeStr(pv(i, PM_UNIT)))) = 0 Then
            AddIssue SEV_WARN, "Planned moves", id & ": '" & SafeStr(pv(i, PM_UNIT)) & "' isn't a unit", "Cancel it and enter it again"
        End If
        If Not IsDate(pv(i, PM_FROM)) Then
            AddIssue SEV_WARN, "Planned moves", id & ": the Effective From date can't be read", "Cancel it and enter it again"
        ElseIf Int(CDate(pv(i, PM_FROM))) < Date - 3 Then
            AddIssue SEV_WARN, "Planned moves", id & ": " & emp & "'s move was due " & Format$(CDate(pv(i, PM_FROM)), "mmm d, yyyy") & _
                     " and hasn't been applied", "Alt+F8 > ApplyPlannedMoves (the Note on the Planned Moves sheet says why)"
        End If
    Next i
    ' v25 (H9): two waiting moves for one person (a hand edit: the code reads them two ways)
    For i = 1 To UBound(pv, 1)
        emp = Trim$(SafeStr(pv(i, PM_EMP)))
        If emp <> "" Then perEmp(emp) = NumOr0(perEmp(emp)) + 1
    Next i
    Dim k As Variant
    For Each k In perEmp.Keys
        If perEmp(k) > 1 Then AddIssue SEV_WARN, "Planned moves", perEmp(k) & " waiting moves for " & k, "Cancel all but one: Needs Attention > Planned unit moves"
    Next k
    End If
    ' v25 (H9): a future crew row that a CANCELLED planned move left behind (a cancel before v25 kept the
    ' crewTo row). A future-dated crew change typed in column E is allowed, so only rows that match a
    ' cancelled move's crew and date are flagged (v25 review).
    Dim canc As Object: Set canc = CreateObject("Scripting.Dictionary")
    canc.CompareMode = 1
    Dim live As Object: Set live = CreateObject("Scripting.Dictionary")      ' the same key on a Planned / Applied move: the row is wanted
    live.CompareMode = 1
    Dim pt As ListObject: Set pt = PlannedTable(False)
    Dim key As String
    If Not pt Is Nothing Then
        If Not pt.DataBodyRange Is Nothing Then
            Dim av As Variant: av = RangeTo2D(pt.DataBodyRange)
            If UBound(av, 2) >= PM_STATUS Then
                For i = 1 To UBound(av, 1)
                    If IsDate(av(i, PM_FROM)) Then
                        key = Trim$(SafeStr(av(i, PM_EMP))) & "|" & UCase$(Trim$(SafeStr(av(i, PM_CREW)))) & "|" & Format$(CDate(av(i, PM_FROM)), "yyyy-mm-dd")
                        If StrComp(Trim$(SafeStr(av(i, PM_STATUS))), "Cancelled", vbTextCompare) = 0 Then canc(key) = True Else live(key) = True
                    End If
                Next i
            End If
        End If
    End If
    Dim lo As ListObject: Set lo = CrewTable()
    If Not lo Is Nothing And canc.Count > 0 Then
        If Not lo.DataBodyRange Is Nothing Then
            Dim cv As Variant: cv = lo.DataBodyRange.Value
            For i = 1 To UBound(cv, 1)
                If IsDate(cv(i, CH_FROM)) Then
                    If Int(CDate(cv(i, CH_FROM))) > Date Then
                        key = Trim$(SafeStr(cv(i, CH_EMP))) & "|" & UCase$(Trim$(SafeStr(cv(i, CH_CREW)))) & "|" & Format$(CDate(cv(i, CH_FROM)), "yyyy-mm-dd")
                        If canc.Exists(key) And Not live.Exists(key) Then
                            AddIssue SEV_WARN, "Planned moves", "Crew History: " & Trim$(SafeStr(cv(i, CH_EMP))) & " has a crew row dated " & _
                                     Format$(CDate(cv(i, CH_FROM)), "mmm d, yyyy") & " left by a cancelled planned move", _
                                     "Remove that row on the Crew History sheet (the move it came with was cancelled)" & ByHandHint(CH_SHEET)
                        End If
                    End If
                End If
            Next i
        End If
    End If
    If n > 0 Then AddIssue SEV_INFO, "Planned moves", n & " unit move(s) waiting for their date", "Listed on Needs Attention"
End Sub

' ---- 8. Kronos ----
Private Sub CheckKronos()
    If KronosIsStale() Then AddIssue SEV_WARN, "Kronos", KronosStatusText(), "Run Import Kronos"
    ' v27 (briefs/v27-kronos.md): Worked (H) must be what the Kronos Totals snapshot + the Worked Days ledger give -
    ' once a year-to-date report has been taken this year (review N3: one-day reports alone, or a January before
    ' the first report, give H nothing to be compared with)
    If KronosCumThrough() < RosterYearFrom() Or KronosCumThrough() = 0 Then Exit Sub
    Dim through As Date, piv As Variant
    piv = PivotFromLedger(through)
    If UBound(piv, 1) < 2 Then Exit Sub                     ' no report imported yet: the pivot sheet fed H, nothing to compare
    Dim expectH As Object: Set expectH = CreateObject("Scripting.Dictionary")
    expectH.CompareMode = 1
    Dim i As Long
    For i = 2 To UBound(piv, 1)
        expectH(Trim$(SafeStr(piv(i, 2)))) = NumOr0(piv(i, 3))
    Next i
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long, nm As String, h As Double, nBad As Long
    For r = R_FIRST To RosterLastRow()
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            If expectH.Exists(nm) Then
                h = NumOr0(ws.Cells(r, C_WORKED).Value)
                If Abs(h - CDbl(expectH(nm))) > 0.001 Then
                    nBad = nBad + 1
                    If nBad <= 10 Then AddIssue SEV_ERR, "Kronos", nm & ": Worked (H) is " & HoursText(h) & " but the Kronos Totals + ledger give " & HoursText(CDbl(expectH(nm))), _
                                                "Run Import Kronos: it rewrites H from the tables (the sanity confirm names who moves)"
                End If
            End If
        End If
    Next r
    If nBad > 10 Then AddIssue SEV_ERR, "Kronos", "... and " & (nBad - 10) & " more people whose Worked (H) differs from the Kronos tables", "Run Import Kronos"
    ' the ledger: every day inside the roster year; the snapshot days since the ledger began with no rows at all
    Dim v As Variant: v = TableValues(WD_TABLE)
    Dim days As Object: Set days = CreateObject("Scripting.Dictionary")
    Dim notCounted As Object: Set notCounted = CreateObject("Scripting.Dictionary")
    notCounted.CompareMode = 1
    Dim d As Date, nOut As Long
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If IsDate(v(i, 3)) Then
                d = CDate(v(i, 3))
                days(CLng(Int(CDbl(d)))) = True
                If d < RosterYearFrom() Or d > RosterYearTo() Then
                    nOut = nOut + 1
                    If nOut <= 5 Then AddIssue SEV_ERR, "Kronos", "Worked Days holds a row dated " & Format$(d, "mmm d, yyyy") & " (" & Trim$(SafeStr(v(i, 2))) & "), outside this roster's year", _
                                               "A report for another year was imported: an admin removes the row (UnlockForRepair) and re-imports"
                End If
                If PremiumMultiplier(Trim$(SafeStr(v(i, 4))), PREMIUM_CODES, PREMIUM_MULTIPLIERS) < 0 And NumOr0(v(i, 5)) <> 0 Then _
                    notCounted(Trim$(SafeStr(v(i, 2))) & ": " & Trim$(SafeStr(v(i, 4)))) = True
            End If
        Next i
    End If
    v = TableValues(KT_TABLE)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If PremiumMultiplier(Trim$(SafeStr(v(i, 3))), PREMIUM_CODES, PREMIUM_MULTIPLIERS) < 0 And NumOr0(v(i, 4)) <> 0 Then _
                notCounted(Trim$(SafeStr(v(i, 2))) & ": " & Trim$(SafeStr(v(i, 3)))) = True
        Next i
    End If
    Dim kk As Variant, nShown As Long
    For Each kk In notCounted.Keys
        nShown = nShown + 1
        If nShown <= 10 Then AddIssue SEV_WARN, "Kronos", "A roster person has hours under a code the roster does not count - " & kk, _
                                     "DBLTIME, Uanpproved OT and LIEU TKN are never an operator's: check the person's codes with payroll, or the PREMIUM_CODES setting"
    Next kk
    If notCounted.Count > 10 Then AddIssue SEV_WARN, "Kronos", "... and " & (notCounted.Count - 10) & " more codes not counted", ""
    Dim lf As Date: lf = KronosLedgerFrom()
    Dim cum As Date: cum = KronosCumThrough()
    If lf > 0 Then
        Dim first As Date: first = lf
        If cum + 1 > first Then first = cum + 1
        Dim missing As String, nMiss As Long
        For d = first To Date - 2                            ' (review) yesterday's report arrives this morning: not missing yet
            If Not days.Exists(CLng(Int(CDbl(d)))) Then
                nMiss = nMiss + 1
                If nMiss <= 10 Then missing = missing & IIf(missing = "", "", ", ") & Format$(d, "ddd mmm d")
            End If
        Next d
        If nMiss > 0 Then AddIssue SEV_WARN, "Kronos", "No one-day Kronos report for " & nMiss & " day(s) since the ledger began: " & missing & IIf(nMiss > 10, " ...", ""), _
                                        "Run the report for each missing day into the Kronos folder, then Import Kronos (or the year to date at the period's end)"
    End If
End Sub


' ====================================================================
'  WHAT THE USER SEES
' ====================================================================

' Roster status line text ("" until a run this session)
Public Function IntegrityStatusText() As String
    If mRanAt = 0 Then Exit Function
    If mLastErr > 0 Then
        IntegrityStatusText = "Integrity: " & mLastErr & " PROBLEM" & IIf(mLastErr = 1, "", "S")
    ElseIf mLastWarn > 0 Then
        IntegrityStatusText = "Integrity: OK (" & mLastWarn & " to check)"
    Else
        IntegrityStatusText = "Integrity: OK"
    End If
End Function

' v27 (item 12): the restore drill runs the check in the backup copy through Application.Run and reads the PROBLEM count
Public Function IntegrityRunCount() As Long
    Dim nErr As Long, nWarn As Long
    RunIntegrity nErr, nWarn
    IntegrityRunCount = nErr
End Function

Public Function IntegrityHasProblems() As Boolean
    IntegrityHasProblems = (mRanAt > 0 And mLastErr > 0)
End Function

' Quiet run when an admin opens the file. Never stops the open.
Public Sub IntegrityOnOpen()
    On Error Resume Next
    If Not IsAdmin() Then Exit Sub
    Dim nErr As Long, nWarn As Long
    RunIntegrity nErr, nWarn
    SetRosterStatus
End Sub

' Alt+F8 > IntegrityCheck: the full report
Public Sub IntegrityCheck()
    On Error GoTo Oops                         ' v24.3.1
    If Not RequireAdmin() Then Exit Sub
    Dim nErr As Long, nWarn As Long, n As Long
    Application.Calculate
    n = RunIntegrity(nErr, nWarn)
    DrawIntegritySheet
    On Error Resume Next
    SetRosterStatus
    On Error GoTo Oops
    If nErr = 0 And nWarn = 0 Then
        MsgBox "Integrity check: everything adds up." & IIf(n > 0, vbNewLine & n & " note(s) on the Integrity sheet.", ""), vbInformation, "Integrity"
    ElseIf nErr = 0 Then
        MsgBox "Integrity check: the hours add up. " & nWarn & " thing(s) to look at are listed on the Integrity sheet.", vbInformation, "Integrity"
    Else
        MsgBox "Integrity check found " & nErr & " PROBLEM(S). They are listed on the Integrity sheet with what to do." & vbNewLine & vbNewLine & _
               "Roster totals may be wrong until they are fixed.", vbExclamation, "Integrity"
    End If
    Exit Sub
Oops:
    ReportError "IntegrityCheck"
End Sub

' Writes the last run to the Integrity sheet (admin, hidden in saved copies). Used by IntegrityCheck and RolloverYear.
' The optional argument keeps it out of the Alt+F8 list; with no run this session there is nothing to draw.
Public Sub DrawIntegritySheet(Optional ByVal fromCode As Boolean = True)
    If mRanAt = 0 Then Exit Sub
    Dim ws As Worksheet
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Failed
    If SheetExists(IC_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(IC_SHEET)
    Else
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = IC_SHEET
        RelockStructure
        SetTabColor ws, RGB(176, 0, 32)                 ' v25.1
    End If
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Range("A1").Value = "INTEGRITY CHECK  -  " & Format$(mRanAt, "ddd mmm d, yyyy h:mm AM/PM")
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 12
    ws.Range("A2").Value = mErr & " PROBLEM(S), " & mWarn & " to check. PROBLEM = totals can't be trusted until fixed. Check = worth a look. " & _
                           "Nothing was changed by this check; run it again after fixing."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    ws.Range("A4:D4").Value = Array("Level", "Check", "What was found", "What to do")
    With ws.Range("A4:D4")
        .Font.Bold = True: .Interior.Color = RGB(176, 0, 32): .Font.Color = vbWhite
    End With
    Dim i As Long, j As Long
    If mN > 0 Then
        Dim out() As Variant: ReDim out(1 To mN, 1 To 4)
        For i = 1 To mN
            For j = 1 To 4
                out(i, j) = mIss(j, i)
            Next j
        Next i
        ws.Range("A5").Resize(mN, 4).NumberFormat = "General"      ' not "@": a text cell over 255 characters shows ####
        ws.Range("A5").Resize(mN, 4).Value = out
        ws.Range("A5").Resize(mN, 4).WrapText = True
        ws.Range("A5").Resize(mN, 4).VerticalAlignment = xlTop
        For i = 1 To mN
            If mIss(1, i) = SEV_ERR Then
                ws.Cells(4 + i, 1).Interior.Color = RGB(255, 205, 210)
                ws.Cells(4 + i, 1).Font.Bold = True
            ElseIf mIss(1, i) = SEV_WARN Then
                ws.Cells(4 + i, 1).Interior.Color = RGB(255, 236, 179)
            End If
        Next i
    Else
        ws.Range("A5").Value = "Everything adds up."
    End If
    ws.Columns(1).ColumnWidth = 11
    ws.Columns(2).ColumnWidth = 15
    ws.Columns(3).ColumnWidth = 78
    ws.Columns(4).ColumnWidth = 62
    ProtectSheet ws
    ws.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    LogError "DrawIntegritySheet", Err.Number, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    MsgBox "The Integrity sheet couldn't be drawn: " & em, vbExclamation, "Call Out Roster"
End Sub
