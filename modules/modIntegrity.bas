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
'             missing formula in a pay-period cell, a duplicate ID.
'    Check    worth a look, nothing is wrong with the hours: people
'             no longer on the roster, a posting whose list and log
'             disagree, a calendar line pointing at nothing, an old
'             payback, old Kronos data.
'    Note     information only.
'
'  The pure parts (PeriodCellCode, DuplicatesIn) take plain values so
'  SelfTest can hand them a planted bad cell / duplicate ID in memory.
' ====================================================================

Public Const IC_VERSION As String = "Integrity v23 (Oct 20)"
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
        For k = 1 To 8
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
                    AddIssue SEV_WARN, "Roster = log", mNames(i) & " is on the roster twice (row " & (R_FIRST + i - 1) & ")", _
                             "Both rows show the same hours; remove or rename one"
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
    Dim r As Long, expSum As Double, expCell As Double, act As Variant, bad As String, nBad As Long
    Dim vI As Variant, vG As Variant, vH As Variant, vC As Variant
    For r = 1 To mRows
        If mNames(r) <> "" Then
            expSum = 0: bad = "": nBad = 0
            For c = C_PERIOD1 To mLastPer
                expCell = NumOr0(sums(mNames(r) & "|" & per(c)))
                expSum = expSum + expCell
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
        End If
    Next r
End Sub

' ---- 3. IDs ----
Private Sub CheckIds()
    Dim keys() As String, n As Long, i As Long, s As String
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
                     " h charged) but isn't on the roster", "If they left, put the roster row back and use Employee > Remove. If the name was typed over by hand, type the old name back and use Employee > Rename"
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
                        AddIssue SEV_WARN, "Not on roster", "Booking " & SafeStr(v(i, AB_ID)) & " is for " & emp & ", who isn't on the roster", _
                                 "Void it in Book Time Off, or fix the name"
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
End Sub

Private Sub OrphanSwitchName(ByVal id As String, ByVal emp As String)
    If emp = "" Then Exit Sub
    If mOnRoster.Exists(emp) Then Exit Sub
    AddIssue SEV_WARN, "Not on roster", "Switch " & id & " names " & emp & ", who isn't on the roster", "Void the switch or fix the name"
End Sub

' ---- 5. postings: the posting's own list against the log ----
Private Sub CheckPostings()
    Dim pv As Variant: pv = TableValues(PT_TABLE)
    If Not IsArray(pv) Then Exit Sub
    Dim lv As Variant: lv = TableValues(PL_TABLE)
    Dim listSum As Object: Set listSum = CreateObject("Scripting.Dictionary")
    listSum.CompareMode = 1
    Dim logSum As Object: Set logSum = CreateObject("Scripting.Dictionary")
    logSum.CompareMode = 1
    Dim logAny As Object: Set logAny = CreateObject("Scripting.Dictionary")
    logAny.CompareMode = 1
    Dim i As Long, id As String, st As String
    If IsArray(lv) Then
        For i = 1 To UBound(lv, 1)
            id = Trim$(SafeStr(lv(i, PL_ID)))
            If id <> "" Then listSum(id) = NumOr0(listSum(id)) + NumOr0(lv(i, PL_CHARGED))
        Next i
    End If
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

' ---- 7. paybacks owed for a long time ----
Private Sub CheckPaybacks()
    Dim txt As String, n As Long
    n = OwedOlderThan(PAYBACK_WARN_DAYS, txt)
    If n > 0 Then
        AddIssue SEV_WARN, "Paybacks", n & " shift-switch payback(s) owed for more than " & PAYBACK_WARN_DAYS & " days: " & txt, _
                 "See SHIFT SWITCH PAYBACKS OWED on Needs Attention"
    End If
End Sub

' ---- 8. Kronos ----
Private Sub CheckKronos()
    If KronosIsStale() Then AddIssue SEV_WARN, "Kronos", KronosStatusText(), "Run Import Kronos"
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
    If Not RequireAdmin() Then Exit Sub
    Dim nErr As Long, nWarn As Long, n As Long
    Application.Calculate
    n = RunIntegrity(nErr, nWarn)
    DrawIntegritySheet
    On Error Resume Next
    SetRosterStatus
    On Error GoTo 0
    If nErr = 0 And nWarn = 0 Then
        MsgBox "Integrity check: everything adds up." & IIf(n > 0, vbNewLine & n & " note(s) on the Integrity sheet.", ""), vbInformation, "Integrity"
    ElseIf nErr = 0 Then
        MsgBox "Integrity check: the hours add up. " & nWarn & " thing(s) to look at are listed on the Integrity sheet.", vbInformation, "Integrity"
    Else
        MsgBox "Integrity check found " & nErr & " PROBLEM(S). They are listed on the Integrity sheet with what to do." & vbNewLine & vbNewLine & _
               "Roster totals may be wrong until they are fixed.", vbExclamation, "Integrity"
    End If
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
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = IC_SHEET
        ws.Tab.Color = RGB(176, 0, 32)
    End If
    ws.Visible = xlSheetVisible
    ws.Unprotect Password:=SHEET_PASSWORD
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
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
    MsgBox "The Integrity sheet couldn't be drawn: " & em, vbExclamation
End Sub
