Attribute VB_Name = "modRollover"
Option Explicit

' ====================================================================
'  YEAR-END ROLLOVER (v23)
'
'  Alt+F8 > RolloverYear, run once a year by the admin with a second
'  person watching (docs/rollover-one-pager.md). In order:
'
'    1. IntegrityCheck must show no PROBLEM (the run stops if it does).
'    2. Each person's final Total / Worked / Refused / Provisional is
'       written to the Year End sheet.
'    3. A frozen copy "Call Out Roster 2026 FINAL.xlsm" is saved in
'       ARCHIVE_FOLDER, reopened read-only and compared with step 2,
'       then marked read-only.
'    4. Last year's OT Log rows (Period Start before the new year's
'       first period) are copied to a sheet "OT Log 2026" and removed
'       from the live log. This is the ONE sanctioned removal: row
'       counts and Hours Charged are compared before the save. ID
'       numbers carry on (high-water marks), they never restart.
'    5. The roster's first period date (T2) is set to the new year.
'       Every period header and formula follows from that one cell.
'       A 27-period year gets column AT; a 26-period year drops it.
'    6. Worked (H) is set to 0 and stays 0 until a Kronos export with
'       data through Jan 1 or later is imported.
'    7. Everything is checked. Any mismatch blocks the save: close
'       WITHOUT saving and the file on disk is exactly as it was.
'
'  Nothing is saved until step 7 passes. The date guard refuses a run
'  before the new year's first pay period starts, unless the file name
'  contains COPY or DRY RUN (the December practice on a copy).
' ====================================================================

Public Const RO_VERSION As String = "Rollover v24.1 (Jan 5)"
Public Const YE_SHEET As String = "Year End"
Public Const YE_TABLE As String = "tblYearEnd"

Private Const YE_NCOLS As Long = 9      ' Year, Employee, Home Unit, Total Hours, Worked, Refused, Provisional, Rolled On, Rolled By
Private Const NOTE_DAYS As Long = 30    ' how long "Rolled over to ..." shows on the status line / Needs Attention
Private Const TOL As Double = 0.001


' ====================================================================
'  PURE HELPERS (SelfTest "archive" and "dryname")
' ====================================================================

' True when a log row belongs to the year being closed: its Period Start
' is a real date before the new year's first period. Blank / text stays.
Public Function GoesToArchive(ByVal periodStart As Variant, ByVal cutDate As Date) As Boolean
    If IsError(periodStart) Then Exit Function
    If IsEmpty(periodStart) Then Exit Function
    If VarType(periodStart) = vbString Then Exit Function     ' text is never a period date (the roster formulas don't count it either)
    Dim x As Double
    If IsDate(periodStart) Then
        x = Int(CDbl(CDate(periodStart)))
    ElseIf IsNumeric(periodStart) Then
        x = Int(CDbl(periodStart))
    Else
        Exit Function
    End If
    If x <= 0 Then Exit Function
    GoesToArchive = (x < Int(CDbl(cutDate)))
End Function

' A practice copy: the file name says so
Public Function IsDryRunName(ByVal fileName As String) As Boolean
    Dim u As String: u = UCase$(fileName)
    IsDryRunName = (InStr(u, "COPY") > 0 Or InStr(u, "DRY RUN") > 0 Or InStr(u, "DRYRUN") > 0 Or InStr(u, "DRY-RUN") > 0)
End Function

Public Function FrozenFileName(ByVal y As Long, ByVal dryRun As Boolean) As String
    FrozenFileName = BACKUP_BASENAME & " " & y & " FINAL" & IIf(dryRun, " (dry run)", "") & ".xlsm"
End Function

Public Function ArchiveSheetName(ByVal y As Long) As String
    ArchiveSheetName = LOG_SHEET & " " & y
End Function


' ====================================================================
'  THE ROLLOVER
' ====================================================================
Public Sub RolloverYear()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If NoSave() Then
        MsgBox "Saving is blocked in this session because an earlier check failed. Close WITHOUT saving, reopen, and start again.", vbExclamation, "Rollover"
        Exit Sub
    End If
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then MsgBox "The OT Log table wasn't found. Run SystemCheck.", vbExclamation, "Rollover": Exit Sub
    If Not LogLayoutCurrent(lo) Then MsgBox "The OT Log is in the old layout. Run UpgradeOTLogLayout first.", vbExclamation, "Rollover": Exit Sub
    If PStartRow() <> 2 Then MsgBox "The roster's date rows are in the old layout. Run ApplyRosterLayout first.", vbExclamation, "Rollover": Exit Sub
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then MsgBox "No employees found on the roster.", vbExclamation, "Rollover": Exit Sub

    ' ---- which year is closing, which one opens ----
    Dim oldFrom As Date, oldYear As Long, newYear As Long, newFrom As Date, lastStart As Date
    Dim havePer As Long, wantPer As Long, gridBad As String
    On Error GoTo BadDates
    oldFrom = RosterYearFrom()
    oldYear = RosterYear()
    newYear = oldYear + 1
    havePer = PeriodLastCol() - C_PERIOD1 + 1
    lastStart = Int(CDate(ws.Cells(PStartRow(), PeriodLastCol()).Value))
    newFrom = FirstPeriodStart(oldFrom, newYear)
    wantPer = PeriodsInYear(oldFrom, newYear)
    gridBad = PeriodGridProblem(ws, oldFrom, havePer)
    On Error GoTo 0
    If gridBad <> "" Then
        MsgBox "The pay-period dates above the roster aren't a clean two-week run: " & gridBad & "." & vbNewLine & vbNewLine & _
               "Someone may have typed over a date in row 2 or 4. Nothing was changed. Send this message to the admin.", vbCritical, "Rollover"
        Exit Sub
    End If
    If newFrom <> lastStart + 14 Then
        MsgBox "The dates on the roster don't make a whole " & oldYear & " year:" & vbNewLine & _
               "the last period starts " & Format$(lastStart, "mmm d, yyyy") & " but " & newYear & " should start " & _
               Format$(newFrom, "mmm d, yyyy") & "." & vbNewLine & vbNewLine & "Nothing was changed. Send this message to the admin.", vbCritical, "Rollover"
        Exit Sub
    End If
    If wantPer <> 26 And wantPer <> 27 Then
        MsgBox newYear & " works out to " & wantPer & " pay periods, which the roster can't hold. Nothing was changed.", vbCritical, "Rollover"
        Exit Sub
    End If

    If lo.ShowTotals Then
        MsgBox "The OT Log table has its Total row switched on. Turn it off (Table Design > Total Row), then run this again. Nothing was changed.", vbExclamation, "Rollover"
        Exit Sub
    End If
    If wantPer = 27 And havePer = 26 Then
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(1, C_PERIODX), ws.Cells(lastRow + 200, C_PERIODX))) > 0 Then
            MsgBox newYear & " has 27 pay periods and needs column " & ColLetter(C_PERIODX) & ", which isn't empty. Clear that column first. Nothing was changed.", vbExclamation, "Rollover"
            Exit Sub
        End If
    End If

    ' ---- date guard: the live file can't be rolled early ----
    Dim dry As Boolean: dry = IsDryRunName(ThisWorkbook.Name)
    If Date < newFrom And Not dry Then
        MsgBox "It's too early to roll over. " & newYear & "'s first pay period starts " & Format$(newFrom, "ddd mmm d, yyyy") & "." & vbNewLine & vbNewLine & _
               "To practise, save a copy of this file with COPY or DRY RUN in its name and run RolloverYear in the copy.", vbExclamation, "Rollover"
        Exit Sub
    End If

    ' ---- where things go ----
    Dim arcName As String: arcName = ArchiveSheetName(oldYear)
    If SheetExists(arcName) Then
        MsgBox "There is already a sheet called '" & arcName & "'. " & oldYear & " looks rolled over already. Nothing was changed.", vbExclamation, "Rollover"
        Exit Sub
    End If
    Dim folder As String: folder = ResolveFolder(ARCHIVE_FOLDER)
    If folder = "" Or Not FolderExists(folder) Then
        MsgBox "The archive folder can't be reached:" & vbNewLine & folder & vbNewLine & vbNewLine & _
               "Check ARCHIVE_FOLDER in modOvertimeEntry. Nothing was changed.", vbExclamation, "Rollover"
        Exit Sub
    End If
    Dim frozen As String: frozen = folder & FrozenFileName(oldYear, dry)
    If Dir$(frozen) <> "" Then
        MsgBox "The frozen copy already exists:" & vbNewLine & frozen & vbNewLine & vbNewLine & _
               "Move or rename it first (it is never overwritten). Nothing was changed.", vbExclamation, "Rollover"
        Exit Sub
    End If

    ' ---- 1. integrity ----
    Dim nErr As Long, nWarn As Long
    Application.Calculate
    RunIntegrity nErr, nWarn
    If nErr > 0 Then
        DrawIntegritySheet
        MsgBox "The rollover did NOT start: the integrity check found " & nErr & " PROBLEM(S)." & vbNewLine & vbNewLine & _
               "They are listed on the Integrity sheet with what to do. Fix them, then run RolloverYear again. Nothing was changed.", vbCritical, "Rollover"
        Exit Sub
    End If

    ' ---- what will move (read only) ----
    Dim v As Variant, nAll As Long, nArc As Long, nKeep As Long, i As Long
    Dim sumAll As Double, sumArc As Double, sumKeep As Double
    If Not lo.DataBodyRange Is Nothing Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then v = RangeTo2D(lo.DataBodyRange)
    End If
    If IsArray(v) Then
        nAll = UBound(v, 1)
        For i = 1 To nAll
            sumAll = sumAll + NumOr0(v(i, L_CHARGED))
            If GoesToArchive(v(i, L_PERIOD), newFrom) Then
                nArc = nArc + 1: sumArc = sumArc + NumOr0(v(i, L_CHARGED))
            Else
                nKeep = nKeep + 1: sumKeep = sumKeep + NumOr0(v(i, L_CHARGED))
            End If
        Next i
    End If
    Dim nPeople As Long, r As Long
    For r = R_FIRST To lastRow
        If Trim$(SafeStr(ws.Cells(r, C_NAME).Value)) <> "" Then nPeople = nPeople + 1
    Next r

    Dim msg As String
    msg = IIf(dry, "PRACTICE RUN on a copy." & vbNewLine & vbNewLine, "") & _
          "Roll the roster over from " & oldYear & " to " & newYear & "?" & vbNewLine & vbNewLine & _
          "- " & newYear & " starts " & Format$(newFrom, "ddd mmm d, yyyy") & " (" & wantPer & " pay periods)" & vbNewLine & _
          "- frozen copy: " & FrozenFileName(oldYear, dry) & vbNewLine & "   in " & folder & vbNewLine & _
          "- " & nPeople & " people: final totals recorded on the Year End sheet" & vbNewLine & _
          "- OT Log: " & Format$(nArc, "#,##0") & " row(s) move to '" & arcName & "', " & Format$(nKeep, "#,##0") & " stay (already logged for " & newYear & ")" & vbNewLine & _
          "- Worked hours (H) go to 0 until the first " & newYear & " Kronos export is imported" & vbNewLine & vbNewLine & _
          IIf(nWarn > 0, "The integrity check has " & nWarn & " thing(s) to look at (not problems). ", "The integrity check is clean. ") & _
          "Nothing is saved unless every check passes."
    If MsgBox(msg, vbOKCancel + vbQuestion + vbDefaultButton2, "Rollover " & oldYear & " -> " & newYear) <> vbOK Then Exit Sub
    Dim typed As String
    typed = InputBox("Type  ROLL " & newYear & "  to go ahead.", "Rollover " & oldYear & " -> " & newYear)
    If UCase$(Trim$(typed)) <> "ROLL " & newYear Then
        MsgBox "Not confirmed. Nothing was changed.", vbInformation, "Rollover"
        Exit Sub
    End If

    ' ====== from here on, any failure blocks the save ======
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim calcWas As XlCalculation: calcWas = Application.Calculation
    Dim changed As Boolean, madeFrozen As Boolean, stepTxt As String, problems As String
    Dim nm() As String, fG() As Double, fH() As Double, fI() As Double, fC() As Double, nRows As Long
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler      ' Esc / Ctrl+Break lands in Failed, never in a half-rolled file

    ' ---- 2. finals ----
    stepTxt = "recording the final totals"
    Application.Calculate
    nRows = lastRow - R_FIRST + 1
    ReDim nm(1 To nRows): ReDim fG(1 To nRows): ReDim fH(1 To nRows): ReDim fI(1 To nRows): ReDim fC(1 To nRows)
    Dim blk As Variant: blk = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, C_REFUSED)).Value2
    For r = 1 To nRows
        nm(r) = Trim$(SafeStr(blk(r, C_NAME)))
        If nm(r) <> "" Then
            If IsError(blk(r, C_TOTAL)) Or IsError(blk(r, C_WORKED)) Or IsError(blk(r, C_REFUSED)) Or IsError(blk(r, C_PROV)) Then
                Err.Raise vbObjectError + 710, , nm(r) & "'s totals show an error"
            End If
            fG(r) = NumOr0(blk(r, C_TOTAL)): fH(r) = NumOr0(blk(r, C_WORKED))
            fI(r) = NumOr0(blk(r, C_REFUSED)): fC(r) = NumOr0(blk(r, C_PROV))
        End If
    Next r
    changed = True
    AppendYearEnd oldYear, blk, nm, fG, fH, fI, fC, nRows
    AppendYearEndFormer oldYear             ' v24: people moved to the Former sheet during the year

    ' ---- 3. frozen copy (saved in the restricted view, like every save) ----
    stepTxt = "saving the frozen copy"
    ws.Activate
    ApplyView False
    ThisWorkbook.SaveCopyAs frozen
    madeFrozen = True
    ApplyView True
    Application.EnableEvents = False
    stepTxt = "checking the frozen copy"
    problems = FrozenCopyProblem(frozen, ws.Name, nm, fG, fH, fI, nRows, lastRow)
    If problems <> "" Then Err.Raise vbObjectError + 711, , "the frozen copy doesn't match: " & problems
    SetAttr frozen, vbReadOnly
    MarkNoSave                                        ' from here until step 7 passes this file must not be saved

    ' ---- 4. the log ----
    stepTxt = "archiving " & oldYear & "'s OT Log rows"
    Application.Calculation = xlCalculationManual
    Dim keep As Variant
    If nAll > 0 Then
        RecordIdMarks v
        keep = ArchiveLogRows(lo, v, arcName, oldYear, newFrom, nArc, nKeep)
    Else
        MakeArchiveSheet lo, arcName, oldYear, Empty, 0
    End If

    ' ---- 5. the roster's year ----
    stepTxt = "moving the roster to " & newYear
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells(2, C_PERIOD1).Value = newFrom
    WritePeriodHeaders ws
    Application.Calculate                             ' calculation is manual here: PStartRow / PeriodLastCol read these dates
    If wantPer <> havePer Then SetPeriodColumns ws, lastRow, (wantPer = 27)

    ' ---- 6. worked hours wait for the new year's Kronos ----
    stepTxt = "resetting Worked hours"
    For r = 1 To nRows
        If nm(r) <> "" Then ws.Cells(R_FIRST + r - 1, C_WORKED).Value = 0
    Next r
    SetKronosAwait DateSerial(newYear, 1, 1)
    If KronosAwaitFrom() <> DateSerial(newYear, 1, 1) Then Err.Raise vbObjectError + 713, , "the 'awaiting " & newYear & " Kronos' marker couldn't be stored"
    StampHoursHistoryYear oldYear
    ProtectSheet ws

    ' ---- 7. verify ----
    stepTxt = "checking the result"
    Application.Calculation = calcWas
    Application.CalculateFull
    problems = RolloverProblems(ws, lo, arcName, lastRow, nm, nRows, keep, newFrom, newYear, wantPer, nAll, nArc, nKeep, sumAll, sumArc, sumKeep)
    If problems = "" Then
        RunIntegrity nErr, nWarn
        If nErr > 0 Then
            DrawIntegritySheet
            problems = "- the integrity check now shows " & nErr & " PROBLEM(S) (see the Integrity sheet)" & vbNewLine
        End If
    End If
    If problems <> "" Then Err.Raise vbObjectError + 712, , "the result did NOT verify:" & vbNewLine & problems

    ' ---- done ----
    ClearNoSave
    Application.EnableCancelKey = xlInterrupt
    Application.EnableEvents = evWas
    ws.Activate
    Dim savedOk As Boolean, t0 As Date
    On Error Resume Next
    SetRosterStatus
    MarkCalDirty
    t0 = FileDateTime(ThisWorkbook.FullName)
    On Error GoTo 0
    Application.ScreenUpdating = True
    SaveIfWritable
    On Error Resume Next                              ' did the file on disk really change?
    Err.Clear
    savedOk = (FileDateTime(ThisWorkbook.FullName) <> t0)
    If Err.Number <> 0 Then savedOk = ThisWorkbook.Saved
    On Error GoTo 0
    MsgBox IIf(dry, "PRACTICE RUN finished (this is a copy)." & vbNewLine & vbNewLine, "") & _
           "Rolled over to " & newYear & " and verified." & vbNewLine & vbNewLine & _
           "- frozen copy (read-only): " & frozen & vbNewLine & _
           "- " & Format$(nArc, "#,##0") & " OT Log row(s) are on '" & arcName & "'; " & Format$(nKeep, "#,##0") & " carried into " & newYear & vbNewLine & _
           "- final " & oldYear & " totals for " & nPeople & " people are on the Year End sheet" & vbNewLine & vbNewLine & _
           "Next: Worked hours stay at 0 until you import a Kronos export saved after Jan 1, " & newYear & ". Click Member View to publish." & _
           IIf(savedOk, "", vbNewLine & vbNewLine & "THE FILE DID NOT SAVE. Press Ctrl+S now and check it saves before you close."), _
           IIf(savedOk, vbInformation, vbExclamation), "Rollover"
    Exit Sub

Failed:
    Dim em As String: em = Err.Description
    If Err.Number = 18 Then em = "it was interrupted (Esc or Ctrl+Break)"
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    If Not changed Then
        Application.ScreenUpdating = True
        MsgBox "The rollover couldn't start (" & stepTxt & "): " & em & vbNewLine & vbNewLine & "Nothing was changed.", vbExclamation, "Rollover"
        Exit Sub
    End If
    MarkNoSave
    If gAdminUnlocked = False Then ApplyView True          ' the view was switched for the frozen copy
    If madeFrozen Then                                      ' a failed run leaves no frozen file behind
        SetAttr frozen, vbNormal
        Kill frozen
    End If
    Application.ScreenUpdating = True
    MsgBox "The rollover STOPPED while " & stepTxt & ":" & vbNewLine & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close this file WITHOUT saving and reopen it: the file on disk is exactly as it was before you started." & _
           vbNewLine & IIf(madeFrozen, "The frozen copy from this attempt was removed. ", "") & "Send this message to the admin before trying again.", vbCritical, "Rollover"
    Exit Sub

BadDates:
    MsgBox "The pay-period dates above the roster couldn't be read (" & Err.Description & ")." & vbNewLine & vbNewLine & _
           "Nothing was changed. Send this message to the admin.", vbCritical, "Rollover"
End Sub

' "" when rows 2 and 4 hold nPer two-week periods from firstStart (start,
' start + 13); otherwise which column is wrong. Read-only.
Private Function PeriodGridProblem(ByVal ws As Worksheet, ByVal firstStart As Date, ByVal nPer As Long) As String
    Dim k As Long, hs As Variant, he As Variant, want As Double
    For k = 0 To nPer - 1
        want = Int(CDbl(firstStart)) + 14 * k
        hs = ws.Cells(2, C_PERIOD1 + k).Value2
        he = ws.Cells(4, C_PERIOD1 + k).Value2
        If IsError(hs) Or IsError(he) Then
            PeriodGridProblem = "column " & ColLetter(C_PERIOD1 + k) & " shows an error"
            Exit Function
        ElseIf IsEmpty(hs) Or IsEmpty(he) Or Not IsNumeric(hs) Or Not IsNumeric(he) Then
            PeriodGridProblem = "column " & ColLetter(C_PERIOD1 + k) & " has no date"
            Exit Function
        ElseIf CDbl(hs) <> want Or CDbl(he) <> want + 13 Then
            PeriodGridProblem = "column " & ColLetter(C_PERIOD1 + k) & " should run " & Format$(CDate(want), "mmm d, yyyy") & " to " & _
                                Format$(CDate(want + 13), "mmm d, yyyy")
            Exit Function
        End If
    Next k
End Function

' Row 2 = each period's start (T2 is the one typed date, the rest are the
' cell to the left + 14), row 3 "to", row 4 = start + 13, for the columns
' in use. Written at every rollover so nothing rests on old header formulas.
Private Sub WritePeriodHeaders(ByVal ws As Worksheet)
    ' the column count is read from row 2 directly: T2 has just changed and row 4 has not recalculated yet
    Dim lastPer As Long: lastPer = C_PERIODN
    Dim a As Variant, b As Variant
    a = ws.Cells(2, C_PERIODN).Value2: b = ws.Cells(2, C_PERIODX).Value2
    If Not IsError(a) And Not IsError(b) Then
        If Not IsEmpty(a) And Not IsEmpty(b) And IsNumeric(a) And IsNumeric(b) Then
            If CDbl(b) = CDbl(a) + 14 Then lastPer = C_PERIODX     ' a 27-period year is in place
        End If
    End If
    ' a date format first: a cell someone set to Text would store the formula as text
    ws.Range(ws.Cells(2, C_PERIOD1), ws.Cells(2, lastPer)).NumberFormat = "mmm d"
    ws.Range(ws.Cells(4, C_PERIOD1), ws.Cells(4, lastPer)).NumberFormat = "mmm d"
    ws.Range(ws.Cells(2, C_PERIOD1 + 1), ws.Cells(2, lastPer)).FormulaR1C1 = "=RC[-1]+14"
    ws.Range(ws.Cells(4, C_PERIOD1), ws.Cells(4, lastPer)).FormulaR1C1 = "=R2C+13"
    ws.Range(ws.Cells(3, C_PERIOD1), ws.Cells(3, lastPer)).Value = "to"
End Sub

' Rows and Hours Charged of a log table, counted the same way everywhere
' in the rollover (NumOr0, so a number stored as text can't cause a false mismatch)
Private Sub TableCountAndSum(ByVal lo As ListObject, ByRef n As Long, ByRef total As Double)
    n = 0: total = 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Sub
    Dim v As Variant: v = RangeTo2D(lo.ListColumns(L_CHARGED).DataBodyRange)
    Dim i As Long
    n = UBound(v, 1)
    For i = 1 To n
        total = total + NumOr0(v(i, 1))
    Next i
End Sub


' ====================================================================
'  YEAR END SHEET (admin; one row per person per closed year)
' ====================================================================
Private Function YearEndTable(ByVal create As Boolean) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(YE_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(YE_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(YE_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = YE_SHEET
            ws.Tab.Color = RGB(96, 125, 139)
        End If
        ws.Unprotect Password:=SHEET_PASSWORD
        ws.Range("A1").Resize(1, YE_NCOLS).Value = Array("Year", "Employee", "Home Unit", "Total Hours", "Worked", "Refused", _
                                                          "Provisional", "Rolled On", "Rolled By")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, YE_NCOLS), , xlYes)
        lo.Name = YE_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 7
        ws.Columns(2).ColumnWidth = 26
        ws.Columns(3).ColumnWidth = 14
        ws.Range(ws.Columns(4), ws.Columns(7)).ColumnWidth = 12
        ws.Columns(8).ColumnWidth = 20
        ws.Columns(9).ColumnWidth = 14
    End If
    Set YearEndTable = lo
End Function

Private Sub AppendYearEnd(ByVal y As Long, ByRef blk As Variant, ByRef nm() As String, ByRef fG() As Double, ByRef fH() As Double, _
                          ByRef fI() As Double, ByRef fC() As Double, ByVal nRows As Long)
    Dim lo As ListObject: Set lo = YearEndTable(True)
    If lo Is Nothing Then Err.Raise vbObjectError + 720, , "the Year End sheet couldn't be created"
    Dim n As Long, r As Long, k As Long
    For r = 1 To nRows
        If nm(r) <> "" Then n = n + 1
    Next r
    If n = 0 Then Exit Sub
    Dim out() As Variant: ReDim out(1 To n, 1 To YE_NCOLS)
    Dim stamp As Date: stamp = Now
    For r = 1 To nRows
        If nm(r) <> "" Then
            k = k + 1
            out(k, 1) = y
            out(k, 2) = nm(r)
            out(k, 3) = SafeStr(blk(r, C_HOME))
            out(k, 4) = fG(r): out(k, 5) = fH(r): out(k, 6) = fI(r): out(k, 7) = fC(r)
            out(k, 8) = stamp
            out(k, 9) = CurrentUser()
        End If
    Next r
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If Not lo.DataBodyRange Is Nothing Then
        If cnt = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    ws.Unprotect Password:=SHEET_PASSWORD
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, YE_NCOLS)
        .Columns(2).NumberFormat = "@"
        .Columns(3).NumberFormat = "@"
        .Value = out
        .Columns(8).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With
    ProtectSheet ws
End Sub

' v24: someone moved to the Former sheet during the closing year has no roster
' row at the rollover. Their final figures were kept when they were removed
' (modPeople.RemoveEmployee) and go on the Year End sheet with everyone else's.
Private Sub AppendYearEndFormer(ByVal y As Long)
    Dim fin() As Variant, n As Long, k As Long
    n = FormerFinals(y, fin)
    If n = 0 Then Exit Sub
    Dim lo As ListObject: Set lo = YearEndTable(True)
    If lo Is Nothing Then Err.Raise vbObjectError + 720, , "the Year End sheet couldn't be created"
    Dim out() As Variant: ReDim out(1 To n, 1 To YE_NCOLS)
    Dim stamp As Date: stamp = Now
    For k = 1 To n
        out(k, 1) = y
        out(k, 2) = fin(k, 1)
        out(k, 3) = fin(k, 2)
        out(k, 4) = fin(k, 3): out(k, 5) = fin(k, 4): out(k, 6) = fin(k, 5): out(k, 7) = fin(k, 6)
        out(k, 8) = stamp
        out(k, 9) = CurrentUser() & " (left during the year)"
    Next k
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If Not lo.DataBodyRange Is Nothing Then
        If cnt = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    ws.Unprotect Password:=SHEET_PASSWORD
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, YE_NCOLS)
        .Columns(2).NumberFormat = "@"
        .Columns(3).NumberFormat = "@"
        .Value = out
        .Columns(8).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With
    ProtectSheet ws
End Sub

' "Rolled over to 2027 on Dec 27, 2026 by gp" for NOTE_DAYS after a rollover ("" otherwise)
Public Function RolloverNote(Optional ByVal shortForm As Boolean = False) As String
    On Error GoTo Quit
    Dim lo As ListObject: Set lo = YearEndTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, best As Long, bestY As Long
    For i = 1 To UBound(v, 1)
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) And IsDate(v(i, 8)) Then
            If CLng(v(i, 1)) >= bestY Then bestY = CLng(v(i, 1)): best = i
        End If
    Next i
    If best = 0 Then Exit Function
    If Int(Date) - Int(CDate(v(best, 8))) > NOTE_DAYS Then Exit Function
    If RosterYear() <= bestY Then Exit Function          ' not moved yet (the frozen copy is saved at this point)
    If shortForm Then
        RolloverNote = "rolled over " & Format$(CDate(v(best, 8)), "mmm d")
    Else
        RolloverNote = "Rolled over to " & (bestY + 1) & " on " & Format$(CDate(v(best, 8)), "mmm d, yyyy") & " by " & SafeStr(v(best, 9))
    End If
Quit:
End Function


' ====================================================================
'  FROZEN COPY
' ====================================================================
' Opens the frozen copy read-only (its macros off) and compares every
' person's Total / Worked / Refused with what was just recorded.
' "" = it matches.
Private Function FrozenCopyProblem(ByVal path As String, ByVal sheetName As String, ByRef nm() As String, ByRef fG() As Double, _
                                   ByRef fH() As Double, ByRef fI() As Double, ByVal nRows As Long, ByVal lastRow As Long) As String
    Dim wb As Workbook, v As Variant, r As Long, bad As String, nBad As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim secWas As Long: secWas = Application.AutomationSecurity
    On Error GoTo Oops
    Application.EnableEvents = False
    Application.AutomationSecurity = 3            ' msoAutomationSecurityForceDisable
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Application.AutomationSecurity = secWas
    With wb.Worksheets(sheetName)
        v = .Range(.Cells(R_FIRST, 1), .Cells(lastRow, C_REFUSED)).Value2
    End With
    wb.Close SaveChanges:=False
    Set wb = Nothing
    ThisWorkbook.Activate
    Application.EnableEvents = evWas
    For r = 1 To nRows
        If nm(r) <> "" Then
            If Trim$(SafeStr(v(r, C_NAME))) <> nm(r) Then
                nBad = nBad + 1
                If nBad <= 3 Then bad = bad & "row " & (R_FIRST + r - 1) & " is " & SafeStr(v(r, C_NAME)) & ", expected " & nm(r) & "; "
            ElseIf IsError(v(r, C_TOTAL)) Or IsError(v(r, C_WORKED)) Or IsError(v(r, C_REFUSED)) Then
                nBad = nBad + 1
                If nBad <= 3 Then bad = bad & nm(r) & " shows an error; "
            ElseIf Abs(NumOr0(v(r, C_TOTAL)) - fG(r)) > TOL Or Abs(NumOr0(v(r, C_WORKED)) - fH(r)) > TOL Or _
                   Abs(NumOr0(v(r, C_REFUSED)) - fI(r)) > TOL Then
                nBad = nBad + 1
                If nBad <= 3 Then bad = bad & nm(r) & " has " & NumOr0(v(r, C_TOTAL)) & " in the copy, " & fG(r) & " recorded; "
            End If
        End If
    Next r
    If nBad > 3 Then bad = bad & "and " & (nBad - 3) & " more"
    FrozenCopyProblem = bad
    Exit Function
Oops:
    FrozenCopyProblem = "it couldn't be opened and read (" & Err.Description & ")"
    On Error Resume Next
    Application.AutomationSecurity = secWas
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.EnableEvents = evWas
End Function


' ====================================================================
'  THE LOG
' ====================================================================
' Highest number used per Ref ID prefix and the highest Entry ID, kept so
' numbering carries on after the old rows leave the live log
Private Sub RecordIdMarks(ByRef v As Variant)
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim i As Long, s As String, p As Long, pre As String, k As Long, mxE As Long, kk As Variant
    For i = 1 To UBound(v, 1)
        s = UCase$(Trim$(SafeStr(v(i, L_REF))))
        p = InStrRev(s, "-")
        If p > 1 And p < Len(s) Then
            If IsNumeric(Mid$(s, p + 1)) Then
                pre = Left$(s, p - 1)
                k = CLng(Val(Mid$(s, p + 1)))
                If Not d.Exists(pre) Then
                    d.Add pre, k
                ElseIf k > d(pre) Then
                    d(pre) = k
                End If
            End If
        End If
        If IsNumeric(v(i, L_ENTRY)) And Not IsEmpty(v(i, L_ENTRY)) Then
            If CLng(v(i, L_ENTRY)) > mxE Then mxE = CLng(v(i, L_ENTRY))
        End If
    Next i
    For Each kk In d.Keys
        If IsPlainPrefix(CStr(kk)) Then
            SetIdMark CStr(kk), CLng(d(kk))
            If IdMark(CStr(kk)) < CLng(d(kk)) Then Err.Raise vbObjectError + 731, , "the ID counter for " & CStr(kk) & "- couldn't be stored"
        End If
    Next kk
    SetIdMark "ENTRY", mxE
    If IdMark("ENTRY") < mxE Then Err.Raise vbObjectError + 732, , "the Entry ID counter couldn't be stored"
End Sub

Private Function IsPlainPrefix(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 8 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[A-Z]" Then Exit Function
    Next i
    IsPlainPrefix = True
End Function

' The archive sheet "OT Log 2026": same columns, values only, a table
' named tblOTLog2026. arc = the rows (or Empty when nArc = 0).
Private Sub MakeArchiveSheet(ByVal lo As ListObject, ByVal arcName As String, ByVal y As Long, ByRef arc As Variant, ByVal nArc As Long)
    Dim wa As Worksheet, la As ListObject
    Set wa = ThisWorkbook.Worksheets.Add(After:=lo.Parent)
    wa.Name = arcName
    wa.Tab.Color = RGB(120, 130, 140)
    wa.Range("A1").Resize(1, L_NCOLS).Value = LogHeaders()
    If nArc > 0 Then
        wa.Range("A2").Resize(nArc, L_NCOLS).Value = arc
        wa.Range("A2").Resize(nArc, 1).Offset(0, L_DATE - 1).NumberFormat = "mm/dd/yyyy"
        wa.Range("A2").Resize(nArc, 1).Offset(0, L_PERIOD - 1).NumberFormat = "mm/dd/yyyy"
        wa.Range("A2").Resize(nArc, 1).Offset(0, L_STAMP - 1).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    End If
    Set la = wa.ListObjects.Add(xlSrcRange, wa.Range("A1").Resize(IIf(nArc > 0, nArc, 1) + 1, L_NCOLS), , xlYes)
    la.Name = "tblOTLog" & y
    la.TableStyle = "TableStyleLight9"
    wa.Range(wa.Columns(1), wa.Columns(L_NCOLS)).ColumnWidth = 13
    wa.Columns(L_EMP).ColumnWidth = 24
    wa.Columns(L_REASON).ColumnWidth = 40
    ProtectSheet wa
End Sub

' Splits the log at cutDate: older rows go to the archive sheet, the rest
' are written back as the whole live table. Returns the kept rows (Empty
' when none). nArc / nKeep were counted by the caller and are checked here.
Private Function ArchiveLogRows(ByVal lo As ListObject, ByRef v As Variant, ByVal arcName As String, ByVal y As Long, _
                                ByVal cutDate As Date, ByVal nArc As Long, ByVal nKeep As Long) As Variant
    Dim arc() As Variant, keep() As Variant, i As Long, c As Long, a As Long, k As Long
    If nArc > 0 Then ReDim arc(1 To nArc, 1 To L_NCOLS)
    If nKeep > 0 Then ReDim keep(1 To nKeep, 1 To L_NCOLS)
    For i = 1 To UBound(v, 1)
        If GoesToArchive(v(i, L_PERIOD), cutDate) Then
            a = a + 1
            For c = 1 To L_NCOLS
                arc(a, c) = v(i, c)
            Next c
        Else
            k = k + 1
            For c = 1 To L_NCOLS
                keep(k, c) = v(i, c)
            Next c
        End If
    Next i
    If a <> nArc Or k <> nKeep Then Err.Raise vbObjectError + 730, , "the OT Log changed while the rollover was running"

    If nArc > 0 Then
        MakeArchiveSheet lo, arcName, y, arc, nArc
    Else
        MakeArchiveSheet lo, arcName, y, Empty, 0
    End If

    Dim wl As Worksheet: Set wl = lo.Parent
    wl.Unprotect Password:=SHEET_PASSWORD
    On Error Resume Next                         ' no filter set = an error; that's fine
    lo.AutoFilter.ShowAllData
    Err.Clear
    On Error GoTo 0
    lo.DataBodyRange.ClearContents
    If nKeep > 0 Then
        With lo.HeaderRowRange.Offset(1, 0).Resize(nKeep, L_NCOLS)
            .Value = keep
            .Columns(L_DATE).NumberFormat = "mm/dd/yyyy"
            .Columns(L_PERIOD).NumberFormat = "mm/dd/yyyy"
            .Columns(L_STAMP).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
        End With
        lo.Resize lo.HeaderRowRange.Resize(nKeep + 1)
        SortLogNewestFirst lo
        ArchiveLogRows = keep
    Else
        lo.Resize lo.HeaderRowRange.Resize(2)    ' one empty row, as on a brand-new log
    End If
    ProtectSheet wl
End Function

' True when an archived year's log holds a non-void row with this Ref ID
' (e.g. a leave topped up before the rollover)
Public Function ArchiveHasActiveRef(ByVal id As String) As Boolean
    id = UCase$(Trim$(id))
    If id = "" Then Exit Function
    Dim ws As Worksheet, lo As ListObject, a As Variant, b As Variant, i As Long
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like LOG_SHEET & " ####" Then
            If ws.ListObjects.Count > 0 Then
                Set lo = ws.ListObjects(1)
                If Not lo.DataBodyRange Is Nothing And lo.ListColumns.Count >= L_STATUS Then
                    a = RangeTo2D(lo.ListColumns(L_REF).DataBodyRange)
                    b = RangeTo2D(lo.ListColumns(L_STATUS).DataBodyRange)
                    For i = 1 To UBound(a, 1)
                        If UCase$(Trim$(SafeStr(a(i, 1)))) = id Then
                            If UCase$(Trim$(SafeStr(b(i, 1)))) <> "VOID" Then ArchiveHasActiveRef = True: Exit Function
                        End If
                    Next i
                End If
            End If
        End If
    Next ws
End Function


' ====================================================================
'  THE 27th PAY-PERIOD COLUMN (AT)
'  A 27-period year (first: 2033) uses AT; the year after drops it.
'  Sheet must be unprotected; T2 already holds the new year's start.
' ====================================================================
Private Sub SetPeriodColumns(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal want27 As Boolean)
    Dim r As Long
    If want27 Then
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(1, C_PERIODX), ws.Cells(lastRow + 200, C_PERIODX))) > 0 Then
            Err.Raise vbObjectError + 740, , "column " & ColLetter(C_PERIODX) & " is needed for the 27th pay period but isn't empty"
        End If
        ws.Columns(C_PERIODN).Copy
        ws.Columns(C_PERIODX).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        ws.Columns(C_PERIODX).ColumnWidth = ws.Columns(C_PERIODN).ColumnWidth
        ws.Columns(C_PERIODX).Hidden = False
        ws.Cells(2, C_PERIODX).FormulaR1C1 = "=RC[-1]+14"
        ws.Cells(3, C_PERIODX).Value = "to"
        ws.Cells(4, C_PERIODX).FormulaR1C1 = "=R2C+13"
    Else
        ws.Range(ws.Cells(2, C_PERIODX), ws.Cells(lastRow, C_PERIODX)).ClearContents
        ws.Columns(C_PERIODX).Hidden = True
    End If
    Application.Calculate                        ' PeriodLastCol reads the header just written
    If (PeriodLastCol() = C_PERIODX) <> want27 Then Err.Raise vbObjectError + 741, , "the pay-period columns couldn't be set for a " & IIf(want27, 27, 26) & "-period year"
    For r = R_FIRST To lastRow
        If Trim$(SafeStr(ws.Cells(r, C_NAME).Value)) <> "" Then
            WritePeriodRow ws, r
            WriteRefusedRow ws, r
            WriteTotalRow ws, r
        End If
    Next r
    If ws.AutoFilterMode Then                    ' the filter must span the period columns in use
        If ws.FilterMode Then ws.ShowAllData
        ws.AutoFilterMode = False
    End If
    EnsureRosterFilter ws
End Sub


' ====================================================================
'  VERIFY
' ====================================================================
Private Function RolloverProblems(ByVal ws As Worksheet, ByVal lo As ListObject, ByVal arcName As String, ByVal lastRow As Long, _
                                  ByRef nm() As String, ByVal nRows As Long, ByRef keep As Variant, ByVal newFrom As Date, _
                                  ByVal newYear As Long, ByVal wantPer As Long, ByVal nAll As Long, ByVal nArc As Long, _
                                  ByVal nKeep As Long, ByVal sumAll As Double, ByVal sumArc As Double, ByVal sumKeep As Double) As String
    Dim p As String, n As Long, s As Double, la As ListObject
    ' the two logs
    If nArc + nKeep <> nAll Then p = p & "- OT Log rows: " & nAll & " before, " & nArc & " archived + " & nKeep & " kept" & vbNewLine
    If Abs(sumArc + sumKeep - sumAll) > TOL Then p = p & "- Hours Charged: " & sumAll & " before, " & (sumArc + sumKeep) & " after" & vbNewLine
    TableCountAndSum lo, n, s
    If n <> nKeep Then p = p & "- live OT Log has " & n & " rows, expected " & nKeep & vbNewLine
    If Abs(s - sumKeep) > TOL Then p = p & "- live OT Log Hours Charged is " & s & ", expected " & sumKeep & vbNewLine
    If Not SheetExists(arcName) Then
        p = p & "- the sheet '" & arcName & "' is missing" & vbNewLine
    Else
        If ThisWorkbook.Worksheets(arcName).ListObjects.Count > 0 Then Set la = ThisWorkbook.Worksheets(arcName).ListObjects(1)
        TableCountAndSum la, n, s
        If n <> nArc Then p = p & "- '" & arcName & "' has " & n & " rows, expected " & nArc & vbNewLine
        If Abs(s - sumArc) > TOL Then p = p & "- '" & arcName & "' Hours Charged is " & s & ", expected " & sumArc & vbNewLine
    End If
    ' the roster's year
    If RosterYearFrom() <> newFrom Or RosterYear() <> newYear Then p = p & "- the roster's first period isn't " & Format$(newFrom, "mmm d, yyyy") & vbNewLine
    If PeriodLastCol() - C_PERIOD1 + 1 <> wantPer Then p = p & "- the roster shows " & (PeriodLastCol() - C_PERIOD1 + 1) & " pay periods, expected " & wantPer & vbNewLine
    ' every period header moved, not just the first
    Dim gridBad As String: gridBad = PeriodGridProblem(ws, newFrom, wantPer)
    If gridBad <> "" Then p = p & "- pay-period dates: " & gridBad & vbNewLine
    If p = "" Then
        If RosterYearTo() <> newFrom + 14 * wantPer - 1 Then p = p & "- the roster's last day isn't " & Format$(newFrom + 14 * wantPer - 1, "mmm d, yyyy") & vbNewLine
    End If
    If p <> "" Then RolloverProblems = p: Exit Function

    ' each person: Refused = the kept rows for the new year, Worked = 0, Total = Worked + Refused + Provisional
    Dim lastStart As Long: lastStart = CLng(Int(CDbl(CDate(ws.Cells(PStartRow(), PeriodLastCol()).Value))))
    Dim sums As Object: Set sums = CreateObject("Scripting.Dictionary")
    sums.CompareMode = 1
    Dim i As Long, emp As String, ps As Double
    If IsArray(keep) Then
        For i = 1 To UBound(keep, 1)
            emp = Trim$(SafeStr(keep(i, L_EMP)))
            If emp <> "" And UCase$(Trim$(SafeStr(keep(i, L_STATUS)))) <> "VOID" Then
                ps = 0
                If VarType(keep(i, L_PERIOD)) = vbString Then
                    ps = 0                                   ' text is never a period date (see GoesToArchive)
                ElseIf IsDate(keep(i, L_PERIOD)) Then
                    ps = Int(CDbl(CDate(keep(i, L_PERIOD))))
                ElseIf IsNumeric(keep(i, L_PERIOD)) And Not IsEmpty(keep(i, L_PERIOD)) Then
                    ps = Int(CDbl(keep(i, L_PERIOD)))
                End If
                ' on a period start of the new year exactly, as the roster formulas match it
                If ps >= CDbl(newFrom) And ps <= lastStart Then
                    If (ps - CDbl(newFrom)) / 14 = Int((ps - CDbl(newFrom)) / 14) Then sums(emp) = NumOr0(sums(emp)) + NumOr0(keep(i, L_CHARGED))
                End If
            End If
        Next i
    End If
    Dim blk As Variant: blk = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, C_REFUSED)).Value2
    Dim r As Long, nBad As Long
    Dim provLive As Boolean: provLive = ProvisionalLive()
    For r = 1 To nRows
        If nm(r) <> "" Then
            If Trim$(SafeStr(blk(r, C_NAME))) <> nm(r) Then
                nBad = nBad + 1
                If nBad <= 5 Then p = p & "- row " & (R_FIRST + r - 1) & " changed name" & vbNewLine
            ElseIf IsError(blk(r, C_TOTAL)) Or IsError(blk(r, C_REFUSED)) Or IsError(blk(r, C_PROV)) Or IsError(blk(r, C_WORKED)) Then
                nBad = nBad + 1
                If nBad <= 5 Then p = p & "- " & nm(r) & " shows an error" & vbNewLine
            ElseIf Abs(NumOr0(blk(r, C_REFUSED)) - NumOr0(sums(nm(r)))) > TOL Then
                nBad = nBad + 1
                If nBad <= 5 Then p = p & "- " & nm(r) & ": Refused shows " & NumOr0(blk(r, C_REFUSED)) & ", the " & newYear & " log rows give " & NumOr0(sums(nm(r))) & vbNewLine
            ElseIf Abs(NumOr0(blk(r, C_WORKED))) > TOL Then
                nBad = nBad + 1
                If nBad <= 5 Then p = p & "- " & nm(r) & ": Worked isn't 0" & vbNewLine
            ElseIf Abs(NumOr0(blk(r, C_TOTAL)) - (NumOr0(blk(r, C_WORKED)) + NumOr0(blk(r, C_REFUSED)) + IIf(provLive, NumOr0(blk(r, C_PROV)), 0))) > TOL Then
                nBad = nBad + 1
                If nBad <= 5 Then p = p & "- " & nm(r) & ": Total isn't Worked + Refused + Provisional" & vbNewLine
            End If
        End If
    Next r
    If nBad > 5 Then p = p & "- ... and " & (nBad - 5) & " more people" & vbNewLine
    RolloverProblems = p
End Function
