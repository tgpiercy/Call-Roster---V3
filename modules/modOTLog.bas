Attribute VB_Name = "modOTLog"
Option Explicit

' ====================================================================
'  OT LOG MAINTENANCE (admin)
'  The OT Log is the source of truth. Roster pay-period cells are
'  SUMIFS formulas over it, so every hour on the roster traces back
'  to a log row. Corrections are made by adding or voiding rows,
'  never by typing over roster cells.
' ====================================================================


' --------------------------------------------------------------------
'  Moves any typed values in the roster's pay-period cells into the log,
'  then (re)writes the formulas for every named row and checks that
'  each employee's period total is unchanged.
'  Used by the migration ("Opening Balance") and by RefreshRosterFormulas.
' --------------------------------------------------------------------
Public Function ConvertRosterToFormulas(ByVal entryType As String, ByVal refPrefix As String, _
                                        ByRef summary As String) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function

    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then summary = "No employees found on the roster.": Exit Function

    Dim calcMode As XlCalculation: calcMode = Application.Calculation

    ' repair any period formulas that were stored as text by an earlier run
    Dim rr As Long, cc As Long
    For rr = R_FIRST To lastRow
        If Trim$(SafeStr(ws.Cells(rr, C_NAME).Value)) <> "" Then
            For cc = C_PERIOD1 To C_PERIODN
                If IsBrokenFormula(ws.Cells(rr, cc)) Then
                    ws.Unprotect Password:=SHEET_PASSWORD
                    WritePeriodRow ws, rr
                    Exit For
                End If
            Next cc
        End If
    Next rr
    Application.Calculate

    Dim nPer As Long: nPer = C_PERIODN - C_PERIOD1 + 1
    Dim nRows As Long: nRows = lastRow - R_FIRST + 1
    Dim before() As Double: ReDim before(1 To nRows)
    Dim buf() As Variant: ReDim buf(1 To nRows * nPer, 1 To L_NCOLS)

    Dim r As Long, c As Long, i As Long, n As Long
    Dim cel As Range, v As Variant, nm As String, note As String
    Dim bad As String, total As Double, amt As Double, existing As Double, ps As Date
    Dim refId As String: refId = NextRefID(refPrefix)
    Dim stamp As Date: stamp = Now
    Dim user As String: user = CurrentUser()
    Dim sRow As Long: sRow = PStartRow()
    Dim oldI() As Variant: ReDim oldI(1 To nRows)
    Dim oldG() As Variant: ReDim oldG(1 To nRows)

    ' ---- 1. collect typed values + record each row's current totals ----
    ' A typed number in a cell that should be a formula is treated as the
    ' NEW total for that employee/period: only the difference from what
    ' the log already holds is added (so nothing is double-counted).
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            oldI(r - R_FIRST + 1) = ws.Cells(r, C_REFUSED).Value2
            oldG(r - R_FIRST + 1) = ws.Cells(r, C_TOTAL).Value2
            For c = C_PERIOD1 To C_PERIODN
                Set cel = ws.Cells(r, c)
                v = cel.Value2
                If IsRealNumber(v) Then before(r - R_FIRST + 1) = before(r - R_FIRST + 1) + CDbl(v)

                If Not IsOurFormula(cel) Then
                    ps = CDate(ws.Cells(sRow, c).Value)
                    If IsError(v) Then
                        bad = bad & cel.Address(False, False) & " (error)  "
                    ElseIf IsEmpty(v) Then
                        ' a cleared cell is only a problem if the log has hours there
                        If LogSumFor(nm, ps) <> 0 Then bad = bad & cel.Address(False, False) & " (cleared)  "
                    ElseIf Not IsRealNumber(v) Then
                        If Trim$(CStr(v)) <> "" Then bad = bad & cel.Address(False, False) & " (text)  "
                    Else
                        existing = LogSumFor(nm, ps)
                        amt = CDbl(v) - existing
                        If Abs(amt) > 0.0000001 Then
                            n = n + 1
                            note = "From roster cell " & cel.Address(False, False)
                            If existing <> 0 Then note = note & " (typed " & CDbl(v) & ", log had " & existing & ")"
                            If Not cel.Comment Is Nothing Then
                                note = note & ": " & Replace(Replace(Trim$(cel.Comment.Text), vbCr, " "), vbLf, " ")
                            End If
                            buf(n, L_REF) = refId
                            buf(n, L_TYPE) = entryType
                            buf(n, L_DATE) = ps
                            buf(n, L_EMP) = nm
                            buf(n, L_REASON) = note
                            buf(n, L_CHARGED) = amt
                            buf(n, L_PERIOD) = ps
                            buf(n, L_USER) = user
                            buf(n, L_STAMP) = stamp
                            total = total + amt
                        End If
                    End If
                End If
            Next c
        End If
    Next r

    If bad <> "" Then
        MsgBox "Nothing was changed. These pay-period cells can't be converted safely:" & vbNewLine & vbNewLine & _
               bad & vbNewLine & vbNewLine & _
               "(text) / (error): fix or clear the cell." & vbNewLine & _
               "(cleared): the log has hours there - put the formula back (copy it from the next cell) " & _
               "and use Void Entry or Add Adjustment instead.", vbCritical, "Roster check"
        Exit Function
    End If

    ' ---- 2. write the log rows FIRST (so the formulas have data) ----
    If n > 0 Then
        Dim data() As Variant
        ReDim data(1 To n, 1 To L_NCOLS)
        For i = 1 To n
            For c = 1 To L_NCOLS
                data(i, c) = buf(i, c)
            Next c
        Next i
        If Not AppendLogRows(data) Then Exit Function
    End If

    ' ---- 3. formulas (locked, so they can't be typed over) ----
    ws.Unprotect Password:=SHEET_PASSWORD
    Application.Calculation = xlCalculationManual
    Dim pf As String, pl As String
    pf = ColLetter(C_PERIOD1): pl = ColLetter(C_PERIODN)
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            If Not ws.Cells(r, C_NAME).HasFormula Then
                If SafeStr(ws.Cells(r, C_NAME).Value) <> nm Then ws.Cells(r, C_NAME).Value = nm
            End If
            WritePeriodRow ws, r
            ws.Cells(r, C_REFUSED).Formula = "=SUM(" & pf & r & ":" & pl & r & ")"
            ws.Cells(r, C_TOTAL).Formula = "=" & ColLetter(C_WORKED) & r & "+" & ColLetter(C_REFUSED) & r
            ws.Cells(r, C_REFUSED).Locked = True
            ws.Cells(r, C_TOTAL).Locked = True
        End If
    Next r
    Application.Calculation = calcMode
    Application.Calculate

    ' ---- 4. verify: period totals, Refused (I) and Total (G) unchanged ----
    Dim afterSum As Double, mism As String, k As Long
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            k = r - R_FIRST + 1
            afterSum = 0
            For c = C_PERIOD1 To C_PERIODN
                v = ws.Cells(r, c).Value2
                If IsRealNumber(v) Then afterSum = afterSum + CDbl(v)
            Next c
            If Abs(afterSum - before(k)) > 0.0001 Then
                mism = mism & "  " & nm & ": periods were " & before(k) & ", now " & afterSum & vbNewLine
            ElseIf Not SameNumber(oldI(k), ws.Cells(r, C_REFUSED).Value2) Then
                mism = mism & "  " & nm & ": Refused (I) was " & SafeStr(oldI(k)) & ", now " & SafeStr(ws.Cells(r, C_REFUSED).Value) & vbNewLine
            ElseIf Not SameNumber(oldG(k), ws.Cells(r, C_TOTAL).Value2) Then
                mism = mism & "  " & nm & ": Total (G) was " & SafeStr(oldG(k)) & ", now " & SafeStr(ws.Cells(r, C_TOTAL).Value) & vbNewLine
            End If
        End If
    Next r
    ProtectSheet ws

    If mism <> "" Then
        MarkNoSave
        MsgBox "STOP - these employees' totals changed after conversion:" & vbNewLine & vbNewLine & _
               mism & vbNewLine & "Saving is now blocked. Close the workbook WITHOUT saving.", _
               vbCritical, "Verification failed"
        Exit Function
    End If

    If n > 0 Then
        summary = n & " pay-period value(s) moved into the OT Log as '" & entryType & _
                  "' (" & refId & "), net " & Format$(total, "#,##0.##") & " hrs." & vbNewLine & _
                  "Every employee's period, Refused and Total hours were checked and are unchanged."
    Else
        summary = "No typed values found. Formulas refreshed for all " & nRows & " roster rows."
    End If
    ConvertRosterToFormulas = True
End Function

' a real number (not text that looks like one, not an error, not blank)
Private Function IsRealNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbDouble, vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal
            IsRealNumber = True
    End Select
End Function

Private Function SameNumber(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then
        SameNumber = (IsError(a) And IsError(b))
    Else
        SameNumber = (Abs(NumOr0(a) - NumOr0(b)) < 0.0001)
    End If
End Function


' --------------------------------------------------------------------
'  Run after adding a new employee row, or if someone typed into a
'  pay-period cell. Typed values are kept (moved into the log as
'  "Manual Entry") before the formulas are restored.
' --------------------------------------------------------------------
Public Sub RefreshRosterFormulas()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim s As String
    If ConvertRosterToFormulas("Manual Entry", "MAN", s) Then
        SaveIfWritable
        MsgBox s, vbInformation, "Roster refreshed"
    End If
End Sub


' --------------------------------------------------------------------
'  ADD ADJUSTMENT - select the employee's row on the roster, then run.
'  Replaces typing into pay-period cells / cell comments.
' --------------------------------------------------------------------
Public Sub AddAdjustment()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = RosterSheet()
    If Not ActiveSheet Is ws Then
        MsgBox "Click a cell in the employee's row on the '" & ws.Name & "' sheet, then run Add Adjustment.", vbInformation
        Exit Sub
    End If
    Dim r As Long: r = ActiveCell.Row
    Dim nm As String: nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
    If r < R_FIRST Or nm = "" Then
        MsgBox "Click a cell in an employee's row first.", vbInformation
        Exit Sub
    End If

    Dim hrs As Variant
    hrs = Application.InputBox(Prompt:="Hours to ADD to " & nm & "'s charged (refused) hours." & vbNewLine & _
                               "Use a negative number to remove hours.", _
                               Title:="Adjustment - " & nm, Type:=1)
    If VarType(hrs) = vbBoolean Then Exit Sub
    If CDbl(hrs) = 0 Then Exit Sub

    Dim dIn As Variant
    dIn = Application.InputBox(Prompt:="Date the adjustment applies to (decides the pay period):", _
                               Title:="Adjustment - " & nm, Default:=Format$(Date, "mm/dd/yyyy"), Type:=2)
    If VarType(dIn) = vbBoolean Then Exit Sub
    If Not IsDate(dIn) Then MsgBox "That isn't a valid date.", vbExclamation: Exit Sub
    Dim d As Date: d = CDate(dIn)
    Dim ps As Date: ps = PeriodStartFor(d)
    If Not ConfirmPeriod(ps) Then Exit Sub

    Dim why As String: why = AskRequired("Reason for the adjustment (kept in the audit log):", "Adjustment - " & nm)
    If why = "" Then Exit Sub

    If MsgBox("Add " & Format$(CDbl(hrs), "0.##") & " hrs to " & nm & vbNewLine & _
              "Pay period starting " & Format$(ps, "mmm d, yyyy") & vbNewLine & _
              "Reason: " & why, vbOKCancel + vbQuestion, "Confirm adjustment") <> vbOK Then Exit Sub

    Dim data() As Variant: ReDim data(1 To 1, 1 To L_NCOLS)
    Dim refId As String: refId = NextRefID("ADJ")
    data(1, L_REF) = refId
    data(1, L_TYPE) = "Adjustment"
    data(1, L_DATE) = d
    data(1, L_EMP) = nm
    data(1, L_REASON) = why
    data(1, L_CHARGED) = CDbl(hrs)
    data(1, L_PERIOD) = ps
    data(1, L_USER) = CurrentUser()
    data(1, L_STAMP) = Now
    If AppendLogRows(data) Then
        SaveIfWritable
        MsgBox "Adjustment " & refId & " added.", vbInformation
    End If
End Sub


' --------------------------------------------------------------------
'  VOID - takes hours back off the roster without deleting history.
'  Enter a Ref ID (S-0012, P-0003, ADJ-0002) to void every row of that
'  session, or #<Entry ID> (e.g. #245) to void a single row.
' --------------------------------------------------------------------
Public Sub VoidEntry()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then MsgBox "The OT Log is empty.": Exit Sub

    Dim kIn As Variant
    kIn = Application.InputBox(Prompt:="Enter the ID to void:" & vbNewLine & _
                               "  S-0012 / P-0003 / ADJ-0002  = every row of that session" & vbNewLine & _
                               "  #245  = one row (the Entry ID column)", _
                               Title:="Void Entry", Type:=2)
    If VarType(kIn) = vbBoolean Then Exit Sub
    Dim k As String: k = UCase$(Trim$(CStr(kIn)))
    If k = "" Then Exit Sub

    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, cnt As Long, hrs As Double, nameList As String
    Dim hit() As Boolean: ReDim hit(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then
            If Left$(k, 1) = "#" Then
                hit(i) = (SafeStr(v(i, L_ENTRY)) = Mid$(k, 2))
            Else
                hit(i) = (UCase$(SafeStr(v(i, L_REF))) = k)
            End If
            If hit(i) Then
                cnt = cnt + 1
                hrs = hrs + NumOr0(v(i, L_CHARGED))
                If cnt <= 15 Then nameList = nameList & "  " & SafeStr(v(i, L_EMP)) & "  (" & NumOr0(v(i, L_CHARGED)) & " hrs)" & vbNewLine
            End If
        End If
    Next i

    If cnt = 0 Then MsgBox "No active (non-void) rows found for " & k & ".", vbInformation: Exit Sub
    If cnt > 15 Then nameList = nameList & "  ...and " & (cnt - 15) & " more" & vbNewLine

    Dim warn As String
    If Left$(k, 3) = "OB-" Then warn = vbNewLine & "WARNING: this is the opening balance carried over from the old roster." & vbNewLine

    If MsgBox("Void " & cnt & " row(s) for " & k & "?" & vbNewLine & vbNewLine & nameList & vbNewLine & _
              "Total hours charged coming off the roster: " & Format$(hrs, "0.##") & vbNewLine & warn & vbNewLine & _
              "The rows stay in the log, marked VOID.", vbOKCancel + vbExclamation, "Confirm void") <> vbOK Then Exit Sub

    Dim why As String: why = AskRequired("Reason for voiding " & k & ":", "Void Entry")
    If why = "" Then Exit Sub

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim stampText As String: stampText = why & " - " & CurrentUser() & " " & Format$(Now, "mm/dd/yyyy hh:mm")
    ws.Unprotect Password:=SHEET_PASSWORD
    For i = 1 To UBound(v, 1)
        If hit(i) Then
            lo.DataBodyRange.Cells(i, L_STATUS).Value = "VOID"
            lo.DataBodyRange.Cells(i, L_VOIDNOTE).Value = stampText
        End If
    Next i
    ProtectSheet ws
    Application.Calculate

    ' take voided call-outs / awards back off the calendar
    Dim calN As Long
    On Error Resume Next
    For i = 1 To UBound(v, 1)
        If hit(i) Then
            Select Case SafeStr(v(i, L_OUTCOME))
                Case "Accepted", "Awarded"
                    calN = calN + ReleaseCoverage(SafeStr(v(i, L_REF)), SafeStr(v(i, L_EMP)))
            End Select
        End If
    Next i
    RefreshCalendar
    On Error GoTo 0

    SaveIfWritable
    MsgBox cnt & " row(s) voided. Roster totals updated." & _
           IIf(calN > 0, vbNewLine & "Calendar: " & calN & " line(s) updated.", ""), vbInformation
End Sub


' --------------------------------------------------------------------
'  RENAME - the log matches on the name in column D, so a name must be
'  changed here (roster + every log row together), not just in the cell.
' --------------------------------------------------------------------
Public Sub RenameEmployee()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Sub

    Dim ws As Worksheet: Set ws = RosterSheet()
    If Not ActiveSheet Is ws Then MsgBox "Click a cell in the employee's row on the roster first.": Exit Sub
    Dim r As Long: r = ActiveCell.Row
    Dim oldName As String: oldName = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
    If r < R_FIRST Or oldName = "" Then MsgBox "Click a cell in an employee's row first.": Exit Sub

    Dim newName As String
    newName = Trim$(SafeStr(Application.InputBox(Prompt:="New name for " & oldName & ":", _
                Title:="Rename employee", Default:=oldName, Type:=2)))
    If newName = "" Or newName = "False" Or newName = oldName Then Exit Sub

    Dim rr As Long
    For rr = R_FIRST To RosterLastRow()
        If rr <> r And StrComp(Trim$(SafeStr(ws.Cells(rr, C_NAME).Value)), newName, vbTextCompare) = 0 Then
            MsgBox newName & " is already on the roster (row " & rr & ").", vbExclamation
            Exit Sub
        End If
    Next rr

    Dim cnt As Long, i As Long, v As Variant
    Dim logWs As Worksheet: Set logWs = lo.Parent
    logWs.Unprotect Password:=SHEET_PASSWORD
    If Not lo.DataBodyRange Is Nothing Then
        v = RangeTo2D(lo.ListColumns(L_EMP).DataBodyRange)
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, 1))), oldName, vbTextCompare) = 0 Then
                lo.ListColumns(L_EMP).DataBodyRange.Cells(i, 1).Value = newName
                cnt = cnt + 1
            End If
        Next i
    End If
    ProtectSheet logWs
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells(r, C_NAME).Value = newName
    ProtectSheet ws
    Application.Calculate
    SaveIfWritable
    MsgBox "Renamed to " & newName & ". " & cnt & " log row(s) updated.", vbInformation
End Sub


' --------------------------------------------------------------------
'  PUBLISH MEMBER VIEW - a values-only .xlsx for union members.
'  No phone numbers, no macros, no passwords. Reasons are left out
'  unless PUBLISH_REASONS = True (they can contain personal details).
' --------------------------------------------------------------------
Public Sub PublishMemberView()
    If Not RequireAdmin() Then Exit Sub
    If SheetExists(LOG_SHEET) Then
        Dim chk As ListObject
        On Error Resume Next
        Set chk = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
        On Error GoTo 0
        If Not chk Is Nothing Then
            If Not LogLayoutCurrent(chk) Then MsgBox "Run UpgradeOTLogLayout first (Alt+F8).", vbExclamation: Exit Sub
        End If
    End If
    Dim src As Worksheet: Set src = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub

    Application.Calculate
    Application.ScreenUpdating = False

    ' ---------- roster values ----------
    Dim cols As Variant
    cols = Array(C_NAME, C_SHIFT, C_TOTAL, C_WORKED, C_REFUSED, C_SENIORITY, _
                 12, 13, 14, 15, 16, 17, 18, 19)
    Dim nFixed As Long: nFixed = UBound(cols) + 1
    Dim nPer As Long: nPer = C_PERIODN - C_PERIOD1 + 1
    Dim v As Variant: v = src.Range(src.Cells(1, 1), src.Cells(lastRow, C_LAST)).Value

    Dim n As Long, r As Long, c As Long, k As Long
    For r = R_FIRST To lastRow
        If Trim$(SafeStr(v(r, C_NAME))) <> "" Then n = n + 1
    Next r
    If n = 0 Then Application.ScreenUpdating = True: Exit Sub

    Dim out() As Variant: ReDim out(1 To n + 1, 1 To nFixed + nPer)
    out(1, 1) = "Employee": out(1, 2) = "Shift": out(1, 3) = "Total Hours"
    out(1, 4) = "Worked Hours": out(1, 5) = "Refused Hours": out(1, 6) = "Plant Seniority"
    For c = 6 To nFixed - 1
        out(1, c + 1) = SafeStr(v(2, cols(c)))
    Next c
    Dim sR As Long: sR = PStartRow()
    For c = 1 To nPer
        If IsDate(v(sR, C_PERIOD1 + c - 1)) Then
            out(1, nFixed + c) = Format$(v(sR, C_PERIOD1 + c - 1), "mmm d") & " - " & _
                                 Format$(CDate(v(sR, C_PERIOD1 + c - 1)) + 13, "mmm d")
        Else
            out(1, nFixed + c) = "Period " & c
        End If
    Next c

    k = 1
    For r = R_FIRST To lastRow
        If Trim$(SafeStr(v(r, C_NAME))) <> "" Then
            k = k + 1
            For c = 0 To nFixed - 1
                out(k, c + 1) = CleanVal(v(r, cols(c)))
            Next c
            For c = 1 To nPer
                out(k, nFixed + c) = CleanVal(v(r, C_PERIOD1 + c - 1))
            Next c
        End If
    Next r

    Dim wb As Workbook: Set wb = Workbooks.Add(xlWBATWorksheet)
    Dim wsR As Worksheet: Set wsR = wb.Worksheets(1)
    wsR.Name = "Roster"
    With wsR.Range("A1").Resize(n + 1, nFixed + nPer)
        .Value = out
        .Sort Key1:=wsR.Range("A1"), Order1:=xlAscending, Header:=xlYes
    End With
    wsR.Rows(1).Font.Bold = True
    wsR.Columns(6).NumberFormat = "mm/dd/yyyy"
    wsR.Range(wsR.Cells(2, nFixed + 1), wsR.Cells(n + 1, nFixed + nPer)).NumberFormat = PERIOD_FMT
    wsR.Columns.AutoFit

    ' ---------- log values ----------
    Dim lo As ListObject
    If SheetExists(LOG_SHEET) Then
        On Error Resume Next
        Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
        On Error GoTo 0
    End If
    If Not lo Is Nothing Then
        Dim keepCols() As Long, nk As Long
        ReDim keepCols(1 To L_NCOLS)
        For c = 1 To L_NCOLS
            If c <> L_VOIDNOTE And (c <> L_REASON Or PUBLISH_REASONS) Then
                nk = nk + 1
                keepCols(nk) = c
            End If
        Next c

        Dim hdr As Variant: hdr = LogHeaders()
        Dim lv As Variant, nl As Long
        If Not lo.DataBodyRange Is Nothing Then
            lv = RangeTo2D(lo.DataBodyRange)
            For r = 1 To UBound(lv, 1)
                If SafeStr(lv(r, L_ENTRY)) <> "" Then nl = nl + 1
            Next r
        End If

        Dim lout() As Variant: ReDim lout(1 To nl + 1, 1 To nk)
        For c = 1 To nk
            lout(1, c) = hdr(keepCols(c) - 1)
        Next c
        k = 1
        If nl > 0 Then
            For r = 1 To UBound(lv, 1)
                If SafeStr(lv(r, L_ENTRY)) <> "" Then
                    k = k + 1
                    For c = 1 To nk
                        lout(k, c) = CleanVal(lv(r, keepCols(c)))
                    Next c
                End If
            Next r
        End If

        Dim wsL As Worksheet: Set wsL = wb.Worksheets.Add(After:=wsR)
        wsL.Name = "OT Log"
        wsL.Range("A1").Resize(nl + 1, nk).Value = lout
        wsL.Rows(1).Font.Bold = True
        For c = 1 To nk
            Select Case keepCols(c)
                Case L_DATE, L_PERIOD: wsL.Columns(c).NumberFormat = "mm/dd/yyyy"
                Case L_STAMP: wsL.Columns(c).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
            End Select
        Next c
        wsL.Columns.AutoFit
        If nl > 0 Then wsL.Range("A1").Resize(nl + 1, nk).AutoFilter
    End If

    ' ---------- about ----------
    Dim wsA As Worksheet: Set wsA = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    wsA.Name = "About"
    wsA.Range("A1").Value = "Overtime Call Out Roster - member view (read only)"
    wsA.Range("A1").Font.Bold = True
    wsA.Range("A3").Value = "Published:": wsA.Range("B3").Value = Now
    wsA.Range("B3").NumberFormat = "mmm d, yyyy h:mm AM/PM"
    wsA.Range("A4").Value = "Published by:": wsA.Range("B4").Value = CurrentUser()
    wsA.Range("A6").Value = "Refused hours for each pay period are the total of that employee's " & _
                            "non-void rows in the OT Log tab for that period."
    wsA.Range("A7").Value = "Rows marked VOID were corrected and no longer count."
    wsA.Columns("A").ColumnWidth = 14

    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        sh.Protect Password:=SHEET_PASSWORD, AllowFiltering:=True, AllowSorting:=True
    Next sh

    ' ---------- save ----------
    Dim folder As String: folder = EnsureSlash(PUBLISH_FOLDER)
    If Not FolderExists(folder) Then folder = EnsureSlash(ThisWorkbook.Path)
    Dim fullPath As String: fullPath = folder & PUBLISH_FILE
    Dim errMsg As String

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then errMsg = Err.Description
    On Error GoTo 0
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    If errMsg <> "" Then
        MsgBox "Could not save the member view to:" & vbNewLine & fullPath & vbNewLine & vbNewLine & errMsg, vbCritical
    Else
        MsgBox "Member view published:" & vbNewLine & fullPath, vbInformation
    End If
End Sub


' ====================================================================
'  helpers
' ====================================================================
Private Function CleanVal(ByVal x As Variant) As Variant
    If IsError(x) Then
        CleanVal = ""
    Else
        CleanVal = x
    End If
End Function

' InputBox that insists on a non-blank answer. Returns "" if cancelled.
Public Function AskRequired(ByVal prompt As String, ByVal title As String) As String
    Dim a As Variant
    Do
        a = Application.InputBox(Prompt:=prompt, Title:=title, Type:=2)
        If VarType(a) = vbBoolean Then Exit Function
        If Trim$(CStr(a)) <> "" Then AskRequired = Trim$(CStr(a)): Exit Function
        MsgBox "This is required.", vbExclamation
    Loop
End Function
