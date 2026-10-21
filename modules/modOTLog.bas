Attribute VB_Name = "modOTLog"
Option Explicit

Public Const OTL_VERSION As String = "OT Log v26.5 (Oct 19)"

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
    EnsureDataThrough                     ' v21: the provisional formulas need the data-through name

    ' repair any period formulas that were stored as text by an earlier run
    Dim rr As Long, cc As Long
    For rr = R_FIRST To lastRow
        If Trim$(SafeStr(ws.Cells(rr, C_NAME).Value)) <> "" Then
            For cc = C_PERIOD1 To PeriodLastCol()
                If IsBrokenFormula(ws.Cells(rr, cc)) Then
                    UnlockSheet ws
                    WritePeriodRow ws, rr
                    Exit For
                End If
            Next cc
        End If
    Next rr
    ProtectSheet ws                       ' v25 review: the repair unlocked it; the refusals below return with it protected
    Application.Calculate

    Dim lastPer As Long: lastPer = PeriodLastCol()             ' AS, or AT in a 27-period year (v23)
    Dim nPer As Long: nPer = lastPer - C_PERIOD1 + 1
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
            For c = C_PERIOD1 To lastPer
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
               "and use Void Entry or Add Adjustment instead." & ByHandHint(ws.Name), vbCritical, "Roster check"
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
    UnlockSheet ws
    On Error GoTo CalcBack                        ' v25 (E5): the calculation mode comes back whatever happens below
    Application.Calculation = xlCalculationManual
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            If Not ws.Cells(r, C_NAME).HasFormula Then
                If SafeStr(ws.Cells(r, C_NAME).Value) <> nm Then ws.Cells(r, C_NAME).Value = nm
            End If
            WritePeriodRow ws, r
            WriteRefusedRow ws, r
            WriteTotalRow ws, r                   ' v21: C = provisional, G = H + I + C
        End If
    Next r
    Application.Calculation = calcMode
    On Error GoTo 0
    Application.Calculate

    ' ---- 4. verify: period totals and Refused (I) unchanged; Total (G) =
    '      Worked + Refused + Provisional (G itself may grow by the
    '      provisional hours the first time v21 writes it) ----
    Dim afterSum As Double, mism As String, k As Long, expG As Double, provTot As Double
    Dim thr As Date: thr = EnsureRefusedThrough()        ' v26.5: once, not per cell
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            k = r - R_FIRST + 1
            afterSum = 0
            For c = C_PERIOD1 To lastPer
                v = ws.Cells(r, c).Value2
                If IsRealNumber(v) Then afterSum = afterSum + CDbl(v)
            Next c
            expG = NumOr0(ws.Cells(r, C_WORKED).Value2) + NumOr0(ws.Cells(r, C_REFUSED).Value2) + NumOr0(ws.Cells(r, C_PROV).Value2)
            provTot = provTot + NumOr0(ws.Cells(r, C_PROV).Value2)
            ' v26.5: Refused (I) = the period cells up to the current period (RefusedThrough), no longer every cell -
            ' so the check is against that sum, not against the old I (the first refresh after v26.5 moves future
            ' periods' charges out of I by design)
            Dim expI As Double: expI = 0
            For c = C_PERIOD1 To lastPer
                v = ws.Cells(r, c).Value2
                If IsRealNumber(v) Then
                    If RefusedCounts(CDate(ws.Cells(sRow, c).Value), thr) Then expI = expI + CDbl(v)
                End If
            Next c
            If Abs(afterSum - before(k)) > 0.0001 Then
                mism = mism & "  " & nm & ": periods were " & before(k) & ", now " & afterSum & vbNewLine
            ElseIf Not SameNumber(expI, ws.Cells(r, C_REFUSED).Value2) Then
                mism = mism & "  " & nm & ": Refused (I) is " & SafeStr(ws.Cells(r, C_REFUSED).Value) & ", expected " & expI & _
                       " (the periods through " & Format$(RefusedThrough(), "mmm d") & "); was " & SafeStr(oldI(k)) & vbNewLine
            ElseIf IsError(ws.Cells(r, C_PROV).Value2) Or IsError(ws.Cells(r, C_TOTAL).Value2) Then
                mism = mism & "  " & nm & ": Provisional (C) or Total (G) shows an error" & vbNewLine
            ElseIf Not SameNumber(expG, ws.Cells(r, C_TOTAL).Value2) Then
                mism = mism & "  " & nm & ": Total (G) is " & SafeStr(ws.Cells(r, C_TOTAL).Value) & ", expected " & expG & _
                       " (Worked + Refused + Provisional); was " & SafeStr(oldG(k)) & vbNewLine
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
                  "' (" & refId & "), net " & HoursText(total) & " hrs." & vbNewLine & _
                  "Every employee's period and Refused hours were checked and are unchanged; " & _
                  "Total = Worked + Refused + Provisional on every row."
    Else
        summary = "No typed values found. Formulas refreshed for all " & nRows & " roster rows " & _
                  "(Total = Worked + Refused + Provisional on every row)."
    End If
    summary = summary & vbNewLine & "Provisional accepted OT (not yet in Kronos, data through " & _
              Format$(EnsureDataThrough(), "mmm d") & "): " & HoursText(provTot) & " hrs in total."
    ConvertRosterToFormulas = True
    Exit Function
CalcBack:
    Dim cbN As Long, cbM As String: cbN = Err.Number: cbM = Err.Description
    Application.Calculation = calcMode
    ProtectSheet ws
    Err.Raise cbN, , cbM

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
'  Run after adding a new employee row, or if a formula was lost in a
'  pay-period cell. Typed values are kept (moved into the log as
'  "Manual Entry") before the formulas are restored.
' --------------------------------------------------------------------
Public Sub RefreshRosterFormulas()
    On Error GoTo Oops                         ' v24.3.1 (ReportError puts calculation back to automatic if the convert raised while manual)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim s As String
    If ConvertRosterToFormulas("Manual Entry", "MAN", s) Then
        SaveIfWritable
        MsgBox s, vbInformation, "Roster refreshed"
    End If
    Exit Sub
Oops:
    ReportError "RefreshRosterFormulas"
End Sub


' --------------------------------------------------------------------
'  ADD ADJUSTMENT - select the employee's row on the roster, then run.
'  Replaces typing into pay-period cells / cell comments.
' --------------------------------------------------------------------
Public Sub AddAdjustment()
    On Error GoTo Oops
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
                               Title:="Adjustment - " & nm, Default:=DateText(Date), Type:=2)
    If VarType(dIn) = vbBoolean Then Exit Sub
    Dim d As Date                                  ' v25: the one parser
    If Not ParseTypedDate(CStr(dIn), d) Then MsgBox "That isn't a valid date (e.g. Oct 12).", vbExclamation: Exit Sub
    Dim ps As Date: ps = PeriodStartFor(d)
    If Not ConfirmPeriod(ps) Then Exit Sub

    ' v21: the type is required (Leave return / Unit move / New hire / Correction / Other)
    Dim adjType As String: adjType = AskAdjustmentType(nm)
    If adjType = "" Then Exit Sub

    Dim why As String: why = AskRequired("Reason for the adjustment (kept in the audit log):", "Adjustment - " & nm)
    If why = "" Then Exit Sub

    If MsgBox("Add " & Format$(CDbl(hrs), "0.##") & " hrs to " & nm & vbNewLine & _
              "Pay period starting " & Format$(ps, "mmm d, yyyy") & vbNewLine & _
              "Type: " & adjType & vbNewLine & _
              "Reason: " & why, vbOKCancel + vbQuestion, "Confirm adjustment") <> vbOK Then Exit Sub

    Dim data() As Variant: ReDim data(1 To 1, 1 To L_NCOLS)
    Dim refId As String: refId = NextRefID("ADJ")
    data(1, L_REF) = refId
    data(1, L_TYPE) = "Adjustment"
    data(1, L_DATE) = d
    data(1, L_TASK) = adjType                      ' Reason column: the adjustment type (v21)
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
    Exit Sub
Oops:
    ReportError "AddAdjustment"
End Sub


' --------------------------------------------------------------------
'  VOID - takes hours back off the roster without deleting history.
'  Enter a Ref ID (S-0012, P-0003, ADJ-0002) to void every row of that
'  session, or #<Entry ID> (e.g. #245) to void a single row.
' --------------------------------------------------------------------
Public Sub VoidEntry()
    On Error GoTo Oops
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
    VoidEntryFor k, ""
    Exit Sub
Oops:
    ReportError "VoidEntry"
End Sub

' v25.5 (A9): the confirm, the reason and the void for one key - Void Entry's own path, shared with Undo last.
' lead = a line put before the question ("" for Void Entry)
Private Sub VoidEntryFor(ByVal k As String, ByVal lead As String)
    Dim lo As ListObject: Set lo = LogTable()
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim cnt As Long, hrs As Double, nameList As String
    Dim hit() As Boolean
    cnt = VoidHits(v, k, hit, hrs, nameList)

    If cnt = 0 Then MsgBox "No active (non-void) rows found for " & k & ".", vbInformation: Exit Sub

    Dim warn As String
    If Left$(k, 3) = "OB-" Then warn = vbNewLine & "WARNING: this is the opening balance carried over from the old roster." & vbNewLine

    If MsgBox(IIf(lead <> "", lead & vbNewLine & vbNewLine, "") & "Void " & cnt & " row(s) for " & k & "?" & vbNewLine & vbNewLine & nameList & vbNewLine & _
              "Total hours charged coming off the roster: " & HoursText(hrs) & vbNewLine & warn & vbNewLine & _
              "The rows stay in the log, marked VOID.", vbOKCancel + vbExclamation, "Confirm void") <> vbOK Then Exit Sub

    Dim why As String: why = AskRequired("Reason for voiding " & k & ":", "Void Entry")
    If why = "" Then Exit Sub

    Dim calN As Long
    cnt = VoidRef(k, why, calN)
    MsgBox cnt & " row(s) voided. Roster totals updated." & _
           IIf(calN > 0, vbNewLine & "Calendar: " & calN & " line(s) updated.", ""), vbInformation
End Sub

' v25.5 (A9) "Undo my last log entry" (the admin Desk; "Undo last entry" on the bar): the newest non-void OT Log entry THIS login
' made TODAY - by its Logged By and Timestamp, the log newest first - is voided through Void Entry's own confirm
' and reason (every row of that Ref ID, as Void Entry does). Nothing new is written. Admin: void is an admin's job.
Public Sub UndoLast()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then MsgBox "The OT Log is empty.", vbInformation, "Undo last entry": Exit Sub
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim u As String: u = LCase$(Trim$(CurrentUser()))
    Dim i As Long, best As Long, bestAt As Date, st As Variant
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" And LCase$(Trim$(SafeStr(v(i, L_USER)))) = u Then
            st = v(i, L_STAMP)
            If IsDate(st) Then
                If Int(CDate(st)) = Date And CDate(st) > bestAt Then best = i: bestAt = CDate(st)
            End If
        End If
    Next i
    If best = 0 Then
        MsgBox "Nothing to undo: no entry logged by " & CurrentUser() & " today is still active." & vbNewLine & vbNewLine & _
               "Void Entry voids any entry by its number.", vbInformation, "Undo last entry"
        Exit Sub
    End If
    Dim k As String: k = UCase$(SafeStr(v(best, L_REF)))
    If k = "" Then MsgBox "Your newest entry today has no Ref ID (row #" & SafeStr(v(best, L_ENTRY)) & "): void it by its Entry ID with Void Entry.", vbInformation, "Undo last entry": Exit Sub
    ' review: an undo voids the OT LOG rows only, as Void Entry does; what a P- / LV- / UM- / NH- ref leaves behind is named
    Dim lead As String
    lead = "Your last OT Log entry today: " & k & " (" & SafeStr(v(best, L_TYPE)) & ", " & SafeStr(v(best, L_EMP)) & ", " & Format$(bestAt, "h:mm AM/PM") & ")."
    If Left$(k, 2) = "P-" Then
        lead = lead & vbNewLine & "The posting stays awarded: reopen it afterwards (Alt+F8 > ReopenPosting, " & k & ") to award it again."
    ElseIf Left$(k, 3) = "LV-" Then
        lead = lead & vbNewLine & "The leave stays closed: book it again with Leave if they are still away."
    ElseIf Left$(k, 3) = "UM-" Or Left$(k, 3) = "NH-" Then
        lead = lead & vbNewLine & "The roster row keeps the move / the new hire: the Employee pop-up takes it back."
    End If
    VoidEntryFor k, lead
    Exit Sub
Oops:
    ReportError "UndoLast"
End Sub

' Which non-void rows a key means: "S-0012" = every row of that Ref ID, "#245" = one
' Entry ID. Fills hit(1..rows); returns how many, their charged hours and a short list.
Private Function VoidHits(ByRef v As Variant, ByVal k As String, ByRef hit() As Boolean, ByRef hrs As Double, _
                          ByRef nameList As String) As Long
    Dim i As Long, cnt As Long
    ReDim hit(1 To UBound(v, 1))
    hrs = 0: nameList = ""
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
    If cnt > 15 Then nameList = nameList & "  ...and " & (cnt - 15) & " more" & vbNewLine
    VoidHits = cnt
End Function

' v24.2: the voiding itself, with no questions. Marks every non-void row of the key
' VOID with the reason, takes voided call-outs / awards back off the calendar and
' saves. Returns the number of rows voided. Void Entry and SmokeTest both call it.
Public Function VoidRef(ByVal k As String, ByVal why As String, ByRef calN As Long) As Long
    calN = 0
    k = UCase$(Trim$(k))
    If k = "" Or Trim$(why) = "" Then Exit Function
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, cnt As Long, hrs As Double, nameList As String
    Dim hit() As Boolean
    cnt = VoidHits(v, k, hit, hrs, nameList)
    If cnt = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim stampText As String: stampText = why & " - " & CurrentUser() & " " & Format$(Now, "mm/dd/yyyy hh:mm")
    ' v26.5.5 (A14 + A3): a stop part-way (an error, Esc) would leave some of the key's rows VOID and
    ' others live, and the sheet unprotected: the save is blocked and the key is named
    Application.EnableCancelKey = xlErrorHandler
    On Error GoTo Failed
    UnlockSheet ws
    For i = 1 To UBound(v, 1)
        If hit(i) Then
            lo.DataBodyRange.Cells(i, L_STATUS).Value = "VOID"
            lo.DataBodyRange.Cells(i, L_VOIDNOTE).Value = stampText
        End If
    Next i
    ProtectSheet ws
    Application.EnableCancelKey = xlInterrupt
    Application.Calculate

    ' take voided call-outs / awards back off the calendar
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
    VoidRef = cnt
    Exit Function
Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    LogError "VoidRef", en, em
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    ProtectSheet ws
    MarkNoSave
    MsgBox "Voiding " & k & " stopped part-way (" & em & "). Some of its " & cnt & " row(s) may be marked VOID and others not, " & _
           "so saving is blocked." & vbNewLine & vbNewLine & "Close WITHOUT saving, reopen, and void " & k & " again.", vbCritical, "Void entry"
    VoidRef = 0
End Function


' --------------------------------------------------------------------
'  RENAME - the name is the key in every table, so it is changed in all
'  of them together by the Employee pop-up (v24: modPeople.RenameEverywhere).
'  This macro keeps its name (old buttons, install notes): it opens the
'  pop-up on Rename, with the person whose row is selected.
' --------------------------------------------------------------------
Public Sub RenameEmployee()
    On Error GoTo Oops
    Dim nm As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = RosterSheet()
    If ActiveSheet Is ws Then
        If ActiveCell.Row >= R_FIRST Then nm = Trim$(SafeStr(ws.Cells(ActiveCell.Row, C_NAME).Value))
    End If
    On Error GoTo Oops
    ShowEmployeeRename nm
    Exit Sub
Oops:
    ReportError "RenameEmployee"
End Sub


' --------------------------------------------------------------------
'  PUBLISH MEMBER VIEW - the Member View button (OnAction name kept).
'  M1: publishes the members' workbook (modPublish.PublishMembers), values
'  only, pushed from the master. The old values-only .xlsx (which also
'  carried Logged By and Timestamp) is no longer written.
' --------------------------------------------------------------------
Public Sub PublishMemberView()
    On Error GoTo Oops
    PublishMembers
    Exit Sub
Oops:
    ReportError "PublishMemberView"
End Sub


' ====================================================================
'  helpers
' ====================================================================

' Numbered pick of the adjustment type (v21). Returns "" if cancelled.
Private Function AskAdjustmentType(ByVal nm As String) As String
    Dim types As Variant: types = AdjustmentTypes()
    Dim i As Long, menu As String
    For i = LBound(types) To UBound(types)
        menu = menu & "  " & (i - LBound(types) + 1) & " = " & types(i) & vbNewLine
    Next i
    Dim a As Variant, k As Long
    Do
        a = Application.InputBox(Prompt:="Type of adjustment - enter the number:" & vbNewLine & vbNewLine & menu, _
                                 Title:="Adjustment - " & nm, Type:=1)
        If VarType(a) = vbBoolean Then Exit Function
        k = CLng(Val(a))
        If k >= 1 And k <= UBound(types) - LBound(types) + 1 Then
            AskAdjustmentType = CStr(types(LBound(types) + k - 1))
            Exit Function
        End If
        MsgBox "Enter a number from 1 to " & (UBound(types) - LBound(types) + 1) & ".", vbExclamation
    Loop
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
