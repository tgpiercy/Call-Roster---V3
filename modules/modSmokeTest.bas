Attribute VB_Name = "modSmokeTest"
Option Explicit

' ====================================================================
'  SMOKE TEST (v24.2)
'
'  Alt+F8 > SmokeTest, on a PRACTICE COPY only (COPY or DRY RUN in the
'  file name). It does what an install note used to ask a person to do,
'  through the SAME functions the pop-ups call, with no pop-ups:
'
'    adds a test employee, creates a posting, marks a partial and a
'    full sign-up and submits the award, logs a call-out with a part
'    acceptance, books time off and a sick call, logs a shift switch,
'    changes the test person's crew from a date (v26.5: through the
'    Employee pop-up's Edit), moves them to another unit, plans a unit
'    move for tomorrow and cancels it (v24.3: the lists must follow the
'    date), edits their phone, seniority and a code through Edit
'    (v26.5), renames them, publishes to a scratch
'    members' file and reads it back, voids everything it logged,
'    removes the test person and runs the integrity check.
'
'  Then every person's totals and the OT Log's charged total must be
'  what they were before it started. The first failure stops the run.
'  It NEVER saves: while it runs gSmokeRun is on (modOvertimeEntry),
'  and when it ends saving is blocked, so the copy can only be closed.
'
'  It holds no rules. If a step here disagrees with a pop-up, the
'  pop-up's function is right and this module is wrong.
'
'  NOT covered (these rules live inside the pop-ups, not in a module):
'  the shift-switch eligibility stops and payback checks, "vacation
'  and personal days on scheduled shifts only", the already-booked-sick
'  check. Queued for v25.
' ====================================================================

Public Const SMK_VERSION As String = "Smoke test v26.5.5 (Oct 21)"

Private Const SMK_EMP As String = "ZZ Test, Person"
Private Const SMK_EMP2 As String = "ZZ Test, Renamed"
Private Const SMK_WHY As String = "SmokeTest"
Private Const SMK_STEPS As Long = 20          ' v25: + no takers / out of dept / reopen, leave return, Kronos headings, planted leave words; v25.3: + the Desk; v25.5: the two v25.2 draw compares out with the legacy writers; v26.5: + the Employee pop-up's Edit

Private Type SmkState
    Emp As String           ' the test person's current name
    Unit As String          ' their home unit for the run
    Unit2 As String         ' the unit they are moved to ("" = no second unit with people)
    Code As String          ' grade code used in Unit
    Code2 As String
    Peer As String          ' a real person in Unit (the other side of the shift switch)
    PostDate As Date
    PostShift As String
    PostId As String
    PostId2 As String       ' v24.3.1: step 9's open posting in Unit (a list row is added by the plan and removed by the cancel)
    PostId3 As String       ' v25: step 13's posting - no takers, voided, reopened; step 16 plants a leave word on it
    PostId4 As String       ' v25: step 13's posting awarded out of department to the test person
    LeaveAbsId As String    ' v25: step 14's leave booking
    LeaveRef As String      ' v25: step 14's leave-return log ref (LV-)
    SessionId As String
    AbsId As String
    SickId As String
    SwitchId As String
    Note As String          ' what the step did, for the report
End Type


Public Sub SmokeTest()
    Const TTL As String = "Smoke test"
    If Not IsDryRunName(ThisWorkbook.Name) Then
        MsgBox "SmokeTest only runs on a practice copy: a file with COPY or DRY RUN in its name." & vbNewLine & vbNewLine & _
               "This is " & ThisWorkbook.Name & ". Save a copy first (with a copy of the members' template beside it) and run it there.", _
               vbExclamation, TTL
        Exit Sub
    End If
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If NoSave() Then
        MsgBox "Saving is already blocked in this copy (SmokeTest or another check ran in it)." & vbNewLine & _
               "Close it without saving, open it again and run SmokeTest once.", vbExclamation, TTL
        Exit Sub
    End If
    If LogTable() Is Nothing Then Exit Sub
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        MsgBox "Run Alt+F8 > SetupPeople in this copy first.", vbExclamation, TTL
        Exit Sub
    End If
    If RosterRowOf(SMK_EMP) > 0 Or RosterRowOf(SMK_EMP2) > 0 Or IsFormer(SMK_EMP) Or IsFormer(SMK_EMP2) Then
        MsgBox "A test person from an earlier run is still in this copy. Start from a fresh copy.", vbExclamation, TTL
        Exit Sub
    End If
    If Date < RosterYearFrom() Then
        MsgBox "This copy's roster year starts " & Format$(RosterYearFrom(), "mmm d, yyyy") & ", which is still ahead." & vbNewLine & _
               "SmokeTest needs a copy of the current year's roster.", vbExclamation, TTL
        Exit Sub
    End If
    If MsgBox("Run the smoke test in this copy?" & vbNewLine & vbNewLine & ThisWorkbook.Name & vbNewLine & vbNewLine & _
              "It adds a test person, logs and voids a posting, a call-out, time off and a switch, publishes to a scratch file, " & _
              "then checks every total is unchanged. About two minutes. Nothing is saved, and afterwards this copy can only be closed.", _
              vbOKCancel + vbQuestion, TTL) <> vbOK Then Exit Sub

    Dim st As SmkState, k As Long, ok As Boolean, why As String, rep As String, failedAt As Long
    Dim repNotes As String                         ' v24.3.1: the notes alone (the Error Log line)
    Dim repSkips As String                         ' v25.2: the skipped steps (the PASSED box shows these; 20 step names don't fit)
    Dim before As Object, logBefore As Double, nErr0 As Long, nWarn0 As Long, diff As String
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    MarkNoSave                                     ' from the first moment: even a run stopped with Esc can't be saved
    On Error GoTo Crashed
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    RunIntegrity nErr0, nWarn0
    st.Emp = SMK_EMP

    gSmokeRun = True
    Application.ScreenUpdating = False
    For k = 1 To SMK_STEPS
        why = "": st.Note = ""
        Application.StatusBar = "Smoke test: step " & k & " of " & SMK_STEPS & " - " & SmkName(k)
        ok = SmkStep(k, st, nErr0, why)
        If ok Then
            If st.Note <> "" Then repNotes = repNotes & k & ": " & st.Note & "; "
            If Left$(st.Note, 7) = "SKIPPED" Then repSkips = repSkips & Format$(k, "00") & "  skip  " & SmkName(k) & "  (" & st.Note & ")" & vbNewLine
            rep = rep & Format$(k, "00") & "  " & IIf(Left$(st.Note, 7) = "SKIPPED", "skip", "OK  ") & "  " & SmkName(k) & _
                  IIf(st.Note <> "", "  (" & st.Note & ")", "") & vbNewLine
        Else
            rep = rep & Format$(k, "00") & "  FAILED  " & SmkName(k) & ": " & why & vbNewLine
            failedAt = k
            Exit For
        End If
    Next k
    gSmokeRun = False
    Application.StatusBar = False
    Application.ScreenUpdating = scrWas
    Application.EnableEvents = evWas               ' a step that failed half-way may have left them off

    ' ---- everyone's totals and the log's charged total must be back where they were ----
    If failedAt = 0 Then
        Application.Calculate
        diff = TotalsDifference(before, RosterTotalsMap())
        If Abs(LogChargedTotal() - logBefore) > 0.0001 Then
            diff = diff & "  OT Log total Hours Charged was " & logBefore & ", now " & LogChargedTotal() & vbNewLine
        End If
    End If
    On Error Resume Next
    LogError "SmokeTest", 0, IIf(failedAt = 0 And diff = "", "all steps OK, totals unchanged", _
                                 IIf(failedAt > 0, "FAILED at step " & failedAt & " (" & SmkName(failedAt) & ")", "totals changed")), _
             IIf(failedAt = 0 And diff = "", Left$(repNotes, 500), Left$(why & diff, 400))
    On Error GoTo 0

    Dim head As String
    If failedAt > 0 Then
        head = "SMOKE TEST FAILED at step " & failedAt & " (" & SmkName(failedAt) & "):" & vbNewLine & Left$(why, 250) & vbNewLine & _
               "Send " & PRIMARY_ADMIN_NAME & " this box (a photo is fine)."
    ElseIf diff <> "" Then
        head = "SMOKE TEST FAILED: every step ran, but the totals did not come back:" & vbNewLine & diff & "Send " & PRIMARY_ADMIN_NAME & " this box."
    Else
        head = "SMOKE TEST PASSED: all " & SMK_STEPS & " steps OK, totals unchanged."
        ' v25.2: 20 step names don't fit a MsgBox with the closing text; a passed run's box shows the skipped
        ' steps (a failed run's box lists the steps: rep)
        rep = IIf(repSkips = "", "No step skipped." & vbNewLine, repSkips)
    End If
    If Len(head) > 420 Then head = Left$(head, 400) & " ..." & vbNewLine
    If Len(rep) > 820 - Len(head) Then rep = "... " & Right$(rep, 800 - Len(head))     ' MsgBox holds about 1,000 characters; the closing text is 150; the LAST lines (the failed one) stay (re-review)
    MsgBox head & vbNewLine & vbNewLine & rep & vbNewLine & _
           "Saving is blocked in this copy. Close it WITHOUT saving; nothing is kept. " & _
           "Make the next practice copy from the real roster, not from this one.", _
           IIf(failedAt = 0 And diff = "", vbInformation, vbCritical), TTL & "  (" & EXPECTED_RELEASE & ")"
    Exit Sub
Crashed:
    Dim em As String: em = Err.Description
    gSmokeRun = False
    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MarkNoSave
    LogError "SmokeTest", 0, "stopped outside a step: " & em, rep
    MsgBox "SMOKE TEST STOPPED: " & em & vbNewLine & vbNewLine & rep & vbNewLine & _
           "Saving is blocked in this copy. Close it; nothing is kept. Send " & PRIMARY_ADMIN_NAME & " this box.", vbCritical, TTL
End Sub

Private Function SmkName(ByVal k As Long) As String
    Select Case k
        Case 1: SmkName = "add a test employee"
        Case 2: SmkName = "create a posting"
        Case 3: SmkName = "partial + full sign-up, submit the award"
        Case 4: SmkName = "call-out with a part acceptance"
        Case 5: SmkName = "book time off and a sick call"
        Case 6: SmkName = "shift switch"
        Case 7: SmkName = "crew change from a date"
        Case 8: SmkName = "move to another unit"
        Case 9: SmkName = "plan a move for tomorrow (+ crew, old-unit posting), cancel it through a filtered list"
        Case 10: SmkName = "edit a person's details through the Employee pop-up (v26.5)"
        Case 11: SmkName = "rename"
        Case 12: SmkName = "publish to a scratch file and read it back"
        Case 13: SmkName = "no takers, award out of dept, void, reopen"
        Case 14: SmkName = "leave going and returning (top-up logged)"
        Case 15: SmkName = "Kronos headings on a made-up sheet (good, then one missing)"
        Case 16: SmkName = "a planted leave word stops the publish (Reason, then a calendar line)"
        Case 17: SmkName = "void everything it logged"
        Case 18: SmkName = "remove the test employee"
        Case 19: SmkName = "integrity check"
        Case 20: SmkName = "the Desk's figures follow the data"
    End Select
End Function

' One step. True = done (st.Note may say what or "SKIPPED ..."); False = why.
Private Function SmkStep(ByVal k As Long, ByRef st As SmkState, ByVal nErr0 As Long, ByRef why As String) As Boolean
    On Error GoTo Failed
    Select Case k
        Case 1: SmkStep = SmkAdd(st, why)
        Case 2: SmkStep = SmkPosting(st, why)
        Case 3: SmkStep = SmkAward(st, why)
        Case 4: SmkStep = SmkCallout(st, why)
        Case 5: SmkStep = SmkTimeOff(st, why)
        Case 6: SmkStep = SmkSwitch(st, why)
        Case 7: SmkStep = SmkCrew(st, why)
        Case 8: SmkStep = SmkMove(st, why)
        Case 9: SmkStep = SmkPlanned(st, why)
        Case 10: SmkStep = SmkEdit(st, why)
        Case 11: SmkStep = SmkRename(st, why)
        Case 12: SmkStep = SmkPublish(st, why)
        Case 13: SmkStep = SmkOutcomes(st, why)
        Case 14: SmkStep = SmkLeaveReturn(st, why)
        Case 15: SmkStep = SmkKronosHeadings(st, why)
        Case 16: SmkStep = SmkPlantedWords(st, why)
        Case 17: SmkStep = SmkVoid(st, why)
        Case 18: SmkStep = SmkRemove(st, why)
        Case 19: SmkStep = SmkIntegrity(st, nErr0, why)
        Case 20: SmkStep = SmkDesk(st, why)
    End Select
    If Not SmkStep And why = "" Then why = "the function returned False (a message above says why)"
    Exit Function
Failed:
    why = "error " & Err.Number & ": " & Err.Description
    SmkStep = False
End Function

' ---- 1. add ----
Private Function SmkAdd(ByRef st As SmkState, ByRef why As String) As Boolean
    ' the unit with the most people gives the later steps a list to work with; the runner-up is the move's target
    Dim depts As Variant: depts = DeptList()
    Dim u As Long, recs() As EmpRec, n As Long, best As Long, best2 As Long, i As Long
    For u = LBound(depts) To UBound(depts)
        n = LoadEligible(CStr(depts(u)), "ALL", recs)
        If n > best Then
            best2 = best: st.Unit2 = st.Unit: st.Code2 = st.Code
            best = n: st.Unit = CStr(depts(u))
            st.Code = recs(1).DeptCode: st.Peer = recs(1).EmpName
            For i = 1 To n                              ' a chargeable grade if the unit has one
                If recs(i).Grade <= CHARGE_MAX_GRADE Then st.Code = recs(i).DeptCode: Exit For
            Next i
        ElseIf n > best2 Then
            best2 = n: st.Unit2 = CStr(depts(u)): st.Code2 = recs(1).DeptCode
        End If
    Next u
    If best < 2 Then why = "no unit has two people with a code, so there is nothing to test against": Exit Function

    Dim pc As PeopleCalc, report As String
    If Not ComputeNewHire(st.Unit, Date, st.Emp, pc) Then why = "ComputeNewHire: " & pc.Warnings: Exit Function
    If Not SaveNewHire(st.Emp, "", "", "DAYS", Date, st.Unit, st.Code, Date, pc, pc.AddHours, "", report) Then Exit Function
    If RosterRowOf(st.Emp) = 0 Then why = "SaveNewHire said yes but the row isn't on the roster": Exit Function
    st.Note = st.Unit & " " & st.Code & ", " & HoursText(pc.AddHours) & " h"
    SmkAdd = True
End Function

' ---- 2. posting ----
Private Function SmkPosting(ByRef st As SmkState, ByRef why As String) As Boolean
    ' the first shift three or more weeks out with at least two people free to be listed
    Dim dd As Long, d As Date, sh As Variant, recs() As EmpRec, n As Long, offTxt As String, found As Boolean
    For dd = 21 To 60
        d = Date + dd
        For Each sh In Array("Days", "Nights")
            n = LoadEligible(st.Unit, CrewFilter(CREW_ALL), recs, DeptPostingMaxGrade(st.Unit), True, d)
            If n > 0 Then n = FilterNotFree(recs, n, d, CStr(sh), offTxt)
            If n >= 2 Then found = True: st.PostDate = d: st.PostShift = CStr(sh): Exit For
        Next sh
        If found Then Exit For
    Next dd
    If Not found Then why = "no shift in the next 60 days has two people free in " & st.Unit: Exit Function

    Dim p As PostRec
    p.Unit = st.Unit
    p.ShiftDate = st.PostDate
    p.ShiftLbl = st.PostShift
    p.Crew = CREW_ALL
    p.Reason = "Unit Operator Support"
    p.Hours = USUAL_SHIFT_HOURS
    p.Positions = 1
    p.ClosesOn = st.PostDate - POSTING_CLOSE_DAYS
    If Not CreatePostingFrom(p, False, False) Then Exit Function
    st.PostId = p.ID
    If PostingStatus(st.PostId) <> "Open" Then why = st.PostId & " isn't Open after being created": Exit Function
    st.Note = st.PostId & " " & Format$(st.PostDate, "mmm d") & " " & st.PostShift
    SmkPosting = True
End Function

' ---- 3. award ----
Private Function SmkAward(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim p As PostRec, lst() As PListRec, n As Long, iFull As Long, iPart As Long
    n = LoadPosting(st.PostId, p, lst)
    If n < 2 Then why = st.PostId & " has " & n & " on its list; two are needed": Exit Function
    ' last on the list takes the whole shift and is awarded; the one before signs up for part of it and
    ' isn't awarded (a full-shift sign-up comes first); anyone above them didn't sign up and is charged
    iFull = n: iPart = n - 1
    lst(iFull).SignedUp = True: lst(iFull).Awarded = True
    lst(iPart).SignedUp = True: lst(iPart).SignupHrs = 8
    ' v26.5: the shift is three or more weeks out (step 2), so its pay period is a later one: a refusal charged by this
    ' award must show in that period's column and stay OUT of Refused (I) until RefusedThrough reaches the period -
    ' checked by moving the hidden name there and back (the brief's acceptance; a practice copy is never saved)
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim pCol As Long: pCol = PeriodColumn(PeriodStartFor(st.PostDate))
    Dim iBefore() As Double, cBefore() As Double, rowOf() As Long, i As Long
    If n > 2 And pCol > 0 Then
        ReDim iBefore(1 To n): ReDim cBefore(1 To n): ReDim rowOf(1 To n)
        Application.Calculate
        For i = 1 To n - 2
            rowOf(i) = RosterRowOf(lst(i).EmpName)
            If rowOf(i) > 0 Then
                iBefore(i) = NumOr0(ws.Cells(rowOf(i), C_REFUSED).Value2)
                cBefore(i) = NumOr0(ws.Cells(rowOf(i), pCol).Value2)
            End If
        Next i
    End If
    If Not SubmitAward(p, lst, n) Then Exit Function
    If PostingStatus(st.PostId) = "Open" Then why = st.PostId & " is still Open after the award": Exit Function
    If Not HasActiveLogRows(st.PostId) Then why = "the award wrote no OT Log rows": Exit Function
    Dim futTxt As String: futTxt = IIf(n > 2, ", the shift's period is beyond this roster", ", nobody above the sign-ups to charge")
    If n > 2 And pCol > 0 Then
        Dim iC As Long: iC = 0
        For i = 1 To n - 2
            If lst(i).Charged > 0 And rowOf(i) > 0 Then iC = i: Exit For
        Next i
        If iC = 0 Then
            futTxt = ", nobody charged (nothing to check for the later period)"
        ElseIf OldRefusedFormula(ws.Cells(rowOf(iC), C_REFUSED)) Then
            futTxt = ", later-period check SKIPPED: this copy's Refused formula is the pre-v26.5 one (run RefreshRosterFormulas)"
        Else
            Application.Calculate
            Dim cAfter As Double: cAfter = NumOr0(ws.Cells(rowOf(iC), pCol).Value2)
            If cAfter - cBefore(iC) < 0.0001 Then why = lst(iC).EmpName & "'s " & Format$(PeriodStartFor(st.PostDate), "mmm d") & " period cell didn't grow with the refusal charge": Exit Function
            If Abs(NumOr0(ws.Cells(rowOf(iC), C_REFUSED).Value2) - iBefore(iC)) > 0.0001 Then why = lst(iC).EmpName & ": Refused (I) counted a refusal for a later pay period": Exit Function
            ' moving the cap to the shift's period brings in every period cell between the old cap and that period
            ' (on a copy of the real roster those may hold other charges), so the expected figure is read from the row
            Dim thrWas As Date: thrWas = RefusedThrough()
            Dim expMoved As Double, cc As Long, hd As Variant
            expMoved = iBefore(iC)
            For cc = C_PERIOD1 To PeriodLastCol()
                hd = ws.Cells(PStartRow(), cc).Value2
                If IsNumeric(hd) Then
                    If CDbl(hd) > CDbl(thrWas) And CDbl(hd) <= CDbl(PeriodStartFor(st.PostDate)) Then expMoved = expMoved + NumOr0(ws.Cells(rowOf(iC), cc).Value2)
                End If
            Next cc
            SetDateName REFUSED_NAME, PeriodStartFor(st.PostDate)
            Application.Calculate
            Dim iMoved As Double: iMoved = NumOr0(ws.Cells(rowOf(iC), C_REFUSED).Value2)
            SetDateName REFUSED_NAME, thrWas
            Application.Calculate
            If Abs(iMoved - expMoved) > 0.0001 Then why = "with RefusedThrough moved to " & Format$(PeriodStartFor(st.PostDate), "mmm d") & ", " & lst(iC).EmpName & "'s Refused showed " & iMoved & ", expected " & expMoved: Exit Function
            If Abs(NumOr0(ws.Cells(rowOf(iC), C_REFUSED).Value2) - iBefore(iC)) > 0.0001 Then why = "Refused (I) didn't come back after RefusedThrough was restored": Exit Function
            futTxt = ", " & lst(iC).EmpName & "'s " & HoursText(cAfter - cBefore(iC)) & " h refusal waits for its period"
        End If
    End If
    st.Note = "awarded #" & iFull & " of " & n & futTxt
    SmkAward = True
End Function

' ---- 4. call-out ----
Private Function SmkCallout(ByRef st As SmkState, ByRef why As String) As Boolean
    ' the first day shift, a week or more after the posting, with at least two people free
    Dim d As Date, dd As Long
    Dim recs() As EmpRec, n As Long, offTxt As String, taken() As Double, i As Long, summary As String
    For dd = 7 To 35
        d = st.PostDate + dd
        n = LoadEligible(st.Unit, CrewFilter(CREW_ALL), recs, 999, CALLOUT_INCLUDE_DAYS, d)
        If n > 0 Then n = FilterNotFree(recs, n, d, "Days", offTxt)
        If n >= 2 Then
            ' v26.5.5 (re-review): the two must be free of a 16 h cap (real OT on a copy of real data), or the charge and the 8 h differ
            If recs(1).MaxHours >= USUAL_SHIFT_HOURS And recs(2).MaxHours >= USUAL_SHIFT_HOURS Then Exit For
            n = 0
        End If
    Next dd
    If n < 2 Then why = "no day shift in the five weeks after the posting has two people free (and uncapped by the 16 h rule) in " & st.Unit: Exit Function
    ReDim taken(1 To n)
    recs(1).Outcome = "Refused"
    recs(2).Outcome = "Accepted": taken(2) = 8                  ' takes 8 of the 12 on offer: the shift closes (v25, D-E4)
    If n >= 3 Then recs(3).Outcome = "Left Message"             ' below the partial: offered nothing, never charged
    If Not SubmitCalloutSession(st.Unit, "Days", CREW_ALL, d, USUAL_SHIFT_HOURS, recs, taken, n, st.SessionId, summary) Then Exit Function

    ' the rows must carry the v25 rule (v25.1: this step still expected v24.1's "the rest offered down the
    ' list" and failed on Graeme's Oct 16 test): the first refusal charged on 12 h, the acceptance 8 h, the
    ' third offered 0 h and charged nothing, its row saying why
    Dim lo As ListObject: Set lo = LogTable()
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim want As Double, got As Double, nRows As Long, accH As Double, belowH As Double, belowWhy As String
    want = CalloutCharge("Refused", recs(1).Grade, USUAL_SHIFT_HOURS)        ' the person below the partial adds nothing
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, L_REF)), st.SessionId, vbTextCompare) = 0 Then
            nRows = nRows + 1
            got = got + NumOr0(v(i, L_CHARGED))
            If SafeStr(v(i, L_OUTCOME)) = "Accepted" Then accH = NumOr0(v(i, L_BASEHRS))
            If SafeStr(v(i, L_OUTCOME)) = "Left Message" Then belowH = NumOr0(v(i, L_BASEHRS)): belowWhy = SafeStr(v(i, L_REASON))
        End If
    Next i
    If nRows <> IIf(n >= 3, 3, 2) Then why = st.SessionId & " has " & nRows & " rows": Exit Function
    If Abs(got - want) > 0.0001 Then why = st.SessionId & " charged " & got & ", expected " & want & " (only the refusal above the partial)": Exit Function
    If Abs(accH - 8) > 0.0001 Then why = "the acceptance shows " & accH & " h, expected 8": Exit Function
    If n >= 3 Then
        If Abs(belowH) > 0.0001 Then why = "the person below the partial was offered " & belowH & " h, expected 0 (the shift closed)": Exit Function
        If InStr(1, belowWhy, "Not offered", vbTextCompare) = 0 Then why = "the row below the partial should say 'Not offered ...'; it says '" & belowWhy & "'": Exit Function
    End If
    ' v26.5.5 (item 9): the 16 h rule sees the acceptance as WORKED time on that day shift - for the night that
    ' follows, the accepted 8 h are the block before it and the refusal counts nothing
    RefreshCalCache
    Dim wb2 As Double, wb1 As Double, mx2 As Double
    wb2 = WorkedBlockBefore(recs(2).EmpName, recs(2).EmpShift, d, True)
    wb1 = WorkedBlockBefore(recs(1).EmpName, recs(1).EmpShift, d, True)
    If Abs(wb2 - 8) > 0.0001 Then why = "the 16 h rule sees " & wb2 & " h worked before the night for the person who accepted 8 h, expected 8": Exit Function
    If Abs(wb1) > 0.0001 Then why = "the 16 h rule sees " & wb1 & " h worked before the night for the person who REFUSED, expected 0": Exit Function
    mx2 = FatigueMaxFor(recs(2).EmpName, recs(2).EmpShift, d, "Nights")
    If Abs(mx2 - FatigueMaxHours(8, HoursToNextWork(recs(2).EmpName, recs(2).EmpShift, d, True))) > 0.0001 Then
        why = "FatigueMaxFor gives " & mx2 & " h for the night after an 8 h acceptance; the core disagrees": Exit Function
    End If
    If mx2 > 8.0001 Then why = "the night after an 8 h acceptance allows " & mx2 & " h, more than 16 - 8": Exit Function
    Dim recsN() As EmpRec, nN As Long, offN As String
    nN = LoadEligible(st.Unit, CrewFilter(CREW_ALL), recsN, 999, CALLOUT_INCLUDE_DAYS, d)
    If nN > 0 Then nN = FilterNotFree(recsN, nN, d, "Nights", offN)
    For i = 1 To nN
        If StrComp(recsN(i).EmpName, recs(2).EmpName, vbTextCompare) = 0 Then
            If Abs(recsN(i).MaxHours - mx2) > 0.0001 Then why = "the night list carries max " & recsN(i).MaxHours & " h for the acceptor, FatigueMaxFor says " & mx2: Exit Function
        End If
    Next i
    st.Note = st.SessionId & ", charged " & HoursText(got) & IIf(n >= 3, " (0 h offered below the partial)", "") & "; 16 h rule: max " & HoursText(mx2) & " h that night"
    SmkCallout = True
End Function

' ---- 5. time off + sick call ----
Private Function SmkTimeOff(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim d As Date: d = SmkWeekday(Date + 10)
    st.AbsId = SaveAbsence(st.Emp, st.Unit, "Vacation", d, d + 1, Empty, "", SMK_WHY, "", "", "", "AB", True)
    If st.AbsId = "" Then why = "SaveAbsence returned no ID": Exit Function
    Dim dirty0 As String: dirty0 = CalDirtyStamp()           ' v25.1: the mark is written on a protected sheet - prove it took
    MarkCalDirty
    If dirty0 <> "(no CalMap)" And CalDirtyStamp() = dirty0 Then why = "MarkCalDirty didn't write the calendar's dirty mark (CalMap Z7 still '" & dirty0 & "')": Exit Function
    Dim sd As Date: sd = SmkWeekday(Date + 20)
    st.SickId = SaveSickCall(st.Emp, sd, "Days", st.Unit, Empty, SMK_WHY)
    If st.SickId = "" Then why = "SaveSickCall returned no ID": Exit Function
    ' v24.3.1: one calendar line per sick call - the absence - whether or not Days is the person's scheduled shift
    Dim slots() As CalSlot, n As Long, i As Long, nAbs As Long, nMine As Long, who As String
    who = CalName(st.Emp)
    RefreshCalCache
    n = SlotsFor(sd, "Days", GroupOfUnit(st.Unit), slots)
    For i = 1 To n
        If StrComp(slots(i).Key, "A:" & st.SickId, vbTextCompare) = 0 Then nAbs = nAbs + 1
        If UCase$(slots(i).TypeCode) = "SIC" And Left$(slots(i).LeftText, Len(who)) = who Then nMine = nMine + 1
    Next i
    If nAbs <> 1 Or nMine <> 1 Then
        why = "the sick call drew " & nMine & " calendar line(s) for " & st.Emp & " on " & Format$(sd, "mmm d") & _
              " (" & nAbs & " from the absence); expected exactly 1"
        Exit Function
    End If
    st.Note = st.AbsId & ", " & st.SickId & " (1 calendar line)"
    SmkTimeOff = True
End Function

' the first Monday-to-Friday date on or after d
Private Function SmkWeekday(ByVal d As Date) As Date
    Do While Weekday(d, vbMonday) > 5
        d = d + 1
    Loop
    SmkWeekday = d
End Function

' ---- 6. shift switch ----
Private Function SmkSwitch(ByRef st As SmkState, ByRef why As String) As Boolean
    st.SwitchId = SaveSwitch(st.PostDate + 14, "Days", st.Unit, st.Peer, st.Emp, "", SMK_WHY)
    If st.SwitchId = "" Then why = "SaveSwitch returned no ID": Exit Function
    st.Note = st.SwitchId
    SmkSwitch = True
End Function

' ---- 7. crew change from a date (v26.5: through the Employee pop-up's Edit - column E + its crew row, as convention 5 says) ----
Private Function SmkCrew(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(st.Emp)
    If r = 0 Then why = st.Emp & " isn't on the roster": Exit Function
    Dim eff As Date: eff = Date + 7
    Dim homeUnit As String, crewTxt As String, phone As String, email As String, sen As String, codes() As String, report As String
    If Not PersonDetails(st.Emp, homeUnit, crewTxt, phone, email, sen, codes) Then why = "PersonDetails returned False": Exit Function
    If Not SaveEmployeeEdit(st.Emp, "", "A", eff, phone, email, sen, codes, report) Then why = "SaveEmployeeEdit (crew A from " & Format$(eff, "mmm d") & ") returned False": Exit Function
    If Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value)) <> "A" Then why = "column E didn't change to A": Exit Function
    If CrewOn(st.Emp, eff + 1) <> "A" Then why = "crew on " & Format$(eff + 1, "mmm d") & " is '" & CrewOn(st.Emp, eff + 1) & "', expected A": Exit Function
    If CrewOn(st.Emp, Date) <> "DAYS" Then why = "crew today is '" & CrewOn(st.Emp, Date) & "', expected DAYS": Exit Function
    st.Note = "DAYS -> A from " & Format$(eff, "mmm d")
    SmkCrew = True
End Function

' ---- 8. unit move ----
Private Function SmkMove(ByRef st As SmkState, ByRef why As String) As Boolean
    If st.Unit2 = "" Then st.Note = "SKIPPED: no second unit with people": SmkMove = True: Exit Function
    Dim pc As PeopleCalc, report As String
    If Not ComputeUnitMove(st.Emp, st.Unit2, Date, pc) Then why = "ComputeUnitMove: " & pc.Warnings: Exit Function
    If Not SaveUnitMove(st.Emp, st.Unit2, st.Code2, st.Unit, "", pc, report) Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(st.Emp)
    If StrComp(Trim$(SafeStr(ws.Cells(r, C_HOME).Value)), st.Unit2, vbTextCompare) <> 0 Then why = "Home Unit didn't change to " & st.Unit2: Exit Function
    If Trim$(SafeStr(ws.Cells(r, DeptCol(st.Unit)).Value)) <> "" Then why = "the old " & st.Unit & " code wasn't cleared": Exit Function
    ' v24.3.1: the move put them on the list of every open Unit2 posting for a shift from today that
    ' can still be signed up for and that they are free for - and on no other
    Dim nOpen As Long
    why = SmkListMismatch(st.Emp, st.Unit2, Date, nOpen)
    If why <> "" Then Exit Function
    st.Note = st.Unit & " -> " & st.Unit2 & ", " & HoursText(pc.AddHours) & " h added, " & nOpen & " list(s) checked"
    SmkMove = True
End Function

' ---- 9. a planned unit move (v24.3): back to the first unit, dated tomorrow, then cancelled ----
Private Function SmkPlanned(ByRef st As SmkState, ByRef why As String) As Boolean
    If st.Unit2 = "" Then st.Note = "SKIPPED: no second unit with people": SmkPlanned = True: Exit Function
    Dim report As String, id As String, eff As Date
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(st.Emp)
    eff = Date + 1
    ' v24.3.1: an open posting in Unit for a shift the person is free for, three weeks or more out,
    ' with at least one other person on its list. The plan must ADD the person to that list and
    ' the cancel must take the row OFF again - through a FILTERED list (DeleteListRow, the only
    ' ListRows.Delete in the codebase, clears the filter itself; a filtered delete would take a
    ' whole sheet row, and a posting's header with it). Both tables are counted before and after.
    Dim p As PostRec, lst() As PListRec, n As Long, dd As Long, d As Date, found As Boolean
    For dd = 21 To 60
        d = Date + dd
        If SmkFreeListed(st.Emp, st.Unit2, d, "Days", CrewFilter(CREW_ALL)) Then   ' free that day (they hold the Unit2 code since step 8)
            If SmkFreeCount(st.Unit, d, "Days") >= 1 Then found = True: Exit For
        End If
    Next dd
    If Not found Then why = "no day shift in the next 60 days has " & st.Emp & " free and someone free in " & st.Unit: Exit Function
    p.Unit = st.Unit
    p.ShiftDate = d
    p.ShiftLbl = "Days"
    p.Crew = CREW_ALL
    p.Reason = "Unit Operator Support"
    p.Hours = USUAL_SHIFT_HOURS
    p.Positions = 1
    p.ClosesOn = d - POSTING_CLOSE_DAYS
    If Not CreatePostingFrom(p, False, False) Then why = "the second posting couldn't be created": Exit Function
    st.PostId2 = p.ID
    Dim rowsWas As Long, postWas As Long
    rowsWas = SmkPostingRows(st.PostId2)
    postWas = SmkTableRows(PT_TABLE)
    If SmkOnPostingList(st.PostId2, st.Emp) Then why = st.Emp & " is on " & st.PostId2 & "'s list before the move was planned": Exit Function
    ' after step 8 the person is in Unit2; the plan moves them back to Unit
    If Not SavePlannedMove(st.Emp, st.Unit, st.Code, st.Unit2, "", eff, report) Then why = "the plan for " & Format$(eff, "mmm d") & " was refused: see the box before this one": Exit Function
    id = PlannedFor(st.Emp)
    If id = "" Then why = "SavePlannedMove said yes but no Planned row was found": GoTo Fail
    If Trim$(SafeStr(ws.Cells(r, DeptCol(st.Unit)).Value)) <> "" Then why = "the roster's " & st.Unit & " code changed today": GoTo Fail
    If StrComp(Trim$(SafeStr(ws.Cells(r, C_HOME).Value)), st.Unit2, vbTextCompare) <> 0 Then why = "the Home Unit changed today": GoTo Fail
    If SmkOnList(st.Emp, st.Unit, Date) Then why = "listed in " & st.Unit & " for TODAY, before the move's date": GoTo Fail
    If Not SmkOnList(st.Emp, st.Unit2, Date) Then why = "not listed in " & st.Unit2 & " for today": GoTo Fail
    If Not SmkOnList(st.Emp, st.Unit, eff) Then why = "not listed in " & st.Unit & " for the move's date": GoTo Fail
    If SmkOnList(st.Emp, st.Unit2, eff) Then why = "still listed in " & st.Unit2 & " for the move's date": GoTo Fail
    If Not SmkOnPostingList(st.PostId2, st.Emp) Then why = "the planned move didn't add " & st.Emp & " to " & st.PostId2 & "'s list": GoTo Fail
    If SmkPostingRows(st.PostId2) <> rowsWas + 1 Then why = st.PostId2 & " has " & SmkPostingRows(st.PostId2) & " list rows after the plan, expected " & (rowsWas + 1): GoTo Fail
    ' the filter goes on now: the row DELETE through a filtered list is the trap under test
    If Not SmkFilterList(True) Then why = "the filter on the posting list didn't take": GoTo Fail
    If Not CancelPlannedMove(id, SMK_WHY, report) Then why = "the plan's cancel was refused: see the box before this one": GoTo Fail
    If PlannedFor(st.Emp) <> "" Then why = "the move is still waiting after the cancel": GoTo Fail
    If SmkOnList(st.Emp, st.Unit, eff) Then why = "still listed in " & st.Unit & " for tomorrow after the cancel": GoTo Fail
    If Not SmkOnList(st.Emp, st.Unit2, eff) Then why = "not back on " & st.Unit2 & "'s list for tomorrow after the cancel": GoTo Fail
    If SmkOnPostingList(st.PostId2, st.Emp) Then why = st.Emp & " is still on " & st.PostId2 & "'s list after the cancel": GoTo Fail
    If SmkPostingRows(st.PostId2) <> rowsWas Then why = st.PostId2 & " has " & SmkPostingRows(st.PostId2) & " list rows after the cancel, expected " & rowsWas: GoTo Fail
    If SmkTableRows(PT_TABLE) <> postWas Then why = "the postings table has " & SmkTableRows(PT_TABLE) & " rows after the cancel; it had " & postWas & " (a filtered delete took a posting's row)": GoTo Fail
    SmkFilterList False
    ' v24.3.2: the words on the posting's calendar line can be changed (Change the text of this line):
    ' a leave word is refused and nothing is written; otherwise the line and the OPEN posting's Reason
    ' change together (SetLineText; SetReasonText goes the other way)
    Dim covId As String, txtRep As String
    Const SMK_TXT As String = "Unit Operator Support (smoke)"
    covId = CovIdOfSource(st.PostId2)
    If covId = "" Then why = st.PostId2 & " has no calendar line to change the text of": Exit Function
    If SetLineText(covId, "Unit Operator Support STD", txtRep) Then why = "a leave word in the line's new text was NOT refused": Exit Function
    If CovCellOf(covId, CV_NEED) = "Unit Operator Support STD" Then why = "the refused text was written to line " & covId: Exit Function
    If Not SetLineText(covId, SMK_TXT, txtRep) Then why = "the line's text couldn't be changed: " & txtRep: Exit Function
    If CovCellOf(covId, CV_NEED) <> SMK_TXT Then why = "line " & covId & " reads '" & CovCellOf(covId, CV_NEED) & "' after the text change": Exit Function
    n = LoadPosting(st.PostId2, p, lst)
    If p.Reason <> SMK_TXT Then why = st.PostId2 & "'s Reason is '" & p.Reason & "' after the text change": Exit Function
    ' and the other way (Change the Reason of posting): the line carrying the Reason's words follows
    Const SMK_TXT2 As String = "Unit Operator Support (smoke 2)"
    If SetReasonText(st.PostId2, "Unit Operator Support LOA", txtRep) Then why = "a leave word in the Reason was NOT refused": Exit Function
    If Not SetReasonText(st.PostId2, SMK_TXT2, txtRep) Then why = "the Reason couldn't be changed: " & txtRep: Exit Function
    If CovCellOf(covId, CV_NEED) <> SMK_TXT2 Then why = "line " & covId & " reads '" & CovCellOf(covId, CV_NEED) & "' after the Reason change": Exit Function
    n = LoadPosting(st.PostId2, p, lst)
    If p.Reason <> SMK_TXT2 Then why = st.PostId2 & "'s Reason is '" & p.Reason & "' after the Reason change": Exit Function
    ' the posting itself is cancelled (nobody is charged), so nothing of it stays open
    If Not SubmitCancel(p, lst, n) Then why = st.PostId2 & " couldn't be cancelled": Exit Function
    If PostingStatus(st.PostId2) <> "Cancelled" Then why = st.PostId2 & " is " & PostingStatus(st.PostId2) & " after the cancel": Exit Function
    ' v25 (item 4c): a second plan, with a CREW change and an open OLD-unit posting. The crew row dated
    ' the move's day must appear at the plan and go at the cancel (H7: never a floor under a later typed
    ' date); the old-unit posting's row must carry the "moves" mark at the plan and lose it at the cancel.
    ' v25.1.1: the move is dated the day AFTER the newest Crew History row (step 7's, a week out), not
    ' tomorrow: SavePlannedMove refuses a crew change dated before an existing later one (Graeme's Oct 16
    ' run: "crew already has a change dated ... after this move's date", a box and a plain False).
    Dim crewWas As String, crewTo As String, crewFound As Boolean, sinceWas As Date, sinceNow As Date, crewNow As String
    crewWas = NewestCrew(st.Emp, crewFound, sinceWas)
    crewTo = IIf(NormCrew(crewWas) = "A", "C", "A")
    Dim eff2 As Date: eff2 = Date + 1
    If crewFound Then
        If Int(sinceWas) >= eff2 Then eff2 = Int(sinceWas) + 1
    End If
    Dim p3 As PostRec, lst3() As PListRec, n3 As Long, d3 As Date, found3 As Boolean
    For dd = CLng(eff2 - Date) To 60                ' a Unit2 shift on or after the move's date they are listed for
        d3 = Date + dd
        If SmkFreeListed(st.Emp, st.Unit2, d3, "Days", CrewFilter(CREW_ALL)) Then found3 = True: Exit For
    Next dd
    If Not found3 Then why = "no " & st.Unit2 & " day shift from " & Format$(eff2, "mmm d") & " to " & Format$(Date + 60, "mmm d") & " lists " & st.Emp: Exit Function
    p3.Unit = st.Unit2: p3.ShiftDate = d3: p3.ShiftLbl = "Days": p3.Crew = CREW_ALL
    p3.Reason = "Unit Operator Support": p3.Hours = USUAL_SHIFT_HOURS: p3.Positions = 1: p3.ClosesOn = d3 - 1
    If Not CreatePostingFrom(p3, False, False) Then why = "the old-unit posting couldn't be created": Exit Function
    If Not SmkOnPostingList(p3.ID, st.Emp) Then why = st.Emp & " isn't on the old-unit posting " & p3.ID: GoTo Fail3
    If SmkExemptNote(p3.ID, st.Emp) <> "" Then why = st.Emp & " is already marked on " & p3.ID & " before the plan": GoTo Fail3
    If Not SavePlannedMove(st.Emp, st.Unit, st.Code, st.Unit2, crewTo, eff2, report) Then why = "the second plan (crew " & crewTo & " from " & Format$(eff2, "mmm d") & ") was refused: see the box before this one": GoTo Fail3
    id = PlannedFor(st.Emp)
    If id = "" Then why = "the second plan left no Planned row": GoTo Fail3
    crewNow = NewestCrew(st.Emp, crewFound, sinceNow)
    If NormCrew(crewNow) <> crewTo Or Int(sinceNow) <> Int(eff2) Then why = "the planned crew " & crewTo & " from " & Format$(eff2, "mmm d") & " isn't the newest Crew History row (" & crewNow & " from " & Format$(sinceNow, "mmm d") & ")": GoTo Fail3
    If Left$(SmkExemptNote(p3.ID, st.Emp), Len(EXEMPT_TAG) + 5) <> EXEMPT_TAG & "moves" Then why = "the old-unit posting " & p3.ID & " didn't get the 'moves' mark (" & SmkExemptNote(p3.ID, st.Emp) & ")": GoTo Fail3
    If Not CancelPlannedMove(id, SMK_WHY, report) Then why = "the second plan's cancel was refused: see the box before this one": GoTo Fail3
    crewNow = NewestCrew(st.Emp, crewFound, sinceNow)
    If NormCrew(crewNow) <> NormCrew(crewWas) Or Int(sinceNow) <> Int(sinceWas) Then why = "after the cancel the Crew History newest row is " & crewNow & " from " & Format$(sinceNow, "mmm d") & ", expected " & crewWas & " from " & Format$(sinceWas, "mmm d") & " (the planned row should be gone)": GoTo Fail3
    If SmkExemptNote(p3.ID, st.Emp) <> "" Then why = "the 'moves' mark is still on " & p3.ID & " after the cancel": GoTo Fail3
    n3 = LoadPosting(p3.ID, p3, lst3)
    If Not SubmitCancel(p3, lst3, n3) Then why = p3.ID & " couldn't be cancelled": Exit Function
    st.Note = id & ": " & st.Unit2 & " -> " & st.Unit & " from " & Format$(eff, "mmm d") & "; +1/-1 list row on " & st.PostId2 & _
              " (filtered); line " & covId & " text and the Reason changed each way, leave words refused; cancelled; " & _
              "crew " & crewWas & "->" & crewTo & " from " & Format$(eff2, "mmm d") & " row added and removed; 'moves' mark on " & p3.ID & " on and off; cancelled"
    SmkPlanned = True
    Exit Function
Fail3:
    On Error Resume Next
    If PlannedFor(st.Emp) <> "" Then CancelPlannedMove PlannedFor(st.Emp), SMK_WHY, report
    n3 = LoadPosting(p3.ID, p3, lst3)
    SubmitCancel p3, lst3, n3
    Exit Function
Fail:
    SmkFilterList False
End Function

' v24.3.1: posting-list helpers for the assertions above. They count rows and ask the
' roster's own list builders; no rule lives here.
Private Function SmkPostingTable(ByVal nm As String) As ListObject
    On Error Resume Next
    Set SmkPostingTable = ThisWorkbook.Worksheets(POSTINGS_SHEET).ListObjects(nm)
End Function

Private Function SmkTableRows(ByVal nm As String) As Long
    Dim lo As ListObject: Set lo = SmkPostingTable(nm)
    If lo Is Nothing Then Exit Function
    SmkTableRows = lo.ListRows.Count
End Function

' rows of the list table that belong to one posting
Private Function SmkPostingRows(ByVal id As String) As Long
    Dim lo As ListObject: Set lo = SmkPostingTable(PL_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, PL_ID))), id, vbTextCompare) = 0 Then SmkPostingRows = SmkPostingRows + 1
    Next i
End Function

' v25: the Result / Exempt note on one person's row of one posting ("" = none, or not on the list)
Private Function SmkExemptNote(ByVal id As String, ByVal emp As String) As String
    Dim p As PostRec, lst() As PListRec, n As Long, i As Long
    n = LoadPosting(id, p, lst)
    For i = 1 To n
        If StrComp(Trim$(lst(i).EmpName), emp, vbTextCompare) = 0 Then SmkExemptNote = lst(i).ExemptNote: Exit Function
    Next i
End Function

Private Function SmkOnPostingList(ByVal id As String, ByVal emp As String) As Boolean
    Dim lo As ListObject: Set lo = SmkPostingTable(PL_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, PL_ID))), id, vbTextCompare) = 0 Then
            If StrComp(Trim$(SafeStr(v(i, PL_EMP))), emp, vbTextCompare) = 0 Then SmkOnPostingList = True: Exit Function
        End If
    Next i
End Function

' Is the person on the unit's list for that date, shift and crew choice, and free for it? (the
' same calls that build a posting's list: LoadEligible by date with the posting's crew, then
' FilterNotFree - a crew-restricted posting lists only that crew's people)
Private Function SmkFreeListed(ByVal emp As String, ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String, _
                               ByVal crewSel As String) As Boolean
    Dim recs() As EmpRec, n As Long, i As Long, offTxt As String
    n = LoadEligible(unit, crewSel, recs, DeptPostingMaxGrade(unit), True, d)
    If n > 0 Then n = FilterNotFree(recs, n, d, shiftLbl, offTxt)
    For i = 1 To n
        If StrComp(recs(i).EmpName, emp, vbTextCompare) = 0 Then SmkFreeListed = True: Exit Function
    Next i
End Function

' how many people a posting in that unit for that shift would list
Private Function SmkFreeCount(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String) As Long
    Dim recs() As EmpRec, n As Long, offTxt As String
    n = LoadEligible(unit, CrewFilter(CREW_ALL), recs, DeptPostingMaxGrade(unit), True, d)
    If n > 0 Then n = FilterNotFree(recs, n, d, shiftLbl, offTxt)
    SmkFreeCount = n
End Function

' v25: ...of whom this many are chargeable (grade CHARGE_MAX_GRADE and below): a No takers needs one
Private Function SmkChargeableFree(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String) As Long
    Dim recs() As EmpRec, n As Long, i As Long, offTxt As String
    n = LoadEligible(unit, CrewFilter(CREW_ALL), recs, DeptPostingMaxGrade(unit), True, d)
    If n > 0 Then n = FilterNotFree(recs, n, d, shiftLbl, offTxt)
    For i = 1 To n
        If recs(i).Grade <= CHARGE_MAX_GRADE Then SmkChargeableFree = SmkChargeableFree + 1
    Next i
End Function

' For every Open posting in the unit for a shift from fromDate that can still be signed up for
' (closing day not passed): the person is on its list exactly when they are listed and free for
' that shift. "" when every posting agrees; otherwise the first disagreement. nOpen = how many
' such postings there are.
Private Function SmkListMismatch(ByVal emp As String, ByVal unit As String, ByVal fromDate As Date, ByRef nOpen As Long) As String
    Dim lo As ListObject: Set lo = SmkPostingTable(PT_TABLE)
    nOpen = 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, sd As Date, cl As Date, id As String, onList As Boolean, shouldBe As Boolean
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, PS_STATUS)) = "Open" And _
           StrComp(Trim$(SafeStr(v(i, PS_UNIT))), unit, vbTextCompare) = 0 And IsDate(v(i, PS_DATE)) Then
            sd = Int(CDate(v(i, PS_DATE)))
            cl = 0
            If IsDate(v(i, PS_CLOSES)) Then cl = Int(CDate(v(i, PS_CLOSES)))
            id = Trim$(SafeStr(v(i, PS_ID)))
            ' the same postings SyncPersonOnPostings touches: Status exactly "Open", at least one list row,
            ' a shift from fromDate that can still be signed up for
            If sd >= Int(fromDate) And sd >= Date And (cl = 0 Or Not PostingClosed(cl)) And SmkPostingRows(id) > 0 Then     ' v25: the one closing test
                nOpen = nOpen + 1
                onList = SmkOnPostingList(id, emp)
                shouldBe = SmkFreeListed(emp, unit, sd, SafeStr(v(i, PS_SHIFT)), CrewFilter(SafeStr(v(i, PS_CREW))))
                If onList <> shouldBe Then
                    SmkListMismatch = id & " (" & Format$(sd, "mmm d") & " " & SafeStr(v(i, PS_SHIFT)) & "): " & emp & _
                                      IIf(shouldBe, " should be on its list and isn't", " is on its list and shouldn't be")
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' a filter on the list table (Rank = 1), as a supervisor might leave one; False clears it.
' Returns True when the filter is really on (so the step cannot pass unfiltered by accident).
Private Function SmkFilterList(ByVal turnOn As Boolean) As Boolean
    Dim lr As ListObject: Set lr = SmkPostingTable(PL_TABLE)
    If lr Is Nothing Then Exit Function
    On Error Resume Next
    If turnOn Then
        lr.Range.AutoFilter Field:=PL_RANK, Criteria1:="1"
        SmkFilterList = lr.AutoFilter.FilterMode
    Else
        lr.AutoFilter.ShowAllData
        SmkFilterList = True
    End If
End Function

' Is the person on the unit's list (anyone with a code, any crew) for that date?
Private Function SmkOnList(ByVal emp As String, ByVal unit As String, ByVal d As Date) As Boolean
    Dim recs() As EmpRec, n As Long, i As Long
    n = LoadEligible(unit, "ALL", recs, 999, True, d)
    For i = 1 To n
        If StrComp(recs(i).EmpName, emp, vbTextCompare) = 0 Then
            SmkOnList = True
            Exit Function
        End If
    Next i
End Function

' ---- 11. rename ----
Private Function SmkRename(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim report As String
    If Not RenameEverywhere(st.Emp, SMK_EMP2, report) Then Exit Function
    If RosterRowOf(SMK_EMP2) = 0 Or RosterRowOf(SMK_EMP) > 0 Then why = "the roster name didn't change": Exit Function
    ' v26.5.5 (A8): the old name is gone from EVERY name column (the log, lists, absences, switches, the
    ' grade / crew / hours histories, planned moves, Worked Periods) and the new one is somewhere
    Dim leftIn As String
    If CountNameEverywhere(SMK_EMP, leftIn) > 0 Then why = "the old name is still in: " & leftIn: Exit Function
    If CountNameEverywhere(SMK_EMP2) = 0 Then why = "the new name is in no table (the rename wrote nothing)": Exit Function
    st.Emp = SMK_EMP2
    SmkRename = True
End Function

' ---- 12. publish to a scratch file, read it back ----
Private Function SmkPublish(ByRef st As SmkState, ByRef why As String) As Boolean
    If Not MemberTemplateExists() Then st.Note = "SKIPPED: no members' template beside this copy": SmkPublish = True: Exit Function
    Dim folder As String: folder = EnsureSlash(Environ$("TEMP")) & "RosterSmoke\"
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    If Not FolderExists(folder) Then why = "couldn't make " & folder: Exit Function
    gSmokeFolder = folder
    gSmokeSaved = ""
    PublishMembers
    Application.ScreenUpdating = False             ' the publish switches it back on; restored by SmokeTest at its end
    If gSmokeSaved = "" Then why = "the publish didn't write its file (a message above says why)": Exit Function

    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim secWas As Long: secWas = Application.AutomationSecurity
    Dim wb As Workbook, nLog As Long, nRos As Long, nAtt As Long, em As String
    On Error GoTo ReadFailed
    Application.EnableEvents = False
    Application.AutomationSecurity = 3              ' msoAutomationSecurityForceDisable: its macros stay off
    Set wb = Workbooks.Open(Filename:=gSmokeSaved, ReadOnly:=True, UpdateLinks:=0)
    Application.AutomationSecurity = secWas
    nLog = SmkRows(wb, DS_LOG)
    nRos = SmkRows(wb, DS_ROSTER)
    nAtt = SmkRows(wb, DS_ATT)
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Application.EnableEvents = evWas
    ThisWorkbook.Activate
    On Error Resume Next
    Kill gSmokeSaved
    On Error GoTo 0
    If nRos < 1 Then why = "the published roster table is empty": Exit Function
    If nLog < 1 Then why = "the published log is empty (My history would show nothing)": Exit Function
    If nAtt < 1 Then why = "the published sick-time counts are empty": Exit Function
    st.Note = nRos & " people, " & nLog & " log rows, " & nAtt & " sick-time rows"
    SmkPublish = True
    Exit Function
ReadFailed:
    em = Err.Description
    On Error Resume Next
    Application.AutomationSecurity = secWas
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.EnableEvents = evWas
    ThisWorkbook.Activate
    why = "couldn't read the scratch file back: " & em
End Function

' data rows of a published table (its first row is the header)
Private Function SmkRows(ByVal wb As Workbook, ByVal sheetName As String) As Long
    Dim ws As Worksheet: Set ws = wb.Worksheets(sheetName)
    Dim last As Long: last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last > 1 Then SmkRows = last - 1
End Function

' ---- 13. no takers, award out of dept, void, reopen (v25) ----
' Two more postings in Unit (the test person holds no code there since step 8, so they are OUT of
' department for them). Posting 3: No takers (everyone on the list charged), its rows voided, then
' reopened (ReopenPostingById) and left Open for step 16. Posting 4: awarded out of department to
' the test person (an Awarded row under their name, voided in step 17).
Private Function SmkOutcomes(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim p As PostRec, lst() As PListRec, n As Long, dd As Long, d As Date, found As Boolean, report As String, calN As Long
    If RosterRowOf(st.Emp) = 0 Then why = st.Emp & " isn't on the roster": Exit Function
    For dd = 2 To 60
        d = Date + dd
        If SmkChargeableFree(st.Unit, d, "Days") >= 1 Then found = True: Exit For
    Next dd
    If Not found Then why = "no day shift in the next 60 days has anyone chargeable free in " & st.Unit: Exit Function
    p.Unit = st.Unit: p.ShiftDate = d: p.ShiftLbl = "Days": p.Crew = CREW_ALL
    p.Reason = "Unit Operator Support": p.Hours = USUAL_SHIFT_HOURS: p.Positions = 1: p.ClosesOn = d - 1
    If Not CreatePostingFrom(p, False, False) Then why = "posting 3 couldn't be created": Exit Function
    st.PostId3 = p.ID
    n = LoadPosting(st.PostId3, p, lst)
    If n < 1 Then why = st.PostId3 & " has nobody on its list": Exit Function
    If Not SubmitNoTakers(p, lst, n) Then why = st.PostId3 & " couldn't be closed as no takers": Exit Function
    If PostingStatus(st.PostId3) <> "No takers" Then why = st.PostId3 & " is " & PostingStatus(st.PostId3) & " after no takers": Exit Function
    If Not HasActiveLogRows(st.PostId3) Then why = "no takers wrote no OT Log rows for " & st.PostId3: Exit Function
    If VoidRef(st.PostId3, SMK_WHY, calN) = 0 Then why = "VoidRef voided nothing for " & st.PostId3: Exit Function
    If HasActiveLogRows(st.PostId3) Then why = st.PostId3 & " still has active OT Log rows after the void": Exit Function
    If Not ReopenPostingById(st.PostId3, report) Then why = st.PostId3 & " couldn't be reopened: " & report: Exit Function
    If PostingStatus(st.PostId3) <> "Open" Then why = st.PostId3 & " is " & PostingStatus(st.PostId3) & " after the reopen": Exit Function
    ' posting 4: out of department to the test person (only when step 8 moved them out of st.Unit)
    If st.Unit2 = "" Then
        st.Note = st.PostId3 & " no takers -> voided -> open again; the out-of-department award skipped (no second unit with people)"
        SmkOutcomes = True
        Exit Function
    End If
    Dim p4 As PostRec, lst4() As PListRec, n4 As Long
    p4.Unit = st.Unit: p4.ShiftDate = d: p4.ShiftLbl = "Days": p4.Crew = CREW_ALL
    p4.Reason = "Unit Operator Support": p4.Hours = USUAL_SHIFT_HOURS: p4.Positions = 1: p4.ClosesOn = d - 1
    If Not CreatePostingFrom(p4, False, False) Then why = "posting 4 couldn't be created": Exit Function
    st.PostId4 = p4.ID
    n4 = LoadPosting(st.PostId4, p4, lst4)
    If SmkOnPostingList(st.PostId4, st.Emp) Then why = st.Emp & " is on " & st.PostId4 & "'s list, so isn't out of department": Exit Function
    If Not SubmitAwardOutside(p4, lst4, n4, st.Emp, 0) Then why = st.PostId4 & " couldn't be awarded out of department": Exit Function
    If PostingStatus(st.PostId4) = "Open" Then why = st.PostId4 & " is still Open after the out-of-department award": Exit Function
    If Not HasActiveLogRows(st.PostId4) Then why = "the out-of-department award wrote no OT Log rows": Exit Function
    st.Note = st.PostId3 & " no takers -> voided -> open again; " & st.PostId4 & " awarded out of dept to " & st.Emp
    SmkOutcomes = True
End Function

' ---- 14. leave going and returning (v25) ----
' A leave booked from 30 days ago, returned today through the leave rule; the LV- row is logged even
' when the top-up is 0 h (L5; no peers or no history give 0 h with a warning, not a refusal), and the
' rule then refuses a second return. The one refusal that is this copy's state, not a fault - a
' return after the roster's last period (the year wasn't rolled over) - is recorded as skipped.
Private Function SmkLeaveReturn(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim lc As LeaveCalc, grp As String, unit As String, avg As Double
    If RosterRowOf(st.Emp) = 0 Then why = st.Emp & " isn't on the roster": Exit Function
    unit = IIf(st.Unit2 <> "", st.Unit2, st.Unit)
    st.LeaveAbsId = SaveAbsence(st.Emp, unit, "LOA", Date - 30, Empty, Empty, "", SMK_WHY, "", "", "", "AB", True)
    If st.LeaveAbsId = "" Then why = "the leave couldn't be booked": Exit Function
    If Not ComputeLeaveReturn(st.LeaveAbsId, Date, lc) Then
        If InStr(1, lc.Warnings, "after the last day", vbTextCompare) > 0 Then
            st.Note = "SKIPPED: " & Left$(lc.Warnings, 160)
            SmkLeaveReturn = True
        Else
            why = "the leave rule refused the return: " & Left$(lc.Warnings, 200)
        End If
        Exit Function
    End If
    avg = lc.Average
    grp = SaveLeaveReturn(lc, lc.Average, "")
    If grp = "" Then why = "the leave return wasn't saved": Exit Function
    st.LeaveRef = grp
    If Not HasActiveLogRows(grp) Then why = "the leave return logged no row (" & grp & ")": Exit Function
    If ComputeLeaveReturn(st.LeaveAbsId, Date, lc) Then why = "a second return on " & st.LeaveAbsId & " wasn't refused": Exit Function
    If InStr(1, lc.Warnings, "already been topped up", vbTextCompare) = 0 Then why = "the second return was refused for the wrong reason: " & lc.Warnings: Exit Function
    st.Note = grp & ": " & Format$(avg, "0.0") & " h logged; a second return refused"
    SmkLeaveReturn = True
End Function

' ---- 15. Kronos headings on a made-up sheet (v25, D-I4) ----
Private Function SmkKronosHeadings(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim src(1 To 4, 1 To 4) As Variant, nameCol As Long, totCol As Long, hdrName As Long, hdrTot As Long, w As String
    src(1, 1) = "Hours report": src(2, 1) = "Payroll #": src(2, 2) = "Employee Name": src(2, 3) = "Regular": src(2, 4) = "Grand Total"
    src(3, 1) = "12345678": src(3, 2) = "Test, Someone": src(3, 3) = 80: src(3, 4) = 80
    w = KronosHeadingProblem(src, nameCol, totCol, hdrName, hdrTot)
    If w <> "" Then why = "a good sheet was refused: " & w: Exit Function
    If nameCol <> 2 Or totCol <> 4 Or hdrName <> 2 Then why = "the headings were found in the wrong place (" & nameCol & ", " & totCol & ", row " & hdrName & ")": Exit Function
    src(2, 4) = "Total"
    w = KronosHeadingProblem(src, nameCol, totCol, hdrName, hdrTot)
    If InStr(w, "Grand Total") = 0 Then why = "a sheet without 'Grand Total' wasn't refused (" & w & ")": Exit Function
    src(2, 4) = "Grand Total": src(2, 2) = "Name": src(4, 2) = "Employee Name"
    w = KronosHeadingProblem(src, nameCol, totCol, hdrName, hdrTot)
    If InStr(w, "same heading row") = 0 Then why = "headings on two rows weren't refused (" & w & ")": Exit Function
    st.Note = "found by heading; 'Total' for 'Grand Total' refused; headings on two rows refused"
    SmkKronosHeadings = True
End Function

' ---- 16. a planted leave word stops the publish (v25) ----
' The posting Reason first, then a calendar line's covered-by; each publish must write nothing and
' must have refused for the planted word (gSmokeRefusal). The words are taken back and posting 3 is
' cancelled (step 17 cancels it when this step is skipped; a FAILED step stops the run, saving blocked).
Private Function SmkPlantedWords(ByRef st As SmkState, ByRef why As String) As Boolean
    If Not MemberTemplateExists() Then st.Note = "SKIPPED: no members' template beside this copy": SmkPlantedWords = True: Exit Function
    If st.PostId3 = "" Then st.Note = "SKIPPED: no open posting from step 13": SmkPlantedWords = True: Exit Function
    If PostingStatus(st.PostId3) <> "Open" Then st.Note = "SKIPPED: " & st.PostId3 & " isn't open": SmkPlantedWords = True: Exit Function
    Dim folder As String: folder = EnsureSlash(Environ$("TEMP")) & "RosterSmoke\"
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    If Not FolderExists(folder) Then why = "couldn't make " & folder: Exit Function
    Dim p As PostRec, lst() As PListRec, n As Long, reasonWas As String
    n = LoadPosting(st.PostId3, p, lst)
    reasonWas = p.Reason
    If Not SetPostingReason(st.PostId3, "Covering J. Smith STD") Then why = "the Reason couldn't be planted": Exit Function
    gSmokeFolder = folder: gSmokeSaved = "": gSmokeRefusal = ""
    PublishMembers
    Application.ScreenUpdating = False             ' the publish switches it back on; restored by SmokeTest at its end
    SetPostingReason st.PostId3, reasonWas
    If gSmokeSaved <> "" Then why = "the publish wrote its file although a posting Reason said STD": GoTo Fail
    If InStr(1, gSmokeRefusal, "STD", vbBinaryCompare) = 0 Or InStr(1, gSmokeRefusal, "Reason", vbTextCompare) = 0 Then
        why = "the publish didn't refuse for the planted Reason: " & Left$(gSmokeRefusal, 200): GoTo Fail
    End If
    If UpdateCoverageBySource(st.PostId3, "NCN STD") = 0 Then why = st.PostId3 & " has no calendar line to plant on": Exit Function
    gSmokeSaved = "": gSmokeRefusal = ""
    PublishMembers
    Application.ScreenUpdating = False             ' restored by SmokeTest at its end
    UpdateCoverageBySource st.PostId3, "POSTED " & st.PostId3
    If gSmokeSaved <> "" Then why = "the publish wrote its file although a calendar line said STD": GoTo Fail
    If InStr(1, gSmokeRefusal, "STD", vbBinaryCompare) = 0 Then why = "the publish didn't refuse for the planted calendar text: " & Left$(gSmokeRefusal, 200): GoTo Fail
    n = LoadPosting(st.PostId3, p, lst)
    If Not SubmitCancel(p, lst, n) Then why = st.PostId3 & " couldn't be cancelled": Exit Function
    st.Note = "Reason and calendar line each refused; " & st.PostId3 & " cancelled"
    SmkPlantedWords = True
    Exit Function
Fail:
    On Error Resume Next
    Kill gSmokeSaved
End Function

' ---- 17. void everything ----
Private Function SmkVoid(ByRef st As SmkState, ByRef why As String) As Boolean
    ' every Ref ID that is the session, the posting, or carries a row under either test name
    Dim lo As ListObject: Set lo = LogTable()
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim refs As Object: Set refs = CreateObject("Scripting.Dictionary")
    refs.CompareMode = 1
    Dim i As Long, ref As String, emp As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then
            ref = Trim$(SafeStr(v(i, L_REF))): emp = Trim$(SafeStr(v(i, L_EMP)))
            If ref <> "" Then
                If StrComp(ref, st.SessionId, vbTextCompare) = 0 Or StrComp(ref, st.PostId, vbTextCompare) = 0 Or _
                   StrComp(ref, st.PostId3, vbTextCompare) = 0 Or StrComp(ref, st.PostId4, vbTextCompare) = 0 Or _
                   StrComp(ref, st.LeaveRef, vbTextCompare) = 0 Or _
                   StrComp(emp, SMK_EMP, vbTextCompare) = 0 Or StrComp(emp, SMK_EMP2, vbTextCompare) = 0 Then refs(ref) = True
            End If
        End If
    Next i
    Dim kk As Variant, nRows As Long, calN As Long, n As Long
    For Each kk In refs.Keys
        n = VoidRef(CStr(kk), SMK_WHY, calN)
        If n = 0 Then why = "VoidRef voided nothing for " & CStr(kk): Exit Function
        nRows = nRows + n
    Next kk
    If st.SwitchId <> "" Then
        If Not VoidSwitch(st.SwitchId, SMK_WHY) Then why = "VoidSwitch " & st.SwitchId & " returned False": Exit Function
    End If
    If st.AbsId <> "" Then
        If Not VoidAbsence(st.AbsId, SMK_WHY) Then why = "VoidAbsence " & st.AbsId & " returned False": Exit Function
    End If
    If st.SickId <> "" Then
        If Not VoidAbsence(st.SickId, SMK_WHY) Then why = "VoidAbsence " & st.SickId & " returned False": Exit Function
    End If
    If st.LeaveAbsId <> "" Then                      ' v25: step 14's leave (its LV- row went with the refs above)
        If Not VoidAbsence(st.LeaveAbsId, SMK_WHY) Then why = "VoidAbsence " & st.LeaveAbsId & " returned False": Exit Function
    End If
    ' v25: posting 3 is still Open when step 16 was skipped (no template): cancelled here so nothing of the run stays Open
    Dim p As PostRec, lst() As PListRec
    If st.PostId3 <> "" Then
        If PostingStatus(st.PostId3) = "Open" Then
            n = LoadPosting(st.PostId3, p, lst)
            If Not SubmitCancel(p, lst, n) Then why = st.PostId3 & " couldn't be cancelled": Exit Function
        End If
    End If
    If st.PostId4 <> "" Then
        If PostingStatus(st.PostId4) = "Open" Then
            n = LoadPosting(st.PostId4, p, lst)
            If Not SubmitCancel(p, lst, n) Then why = st.PostId4 & " couldn't be cancelled": Exit Function
        End If
    End If
    MarkCalDirty
    st.Note = refs.Count & " Ref IDs, " & nRows & " rows"
    SmkVoid = True
End Function

' ---- 18. remove ----
Private Function SmkRemove(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim report As String
    If Not RemoveEmployee(st.Emp, Date, report) Then Exit Function
    If RosterRowOf(st.Emp) > 0 Then why = st.Emp & " is still on the roster": Exit Function
    If Not IsFormer(st.Emp) Then why = st.Emp & " isn't on the Former sheet": Exit Function
    SmkRemove = True
End Function

' ---- 19. integrity ----
Private Function SmkIntegrity(ByRef st As SmkState, ByVal nErr0 As Long, ByRef why As String) As Boolean
    Dim nErr As Long, nWarn As Long
    Application.Calculate
    RunIntegrity nErr, nWarn
    If nErr > nErr0 Then
        why = "IntegrityCheck has " & nErr & " PROBLEM line(s); it had " & nErr0 & " before the run. Run IntegrityCheck in this copy to see them."
        Exit Function
    End If
    st.Note = nErr & " PROBLEM (" & nErr0 & " before), " & nWarn & " Check"
    SmkIntegrity = True
End Function

' (v25.5: the v25.2 compare steps 19-20 - the Calendar and Needs Attention drawn both ways - went with the
'  legacy writers; the compares passed on the real roster Oct 17, 2026)

' ---- 10. v26.5: the Employee pop-up's Edit, through its one write path: a new phone, a new seniority date and a code
'          added in the FIRST unit (cleared by step 8's move) then cleared again; the cells show it, the totals come back
'          (SaveEmployeeEdit's own check), the Grade Log grew by two ----
Private Function SmkEdit(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(st.Emp)
    If r = 0 Then why = st.Emp & " isn't on the roster": Exit Function
    Dim homeUnit As String, crewTxt As String, phone As String, email As String, sen As String, codes() As String
    If Not PersonDetails(st.Emp, homeUnit, crewTxt, phone, email, sen, codes) Then why = "PersonDetails returned False": Exit Function
    Dim glBefore As Long
    If Not GradeLogTable(False) Is Nothing Then
        If Not GradeLogTable(False).DataBodyRange Is Nothing Then glBefore = GradeLogTable(False).ListRows.Count
    End If
    ' the code goes into the FIRST unit's column (cleared by step 8's move) and is cleared again by a second save, so the
    ' later steps find the person as they expect (out of department in Unit since step 8)
    Dim c1 As Long, addedCode As Boolean
    c1 = DeptCol(st.Unit) - UnitColFirst()
    If codes(c1) = "" Then codes(c1) = "Z9": addedCode = True
    Dim report As String, senDate As Date: senDate = DateSerial(2011, 3, 3)
    If Not SaveEmployeeEdit(st.Emp, "", "", 0, "ext 0100", email, senDate, codes, report) Then why = "SaveEmployeeEdit returned False (a message above says why)": Exit Function
    If Trim$(SafeStr(ws.Cells(r, C_PHONE).Value)) <> "ext 0100" Then why = "the phone didn't change": Exit Function
    If Not IsDate(ws.Cells(r, C_SENIORITY).Value) Then why = "the seniority date isn't a date": Exit Function
    If Int(CDate(ws.Cells(r, C_SENIORITY).Value)) <> senDate Then why = "the seniority date didn't change": Exit Function
    If addedCode Then
        If Trim$(SafeStr(ws.Cells(r, DeptCol(st.Unit)).Value)) <> "Z9" Then why = "the " & st.Unit & " code wasn't written": Exit Function
        If Not GradeLogTable(False) Is Nothing Then
            If GradeLogTable(False).ListRows.Count <= glBefore Then why = "the Grade Log didn't get a row for the new code": Exit Function
        End If
        codes(c1) = ""                               ' and back off again: the person holds no Unit code for the steps that follow
        If Not SaveEmployeeEdit(st.Emp, "", "", 0, "ext 0100", email, senDate, codes, report) Then why = "the second save (clearing the code) returned False": Exit Function
        If Trim$(SafeStr(ws.Cells(r, DeptCol(st.Unit)).Value)) <> "" Then why = "the " & st.Unit & " code wasn't cleared again": Exit Function
    End If
    st.Note = "phone, seniority" & IIf(addedCode, ", a " & st.Unit & " code added then cleared (Grade Log + 2)", "") & " - totals unchanged"
    SmkEdit = True
End Function

' ---- 20. the Desk (v25.3; step 21 until v25.5, 19 until v26.5): its four figures come from the roster's own functions. A vacation
' for the peer on a day shift inside the window adds at least one open line; a posting closing today adds one ready to award
' once the closing moment has passed; both are taken back and the figures return; ViewModeFor on four logins (v25.5:
' + a named supervisor).
Private Function SmkDesk(ByRef st As SmkState, ByRef why As String) As Boolean
    If ViewModeFor("boss", "ann,boss", "boss", "sue") <> "desk" Or ViewModeFor("ann", "ann,boss", "boss", "sue") <> "admin" Or _
       ViewModeFor("boss", "ann", "boss", "sue") <> "user" Or ViewModeFor("sue", "ann,boss", "boss", "sue") <> "supervisor" Then _
       why = "ViewModeFor gives the wrong view for a made-up login": Exit Function
    Dim c0 As Long, r0 As Long, o0 As Long, c1 As Long, r1 As Long, o1 As Long, c2 As Long, r2 As Long, o2 As Long, pubAt As Date
    DeskCounts c0, r0, o0, pubAt
    Dim d As Date, dd As Long, found As Boolean
    For dd = 1 To DESK_AHEAD_DAYS - 1                ' a day inside the tile's window the peer works days on
        d = Date + dd
        If EmpShiftOn(st.Peer, d) = "Days" Then found = True: Exit For
    Next dd
    If Not found Then st.Note = "SKIPPED: " & st.Peer & " has no day shift in the next " & DESK_AHEAD_DAYS & " days": SmkDesk = True: Exit Function
    Dim absId As String: absId = SaveAbsence(st.Peer, st.Unit, "Vacation", d, d, Empty, "", SMK_WHY, "", "", "", "AB", True)
    If absId = "" Then why = "SaveAbsence returned no ID": Exit Function
    MarkCalDirty
    Dim p As PostRec, lst() As PListRec, n As Long, readyAdd As Long, nearClose As Boolean
    ' the shift step 2 proved has people free (21+ days out: outside the cover tile's window); it closes today
    p.Unit = st.Unit: p.ShiftDate = st.PostDate: p.ShiftLbl = st.PostShift: p.Crew = CREW_ALL
    p.Reason = "Unit Operator Support": p.Hours = USUAL_SHIFT_HOURS: p.Positions = 1: p.ClosesOn = Date
    If Not CreatePostingFrom(p, False, False) Then why = "a posting closing today couldn't be created": GoTo Undo
    readyAdd = IIf(PostingClosed(Date), 1, 0)       ' closed from the closing time on the closing date
    ' a run crossing the closing moment would move every posting closing today: the ready compare is skipped then (review)
    nearClose = (Abs(Now - (Date + TimeValue(POSTING_CLOSE_TIME))) < 2 / 1440)
    DeskCounts c1, r1, o1, pubAt
    If c1 < c0 + 1 Then why = "shifts needing cover read " & c1 & " after a vacation in the window; expected at least " & (c0 + 1): GoTo Undo
    If Not nearClose Then
        If r1 <> r0 + readyAdd Then why = "postings ready to award read " & r1 & " with a posting closing today; expected " & (r0 + readyAdd): GoTo Undo
    End If
    n = LoadPosting(p.ID, p, lst)
    If Not SubmitCancel(p, lst, n) Then why = p.ID & " couldn't be cancelled": Exit Function
    If Not VoidAbsence(absId, SMK_WHY) Then why = "VoidAbsence " & absId & " returned False": Exit Function
    MarkCalDirty
    DeskCounts c2, r2, o2, pubAt
    If c2 <> c0 Or (r2 <> r0 And Not nearClose) Or o2 <> o0 Then why = "after the cancel and the void the figures read " & c2 & " / " & r2 & " / " & o2 & "; expected " & c0 & " / " & r0 & " / " & o0: Exit Function
    st.Note = "cover " & c0 & " -> " & c1 & " -> " & c2 & ", ready " & r0 & " -> " & r1 & " -> " & r2 & IIf(nearClose, " (not compared: the closing moment)", "") & ", owed " & o0 & "; ViewModeFor desk / admin / user / supervisor"
    SmkDesk = True
    Exit Function
Undo:
    On Error Resume Next
    If p.ID <> "" Then n = LoadPosting(p.ID, p, lst): SubmitCancel p, lst, n
    VoidAbsence absId, SMK_WHY
    MarkCalDirty
End Function
