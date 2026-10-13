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
'    changes the test person's crew from a date, moves them to another
'    unit, plans a unit move for tomorrow and cancels it (v24.3: the
'    lists must follow the date), renames them, publishes to a scratch
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

Public Const SMK_VERSION As String = "Smoke test v24.3.1 (Oct 14)"

Private Const SMK_EMP As String = "ZZ Test, Person"
Private Const SMK_EMP2 As String = "ZZ Test, Renamed"
Private Const SMK_WHY As String = "SmokeTest"
Private Const SMK_STEPS As Long = 14

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
              "then checks every total is unchanged. About a minute. Nothing is saved, and afterwards this copy can only be closed.", _
              vbOKCancel + vbQuestion, TTL) <> vbOK Then Exit Sub

    Dim st As SmkState, k As Long, ok As Boolean, why As String, rep As String, failedAt As Long
    Dim repShort As String, repNotes As String     ' v24.3.1: the step lines without their notes (the PASSED box), and the notes alone (the Error Log line)
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
            repShort = repShort & Format$(k, "00") & "  " & IIf(Left$(st.Note, 7) = "SKIPPED", "skip", "OK  ") & "  " & SmkName(k) & vbNewLine
            If st.Note <> "" Then repNotes = repNotes & k & ": " & st.Note & "; "
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
               "Send me this box (a photo is fine)."
    ElseIf diff <> "" Then
        head = "SMOKE TEST FAILED: every step ran, but the totals did not come back:" & vbNewLine & diff & "Send me this box."
    Else
        head = "SMOKE TEST PASSED: all steps OK, totals unchanged."
        rep = repShort                             ' 14 plain lines fit in one box; the notes are on the Error Log line
    End If
    If Len(head) > 420 Then head = Left$(head, 400) & " ..." & vbNewLine
    If Len(rep) > 880 - Len(head) Then rep = Left$(rep, 860 - Len(head)) & " ..." & vbNewLine     ' MsgBox holds about 1,000 characters
    MsgBox head & vbNewLine & vbNewLine & rep & vbNewLine & _
           "Saving is blocked in this copy. Close it; nothing is kept.", _
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
           "Saving is blocked in this copy. Close it; nothing is kept. Send me this box.", vbCritical, TTL
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
        Case 9: SmkName = "plan a unit move for tomorrow, cancel it through a filtered list, cancel that posting"
        Case 10: SmkName = "rename"
        Case 11: SmkName = "publish to a scratch file and read it back"
        Case 12: SmkName = "void everything it logged"
        Case 13: SmkName = "remove the test employee"
        Case 14: SmkName = "integrity check"
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
        Case 10: SmkStep = SmkRename(st, why)
        Case 11: SmkStep = SmkPublish(st, why)
        Case 12: SmkStep = SmkVoid(st, why)
        Case 13: SmkStep = SmkRemove(st, why)
        Case 14: SmkStep = SmkIntegrity(st, nErr0, why)
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
    If Not SubmitAward(p, lst, n) Then Exit Function
    If PostingStatus(st.PostId) = "Open" Then why = st.PostId & " is still Open after the award": Exit Function
    If Not HasActiveLogRows(st.PostId) Then why = "the award wrote no OT Log rows": Exit Function
    st.Note = "awarded #" & iFull & " of " & n
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
        If n >= 2 Then Exit For
    Next dd
    If n < 2 Then why = "no day shift in the five weeks after the posting has two people free in " & st.Unit: Exit Function
    ReDim taken(1 To n)
    recs(1).Outcome = "Refused"
    recs(2).Outcome = "Accepted": taken(2) = 8                  ' takes 8 of the 12 on offer
    If n >= 3 Then recs(3).Outcome = "Left Message"             ' offered the 4 that are left
    If Not SubmitCalloutSession(st.Unit, "Days", CREW_ALL, d, USUAL_SHIFT_HOURS, recs, taken, n, st.SessionId, summary) Then Exit Function

    ' the rows must carry the rule: the first refusal on 12 h, the acceptance 8 h, the third on the 4 h left
    Dim lo As ListObject: Set lo = LogTable()
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim want As Double, got As Double, nRows As Long, accH As Double
    want = CalloutCharge("Refused", recs(1).Grade, USUAL_SHIFT_HOURS)
    If n >= 3 Then want = want + CalloutCharge("Left Message", recs(3).Grade, USUAL_SHIFT_HOURS - 8)
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, L_REF)), st.SessionId, vbTextCompare) = 0 Then
            nRows = nRows + 1
            got = got + NumOr0(v(i, L_CHARGED))
            If SafeStr(v(i, L_OUTCOME)) = "Accepted" Then accH = NumOr0(v(i, L_BASEHRS))
        End If
    Next i
    If nRows <> IIf(n >= 3, 3, 2) Then why = st.SessionId & " has " & nRows & " rows": Exit Function
    If Abs(got - want) > 0.0001 Then why = st.SessionId & " charged " & got & ", expected " & want: Exit Function
    If Abs(accH - 8) > 0.0001 Then why = "the acceptance shows " & accH & " h, expected 8": Exit Function
    st.Note = st.SessionId & ", charged " & HoursText(got)
    SmkCallout = True
End Function

' ---- 5. time off + sick call ----
Private Function SmkTimeOff(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim d As Date: d = SmkWeekday(Date + 10)
    st.AbsId = SaveAbsence(st.Emp, st.Unit, "Vacation", d, d + 1, Empty, "", SMK_WHY, "", "", "", "AB", True)
    If st.AbsId = "" Then why = "SaveAbsence returned no ID": Exit Function
    MarkCalDirty
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

' ---- 7. crew change from a date (column E + its crew row, as convention 5 says) ----
Private Function SmkCrew(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(st.Emp)
    If r = 0 Then why = st.Emp & " isn't on the roster": Exit Function
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim eff As Date: eff = Date + 7
    Application.EnableEvents = False
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Cells(r, C_SHIFT).Value = "A"
    ProtectSheet ws
    Application.EnableEvents = evWas
    If Not AddCrewRow(st.Emp, "A", eff) Then why = "AddCrewRow returned False": Exit Function
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
    If Not SavePlannedMove(st.Emp, st.Unit, st.Code, st.Unit2, "", eff, report) Then Exit Function
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
    If Not CancelPlannedMove(id, SMK_WHY, report) Then GoTo Fail
    If PlannedFor(st.Emp) <> "" Then why = "the move is still waiting after the cancel": GoTo Fail
    If SmkOnList(st.Emp, st.Unit, eff) Then why = "still listed in " & st.Unit & " for tomorrow after the cancel": GoTo Fail
    If Not SmkOnList(st.Emp, st.Unit2, eff) Then why = "not back on " & st.Unit2 & "'s list for tomorrow after the cancel": GoTo Fail
    If SmkOnPostingList(st.PostId2, st.Emp) Then why = st.Emp & " is still on " & st.PostId2 & "'s list after the cancel": GoTo Fail
    If SmkPostingRows(st.PostId2) <> rowsWas Then why = st.PostId2 & " has " & SmkPostingRows(st.PostId2) & " list rows after the cancel, expected " & rowsWas: GoTo Fail
    If SmkTableRows(PT_TABLE) <> postWas Then why = "the postings table has " & SmkTableRows(PT_TABLE) & " rows after the cancel; it had " & postWas & " (a filtered delete took a posting's row)": GoTo Fail
    SmkFilterList False
    ' the posting itself is cancelled (nobody is charged), so nothing of it stays open
    n = LoadPosting(st.PostId2, p, lst)
    If Not SubmitCancel(p, lst, n) Then why = st.PostId2 & " couldn't be cancelled": Exit Function
    If PostingStatus(st.PostId2) <> "Cancelled" Then why = st.PostId2 & " is " & PostingStatus(st.PostId2) & " after the cancel": Exit Function
    st.Note = id & ": " & st.Unit2 & " -> " & st.Unit & " from " & Format$(eff, "mmm d") & "; +1/-1 list row on " & st.PostId2 & " (filtered); cancelled"
    SmkPlanned = True
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
            If sd >= Int(fromDate) And sd >= Date And (cl = 0 Or Date <= cl) And SmkPostingRows(id) > 0 Then
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

' ---- 10. rename ----
Private Function SmkRename(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim report As String
    If Not RenameEverywhere(st.Emp, SMK_EMP2, report) Then Exit Function
    If RosterRowOf(SMK_EMP2) = 0 Or RosterRowOf(SMK_EMP) > 0 Then why = "the roster name didn't change": Exit Function
    st.Emp = SMK_EMP2
    SmkRename = True
End Function

' ---- 11. publish to a scratch file, read it back ----
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
    Application.ScreenUpdating = False             ' the publish switches it back on
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

' ---- 12. void everything ----
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
    MarkCalDirty
    st.Note = refs.Count & " Ref IDs, " & nRows & " rows"
    SmkVoid = True
End Function

' ---- 13. remove ----
Private Function SmkRemove(ByRef st As SmkState, ByRef why As String) As Boolean
    Dim report As String
    If Not RemoveEmployee(st.Emp, Date, report) Then Exit Function
    If RosterRowOf(st.Emp) > 0 Then why = st.Emp & " is still on the roster": Exit Function
    If Not IsFormer(st.Emp) Then why = st.Emp & " isn't on the Former sheet": Exit Function
    SmkRemove = True
End Function

' ---- 14. integrity ----
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
