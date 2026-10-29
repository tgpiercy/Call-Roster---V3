Attribute VB_Name = "modPostingSync"
Option Explicit
' v27.1 (the refactor, Oct 2026): a person joining or leaving an open posting's list, the posting records, reopen, from modPostings; moved unchanged (tests/movecheck.py proves it)
Public Const PSYNC_VERSION As String = "Posting sync v29.0.3 (Oct 28)"

' What goes in the Result cell for one list person
Private Function ResultCellText(ByRef rec As PListRec) As String
    If rec.ExemptNote = "" Or rec.Result = rec.ExemptNote Then
        ResultCellText = rec.Result
    ElseIf rec.Result = "" Then
        ResultCellText = rec.ExemptNote
    Else
        ResultCellText = rec.Result & " | " & rec.ExemptNote
    End If
End Function
' ====================================================================
'  ONE PERSON JOINS OR LEAVES AN OPEN POSTING'S LIST  (v24.3.1)
'
'  A posting's list is frozen when it is created. Two things may change
'  it afterwards, both for a unit move:
'   - someone who moves INTO the unit (planned or today) is added to the
'     open postings for shifts on or after the move that can still be
'     signed up for, in the place call order gives them among the frozen
'     rows;
'   - a planned move that is cancelled takes them off again (and puts
'     them back on the old unit's postings made while it was waiting).
'  Nobody else's row, code or posted hours change; only Rank numbers
'  move to make room. A row is deleted only while the posting is Open
'  and only when it carries no sign-up, outcome, charge or Exempt note.
' ====================================================================
' Brings emp into line on the OPEN postings of unit for shifts on or after
' fromDate that were created on or after postedFrom (0 = any) or are named
' in alsoIds ("|P-0012|P-0014|").
'   addMissing  not on the list, holds a code in the unit on the shift date,
'               passes the posting's crew and grade filters, is free that
'               shift, and the posting can still be signed up for -> added
'   dropGone    on the list but holds no code in the unit on the shift date
'               -> taken off
'   refresh     on the list, holds a code, and the row's code or hours are
'               not today's (a cancelled move left the planned ones)
'               -> taken off and put back at today's code and hours
' addedTxt / removedTxt get the posting IDs; blockedTxt the postings where a
' row should have come off but was left (the only name on the list, or it
' carries a sign-up or an outcome), or where a name was not added because the
' posting can no longer be signed up for. An error raises to the caller,
' whose handler blocks the save.
Public Sub SyncPersonOnPostings(ByVal emp As String, ByVal unit As String, ByVal fromDate As Date, ByVal postedFrom As Date, _
                                ByVal addMissing As Boolean, ByVal dropGone As Boolean, ByVal refresh As Boolean, _
                                ByRef addedTxt As String, ByRef removedTxt As String, _
                                Optional ByRef blockedTxt As String, Optional ByVal alsoIds As String = "")
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Sub
    emp = Trim$(emp)
    Dim dc As Long: dc = DeptCol(unit)
    If emp = "" Or dc = 0 Then Exit Sub
    Dim rosterRow As Long: rosterRow = RosterRowOf(emp)
    If rosterRow = 0 Then Exit Sub                      ' not on the roster: the "left" marks cover them

    ' the postings first (the list table changes while we work)
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim ids() As String, nIds As Long, i As Long, id As String
    ReDim ids(1 To UBound(pv, 1))
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And SafeStr(pv(i, PS_STATUS)) = "Open" Then
            If StrComp(Trim$(SafeStr(pv(i, PS_UNIT))), Trim$(unit), vbTextCompare) = 0 Then
                If Int(ToDate(pv(i, PS_DATE))) >= Int(fromDate) Then
                    If Int(ToDate(pv(i, PS_POSTED))) >= Int(postedFrom) Or InStr(1, alsoIds, "|" & id & "|", vbTextCompare) > 0 Then
                        nIds = nIds + 1
                        ids(nIds) = id
                    End If
                End If
            End If
        End If
    Next i
    If nIds = 0 Then Exit Sub
    RefreshCalCache                                     ' "free that shift" must see time off booked earlier in this session

    Dim p As PostRec, lst() As PListRec, n As Long, j As Long, idx As Long, wasOn As Boolean, wantOff As Boolean
    Dim recs() As EmpRec, k As Long, hit As Long, codeNow As String, hrsNow As Double, holds As Boolean, changed As Boolean
    Dim blankP As PostRec
    codeNow = Trim$(SafeStr(RosterSheet().Cells(rosterRow, dc).Value))
    hrsNow = NumOr0(RosterSheet().Cells(rosterRow, C_TOTAL).Value2)
    For i = 1 To nIds
        p = blankP
        n = LoadPosting(ids(i), p, lst)
        If n > 0 And p.Status = "Open" Then
            idx = 0
            For j = 1 To n
                If StrComp(Trim$(lst(j).EmpName), emp, vbTextCompare) = 0 Then
                    idx = j
                    Exit For
                End If
            Next j
            wasOn = False
            ' do they hold a code in this unit on the shift date (a waiting move counts from its date)?
            holds = (UnitCodeOn(emp, unit, p.ShiftDate, codeNow) <> "")

            wantOff = False
            If idx > 0 Then
                If dropGone And Not holds Then wantOff = True
                If refresh And holds Then
                    ' only a row that isn't today's: the planned code, or hours that aren't their total
                    If StrComp(Trim$(lst(idx).DeptCode), codeNow, vbTextCompare) <> 0 Or Abs(lst(idx).TotalHrs - hrsNow) > 0.01 Then wantOff = True
                End If
            End If
            If wantOff Then
                If lst(idx).ExemptNote <> "" Then
                    ' already marked never-charged for another reason (moved, left): leave the row and its note
                ElseIf n > 1 And lst(idx).Result = "" And Not lst(idx).SignedUp And lst(idx).Charged = 0 Then
                    DeleteListRow p.ID, emp, n
                    wasOn = True
                    changed = True
                    idx = 0
                    n = LoadPosting(ids(i), p, lst)         ' the list without them
                Else
                    blockedTxt = blockedTxt & IIf(blockedTxt = "", "", ", ") & p.ID
                End If
            End If

            ' (a refresh puts back the row it has just taken off, whatever addMissing says;
            '  a new name only goes on a posting that can still be signed up for)
            If idx = 0 And holds And n > 0 And (wasOn Or (addMissing And StillOpenForSignup(p))) Then
                ' the same list CreatePostingFrom would build today for this posting: are they on it?
                k = LoadEligible(unit, CrewFilter(p.Crew), recs, DeptPostingMaxGrade(unit), True, p.ShiftDate)
                hit = 0
                For j = 1 To k
                    If StrComp(recs(j).EmpName, emp, vbTextCompare) = 0 Then
                        hit = j
                        Exit For
                    End If
                Next j
                If hit > 0 Then
                    If NotFreeReason(emp, recs(hit).EmpShift, p.ShiftDate, p.ShiftLbl) = "" And Not OtRestrictedOn(emp, p.ShiftDate) Then   ' v27 (item 21): a restricted mover is not added
                        InsertListRow p, lst, n, recs(hit)
                        addedTxt = addedTxt & IIf(addedTxt = "", "", ", ") & p.ID
                        changed = True
                        wasOn = False
                    End If
                End If
            End If
            If wasOn Then removedTxt = removedTxt & IIf(removedTxt = "", "", ", ") & p.ID      ' taken off and not put back
            ' not added only because the posting can no longer be signed up for: say so
            If idx = 0 And holds And n > 0 And addMissing And Not wasOn Then
                If Not StillOpenForSignup(p) Then blockedTxt = blockedTxt & IIf(blockedTxt = "", "", ", ") & p.ID
            End If
        End If
    Next i
    If changed Then
        MarkMembersDirty
        MarkCalDirty
    End If
End Sub

' Can someone still sign up for this posting? Not once its closing time has passed (PostingClosed)
' or the shift itself is in the past. (A name added after that could only be charged.)
Private Function StillOpenForSignup(ByRef p As PostRec) As Boolean
    If Int(p.ShiftDate) < Date Then Exit Function
    If PostingClosed(p.ClosesOn) Then Exit Function     ' v25 (P6): the one closing test
    StillOpenForSignup = True
End Function

' Adds one row to posting p's frozen list (lst, n: in rank order, as LoadPosting
' returns it) for the person in rec. Their place is before the first frozen row
' they come before in call order (the frozen code and posted hours, seniority
' from the roster); rows from there down move one rank.
Private Sub InsertListRow(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByRef rec As EmpRec)
    Dim frozen() As EmpRec, i As Long, pos As Long, newRank As Long
    ReDim frozen(1 To n)
    For i = 1 To n
        frozen(i).EmpName = lst(i).EmpName
        frozen(i).Grade = lst(i).Grade
        frozen(i).TotalHrs = lst(i).TotalHrs
        frozen(i).Seniority = RosterSeniorityKey(lst(i).EmpName)
    Next i
    pos = ListInsertIndex(frozen, n, rec)
    If pos <= n Then newRank = lst(pos).Rank Else newRank = lst(n).Rank + 1

    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    Dim ws As Worksheet: Set ws = lr.Parent
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    UnlockSheet ws
    For i = 1 To UBound(lv, 1)
        If UCase$(Trim$(SafeStr(lv(i, PL_ID)))) = UCase$(p.ID) Then
            If CLng(NumOr0(lv(i, PL_RANK))) >= newRank Then lr.DataBodyRange.Cells(i, PL_RANK).Value = CLng(NumOr0(lv(i, PL_RANK))) + 1
        End If
    Next i
    ProtectSheet ws

    Dim one() As Variant: ReDim one(1 To 1, 1 To PL_N)
    one(1, PL_ID) = p.ID
    one(1, PL_RANK) = newRank
    one(1, PL_EMP) = rec.EmpName
    one(1, PL_SHIFT) = rec.EmpShift
    one(1, PL_CODE) = rec.DeptCode
    one(1, PL_HRS) = rec.TotalHrs
    If Not AppendRows(lr, one, PL_N) Then Err.Raise vbObjectError + 771, , "the posting list row for " & p.ID & " couldn't be written"
    If ListRowCount(p.ID) <> n + 1 Then
        Err.Raise vbObjectError + 773, , p.ID & " has " & ListRowCount(p.ID) & " list rows after adding one (expected " & (n + 1) & ")"
    End If
End Sub

' How many rows of the list table belong to one posting
Private Function ListRowCount(ByVal id As String) As Long
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lr Is Nothing Then Exit Function
    If lr.DataBodyRange Is Nothing Then Exit Function
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(lv, 1)
        If UCase$(Trim$(SafeStr(lv(i, PL_ID)))) = UCase$(id) Then ListRowCount = ListRowCount + 1
    Next i
End Function

' Takes one person's row off an OPEN posting's frozen list (nWas = how many rows the
' posting had). The rows below move up one rank, so the numbers stay 1, 2, 3 ...
' The Postings sheet holds two tables side by side (postings in A:O, lists in Q:Z).
' Any filter is cleared first: in a filtered list Excel would take the whole sheet
' row, and a posting's header with it. The check afterwards counts BOTH tables.
Private Sub DeleteListRow(ByVal id As String, ByVal emp As String, ByVal nWas As Long)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim ws As Worksheet: Set ws = lr.Parent
    Dim lv As Variant, i As Long, rowAt As Long, cntWas As Long, rankWas As Long, postWas As Long
    UnlockSheet ws
    On Error Resume Next                                ' (each of these raises when there is nothing to clear)
    lr.AutoFilter.ShowAllData
    lo.AutoFilter.ShowAllData
    ws.ShowAllData
    On Error GoTo 0
    lv = RangeTo2D(lr.DataBodyRange)
    For i = 1 To UBound(lv, 1)
        If UCase$(Trim$(SafeStr(lv(i, PL_ID)))) = UCase$(id) Then
            If StrComp(Trim$(SafeStr(lv(i, PL_EMP))), Trim$(emp), vbTextCompare) = 0 Then
                rowAt = i
                Exit For
            End If
        End If
    Next i
    If rowAt = 0 Then
        ProtectSheet ws
        Err.Raise vbObjectError + 775, , "the list row for " & emp & " on " & id & " wasn't found"
    End If
    rankWas = CLng(NumOr0(lv(rowAt, PL_RANK)))
    cntWas = lr.ListRows.Count
    postWas = lo.ListRows.Count
    lr.ListRows(rowAt).Delete
    If Not lr.DataBodyRange Is Nothing And rankWas > 0 Then
        lv = RangeTo2D(lr.DataBodyRange)                ' the table as it is now
        For i = 1 To UBound(lv, 1)
            If UCase$(Trim$(SafeStr(lv(i, PL_ID)))) = UCase$(id) Then
                If CLng(NumOr0(lv(i, PL_RANK))) > rankWas Then lr.DataBodyRange.Cells(i, PL_RANK).Value = CLng(NumOr0(lv(i, PL_RANK))) - 1
            End If
        Next i
    End If
    ProtectSheet ws
    If lr.ListRows.Count <> cntWas - 1 Or ListRowCount(id) <> nWas - 1 Then
        Err.Raise vbObjectError + 772, , id & " has " & ListRowCount(id) & " list rows after taking one off (expected " & (nWas - 1) & ")"
    End If
    If lo.ListRows.Count <> postWas Then
        Err.Raise vbObjectError + 774, , "the postings table has " & lo.ListRows.Count & " rows after a list row was taken off (it had " & postWas & ")"
    End If
End Sub

' Seniority as call order uses it, for a name on a frozen list (blank / not on the roster sorts last)
Private Function RosterSeniorityKey(ByVal emp As String) As Double
    Dim r As Long: r = RosterRowOf(emp)
    If r = 0 Then
        RosterSeniorityKey = SeniorityKey(Empty)
    Else
        RosterSeniorityKey = SeniorityKey(RosterSheet().Cells(r, C_SENIORITY).Value)
    End If
End Function

' v24.3: the person now moves (or has moved) into unit after all, so a
' "Exempt - move cancelled (PM-n)" mark left by an earlier, cancelled move no
' longer applies on that unit's OPEN postings for shifts on or after fromDate.
' Returns the posting IDs ("" = none).
Public Function ClearCancelledMoveMarks(ByVal emp As String, ByVal unit As String, ByVal fromDate As Date) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Function
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    Dim openIds As Object: Set openIds = CreateObject("Scripting.Dictionary")
    openIds.CompareMode = 1
    Dim i As Long, id As String, hit As Boolean, outS As String, res As String, note As String
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And SafeStr(pv(i, PS_STATUS)) = "Open" Then
            If Int(ToDate(pv(i, PS_DATE))) >= Int(fromDate) Then
                If StrComp(Trim$(SafeStr(pv(i, PS_UNIT))), Trim$(unit), vbTextCompare) = 0 Then openIds(id) = True
            End If
        End If
    Next i
    If openIds.Count = 0 Then Exit Function
    For i = 1 To UBound(lv, 1)
        id = Trim$(SafeStr(lv(i, PL_ID)))
        If openIds.Exists(id) Then
            If StrComp(Trim$(SafeStr(lv(i, PL_EMP))), Trim$(emp), vbTextCompare) = 0 Then
                res = SafeStr(lv(i, PL_RESULT))
                SplitExempt res, note
                If InStr(1, note, EXEMPT_TAG & "move cancelled (PM-", vbTextCompare) > 0 Then
                    If Not hit Then
                        UnlockSheet lr.Parent
                        hit = True
                    End If
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = res
                    outS = outS & IIf(outS = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ClearCancelledMoveMarks = outS
End Function

' v24.3: a planned unit move was cancelled, so its "Exempt - moves ... (PM-0007)"
' marks come off OPEN postings. Only marks carrying that move's ID are touched.
' Returns the posting IDs ("" = none).
Public Function ClearMoveExemptions(ByVal emp As String, ByVal moveId As String) As String
    If Trim$(moveId) = "" Then Exit Function
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Function
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    Dim openIds As Object: Set openIds = CreateObject("Scripting.Dictionary")
    openIds.CompareMode = 1
    Dim i As Long, id As String, hit As Boolean, outS As String, res As String, note As String
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And SafeStr(pv(i, PS_STATUS)) = "Open" Then openIds(id) = True
    Next i
    If openIds.Count = 0 Then Exit Function
    For i = 1 To UBound(lv, 1)
        id = Trim$(SafeStr(lv(i, PL_ID)))
        If openIds.Exists(id) Then
            If StrComp(Trim$(SafeStr(lv(i, PL_EMP))), Trim$(emp), vbTextCompare) = 0 Then
                res = SafeStr(lv(i, PL_RESULT))
                SplitExempt res, note
                If InStr(1, note, "(" & Trim$(moveId) & ")", vbTextCompare) > 0 Then
                    If Not hit Then
                        UnlockSheet lr.Parent
                        hit = True
                    End If
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = res
                    outS = outS & IIf(outS = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ClearMoveExemptions = outS
End Function

Public Function HasActiveLogRows(ByVal id As String) As Boolean
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant, i As Long
    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, L_REF)))) = UCase$(id) And UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then
            HasActiveLogRows = True
            Exit Function
        End If
    Next i
End Function

Public Sub UpdatePostingRecords(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                                 ByVal newStatus As String, ByVal awardNames As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(POSTINGS_SHEET)
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    Dim v As Variant, i As Long, j As Long
    UnlockSheet ws

    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(p.ID) Then
            lo.DataBodyRange.Cells(i, PS_STATUS).Value = newStatus
            If newStatus = "Open" Then
                lo.DataBodyRange.Cells(i, PS_AWARDTO).ClearContents
                lo.DataBodyRange.Cells(i, PS_AWARDON).ClearContents
                lo.DataBodyRange.Cells(i, PS_AWARDBY).ClearContents
            Else
                lo.DataBodyRange.Cells(i, PS_AWARDTO).Value = awardNames
                lo.DataBodyRange.Cells(i, PS_AWARDON).Value = Now
                lo.DataBodyRange.Cells(i, PS_AWARDBY).Value = CurrentUser()
            End If
            Exit For
        End If
    Next i

    v = RangeTo2D(lr.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PL_ID)))) = UCase$(p.ID) Then
            For j = 1 To n
                If CLng(NumOr0(v(i, PL_RANK))) = lst(j).Rank Then
                    If newStatus = "Open" Then
                        lr.DataBodyRange.Cells(i, PL_SIGNED).ClearContents
                        If lst(j).ExemptNote <> "" Then          ' a reopened posting keeps "moved / left" (v24)
                            lr.DataBodyRange.Cells(i, PL_RESULT).Value = lst(j).ExemptNote
                        Else
                            lr.DataBodyRange.Cells(i, PL_RESULT).ClearContents
                        End If
                        lr.DataBodyRange.Cells(i, PL_CHARGED).ClearContents
                        If lr.ListColumns.Count >= PL_SIGNHRS Then lr.DataBodyRange.Cells(i, PL_SIGNHRS).ClearContents
                    Else
                        lr.DataBodyRange.Cells(i, PL_SIGNED).Value = IIf(lst(j).SignedUp, "Yes", "")
                        lr.DataBodyRange.Cells(i, PL_RESULT).Value = ResultCellText(lst(j))
                        lr.DataBodyRange.Cells(i, PL_CHARGED).Value = lst(j).Charged
                        If lr.ListColumns.Count >= PL_SIGNHRS Then
                            If lst(j).SignedUp And IsPartialSignup(p.Hours, lst(j).SignupHrs) Then
                                lr.DataBodyRange.Cells(i, PL_SIGNHRS).Value = lst(j).SignupHrs
                            Else
                                lr.DataBodyRange.Cells(i, PL_SIGNHRS).ClearContents
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    ProtectSheet ws
End Sub

' Reopen an awarded posting (after its log rows have been voided)
Public Sub ReopenPosting()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim id As Variant
    id = Application.InputBox(Prompt:="Posting ID to reopen (e.g. P-0012):", Title:="Reopen Posting", Type:=2)
    If VarType(id) = vbBoolean Then Exit Sub
    Dim report As String
    ReopenPostingById UCase$(Trim$(CStr(id))), report
    Exit Sub
Oops:
    ReportError "ReopenPosting"
End Sub

' v25: the reopen itself (ReopenPosting asks for the ID; SmokeTest calls this). True when the posting
' is open again; report = the outcome text (also shown, unless SmokeTest is running).
Public Function ReopenPostingById(ByVal id As String, ByRef report As String) As Boolean
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If Not PostingTablesReady() Then Exit Function
    Dim p As PostRec, lst() As PListRec, n As Long
    n = LoadPosting(CStr(id), p, lst)
    If n = 0 Then report = "Posting " & id & " not found.": Tell report, vbExclamation, "Reopen": Exit Function
    If p.Status = "Open" Then report = id & " is already open.": Tell report, vbInformation, "Reopen": Exit Function

    ' its OT Log rows must be voided first so nothing is charged twice
    If HasActiveLogRows(CStr(id)) Then
        report = id & " still has active OT Log rows. Void them first (Void Entry > " & id & "), then reopen."
        Tell report, vbExclamation, "Reopen"
        Exit Function
    End If
    If Int(p.ShiftDate) < RosterYearFrom() Then     ' v25 (P4): a shift in a closed year is never reopened
        report = id & " is for " & Format$(p.ShiftDate, "mmm d, yyyy") & ", before this roster year (" & Format$(RosterYearFrom(), "mmm d, yyyy") & "): it can't be reopened."
        Tell report, vbExclamation, "Reopen"
        Exit Function
    End If
    If ArchiveHasActiveRef(CStr(id)) Then          ' v23: its rows left the live log at the year-end rollover
        report = id & " belongs to a closed year, so it can't be reopened. Its OT Log rows are on the archive sheet and that year's file is frozen."
        Tell report, vbExclamation, "Reopen"
        Exit Function
    End If
    If Ask("Reopen " & id & " so it can be awarded again?", vbYesNo + vbQuestion, "Reopen", vbYes) = vbNo Then Exit Function
    Dim newlyExempt As String                      ' v25 (P5): who is marked not chargeable by this reopen
    ' v24: the list is frozen, the people aren't. Someone who has left by the shift
    ' date, or no longer holds a code in this unit, must not be charged on the re-award.
    ' (These two marks stay until the posting is closed. If the roster was wrong and is
    ' corrected while the posting is open, clear the person's Result cell on the Postings sheet.)
    Dim j As Long, leftOn As Date, dc As Long
    dc = DeptCol(p.Unit)
    Dim planId As String, planOn As Date, planUnit As String, codeNow As String
    For j = 1 To n
        ' v24.3: a mark left by a planned unit move that was cancelled since no longer applies
        ' ("Exempt - move cancelled (PM-n)" is a different mark and stays: that person never held the code)
        If InStr(1, lst(j).ExemptNote, EXEMPT_TAG & "moves ", vbTextCompare) > 0 Then
            If StrComp(PlannedStatusOf(PlanIdIn(lst(j).ExemptNote)), "Cancelled", vbTextCompare) = 0 Then lst(j).ExemptNote = ""
        End If
        If lst(j).ExemptNote = "" Then
            If RosterRowOf(lst(j).EmpName) = 0 Then
                leftOn = FormerLeftOn(lst(j).EmpName)
                If leftOn = 0 Then
                    lst(j).ExemptNote = EXEMPT_TAG & "not on the roster"
                    newlyExempt = newlyExempt & IIf(newlyExempt = "", "", ", ") & lst(j).EmpName & " (" & Mid$(lst(j).ExemptNote, Len(EXEMPT_TAG) + 1) & ")"
                ElseIf Int(leftOn) <= Int(p.ShiftDate) Then
                    lst(j).ExemptNote = EXEMPT_TAG & "left " & Format$(leftOn, "mmm d")
                    newlyExempt = newlyExempt & IIf(newlyExempt = "", "", ", ") & lst(j).EmpName & " (" & Mid$(lst(j).ExemptNote, Len(EXEMPT_TAG) + 1) & ")"
                End If
            ElseIf dc > 0 Then
                ' v24.3: the code they hold on the SHIFT date (a waiting unit move counts from its date,
                ' and its mark carries the move's ID so a cancel can take it off again)
                codeNow = Trim$(SafeStr(RosterSheet().Cells(RosterRowOf(lst(j).EmpName), dc).Value))
                If UnitCodeOn(lst(j).EmpName, p.Unit, p.ShiftDate, codeNow) = "" Then
                    If codeNow = "" Then
                        lst(j).ExemptNote = EXEMPT_TAG & "no code in " & p.Unit & " now"
                        newlyExempt = newlyExempt & IIf(newlyExempt = "", "", ", ") & lst(j).EmpName & " (" & Mid$(lst(j).ExemptNote, Len(EXEMPT_TAG) + 1) & ")"
                    Else
                        planId = PlannedFor(lst(j).EmpName, planOn, planUnit)
                        lst(j).ExemptNote = EXEMPT_TAG & "moves " & Format$(planOn, "mmm d") & " (" & planId & ")"
                        newlyExempt = newlyExempt & IIf(newlyExempt = "", "", ", ") & lst(j).EmpName & " (" & Mid$(lst(j).ExemptNote, Len(EXEMPT_TAG) + 1) & ")"
                    End If
                End If
            End If
        End If
    Next j
    UpdatePostingRecords p, lst, n, "Open", ""
    On Error Resume Next
    UpdateCoverageBySource p.ID, "POSTED " & p.ID
    RefreshCalendar
    On Error GoTo Oops
    If Not gSmokeRun Then RestoreSignups CStr(id)             ' v24.1: its electronic sign-ups count again (v25: never from SmokeTest)
    MarkMembersDirty
    SaveIfWritable
    report = id & " is open again." & IIf(newlyExempt <> "", vbNewLine & vbNewLine & "Not chargeable on it (moved, left, or no code now): " & newlyExempt, "")
    Tell report, vbInformation, "Reopen"
    ReopenPostingById = True
    Exit Function
Oops:
    ReportError "ReopenPostingById"
End Function


