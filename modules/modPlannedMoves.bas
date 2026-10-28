Attribute VB_Name = "modPlannedMoves"
Option Explicit
' v27.1 (the refactor, Oct 2026): planned unit moves (tblPlannedMoves), from modPeople; moved unchanged (tests/movecheck.py proves it)
Public Const PLANMOVE_VERSION As String = "Planned moves v29.0.2 (Oct 28)"

Private mPlan As Variant            ' v24.3: the Planned rows of tblPlannedMoves (Empty = none)
Private mPlanLoaded As Boolean
Private Const ADDED_TAG As String = "[added to "     ' v24.3.1: in a planned row's Note, the postings its entry put them on
' ====================================================================
'  PLANNED UNIT MOVES  (v24.3: tblPlannedMoves)
'
'  A unit move dated after today changes nothing on the roster. One row
'  goes into tblPlannedMoves (Status Planned). From then on every list
'  built for a shift ON OR AFTER that date treats the person as moved:
'  LoadEligible asks UnitCodeOn for each code, so they are off the
'  cleared units' lists and on the new unit's, at the hours they would
'  have after the top-up (PlannedListHours: for call order only, nothing
'  is logged). The first admin open on or after the date applies the move
'  through SaveUnitMove and the row becomes Applied. A planned move is
'  never edited: cancel it and enter it again.
' ====================================================================

' ---- pure rules (mirrored in tests/rules.py) ----
' The code emp holds in unit on date d once waiting moves are counted.
' plans = rows in the tblPlannedMoves layout (or Empty). Only Planned rows
' dated on or before d count; of two, the later date. movedIn = True when
' the answer is the planned code in the NEW unit.
Public Function PlannedCodeCore(ByRef plans As Variant, ByVal emp As String, ByVal unit As String, ByVal d As Date, _
                                ByVal rosterCode As String, Optional ByRef movedIn As Boolean) As String
    PlannedCodeCore = rosterCode
    movedIn = False
    If Not IsArray(plans) Then Exit Function
    Dim i As Long, best As Long, bestFrom As Double, t As Double
    For i = 1 To UBound(plans, 1)
        If IsPlannedStatus(plans(i, PM_STATUS)) Then
            If StrComp(Trim$(SafeStr(plans(i, PM_EMP))), Trim$(emp), vbTextCompare) = 0 Then
                If IsDate(plans(i, PM_FROM)) Then
                    t = Int(CDbl(CDate(plans(i, PM_FROM))))
                    If t <= Int(CDbl(d)) Then
                        If best = 0 Or t >= bestFrom Then
                            best = i
                            bestFrom = t
                        End If
                    End If
                End If
            End If
        End If
    Next i
    If best = 0 Then Exit Function
    If StrComp(Trim$(SafeStr(plans(best, PM_UNIT))), Trim$(unit), vbTextCompare) = 0 Then
        PlannedCodeCore = Trim$(SafeStr(plans(best, PM_CODE)))
        movedIn = True
    ElseIf InBarList(SafeStr(plans(best, PM_CLEAR)), unit) Then
        PlannedCodeCore = ""
    End If
End Function

' Is item one of the "|"-separated entries of listTxt? (trimmed, any case; a blank item never is)
' (TrimBarList gives the same list cleaned up: "A | B |" -> "A|B")
Public Function InBarList(ByVal listTxt As String, ByVal item As String) As Boolean
    Dim part As Variant
    item = Trim$(item)
    If item = "" Then Exit Function
    For Each part In Split(listTxt, "|")
        If StrComp(Trim$(CStr(part)), item, vbTextCompare) = 0 Then
            InBarList = True
            Exit Function
        End If
    Next part
End Function

' The Total a planned mover is listed at in the new unit before the top-up:
' the highest total among the others at grade PEER_MAX_GRADE and below
' (never another mover), never under their own. Nobody to match = their own.
Public Function PlannedListHours(ByVal own As Double, ByRef grades() As Double, ByRef isMover() As Boolean, _
                                 ByRef totals() As Double, ByVal n As Long) As Double
    Dim found As Boolean, top As Double
    top = HighestPeerTotal(grades, isMover, totals, n, found)
    PlannedListHours = own
    If found Then
        If top > own Then PlannedListHours = top
    End If
End Function

Private Function IsPlannedStatus(ByVal x As Variant) As Boolean
    IsPlannedStatus = (StrComp(Trim$(SafeStr(x)), PM_PLANNED, vbTextCompare) = 0)
End Function

' ---- lookups (the Planned rows are cached; any write to the table clears the cache) ----
Private Sub EnsurePlans()
    If mPlanLoaded Then Exit Sub
    mPlan = Empty
    Dim v As Variant: v = BodyOf(PlannedTable(False))
    Dim i As Long, c As Long, n As Long
    Dim a() As Variant
    If IsArray(v) Then
        If UBound(v, 2) >= PM_N Then
            For i = 1 To UBound(v, 1)
                If IsPlannedStatus(v(i, PM_STATUS)) Then n = n + 1
            Next i
            If n > 0 Then
                ReDim a(1 To n, 1 To PM_N)
                n = 0
                For i = 1 To UBound(v, 1)
                    If IsPlannedStatus(v(i, PM_STATUS)) Then
                        n = n + 1
                        For c = 1 To PM_N
                            a(n, c) = v(i, c)
                        Next c
                    End If
                Next i
                mPlan = a
            End If
        End If
    End If
    mPlanLoaded = True
End Sub

Public Sub ResetPlanCache()
    mPlan = Empty
    mPlanLoaded = False
End Sub

' The waiting moves, in the tblPlannedMoves layout (Empty = none)
Public Function PlannedRows() As Variant
    EnsurePlans
    PlannedRows = mPlan
End Function

Public Function HasPlannedMoves() As Boolean
    EnsurePlans
    HasPlannedMoves = IsArray(mPlan)
End Function

' The code emp holds in unit on date d (LoadEligible and the posting checks).
' rosterCode = what the roster's unit column holds today.
Public Function UnitCodeOn(ByVal emp As String, ByVal unit As String, ByVal d As Date, ByVal rosterCode As String, _
                           Optional ByRef movedIn As Boolean) As String
    movedIn = False
    EnsurePlans
    If Not IsArray(mPlan) Then
        UnitCodeOn = rosterCode
        Exit Function
    End If
    UnitCodeOn = PlannedCodeCore(mPlan, emp, unit, d, rosterCode, movedIn)
End Function

' The ID of the person's waiting move ("" = none), with its date and unit
Public Function PlannedFor(ByVal emp As String, Optional ByRef eff As Date, Optional ByRef toUnit As String) As String
    eff = 0: toUnit = ""
    EnsurePlans
    If Not IsArray(mPlan) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(mPlan, 1)
        If StrComp(Trim$(SafeStr(mPlan(i, PM_EMP))), Trim$(emp), vbTextCompare) = 0 Then
            PlannedFor = Trim$(SafeStr(mPlan(i, PM_ID)))
            If IsDate(mPlan(i, PM_FROM)) Then eff = Int(CDate(mPlan(i, PM_FROM)))
            toUnit = Trim$(SafeStr(mPlan(i, PM_UNIT)))
            Exit Function
        End If
    Next i
End Function

' "Phosphine | Cypure |" -> "Phosphine|Cypure" (each entry trimmed, blanks dropped)
Public Function TrimBarList(ByVal listTxt As String) As String
    Dim part As Variant, s As String
    For Each part In Split(listTxt, "|")
        If Trim$(CStr(part)) <> "" Then s = s & IIf(s = "", "", "|") & Trim$(CStr(part))
    Next part
    TrimBarList = s
End Function

' The first entry of a "|" list that isn't a unit name ("" = they all are)
Private Function FirstUnknownUnit(ByVal listTxt As String) As String
    Dim part As Variant
    For Each part In Split(listTxt, "|")
        If Trim$(CStr(part)) <> "" Then
            If DeptCol(Trim$(CStr(part))) = 0 Then
                FirstUnknownUnit = Trim$(CStr(part))
                Exit Function
            End If
        End If
    Next part
End Function

' Planned / Applied / Cancelled for one ID, whatever its state ("" = no such row)
Public Function PlannedStatusOf(ByVal id As String) As String
    Dim v As Variant, i As Long
    i = PlannedRowById(id, v)
    If i > 0 Then PlannedStatusOf = Trim$(SafeStr(v(i, PM_STATUS)))
End Function

' True when a waiting move takes effect after d1 and on or before d2
' (the publish's Who's next then loads its lists per date)
Public Function PlannedWithin(ByVal d1 As Date, ByVal d2 As Date) As Boolean
    EnsurePlans
    If Not IsArray(mPlan) Then Exit Function
    Dim i As Long, t As Date
    For i = 1 To UBound(mPlan, 1)
        If IsDate(mPlan(i, PM_FROM)) Then
            t = Int(CDate(mPlan(i, PM_FROM)))
            If t > Int(d1) And t <= Int(d2) Then
                PlannedWithin = True
                Exit Function
            End If
        End If
    Next i
End Function

' ---- the table ----
Public Function PlannedTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(PM_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(PM_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(PM_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet       ' adding a sheet activates it - put the user back
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = PM_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, PM_N).Value = Array("ID", "Employee", "New Unit", "Grade Code", "Codes To Clear", "New Crew", _
                                                    "Old Crew", "Effective From", "Status", "Entered By", "Entered On", "Closed On", "Note")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, PM_N), , xlYes)
        lo.Name = PM_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Range(ws.Columns(1), ws.Columns(PM_N)).ColumnWidth = 14
        ws.Columns(PM_EMP).ColumnWidth = 26
        ws.Columns(PM_CLEAR).ColumnWidth = 24
        ws.Columns(PM_FROM).ColumnWidth = 17
        ws.Columns(PM_STAMP).ColumnWidth = 20
        ws.Columns(PM_NOTE).ColumnWidth = 60
        ws.Range(ws.Columns(1), ws.Columns(PM_N)).NumberFormat = "@"
        ws.Columns(PM_FROM).NumberFormat = "ddd mmm d, yyyy"
        ws.Columns(PM_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ws.Columns(PM_CLOSED).NumberFormat = "mmm d, yyyy"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set PlannedTable = lo
End Function

' Appends rec(1, 1..PM_N). True when it is in the table.
Private Function AppendPlannedRow(ByRef rec() As Variant) As Boolean
    Dim lo As ListObject: Set lo = PlannedTable(True)
    If lo Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim em As String
    On Error GoTo Fail
    Application.EnableEvents = False
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, PM_N)
        .NumberFormat = "@"
        .Columns(PM_FROM).NumberFormat = "ddd mmm d, yyyy"
        .Columns(PM_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        .Columns(PM_CLOSED).NumberFormat = "mmm d, yyyy"
        .Value = rec
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetPlanCache
    AppendPlannedRow = True
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendPlannedRow", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetPlanCache
End Function

' The table's rows and the row number of one ID (0 = not there)
Private Function PlannedRowById(ByVal id As String, ByRef v As Variant) As Long
    v = BodyOf(PlannedTable(False))
    If Not IsArray(v) Then Exit Function
    If UBound(v, 2) < PM_N Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, PM_ID))), Trim$(id), vbTextCompare) = 0 Then
            PlannedRowById = i
            Exit Function
        End If
    Next i
End Function

' Sets Status + Closed On (newStatus "" = leave both) and the Note of one row
Private Function MarkPlannedRow(ByVal id As String, ByVal newStatus As String, ByVal noteTxt As String) As Boolean
    Dim v As Variant, r As Long
    r = PlannedRowById(id, v)
    If r = 0 Then Exit Function
    Dim lo As ListObject: Set lo = PlannedTable(False)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim em As String
    On Error GoTo Fail
    Application.EnableEvents = False
    UnlockSheet ws
    If newStatus <> "" Then
        lo.DataBodyRange.Cells(r, PM_STATUS).Value = newStatus
        lo.DataBodyRange.Cells(r, PM_CLOSED).Value = Date
    End If
    ' the "[added to P-0012, P-0014]" tag (postings they were put on when the move was entered) outlives later notes
    Dim tagTxt As String: tagTxt = AddedTag(SafeStr(v(r, PM_NOTE)))
    If tagTxt <> "" And InStr(1, noteTxt, ADDED_TAG, vbTextCompare) = 0 Then
        If Len(tagTxt) >= 240 Then
            noteTxt = tagTxt                    ' (never in practice: SavePlannedMove keeps the tag under 200)
        Else
            noteTxt = Left$(noteTxt, 249 - Len(tagTxt)) & " " & tagTxt
        End If
    End If
    lo.DataBodyRange.Cells(r, PM_NOTE).Value = Left$(noteTxt, 250)
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetPlanCache
    MarkPlannedRow = True
    Exit Function
Fail:
    em = Err.Description
    LogError "MarkPlannedRow", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetPlanCache
End Function

' "... [added to P-0012, P-0014]" -> "[added to P-0012, P-0014]" ("" = no tag)
Private Function AddedTag(ByVal noteTxt As String) As String
    Dim a As Long, b As Long
    a = InStr(1, noteTxt, ADDED_TAG, vbTextCompare)
    If a = 0 Then Exit Function
    b = InStr(a, noteTxt, "]")
    If b > a Then AddedTag = Mid$(noteTxt, a, b - a + 1)
End Function

' The posting IDs in a note's tag as "|P-0012|P-0014|" ("" = none)
Private Function AddedIds(ByVal noteTxt As String) As String
    Dim tagTxt As String: tagTxt = AddedTag(noteTxt)
    If tagTxt = "" Then Exit Function
    Dim part As Variant, s As String
    tagTxt = Mid$(tagTxt, Len(ADDED_TAG) + 1, Len(tagTxt) - Len(ADDED_TAG) - 1)
    For Each part In Split(tagTxt, ",")
        If Trim$(CStr(part)) <> "" Then s = s & "|" & Trim$(CStr(part))
    Next part
    If s <> "" Then AddedIds = s & "|"
End Function

' PM-0001, PM-0002 ... (rows are never removed, so the highest + 1 is free)
Private Function NextPlannedId() As String
    Dim v As Variant: v = BodyOf(PlannedTable(False))
    Dim i As Long, k As Long, mx As Long, s As String
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            s = Trim$(SafeStr(v(i, PM_ID)))
            If UCase$(Left$(s, 3)) = "PM-" Then
                If IsNumeric(Mid$(s, 4)) Then
                    k = CLng(Val(Mid$(s, 4)))
                    If k > mx Then mx = k
                End If
            End If
        Next i
    End If
    NextPlannedId = "PM-" & Format$(mx + 1, "0000")
End Function

' ---- entering one ----
' A unit move dated after today. Nothing on the roster's codes or Home Unit
' changes; a crew change that comes with it is dated now (crew history), so
' lists for later shifts use the right schedule. Verified like every people
' change: nobody's totals may move.
Public Function SavePlannedMove(ByVal emp As String, ByVal newUnit As String, ByVal gradeCode As String, _
                                ByVal clearUnits As String, ByVal newCrew As String, ByVal eff As Date, _
                                ByRef report As String) As Boolean
    Const TTL As String = "Planned move"
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim nRows As Long, r As Long
    r = RosterRowOf(emp, nRows)
    If r = 0 Then MsgBox emp & " isn't on the roster.", vbExclamation, TTL: Exit Function
    If nRows > 1 Then MsgBox emp & " is on the roster " & nRows & " times. Remove the extra row first.", vbExclamation, TTL: Exit Function
    gradeCode = Trim$(gradeCode)
    If DeptCol(newUnit) = 0 Or gradeCode = "" Then MsgBox "Pick the new unit and enter the grade code.", vbExclamation, TTL: Exit Function
    eff = Int(eff)
    If eff <= Date Then
        MsgBox "A planned move needs a date after today. A move dated today or earlier is saved straight away.", vbExclamation, TTL
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = RosterSheet()
    emp = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
    Dim had As String: had = PlannedFor(emp)
    If had <> "" Then
        MsgBox emp & " already has a planned move (" & had & "). Cancel it first: Needs Attention > Planned unit moves.", vbExclamation, TTL
        Exit Function
    End If

    ' crew change (optional): dated now, so "scheduled" and the rest rule are right for later shifts
    Dim crewNow As String: crewNow = NormCrew(SafeStr(ws.Cells(r, C_SHIFT).Value))
    Dim crewTo As String: crewTo = NormCrew(newCrew)
    Dim crewChanges As Boolean, found As Boolean, since As Date
    If crewTo <> "" And crewTo <> crewNow Then
        If Not IsCrewCode(crewTo) Then MsgBox "'" & newCrew & "' isn't a crew.", vbExclamation, TTL: Exit Function
        NewestCrew emp, found, since
        If found And since > eff Then
            MsgBox emp & "'s crew already has a change dated " & Format$(since, "mmm d, yyyy") & ", after this move's date. " & _
                   "Leave the crew box blank here and change the crew on the roster instead.", vbExclamation, TTL
            Exit Function
        End If
        crewChanges = True
    End If

    ' the codes that will be cleared on the day: only units they hold a code in, never the new unit
    Dim u As Variant, cleared As String
    For Each u In Split(clearUnits, "|")
        If CStr(u) <> "" And StrComp(CStr(u), newUnit, vbTextCompare) <> 0 Then
            If DeptCol(CStr(u)) > 0 Then
                If Trim$(SafeStr(ws.Cells(r, DeptCol(CStr(u))).Value)) <> "" Then cleared = cleared & IIf(cleared = "", "", "|") & CStr(u)
            End If
        End If
    Next u

    Dim bk As String
    If Not PeopleBackup("planned move", bk) Then Exit Function
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String, id As String, postTxt As String
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To PM_N)
    Application.Calculate
    Set before = RosterTotalsMap()
    Dim beforeCodes As Object: Set beforeCodes = RosterCodesMap()     ' v25: nothing on the roster changes today, codes and Home Unit included
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    id = NextPlannedId()
    rec(1, PM_ID) = id
    rec(1, PM_EMP) = emp
    rec(1, PM_UNIT) = newUnit
    rec(1, PM_CODE) = gradeCode
    rec(1, PM_CLEAR) = cleared
    rec(1, PM_CREW) = IIf(crewChanges, crewTo, "")
    rec(1, PM_OLDCREW) = IIf(crewChanges, crewNow, "")
    rec(1, PM_FROM) = eff
    rec(1, PM_STATUS) = PM_PLANNED
    rec(1, PM_USER) = CurrentUser()
    rec(1, PM_STAMP) = Now
    If Not AppendPlannedRow(rec) Then Err.Raise vbObjectError + 751, , "the Planned Moves sheet couldn't be written"
    ClearCancelledMoveMarks emp, newUnit, eff     ' an earlier, cancelled move to this unit left them not chargeable there
    If crewChanges Then
        UnlockSheet ws
        ws.Cells(r, C_SHIFT).Value = crewTo
        ProtectSheet ws
        If Not AddCrewRow(emp, crewTo, eff) Then Err.Raise vbObjectError + 752, , "the crew history row couldn't be written"
    End If
    If cleared <> "" Then
        postTxt = ExemptFromOpenPostings(emp, eff, "|" & cleared & "|", "moves " & Format$(eff, "mmm d") & " (" & id & ")")
    End If
    ' v24.3.1: postings already open in the new unit for shifts on or after the date get them on the list
    Dim addTxt As String, dropTxt As String
    Dim lateTxt As String, tagIds As String, cutAt As Long
    SyncPersonOnPostings emp, newUnit, eff, 0, True, False, False, addTxt, dropTxt, lateTxt
    If addTxt <> "" Then                        ' remembered on the row: a cancel puts exactly these right
        tagIds = addTxt
        Do While Len(tagIds) > 180              ' whole IDs only, so the tag always fits the Note
            cutAt = InStrRev(tagIds, ", ")
            If cutAt = 0 Then Exit Do
            tagIds = Left$(tagIds, cutAt - 1)
        Loop
        If Not MarkPlannedRow(id, "", ADDED_TAG & tagIds & "]") Then Err.Raise vbObjectError + 753, , "the Planned Moves row couldn't be updated"
    End If
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap())
    diff = diff & TotalsDifference(beforeCodes, RosterCodesMap())
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "The planned move did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, TTL
        Exit Function
    End If
    AfterPeopleChange
    SaveIfWritable
    report = id & ": " & emp & " moves to " & newUnit & " (" & gradeCode & ") on " & Format$(eff, "ddd mmm d, yyyy") & "." & vbNewLine & _
             "Nothing on the roster changed today. From that date they are called in " & newUnit & _
             IIf(cleared <> "", " and not in " & Replace(cleared, "|", ", "), "") & "." & vbNewLine
    If crewChanges Then
        report = report & "Crew " & crewNow & " -> " & crewTo & " from " & Format$(eff, "mmm d") & " (in the Crew History now)." & vbNewLine
    End If
    If postTxt <> "" Then report = report & "Not chargeable on open posting(s): " & postTxt & vbNewLine
    If addTxt <> "" Then report = report & "Added to the list of open " & newUnit & " posting(s): " & addTxt & ". Print those sheets again." & vbNewLine
    If lateTxt <> "" Then report = report & "NOT added to " & newUnit & " posting(s) " & lateTxt & ": their closing time has passed." & vbNewLine
    report = report & "It is applied the first time an admin opens the file on or after that date." & vbNewLine & _
             "To cancel it: Needs Attention > Planned unit moves, double-click the line." & vbNewLine & "Backup: " & bk
    SavePlannedMove = True
    Exit Function
Failed:
    em = Err.Description
    LogError "SavePlannedMove", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "The planned move stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, TTL
End Function

' ---- cancelling one ----
' The work of a cancel, with no questions and no save. The caller has events
' off and mBusy set. False (didTxt says why) = nothing was written; a write
' that fails raises, so the caller's handler blocks the save.
' syncPostings (v24.3.1): take them off the new unit's open posting lists and put
' them back on the old unit's lists made while the move was waiting. False when
' the person is being removed from the roster (the "left" marks cover that).
Public Function CancelPlannedCore(ByVal id As String, ByVal reason As String, ByRef didTxt As String, _
                                   Optional ByVal syncPostings As Boolean = True) As Boolean
    Dim v As Variant, i As Long
    i = PlannedRowById(id, v)
    If i = 0 Then didTxt = id & " isn't on the Planned Moves sheet.": Exit Function
    If Not IsPlannedStatus(v(i, PM_STATUS)) Then
        didTxt = id & " is already " & LCase$(Trim$(SafeStr(v(i, PM_STATUS)))) & "."
        Exit Function
    End If
    Dim emp As String: emp = Trim$(SafeStr(v(i, PM_EMP)))
    Dim crewTo As String: crewTo = NormCrew(SafeStr(v(i, PM_CREW)))
    Dim crewOld As String: crewOld = NormCrew(SafeStr(v(i, PM_OLDCREW)))
    Dim eff As Date
    If IsDate(v(i, PM_FROM)) Then eff = Int(CDate(v(i, PM_FROM)))
    Dim crewTxt As String, postTxt As String, r As Long, found As Boolean, since As Date, newest As String
    Dim ws As Worksheet
    Dim newUnit As String: newUnit = Trim$(SafeStr(v(i, PM_UNIT)))
    Dim clearTxt As String: clearTxt = TrimBarList(SafeStr(v(i, PM_CLEAR)))
    Dim stuckTxt As String, entered As Date
    If IsDate(v(i, PM_STAMP)) Then entered = Int(CDate(v(i, PM_STAMP)))
    Dim addedIds0 As String: addedIds0 = AddedIds(SafeStr(v(i, PM_NOTE)))      ' the postings the entry put them on
    r = RosterRowOf(emp)
    If crewTo <> "" And eff > 0 Then
        newest = NewestCrew(emp, found, since)
        If r > 0 Then
            If crewOld = "" Then
                crewTxt = " They had no crew before this move, so crew " & crewTo & " (from " & Format$(eff, "mmm d") & _
                          ") was left in place: correct column E by hand if it is wrong."
            ElseIf found And newest = crewTo And Int(since) = Int(eff) Then
                Set ws = RosterSheet()
                UnlockSheet ws
                ws.Cells(r, C_SHIFT).Value = crewOld
                ProtectSheet ws
                If eff > Date Then
                    ' v25 (H7, D-H3): the change never took effect - the row goes, so nothing sets a floor
                    ' under a later typed crew date and the history reads as if it had never been entered
                    If Not RemoveCrewRow(emp, crewTo, eff) Then Err.Raise vbObjectError + 761, , "the planned crew row couldn't be removed"
                    crewTxt = " The planned crew change (" & crewTo & " from " & Format$(eff, "mmm d") & ") was taken out of the Crew History; crew " & crewOld & " stands."
                Else
                    ' v25 (H6): the date has passed, so days were lived on the new crew - the undo is dated TODAY
                    If Not AddCrewRow(emp, crewOld, Date) Then Err.Raise vbObjectError + 761, , "the crew history row couldn't be written"
                    crewTxt = " Crew back to " & crewOld & " from today" & IIf(Date - 1 >= eff, " (the move's date had passed: " & Format$(eff, "mmm d") & " to " & _
                              Format$(Date - 1, "mmm d") & " stay on crew " & crewTo & " in the history).", " (the move's date was today).")
                End If
            Else
                crewTxt = " The crew was changed again after this move was entered, so it was left alone: check column E and the " & _
                          "Crew History row dated " & Format$(eff, "mmm d") & "."
            End If
        End If
    End If
    postTxt = ClearMoveExemptions(emp, id)
    If Not MarkPlannedRow(id, "Cancelled", reason) Then Err.Raise vbObjectError + 762, , "the Planned Moves row couldn't be updated"
    ' From here the move no longer counts on any list (MarkPlannedRow cleared the cache).
    ' Postings made or changed while it was waiting carry it on their frozen lists: put those right.
    Dim offTxt As String, reTxt As String, backTxt As String, skipTxt As String, skipTxt2 As String, keptTxt As String, u As Variant
    Dim lateBack As String
    Dim listsFrom As Date
    If syncPostings And eff > 0 And r > 0 Then
        ' new unit, no code there on the shift date -> off the list (whenever the posting was made)
        SyncPersonOnPostings emp, newUnit, eff, 0, False, True, False, skipTxt, offTxt, skipTxt2
        ' new unit, holds a code -> back to today's code and hours on the lists that carry the plan:
        ' the ones made since it was entered, and the ones its entry put them on
        listsFrom = entered
        If listsFrom = 0 Then listsFrom = DateSerial(9999, 12, 31)
        SyncPersonOnPostings emp, newUnit, eff, listsFrom, False, False, True, reTxt, offTxt, keptTxt, addedIds0
        ' old unit(s): back on the lists made without them
        If entered > 0 Then
            For Each u In Split(clearTxt, "|")
                If CStr(u) <> "" Then SyncPersonOnPostings emp, CStr(u), eff, entered, True, False, False, backTxt, skipTxt2, lateBack
            Next u
        End If
    End If
    ' whatever could not be taken off (the only name on a list, or a removal) is at least never charged
    If eff > 0 And r > 0 And DeptCol(newUnit) > 0 Then
        If Trim$(SafeStr(RosterSheet().Cells(r, DeptCol(newUnit)).Value)) = "" Then
            stuckTxt = ExemptFromOpenPostings(emp, eff, "|" & newUnit & "|", "move cancelled (" & id & ")")
        End If
    End If
    didTxt = id & " cancelled: " & emp & " stays where they are." & crewTxt
    If postTxt <> "" Then didTxt = didTxt & vbNewLine & "Chargeable again on open posting(s): " & postTxt & "."
    If offTxt <> "" Then didTxt = didTxt & vbNewLine & "Taken off the list of " & newUnit & " posting(s): " & offTxt & "."
    If reTxt <> "" Then didTxt = didTxt & vbNewLine & newUnit & " posting(s) " & reTxt & " list them at today's code and hours again."
    If backTxt <> "" Then didTxt = didTxt & vbNewLine & "Put back on the list of posting(s): " & backTxt & "."
    If stuckTxt <> "" Then
        didTxt = didTxt & vbNewLine & "Still on the list of " & newUnit & " posting(s) " & stuckTxt & ", marked not chargeable."
    End If
    If keptTxt <> "" Then
        didTxt = didTxt & vbNewLine & newUnit & " posting(s) " & keptTxt & " may still list them at the planned code and hours " & _
                 "(a sign-up or a result is on the row, or theirs is the only name): check those lists."
    End If
    If lateBack <> "" Then
        didTxt = didTxt & vbNewLine & "NOT put back on posting(s) " & lateBack & ": their closing time has passed."
    End If
    If syncPostings And eff > 0 And r > 0 And entered = 0 Then
        didTxt = didTxt & vbNewLine & "The day this move was entered can't be read, so posting lists made since then were left as they are: check the open postings."
    End If
    If offTxt <> "" Or reTxt <> "" Or backTxt <> "" Then didTxt = didTxt & vbNewLine & "Print those posting sheets again."
    CancelPlannedCore = True
End Function

' Cancels a waiting move (Needs Attention, SmokeTest). Verified: nobody's totals may move.
Public Function CancelPlannedMove(ByVal id As String, ByVal reason As String, ByRef report As String) As Boolean
    Const TTL As String = "Planned move"
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim v0 As Variant, i0 As Long
    i0 = PlannedRowById(id, v0)
    If i0 = 0 Then MsgBox id & " isn't on the Planned Moves sheet.", vbExclamation, TTL: Exit Function
    If Not IsPlannedStatus(v0(i0, PM_STATUS)) Then
        MsgBox id & " is already " & LCase$(Trim$(SafeStr(v0(i0, PM_STATUS)))) & ".", vbInformation, TTL
        Exit Function
    End If
    Dim bk As String
    If Not PeopleBackup("cancel planned move", bk) Then Exit Function
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String, didTxt As String, ok As Boolean
    reason = Trim$(reason)
    If reason = "" Then reason = "cancelled"
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    ok = CancelPlannedCore(id, reason & " (" & CurrentUser() & ", " & Format$(Date, "mmm d") & ")", didTxt)
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap())
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "Cancelling " & id & " did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, TTL
        Exit Function
    End If
    If Not ok Then
        MsgBox didTxt, vbExclamation, TTL
        Exit Function
    End If
    AfterPeopleChange
    SaveIfWritable
    report = didTxt & vbNewLine & "Backup: " & bk
    CancelPlannedMove = True
    Exit Function
Failed:
    em = Err.Description
    LogError "CancelPlannedMove", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "Cancelling " & id & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, TTL
End Function

' ---- applying the ones that are due ----
' Every Planned row dated today or earlier goes through ComputeUnitMove (hours
' as of the effective date) and SaveUnitMove, oldest first. One backup, one
' save, one message. A move that can't be applied stays Planned with the
' reason in its Note. byHand = run from Alt+F8 (says so when nothing is due).
' askFirst (v24.3.1, decided Oct 13 - review H1): the open offers to apply the due moves and
' No leaves them waiting. Before this the open applied them by itself, and a move whose
' after-check failed blocked the save on EVERY open (cancelling it needs a save too).
' onlyId (v25, H3): apply that one move (the Needs Attention line clicked), not every due move.
Public Sub ApplyDuePlannedMoves(Optional ByVal byHand As Boolean = False, Optional ByVal askFirst As Boolean = False, _
                                Optional ByVal onlyId As String = "")
    Const TTL As String = "Planned unit moves"
    If ThisWorkbook.ReadOnly Then
        If byHand Then MsgBox "This copy is read-only: nothing can be applied in it.", vbExclamation, TTL
        Exit Sub
    End If
    If NoSave() Then                                     ' v25 (H5): say so by hand, instead of nothing
        If byHand Then MsgBox "Saving is blocked in this file (an earlier check failed): close it WITHOUT saving, open it again, then apply the move.", vbExclamation, TTL
        Exit Sub
    End If
    If Not IsAdmin() Then Exit Sub
    Dim v As Variant: v = BodyOf(PlannedTable(False))
    Dim due() As Long, nDue As Long, i As Long, j As Long, t As Long
    If IsArray(v) Then
        If UBound(v, 2) >= PM_N Then
            ReDim due(1 To UBound(v, 1))
            For i = 1 To UBound(v, 1)
                If IsPlannedStatus(v(i, PM_STATUS)) Then
                    If IsDate(v(i, PM_FROM)) And (onlyId = "" Or StrComp(Trim$(SafeStr(v(i, PM_ID))), onlyId, vbTextCompare) = 0) Then
                        If Int(CDate(v(i, PM_FROM))) <= Date Then
                            nDue = nDue + 1
                            due(nDue) = i
                        End If
                    End If
                End If
            Next i
        End If
    End If
    If nDue = 0 Then
        If byHand Then MsgBox IIf(onlyId = "", "No planned move is due today.", onlyId & " isn't due yet, or isn't waiting."), vbInformation, TTL
        Exit Sub
    End If
    If LogTable() Is Nothing Then Exit Sub
    For i = 2 To nDue                                   ' oldest first
        t = due(i)
        j = i - 1
        Do While j >= 1
            If CDate(v(due(j), PM_FROM)) <= CDate(v(t, PM_FROM)) Then Exit Do
            due(j + 1) = due(j)
            j = j - 1
        Loop
        due(j + 1) = t
    Next i
    If askFirst Then
        Dim lines As String
        For i = 1 To nDue
            If i <= 10 Then
                lines = lines & "  " & Trim$(SafeStr(v(due(i), PM_ID))) & "  " & Trim$(SafeStr(v(due(i), PM_EMP))) & " -> " & _
                        Trim$(SafeStr(v(due(i), PM_UNIT))) & " (" & Trim$(SafeStr(v(due(i), PM_CODE))) & "), due " & _
                        Format$(CDate(v(due(i), PM_FROM)), "ddd mmm d") & vbNewLine
            End If
        Next i
        If nDue > 10 Then lines = lines & "  ... and " & (nDue - 10) & " more" & vbNewLine
        If MsgBox(nDue & " planned unit move(s) " & IIf(nDue = 1, "is", "are") & " due:" & vbNewLine & lines & vbNewLine & _
                  "Apply " & IIf(nDue = 1, "it", "them") & " now?" & vbNewLine & vbNewLine & _
                  "Yes = apply (a backup copy is saved first; one box then says what went in)." & vbNewLine & _
                  "No = leave " & IIf(nDue = 1, "it", "them") & " waiting. Needs Attention lists " & IIf(nDue = 1, "it", "them") & _
                  ", you are asked again at the next open, and Alt+F8 > ApplyPlannedMoves does it any time.", _
                  vbYesNo + vbQuestion, TTL) <> vbYes Then Exit Sub
    End If

    Dim bk As String, backedUp As Boolean
    Dim pc As PeopleCalc, rep As String, why As String, ok As Boolean, nRows As Long
    Dim id As String, emp As String, unit As String, code As String, eff As Date
    Dim okTxt As String, failTxt As String, nOk As Long, nFail As Long, noteTxt As String, clearList As String, repLine As Variant
    Dim refOut As String
    For i = 1 To nDue
        id = Trim$(SafeStr(v(due(i), PM_ID)))
        emp = Trim$(SafeStr(v(due(i), PM_EMP)))
        unit = Trim$(SafeStr(v(due(i), PM_UNIT)))
        code = Trim$(SafeStr(v(due(i), PM_CODE)))
        eff = Int(CDate(v(due(i), PM_FROM)))
        clearList = TrimBarList(SafeStr(v(due(i), PM_CLEAR)))
        ok = False: why = "": rep = ""
        If RosterRowOf(emp, nRows) = 0 Then
            why = "not on the roster"
        ElseIf nRows > 1 Then
            why = "on the roster " & nRows & " times"
        ElseIf DeptCol(unit) = 0 Then
            why = "'" & unit & "' isn't a unit"
        ElseIf code = "" Then
            why = "no grade code on the Planned Moves row"
        ElseIf FirstUnknownUnit(clearList) <> "" Then
            why = "'" & FirstUnknownUnit(clearList) & "' in Codes To Clear isn't a unit"
        ElseIf Not ComputeUnitMove(emp, unit, eff, pc) Then
            why = pc.Warnings
        Else
            If Not backedUp Then                        ' one backup, and only when a move is really about to go in
                If Not PeopleBackup("planned moves", bk) Then Exit Sub
                backedUp = True
            End If
            ok = SaveUnitMove(emp, unit, code, clearList, "", pc, rep, True, why, refOut)
        End If
        If NoSave() Then Exit Sub                       ' a failed check: SaveUnitMove has said so and blocked the save
        If ok Then
            nOk = nOk + 1
            noteTxt = IIf(pc.AddHours > 0, Format$(pc.AddHours, "0.0#") & " h added", "no hours added") & IIf(refOut <> "", " (" & refOut & ")", "")
            okTxt = okTxt & "  " & emp & " -> " & unit & " (" & code & "), " & noteTxt & vbNewLine
            For Each repLine In Split(rep, vbNewLine)   ' (v24.3.1 / v25 H4) the warnings, the exempt postings and the lists they were added to
                If InStr(1, CStr(repLine), "Added to the list", vbTextCompare) > 0 Or InStr(1, CStr(repLine), "Not chargeable", vbTextCompare) > 0 _
                   Or Left$(CStr(repLine), 6) = "Check:" Then okTxt = okTxt & "    " & CStr(repLine) & vbNewLine
                If Left$(CStr(repLine), 6) = "Check:" Then noteTxt = noteTxt & "; " & CStr(repLine)
            Next repLine
            If Not MarkPlannedRow(id, "Applied", "applied by " & CurrentUser() & ": " & noteTxt) Then
                ' the roster has moved but the row still says Planned: saving that would apply it again at the next open
                MarkNoSave
                MsgBox emp & "'s move went in, but its row on the Planned Moves sheet couldn't be marked Applied." & vbNewLine & vbNewLine & _
                       "Saving is blocked. Close WITHOUT saving and open the file again: the move is then applied afresh." & vbNewLine & _
                       "If it happens twice, send " & PRIMARY_ADMIN_NAME & " this box. Backup: " & bk, vbCritical, TTL
                Exit Sub
            End If
        Else
            nFail = nFail + 1
            If why = "" Then why = "the move didn't go through"
            failTxt = failTxt & "  " & id & "  " & emp & " -> " & unit & ": " & why & vbNewLine
            MarkPlannedRow id, "", "NOT applied " & Format$(Date, "mmm d") & ": " & why
        End If
    Next i
    AfterPeopleChange
    SaveIfWritable
    Dim msg As String
    If nOk > 0 Then msg = nOk & " planned unit move(s) applied:" & vbNewLine & Left$(okTxt, IIf(nFail > 0, 300, 500)) & vbNewLine
    If nFail > 0 Then
        msg = msg & nFail & " could NOT be applied and still wait(s):" & vbNewLine & Left$(failTxt, 400) & vbNewLine & _
              "Fix what the line says, then Alt+F8 > ApplyPlannedMoves. Or cancel it: Needs Attention > Planned unit moves." & vbNewLine & vbNewLine
    End If
    If backedUp Then msg = msg & "Backup: " & bk
    MsgBox msg, IIf(nFail > 0, vbExclamation, vbInformation), TTL
End Sub

' Alt+F8: applies the moves that are due (the open does this by itself;
' this is for a move that came due while the file was open)
Public Sub ApplyPlannedMoves()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    ApplyDuePlannedMoves True
    Exit Sub
Oops:
    ReportError "ApplyPlannedMoves"
End Sub

' Needs Attention, double-click on a planned move's line
Public Sub PlannedMoveMenu(ByVal id As String)
    Const TTL As String = "Planned move"
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim v As Variant, i As Long, eff As Date, txt As String, ans As Long, report As String
    i = PlannedRowById(id, v)
    If i = 0 Then MsgBox id & " isn't on the Planned Moves sheet.", vbExclamation, TTL: Exit Sub
    If Not IsPlannedStatus(v(i, PM_STATUS)) Then
        MsgBox id & " is already " & LCase$(Trim$(SafeStr(v(i, PM_STATUS)))) & ".", vbInformation, TTL
        Exit Sub
    End If
    If IsDate(v(i, PM_FROM)) Then eff = Int(CDate(v(i, PM_FROM)))
    txt = id & ":  " & Trim$(SafeStr(v(i, PM_EMP))) & " moves to " & Trim$(SafeStr(v(i, PM_UNIT))) & _
          " (" & Trim$(SafeStr(v(i, PM_CODE))) & ") on " & Format$(eff, "ddd mmm d, yyyy") & "." & vbNewLine & _
          "Codes cleared on the day: " & IIf(Trim$(SafeStr(v(i, PM_CLEAR))) = "", "none", Replace(SafeStr(v(i, PM_CLEAR)), "|", ", ")) & vbNewLine & _
          IIf(Trim$(SafeStr(v(i, PM_CREW))) <> "", "Crew: " & SafeStr(v(i, PM_OLDCREW)) & " -> " & SafeStr(v(i, PM_CREW)) & vbNewLine, "") & _
          IIf(Trim$(SafeStr(v(i, PM_NOTE))) <> "", "Note: " & SafeStr(v(i, PM_NOTE)) & vbNewLine, "") & vbNewLine
    If eff > 0 And eff <= Date Then
        ' v25 (H2 / D-H4-H5): two plain questions, and only THIS move is applied
        ans = MsgBox(txt & "This move is due. Apply it now?" & vbNewLine & vbNewLine & _
                     "Yes = apply this move (a backup copy is saved first).   No = not now.", _
                     vbYesNo + vbQuestion + vbDefaultButton2, TTL)
        If ans = vbYes Then
            ApplyDuePlannedMoves True, False, id
            RedrawNeedsIfShown
            Exit Sub
        End If
        If MsgBox("Cancel the move instead?" & vbNewLine & vbNewLine & _
                  "Yes = cancel it (the person stays where they are; postings are put right).   No = leave it waiting.", _
                  vbYesNo + vbQuestion + vbDefaultButton2, TTL) <> vbYes Then Exit Sub
    Else
        If MsgBox(txt & "Cancel this planned move?" & vbNewLine & vbNewLine & _
                  "To change it, cancel it and enter it again (Employee > Move to unit).", _
                  vbYesNo + vbQuestion + vbDefaultButton2, TTL) <> vbYes Then Exit Sub
    End If
    If CancelPlannedMove(id, "cancelled from Needs Attention", report) Then
        RedrawNeedsIfShown
        MsgBox report, vbInformation, TTL
    End If
    Exit Sub
Oops:
    ReportError "PlannedMoveMenu"
End Sub

Private Sub RedrawNeedsIfShown()
    Dim nm As String
    On Error Resume Next
    nm = ActiveSheet.Name
    On Error GoTo 0
    If nm = NA_SHEET Then NARefresh
End Sub


