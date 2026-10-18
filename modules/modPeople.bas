Attribute VB_Name = "modPeople"
Option Explicit

' ====================================================================
'  PEOPLE (v24)
'
'  CREW HISTORY. tblCrewHistory on the admin sheet "Crew History" holds
'  one row per person per crew: Employee | Crew | Effective From |
'  Logged By | Timestamp. CrewOn(employee, date) answers "which crew was
'  this person on that day": the row with the latest Effective From on or
'  before the date. A date before the person's first row gets that first
'  row's crew; a person with no rows gets the roster's crew column.
'  So a crew change never moves a past booking onto different shifts.
'
'  Roster column E always shows the newest crew. Typing a different crew
'  there asks for the effective date and appends a row; Cancel puts the
'  old crew back. A change made with macros off is picked up the next
'  time an admin opens the file, dated that day.
'
'  THE EMPLOYEE POP-UP (frmEmployee, ADMIN > Employee) calls:
'   SaveNewHire      roster row + crew row + ONE Adjustment (NH-####):
'                    the home unit's average Total Hours (grade 2 and
'                    below, not on leave) on the start date
'   SaveUnitMove     codes, Home Unit, Grade Log, crew, and ONE Adjustment
'                    (UM-####) up to the highest Total Hours among grade 2
'                    and below in the new unit; never lowered
'   RenameEverywhere roster, OT Log, posting lists and awards, absences,
'                    switches, calendar lines, Grade Log, Hours History,
'                    crew history - together, with a backup and a check
'   RemoveEmployee   moves the roster row to the Former sheet (values);
'                    BringBack puts it back with its formulas
'  Every one of them compares everybody's roster totals and the OT Log
'  total before and after, and blocks the save on a mismatch. Hours
'  changes are Adjustment rows, undone with Void Entry.
'
'  Pure (mirrored in tests/rules.py, checked by SelfTest): CrewOnCore,
'  NewHireAverage, HighestPeerTotal, UnitMoveTopUp, RenameInList,
'  ReplaceNameToken, PlannedCodeCore, PlannedListHours.
'
'  PLANNED UNIT MOVES (v24.3): a move dated after today is a row in
'  tblPlannedMoves; lists for shifts on or after the date already treat
'  the person as moved (UnitCodeOn, called by LoadEligible); the first
'  admin open on or after the date applies it through SaveUnitMove.
' ====================================================================

Public Const PPL_VERSION As String = "People v25.2 (Oct 17)"
Public Const CH_SHEET As String = "Crew History"
Public Const CH_TABLE As String = "tblCrewHistory"
Public Const CH_EMP As Long = 1
Public Const CH_CREW As Long = 2
Public Const CH_FROM As Long = 3
Public Const CH_USER As Long = 4
Public Const CH_STAMP As Long = 5
Public Const CH_N As Long = 5

' tblFormer (v24): one row per person taken off the roster, as values
Public Const FM_SHEET As String = "Former"
Public Const FM_TABLE As String = "tblFormer"
Public Const FM_EMP As Long = 1
Public Const FM_LEFT As Long = 2
Public Const FM_CREW As Long = 3
Public Const FM_HOME As Long = 4
Public Const FM_SENIORITY As Long = 5
Public Const FM_CODE1 As Long = 6           ' the eight unit codes, in roster column order (L:S)
Public Const FM_TOTAL As Long = 14
Public Const FM_WORKED As Long = 15
Public Const FM_REFUSED As Long = 16
Public Const FM_PROV As Long = 17
Public Const FM_YEAR As Long = 18           ' the roster year they were removed in
Public Const FM_USER As Long = 19
Public Const FM_STAMP As Long = 20
Public Const FM_STATUS As Long = 21         ' Former / Returned
Public Const FM_BACK As Long = 22           ' Returned On
Public Const FM_PAYROLL As Long = 23        ' v24.2: Payroll # (added on the end, so the columns before it keep their numbers)
Public Const FM_N As Long = 23

' tblPlannedMoves (v24.3): unit moves entered ahead of time. Rows are never removed.
Public Const PM_SHEET As String = "Planned Moves"
Public Const PM_TABLE As String = "tblPlannedMoves"
Public Const PM_ID As Long = 1              ' PM-0001
Public Const PM_EMP As Long = 2
Public Const PM_UNIT As Long = 3            ' the new unit
Public Const PM_CODE As Long = 4            ' the grade code they will hold there
Public Const PM_CLEAR As Long = 5           ' units whose code is cleared on the day, "|" between them
Public Const PM_CREW As Long = 6            ' new crew ("" = no change); its crew history row is written at entry
Public Const PM_OLDCREW As Long = 7         ' what column E held at entry (a cancel puts it back)
Public Const PM_FROM As Long = 8            ' Effective From
Public Const PM_STATUS As Long = 9          ' Planned / Applied / Cancelled
Public Const PM_USER As Long = 10
Public Const PM_STAMP As Long = 11
Public Const PM_CLOSED As Long = 12         ' the day it was applied or cancelled
Public Const PM_NOTE As Long = 13
Public Const PM_N As Long = 13
Public Const PM_PLANNED As String = "Planned"

' what the Employee pop-up shows before it saves a new hire or a unit move
Public Type PeopleCalc
    Emp As String
    Unit As String
    AsOf As Date
    NPeers As Long
    PeerNames As String
    Target As Double            ' new hire: the peers' average; unit move: the highest peer total
    Own As Double
    AddHours As Double
    Summary As String
    Warnings As String
End Type

' what CrewOnCore keeps while it walks a person's rows
Private Type CrewPick
    HaveLE As Boolean           ' a row dated on or before the date wanted
    DateLE As Double
    CrewLE As String
    HaveAny As Boolean          ' the person's earliest row
    DateFirst As Double
    CrewFirst As String
End Type

Private Const NO_CREW_ROW As String = "~no dated row~"

Private mHist As Variant            ' the table's rows (Empty = no rows)
Private mIdx As Object              ' employee -> Collection of row numbers in mHist
Private mHistLoaded As Boolean
Private mTestMode As Boolean        ' SelfTest feeds mHist from memory
Private mBusy As Boolean            ' our own writes to the roster fire change events
Private mPlan As Variant            ' v24.3: the Planned rows of tblPlannedMoves (Empty = none)
Private mPlanLoaded As Boolean
Private Const ADDED_TAG As String = "[added to "     ' v24.3.1: in a planned row's Note, the postings its entry put them on


' ====================================================================
'  PURE RULE
' ====================================================================
' The crew emp was on at date d. hist = rows in the tblCrewHistory layout.
' rowList (optional) = the row numbers that belong to emp, so the live
' lookup doesn't walk the whole table. fallback = what to answer when the
' person has no dated row at all.
Public Function CrewOnCore(ByRef hist As Variant, ByVal emp As String, ByVal d As Date, ByVal fallback As String, _
                           Optional ByVal rowList As Collection) As String
    CrewOnCore = fallback
    If Not IsArray(hist) Then Exit Function
    Dim st As CrewPick, i As Long, k As Variant
    If rowList Is Nothing Then
        For i = 1 To UBound(hist, 1)
            If StrComp(Trim$(SafeStr(hist(i, CH_EMP))), Trim$(emp), vbTextCompare) = 0 Then WeighCrewRow hist, i, d, st
        Next i
    Else
        For Each k In rowList
            WeighCrewRow hist, CLng(k), d, st
        Next k
    End If
    If st.HaveLE Then
        CrewOnCore = st.CrewLE
    ElseIf st.HaveAny Then
        CrewOnCore = st.CrewFirst
    End If
End Function

' Rows are weighed in table order: of two rows with the same date the later one wins.
Private Sub WeighCrewRow(ByRef hist As Variant, ByVal i As Long, ByVal d As Date, ByRef st As CrewPick)
    If Not IsDate(hist(i, CH_FROM)) Then Exit Sub
    Dim t As Double: t = Int(CDbl(CDate(hist(i, CH_FROM))))
    Dim crewTxt As String: crewTxt = NormCrew(SafeStr(hist(i, CH_CREW)))
    If t <= Int(CDbl(d)) Then
        If Not st.HaveLE Or t >= st.DateLE Then
            st.HaveLE = True: st.DateLE = t: st.CrewLE = crewTxt
        End If
    End If
    If Not st.HaveAny Or t < st.DateFirst Then
        st.HaveAny = True: st.DateFirst = t: st.CrewFirst = crewTxt
    End If
End Sub

' "c / d" -> "C/D", "a&b" -> "A/B", "Day" -> "DAYS"
Public Function NormCrew(ByVal s As String) As String
    s = UCase$(Replace(Trim$(s), " ", ""))
    s = Replace(s, "&", "/")
    If s = "DAY" Then s = "DAYS"
    NormCrew = s
End Function

' True for the crew codes the rotation knows (ShiftList in modOvertimeEntry)
Public Function IsCrewCode(ByVal s As String) As Boolean
    Dim c As Variant
    s = NormCrew(s)
    For Each c In ShiftList()
        If s = CStr(c) Then IsCrewCode = True: Exit Function
    Next c
End Function


' ====================================================================
'  LOOKUP
' ====================================================================
' The crew emp is on at date d. fallback = the roster crew when the caller
' already has it (saves a roster read); left out, the roster is read.
Public Function CrewOn(ByVal emp As String, ByVal d As Date, Optional ByVal fallback As Variant) As String
    EnsureHist
    Dim nm As String: nm = Trim$(emp)
    If Not mIdx Is Nothing Then
        If mIdx.Exists(nm) Then
            Dim rowList As Collection
            Set rowList = mIdx(nm)
            CrewOn = CrewOnCore(mHist, nm, d, NO_CREW_ROW, rowList)
            If CrewOn <> NO_CREW_ROW Then Exit Function         ' (rows with no readable date fall through to the roster crew)
            CrewOn = ""
        End If
    End If
    If IsMissing(fallback) Then
        If Not mTestMode Then CrewOn = EmpCrewCode(nm)
    Else
        CrewOn = CStr(fallback)
    End If
End Function

' The crew on the person's newest row (latest Effective From; of two rows
' with the same date, the later one). found = False when they have no dated row.
Public Function NewestCrew(ByVal emp As String, ByRef found As Boolean, Optional ByRef since As Date) As String
    found = False: since = 0
    EnsureHist
    If mIdx Is Nothing Then Exit Function
    Dim nm As String: nm = Trim$(emp)
    If Not mIdx.Exists(nm) Then Exit Function
    Dim rowList As Collection, k As Variant, t As Date
    Set rowList = mIdx(nm)
    For Each k In rowList
        If IsDate(mHist(CLng(k), CH_FROM)) Then
            t = Int(CDate(mHist(CLng(k), CH_FROM)))
            If Not found Or t >= since Then
                found = True: since = t
                NewestCrew = NormCrew(SafeStr(mHist(CLng(k), CH_CREW)))
            End If
        End If
    Next k
End Function

Private Sub EnsureHist()
    If mHistLoaded Then Exit Sub
    mHist = Empty
    Dim lo As ListObject: Set lo = CrewTable(False)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then mHist = RangeTo2D(lo.DataBodyRange)
        End If
    End If
    BuildCrewIndex
    mHistLoaded = True
End Sub

Private Sub BuildCrewIndex()
    Set mIdx = CreateObject("Scripting.Dictionary")
    mIdx.CompareMode = 1
    If Not IsArray(mHist) Then Exit Sub
    Dim i As Long, nm As String, rowList As Collection
    For i = 1 To UBound(mHist, 1)
        nm = Trim$(SafeStr(mHist(i, CH_EMP)))
        If nm <> "" Then
            If mIdx.Exists(nm) Then
                Set rowList = mIdx(nm)
            Else
                Set rowList = New Collection
                mIdx.Add nm, rowList
            End If
            rowList.Add i
        End If
    Next i
End Sub

' Forget the cached rows (after any write to the table, or an edit on its sheet)
Public Sub ResetCrewCache()
    If mTestMode Then Exit Sub
    mHist = Empty
    Set mIdx = Nothing
    mHistLoaded = False
End Sub

' SelfTest: rows in the tblCrewHistory layout (or Empty) replace the table
' until ClearTestCrewHistory. The sheet is never read or written.
Public Sub SetTestCrewHistory(ByVal rows As Variant)
    mTestMode = True
    mHist = rows
    BuildCrewIndex
    mHistLoaded = True
End Sub

Public Sub ClearTestCrewHistory()
    mTestMode = False
    mHist = Empty
    Set mIdx = Nothing
    mHistLoaded = False
End Sub


' ====================================================================
'  TABLE
' ====================================================================
Public Function CrewTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(CH_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(CH_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(CH_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet       ' adding a sheet activates it - put the user back
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CH_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, CH_N).Value = Array("Employee", "Crew", "Effective From", "Logged By", "Timestamp")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, CH_N), , xlYes)
        lo.Name = CH_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(CH_EMP).ColumnWidth = 26
        ws.Columns(CH_CREW).ColumnWidth = 8
        ws.Columns(CH_FROM).ColumnWidth = 17
        ws.Columns(CH_USER).ColumnWidth = 22
        ws.Columns(CH_STAMP).ColumnWidth = 20
        ws.Columns(CH_CREW).NumberFormat = "@"
        ws.Columns(CH_FROM).NumberFormat = "ddd mmm d, yyyy"
        ws.Columns(CH_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set CrewTable = lo
End Function

' Appends rows(1..n, 1..CH_N). True when they are in the table.
Private Function AppendCrewRows(ByRef rows() As Variant, ByVal n As Long) As Boolean
    If n < 1 Then AppendCrewRows = True: Exit Function
    Dim lo As ListObject: Set lo = CrewTable(True)
    If lo Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim outA() As Variant: ReDim outA(1 To n, 1 To CH_N)
    Dim i As Long, c As Long
    For i = 1 To n
        For c = 1 To CH_N
            outA(i, c) = rows(i, c)
        Next c
    Next i
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim em As String
    On Error GoTo Fail
    Application.EnableEvents = False
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, CH_N)
        .Columns(CH_CREW).NumberFormat = "@"
        .Value = outA
        .Columns(CH_FROM).NumberFormat = "ddd mmm d, yyyy"
        .Columns(CH_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetCrewCache
    ResetCalCacheAll
    MarkCalDirty
    AppendCrewRows = True
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendCrewRows", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetCrewCache
    MsgBox "The crew change couldn't be written to the Crew History sheet: " & em & vbNewLine & _
           "Send me this message.", vbExclamation, "Crew History"
End Function

' One crew row from code (the Employee pop-up). True when written.
Public Function AddCrewRow(ByVal emp As String, ByVal crewTxt As String, ByVal fromDate As Date) As Boolean
    Dim rows() As Variant: ReDim rows(1 To 1, 1 To CH_N)
    rows(1, CH_EMP) = Trim$(emp)
    rows(1, CH_CREW) = NormCrew(crewTxt)
    rows(1, CH_FROM) = Int(fromDate)
    rows(1, CH_USER) = CurrentUser()
    rows(1, CH_STAMP) = Now
    AddCrewRow = AppendCrewRows(rows, 1)
End Function

' v25 (H7): takes out ONE crew history row (the last one matching person, crew and date) - only a
' planned change that never took effect is ever removed. Any filter is cleared first (a filtered
' delete would take the whole sheet row); the count is checked after. False = not found / not removed.
Public Function RemoveCrewRow(ByVal emp As String, ByVal crewTxt As String, ByVal fromDate As Date) As Boolean
    Dim lo As ListObject: Set lo = CrewTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim v As Variant, i As Long, hit As Long, nWas As Long
    UnlockSheet ws
    On Error Resume Next
    lo.AutoFilter.ShowAllData
    ws.ShowAllData
    On Error GoTo 0
    v = lo.DataBodyRange.Value
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, CH_EMP))), Trim$(emp), vbTextCompare) = 0 Then
            If NormCrew(SafeStr(v(i, CH_CREW))) = NormCrew(crewTxt) And IsDate(v(i, CH_FROM)) Then
                If Int(CDate(v(i, CH_FROM))) = Int(fromDate) Then hit = i
            End If
        End If
    Next i
    If hit = 0 Then ProtectSheet ws: Exit Function
    nWas = lo.ListRows.Count
    lo.ListRows(hit).Delete
    ProtectSheet ws
    RemoveCrewRow = (lo.ListRows.Count = nWas - 1)
    ResetCrewCache
End Function

' Everyone on the roster without a dated row gets one: their roster crew,
' effective from the start of the roster year. Returns how many were added
' (-1 = the rows couldn't be written).
Public Function SeedCrewHistory() As Long
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_SHIFT)).Value    ' D:E, always 2-D
    Dim rows() As Variant: ReDim rows(1 To UBound(v, 1), 1 To CH_N)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Dim i As Long, n As Long, nm As String, found As Boolean
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    Dim who As String: who = CurrentUser()
    Dim stamp As Date: stamp = Now
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And Not seen.Exists(nm) Then
            seen.Add nm, True
            NewestCrew nm, found
            If Not found Then
                n = n + 1
                rows(n, CH_EMP) = nm
                rows(n, CH_CREW) = NormCrew(SafeStr(v(i, 2)))
                rows(n, CH_FROM) = yearFrom
                rows(n, CH_USER) = who
                rows(n, CH_STAMP) = stamp
            End If
        End If
    Next i
    If n > 0 Then
        If Not AppendCrewRows(rows, n) Then SeedCrewHistory = -1: Exit Function
    End If
    SeedCrewHistory = n
End Function


' ====================================================================
'  A CREW TYPED ON THE ROSTER
' ====================================================================
' From ThisWorkbook.Workbook_SheetChange. A sort hands us the whole block:
' every row is checked by NAME against the person's newest history row, so
' a sort (which changes no name -> crew pair) writes nothing.
Public Sub PeopleSheetChanged(ByVal Sh As Object, ByVal Target As Range)
    If mBusy Then Exit Sub
    If Target Is Nothing Then Exit Sub
    On Error GoTo Quit
    If Sh.Name = CH_SHEET Then              ' the table edited by hand
        ResetCrewCache
        MarkCalDirty
        Exit Sub
    End If
    If Sh.Name = PM_SHEET Then              ' v24.3: the planned moves edited by hand
        ResetPlanCache
        MarkCalDirty
        Exit Sub
    End If
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = RosterSheet()
    If Not Sh Is ws Then Exit Sub
    If CrewTable(False) Is Nothing Then Exit Sub        ' SetupPeople hasn't been run yet
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    Dim hit As Range
    Set hit = Application.Intersect(Target, ws.Range(ws.Cells(R_FIRST, C_SHIFT), ws.Cells(lastRow, C_SHIFT)))
    If hit Is Nothing Then Exit Sub

    Dim r1 As Long, r2 As Long, area As Range
    r1 = lastRow: r2 = R_FIRST
    For Each area In hit.Areas
        If area.Row < r1 Then r1 = area.Row
        If area.Row + area.Rows.Count - 1 > r2 Then r2 = area.Row + area.Rows.Count - 1
    Next area
    mBusy = True
    ReconcileCrew ws, r1, r2, True, CurrentUser()
Quit:
    mBusy = False
End Sub

' Workbook_Open (admins, writable): the crew history is brought up to date,
' then any planned unit move that is due is applied (v24.3). The two are
' separate so a problem in one can't hide the other.
Public Sub PeopleOnOpen()
    If ThisWorkbook.ReadOnly Or Not IsAdmin() Then Exit Sub
    CrewOnOpen
    PlannedOnOpen
End Sub

' Anyone new on the roster gets their first crew row; a crew changed while
' macros were off is dated today. Quiet.
Private Sub CrewOnOpen()
    On Error GoTo Quit
    If CrewTable(False) Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    mBusy = True
    ReconcileCrew ws, R_FIRST, lastRow, False, CurrentUser() & " (on open)"
Quit:
    mBusy = False
End Sub

Private Sub PlannedOnOpen()
    Dim en As Long, em As String
    On Error GoTo Oops
    ApplyDuePlannedMoves False, True            ' v24.3.1: the open ASKS; No leaves the moves waiting
    Exit Sub
Oops:
    en = Err.Number: em = Err.Description
    mBusy = False
    LogError "PlannedOnOpen", en, em
End Sub

' Compares roster rows r1..r2 with the crew history, by name.
'   no row yet       -> first row, effective from the roster year start
'   crew differs     -> interactive: ask the effective date (Cancel puts the
'                       old crew back); not interactive: effective today
Private Sub ReconcileCrew(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                          ByVal interactive As Boolean, ByVal who As String)
    Dim v As Variant: v = ws.Range(ws.Cells(r1, C_NAME), ws.Cells(r2, C_SHIFT)).Value    ' D:E, always 2-D
    Dim nRows As Long: nRows = UBound(v, 1)
    Dim rows() As Variant: ReDim rows(1 To nRows, 1 To CH_N)
    Dim chgIdx() As Long: ReDim chgIdx(1 To nRows)
    Dim chgOld() As String: ReDim chgOld(1 To nRows)
    Dim chgNew() As String: ReDim chgNew(1 To nRows)
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    Dim i As Long, n As Long, nChg As Long, nm As String, nowCrew As String, oldCrew As String
    Dim found As Boolean, since As Date, latest As Date, listTxt As String, bad As String
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    Dim stamp As Date: stamp = Now

    For i = 1 To nRows
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" And Not seen.Exists(nm) Then
            seen.Add nm, True
            nowCrew = NormCrew(SafeStr(v(i, 2)))
            oldCrew = NewestCrew(nm, found, since)
            If Not found Then
                n = n + 1
                rows(n, CH_EMP) = nm
                rows(n, CH_CREW) = nowCrew
                rows(n, CH_FROM) = yearFrom
                rows(n, CH_USER) = who
                rows(n, CH_STAMP) = stamp
            ElseIf oldCrew <> nowCrew Then
                nChg = nChg + 1
                chgIdx(nChg) = i
                chgOld(nChg) = oldCrew
                chgNew(nChg) = nowCrew
                If since > latest Then latest = since
                If nChg <= 8 Then
                    listTxt = listTxt & "  " & nm & ":  " & IIf(oldCrew = "", "(none)", oldCrew) & " -> " & _
                              IIf(nowCrew = "", "(none)", nowCrew) & vbNewLine
                End If
                If nowCrew <> "" And Not IsCrewCode(nowCrew) Then bad = bad & "  " & nm & ": " & nowCrew & vbNewLine
            End If
        End If
    Next i
    If nChg > 8 Then listTxt = listTxt & "  ...and " & (nChg - 8) & " more" & vbNewLine

    If nChg > 0 Then
        Dim eff As Date, keep As Boolean
        keep = True
        If interactive Then
            If bad <> "" Then
                MsgBox "Not a crew:" & vbNewLine & bad & vbNewLine & _
                       "Use A, B, C, D, A/B, C/D or DAYS. The old crew has been put back.", vbExclamation, "Crew"
                keep = False
            Else
                keep = AskCrewDate(listTxt, latest, eff)
            End If
            If Not keep Then PutCrewBack ws, r1, chgIdx, chgOld, nChg
        Else
            eff = Date
            If eff < latest Then eff = latest
        End If
        If keep Then
            For i = 1 To nChg
                n = n + 1
                rows(n, CH_EMP) = Trim$(SafeStr(v(chgIdx(i), 1)))
                rows(n, CH_CREW) = chgNew(i)
                rows(n, CH_FROM) = eff
                rows(n, CH_USER) = who
                rows(n, CH_STAMP) = stamp
            Next i
        End If
    End If
    If n > 0 Then AppendCrewRows rows, n
End Sub

' Asks when the typed crew change takes effect. False = cancelled.
Private Function AskCrewDate(ByVal listTxt As String, ByVal notBefore As Date, ByRef eff As Date) As Boolean
    Dim s As String, d As Date
    Do
        s = VBA.InputBox("Crew change:" & vbNewLine & listTxt & vbNewLine & _
                         "From which date? Time off and switches before that date keep the old crew's shifts." & vbNewLine & _
                         "Cancel puts the old crew back.", "Crew change", Format$(Date, "mmm d, yyyy"))
        If Trim$(s) = "" Then Exit Function
        If Not ParseTypedDate(s, d) Then
            MsgBox "That isn't a date I can read. Try  Jan 12  or  Jan 12, 2027.", vbExclamation, "Crew change"
        ElseIf notBefore > 0 And d < notBefore Then
            MsgBox "The date must be " & Format$(notBefore, "mmm d, yyyy") & " or later: that is when the crew being replaced started." & _
                   vbNewLine & "To correct an earlier row, edit it on the Crew History sheet." & ByHandHint(CH_SHEET), vbExclamation, "Crew change"
        ElseIf MsgBox("Crew change from " & Format$(d, "dddd, mmmm d, yyyy") & "?", vbYesNo + vbQuestion, "Crew change") = vbYes Then
            eff = d
            AskCrewDate = True
            Exit Function
        End If
    Loop
End Function

Private Sub PutCrewBack(ByVal ws As Worksheet, ByVal r1 As Long, ByRef chgIdx() As Long, _
                        ByRef chgOld() As String, ByVal nChg As Long)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim i As Long
    On Error Resume Next
    Application.EnableEvents = False
    UnlockSheet ws
    For i = 1 To nChg
        ws.Cells(r1 + chgIdx(i) - 1, C_SHIFT).Value = chgOld(i)
    Next i
    ProtectSheet ws
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  TYPED DATES
' ====================================================================
' v25 (D-C7): THE one typed-date rule, for every date box and date prompt. "" = d holds the
' date; otherwise the word for the hint: "blank", "not a date", "check the year". No year typed
' ("Jan 12", "1/12") = this year, or NEXT year when that is more than 90 days back (typing
' January in December). More than ~13 months from today either way is refused. A time typed
' after the date is ignored ("Oct 12 7pm", "10/12 7:00"). Pure: today is passed in (SelfTest
' suite typeddate; the year of a yearless date comes from today, never from the PC clock).
' v25 review: read by hand, never through CDate - CDate took "Feb 29" as Feb 2029 and "Sep 31" as
' Sep 1931, and a slash date in whatever order the PC's locale says. A yearless date lands in the
' year that puts it inside today - 90 .. today + 275 (one 365-day window: "Dec 20" typed on Jan 3
' is LAST December, "Jan 5" typed in December is next January).
Public Function TypedDateCore(ByVal t As String, ByVal today As Date, ByRef d As Date) As String
    t = Trim$(t)
    If t = "" Then TypedDateCore = "blank": Exit Function
    Dim y As Long, m As Long, dy As Long, hasYear As Boolean, dd As Date
    If Not TypedDateParts(t, y, m, dy, hasYear) Then TypedDateCore = "not a date": Exit Function
    If Not hasYear Then
        y = Year(today)
        If m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then TypedDateCore = "not a date": Exit Function
        dd = DateSerial(y, m, IIf(dy > 28, 28, dy))    ' the year is chosen first ("Feb 29" typed before a leap February)
        If dd < today - 90 Then
            y = y + 1
        ElseIf dd > today + 275 Then
            y = y - 1
        End If
    End If
    If Not ValidYmd(y, m, dy) Then TypedDateCore = "not a date": Exit Function
    dd = DateSerial(y, m, dy)
    If Abs(dd - today) > 400 Then TypedDateCore = "check the year": Exit Function
    d = dd
End Function

' The pieces of a typed date: "Oct 12", "October 12", "12 Oct", "Oct 12, 2026", "Oct 12 26",
' "2026-10-12", "10/12", "10/12/2026" (month / day, as the boxes' hints say). A time after it
' ("7:00", "7pm", "7 pm") is dropped. False = none of these.
Private Function TypedDateParts(ByVal t As String, ByRef y As Long, ByRef m As Long, ByRef dy As Long, ByRef hasYear As Boolean) As Boolean
    Dim raw As Variant, toks() As String, n As Long, w As String, i As Long, clockSeen As Boolean
    raw = Split(Trim$(Replace(Replace(t, ",", " "), ".", " ")), " ")
    ReDim toks(0 To UBound(raw) + 1)
    For i = 0 To UBound(raw)
        w = Trim$(CStr(raw(i)))
        If w = "" Then
        ElseIf InStr(w, ":") > 0 Then                               ' 7:00 - the time, dropped
            clockSeen = True
        ElseIf LCase$(w) Like "#*[ap]m" Then                        ' 7pm
        ElseIf LCase$(w) = "am" Or LCase$(w) = "pm" Then            ' "7 pm": the 7 goes too ("7:00 pm": nothing more to drop)
            If n > 0 And Not clockSeen Then
                If toks(n - 1) Like "#*" And Not (toks(n - 1) Like "*[-/]*") Then n = n - 1
            End If
        Else
            toks(n) = w: n = n + 1
        End If
    Next i
    hasYear = False
    If n = 1 Then
        If toks(0) Like "####-##-##" Or toks(0) Like "####-#-#" Or toks(0) Like "####-##-#" Or toks(0) Like "####-#-##" Then
            Dim p As Variant: p = Split(toks(0), "-")
            y = CLng(p(0)): m = CLng(p(1)): dy = CLng(p(2)): hasYear = True
            TypedDateParts = True
        ElseIf InStr(toks(0), "/") > 0 Then
            p = Split(toks(0), "/")
            If UBound(p) < 1 Or UBound(p) > 2 Then Exit Function
            For i = 0 To UBound(p)
                If Not (CStr(p(i)) Like "#" Or CStr(p(i)) Like "##" Or (i = 2 And CStr(p(i)) Like "####")) Then Exit Function
            Next i
            m = CLng(p(0)): dy = CLng(p(1))
            If UBound(p) = 2 Then
                y = CLng(p(2)): hasYear = True
                If y < 100 Then y = y + 2000
            End If
            TypedDateParts = True
        End If
        Exit Function
    End If
    If n < 2 Or n > 3 Then Exit Function
    Dim mo As Long
    mo = MonthIndex(toks(0))
    If mo > 0 Then
        If Not IsDigits(toks(1)) Then Exit Function
        m = mo: dy = CLng(toks(1))
    Else
        mo = MonthIndex(toks(1))
        If mo = 0 Or Not IsDigits(toks(0)) Then Exit Function
        m = mo: dy = CLng(toks(0))
    End If
    If n = 3 Then
        If Not IsDigits(toks(2)) Then Exit Function
        y = CLng(toks(2)): hasYear = True
        If y < 100 Then y = y + 2000
    End If
    TypedDateParts = True
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = (s <> "" And Len(s) <= 4 And s Like String$(Len(s), "#"))
End Function

' "Oct" / "October" / "oct" -> 10; 0 when not a month name
Private Function MonthIndex(ByVal s As String) As Long
    Dim k As Long, names As Variant
    names = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    s = LCase$(s)
    If Len(s) < 3 Then Exit Function
    For k = 0 To 11
        If Left$(s, 3) = names(k) Then
            If s Like "[a-z]*" And Not (s Like "*[!a-z]*") Then MonthIndex = k + 1
            Exit Function
        End If
    Next k
End Function

Private Function ValidYmd(ByVal y As Long, ByVal m As Long, ByVal dy As Long) As Boolean
    If y < 100 Or y > 9999 Or m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then Exit Function
    ValidYmd = (Day(DateSerial(y, m, dy)) = dy)
End Function

' Reads a typed date (no time part) as of today: True and d, or False (the box shows why).
Public Function ParseTypedDate(ByVal t As String, ByRef d As Date) As Boolean
    ParseTypedDate = (TypedDateCore(t, Date, d) = "")
End Function

' v25: the hint under a date box: "= Mon Oct 12, 2026" in green, or the word in red ("" when blank)
Public Sub TypedDateHint(ByVal t As String, ByVal lbl As Object, Optional ByVal suffix As String = "")
    Dim d As Date, w As String
    w = TypedDateCore(t, Date, d)
    If w = "blank" Then
        lbl.Caption = "": lbl.ForeColor = RGB(120, 120, 120)
    ElseIf w = "" Then
        lbl.Caption = "= " & Format$(d, "ddd mmm d, yyyy") & suffix: lbl.ForeColor = RGB(0, 110, 60)
    Else
        lbl.Caption = w: lbl.ForeColor = RGB(190, 30, 30)
    End If
End Sub

' v25: how every date box is pre-filled (no ambiguity about month / day order)
Public Function DateText(ByVal d As Date) As String
    DateText = Format$(d, "mmm d, yyyy")
End Function

' A typed date that must carry its year (a seniority date can be decades
' back, so nothing is guessed): "Mar 3, 2011", "3/3/2011". The pop-up shows
' the date back in words before anything is saved.
Public Function ParseFullDate(ByVal t As String, ByRef d As Date) As Boolean
    t = Trim$(t)
    If t = "" Then Exit Function
    If Not IsDate(t) Then Exit Function
    If Not TypedHasYear(t) Then Exit Function
    d = Int(CDate(t))
    ParseFullDate = (Year(d) >= 1950 And d <= Date + 400)
End Function

' True when the text includes a year: a 3+ digit number, three numbers
' (1/12/27), or a month name plus two numbers (Jan 12 27). v25: a time typed
' after the date ("10/12 7:00", "Oct 12 7pm") is left out first.
Private Function TypedHasYear(ByVal t As String) As Boolean
    Dim i As Long, ch As String, tok As String
    Dim nNum As Long, hasAlpha As Boolean, bigNum As Boolean
    Dim p As Long: p = InStr(t, ":")
    If p > 0 Then t = Left$(t, InStrRev(Left$(t, p), " "))
    Dim w As Variant, kept As String
    For Each w In Split(t, " ")
        If Not (LCase$(CStr(w)) Like "#*[ap]m") Then kept = kept & CStr(w) & " "
    Next w
    t = kept & " "
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[0-9]" Then
            tok = tok & ch
        Else
            If tok <> "" Then
                nNum = nNum + 1
                If Len(tok) >= 3 Then bigNum = True
                tok = ""
            End If
            If ch Like "[A-Za-z]" Then hasAlpha = True
        End If
    Next i
    TypedHasYear = bigNum Or nNum >= 3 Or (hasAlpha And nNum >= 2)
End Function


' ====================================================================
'  VERIFIED CHANGES: shared pieces
' ====================================================================
' Saves a dated backup copy next to the roster. False (and says why) if it couldn't.
Public Function PeopleBackup(ByVal what As String, ByRef bk As String) As Boolean
    If gSmokeRun Then                              ' v24.2: SmokeTest runs on a throw-away copy and never saves
        bk = "(none: SmokeTest)"
        PeopleBackup = True
        Exit Function
    End If
    bk = EnsureSlash(ThisWorkbook.Path) & "Backup before " & what & " " & Format$(Now, "yyyy-mm-dd hhmmss") & " - " & ThisWorkbook.Name
    Dim e As String
    On Error Resume Next
    SaveRestrictedCopy bk                          ' v25 (I2)
    If Err.Number <> 0 Then e = Err.Description
    On Error GoTo 0
    If e <> "" Then
        MsgBox "Couldn't save the backup copy, so nothing was changed:" & vbNewLine & e, vbCritical, "Backup"
        Exit Function
    End If
    PeopleBackup = True
End Function

' name -> "Total|Refused|sum of the pay-period cells" for every roster row.
' Two of these, taken before and after a change, show whose hours moved.
' v25 (H4 design note): what SavePlannedMove must leave untouched TODAY - each person's Home Unit
' and the eight unit-code columns (RosterTotalsMap covers the hours)
Public Function RosterCodesMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set RosterCodesMap = d
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, UnitColLast())).Value2
    Dim i As Long, c As Long, nm As String, s As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" Then
            s = Trim$(SafeStr(v(i, C_HOME)))
            For c = UnitColFirst() To UnitColLast()
                s = s & "|" & Trim$(SafeStr(v(i, c)))
            Next c
            d(nm) = s
        End If
    Next i
End Function

Public Function RosterTotalsMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set RosterTotalsMap = d
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim lastPer As Long: lastPer = PeriodLastCol()
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, lastPer)).Value2
    Dim i As Long, c As Long, nm As String, perSum As Double, bad As Boolean
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" Then
            perSum = 0: bad = False
            For c = C_PERIOD1 To lastPer
                If IsError(v(i, c)) Then bad = True Else perSum = perSum + NumOr0(v(i, c))
            Next c
            d(nm) = TotText(v(i, C_TOTAL)) & "|" & TotText(v(i, C_REFUSED)) & "|" & IIf(bad, "ERR", Format$(perSum, "0.0000"))
        End If
    Next i
End Function

Private Function TotText(ByVal x As Variant) As String
    If IsError(x) Then TotText = "ERR" Else TotText = Format$(NumOr0(x), "0.0000")
End Function

' Lists every name whose totals differ between two RosterTotalsMap results,
' or that is in one and not the other. skipName = the person the change is
' about (their own figures are checked by the caller). "" = nothing moved.
Public Function TotalsDifference(ByVal before As Object, ByVal after As Object, Optional ByVal skipName As String = "") As String
    Dim k As Variant, s As String, n As Long
    For Each k In before.Keys
        If StrComp(CStr(k), skipName, vbTextCompare) <> 0 Then
            If Not after.Exists(k) Then
                n = n + 1
                If n <= 10 Then s = s & "  " & k & ": no longer on the roster" & vbNewLine
            ElseIf before(k) <> after(k) Then
                n = n + 1
                If n <= 10 Then s = s & "  " & k & ": was " & TotalsPlain(before(k)) & ", now " & TotalsPlain(after(k)) & vbNewLine
            End If
        End If
    Next k
    For Each k In after.Keys
        If StrComp(CStr(k), skipName, vbTextCompare) <> 0 Then
            If Not before.Exists(k) Then
                n = n + 1
                If n <= 10 Then s = s & "  " & k & ": appeared on the roster" & vbNewLine
            End If
        End If
    Next k
    If n > 10 Then s = s & "  ...and " & (n - 10) & " more" & vbNewLine
    TotalsDifference = s
End Function

' "12.0000|21.0000|21.0000" -> "Total 12, Refused 21, periods 21"
Private Function TotalsPlain(ByVal packed As String) As String
    Dim p As Variant: p = Split(packed, "|")
    If UBound(p) < 2 Then TotalsPlain = packed: Exit Function
    TotalsPlain = "Total " & PlainNum(CStr(p(0))) & ", Refused " & PlainNum(CStr(p(1))) & ", periods " & PlainNum(CStr(p(2)))
End Function

Private Function PlainNum(ByVal s As String) As String
    If IsNumeric(s) Then PlainNum = HoursText(CDbl(s)) Else PlainNum = s
    If Right$(PlainNum, 1) = "." Or Right$(PlainNum, 1) = "," Then PlainNum = Left$(PlainNum, Len(PlainNum) - 1)
End Function

' Non-void Hours Charged in the whole OT Log
Public Function LogChargedTotal() As Double
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then LogChargedTotal = LogChargedTotal + NumOr0(v(i, L_CHARGED))
    Next i
End Function


' ====================================================================
'  PURE RULES: new hire, unit move
' ====================================================================
' New hire's starting hours: the average Total Hours of the home unit's
' peers (grade PEER_MAX_GRADE and below, not on leave on the start date),
' rounded to 0.1 like the leave top-up. nPeers = how many counted.
Public Function NewHireAverage(ByRef grades() As Double, ByRef onLeave() As Boolean, ByRef totals() As Double, _
                               ByVal n As Long, ByRef nPeers As Long) As Double
    nPeers = 0
    If n <= 0 Then Exit Function
    Dim keep() As Double: ReDim keep(1 To n)
    Dim i As Long
    For i = 1 To n
        If PeerEligible(grades(i), False, onLeave(i), True) Then
            nPeers = nPeers + 1
            keep(nPeers) = totals(i)
        End If
    Next i
    NewHireAverage = LeaveAverage(keep, nPeers)
End Function

' Unit move target: the highest Total Hours among the new unit's peers
' (grade PEER_MAX_GRADE and below, not the mover). found = False when
' there is no such peer.
Public Function HighestPeerTotal(ByRef grades() As Double, ByRef isMover() As Boolean, ByRef totals() As Double, _
                                 ByVal n As Long, ByRef found As Boolean) As Double
    found = False
    Dim i As Long
    For i = 1 To n
        If grades(i) <= PEER_MAX_GRADE And Not isMover(i) Then
            If Not found Or totals(i) > HighestPeerTotal Then HighestPeerTotal = totals(i)
            found = True
        End If
    Next i
End Function

' Hours added on a unit move: up to the target, never lowered (2 decimals)
Public Function UnitMoveTopUp(ByVal own As Double, ByVal highestPeer As Double) As Double
    If highestPeer > own Then UnitMoveTopUp = Int((highestPeer - own) * 100 + 0.5 + 0.000001) / 100
End Function

' One name inside "A; B (out of dept)" style text (tblPostings Awarded To)
Public Function RenameInList(ByVal txt As String, ByVal oldName As String, ByVal newName As String, ByRef hits As Long) As String
    Dim parts As Variant, i As Long, p As String, tail As String, outS As String
    If Trim$(txt) = "" Then RenameInList = txt: Exit Function
    parts = Split(txt, ";")
    For i = 0 To UBound(parts)
        p = Trim$(CStr(parts(i)))
        tail = ""
        If LCase$(Right$(p, 14)) = " (out of dept)" Then
            tail = Right$(p, 14)
            p = Trim$(Left$(p, Len(p) - 14))
        End If
        If StrComp(p, oldName, vbTextCompare) = 0 Then
            p = newName
            hits = hits + 1
        End If
        outS = outS & IIf(i = 0, "", "; ") & p & tail
    Next i
    RenameInList = outS
End Function

' A short calendar name ("Fred C") inside free text such as "Fred C 8h / Sam K":
' replaced only where it stands as a whole name, never inside a longer one.
Public Function ReplaceNameToken(ByVal txt As String, ByVal oldT As String, ByVal newT As String, ByRef hits As Long) As String
    Dim p As Long, startAt As Long, chBefore As String, chAfter As String
    startAt = 1
    If oldT = "" Then ReplaceNameToken = txt: Exit Function
    Do
        p = InStr(startAt, txt, oldT, vbTextCompare)
        If p = 0 Then Exit Do
        If p = 1 Then chBefore = "" Else chBefore = Mid$(txt, p - 1, 1)
        chAfter = Mid$(txt, p + Len(oldT), 1)
        If (chBefore = "" Or chBefore = " " Or chBefore = "/" Or chBefore = ";") And _
           (chAfter = "" Or chAfter = " " Or chAfter = "/" Or chAfter = ";" Or chAfter = "(") Then
            txt = Left$(txt, p - 1) & newT & Mid$(txt, p + Len(oldT))
            hits = hits + 1
            startAt = p + Len(newT)
        Else
            startAt = p + 1
        End If
    Loop
    ReplaceNameToken = txt
End Function


' ====================================================================
'  ROSTER LOOKUPS
' ====================================================================
' Roster row of a name (0 = not there). howMany (optional) = how many rows carry it.
Public Function RosterRowOf(ByVal emp As String, Optional ByRef howMany As Long) As Long
    howMany = 0
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = RangeTo2D(ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)))
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), Trim$(emp), vbTextCompare) = 0 Then
            howMany = howMany + 1
            If RosterRowOf = 0 Then RosterRowOf = R_FIRST + i - 1
        End If
    Next i
End Function

' The unit codes a person holds: units(1..n) / codes(1..n). Returns n.
Public Function PersonCodes(ByVal emp As String, ByRef units() As String, ByRef codes() As String) As Long
    Dim r As Long: r = RosterRowOf(emp)
    ReDim units(1 To 8): ReDim codes(1 To 8)
    If r = 0 Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim c As Long, n As Long, t As String
    For c = UnitColFirst() To UnitColLast()
        t = Trim$(SafeStr(ws.Cells(r, c).Value))
        If t <> "" Then
            n = n + 1
            units(n) = DeptOfCol(c)
            codes(n) = t
        End If
    Next c
    PersonCodes = n
End Function

' "Home unit Derivatives  |  crew C  |  codes: D2 Derivatives, P3 Phosphine"
Public Function PersonSummary(ByVal emp As String) As String
    Dim r As Long: r = RosterRowOf(emp)
    If r = 0 Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim units() As String, codes() As String, n As Long, i As Long, s As String
    n = PersonCodes(emp, units, codes)
    For i = 1 To n
        s = s & IIf(i = 1, "", ", ") & codes(i) & " " & units(i)
    Next i
    Dim hu As String: hu = Trim$(SafeStr(ws.Cells(r, C_HOME).Value))
    PersonSummary = "Home unit " & IIf(hu = "", "(blank)", hu) & "  |  crew " & Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value)) & _
                    "  |  codes: " & IIf(s = "", "(none)", s)
End Function

Private Function TableOn(ByVal sheetName As String, ByVal tableName As String) As ListObject
    If Not SheetExists(sheetName) Then Exit Function
    On Error Resume Next
    Set TableOn = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function BodyOf(ByVal lo As ListObject) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    BodyOf = RangeTo2D(lo.DataBodyRange)
End Function

' Application state for the people macros: captured before the error handler
' is set, restored on every exit path
Private Sub QuietOn(ByRef evWas As Boolean, ByRef scrWas As Boolean, ByRef calcWas As Long)
    evWas = Application.EnableEvents
    scrWas = Application.ScreenUpdating
    calcWas = Application.Calculation
End Sub

Private Sub QuietOff(ByVal evWas As Boolean, ByVal scrWas As Boolean, ByVal calcWas As Long)
    On Error Resume Next
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    mBusy = False
End Sub

Private Sub AfterPeopleChange()
    On Error Resume Next
    ResetCrewCache
    ResetPlanCache                          ' v24.3
    ResetCalCacheAll
    ResetSplitCache                         ' v24.1: the status bar's refused split is keyed on names
    MarkMembersDirty                        ' v24.1: the members' file is behind now
    MarkCalDirty
    SetRosterStatus
End Sub


' ====================================================================
'  FORMER EMPLOYEES (tblFormer)
' ====================================================================
Public Function FormerTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(FM_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(FM_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(FM_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = FM_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        Dim hdr() As Variant: ReDim hdr(1 To FM_N)
        Dim c As Long
        hdr(FM_EMP) = "Employee": hdr(FM_LEFT) = "Left On": hdr(FM_CREW) = "Crew": hdr(FM_HOME) = "Home Unit"
        hdr(FM_SENIORITY) = "Plant Seniority"
        For c = UnitColFirst() To UnitColLast()
            hdr(FM_CODE1 + c - UnitColFirst()) = DeptOfCol(c)
        Next c
        hdr(FM_TOTAL) = "Total Hours": hdr(FM_WORKED) = "Worked": hdr(FM_REFUSED) = "Refused": hdr(FM_PROV) = "Provisional"
        hdr(FM_YEAR) = "Roster Year": hdr(FM_USER) = "Removed By": hdr(FM_STAMP) = "Timestamp"
        hdr(FM_STATUS) = "Status": hdr(FM_BACK) = "Returned On": hdr(FM_PAYROLL) = "Payroll #"
        ws.Range("A1").Resize(1, FM_N).Value = hdr
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, FM_N), , xlYes)
        lo.Name = FM_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(FM_EMP).ColumnWidth = 26
        ws.Range(ws.Columns(FM_LEFT), ws.Columns(FM_N)).ColumnWidth = 12
        ws.Columns(FM_HOME).ColumnWidth = 16
        ws.Columns(FM_STAMP).ColumnWidth = 20
        ws.Rows(1).WrapText = True
        FormatFormer ws
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    If Not lo Is Nothing And create Then EnsureFormerColumns lo      ' never from a read (IntegrityCheck, SystemCheck)
    Set FormerTable = lo
End Function

' A v24 Former table has 22 columns: Payroll # (v24.2) goes on the end, once
Private Sub EnsureFormerColumns(ByVal lo As ListObject)
    If lo.ListColumns.Count >= FM_N Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    lo.ListColumns.Add.Name = "Payroll #"
    ws.Columns(FM_PAYROLL).ColumnWidth = 12
    ws.Columns(FM_PAYROLL).NumberFormat = "@"
    ProtectSheet ws
    On Error GoTo 0
End Sub

Private Sub FormatFormer(ByVal ws As Worksheet)
    ws.Columns(FM_LEFT).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_BACK).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_SENIORITY).NumberFormat = "mmm d, yyyy"
    ws.Columns(FM_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ws.Columns(FM_CREW).NumberFormat = "@"
    ws.Columns(FM_PAYROLL).NumberFormat = "@"
End Sub

' Row in tblFormer of a person who is still "Former" (0 = none)
Private Function FormerRowOf(ByVal emp As String, ByRef v As Variant) As Long
    v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = UBound(v, 1) To 1 Step -1
        If StrComp(Trim$(SafeStr(v(i, FM_EMP))), Trim$(emp), vbTextCompare) = 0 And _
           StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then FormerRowOf = i: Exit Function
    Next i
End Function

' True for someone moved to the Former sheet and not brought back
Public Function IsFormer(ByVal emp As String) As Boolean
    Dim v As Variant
    IsFormer = (FormerRowOf(emp, v) > 0)
End Function

' The day a person still on the Former sheet left (0 = not there)
Public Function FormerLeftOn(ByVal emp As String) As Date
    Dim v As Variant, fr As Long
    fr = FormerRowOf(emp, v)
    If fr = 0 Then Exit Function
    If IsDate(v(fr, FM_LEFT)) Then FormerLeftOn = Int(CDate(v(fr, FM_LEFT)))
End Function

' Everyone on the Former sheet (Former or Returned): name -> True.
' IntegrityCheck and the Kronos import use it: these names are known.
Public Function FormerNameSet() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set FormerNameSet = d
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim i As Long, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, FM_EMP)))
        If nm <> "" Then d(nm) = True
    Next i
End Function

' Names still "Former", A-Z, for the Bring back list (Empty if none)
Public Function FormerNames() As Variant
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim a() As String: ReDim a(0 To UBound(v, 1) - 1)
    Dim i As Long, j As Long, n As Long, t As String
    For i = 1 To UBound(v, 1)
        t = Trim$(SafeStr(v(i, FM_EMP)))
        If t <> "" And StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then
            a(n) = t & "   (left " & IIf(IsDate(v(i, FM_LEFT)), Format$(v(i, FM_LEFT), "mmm d, yyyy"), "?") & ")"
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve a(0 To n - 1)
    For i = 1 To n - 1
        t = a(i): j = i - 1
        Do While j >= 0
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    FormerNames = a
End Function

' Rows of the Former sheet removed in roster year y, for the rollover's Year
' End record: out(1..n, 1..6) = name, home unit, total, worked, refused, provisional
Public Function FormerFinals(ByVal y As Long, ByRef out() As Variant) As Long
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    ReDim out(1 To UBound(v, 1), 1 To 6)
    Dim i As Long, n As Long
    For i = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(i, FM_EMP))) <> "" And CLng(NumOr0(v(i, FM_YEAR))) = y And _
           StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then
            n = n + 1
            out(n, 1) = Trim$(SafeStr(v(i, FM_EMP)))
            out(n, 2) = SafeStr(v(i, FM_HOME))
            out(n, 3) = NumOr0(v(i, FM_TOTAL))
            out(n, 4) = NumOr0(v(i, FM_WORKED))
            out(n, 5) = NumOr0(v(i, FM_REFUSED))
            out(n, 6) = NumOr0(v(i, FM_PROV))
        End If
    Next i
    FormerFinals = n
End Function

Private Function AppendFormerRow(ByRef rec() As Variant) As Boolean
    Dim lo As ListObject: Set lo = FormerTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < FM_N Then Exit Function      ' its Payroll # column couldn't be added (read-only?)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    FormatFormer ws
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, FM_N).Value = rec
    ProtectSheet ws
    AppendFormerRow = True
End Function


' ====================================================================
'  A NEW ROSTER ROW (Add employee, Bring back)
' ====================================================================
' Writes the row under the last name: formats and drop-downs copied from the
' row above, the values, and the same formulas every other row has. Events
' must be off and the caller handles errors. codes(0..7) = the unit codes in
' roster column order (L:S). Returns the row number.
Private Function AddRosterRow(ByVal ws As Worksheet, ByVal nm As String, ByVal crewTxt As String, ByVal homeUnit As String, _
                              ByVal seniority As Variant, ByRef codes() As String, ByVal phone As String, _
                              ByVal email As String, Optional ByVal payrollNo As String = "") As Long
    ' payrollNo (v24.2): Bring back passes the number kept on the Former sheet; a new hire's is filled by the next Kronos import
    Dim r As Long: r = RosterLastRow() + 1
    If r < R_FIRST Then r = R_FIRST
    UnlockSheet ws
    If ws.FilterMode Then ws.ShowAllData
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))) > 0 Then
        Err.Raise vbObjectError + 711, , "row " & r & " (under the last name on the roster) isn't empty - clear it and try again" & ByHandHint(ws.Name)
    End If
    If r > R_FIRST Then
        ws.Range(ws.Cells(r - 1, 1), ws.Cells(r - 1, C_LAST)).Copy
        ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST)).PasteSpecial Paste:=xlPasteFormats
        ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST)).PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False
        ws.Rows(r).RowHeight = ws.Rows(r - 1).RowHeight
    End If
    ws.Cells(r, C_NAME).Interior.ColorIndex = xlNone        ' the Kronos import colours matched names
    ws.Cells(r, C_PAYROLL).NumberFormat = "@"
    ws.Cells(r, C_PAYROLL).Value = payrollNo
    ws.Cells(r, C_HOME).Value = homeUnit
    ws.Cells(r, C_NAME).Value = nm
    ws.Cells(r, C_SHIFT).Value = crewTxt
    ws.Cells(r, C_PHONE).Value = phone
    If IsDate(seniority) Then
        ws.Cells(r, C_SENIORITY).Value = Int(CDate(seniority))
    ElseIf Not IsEmpty(seniority) Then
        ws.Cells(r, C_SENIORITY).Value = seniority
    End If
    ws.Cells(r, C_EMAIL).Value = email
    Dim c As Long
    For c = UnitColFirst() To UnitColLast()
        ws.Cells(r, c).Value = codes(c - UnitColFirst())
    Next c
    ws.Cells(r, C_WORKED).Value = 0                         ' Kronos fills it at the next import
    WritePeriodRow ws, r
    WriteRefusedRow ws, r
    WriteTotalRow ws, r
    AddRosterRow = r
End Function


' ====================================================================
'  ADD EMPLOYEE  (NH-####)
' ====================================================================
' Works out a new hire's starting hours. False = can't (pc.Warnings says why).
Public Function ComputeNewHire(ByVal homeUnit As String, ByVal startDate As Date, ByVal newName As String, _
                               ByRef pc As PeopleCalc) As Boolean
    Dim blank As PeopleCalc: pc = blank
    If DeptCol(homeUnit) = 0 Then pc.Warnings = "Pick the home unit.": Exit Function
    pc.Emp = Trim$(newName): pc.Unit = homeUnit
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    pc.AsOf = Int(startDate)
    If pc.AsOf < yearFrom Then
        pc.AsOf = yearFrom
        pc.Warnings = "The start date is before this roster year, so hours are taken at " & Format$(yearFrom, "mmm d, yyyy") & ". "
    ElseIf pc.AsOf > Date Then
        pc.Warnings = "The start date is in the future: the hours used are today's. "
    End If

    Dim recs() As EmpRec, n As Long, i As Long, nP As Long, snap As Date, note As String, flagged As Boolean
    On Error GoTo Failed
    BeginHoursBatch
    RefreshCalCache
    n = LoadEligible(homeUnit, "ALL", recs)
    Dim grades() As Double, onLeave() As Boolean, totals() As Double
    ReDim grades(1 To IIf(n > 0, n, 1)): ReDim onLeave(1 To IIf(n > 0, n, 1)): ReDim totals(1 To IIf(n > 0, n, 1))
    Dim k As Long
    For i = 1 To n
        If StrComp(recs(i).EmpName, pc.Emp, vbTextCompare) <> 0 Then
            k = k + 1
            grades(k) = recs(i).Grade
            onLeave(k) = HasLeaveOverlapping(recs(i).EmpName, pc.AsOf, pc.AsOf)
            If grades(k) <= PEER_MAX_GRADE And Not onLeave(k) Then
                totals(k) = TotalHoursAsOf(recs(i).EmpName, pc.AsOf, snap, note)
                pc.PeerNames = pc.PeerNames & IIf(pc.PeerNames = "", "", ", ") & CalName(recs(i).EmpName) & " " & Format$(totals(k), "0.0")
                If note <> "" And Not flagged Then flagged = True: pc.Warnings = pc.Warnings & "Hours: " & note & ". "
            End If
        End If
    Next i
    EndHoursBatch
    pc.Target = NewHireAverage(grades, onLeave, totals, k, nP)
    pc.NPeers = nP
    pc.Own = 0
    pc.AddHours = pc.Target
    If nP = 0 Then pc.Warnings = pc.Warnings & "Nobody in " & homeUnit & " at grade " & PEER_MAX_GRADE & " and below is available to average, so the starting hours are 0. "
    pc.Summary = homeUnit & ": " & nP & " operator" & IIf(nP = 1, "", "s") & " at grade " & PEER_MAX_GRADE & " and below, not on leave, as of " & _
                 Format$(pc.AsOf, "mmm d, yyyy") & ". Average " & Format$(pc.Target, "0.0") & " h. " & _
                 IIf(pc.Emp = "", "The new hire", pc.Emp) & " starts at " & Format$(pc.AddHours, "0.0") & " h."
    ComputeNewHire = True
    Exit Function
Failed:
    pc.Warnings = "The calculation stopped: " & Err.Description
    On Error Resume Next
    EndHoursBatch
End Function

' Adds the roster row, the crew history row and (when the average isn't 0)
' ONE Adjustment (Ref NH-####, Reason "New hire", Note = the calculation).
' hoursToAdd may be an admin override (then overrideReason goes in the Note).
Public Function SaveNewHire(ByVal nm As String, ByVal phone As String, ByVal email As String, ByVal crewTxt As String, _
                            ByVal seniority As Variant, ByVal homeUnit As String, ByVal gradeCode As String, _
                            ByVal startDate As Date, ByRef pc As PeopleCalc, ByVal hoursToAdd As Double, _
                            ByVal overrideReason As String, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    nm = Trim$(nm): gradeCode = Trim$(gradeCode): crewTxt = NormCrew(crewTxt)
    If nm = "" Then MsgBox "Enter the name.", vbExclamation, "Add employee": Exit Function
    If InStr(nm, ";") > 0 Or InStr(nm, "|") > 0 Or InStr(nm, "/") > 0 Then
        MsgBox "A name can't contain  ;  |  or  /  (the calendar and the postings use them between names).", vbExclamation, "Add employee"
        Exit Function
    End If
    If RosterRowOf(nm) > 0 Then MsgBox nm & " is already on the roster.", vbExclamation, "Add employee": Exit Function
    If IsFormer(nm) Then
        MsgBox nm & " is on the Former sheet. Use Bring back instead, so their history comes back with them.", vbExclamation, "Add employee"
        Exit Function
    End If
    If DeptCol(homeUnit) = 0 Or gradeCode = "" Then MsgBox "Pick the home unit and enter the grade code.", vbExclamation, "Add employee": Exit Function
    If Not IsCrewCode(crewTxt) Then MsgBox "Pick the crew.", vbExclamation, "Add employee": Exit Function

    Dim ps As Date: ps = PeriodStartFor(pc.AsOf)
    If hoursToAdd <> 0 Then
        If Not ConfirmPeriod(ps) Then Exit Function
    End If
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String, r As Long, refId As String
    Dim codes() As String: ReDim codes(0 To UnitColLast() - UnitColFirst())
    codes(DeptCol(homeUnit) - UnitColFirst()) = gradeCode
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    Dim ownBefore As Double: ownBefore = LogTotalOnRoster(nm)     ' hours already in the log under this name (0 for a real new hire)
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    r = AddRosterRow(ws, nm, crewTxt, homeUnit, seniority, codes, Trim$(phone), Trim$(email))
    EnsureRosterFilter ws
    ProtectSheet ws
    If hoursToAdd <> 0 Then
        refId = NextRefID("NH")
        Dim data() As Variant: ReDim data(1 To 1, 1 To L_NCOLS)
        data(1, L_REF) = refId
        data(1, L_TYPE) = "Adjustment"
        data(1, L_DATE) = pc.AsOf
        data(1, L_DEPT) = homeUnit
        data(1, L_TASK) = "New hire"
        data(1, L_EMP) = nm
        data(1, L_CODE) = gradeCode
        data(1, L_EMPHRS) = 0
        data(1, L_REASON) = "New hire, start " & Format$(startDate, "mmm d, yyyy") & ": " & homeUnit & " average of " & pc.NPeers & _
                            " peer(s) [" & pc.PeerNames & "] as of " & Format$(pc.AsOf, "mmm d") & " = " & Format$(pc.Target, "0.0") & " h" & _
                            IIf(hoursToAdd <> pc.Target, " - OVERRIDDEN to " & Format$(hoursToAdd, "0.0") & " h: " & overrideReason, "")
        data(1, L_CHARGED) = hoursToAdd
        data(1, L_PERIOD) = ps
        data(1, L_USER) = CurrentUser()
        data(1, L_STAMP) = Now
        If Not AppendLogRows(data) Then Err.Raise vbObjectError + 712, , "the starting-hours entry couldn't be written to the OT Log"
    End If
    If Not AddCrewRow(nm, crewTxt, IIf(Int(startDate) < RosterYearFrom(), RosterYearFrom(), Int(startDate))) Then
        Err.Raise vbObjectError + 713, , "the crew history row couldn't be written"
    End If
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap(), nm)
    If Abs(LogChargedTotal() - logBefore - hoursToAdd) > 0.0001 Then diff = diff & "  OT Log total Hours Charged is off" & vbNewLine
    If PeriodColumn(ps) > 0 Or hoursToAdd = 0 Then
        If Abs(NumOr0(ws.Cells(r, C_REFUSED).Value2) - ownBefore - hoursToAdd) > 0.0001 Then
            diff = diff & "  " & nm & ": Refused shows " & NumOr0(ws.Cells(r, C_REFUSED).Value2) & ", expected " & (ownBefore + hoursToAdd) & vbNewLine
        End If
    End If
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "Adding " & nm & " did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Add employee"
        Exit Function
    End If
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = IIf(ownBefore <> 0, "NOTE: the OT Log already held " & HoursText(ownBefore) & " h under the name " & nm & _
                 " this year, and they now show on the new row as well. If this is a different person, rename one of them." & vbNewLine & vbNewLine, "") & _
             nm & " added on row " & r & " with " & Format$(hoursToAdd, "0.0") & " h" & _
             IIf(refId <> "", " (" & refId & " in the OT Log; Void Entry undoes the hours)", "") & "." & vbNewLine & _
             "Worked hours stay 0 until the next Kronos import: the name must be spelled the way Kronos spells it."
    SaveNewHire = True
    Exit Function
Failed:
    em = Err.Description
    LogError "SaveNewHire", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Adding " & nm & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Add employee"
End Function


' ====================================================================
'  MOVE TO UNIT  (UM-####)
' ====================================================================
' Works out the top-up for a full move. False = can't (pc.Warnings says why).
' allowFuture (v24.3, the Employee pop-up): a date after today is accepted and
' the result is a preview for a PLANNED move, worked out again on the day.
Public Function ComputeUnitMove(ByVal emp As String, ByVal newUnit As String, ByVal eff As Date, ByRef pc As PeopleCalc, _
                                Optional ByVal allowFuture As Boolean = False) As Boolean
    Dim blank As PeopleCalc: pc = blank
    If RosterRowOf(emp) = 0 Then pc.Warnings = "Pick the person.": Exit Function
    If DeptCol(newUnit) = 0 Then pc.Warnings = "Pick the new unit.": Exit Function
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    If Int(eff) > Date And Not allowFuture Then
        pc.Warnings = "The effective date can't be in the future: unit codes change every list from the moment they are entered. Enter the move on the day."
        Exit Function
    End If
    If Int(eff) < yearFrom Then
        pc.Warnings = "The effective date is before this roster year (" & Format$(yearFrom, "mmm d, yyyy") & "). That year is closed."
        Exit Function
    End If
    pc.Emp = Trim$(emp): pc.Unit = newUnit: pc.AsOf = Int(eff)

    Dim recs() As EmpRec, n As Long, i As Long, snap As Date, note As String, flagged As Boolean, found As Boolean
    On Error GoTo Failed
    BeginHoursBatch
    n = LoadEligible(newUnit, "ALL", recs)
    Dim grades() As Double, isMover() As Boolean, totals() As Double
    ReDim grades(1 To IIf(n > 0, n, 1)): ReDim isMover(1 To IIf(n > 0, n, 1)): ReDim totals(1 To IIf(n > 0, n, 1))
    For i = 1 To n
        grades(i) = recs(i).Grade
        isMover(i) = (StrComp(recs(i).EmpName, pc.Emp, vbTextCompare) = 0)
        If grades(i) <= PEER_MAX_GRADE And Not isMover(i) Then
            totals(i) = TotalHoursAsOf(recs(i).EmpName, pc.AsOf, snap, note)
            pc.NPeers = pc.NPeers + 1
            pc.PeerNames = pc.PeerNames & IIf(pc.PeerNames = "", "", ", ") & CalName(recs(i).EmpName) & " " & Format$(totals(i), "0.0")
            If note <> "" And Not flagged Then flagged = True: pc.Warnings = pc.Warnings & "Hours: " & note & ". "
        End If
    Next i
    pc.Target = HighestPeerTotal(grades, isMover, totals, n, found)
    pc.Own = TotalHoursAsOf(pc.Emp, pc.AsOf, snap, note)
    If note <> "" Then pc.Warnings = pc.Warnings & pc.Emp & ": " & note & ". "
    EndHoursBatch
    If found Then pc.AddHours = UnitMoveTopUp(pc.Own, pc.Target)
    If Not found Then
        pc.Warnings = pc.Warnings & "Nobody else in " & newUnit & " is at grade " & PEER_MAX_GRADE & " or below, so there is nothing to match. "
        pc.Summary = pc.Emp & " moves to " & newUnit & " on " & Format$(pc.AsOf, "mmm d, yyyy") & ". No hours are added."
    ElseIf pc.AddHours > 0 Then
        pc.Summary = pc.Emp & " moves to " & newUnit & " on " & Format$(pc.AsOf, "mmm d, yyyy") & ". Highest in " & newUnit & _
                     " (grade " & PEER_MAX_GRADE & " and below) is " & Format$(pc.Target, "0.0#") & " h; " & pc.Emp & " has " & _
                     Format$(pc.Own, "0.0#") & " h. Add " & Format$(pc.AddHours, "0.0#") & " h (" & Format$(pc.Own, "0.0#") & " -> " & _
                     Format$(pc.Own + pc.AddHours, "0.0#") & ")."
    Else
        pc.Summary = pc.Emp & " moves to " & newUnit & " on " & Format$(pc.AsOf, "mmm d, yyyy") & ". Highest in " & newUnit & _
                     " is " & Format$(pc.Target, "0.0#") & " h; " & pc.Emp & " already has " & Format$(pc.Own, "0.0#") & _
                     " h. No hours are added (hours are never lowered)."
    End If
    If pc.AsOf > Date Then
        pc.Summary = "PLANNED: " & pc.Summary & " This is today's estimate; the hours are worked out again on the day."
    End If
    ComputeUnitMove = True
    Exit Function
Failed:
    pc.Warnings = "The calculation stopped: " & Err.Description
    On Error Resume Next
    EndHoursBatch
End Function

' Clears the old code(s), sets the new one and the Home Unit, logs the grade
' changes, the crew change (if any) and, when the mover is below the target,
' ONE Adjustment (Ref UM-####, Reason "Unit move").
' clearUnits = the units whose code is cleared, "|" between them.
' batch (v24.3, a planned move being applied): no pay-period question, no
' save and no "can't do that" boxes - failTxt carries the reason instead.
' The check after the change, and its message, are the same either way.
Public Function SaveUnitMove(ByVal emp As String, ByVal newUnit As String, ByVal gradeCode As String, _
                             ByVal clearUnits As String, ByVal newCrew As String, ByRef pc As PeopleCalc, _
                             ByRef report As String, Optional ByVal batch As Boolean = False, _
                             Optional ByRef failTxt As String, Optional ByRef refOut As String) As Boolean
    failTxt = "": refOut = ""                          ' v25 (H4): the batch records the UM- ref
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim nRows As Long, r As Long
    r = RosterRowOf(emp, nRows)
    If r = 0 Then MoveStop batch, failTxt, emp & " isn't on the roster.": Exit Function
    If nRows > 1 Then MoveStop batch, failTxt, emp & " is on the roster " & nRows & " times. Remove the extra row first.": Exit Function
    gradeCode = Trim$(gradeCode)
    If DeptCol(newUnit) = 0 Or gradeCode = "" Then MoveStop batch, failTxt, "Pick the new unit and enter the grade code.": Exit Function
    If Not batch Then                       ' v24.3: one move at a time per person
        Dim had As String: had = PlannedFor(emp)
        If had <> "" Then
            MoveStop batch, failTxt, emp & " has a planned move waiting (" & had & "). Cancel it first: Needs Attention > Planned unit moves."
            Exit Function
        End If
    End If
    Dim ws As Worksheet: Set ws = RosterSheet()

    ' crew change (optional)
    Dim crewNow As String: crewNow = NormCrew(SafeStr(ws.Cells(r, C_SHIFT).Value))
    Dim crewTo As String: crewTo = NormCrew(newCrew)
    Dim crewChanges As Boolean, found As Boolean, since As Date
    If crewTo <> "" And crewTo <> crewNow Then
        If Not IsCrewCode(crewTo) Then MoveStop batch, failTxt, "'" & newCrew & "' isn't a crew.": Exit Function
        NewestCrew emp, found, since
        If found And since > pc.AsOf Then
            MoveStop batch, failTxt, emp & "'s crew was last changed on " & Format$(since, "mmm d, yyyy") & ", after this move's date. " & _
                     "Leave the crew box blank here and change the crew on the roster instead."
            Exit Function
        End If
        crewChanges = True
    End If

    Dim ps As Date: ps = PeriodStartFor(pc.AsOf)
    If pc.AddHours > 0 Then
        If batch Then
            If PeriodColumn(ps) = 0 Then
                failTxt = "its pay period (" & Format$(ps, "mmm d, yyyy") & ") isn't one of this roster's columns - run the year-end rollover first"
                Exit Function
            End If
        ElseIf Not ConfirmPeriod(ps) Then
            Exit Function
        End If
    End If
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim suppressWas As Boolean: suppressWas = gSuppressGradeLog
    Dim before As Object, logBefore As Double, iBefore As Double, em As String, diff As String, refId As String
    Dim u As Variant, cleared As String, postTxt As String
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    iBefore = NumOr0(ws.Cells(r, C_REFUSED).Value2)
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    gSuppressGradeLog = True
    UnlockSheet ws
    For Each u In Split(clearUnits, "|")
        If CStr(u) <> "" And StrComp(CStr(u), newUnit, vbTextCompare) <> 0 Then
            If DeptCol(CStr(u)) > 0 Then
                If Trim$(SafeStr(ws.Cells(r, DeptCol(CStr(u))).Value)) <> "" Then
                    ws.Cells(r, DeptCol(CStr(u))).ClearContents
                    cleared = cleared & "|" & CStr(u)
                End If
            End If
        End If
    Next u
    If cleared <> "" Then cleared = cleared & "|"
    ws.Cells(r, DeptCol(newUnit)).Value = gradeCode
    ws.Cells(r, C_HOME).Value = newUnit
    If crewChanges Then ws.Cells(r, C_SHIFT).Value = crewTo
    ProtectSheet ws
    gSuppressGradeLog = suppressWas
    RecordGradeRows r, r, CurrentUser()                 ' Grade Log rows for every code that changed
    If crewChanges Then
        If Not AddCrewRow(emp, crewTo, pc.AsOf) Then Err.Raise vbObjectError + 721, , "the crew history row couldn't be written"
    End If
    If pc.AddHours > 0 Then
        refId = NextRefID("UM")
        Dim data() As Variant: ReDim data(1 To 1, 1 To L_NCOLS)
        data(1, L_REF) = refId
        data(1, L_TYPE) = "Adjustment"
        data(1, L_DATE) = pc.AsOf
        data(1, L_DEPT) = newUnit
        data(1, L_TASK) = "Unit move"
        data(1, L_EMP) = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        data(1, L_CODE) = gradeCode
        data(1, L_EMPHRS) = pc.Own
        data(1, L_REASON) = "Unit move to " & newUnit & " on " & Format$(pc.AsOf, "mmm d, yyyy") & ": highest of " & pc.NPeers & _
                            " peer(s) [" & pc.PeerNames & "] = " & Format$(pc.Target, "0.0#") & " h; own " & Format$(pc.Own, "0.0#") & _
                            " h; added " & Format$(pc.AddHours, "0.0#") & " h"
        data(1, L_CHARGED) = pc.AddHours
        data(1, L_PERIOD) = ps
        data(1, L_USER) = CurrentUser()
        data(1, L_STAMP) = Now
        If Not AppendLogRows(data) Then Err.Raise vbObjectError + 722, , "the unit-move entry couldn't be written to the OT Log"
    End If
    If cleared <> "" Then postTxt = ExemptFromOpenPostings(emp, pc.AsOf, cleared, "moved " & Format$(pc.AsOf, "mmm d"))
    ClearCancelledMoveMarks emp, newUnit, pc.AsOf      ' v24.3: they hold the code now, so a cancelled move's mark no longer applies
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap(), emp)
    If Abs(LogChargedTotal() - logBefore - pc.AddHours) > 0.0001 Then diff = diff & "  OT Log total Hours Charged is off" & vbNewLine
    If PeriodColumn(ps) > 0 Or pc.AddHours = 0 Then
        If Abs(NumOr0(ws.Cells(r, C_REFUSED).Value2) - iBefore - pc.AddHours) > 0.0001 Then
            diff = diff & "  " & emp & ": Refused went from " & iBefore & " to " & NumOr0(ws.Cells(r, C_REFUSED).Value2) & _
                   ", expected + " & pc.AddHours & vbNewLine
        End If
    End If
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        failTxt = "the check after the change failed"
        MsgBox "The move did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Move to unit"
        Exit Function
    End If
    ' v24.3.1: open postings in the new unit for shifts on or after the move get them on the list
    ' (after the check, so the hours on the new row include the top-up; a planned mover is on them already)
    Dim addTxt As String, dropTxt As String
    SyncPersonOnPostings emp, newUnit, pc.AsOf, 0, True, False, False, addTxt, dropTxt
    AfterPeopleChange
    If Not batch Then SaveIfWritable
    refOut = refId
    report = emp & " moved to " & newUnit & " (" & gradeCode & ")." & vbNewLine & _
             IIf(pc.AddHours > 0, Format$(pc.AddHours, "0.0#") & " h added (" & refId & "; Void Entry undoes the hours).", "No hours added.") & _
             IIf(Trim$(pc.Warnings) <> "", vbNewLine & "Check: " & Trim$(pc.Warnings), "") & _
             IIf(crewChanges, vbNewLine & "Crew " & crewNow & " -> " & crewTo & " from " & Format$(pc.AsOf, "mmm d") & ".", "") & _
             IIf(postTxt <> "", vbNewLine & "Not chargeable on open posting(s): " & postTxt, "") & _
             IIf(addTxt <> "", vbNewLine & "Added to the list of open " & newUnit & " posting(s): " & addTxt & ". Print those sheets again.", "")
    SaveUnitMove = True
    Exit Function
Failed:
    em = Err.Description
    LogError "SaveUnitMove", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    failTxt = em
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    gSuppressGradeLog = suppressWas
    ProtectSheet ws
    ProtectAll
    MsgBox "The move stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Move to unit"
End Function

' A "can't do that" from SaveUnitMove: a box for the pop-up, text only for a batch
Private Sub MoveStop(ByVal batch As Boolean, ByRef failTxt As String, ByVal msg As String)
    failTxt = msg
    If Not batch Then MsgBox msg, vbExclamation, "Move to unit"
End Sub



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
Private Function CancelPlannedCore(ByVal id As String, ByVal reason As String, ByRef didTxt As String, _
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
                       "If it happens twice, send me this box. Backup: " & bk, vbCritical, TTL
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


' ====================================================================
'  RENAME  (every table together, verified)
' ====================================================================
' The tables and columns a person's full name lives in. Fills los / cols /
' labels (1..n) and returns n. Tables that don't exist yet are left out.
Private Function NameColumns(ByRef los() As ListObject, ByRef cols() As Long, ByRef labels() As String) As Long
    ReDim los(1 To 12): ReDim cols(1 To 12): ReDim labels(1 To 12)
    Dim n As Long
    AddNameCol los, cols, labels, n, TableOn(LOG_SHEET, LOG_TABLE), L_EMP, "OT Log"
    AddNameCol los, cols, labels, n, TableOn(POSTINGS_SHEET, PL_TABLE), PL_EMP, "Posting lists"
    AddNameCol los, cols, labels, n, AbsTable(False), AB_EMP, "Absences"
    AddNameCol los, cols, labels, n, SwitchTable(False), SW_OFF, "Shift switches (gives up)"
    AddNameCol los, cols, labels, n, SwitchTable(False), SW_WORKS, "Shift switches (works)"
    AddNameCol los, cols, labels, n, GradeLogTable(False), 2, "Grade Log"
    AddNameCol los, cols, labels, n, HoursHistoryTable(False), 4, "Hours History"
    AddNameCol los, cols, labels, n, CrewTable(False), CH_EMP, "Crew History"
    AddNameCol los, cols, labels, n, PlannedTable(False), PM_EMP, "Planned moves"     ' v24.3
    NameColumns = n
End Function

Private Sub AddNameCol(ByRef los() As ListObject, ByRef cols() As Long, ByRef labels() As String, ByRef n As Long, _
                       ByVal lo As ListObject, ByVal col As Long, ByVal label As String)
    If lo Is Nothing Then Exit Sub
    If col > lo.ListColumns.Count Then Exit Sub
    n = n + 1
    Set los(n) = lo
    cols(n) = col
    labels(n) = label
End Sub

' Cells in one table column holding nameTxt (trimmed; exact = same capitals too)
Private Function CountName(ByVal lo As ListObject, ByVal col As Long, ByVal nameTxt As String, _
                           Optional ByVal exact As Boolean = False) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.ListColumns(col).DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), nameTxt, IIf(exact, vbBinaryCompare, vbTextCompare)) = 0 Then CountName = CountName + 1
    Next i
End Function

' Rewrites the matching cells one by one (so a filter on the table can't
' hide a row from it). The sheet must be unprotected. Returns how many.
Private Function RenameColumn(ByVal lo As ListObject, ByVal col As Long, ByVal oldName As String, ByVal newName As String) As Long
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim rng As Range: Set rng = lo.ListColumns(col).DataBodyRange
    Dim v As Variant: v = RangeTo2D(rng)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), oldName, vbTextCompare) = 0 Then
            rng.Cells(i, 1).Value = newName
            RenameColumn = RenameColumn + 1
        End If
    Next i
End Function

' What ties a calendar line to one person (v24 rename). Two people can share a
' short calendar name ("Fred C"), so a line is theirs only when:
'   tied     its Source is a call-out / posting they ACCEPTED or were AWARDED,
'            or a shift switch that names them
'   absTied  it covers one of their own absences
Private Sub CalendarTies(ByVal emp As String, ByRef tied As Object, ByRef absTied As Object)
    Set tied = CreateObject("Scripting.Dictionary")
    tied.CompareMode = 1
    Set absTied = CreateObject("Scripting.Dictionary")
    absTied.CompareMode = 1
    Dim v As Variant, i As Long
    v = BodyOf(TableOn(LOG_SHEET, LOG_TABLE))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, L_EMP))), emp, vbTextCompare) = 0 And Trim$(SafeStr(v(i, L_REF))) <> "" Then
                Select Case SafeStr(v(i, L_OUTCOME))
                    Case "Accepted", "Awarded": tied(Trim$(SafeStr(v(i, L_REF)))) = True
                End Select
            End If
        Next i
    End If
    v = BodyOf(SwitchTable(False))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, SW_OFF))), emp, vbTextCompare) = 0 Or _
               StrComp(Trim$(SafeStr(v(i, SW_WORKS))), emp, vbTextCompare) = 0 Then tied(Trim$(SafeStr(v(i, SW_ID)))) = True
        Next i
    End If
    v = BodyOf(AbsTable(False))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, AB_EMP))), emp, vbTextCompare) = 0 And Trim$(SafeStr(v(i, AB_ID))) <> "" Then
                absTied(Trim$(SafeStr(v(i, AB_ID)))) = True
            End If
        Next i
    End If
End Sub

' Live calendar lines that mention the short name but are NOT tied to the
' person: the rename leaves them alone and says which ("" = none).
Private Function CalendarLinesLeftAlone(ByVal shortName As String, ByVal tied As Object, ByVal absTied As Object) As String
    Dim v As Variant: v = BodyOf(CovTable(False))
    If Not IsArray(v) Then Exit Function
    Dim i As Long, hits As Long, n As Long, s As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            If Not tied.Exists(Trim$(SafeStr(v(i, CV_SOURCE)))) And Not absTied.Exists(Trim$(SafeStr(v(i, CV_ABS)))) Then
                hits = 0
                ReplaceNameToken SafeStr(v(i, CV_BY)) & " / " & SafeStr(v(i, CV_NEED)), shortName, shortName, hits
                If hits > 0 Then
                    n = n + 1
                    If n <= 8 Then s = s & "  " & SafeStr(v(i, CV_ID)) & " " & IIf(IsDate(v(i, CV_DATE)), Format$(v(i, CV_DATE), "mmm d"), "") & vbNewLine
                End If
            End If
        End If
    Next i
    If n > 8 Then s = s & "  ...and " & (n - 8) & " more" & vbNewLine
    CalendarLinesLeftAlone = s
End Function

' What a rename would touch, for the confirm. False = not allowed (txt says why).
Public Function RenamePreview(ByVal oldName As String, ByVal newName As String, ByRef txt As String) As Boolean
    txt = ""
    oldName = Trim$(oldName): newName = Trim$(newName)
    Dim nRows As Long
    If RosterRowOf(oldName, nRows) = 0 Then txt = "Pick the person.": Exit Function
    If nRows > 1 Then txt = oldName & " is on the roster " & nRows & " times. Remove the extra row first.": Exit Function
    If newName = "" Then txt = "Enter the new name.": Exit Function
    If StrComp(oldName, newName, vbBinaryCompare) = 0 Then txt = "That is the same name.": Exit Function
    If InStr(newName, ";") > 0 Or InStr(newName, "|") > 0 Or InStr(newName, "/") > 0 Then
        txt = "A name can't contain  ;  |  or  /  (the calendar and the postings use them between names).": Exit Function
    End If
    Dim sameName As Boolean: sameName = (StrComp(oldName, newName, vbTextCompare) = 0)
    If Not sameName Then
        If RosterRowOf(newName) > 0 Then txt = newName & " is already on the roster.": Exit Function
        If IsFormer(newName) Then txt = newName & " is on the Former sheet.": Exit Function
    End If
    Dim los() As ListObject, cols() As Long, labels() As String, n As Long, i As Long, a As Long, b As Long, mergeTxt As String
    n = NameColumns(los, cols, labels)
    txt = oldName & "  ->  " & newName & vbNewLine & vbNewLine & "Rows that will change:" & vbNewLine & "  Roster: 1" & vbNewLine
    For i = 1 To n
        a = CountName(los(i), cols(i), oldName)
        If a > 0 Then txt = txt & "  " & labels(i) & ": " & a & vbNewLine
        If Not sameName Then
            b = CountName(los(i), cols(i), newName)
            If b > 0 Then mergeTxt = mergeTxt & "  " & labels(i) & ": " & b & vbNewLine
        End If
    Next i
    If CalName(oldName) <> CalName(newName) Then
        Dim tied As Object, absTied As Object, leftTxt As String
        CalendarTies oldName, tied, absTied
        leftTxt = CalendarLinesLeftAlone(CalName(oldName), tied, absTied)
        txt = txt & "  Calendar lines: '" & CalName(oldName) & "' becomes '" & CalName(newName) & "' on lines tied to their own call-outs, awards, switches and time off" & vbNewLine
        If leftTxt <> "" Then
            txt = txt & vbNewLine & "These calendar lines also say '" & CalName(oldName) & "' but aren't tied to this person (someone else with the same short name, or typed by hand). They will be LEFT ALONE:" & vbNewLine & leftTxt
        End If
    End If
    If mergeTxt <> "" Then
        txt = txt & vbNewLine & "CAREFUL: the name " & newName & " already has rows of its own. They would be combined with " & oldName & "'s:" & vbNewLine & mergeTxt
    End If
    txt = txt & vbNewLine & "Closed years (the OT Log archive sheets and Year End) keep the old name; the Grade Log gets a note." & vbNewLine & _
          "Kronos matches on the name: spell it the way Kronos does, or Worked hours stop updating for this person."
    RenamePreview = True
End Function

' The rename. A backup is saved first; counts per table, every roster total
' and the OT Log total are checked after; a mismatch blocks the save.
Public Function RenameEverywhere(ByVal oldName As String, ByVal newName As String, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    oldName = Trim$(oldName): newName = Trim$(newName)
    Dim chk As String
    If Not RenamePreview(oldName, newName, chk) Then MsgBox chk, vbExclamation, "Rename": Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(oldName)
    Dim sameName As Boolean: sameName = (StrComp(oldName, newName, vbTextCompare) = 0)
    Dim bk As String
    If Not PeopleBackup("rename", bk) Then Exit Function

    Dim los() As ListObject, cols() As Long, labels() As String, n As Long, i As Long
    Dim aBefore() As Long, bBefore() As Long, done As Long, total As Long
    n = NameColumns(los, cols, labels)
    ReDim aBefore(1 To IIf(n > 0, n, 1)): ReDim bBefore(1 To IIf(n > 0, n, 1))
    For i = 1 To n
        aBefore(i) = CountName(los(i), cols(i), oldName)
        If Not sameName Then bBefore(i) = CountName(los(i), cols(i), newName)
    Next i

    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String
    Dim tied As Object, absTied As Object, lo As ListObject, v As Variant, hits As Long, newTxt As String
    Dim oldShort As String, newShort As String, covDone As Long, covLeft As String, awardDone As Long
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    mBusy = True

    ' what ties a calendar line to this person (read BEFORE the rename)
    CalendarTies oldName, tied, absTied

    ' the name columns
    For i = 1 To n
        If aBefore(i) > 0 Then
            UnlockSheet los(i).Parent
            done = RenameColumn(los(i), cols(i), oldName, newName)
            ProtectSheet los(i).Parent
            total = total + done
        End If
    Next i

    ' postings: the "Awarded To" text
    Set lo = TableOn(POSTINGS_SHEET, PT_TABLE)
    v = BodyOf(lo)
    If IsArray(v) Then
        UnlockSheet lo.Parent
        For i = 1 To UBound(v, 1)
            hits = 0
            newTxt = RenameInList(SafeStr(v(i, PS_AWARDTO)), oldName, newName, hits)
            If hits > 0 Then lo.DataBodyRange.Cells(i, PS_AWARDTO).Value = newTxt: awardDone = awardDone + 1
        Next i
        ProtectSheet lo.Parent
    End If

    ' calendar lines: the short name, only on lines tied to this person
    oldShort = CalName(oldName): newShort = CalName(newName)
    If StrComp(oldShort, newShort, vbBinaryCompare) <> 0 Then
        covLeft = CalendarLinesLeftAlone(oldShort, tied, absTied)
        Set lo = CovTable(False)
        v = BodyOf(lo)
        If IsArray(v) Then
            UnlockSheet lo.Parent
            For i = 1 To UBound(v, 1)
                If UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
                    ' "Covered By" is theirs when the line came from their own call-out, award or switch
                    If tied.Exists(Trim$(SafeStr(v(i, CV_SOURCE)))) Then
                        hits = 0
                        newTxt = ReplaceNameToken(SafeStr(v(i, CV_BY)), oldShort, newShort, hits)
                        If hits > 0 Then lo.DataBodyRange.Cells(i, CV_BY).Value = newTxt: covDone = covDone + 1
                    End If
                    ' "Off / Need" is theirs on a switch of theirs, or a line covering their own time off
                    If Left$(UCase$(Trim$(SafeStr(v(i, CV_SOURCE)))), 3) = "SS-" And tied.Exists(Trim$(SafeStr(v(i, CV_SOURCE)))) _
                       Or absTied.Exists(Trim$(SafeStr(v(i, CV_ABS)))) Then
                        hits = 0
                        newTxt = ReplaceNameToken(SafeStr(v(i, CV_NEED)), oldShort, newShort, hits)
                        If hits > 0 Then lo.DataBodyRange.Cells(i, CV_NEED).Value = newTxt: covDone = covDone + 1
                    End If
                End If
            Next i
            ProtectSheet lo.Parent
        End If
    End If

    ' the roster
    UnlockSheet ws
    ws.Cells(r, C_NAME).Value = newName
    ProtectSheet ws
    Application.Calculation = calcWas
    Application.Calculate

    ' ---- verify ----
    For i = 1 To n
        If sameName Then
            If CountName(los(i), cols(i), newName, True) <> aBefore(i) Then diff = diff & "  " & labels(i) & ": not every row took the new spelling" & vbNewLine
        Else
            If CountName(los(i), cols(i), oldName) <> 0 Then diff = diff & "  " & labels(i) & ": the old name is still there" & vbNewLine
            If CountName(los(i), cols(i), newName) <> aBefore(i) + bBefore(i) Then
                diff = diff & "  " & labels(i) & ": " & CountName(los(i), cols(i), newName) & " rows under the new name, expected " & _
                       (aBefore(i) + bBefore(i)) & vbNewLine
            End If
        End If
    Next i
    If Not sameName Then
        If before.Exists(oldName) Then
            before(newName) = before(oldName)
            before.Remove oldName
        End If
    End If
    ' (when the new name already had OT Log rows of its own, the two people's hours are now one row: only the others are compared)
    If bBefore(1) = 0 Then
        diff = diff & TotalsDifference(before, RosterTotalsMap())
    Else
        diff = diff & TotalsDifference(before, RosterTotalsMap(), newName)
    End If
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "The rename did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Rename"
        Exit Function
    End If
    LogRename oldName, newName
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = "Renamed to " & newName & "." & vbNewLine & total & " row(s) updated in the tables" & _
             IIf(awardDone > 0, ", " & awardDone & " posting award(s)", "") & IIf(covDone > 0, ", " & covDone & " calendar line(s)", "") & _
             ". Roster totals and the OT Log are unchanged." & _
             IIf(covLeft <> "", vbNewLine & vbNewLine & "These calendar lines mention '" & oldShort & "' but aren't tied to this person, so they were left alone:" & vbNewLine & covLeft, "") & _
             vbNewLine & vbNewLine & "Publish the members' file again so it shows the new name." & vbNewLine & "Backup: " & bk
    RenameEverywhere = True
    Exit Function
Failed:
    em = Err.Description
    LogError "RenameEverywhere", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "The rename stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Rename"
End Function


' ====================================================================
'  REMOVE / BRING BACK  (the Former sheet)
' ====================================================================
' Moves the person's roster row to the Former sheet as values. Their OT Log,
' absence, switch and posting rows stay. A backup is saved first; every other
' row's totals and the OT Log total are checked after.
Public Function RemoveEmployee(ByVal emp As String, ByVal leftOn As Date, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim nRows As Long, r As Long
    r = RosterRowOf(emp, nRows)
    If r = 0 Then MsgBox emp & " isn't on the roster.", vbExclamation, "Remove": Exit Function
    If nRows > 1 Then
        MsgBox emp & " is on the roster " & nRows & " times. Delete the extra row by hand first (IntegrityCheck shows which)." & ByHandHint(RosterSheet().Name), vbExclamation, "Remove"
        Exit Function
    End If
    If Int(leftOn) > Date Then
        MsgBox "Remove takes the person off every list for every date, so it is done on or after their last day." & vbNewLine & _
               "Until then, book them off (Book Time Off).", vbExclamation, "Remove"
        Exit Function
    End If
    emp = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
    Dim bk As String
    If Not PeopleBackup("remove", bk) Then Exit Function

    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, after As Object, logBefore As Double, em As String, diff As String, postTxt As String
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To FM_N)
    Dim c As Long, logHrs As Double
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    logHrs = LogTotalOnRoster(emp)
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    ' v24.3: a unit move still waiting is cancelled first, while the row is here
    ' (it may have dated a crew change, and the Former row keeps the crew)
    Dim planId As String, planTxt As String
    planId = PlannedFor(emp)
    If planId <> "" Then
        If Not CancelPlannedCore(planId, "removed from the roster, left " & Format$(leftOn, "mmm d, yyyy") & _
                                 " (" & CurrentUser() & ", " & Format$(Date, "mmm d") & ")", planTxt, False) Then
            Err.Raise vbObjectError + 732, , planTxt
        End If
        planTxt = vbNewLine & "Their planned move " & planId & " was cancelled."
    End If
    rec(1, FM_EMP) = emp
    rec(1, FM_LEFT) = Int(leftOn)
    rec(1, FM_CREW) = Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value))
    rec(1, FM_HOME) = Trim$(SafeStr(ws.Cells(r, C_HOME).Value))
    rec(1, FM_SENIORITY) = ws.Cells(r, C_SENIORITY).Value
    For c = UnitColFirst() To UnitColLast()
        rec(1, FM_CODE1 + c - UnitColFirst()) = Trim$(SafeStr(ws.Cells(r, c).Value))
    Next c
    rec(1, FM_TOTAL) = NumOr0(ws.Cells(r, C_TOTAL).Value2)
    rec(1, FM_WORKED) = NumOr0(ws.Cells(r, C_WORKED).Value2)
    rec(1, FM_REFUSED) = NumOr0(ws.Cells(r, C_REFUSED).Value2)
    rec(1, FM_PROV) = NumOr0(ws.Cells(r, C_PROV).Value2)
    rec(1, FM_YEAR) = RosterYear()
    rec(1, FM_USER) = CurrentUser()
    rec(1, FM_STAMP) = Now
    rec(1, FM_STATUS) = "Former"
    rec(1, FM_PAYROLL) = PayrollText(ws.Cells(r, C_PAYROLL).Value)
    If Not AppendFormerRow(rec) Then Err.Raise vbObjectError + 731, , "the Former sheet couldn't be written"

    UnlockSheet ws
    If ws.FilterMode Then ws.ShowAllData
    ws.Rows(r).Delete Shift:=xlUp
    EnsureRosterFilter ws
    ProtectSheet ws
    Application.Calculate
    Set after = RosterTotalsMap()
    diff = TotalsDifference(before, after, emp)
    If after.Exists(emp) Then diff = diff & "  " & emp & " is still on the roster" & vbNewLine
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    If Not IsFormer(emp) Then diff = diff & "  " & emp & " isn't on the Former sheet" & vbNewLine
    If diff <> "" Then
        QuietOff evWas, scrWas, calcWas
        MarkNoSave
        MsgBox "Removing " & emp & " did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Remove"
        Exit Function
    End If
    postTxt = ExemptFromOpenPostings(emp, Int(leftOn), "", "left " & Format$(leftOn, "mmm d"))
    QuietOff evWas, scrWas, calcWas
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = emp & " is on the Former sheet (left " & Format$(leftOn, "mmm d, yyyy") & ") and off the roster." & vbNewLine & _
             "Final figures kept: Total " & HoursText(rec(1, FM_TOTAL)) & ", Worked " & HoursText(rec(1, FM_WORKED)) & _
             ", Refused " & HoursText(rec(1, FM_REFUSED)) & ". Their " & HoursText(logHrs) & " h in the OT Log stay there." & _
             IIf(postTxt <> "", vbNewLine & "Not chargeable on open posting(s): " & postTxt, "") & planTxt & vbNewLine & _
             "Everyone else's totals are unchanged. Publish the members' file again." & vbNewLine & "Backup: " & bk
    RemoveEmployee = True
    Exit Function
Failed:
    em = Err.Description
    LogError "RemoveEmployee", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "Removing " & emp & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Remove"
End Function

' Puts a former person back on the roster with the same formulas as everyone
' else, so their OT Log hours for this year show again.
Public Function BringBack(ByVal emp As String, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim fv As Variant, fr As Long
    fr = FormerRowOf(emp, fv)
    If fr = 0 Then MsgBox emp & " isn't on the Former sheet.", vbExclamation, "Bring back": Exit Function
    emp = Trim$(SafeStr(fv(fr, FM_EMP)))
    If RosterRowOf(emp) > 0 Then MsgBox emp & " is already on the roster.", vbExclamation, "Bring back": Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim bk As String
    If Not PeopleBackup("bring back", bk) Then Exit Function

    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String, r As Long, c As Long
    Dim codes() As String: ReDim codes(0 To UnitColLast() - UnitColFirst())
    Dim lo As ListObject, found As Boolean, crewTxt As String, newest As String, postTxt As String
    For c = UnitColFirst() To UnitColLast()
        codes(c - UnitColFirst()) = Trim$(SafeStr(fv(fr, FM_CODE1 + c - UnitColFirst())))
    Next c
    crewTxt = NormCrew(SafeStr(fv(fr, FM_CREW)))
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    Dim payNo As String
    If UBound(fv, 2) >= FM_PAYROLL Then payNo = PayrollText(fv(fr, FM_PAYROLL))
    r = AddRosterRow(ws, emp, crewTxt, SafeStr(fv(fr, FM_HOME)), fv(fr, FM_SENIORITY), codes, "", "", payNo)
    EnsureRosterFilter ws
    ProtectSheet ws
    Set lo = FormerTable(False)
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(fr, FM_STATUS).Value = "Returned"
    lo.DataBodyRange.Cells(fr, FM_BACK).Value = Date
    ProtectSheet lo.Parent
    newest = NewestCrew(emp, found)
    If Not found Or newest <> crewTxt Then
        If Not AddCrewRow(emp, crewTxt, Date) Then Err.Raise vbObjectError + 741, , "the crew history row couldn't be written"
    End If
    postTxt = ClearLeftExemptions(emp)                   ' chargeable again on open postings
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap(), emp)
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    If Abs(NumOr0(ws.Cells(r, C_REFUSED).Value2) - LogTotalOnRoster(emp)) > 0.0001 Then
        diff = diff & "  " & emp & ": Refused shows " & NumOr0(ws.Cells(r, C_REFUSED).Value2) & ", the OT Log has " & LogTotalOnRoster(emp) & vbNewLine
    End If
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "Bringing " & emp & " back did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Bring back"
        Exit Function
    End If
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = emp & " is back on the roster (row " & r & ") with " & HoursText(NumOr0(ws.Cells(r, C_REFUSED).Value2)) & _
             " h charged this year from the OT Log." & vbNewLine & _
             "Type their phone and email in again (they aren't kept on the Former sheet). Worked hours fill in at the next Kronos import." & _
             IIf(postTxt <> "", vbNewLine & "Chargeable again on open posting(s): " & postTxt, "") & _
             vbNewLine & "Backup: " & bk
    BringBack = True
    Exit Function
Failed:
    em = Err.Description
    LogError "BringBack", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "Bringing " & emp & " back stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Bring back"
End Function


' ====================================================================
'  LAUNCHER (ADMIN > Employee)
' ====================================================================
Public Sub ShowEmployeeForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        MsgBox "Run Alt+F8 > SetupPeople once first (it adds the Crew History and Former sheets).", vbExclamation, "Employee"
        Exit Sub
    End If
    frmEmployee.Show
    Exit Sub
Oops:
    ReportError "ShowEmployeeForm"
End Sub

' The old Rename Employee macro: the pop-up opens on Rename with that person picked
Public Sub ShowEmployeeRename(ByVal emp As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        MsgBox "Run Alt+F8 > SetupPeople once first (it adds the Crew History and Former sheets).", vbExclamation, "Employee"
        Exit Sub
    End If
    frmEmployee.StartRename emp
    frmEmployee.Show
End Sub


' ====================================================================
'  ONE-TIME SETUP (Alt+F8 > SetupPeople; safe to run again)
' ====================================================================
Public Sub SetupPeople()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If RosterLastRow() < R_FIRST Then MsgBox "No employees found on the roster.", vbExclamation, "Setup": Exit Sub
    If MsgBox("Set up the v24 people sheets?" & vbNewLine & vbNewLine & _
              "- Crew History: one row per person, their crew from " & Format$(RosterYearFrom(), "mmm d, yyyy") & vbNewLine & _
              "- Former: empty, for people taken off the roster with Employee > Remove" & vbNewLine & _
              "- Postings: a Signup Hours column on the end of the posted lists (partial sign-ups)" & vbNewLine & vbNewLine & _
              "Nothing on the roster or in the OT Log changes. A backup copy is saved first and every roster total is checked after.", _
              vbOKCancel + vbQuestion, "Setup People") <> vbOK Then Exit Sub
    Dim bk As String
    If Not PeopleBackup("v24 setup", bk) Then Exit Sub

    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim em As String, nSeed As Long, diff As String
    Dim before As Object, logBefore As Double
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    mBusy = True
    If CrewTable(True) Is Nothing Then Err.Raise vbObjectError + 701, , "the Crew History sheet couldn't be created"
    nSeed = SeedCrewHistory()
    If nSeed < 0 Then Err.Raise vbObjectError + 702, , "the crew rows couldn't be written"
    If FormerTable(True) Is Nothing Then Err.Raise vbObjectError + 703, , "the Former sheet couldn't be created"
    Dim whyNot As String
    If Not UpgradePostingList(whyNot) Then Err.Raise vbObjectError + 704, , whyNot
    mBusy = False
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap())
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    Application.ScreenUpdating = scrWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "The setup did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Setup People"
        Exit Sub
    End If
    SaveIfWritable
    MsgBox "People sheets are set up." & vbNewLine & vbNewLine & _
           "Crew History: " & nSeed & " row(s) added" & IIf(nSeed = 0, " (everyone already had one)", "") & "." & vbNewLine & _
           "Former sheet and the posting lists' Signup Hours column: ready." & vbNewLine & _
           "Roster totals and the OT Log are unchanged." & vbNewLine & vbNewLine & "Backup: " & bk, vbInformation, "Setup People"
    Exit Sub
Failed:
    em = Err.Description
    LogError "SetupPeople", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    On Error Resume Next
    mBusy = False
    Application.ScreenUpdating = scrWas
    MsgBox "The setup stopped part-way: " & em & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Setup People"
End Sub
