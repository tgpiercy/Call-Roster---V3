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
'  ReplaceNameToken.
' ====================================================================

Public Const PPL_VERSION As String = "People v24 (Jan 5)"
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
Public Const FM_N As Long = 22

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

Private mHist As Variant            ' the table's rows (Empty = no rows)
Private mIdx As Object              ' employee -> Collection of row numbers in mHist
Private mHistLoaded As Boolean
Private mTestMode As Boolean        ' SelfTest feeds mHist from memory
Private mBusy As Boolean            ' our own writes to the roster fire change events


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
            CrewOn = CrewOnCore(mHist, nm, d, "", rowList)
            Exit Function
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
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CH_SHEET
            ws.Tab.Color = RGB(96, 125, 139)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        ws.Unprotect Password:=SHEET_PASSWORD
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
        ws.Visible = IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
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
    ws.Unprotect Password:=SHEET_PASSWORD
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

' Workbook_Open (admins, writable): anyone new on the roster gets their
' first row; a crew changed while macros were off is dated today. Quiet.
Public Sub PeopleOnOpen()
    If ThisWorkbook.ReadOnly Or Not IsAdmin() Then Exit Sub
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
            MsgBox "That isn't a date I can read. Try  Jan 12  or  1/12/2027.", vbExclamation, "Crew change"
        ElseIf notBefore > 0 And d < notBefore Then
            MsgBox "The date must be " & Format$(notBefore, "mmm d, yyyy") & " or later: that is when the crew being replaced started." & _
                   vbNewLine & "To correct an earlier row, edit it on the Crew History sheet.", vbExclamation, "Crew change"
        Else
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
    ws.Unprotect Password:=SHEET_PASSWORD
    For i = 1 To nChg
        ws.Cells(r1 + chgIdx(i) - 1, C_SHIFT).Value = chgOld(i)
    Next i
    ProtectSheet ws
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  TYPED DATES
' ====================================================================
' Reads a typed date (no time part). With no year typed ("Jan 12", "1/12"),
' a date more than 90 days ago means next year - typing January in December.
' More than ~13 months from today is refused. Same rule as the Shift Switch
' pop-up's date box.
Public Function ParseTypedDate(ByVal t As String, ByRef d As Date) As Boolean
    t = Trim$(t)
    If t = "" Then Exit Function
    If Not IsDate(t) Then Exit Function
    Dim dd As Date: dd = Int(CDate(t))
    If Not TypedHasYear(t) Then
        If dd < Date - 90 Then dd = DateAdd("yyyy", 1, dd)
    End If
    If Abs(dd - Date) > 400 Then Exit Function
    d = dd
    ParseTypedDate = True
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
' (1/12/27), or a month name plus two numbers (Jan 12 27).
Private Function TypedHasYear(ByVal t As String) As Boolean
    Dim i As Long, ch As String, tok As String
    Dim nNum As Long, hasAlpha As Boolean, bigNum As Boolean
    t = t & " "
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
    bk = EnsureSlash(ThisWorkbook.Path) & "Backup before " & what & " " & Format$(Now, "yyyy-mm-dd hhmmss") & " - " & ThisWorkbook.Name
    Dim e As String
    On Error Resume Next
    ThisWorkbook.SaveCopyAs bk
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
    If IsNumeric(s) Then PlainNum = Format$(CDbl(s), "0.##") Else PlainNum = s
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
    ResetCalCacheAll
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
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = FM_SHEET
            ws.Tab.Color = RGB(96, 125, 139)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        ws.Unprotect Password:=SHEET_PASSWORD
        Dim hdr() As Variant: ReDim hdr(1 To FM_N)
        Dim c As Long
        hdr(FM_EMP) = "Employee": hdr(FM_LEFT) = "Left On": hdr(FM_CREW) = "Crew": hdr(FM_HOME) = "Home Unit"
        hdr(FM_SENIORITY) = "Plant Seniority"
        For c = UnitColFirst() To UnitColLast()
            hdr(FM_CODE1 + c - UnitColFirst()) = DeptOfCol(c)
        Next c
        hdr(FM_TOTAL) = "Total Hours": hdr(FM_WORKED) = "Worked": hdr(FM_REFUSED) = "Refused": hdr(FM_PROV) = "Provisional"
        hdr(FM_YEAR) = "Roster Year": hdr(FM_USER) = "Removed By": hdr(FM_STAMP) = "Timestamp"
        hdr(FM_STATUS) = "Status": hdr(FM_BACK) = "Returned On"
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
        ws.Visible = IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set FormerTable = lo
End Function

Private Sub FormatFormer(ByVal ws As Worksheet)
    ws.Columns(FM_LEFT).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_BACK).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_SENIORITY).NumberFormat = "mmm d, yyyy"
    ws.Columns(FM_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ws.Columns(FM_CREW).NumberFormat = "@"
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
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    ws.Unprotect Password:=SHEET_PASSWORD
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
                              ByVal email As String) As Long
    Dim r As Long: r = RosterLastRow() + 1
    If r < R_FIRST Then r = R_FIRST
    ws.Unprotect Password:=SHEET_PASSWORD
    If ws.FilterMode Then ws.ShowAllData
    If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST))) > 0 Then
        Err.Raise vbObjectError + 711, , "row " & r & " (under the last name on the roster) isn't empty - clear it and try again"
    End If
    If r > R_FIRST Then
        ws.Range(ws.Cells(r - 1, 1), ws.Cells(r - 1, C_LAST)).Copy
        ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST)).PasteSpecial Paste:=xlPasteFormats
        ws.Range(ws.Cells(r, 1), ws.Cells(r, C_LAST)).PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False
        ws.Rows(r).RowHeight = ws.Rows(r - 1).RowHeight
    End If
    ws.Cells(r, C_NAME).Interior.ColorIndex = xlNone        ' the Kronos import colours matched names
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
        If Abs(NumOr0(ws.Cells(r, C_REFUSED).Value2) - hoursToAdd) > 0.0001 Then
            diff = diff & "  " & nm & ": starting hours show " & NumOr0(ws.Cells(r, C_REFUSED).Value2) & ", expected " & hoursToAdd & vbNewLine
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
    report = nm & " added on row " & r & " with " & Format$(hoursToAdd, "0.0") & " h" & _
             IIf(refId <> "", " (" & refId & " in the OT Log; Void Entry undoes the hours)", "") & "." & vbNewLine & _
             "Worked hours stay 0 until the next Kronos import: the name must be spelled the way Kronos spells it."
    SaveNewHire = True
    Exit Function
Failed:
    em = Err.Description
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectSheet ws
    MarkNoSave
    MsgBox "Adding " & nm & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Add employee"
End Function


' ====================================================================
'  MOVE TO UNIT  (UM-####)
' ====================================================================
' Works out the top-up for a full move. False = can't (pc.Warnings says why).
Public Function ComputeUnitMove(ByVal emp As String, ByVal newUnit As String, ByVal eff As Date, ByRef pc As PeopleCalc) As Boolean
    Dim blank As PeopleCalc: pc = blank
    If RosterRowOf(emp) = 0 Then pc.Warnings = "Pick the person.": Exit Function
    If DeptCol(newUnit) = 0 Then pc.Warnings = "Pick the new unit.": Exit Function
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    If Int(eff) > Date Then
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
Public Function SaveUnitMove(ByVal emp As String, ByVal newUnit As String, ByVal gradeCode As String, _
                             ByVal clearUnits As String, ByVal newCrew As String, ByRef pc As PeopleCalc, _
                             ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim nRows As Long, r As Long
    r = RosterRowOf(emp, nRows)
    If r = 0 Then MsgBox emp & " isn't on the roster.", vbExclamation, "Move to unit": Exit Function
    If nRows > 1 Then MsgBox emp & " is on the roster " & nRows & " times. Remove the extra row first.", vbExclamation, "Move to unit": Exit Function
    gradeCode = Trim$(gradeCode)
    If DeptCol(newUnit) = 0 Or gradeCode = "" Then MsgBox "Pick the new unit and enter the grade code.", vbExclamation, "Move to unit": Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()

    ' crew change (optional)
    Dim crewNow As String: crewNow = NormCrew(SafeStr(ws.Cells(r, C_SHIFT).Value))
    Dim crewTo As String: crewTo = NormCrew(newCrew)
    Dim crewChanges As Boolean, found As Boolean, since As Date
    If crewTo <> "" And crewTo <> crewNow Then
        If Not IsCrewCode(crewTo) Then MsgBox "'" & newCrew & "' isn't a crew.", vbExclamation, "Move to unit": Exit Function
        NewestCrew emp, found, since
        If found And since > pc.AsOf Then
            MsgBox emp & "'s crew was last changed on " & Format$(since, "mmm d, yyyy") & ", after this move's date. " & _
                   "Leave the crew box blank here and change the crew on the roster instead.", vbExclamation, "Move to unit"
            Exit Function
        End If
        crewChanges = True
    End If

    Dim ps As Date: ps = PeriodStartFor(pc.AsOf)
    If pc.AddHours > 0 Then
        If Not ConfirmPeriod(ps) Then Exit Function
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
    ws.Unprotect Password:=SHEET_PASSWORD
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
        MsgBox "The move did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Move to unit"
        Exit Function
    End If
    AfterPeopleChange
    SaveIfWritable
    report = emp & " moved to " & newUnit & " (" & gradeCode & ")." & vbNewLine & _
             IIf(pc.AddHours > 0, Format$(pc.AddHours, "0.0#") & " h added (" & refId & "; Void Entry undoes the hours).", "No hours added.") & _
             IIf(crewChanges, vbNewLine & "Crew " & crewNow & " -> " & crewTo & " from " & Format$(pc.AsOf, "mmm d") & ".", "") & _
             IIf(postTxt <> "", vbNewLine & "Not chargeable on open posting(s): " & postTxt, "")
    SaveUnitMove = True
    Exit Function
Failed:
    em = Err.Description
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    gSuppressGradeLog = suppressWas
    ProtectSheet ws
    MarkNoSave
    MsgBox "The move stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving and reopen.", vbCritical, "Move to unit"
End Function


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
        txt = txt & "  Calendar lines: '" & CalName(oldName) & "' becomes '" & CalName(newName) & "' on lines tied to their call-outs, postings and switches" & vbNewLine
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
    Dim tied As Object, lo As ListObject, v As Variant, hits As Long, newTxt As String
    Dim oldShort As String, newShort As String, covDone As Long, covLeft As String, awardDone As Long
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    mBusy = True

    ' what ties a calendar line to this person: their OT Log refs and their switches (read BEFORE the rename)
    Set tied = CreateObject("Scripting.Dictionary")
    tied.CompareMode = 1
    v = BodyOf(TableOn(LOG_SHEET, LOG_TABLE))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, L_EMP))), oldName, vbTextCompare) = 0 And Trim$(SafeStr(v(i, L_REF))) <> "" Then tied(Trim$(SafeStr(v(i, L_REF)))) = True
        Next i
    End If
    v = BodyOf(SwitchTable(False))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, SW_OFF))), oldName, vbTextCompare) = 0 Or _
               StrComp(Trim$(SafeStr(v(i, SW_WORKS))), oldName, vbTextCompare) = 0 Then tied(Trim$(SafeStr(v(i, SW_ID)))) = True
        Next i
    End If

    ' the name columns
    For i = 1 To n
        If aBefore(i) > 0 Then
            los(i).Parent.Unprotect Password:=SHEET_PASSWORD
            done = RenameColumn(los(i), cols(i), oldName, newName)
            ProtectSheet los(i).Parent
            total = total + done
        End If
    Next i

    ' postings: the "Awarded To" text
    Set lo = TableOn(POSTINGS_SHEET, PT_TABLE)
    v = BodyOf(lo)
    If IsArray(v) Then
        lo.Parent.Unprotect Password:=SHEET_PASSWORD
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
        Set lo = CovTable(False)
        v = BodyOf(lo)
        If IsArray(v) Then
            lo.Parent.Unprotect Password:=SHEET_PASSWORD
            For i = 1 To UBound(v, 1)
                If UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
                    If tied.Exists(Trim$(SafeStr(v(i, CV_SOURCE)))) Then
                        hits = 0
                        newTxt = ReplaceNameToken(SafeStr(v(i, CV_BY)), oldShort, newShort, hits)
                        If hits > 0 Then lo.DataBodyRange.Cells(i, CV_BY).Value = newTxt
                        covDone = covDone + IIf(hits > 0, 1, 0)
                        hits = 0
                        newTxt = ReplaceNameToken(SafeStr(v(i, CV_NEED)), oldShort, newShort, hits)
                        If hits > 0 Then lo.DataBodyRange.Cells(i, CV_NEED).Value = newTxt
                        covDone = covDone + IIf(hits > 0, 1, 0)
                    Else
                        hits = 0
                        ReplaceNameToken SafeStr(v(i, CV_BY)) & " / " & SafeStr(v(i, CV_NEED)), oldShort, newShort, hits
                        If hits > 0 And Len(covLeft) < 300 Then covLeft = covLeft & "  " & SafeStr(v(i, CV_ID)) & " " & _
                            IIf(IsDate(v(i, CV_DATE)), Format$(v(i, CV_DATE), "mmm d"), "") & vbNewLine
                    End If
                End If
            Next i
            ProtectSheet lo.Parent
        End If
    End If

    ' the roster
    ws.Unprotect Password:=SHEET_PASSWORD
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
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MarkNoSave
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
        MsgBox emp & " is on the roster " & nRows & " times. Delete the extra row by hand first (IntegrityCheck shows which).", vbExclamation, "Remove"
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
    If Not AppendFormerRow(rec) Then Err.Raise vbObjectError + 731, , "the Former sheet couldn't be written"

    ws.Unprotect Password:=SHEET_PASSWORD
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
             "Final figures kept: Total " & Format$(rec(1, FM_TOTAL), "0.##") & ", Worked " & Format$(rec(1, FM_WORKED), "0.##") & _
             ", Refused " & Format$(rec(1, FM_REFUSED), "0.##") & ". Their " & Format$(logHrs, "0.##") & " h in the OT Log stay there." & _
             IIf(postTxt <> "", vbNewLine & "Not chargeable on open posting(s): " & postTxt, "") & vbNewLine & _
             "Everyone else's totals are unchanged. Publish the members' file again." & vbNewLine & "Backup: " & bk
    RemoveEmployee = True
    Exit Function
Failed:
    em = Err.Description
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectSheet ws
    MarkNoSave
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
    Dim lo As ListObject, found As Boolean, crewTxt As String, newest As String
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
    r = AddRosterRow(ws, emp, crewTxt, SafeStr(fv(fr, FM_HOME)), fv(fr, FM_SENIORITY), codes, "", "")
    EnsureRosterFilter ws
    ProtectSheet ws
    Set lo = FormerTable(False)
    lo.Parent.Unprotect Password:=SHEET_PASSWORD
    lo.DataBodyRange.Cells(fr, FM_STATUS).Value = "Returned"
    lo.DataBodyRange.Cells(fr, FM_BACK).Value = Date
    ProtectSheet lo.Parent
    newest = NewestCrew(emp, found)
    If Not found Or newest <> crewTxt Then
        If Not AddCrewRow(emp, crewTxt, Date) Then Err.Raise vbObjectError + 741, , "the crew history row couldn't be written"
    End If
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
    report = emp & " is back on the roster (row " & r & ") with " & Format$(NumOr0(ws.Cells(r, C_REFUSED).Value2), "0.##") & _
             " h charged this year from the OT Log." & vbNewLine & _
             "Type their phone and email in again (they aren't kept on the Former sheet). Worked hours fill in at the next Kronos import." & _
             vbNewLine & "Backup: " & bk
    BringBack = True
    Exit Function
Failed:
    em = Err.Description
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectSheet ws
    MarkNoSave
    MsgBox "Bringing " & emp & " back stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Bring back"
End Function


' ====================================================================
'  LAUNCHER (ADMIN > Employee)
' ====================================================================
Public Sub ShowEmployeeForm()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        MsgBox "Run Alt+F8 > SetupPeople once first (it adds the Crew History and Former sheets).", vbExclamation, "Employee"
        Exit Sub
    End If
    frmEmployee.Show
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
    On Error Resume Next
    mBusy = False
    Application.ScreenUpdating = scrWas
    MarkNoSave
    MsgBox "The setup stopped part-way: " & em & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Setup People"
End Sub
