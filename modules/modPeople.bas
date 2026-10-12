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
'  Pure (mirrored in tests/rules.py, checked by SelfTest): CrewOnCore.
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
'  ONE-TIME SETUP (Alt+F8 > SetupPeople; safe to run again)
' ====================================================================
Public Sub SetupPeople()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If RosterLastRow() < R_FIRST Then MsgBox "No employees found on the roster.", vbExclamation, "Setup": Exit Sub
    If MsgBox("Set up the v24 people sheets?" & vbNewLine & vbNewLine & _
              "- Crew History: one row per person, their crew from " & Format$(RosterYearFrom(), "mmm d, yyyy") & vbNewLine & vbNewLine & _
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
