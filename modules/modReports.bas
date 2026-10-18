Attribute VB_Name = "modReports"
Option Explicit

' ====================================================================
'  REFUSED HOURS BY KIND (v24.1)
'
'  Refused (roster column I) stays ONE figure: it drives Total Hours and
'  call order. This module only shows what it is made of:
'     call-out   charges on OT Log rows of Type "Callout"
'     posting    charges on rows of Type "Posting"
'     other      everything else (adjustments: ADJ-, LV-, NH-, UM-;
'                opening balance; values moved in from typed cells)
'  Same test as the roster's own formulas: not VOID, Period Start inside
'  this roster year. call-out + posting + other = Refused, always.
'
'  Where it shows (decided Oct 12, 2026: no new roster columns):
'   - Excel's status bar (bottom left) for the roster row that is selected
'   - the members' file (modPublish adds the three figures to dRoster)
'   - RefusalsReport: a printable sheet, per person and per unit
'
'  Pure (mirrored in tests/rules.py, checked by SelfTest): RefusedBucket,
'  SplitTotals.
' ====================================================================

Public Const REP_VERSION As String = "Reports v25.3 (Oct 17)"

' what SplitTotals keeps per person: a Variant array with these indexes
Public Const SP_CALLOUT As Long = 0     ' hours
Public Const SP_POSTING As Long = 1
Public Const SP_OTHER As Long = 2
Public Const SP_NCALLOUT As Long = 3    ' how many charged rows
Public Const SP_NPOSTING As Long = 4

Private mSplit As Object                ' name -> the array above, for the status bar
Private mSplitStamp As String
Private mBarShown As Boolean            ' True while the status bar holds our text


' ====================================================================
'  PURE RULES
' ====================================================================
' Which part of Refused a log row's charge belongs to
Public Function RefusedBucket(ByVal typ As String) As String
    Select Case LCase$(Trim$(typ))
        Case "callout": RefusedBucket = "callout"
        Case "posting": RefusedBucket = "posting"
        Case Else: RefusedBucket = "other"
    End Select
End Function

' lg = rows in the OT Log layout. firstStart / lastStart = the first and
' last Period Start of the roster year, as whole numbers. Returns
' name -> Array(call-out h, posting h, other h, call-out rows, posting rows).
Public Function SplitTotals(ByRef lg As Variant, ByVal firstStart As Long, ByVal lastStart As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set SplitTotals = d
    If Not IsArray(lg) Then Exit Function
    Dim i As Long, emp As String, h As Double, per As Long, a As Variant
    For i = 1 To UBound(lg, 1)
        emp = Trim$(SafeStr(lg(i, L_EMP)))
        If emp <> "" And UCase$(Trim$(SafeStr(lg(i, L_STATUS)))) <> "VOID" Then
            per = PeriodNumber(lg(i, L_PERIOD))
            h = NumOr0(lg(i, L_CHARGED))
            If h <> 0 And per >= firstStart And per <= lastStart Then
                If d.Exists(emp) Then
                    a = d(emp)
                Else
                    a = Array(0#, 0#, 0#, 0&, 0&)
                End If
                Select Case RefusedBucket(SafeStr(lg(i, L_TYPE)))
                    Case "callout"
                        a(SP_CALLOUT) = a(SP_CALLOUT) + h
                        If h > 0 Then a(SP_NCALLOUT) = a(SP_NCALLOUT) + 1
                    Case "posting"
                        a(SP_POSTING) = a(SP_POSTING) + h
                        If h > 0 Then a(SP_NPOSTING) = a(SP_NPOSTING) + 1
                    Case Else
                        a(SP_OTHER) = a(SP_OTHER) + h
                End Select
                d(emp) = a
            End If
        End If
    Next i
End Function

' A Period Start cell as a whole number (0 = not a date)
Private Function PeriodNumber(ByVal v As Variant) As Long
    If IsError(v) Then Exit Function
    If IsDate(v) Then
        PeriodNumber = CLng(Int(CDbl(CDate(v))))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        PeriodNumber = CLng(Int(CDbl(v)))
    End If
End Function


' ====================================================================
'  THE LIVE FIGURES
' ====================================================================
' The OT Log table without any message (Nothing if it isn't there)
Private Function QuietLog() As ListObject
    If Not SheetExists(LOG_SHEET) Then Exit Function
    On Error Resume Next
    Set QuietLog = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo 0
End Function

' Everyone's split for this roster year, straight from the OT Log
Public Function SplitNow() As Object
    Dim lo As ListObject: Set lo = QuietLog()
    Dim lg As Variant
    If Not lo Is Nothing Then
        If LogLayoutCurrent(lo) Then
            If Not lo.DataBodyRange Is Nothing Then lg = RangeTo2D(lo.DataBodyRange)
        End If
    End If
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim sRow As Long: sRow = PStartRow()
    Set SplitNow = SplitTotals(lg, CLng(ws.Cells(sRow, C_PERIOD1).Value2), CLng(ws.Cells(sRow, PeriodLastCol()).Value2))
End Function

' v24.2: the same split for part of the year (period start numbers, both ends included)
Public Function SplitBetween(ByVal firstStart As Long, ByVal lastStart As Long) As Object
    Dim lo As ListObject: Set lo = QuietLog()
    Dim lg As Variant
    If Not lo Is Nothing Then
        If LogLayoutCurrent(lo) Then
            If Not lo.DataBodyRange Is Nothing Then lg = RangeTo2D(lo.DataBodyRange)
        End If
    End If
    Set SplitBetween = SplitTotals(lg, firstStart, lastStart)
End Function

' v24.2: which pay periods the report covers. The box opens on "1-<current period>"
' (year to date). False = cancelled.
Private Function AskReportRange(ByRef firstStart As Long, ByRef lastStart As Long, ByRef rangeTxt As String, _
                                ByRef wholeYear As Boolean) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim nPer As Long: nPer = PeriodLastCol() - C_PERIOD1 + 1
    Dim base As Long: base = CLng(ws.Cells(PStartRow(), C_PERIOD1).Value2)
    Dim cur As Long: cur = PeriodColumn(PeriodStartFor(Date)) - C_PERIOD1 + 1
    If cur < 1 Or cur > nPer Then cur = nPer
    Dim a As Variant, t As String, dash As Long, pFrom As Long, pTo As Long, s1 As String, s2 As String
    Do
        a = Application.InputBox("Pay periods to include, as  from-to  (1 to " & nPer & ")." & vbNewLine & vbNewLine & _
                                 "Period " & cur & " is the current one, so  1-" & cur & "  is the year to date." & vbNewLine & _
                                 "For the whole year enter  1-" & nPer & ".", "Refusals Report", "1-" & cur, Type:=2)
        If VarType(a) = vbBoolean Then Exit Function
        t = Replace(Trim$(CStr(a)), " ", "")
        pFrom = 0: pTo = 0
        dash = InStr(t, "-")
        If dash > 1 Then
            s1 = Left$(t, dash - 1): s2 = Mid$(t, dash + 1)
        Else
            s1 = t: s2 = t
        End If
        If Len(s1) >= 1 And Len(s1) <= 2 And Len(s2) >= 1 And Len(s2) <= 2 Then
            If s1 Like String$(Len(s1), "#") And s2 Like String$(Len(s2), "#") Then pFrom = CLng(s1): pTo = CLng(s2)
        End If
        If pFrom >= 1 And pTo >= pFrom And pTo <= nPer Then Exit Do
        MsgBox "Enter two period numbers with a dash between them, for example  1-" & cur & ".  They run from 1 to " & nPer & ".", _
               vbExclamation, "Refusals Report"
    Loop
    firstStart = base + 14 * (pFrom - 1)
    lastStart = base + 14 * (pTo - 1)
    wholeYear = (pFrom = 1 And pTo = nPer)
    rangeTxt = "pay periods " & pFrom & " to " & pTo & " (" & Format$(CDate(firstStart), "mmm d, yyyy") & " to " & _
               Format$(CDate(lastStart + 13), "mmm d, yyyy") & ")"
    AskReportRange = True
End Function

' One person's split. False when they have nothing charged this year.
Public Function RefusedSplitOf(ByVal splitMap As Object, ByVal emp As String, ByRef callout As Double, _
                               ByRef posting As Double, ByRef other As Double) As Boolean
    callout = 0: posting = 0: other = 0
    If splitMap Is Nothing Then Exit Function
    If Not splitMap.Exists(Trim$(emp)) Then Exit Function
    Dim a As Variant: a = splitMap(Trim$(emp))
    callout = a(SP_CALLOUT): posting = a(SP_POSTING): other = a(SP_OTHER)
    RefusedSplitOf = True
End Function

' "call-out 42 / posting 21" (+ " / other 12.3" only when it isn't zero)
Public Function SplitText(ByVal callout As Double, ByVal posting As Double, ByVal other As Double) As String
    SplitText = "call-out " & HoursText(callout) & " / posting " & HoursText(posting)
    If Abs(other) > 0.0001 Then SplitText = SplitText & " / other " & HoursText(other)
End Function

' The cached map for the status bar: rebuilt when the log changes (row
' count, total charged or number of voids) or the roster year moves.
Private Function CachedSplit() As Object
    Dim lo As ListObject: Set lo = QuietLog()
    If lo Is Nothing Then Exit Function
    Dim stamp As String
    If Not lo.DataBodyRange Is Nothing Then
        stamp = lo.ListRows.Count & "|" & _
                Application.WorksheetFunction.Sum(lo.ListColumns(L_CHARGED).DataBodyRange) & "|" & _
                Application.WorksheetFunction.CountIf(lo.ListColumns(L_STATUS).DataBodyRange, "VOID")
    End If
    stamp = stamp & "|" & CStr(RosterSheet().Cells(PStartRow(), C_PERIOD1).Value2)
    If mSplit Is Nothing Or stamp <> mSplitStamp Then
        Set mSplit = SplitNow()
        mSplitStamp = stamp
    End If
    Set CachedSplit = mSplit
End Function

' A rename changes names without changing any hours: forget the cache
Public Sub ResetSplitCache(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    Set mSplit = Nothing
    mSplitStamp = ""
End Sub


' ====================================================================
'  STATUS BAR: the split for the selected roster row
' ====================================================================
' From ThisWorkbook.Workbook_SheetSelectionChange. Writes nothing to the
' workbook (a cell write on every click would clear Undo), only Excel's
' status bar. Must never raise or slow a click.
Public Sub RosterSelectionChanged(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Quit
    Dim txt As String, nm As String, c As Double, p As Double, o As Double
    If Not Target Is Nothing Then
        If Sh.Name = RosterSheet().Name And Target.Rows.Count = 1 And Target.Row >= R_FIRST Then
            nm = Trim$(SafeStr(Sh.Cells(Target.Row, C_NAME).Value))
            If nm <> "" Then
                RefusedSplitOf CachedSplit(), nm, c, p, o
                txt = nm & ":   Refused " & HoursText(c + p + o) & " h  =  " & SplitText(c, p, o)
            End If
        End If
    End If
    If txt <> "" Then
        Application.StatusBar = txt
        mBarShown = True
    ElseIf mBarShown Then
        Application.StatusBar = False
        mBarShown = False
    End If
    Exit Sub
Quit:
    On Error Resume Next
    If mBarShown Then Application.StatusBar = False: mBarShown = False
End Sub

' Gives Excel's status bar back if it holds our text. Called when the workbook
' loses focus or closes, and by macros that switch tabs with events off.
Public Sub ReleaseStatusBar(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    On Error Resume Next
    If mBarShown Then Application.StatusBar = False: mBarShown = False
End Sub

' Leaving the roster tab: give the status bar back
Public Sub ReportsSheetActivated(ByVal Sh As Object)
    On Error Resume Next
    If Not mBarShown Then Exit Sub
    If Sh.Name <> RosterSheet().Name Then Application.StatusBar = False: mBarShown = False
End Sub


' ====================================================================
'  REFUSALS REPORT (PRINT / PUBLISH > Refusals Report)
' ====================================================================
' One row per person with charged hours this roster year: call-out and
' posting refusals (hours and how many), other, total; a sub-total per
' Home Unit. Drawn on the output-only "Refusal Report" sheet. No notes,
' no names of who logged anything.
Public Sub RefusalsReport()
    If Not RequireAdmin() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Dim unit As String, allUnits As Boolean
    If Not AskReportUnit(unit, allUnits) Then Exit Sub
    Dim firstStart As Long, lastStart As Long, rangeTxt As String, wholeYear As Boolean
    If Not AskReportRange(firstStart, lastStart, rangeTxt, wholeYear) Then Exit Sub

    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim ws As Worksheet, em As String
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- who, and their home unit ----
    Dim splitMap As Object: Set splitMap = SplitBetween(firstStart, lastStart)
    Dim homeOf As Object: Set homeOf = CreateObject("Scripting.Dictionary")
    homeOf.CompareMode = 1
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim v As Variant, i As Long, nm As String
    If lastRow >= R_FIRST Then
        v = rs.Range(rs.Cells(R_FIRST, 1), rs.Cells(lastRow, C_NAME)).Value        ' A:D, always 2-D
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, C_NAME)))
            If nm <> "" Then homeOf(nm) = Trim$(SafeStr(v(i, C_HOME)))
        Next i
    End If
    Dim n As Long, k As Variant, hu As String
    Dim names() As String, units() As String, keys() As String
    ReDim names(1 To splitMap.Count + 1): ReDim units(1 To splitMap.Count + 1): ReDim keys(1 To splitMap.Count + 1)
    For Each k In splitMap.Keys
        If homeOf.Exists(k) Then
            hu = homeOf(k)
            If hu = "" Then hu = "(no home unit)"
        Else
            hu = "(not on the roster)"
        End If
        If allUnits Or StrComp(hu, unit, vbTextCompare) = 0 Then
            n = n + 1
            names(n) = CStr(k): units(n) = hu
            keys(n) = LCase$(hu) & "|" & LCase$(CStr(k))
        End If
    Next k
    ' sort by unit, then name
    Dim j As Long, tK As String, tN As String, tU As String
    For i = 2 To n
        tK = keys(i): tN = names(i): tU = units(i): j = i - 1
        Do While j >= 1
            If keys(j) <= tK Then Exit Do
            keys(j + 1) = keys(j): names(j + 1) = names(j): units(j + 1) = units(j)
            j = j - 1
        Loop
        keys(j + 1) = tK: names(j + 1) = tN: units(j + 1) = tU
    Next i

    ' ---- the sheet ----
    Set ws = ReportSheet()
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    Const NC As Long = 8
    ws.Range("A1").Value = "REFUSED HOURS BY KIND  -  " & IIf(allUnits, "all units", unit) & "  -  " & RosterYear()
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = "Charged hours in the OT Log for " & rangeTxt & ", not void.  " & _
                           IIf(wholeYear, "Call-out + Posting + Other = Refused on the roster.  ", _
                               "Part of the year: the totals are less than Refused on the roster for anyone charged outside these periods.  ") & _
                           "Drawn " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & "."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    ws.Range("A4").Resize(1, NC).Value = Array("Home Unit", "Employee", "Call-out hours", "Call-out refusals", _
                                               "Posting hours", "Posting refusals", "Other hours", "Total refused")
    With ws.Range("A4").Resize(1, NC)
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(30, 56, 100)
        .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    ws.Rows(4).RowHeight = 30

    Dim r As Long: r = 5
    Dim a As Variant, sub1(0 To 4) As Double, tot(0 To 4) As Double, q As Long, curUnit As String
    For i = 1 To n
        If i > 1 And units(i) <> curUnit Then
            ReportTotalRow ws, r, curUnit & " total", sub1
            r = r + 1
            For q = 0 To 4
                sub1(q) = 0
            Next q
        End If
        curUnit = units(i)
        a = splitMap(names(i))
        ws.Cells(r, 1).Value = units(i)
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = names(i)
        ws.Cells(r, 3).Value = a(SP_CALLOUT)
        ws.Cells(r, 4).Value = a(SP_NCALLOUT)
        ws.Cells(r, 5).Value = a(SP_POSTING)
        ws.Cells(r, 6).Value = a(SP_NPOSTING)
        ws.Cells(r, 7).Value = a(SP_OTHER)
        ws.Cells(r, 8).Value = a(SP_CALLOUT) + a(SP_POSTING) + a(SP_OTHER)
        For q = 0 To 4
            sub1(q) = sub1(q) + a(q): tot(q) = tot(q) + a(q)
        Next q
        r = r + 1
    Next i
    If n > 0 Then
        ReportTotalRow ws, r, curUnit & " total", sub1
        r = r + 1
        If allUnits Then ReportTotalRow ws, r, "ALL UNITS", tot: r = r + 1
    Else
        ws.Cells(r, 1).Value = "Nobody has charged hours in these pay periods" & IIf(allUnits, ".", " in " & unit & ".")
        r = r + 1
    End If

    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 28
    ws.Range(ws.Columns(3), ws.Columns(NC)).ColumnWidth = 12
    ws.Range(ws.Cells(5, 3), ws.Cells(r, NC)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(4, 1), ws.Cells(r - 1, NC)).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, NC)).Address
        .PrintTitleRows = "$4:$4"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo Failed
    ProtectSheet ws
    ws.Activate
    ReleaseStatusBar                        ' the tab changed with events off
    ActiveWindow.DisplayGridlines = False
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas

    If n > 0 Then
        If MsgBox("The report is on the '" & ws.Name & "' tab (" & n & " people)." & vbNewLine & vbNewLine & _
                  "Save a PDF as well?", vbYesNo + vbQuestion, "Refusals report") = vbYes Then ReportPdf ws, IIf(allUnits, "all units", unit)
    End If
    Exit Sub
Failed:
    em = Err.Description
    LogError "RefusalsReport", Err.Number, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    MsgBox "The refusals report couldn't be drawn: " & em, vbExclamation, "Refusals report"
End Sub

Private Sub ReportTotalRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByRef t() As Double)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 3).Value = t(SP_CALLOUT)
    ws.Cells(r, 4).Value = t(SP_NCALLOUT)
    ws.Cells(r, 5).Value = t(SP_POSTING)
    ws.Cells(r, 6).Value = t(SP_NPOSTING)
    ws.Cells(r, 7).Value = t(SP_OTHER)
    ws.Cells(r, 8).Value = t(SP_CALLOUT) + t(SP_POSTING) + t(SP_OTHER)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Font.Bold = True
        .Interior.Color = RGB(236, 239, 243)
    End With
End Sub

Private Function ReportSheet() As Worksheet
    If Not SheetExists(REFUSAL_SHEET) Then
        Dim was As Object: Set was = ActiveSheet
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = REFUSAL_SHEET
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)                ' v25.1
        On Error Resume Next
        If Not was Is Nothing Then was.Activate
        On Error GoTo 0
    End If
    Set ReportSheet = ThisWorkbook.Worksheets(REFUSAL_SHEET)
End Function

' Numbered pick: 0 = all units, 1..8 = one unit. False = cancelled.
Private Function AskReportUnit(ByRef unit As String, ByRef allUnits As Boolean) As Boolean
    Dim depts As Variant: depts = DeptList()
    Dim i As Long, menu As String, s As String, k As Long
    menu = "  0 = All units" & vbNewLine
    For i = LBound(depts) To UBound(depts)
        menu = menu & "  " & (i - LBound(depts) + 1) & " = " & depts(i) & vbNewLine
    Next i
    Do
        s = Trim$(VBA.InputBox("Refused hours by kind (call-out / posting), this roster year." & vbNewLine & vbNewLine & _
                               "Which unit? Enter the number:" & vbNewLine & vbNewLine & menu, "Refusals report", "0"))
        If s = "" Then Exit Function
        k = -1
        If s Like "#" Or s Like "##" Then k = CLng(s)        ' digits only: "$5" or "1,5" must not pick a unit
        If k >= 0 Then
            If k = 0 Then allUnits = True: AskReportUnit = True: Exit Function
            If k >= 1 And k <= UBound(depts) - LBound(depts) + 1 Then
                unit = CStr(depts(LBound(depts) + k - 1))
                AskReportUnit = True
                Exit Function
            End If
        End If
        MsgBox "Enter a number from 0 to " & (UBound(depts) - LBound(depts) + 1) & ".", vbExclamation, "Refusals report"
    Loop
End Function

Private Sub ReportPdf(ByVal ws As Worksheet, ByVal what As String)
    Dim folder As String: folder = EnsureSlash(ReportFolder())
    If Not FolderExists(folder) Then folder = EnsureSlash(ThisWorkbook.Path)
    Dim pdfPath As String: pdfPath = folder & "Refused hours " & what & " " & Format$(Date, "yyyy-mm-dd") & ".pdf"
    Dim em As String
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then em = Err.Description
    On Error GoTo 0
    If em = "" Then
        MsgBox "PDF saved:" & vbNewLine & pdfPath, vbInformation, "Refusals report"
    Else
        MsgBox "The PDF couldn't be saved (" & em & "). Is it open?", vbExclamation, "Refusals report"
    End If
End Sub
