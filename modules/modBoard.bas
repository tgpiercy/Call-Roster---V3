Attribute VB_Name = "modBoard"
Option Explicit
' v27.1 (the refactor, Oct 2026): the two-week board (BoardPrint), from modCalendar; moved unchanged (tests/movecheck.py proves it)
Public Const BOARD_VERSION As String = "Board v27.1 (Oct 22)"

' ====================================================================
'  THE TWO-WEEK BOARD (v26, kit item 3)
' ====================================================================
' Today .. today + 13, one page per calendar group (the named groups; "All
' units" would repeat them), landscape: one row per date with the crews on,
' the Days lines and the Nights lines as the calendar shows them (OPEN in
' red). Drawn from SlotsFor over the caches on the output-only Board sheet,
' saved as a PDF beside the month calendars, then "print as well?".
' Supervisors and admins. Reads only; writes the Board sheet. (BOARD_SHEET / BOARD_DAYS
' are declared at the module head: a Const after a procedure does not compile.)
Public Sub BoardPrint()
    On Error GoTo Oops
    If Not IsSupervisor() Then
        MsgBox "The board is printed by a supervisor or an admin.", vbInformation, "Board"
        Exit Sub
    End If
    Dim ws As Worksheet
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim calcWas As Long: calcWas = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Failed
    Set ws = BoardSheet()                           ' review: made with the screen and events off (no flash, no CalActivated)
    LoadCache
    LoadPostCache
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    On Error Resume Next
    ws.ResetAllPageBreaks
    On Error GoTo Failed
    With ws.Cells.Font
        .Name = CAL_FONT: .Size = 9
    End With
    ws.Columns(1).ColumnWidth = 15
    ws.Columns(2).ColumnWidth = 15
    ws.Columns(3).ColumnWidth = 58
    ws.Columns(4).ColumnWidth = 58

    Dim d0 As Date: d0 = Date
    Dim g As Variant, parts As Variant, grp As String, r As Long, dd As Long, d As Date, nOpenAll As Long
    Dim slots() As CalSlot, n As Long, i As Long, txtD As String, txtN As String, openD As Boolean, openN As Boolean
    Dim firstRow As Long, lastGroupStart As Long
    r = 1
    For Each g In CalGroups()
        parts = Split(g, "|")
        If parts(1) <> "*" Then
            grp = CStr(parts(0))
            If r > 1 Then
                On Error Resume Next
                ws.Rows(r).PageBreak = xlPageBreakManual      ' one page per group
                On Error GoTo Failed
            End If
            lastGroupStart = r
            With ws.Cells(r, 1)
                .Value = UCase$(grp) & "  -  BOARD  " & Format$(d0, "ddd mmm d") & " to " & Format$(d0 + BOARD_DAYS - 1, "ddd mmm d")
                .Font.Name = CAL_FONT_B: .Font.Size = 14: .Font.Color = thHead()
            End With
            r = r + 1
            ws.Cells(r, 1).Value = "Drawn " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & ".  OPEN = nobody covers it yet; NCN = no coverage needed; " & _
                                   "POSTED = waiting for sign-ups.  The calendar tab has the details and the options.  " & TrustStrip()   ' v26.5.5 (5.2)
            ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = thMuted()
            r = r + 1
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Value = Array("Date", "Crews on", "DAYS  7 AM - 7 PM", "NIGHTS  7 PM - 7 AM")
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 4))
                .Font.Name = CAL_FONT_B: .Font.Color = vbWhite: .Interior.Color = thHead()
                .VerticalAlignment = xlCenter: .IndentLevel = 1
            End With
            ws.Rows(r).RowHeight = 18
            r = r + 1
            firstRow = r
            For dd = 0 To BOARD_DAYS - 1
                d = d0 + dd
                n = SlotsFor(d, "Days", grp, slots, True)
                txtD = BoardLines(slots, n, openD)
                n = SlotsFor(d, "Nights", grp, slots, True)
                txtN = BoardLines(slots, n, openN)
                ws.Cells(r, 1).NumberFormat = "@"
                ws.Cells(r, 1).Value = Format$(d, "ddd mmm d") & IIf(HolidayLabel(d) <> "", vbLf & HolidayLabel(d), "")
                ws.Cells(r, 2).Value = DaysCrew(d) & " days" & vbLf & NightsCrew(d) & " nights"
                ws.Cells(r, 3).Value = IIf(txtD = "", "-", txtD)
                ws.Cells(r, 4).Value = IIf(txtN = "", "-", txtN)
                If openD Then ws.Cells(r, 3).Font.Color = thOpen(): ws.Cells(r, 3).Font.Name = CAL_FONT_B: nOpenAll = nOpenAll + 1
                If openN Then ws.Cells(r, 4).Font.Color = thOpen(): ws.Cells(r, 4).Font.Name = CAL_FONT_B: nOpenAll = nOpenAll + 1
                If dd = 0 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.Color = thToday()
                If Weekday(d) = vbSaturday Or Weekday(d) = vbSunday Then ws.Cells(r, 1).Font.Color = thMuted()
                r = r + 1
            Next dd
            With ws.Range(ws.Cells(firstRow, 1), ws.Cells(r - 1, 4))
                .WrapText = True: .VerticalAlignment = xlTop: .IndentLevel = 1
                .Borders(xlInsideHorizontal).Color = thGrid(): .Borders(xlEdgeBottom).Color = thGrid()
            End With
            ws.Range(ws.Rows(firstRow), ws.Rows(r - 1)).AutoFit
            r = r + 1
        End If
    Next g
    If lastGroupStart = 0 Then ws.Cells(1, 1).Value = "No calendar groups are set up (CalGroups in modCalendar)."

    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 4)).Address
        .PrintTitleRows = ""
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .PrintGridlines = False
    End With
    Application.PrintCommunication = True
    On Error GoTo Failed
    ProtectSheet ws
    ws.Activate
    ReleaseStatusBar
    ActiveWindow.DisplayGridlines = False
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas

    ' the PDF beside the month calendars, then the printer
    Dim folder As String, subF As String, fName As String, msg As String
    folder = EnsureSlash(ReportFolder())
    If FolderExists(folder) Then
        subF = folder & "Calendars" & Application.PathSeparator
        If Not FolderExists(subF) Then
            On Error Resume Next
            MkDir subF
            On Error GoTo Failed
        End If
        If FolderExists(subF) Then folder = subF
    Else
        folder = EnsureSlash(ThisWorkbook.Path)
    End If
    fName = "Board " & Format$(d0, "yyyy-mm-dd") & ".pdf"
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & fName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then msg = "PDF NOT saved: " & Err.Description Else msg = "PDF saved: " & folder & fName
    Err.Clear
    On Error GoTo Failed
    If MsgBox(msg & vbNewLine & vbNewLine & "The board is on the '" & BOARD_SHEET & "' tab: " & _
              "one page per unit group, " & BOARD_DAYS & " days, " & nOpenAll & " shift(s) with an open line." & vbNewLine & vbNewLine & _
              "Print it as well?", vbYesNo + vbQuestion, "Board") = vbYes Then
        On Error Resume Next
        ws.PrintOut Copies:=1
        If Err.Number <> 0 Then MsgBox "NOT printed: " & Err.Description, vbExclamation, "Board"
        Err.Clear
        On Error GoTo Failed
    End If
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    LogError "BoardPrint", Err.Number, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    MsgBox "The board couldn't be drawn: " & em, vbExclamation, "Board"
    Exit Sub
Oops:
    ReportError "BoardPrint"
End Sub

' One shift's lines as the board shows them, one per line: "Smith P (Vac)  -  Jones K", "... - OPEN",
' "... - NO TAKERS", "... - POSTED P-0012", "... - NCN". anyOpen = an OPEN or NO TAKERS line among them.
Private Function BoardLines(ByRef slots() As CalSlot, ByVal n As Long, ByRef anyOpen As Boolean) As String
    Dim i As Long, t As String, rt As String, lt As String
    anyOpen = False
    For i = 1 To n
        rt = slots(i).RightText
        If rt = "" Then
            rt = "OPEN": anyOpen = True
        ElseIf UCase$(rt) Like "NO TAKERS*" Then
            rt = "NO TAKERS": anyOpen = True
        End If
        ' (review, D-M1) the board hangs where members read it: a sick call or a personal day prints as "Off"
        ' (PublicLeft, as the members' file), and a line whose free text names a kind of leave is not printed as written
        lt = PublicLeft(slots(i).LeftText, slots(i).TypeCode)
        If LeaveWordsIn(lt) <> "" Then lt = "(a line the board can't show as written - see the calendar)"
        If LeaveWordsIn(rt) <> "" Then rt = "(see the calendar)"
        t = t & IIf(t = "", "", vbLf) & lt & "  -  " & rt
    Next i
    BoardLines = t
End Function

' The Board sheet, made on first use (output only; protected; shown when drawn)
Private Function BoardSheet() As Worksheet
    If Not SheetExists(BOARD_SHEET) Then
        Dim was As Object: Set was = ActiveSheet
        Dim ws As Worksheet
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = BOARD_SHEET
        RelockStructure
        SetTabColor ws, thHead()
        On Error Resume Next
        If Not was Is Nothing Then was.Activate
        On Error GoTo 0
    End If
    Set BoardSheet = ThisWorkbook.Worksheets(BOARD_SHEET)
End Function


