Attribute VB_Name = "modReconcile"
Option Explicit

' ====================================================================
'  RECONCILE (v27, item 4 - briefs/v27.md; review 5.3): the paper call
'  order beside the roster's, for the parallel run of the alpha
'
'  Paper runs alongside the roster for at least one pay period (Graeme,
'  Oct 20). On the admin-only Reconcile sheet the admin picks a unit and
'  the date the paper roster was good for, types the paper's call order
'  (names, top to bottom) once, and clicks Compare: the roster's order for
'  that unit is drawn beside it (the crews and a planned move as of that
'  day; Total Hours as of now) - the SAME list the printed
'  call-out roster and the call-out pop-up use (LoadEligible with asOf, every
'  grade, every crew) - and every position that differs gets the roster's
'  reason: where the roster puts the paper's person and which of the three
'  keys decided it (grade code, then Total Hours, then seniority). A
'  Decision (Paper wrong / Rule missing / Data error) and a Note are typed
'  per line. Save PDF files one page per period beside the master for the
'  Nov 15 pack. Nothing here changes any data; the sheet is never published.
'  The typed cells (unit, date, paper names, decisions, notes) are the only
'  unlocked ones, for an admin (LockReconcileCells), as the Min Staff sheet's.
'  ReconcileDiffCore is the one pure rule (suite "reconcile").
'  "From the old log" (the legacy Entry Log replayed as the paper) waits
'  for that sheet's layout (not in the repo): a report-back question.
' ====================================================================

Public Const RC_VERSION As String = "Reconcile v29.0.1 (Oct 28)"     ' stamped with the release at release prep (v27)
Public Const RECON_SHEET As String = "Reconcile"
Private Const RC_FIRST As Long = 9                   ' the first position row
Private Const RC_POSITIONS As Long = 60              ' rows RC_FIRST .. RC_FIRST + 59
Private Const RC_UNIT As String = "C4"
Private Const RC_DATE As String = "C5"
Private Const RC_PAPER As String = "C6"
Private Const RC_SUMMARY As String = "A7"
Private Const RC_DECISIONS As String = "Paper wrong,Rule missing,Data error,Agreed"

' ====================================================================
'  THE PURE RULE (mirrored in tests/rules.py)
' ====================================================================
' paper and roster: names top to bottom, ";" between. Returns "positions that differ (;)|paper names the roster
' does not list (;)|roster names not on the paper (;)" - "-" for an empty part. Names compare trimmed, any case;
' a blank paper line is a position the paper left empty (it differs when the roster has someone there).
Public Function ReconcileDiffCore(ByVal paper As String, ByVal roster As String) As String
    Dim p As Variant: p = SplitNames(paper)
    Dim r As Variant: r = SplitNames(roster)
    Dim np As Long, nr As Long, n As Long, i As Long, j As Long
    np = UBound(p) + 1: nr = UBound(r) + 1
    n = IIf(np > nr, np, nr)
    Dim diff As String, missing As String, extra As String, found As Boolean
    For i = 0 To n - 1
        If StrComp(NameAt(p, i), NameAt(r, i), vbTextCompare) <> 0 Then diff = diff & IIf(diff = "", "", ";") & (i + 1)
    Next i
    For i = 0 To np - 1
        If NameAt(p, i) <> "" Then
            found = False
            For j = 0 To nr - 1
                If StrComp(NameAt(p, i), NameAt(r, j), vbTextCompare) = 0 Then found = True: Exit For
            Next j
            If Not found Then missing = missing & IIf(missing = "", "", ";") & NameAt(p, i)
        End If
    Next i
    For j = 0 To nr - 1
        If NameAt(r, j) <> "" Then
            found = False
            For i = 0 To np - 1
                If StrComp(NameAt(p, i), NameAt(r, j), vbTextCompare) = 0 Then found = True: Exit For
            Next i
            If Not found Then extra = extra & IIf(extra = "", "", ";") & NameAt(r, j)
        End If
    Next j
    ReconcileDiffCore = IIf(diff = "", "-", diff) & "|" & IIf(missing = "", "-", missing) & "|" & IIf(extra = "", "-", extra)
End Function

Private Function SplitNames(ByVal s As String) As Variant
    If Trim$(s) = "" Then
        SplitNames = Split("", ";")                  ' UBound -1: no names
    Else
        SplitNames = Split(s, ";")
    End If
End Function

Private Function NameAt(ByRef arr As Variant, ByVal i As Long) As String
    If i >= 0 And i <= UBound(arr) Then NameAt = Trim$(CStr(arr(i)))
End Function

' ====================================================================
'  THE SHEET
' ====================================================================
Private Function ReconcileSheet(Optional ByVal create As Boolean = False) As Worksheet
    If Not SheetExists(RECON_SHEET) Then
        If Not create Then Exit Function
        Dim was As Object: Set was = ActiveSheet
        Dim ws As Worksheet
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RECON_SHEET
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)
        DrawFrame ws
        On Error Resume Next
        If Not was Is Nothing Then was.Activate
        On Error GoTo 0
    End If
    Set ReconcileSheet = ThisWorkbook.Worksheets(RECON_SHEET)
End Function

' The fixed part: titles, the three typed cells with their hints, the headings, the position numbers, the
' validation lists, the three buttons. Drawn when the sheet is made, and again by ShowReconcile when the frame's version
' in L1 differs from this release (the typed cells are kept across the redraw).
Private Sub DrawFrame(ByVal ws As Worksheet)
    UnlockSheet ws
    ws.Cells.Clear
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "rc_" Then ws.Shapes(i).Delete
    Next i
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    With ws.Range("A1")
        .Value = "RECONCILE: the paper call order beside the roster's"
        .Font.Size = 16: .Font.Bold = True: .Font.Color = RGB(30, 56, 100)
    End With
    ws.Range("A2").Value = "Type the paper roster's call order for one unit (names top to bottom, column B), the unit and the date the paper was good for, then click Compare. " & _
                           "Every position that differs shows the roster's reason. Type a Decision and a Note per line; Save PDF files the page beside the roster for the Nov 15 pack."
    ws.Range("A2").Font.Color = RGB(90, 100, 110)
    ws.Range("B4").Value = "Unit:": ws.Range("B5").Value = "Paper good for (date):": ws.Range("B6").Value = "Paper roster (which sheet):"
    ws.Range("B4:B6").Font.Bold = True
    ws.Range("D4").Value = "pick the unit the paper sheet is for"
    ws.Range("D5").Value = "the day the paper order was read (type it as Oct 12, 12 Oct 2026 or 10/12): crews and planned moves as of that day; Total Hours as of now"
    ws.Range("D6").Value = "optional: which paper sheet, who read it"
    ws.Range("D4:D6").Font.Color = RGB(120, 130, 140): ws.Range("D4:D6").Font.Italic = True
    ws.Range(RC_UNIT & ":" & RC_PAPER).NumberFormat = "@"
    ws.Range(RC_UNIT & ":" & RC_PAPER).Interior.Color = RGB(255, 250, 230)
    ws.Range(RC_DATE).Value = DateText(PeriodStartFor(Date))
    On Error Resume Next
    With ws.Range(RC_UNIT).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(DeptList(), ",")
        .InCellDropdown = True
        .ErrorTitle = "Unit": .ErrorMessage = "Pick a unit from the list."
    End With
    On Error GoTo 0
    ' the table
    Dim hdr As Variant: hdr = Array("#", "Paper (type the names)", "Roster's order", "Code", "Total h", "Seniority", "Differs?", "The roster's reason", "Decision", "Note")
    Dim r As Long: r = RC_FIRST - 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Value = hdr
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 10))
        .Font.Bold = True: .Interior.Color = RGB(30, 56, 100): .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 20
    For i = 1 To RC_POSITIONS
        ws.Cells(RC_FIRST + i - 1, 1).Value = i
    Next i
    ws.Range(ws.Cells(RC_FIRST, 1), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 1)).Font.Color = RGB(120, 130, 140)
    ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 2)).Interior.Color = RGB(255, 250, 230)
    ws.Range(ws.Cells(RC_FIRST, 9), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 10)).Interior.Color = RGB(255, 250, 230)
    ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 10)).NumberFormat = "@"
    On Error Resume Next
    With ws.Range(ws.Cells(RC_FIRST, 9), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 9)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=RC_DECISIONS
        .InCellDropdown = True
        .ErrorTitle = "Decision": .ErrorMessage = "Paper wrong, Rule missing, Data error or Agreed."
    End With
    On Error GoTo 0
    ws.Columns(1).ColumnWidth = 4: ws.Columns(2).ColumnWidth = 26: ws.Columns(3).ColumnWidth = 26: ws.Columns(4).ColumnWidth = 7
    ws.Columns(5).ColumnWidth = 8: ws.Columns(6).ColumnWidth = 13: ws.Columns(7).ColumnWidth = 8: ws.Columns(8).ColumnWidth = 60
    ws.Columns(9).ColumnWidth = 13: ws.Columns(10).ColumnWidth = 36
    ws.Range(ws.Cells(RC_FIRST, 8), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 8)).WrapText = True
    ' the buttons (shapes, as the Desk's), on row 3 so the hints in column D stay readable (review)
    ws.Rows(3).RowHeight = 28
    Dim x As Double: x = ws.Range("B3").Left
    RcButton ws, "rc_btn_compare", "Compare", "modReconcile.ReconcileCompare", x, ws.Range("B3").Top + 2, RGB(0, 128, 70)
    RcButton ws, "rc_btn_pdf", "Save PDF", "modReconcile.ReconcilePdf", x + 110, ws.Range("B3").Top + 2, RGB(30, 56, 100)
    RcButton ws, "rc_btn_clear", "Clear the page", "modReconcile.ReconcileClear", x + 220, ws.Range("B3").Top + 2, RGB(120, 130, 140)
    ' the summary line wraps across the page (review M4: a one-cell line is cut off in the PDF)
    With ws.Range("A7:J7")
        .Merge: .WrapText = True: .VerticalAlignment = xlTop
    End With
    ws.Rows(7).RowHeight = 54
    ws.Range(RC_SUMMARY).Value = "Not compared yet."
    ws.Range(RC_SUMMARY).Font.Italic = True
    ws.Range("L1").Value = RC_VERSION: ws.Range("L1").Font.Size = 7: ws.Range("L1").Font.Color = RGB(176, 190, 197)   ' the frame's version (ShowReconcile redraws on a change)
    ' the page
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 10)).Address
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = "Reconcile - Page &P of &N"
    End With
    On Error GoTo 0
    ProtectSheet ws
End Sub

Private Sub RcButton(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal fill As Long)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, 100, 24)
    s.Name = nm
    s.Fill.ForeColor.RGB = fill
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.OnAction = mac
    On Error Resume Next                            ' never raises
    s.DrawingObject.PrintObject = False
    With s.TextFrame2
        .WordWrap = msoFalse
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = cap
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = "Segoe UI Semibold": .Font.Size = 10: .Font.Bold = msoFalse
            .Font.Fill.ForeColor.RGB = vbWhite
        End With
    End With
End Sub

' A changed frame (a new release) redrawn around what the admin typed: the unit, the date, the paper's sheet, the names,
' the decisions and the notes are read first and written back after DrawFrame
Private Sub RedrawFrameKeepingTyped(ByVal ws As Worksheet)
    Dim last As Long: last = RC_FIRST + RC_POSITIONS - 1
    Dim top3 As Variant: top3 = ws.Range(RC_UNIT & ":" & RC_PAPER).Value
    Dim names As Variant: names = ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(last, 2)).Value
    Dim decs As Variant: decs = ws.Range(ws.Cells(RC_FIRST, 9), ws.Cells(last, 10)).Value
    DrawFrame ws
    UnlockSheet ws
    ws.Range(RC_UNIT & ":" & RC_PAPER).Value = top3
    ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(last, 2)).Value = names
    ws.Range(ws.Cells(RC_FIRST, 9), ws.Cells(last, 10)).Value = decs
    ProtectSheet ws
End Sub

' The typed cells open for an admin, everything else locked (ApplyViewCore calls it with the other typed sheets)
Public Sub LockReconcileCells(ByVal adminView As Boolean)
    Dim ws As Worksheet: Set ws = ReconcileSheet(False)
    If ws Is Nothing Then Exit Sub
    UnlockSheet ws
    ws.Cells.Locked = True
    If adminView Then
        ws.Range(RC_UNIT & ":" & RC_PAPER).Locked = False
        ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 2)).Locked = False
        ws.Range(ws.Cells(RC_FIRST, 9), ws.Cells(RC_FIRST + RC_POSITIONS - 1, 10)).Locked = False
    End If
    ProtectSheet ws
End Sub

' ====================================================================
'  THE ENTRY POINTS (admin)
' ====================================================================
' Alt+F8 / the Desk: the sheet, made on first use, in front
Public Sub ShowReconcile()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet: Set ws = ReconcileSheet(True)
    If SafeStr(ws.Range("L1").Value) <> RC_VERSION Then RedrawFrameKeepingTyped ws     ' an older frame: redrawn, the typed cells kept
    LockReconcileCells True
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    ws.Range(RC_UNIT).Select
    Exit Sub
Oops:
    ReportError "ShowReconcile"
End Sub

' The Compare button: the roster's order for the unit on the date beside the paper's, with the reasons
Public Sub ReconcileCompare()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet: Set ws = ReconcileSheet(True)
    Dim unit As String: unit = Trim$(SafeStr(ws.Range(RC_UNIT).Value))
    If unit = "" Then MsgBox "Pick the unit the paper sheet is for (cell C4).", vbExclamation, "Reconcile": Exit Sub
    If DeptCol(unit) = 0 Then MsgBox "'" & unit & "' is not one of the roster's units. Pick one from the list in C4.", vbExclamation, "Reconcile": Exit Sub
    Dim d As Date
    If Not ParseTypedDate(SafeStr(ws.Range(RC_DATE).Value), d) Then
        MsgBox "The date in C5 could not be read. Type it as Oct 12, 12 Oct 2026 or 10/12 (month first).", vbExclamation, "Reconcile": Exit Sub
    End If
    Dim last As Long: last = RC_FIRST + RC_POSITIONS - 1
    Dim paperV As Variant: paperV = ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(last, 2)).Value
    Dim paper As String, i As Long, nPaper As Long
    For i = 1 To RC_POSITIONS
        If Trim$(SafeStr(paperV(i, 1))) <> "" Then nPaper = i
    Next i
    For i = 1 To nPaper
        paper = paper & IIf(i = 1, "", ";") & Trim$(SafeStr(paperV(i, 1)))
    Next i
    If nPaper = 0 Then MsgBox "Type the paper roster's names in column B first (top to bottom).", vbExclamation, "Reconcile": Exit Sub
    If d > Date Then
        If Ask("The date in C5 is in the future (" & Format$(d, "ddd mmm d, yyyy") & "). Compare as of that day anyway?", vbYesNo + vbQuestion + vbDefaultButton2, "Reconcile", vbNo) <> vbYes Then Exit Sub
    End If
    ' the roster's list: the same call as the printed call-out roster, for that day (the crews and a planned move as of the
    ' day; Total Hours as of now - the sheet says so)
    Dim recs() As EmpRec, n As Long, nAll As Long
    nAll = LoadEligible(unit, "ALL", recs, 999, True, d)
    n = nAll
    If n > RC_POSITIONS Then n = RC_POSITIONS            ' the drawn rows; the list and posOf keep everyone (review)
    Dim roster As String, posOf As Object: Set posOf = CreateObject("Scripting.Dictionary"): posOf.CompareMode = 1
    For i = 1 To nAll
        roster = roster & IIf(i = 1, "", ";") & recs(i).EmpName
        posOf(recs(i).EmpName) = i
    Next i
    Dim res As String: res = ReconcileDiffCore(paper, roster)
    Dim parts As Variant: parts = Split(res, "|")
    Dim diffPos As Object: Set diffPos = CreateObject("Scripting.Dictionary")
    Dim pv As Variant, k As Variant, nDiff As Long
    If CStr(parts(0)) <> "-" Then
        pv = Split(CStr(parts(0)), ";")
        For Each k In pv
            diffPos(CLng(k)) = True
        Next k
        nDiff = UBound(pv) + 1                            ' (review N3) every position, not only the drawn ones
    End If
    ' the drawn columns
    Dim outA() As Variant: ReDim outA(1 To RC_POSITIONS, 1 To 6)
    Dim pName As String, why As String
    For i = 1 To RC_POSITIONS
        pName = ""
        If i <= nPaper Then pName = Trim$(SafeStr(paperV(i, 1)))
        If i <= n Then
            outA(i, 1) = recs(i).EmpName
            outA(i, 2) = recs(i).DeptCode
            outA(i, 3) = HoursText(recs(i).TotalHrs) & IIf(recs(i).EstHrs, "~", "")
            outA(i, 4) = SenText(recs(i).Seniority)
        End If
        If diffPos.Exists(i) Then
            outA(i, 5) = "yes"
            outA(i, 6) = DiffReason(pName, i, recs, n, posOf, unit)
        ElseIf i <= n Or pName <> "" Then
            outA(i, 5) = "same"
        End If
    Next i
    UnlockSheet ws
    ws.Range(ws.Cells(RC_FIRST, 3), ws.Cells(last, 8)).ClearContents
    ws.Range(ws.Cells(RC_FIRST, 3), ws.Cells(last, 8)).Value = outA
    ws.Range(ws.Cells(RC_FIRST, 3), ws.Cells(last, 8)).Font.Color = vbBlack
    For i = 1 To RC_POSITIONS
        If diffPos.Exists(i) Then ws.Range(ws.Cells(RC_FIRST + i - 1, 7), ws.Cells(RC_FIRST + i - 1, 8)).Font.Color = RGB(176, 0, 32)
    Next i
    Dim summ As String
    summ = "Compared " & Format$(Now, "ddd mmm d, h:mm AM/PM") & ": " & unit & ", crews and planned moves as of " & Format$(d, "ddd mmm d, yyyy") & ", Total Hours as of now - paper " & nPaper & " names, roster " & nAll & _
           IIf(nAll > RC_POSITIONS, " (the first " & RC_POSITIONS & " drawn)", "") & _
           ". " & IIf(nDiff = 0, "Every position agrees.", nDiff & " position(s) differ.") & _
           IIf(CStr(parts(1)) <> "-", " On the paper, not listed by the roster: " & Replace(CStr(parts(1)), ";", ", ") & ".", "") & _
           IIf(CStr(parts(2)) <> "-", " Listed by the roster, not on the paper: " & Replace(CStr(parts(2)), ";", ", ") & ".", "") & _
           "  " & TrustStrip()
    ws.Range(RC_SUMMARY).Value = summ
    ws.Range(RC_SUMMARY).Font.Italic = False
    ws.Range(RC_SUMMARY).Font.Color = IIf(nDiff = 0, RGB(0, 128, 70), RGB(176, 0, 32))
    ProtectSheet ws
    LockReconcileCells True
    Exit Sub
Oops:
    ReportError "ReconcileCompare"
End Sub

' Why the roster has someone else at position i: where it puts the paper's person and which key decided it
Private Function DiffReason(ByVal pName As String, ByVal i As Long, ByRef recs() As EmpRec, ByVal n As Long, ByVal posOf As Object, ByVal unit As String) As String
    If pName = "" Then
        If i <= n Then DiffReason = "The paper has nobody here; the roster lists " & recs(i).EmpName & "."
        Exit Function
    End If
    If Not posOf.Exists(pName) Then
        DiffReason = pName & " is not on the roster's list for " & unit & " on that day (no code in the unit, moved, or spelled differently on the roster)."
        Exit Function
    End If
    Dim k As Long: k = CLng(posOf(pName))
    If i > n Then DiffReason = "The roster puts " & pName & " at #" & k & "; it lists only " & n & " people.": Exit Function
    Dim pr As EmpRec: pr = recs(k)
    Dim rr As EmpRec: rr = recs(i)
    Dim key As String
    If pr.Grade <> rr.Grade Then
        key = "grade code (" & pr.DeptCode & " vs " & rr.DeptCode & ")"
    ElseIf Abs(pr.TotalHrs - rr.TotalHrs) > 0.001 Then
        key = "Total Hours (" & HoursText(pr.TotalHrs) & IIf(pr.EstHrs, "~", "") & " vs " & HoursText(rr.TotalHrs) & IIf(rr.EstHrs, "~", "") & ")"
    ElseIf pr.Seniority <> rr.Seniority Then
        key = "seniority (" & SenText(pr.Seniority) & " vs " & SenText(rr.Seniority) & ")"
    Else
        key = "the same three keys: the name breaks the tie"
    End If
    DiffReason = "The roster puts " & pName & " at #" & k & " and " & rr.EmpName & " here: " & key & "."
End Function

Private Function SenText(ByVal key As Double) As String
    If key > 0 And key < 100000 Then SenText = Format$(CDate(key), "mmm d, yyyy") Else SenText = "-"
End Function

' The Save PDF button: this page, one file per unit and date, beside the roster
Public Sub ReconcilePdf()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet: Set ws = ReconcileSheet(False)
    If ws Is Nothing Then Exit Sub
    Dim unit As String: unit = Trim$(SafeStr(ws.Range(RC_UNIT).Value))
    Dim d As Date
    If Not ParseTypedDate(SafeStr(ws.Range(RC_DATE).Value), d) Then d = Date
    If InStr(SafeStr(ws.Range(RC_SUMMARY).Value), "Compared ") <> 1 Then
        If Ask("The page has not been compared since it was last changed. Save it as it is?", vbYesNo + vbQuestion, "Reconcile", vbYes) <> vbYes Then Exit Sub
    End If
    ReportPdf ws, "Reconcile " & IIf(unit = "", "unit", unit) & " " & Format$(d, "yyyy-mm-dd"), "Reconcile", True
    Exit Sub
Oops:
    ReportError "ReconcilePdf"
End Sub

' The Clear button: the typed names, decisions and notes and the drawn columns, after a question
Public Sub ReconcileClear()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim ws As Worksheet: Set ws = ReconcileSheet(False)
    If ws Is Nothing Then Exit Sub
    If Ask("Clear the paper names, the decisions, the notes and the comparison? (Save PDF first if this page should be kept.)", _
           vbYesNo + vbQuestion + vbDefaultButton2, "Reconcile", vbYes) <> vbYes Then Exit Sub
    Dim last As Long: last = RC_FIRST + RC_POSITIONS - 1
    UnlockSheet ws
    ws.Range(ws.Cells(RC_FIRST, 2), ws.Cells(last, 10)).ClearContents
    ws.Range(ws.Cells(RC_FIRST, 3), ws.Cells(last, 8)).Font.Color = vbBlack
    ws.Range(RC_PAPER).ClearContents
    ws.Range(RC_SUMMARY).Value = "Not compared yet."
    ws.Range(RC_SUMMARY).Font.Italic = True: ws.Range(RC_SUMMARY).Font.Color = vbBlack
    ProtectSheet ws
    LockReconcileCells True
    Exit Sub
Oops:
    ReportError "ReconcileClear"
End Sub
