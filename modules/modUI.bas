Attribute VB_Name = "modUI"
Option Explicit

' ====================================================================
'  ROSTER LOOK & FEEL
'  ApplyRosterLayout - safe to run any time (and again after adding
'  employees). It never changes hours. It:
'   - renames the roster tab to ROSTER_SHEET ("OT Roster")
'   - puts pay-period START dates in row 2, "to" in row 3, END in row 4
'     (rewrites the period formulas and verifies every total)
'   - sorts the OT Log newest first
'   - builds the command bar inside the frozen area so it never scrolls;
'     Print/Publish and Admin buttons only show for admins
'   - freezes rows 1-4 and columns A:S (names + details stay put while
'     the pay periods scroll), filter dropdowns on L:S only
'   - formats: navy headers, banded rows, current pay period highlighted,
'     grades above the charge limit greyed, whole-hour totals
' ====================================================================

Private Const STATUS_CELL As String = "D3"      ' status (3 lines, plain text) - in the frozen names column
Public Const UI_LARGER_SCALE As Double = 1.2    ' v27 (item 20, B14): the Desk's "Larger type" scale for a login (UiScale below)
Public Const UI_VERSION As String = "Roster UI v29.1 (Oct 28)"
Private Const NAVY As Long = 6567966            ' RGB(30, 56, 100)
Private Const INK As Long = 3355443             ' RGB(51, 51, 51)
Private Const LIGHT As Long = 15787486          ' RGB(222, 229, 240)

Private mFlipStarted As Boolean
Private mFormScale As Double                    ' v27 (review M6): the scale ScaleForm chose for the form being built (clamped); ScaleWidths / UiPx use it

Public Sub ApplyRosterLayout()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub        ' run RunMigration first

    Dim ws As Worksheet: Set ws = RosterSheet()
    Application.ScreenUpdating = False
    UnlockSheet ws
    On Error GoTo Failed

    RenameRosterTab ws
    If Not FlipPeriodRows(ws) Then GoTo Done      ' message already shown

    ClearRosterFilter ws
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then GoTo Done

    ' log: newest first
    Dim lo As ListObject: Set lo = LogTable()
    If Not lo Is Nothing Then
        UnlockSheet lo.Parent
        SortLogNewestFirst lo
        ProtectSheet lo.Parent
    End If

    ws.Activate
    FormatRoster ws, lastRow
    BuildCommandBar ws
    EnsureRosterFilter ws
    SetRosterStatus "All"
    ApplyView IsAdmin(), gDeskMode                ' shows/hides admin buttons, re-protects; v25.3: a Desk user keeps the Desk view

    ws.Activate
    ws.Range("D5").Select
    Application.ScreenUpdating = True
    SaveIfWritable
    MsgBox "Roster layout applied.", vbInformation, "Call Out Roster"
    Exit Sub

Done:
    ProtectSheet ws
    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Dim em As String: em = Err.Description
    LogError "ApplyRosterLayout", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.ScreenUpdating = True
    If mFlipStarted Then
        MarkNoSave
        MsgBox "Layout stopped while switching the date rows: " & em & vbNewLine & vbNewLine & _
               "Roster totals may show wrong until fixed. Saving is blocked - close WITHOUT saving " & _
               "and run ApplyRosterLayout again. (The OT Log itself is untouched.)", vbCritical, "Call Out Roster"
    Else
        MsgBox "Layout stopped part-way: " & em & vbNewLine & _
               "Hours are untouched; fix the cause and run it again.", vbExclamation, "Call Out Roster"
    End If
End Sub

' Status line in row 3: current pay period | OT Log size | current view.
' Written as plain text (no formula). Called by the sort buttons with the
' view name, and with no argument (keeps the view) after log writes and
' whenever the view is applied (open / save).
Public Sub SetRosterStatus(Optional ByVal viewTxt As String = "")
    On Error GoTo Quit
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim cur As String: cur = SafeStr(ws.Range(STATUS_CELL).Value)
    If viewTxt = "" Then
        Dim p As Long: p = InStr(1, cur, "View: ")
        If p > 0 Then viewTxt = Trim$(Mid$(cur, p + 6)) Else viewTxt = "All"
    End If

    Dim n As Long, lo As ListObject
    If SheetExists(LOG_SHEET) Then
        On Error Resume Next
        Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
        If Not lo Is Nothing Then n = Application.WorksheetFunction.Count(lo.ListColumns(L_ENTRY).Range)
        On Error GoTo Quit
    End If

    Dim ps As Date: ps = PeriodStartFor(Date)
    Dim txt As String
    Dim kTxt As String: kTxt = KronosStatusText()        ' v20: how old the worked hours are
    Dim dt As Date: dt = KronosDataThrough()             ' v21: provisional OT counts after this date
    Dim icTxt As String: icTxt = IntegrityStatusText()   ' v23: "" until IntegrityCheck has run this session
    Dim roTxt As String: roTxt = RolloverNote(True)      ' v23: "rolled over Dec 27", for a month after the year-end rollover
    ' one text-format cell: keep it well under 255 characters or Excel shows ####
    ' v27 (item 20, B15): line 3 starts with who is signed in and what that login may do (the gating levels)
    txt = "Pay period: " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d") & IIf(roTxt <> "", "  |  " & roTxt, "") & vbLf & _
          kTxt & IIf(dt > 0 And KronosAwaitFrom() = 0, "  |  data through " & Format$(dt, "mmm d"), "") & vbLf & _
          SignedInText() & "  |  Log: " & Format$(n, "#,##0") & IIf(icTxt <> "", "  |  " & icTxt, "") & "  |  View: " & Left$(viewTxt, 50)
    If txt = Replace(cur, vbCr, "") Then Exit Sub

    Dim wasLocked As Boolean: wasLocked = ws.ProtectContents
    If wasLocked Then UnlockSheet ws
    ' leftovers from earlier versions of the status line
    If Left$(SafeStr(ws.Range("E3").Value), 10) = "Pay period" Then ws.Range("E3").ClearContents
    If Left$(SafeStr(ws.Range("L3").Value), 5) = "View:" Then ws.Range("L3").ClearContents
    If ws.Rows(3).RowHeight < 46 Then ws.Rows(3).RowHeight = 46   ' room for the lines (v27: the signed-in line made line 3 wrap)
    With ws.Range(STATUS_CELL)
        .NumberFormat = "@"
        .WrapText = True
        .Value = Left$(txt, 250)                      ' one text cell: past 255 Excel shows ####
        .Font.Color = NAVY
        If KronosIsStale() Then .Characters(InStr(txt, kTxt), Len(kTxt)).Font.Color = RGB(190, 30, 30)
        If icTxt <> "" And IntegrityHasProblems() Then .Characters(InStr(txt, icTxt), Len(icTxt)).Font.Color = RGB(190, 30, 30)
    End With
    If wasLocked Then ProtectSheet ws
    Exit Sub
Quit:
    On Error Resume Next
    If wasLocked Then ProtectSheet ws
End Sub


' --------------------------------------------------------------------
'  TAB NAME
' --------------------------------------------------------------------
Private Sub RenameRosterTab(ByVal ws As Worksheet)
    If ws.Name = ROSTER_SHEET Then Exit Sub
    If SheetExists(ROSTER_SHEET) Then Exit Sub    ' name taken - leave it
    UnlockStructure
    ws.Name = ROSTER_SHEET
    RelockStructure
End Sub


' --------------------------------------------------------------------
'  PERIOD DATES: start (row 2) / "to" (row 3) / end (row 4)
'  Switches the old layout (start dates in row 4) and (re)writes every
'  row's period formulas, then checks each row against the OT Log.
'  Also repairs formulas that were stored as text by an earlier run.
' --------------------------------------------------------------------
Private Function FlipPeriodRows(ByVal ws As Worksheet) As Boolean
    Dim lastPer As Long: lastPer = PeriodLastCol()             ' AS, or AT in a 27-period year (v23)
    Dim pf As String: pf = ColLetter(C_PERIOD1)
    Dim pl As String: pl = ColLetter(lastPer)
    Dim p2 As String: p2 = ColLetter(C_PERIOD1 + 1)
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim r As Long, c As Long, nm As String

    ' rows holding typed numbers in the pay-period cells are left for Refresh Formulas
    Dim typed As String, hasTyped As Boolean, cel As Range
    Dim skip() As Boolean: ReDim skip(R_FIRST To Application.Max(lastRow, R_FIRST))
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" Then
            hasTyped = False
            For c = C_PERIOD1 To lastPer
                Set cel = ws.Cells(r, c)
                If Not IsOurFormula(cel) And Not IsBrokenFormula(cel) Then
                    If Not IsEmpty(cel.Value2) Then hasTyped = True: Exit For
                End If
            Next c
            If hasTyped Then typed = typed & "  " & nm & vbNewLine: skip(r) = True
        End If
    Next r

    If PStartRow() <> 2 Then
        If typed <> "" Then
            MsgBox "Nothing was changed. These rows have typed numbers in the pay-period cells." & vbNewLine & _
                   "Run Alt+F8 > RefreshRosterFormulas first (it moves them into the OT Log), then run this again:" & _
                   vbNewLine & vbNewLine & typed, vbExclamation, "Call Out Roster"
            Exit Function
        End If
        Dim firstStart As Date: firstStart = CDate(ws.Range(pf & "4").Value)
        mFlipStarted = True
        ws.Range(pf & "2:" & pl & "4").NumberFormat = "General"
        ws.Range(pf & "2").Value = firstStart
        ws.Range(p2 & "2:" & pl & "2").FormulaR1C1 = "=RC[-1]+14"
        ws.Range(pf & "4:" & pl & "4").FormulaR1C1 = "=R2C+13"
        Application.Calculate
        If PStartRow() <> 2 Then
            MsgBox "Couldn't switch the pay-period rows.", vbExclamation, "Call Out Roster"
            Exit Function
        End If
    End If
    ws.Range(pf & "3:" & pl & "3").NumberFormat = "General"
    ws.Range(pf & "3:" & pl & "3").Value = "to"

    ' rewrite each row's formulas (rows with typed numbers were listed above).
    ' v21: once Total includes the provisional term, C and G are rewritten too,
    ' so their year window ($T$2 / $AS$4) always matches the date rows.
    Dim provLive As Boolean: provLive = ProvisionalLive()
    For r = R_FIRST To lastRow
        If Trim$(SafeStr(ws.Cells(r, C_NAME).Value)) <> "" And Not skip(r) Then
            WritePeriodRow ws, r
            If provLive Then WriteTotalRow ws, r
        End If
    Next r
    Application.Calculate

    ' verify every rewritten row against the OT Log
    Dim rowSum As Double, mism As String, v As Variant
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If nm <> "" And IsOurFormula(ws.Cells(r, C_PERIOD1)) Then
            rowSum = 0
            For c = C_PERIOD1 To lastPer
                v = ws.Cells(r, c).Value2
                If VarType(v) = vbDouble Then rowSum = rowSum + v
            Next c
            If Abs(rowSum - LogTotalOnRoster(nm)) > 0.0001 Then
                mism = mism & "  " & nm & ": roster " & rowSum & ", OT Log " & LogTotalOnRoster(nm) & vbNewLine
            End If
        End If
    Next r
    If mism <> "" Then
        MarkNoSave
        MsgBox "STOP - these roster rows don't match the OT Log:" & vbNewLine & mism & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving and send " & PRIMARY_ADMIN_NAME & " this message.", vbCritical, "Call Out Roster"
        Exit Function
    End If
    If typed <> "" Then
        MsgBox "These rows have typed numbers in the pay-period cells and were left alone." & vbNewLine & _
               "Run Alt+F8 > RefreshRosterFormulas to move them into the OT Log:" & vbNewLine & vbNewLine & typed, vbInformation, "Call Out Roster"
    End If
    ' v26.5.5 (A9): the block is lifted only for the flip THIS call started and verified. A NoSave set by
    ' another failed check (a people save, the rollover) stood for a reason: it stays, and the box says so.
    If mFlipStarted Then
        mFlipStarted = False
        ClearNoSave                 ' every row now matches the OT Log
    ElseIf NoSave() Then
        MsgBox "The date rows were already in place, so nothing was rewritten here." & vbNewLine & vbNewLine & _
               "Saving stays BLOCKED from an earlier failed check in this session: close WITHOUT saving and reopen, " & _
               "then run the macro that failed again.", vbExclamation, "Saving is blocked"
    End If
    FlipPeriodRows = True
End Function


' --------------------------------------------------------------------
'  COMMAND BAR - fits inside the frozen columns (D:S) so it never moves.
'  Names starting "ui_adm_" are admin-only (ApplyView hides them); "ui_sup_" (v25.5) go to the logins on
'  SUPERVISOR_USERS and admins; "ui_btn_" and "ui_srt_" to everyone.
' --------------------------------------------------------------------
Private Sub BuildCommandBar(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    ws.Buttons.Delete                                  ' old grey form buttons
    On Error GoTo 0
    For i = ws.Shapes.Count To 1 Step -1               ' previous runs of this macro
        If Left$(ws.Shapes(i).Name, 3) = "ui_" Then ws.Shapes(i).Delete
    Next i

    ws.Rows(1).RowHeight = 96
    Dim x0 As Double: x0 = ws.Cells(1, C_SHIFT).Left + 4          ' starts at column E

    ' base widths (points), scaled down if the frozen area is narrower
    Const G As Double = 3          ' gap between buttons
    Const GG As Double = 8         ' gap between groups
    Dim wSrt As Double, wL As Double, wP As Double, wA As Double, wE As Double, sc As Double
    wSrt = 62: wL = 90: wP = 88: wA = 88: wE = 90          ' v22: a little narrower, room for the PEOPLE group
    Dim fixedGaps As Double: fixedGaps = 2 * G + G + G + 4 * GG
    sc = 1
    wSrt = wSrt * sc: wL = wL * sc: wP = wP * sc: wA = wA * sc: wE = wE * sc

    Dim y0 As Double: y0 = ws.Range("A1").Top + 3
    Dim y(1 To 3) As Double
    y(1) = y0 + 14: y(2) = y(1) + 27: y(3) = y(2) + 27
    Dim x As Double: x = x0

    ' --- Sort & filter (everyone) ---
    GroupLabel ws, "ui_lbl_sort", "SORT & FILTER", x, y0, 3 * wSrt + 2 * G
    SortBtn ws, "Derivatives", "Derivatives", x, y(1), wSrt
    SortBtn ws, "Phosphine", "Phosphine", x + (wSrt + G), y(1), wSrt
    SortBtn ws, "Cypure", "Cypure", x + 2 * (wSrt + G), y(1), wSrt
    SortBtn ws, "Fumigation", "Mat_Handler", x, y(2), wSrt
    SortBtn ws, "QC - Lab", "QC_Lab", x + (wSrt + G), y(2), wSrt
    SortBtn ws, "QC - Gases", "QC_Gas", x + 2 * (wSrt + G), y(2), wSrt
    SortBtn ws, "Utilities", "Utilities", x, y(3), wSrt
    SortBtn ws, "SBP", "SBP", x + (wSrt + G), y(3), wSrt
    SortBtn ws, "Show All", "ShowAllRows", x + 2 * (wSrt + G), y(3), wSrt
    x = x + 3 * wSrt + 2 * G + GG

    ' --- Log (everyone with edit access) ---
    GroupLabel ws, "ui_lbl_log", "LOG", x, y0, wL
    Btn ws, "ui_btn_LogCallOut", "Log Call-Out", "modOvertimeEntry.ShowOvertimeForm", x, y(1), wL, RGB(0, 128, 70), vbWhite, True
    Btn ws, "ui_adm_Switch", "Shift Switch", "modShiftSwitch.ShowShiftSwitchForm", x, y(2), wL, RGB(180, 100, 0), vbWhite, True
    Btn ws, "ui_adm_TimeOff", "Book Time Off", "modCalendar.ShowTimeOffForm", x, y(3), wL, RGB(31, 56, 100), vbWhite, True
    x = x + wL + GG

    ' --- People (v22; v25.3: Sick Call for every supervisor, the rest admin) ---
    GroupLabel ws, "ui_sup_lbl_people", "PEOPLE", x, y0, wE                 ' v25.3: Sick Call is a supervisor's job; v25.5: the heading goes with its buttons (a plain user has none here)
    Btn ws, "ui_sup_Sick", "Sick Call", "modLeave.ShowSickCallForm", x, y(1), wE, RGB(230, 81, 0), vbWhite, True   ' v25.3: every supervisor; v25.5: ui_sup_ = the logins on SUPERVISOR_USERS + admins
    Btn ws, "ui_adm_Leave", "Leave", "modLeave.ShowLeaveForm", x, y(2), wE, RGB(106, 27, 154), vbWhite, True
    Btn ws, "ui_adm_Convert", "Convert imports", "modLeave.ShowConvertForm", x, y(3), wE, RGB(21, 101, 192), vbWhite, False
    x = x + wE + GG

    ' --- Postings (v25.3: New Posting and Award Posting for every supervisor; Load Posting admin) ---
    Dim grn As Long: grn = RGB(0, 128, 70)
    GroupLabel ws, "ui_sup_lbl_post", "POSTINGS", x, y0, wL                  ' v25.3: New Posting and Award Posting are a supervisor's jobs; v25.5: the heading goes with them
    Btn ws, "ui_sup_NewPost", "New Posting", "modPostings.ShowNewPostingForm", x, y(1), wL, grn, vbWhite, True   ' v25.3: every supervisor; v25.5: ui_sup_
    Btn ws, "ui_adm_LoadPost", "Load Posting", "modPostings.ShowLoadPostingForm", x, y(2), wL, grn, vbWhite, True
    Btn ws, "ui_sup_Award", "Award Posting", "modOvertimeEntry.ShowPostingForm", x, y(3), wL, grn, vbWhite, True   ' v25.3: every supervisor; v25.5: ui_sup_
    x = x + wL + GG

    ' --- Print / publish (v25.3: Calendar for everyone; v25.5: Call Out Roster and Member View for supervisors and admins, the report admin) ---
    GroupLabel ws, "ui_lbl_print", "PRINT / PUBLISH", x, y0, 2 * wP + G      ' v25.3: Call Out Roster and Calendar are everyone's
    Btn ws, "ui_sup_CallOut", "Call Out Roster", "Call_Out", x, y(1), wP, NAVY, vbWhite, False   ' v25.3: every supervisor; v25.5: ui_sup_
    Btn ws, "ui_sup_Member", "Member View", "PublishMemberView", x + wP + G, y(1), wP, NAVY, vbWhite, False   ' v25.5: supervisors publish too (bundle item 8)
    Btn ws, "ui_btn_Calendar", "Calendar", "modCalendar.ShowCalendar", x, y(2), wP, RGB(0, 128, 70), vbWhite, True
    Btn ws, "ui_adm_Refusals", "Refusals Report", "modReports.RefusalsReport", x + wP + G, y(2), wP, NAVY, vbWhite, False   ' v24.1
    ' v25.3: the Desk (the second admin's screen), under Calendar so it is in view in every layout. ui_adm_ so
    ' non-admins never see it; ApplyViewCore keeps it visible in the Desk view, where the other admin buttons hide
    Btn ws, "ui_adm_Desk", "< Desk", "modDesk.ShowDesk", x, y(3), wP, NAVY, vbWhite, False
    Btn ws, "ui_adm_Undo", "Undo last entry", "modOTLog.UndoLast", x + wP + G, y(3), wP, RGB(236, 239, 243), INK, False   ' v25.5 (A9): the free slot beside "< Desk" (the ADMIN group is full; gry is set below)
    x = x + 2 * wP + G + GG

    ' --- Admin ---
    Dim gry As Long: gry = RGB(236, 239, 243)
    GroupLabel ws, "ui_adm_lbl_admin", "ADMIN", x, y0, 2 * wA + G
    Btn ws, "ui_adm_Kronos", "Import Kronos", "SyncKronosData", x, y(1), wA, gry, INK, False
    Btn ws, "ui_adm_Adjust", "Add Adjustment", "AddAdjustment", x + wA + G, y(1), wA, gry, INK, False
    Btn ws, "ui_adm_Void", "Void Entry", "VoidEntry", x, y(2), wA, gry, INK, False
    Btn ws, "ui_adm_Employee", "Employee", "modPeople.ShowEmployeeForm", x + wA + G, y(2), wA, gry, INK, False   ' v24: add / move / rename / remove
    Btn ws, "ui_adm_Access", "Full Access", "LoginFullAccess.LoginFullAccess", x, y(3), wA, gry, INK, False
    Btn ws, "ui_adm_Attend", "Attendance", "modAttendance.ShowAttendance", x + wA + G, y(3), wA, gry, INK, False   ' v22.1
    ' Refresh Formulas is a repair tool now (roster totals are live formulas): Alt+F8 > RefreshRosterFormulas
    x = x + 2 * wA + G + GG

    ' --- Find (v25.5, A10: supervisors and admins; every other group was full) ---
    GroupLabel ws, "ui_sup_lbl_find", "FIND", x, y0, wP
    Btn ws, "ui_sup_Find", "Find a person", "modReports.FindPerson", x, y(1), wP, vbWhite, NAVY, False
    Btn ws, "ui_sup_History", "Employee history", "modReports.EmployeeHistory", x, y(2), wP, vbWhite, NAVY, False   ' v26: the report (admin: full or member-safe; supervisor: member-safe)
    Btn ws, "ui_btn_Help", "How do I ...", "modDesk.ShowHelp", x, y(3), wP, vbWhite, NAVY, False   ' v27 (item 20, B15): the Help sheet, for everyone (ui_btn_ = always shown)
End Sub

Private Sub GroupLabel(ByVal ws As Worksheet, ByVal nm As String, ByVal txt As String, _
                       ByVal l As Double, ByVal t As Double, ByVal w As Double)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, 12)
    s.Name = nm
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse
        With .TextRange
            .Text = txt
            .Font.Size = 9                                ' v27 (item 20, B14): the group names (was 7.5)
            .Font.Bold = msoTrue
            .Font.Name = "Segoe UI"
            .Font.Fill.ForeColor.RGB = RGB(110, 120, 135)
        End With
    End With
End Sub

Private Sub SortBtn(ByVal ws As Worksheet, ByVal cap As String, ByVal mac As String, _
                    ByVal l As Double, ByVal t As Double, ByVal w As Double)
    Dim s As Shape
    Set s = Btn(ws, "ui_srt_" & Replace(Replace(cap, " ", ""), "-", ""), cap, mac, l, t, w, vbWhite, NAVY, False)
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = RGB(160, 175, 200)
    s.Line.Weight = 0.75
End Sub

Private Function Btn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                     ByVal l As Double, ByVal t As Double, ByVal w As Double, _
                     ByVal fillCol As Long, ByVal textCol As Long, ByVal isBold As Boolean) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, 24)
    s.Name = nm
    s.Adjustments.Item(1) = 0.18
    s.Fill.ForeColor.RGB = fillCol
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.OnAction = mac
    With s.TextFrame2
        .MarginLeft = 1: .MarginRight = 1: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = cap
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = "Segoe UI"
            .Font.Size = 10                               ' v27 (item 20, B14): every label floors at 10 pt (was 8.5)
            .Font.Bold = IIf(isBold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = textCol
        End With
    End With
    Set Btn = s
End Function


' --------------------------------------------------------------------
'  SHEET FORMATTING
' --------------------------------------------------------------------
Private Sub FormatRoster(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim lastPer As Long: lastPer = PeriodLastCol()             ' AS, or AT in a 27-period year (v23)
    Dim pf As String: pf = ColLetter(C_PERIOD1)
    Dim pl As String: pl = ColLetter(lastPer)
    Dim body As Range: Set body = ws.Range(ws.Cells(R_FIRST, 3), ws.Cells(lastRow, lastPer))
    Dim hdr As Range: Set hdr = ws.Range(ws.Cells(2, 3), ws.Cells(4, lastPer))

    ' clear old header/body conditional formats
    ws.Range("A2:XFD4").FormatConditions.Delete
    body.FormatConditions.Delete

    With ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, lastPer))
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = INK
    End With

    ' ---------- header block: rows 2-4 ----------
    hdr.Interior.Pattern = xlNone
    hdr.Borders.LineStyle = xlNone
    ws.Rows(2).RowHeight = 30
    ws.Rows(3).RowHeight = 34                 ' status: 3 lines (v20)
    ws.Rows(4).RowHeight = 17

    ' row 2, B:S - column titles (B = Home Unit from v21; SetupHomeUnit writes the title)
    With ws.Range(ws.Cells(2, C_HOME), ws.Cells(2, 19))
        .Interior.Color = NAVY
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 8
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(R_FIRST, C_HOME), ws.Cells(lastRow, C_HOME))
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = INK
        .HorizontalAlignment = xlCenter
    End With
    HomeUnitValidation ws, lastRow
    ' row 3, C:S - status lives in D3 (two lines, stays visible when scrolling)
    With ws.Range(ws.Cells(3, 3), ws.Cells(3, 19))
        .ClearContents                            ' removes the old D3 formula / L3 text
        .WrapText = False
        .ShrinkToFit = False
    End With
    With ws.Range(STATUS_CELL)
        .NumberFormat = "@"
        .WrapText = True
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = NAVY
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ' row 4, D:S - filter row
    With ws.Range(ws.Cells(4, 3), ws.Cells(4, 19))
        .Interior.Color = LIGHT
        .Font.Size = 8
    End With

    ' T:AS rows 2-4 - one navy date block: start / to / end
    With ws.Range(pf & "2:" & pl & "4")
        .Interior.Color = NAVY
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .WrapText = False
    End With
    With ws.Range(pf & "2:" & pl & "2")
        .NumberFormat = "mmm d"
        .Font.Bold = True
        .Font.Size = 8.5
        .VerticalAlignment = xlBottom
    End With
    With ws.Range(pf & "3:" & pl & "3")
        .Font.Italic = True
        .Font.Size = 7
        .Font.Color = RGB(190, 200, 220)
        .VerticalAlignment = xlCenter
    End With
    With ws.Range(pf & "4:" & pl & "4")
        .NumberFormat = "mmm d"
        .Font.Bold = True
        .Font.Size = 8.5
        .VerticalAlignment = xlTop
    End With
    ' thin separators between pay periods in the date block
    With ws.Range(pf & "2:" & pl & "4").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Color = RGB(70, 95, 140)
    End With

    ' ---------- body ----------
    body.Interior.Pattern = xlNone
    body.Borders.LineStyle = xlNone
    With body.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Color = RGB(228, 232, 238)
        .Weight = xlThin
    End With
    With body.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(200, 206, 215)
    End With
    ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, C_REFUSED)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, C_TOTAL)).Font.Bold = True
    ws.Range(ws.Cells(R_FIRST, C_SENIORITY), ws.Cells(lastRow, C_SENIORITY)).NumberFormat = "mmm d, yyyy"
    With ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME))
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Range(ws.Cells(R_FIRST, C_SHIFT), ws.Cells(lastRow, C_PHONE)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, 19)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(R_FIRST, C_PERIOD1), ws.Cells(lastRow, lastPer)).HorizontalAlignment = xlCenter
    ' divider between the frozen details and the pay periods
    With ws.Range(ws.Cells(2, C_PERIOD1), ws.Cells(lastRow, C_PERIOD1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = RGB(150, 160, 175)
        .Weight = xlMedium
    End With

    ' ---------- conditional formats (relative to the selected cell) ----------
    Dim fc As FormatCondition
    ws.Range("E" & R_FIRST).Select
    Set fc = ws.Range(ws.Cells(R_FIRST, C_SHIFT), ws.Cells(lastRow, lastPer)).FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=MOD(SUBTOTAL(103,$D$" & R_FIRST & ":$D" & R_FIRST & "),2)=0")
    fc.Interior.Color = RGB(245, 247, 250)

    ws.Range("L" & R_FIRST).Select
    Set fc = ws.Range(ws.Cells(R_FIRST, 12), ws.Cells(lastRow, 19)).FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=IFERROR(VALUE(MID(L" & R_FIRST & ",2,9))>" & CHARGE_MAX_GRADE & ",FALSE)")
    fc.Font.Color = RGB(160, 165, 175)
    fc.Font.Italic = True

    ws.Range(pf & "2").Select
    Set fc = ws.Range(pf & "2:" & pl & lastRow).FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=AND(" & pf & "$2<=TODAY(),TODAY()<=" & pf & "$4)")
    fc.Interior.Color = RGB(255, 236, 179)
    fc.Font.Color = RGB(60, 45, 0)
    fc.Font.Bold = True
    fc.SetFirstPriority

    ' ---------- columns ----------
    ' C = provisional accepted OT (v21): a hidden helper that Total Hours adds in
    With ws.Cells(2, C_PROV)
        .Value = "PROV"
        .Interior.Color = NAVY
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 8
    End With
    ws.Range(ws.Cells(R_FIRST, C_PROV), ws.Cells(lastRow, C_PROV)).NumberFormat = PERIOD_FMT
    ws.Columns(C_PROV).ColumnWidth = 7
    ws.Columns(C_PROV).Hidden = True
    ws.Columns(C_HOME).ColumnWidth = 14          ' Home Unit (v21)
    ws.Columns(C_HOME).Hidden = False
    ws.Columns(C_PAYROLL).ColumnWidth = 11       ' v24.2: Payroll # (filled by the Kronos import)
    ws.Columns(C_PAYROLL).Hidden = Not gAdminUnlocked   ' admin view only; ApplyView hides it in every saved copy
    ws.Columns(C_EMAIL).ColumnWidth = 26           ' EMAIL (ApplyView hides it for non-admins)
    ws.Columns(C_NAME).ColumnWidth = 26
    ws.Columns(C_SHIFT).ColumnWidth = 6.5
    ws.Columns(C_PHONE).ColumnWidth = 12.5
    ws.Range(ws.Columns(C_TOTAL), ws.Columns(C_REFUSED)).ColumnWidth = 8
    ws.Columns(C_SENIORITY).ColumnWidth = 11.5
    ws.Range(ws.Columns(12), ws.Columns(19)).ColumnWidth = 7
    ws.Range(ws.Columns(C_PERIOD1), ws.Columns(lastPer)).ColumnWidth = 7
    ws.Range(ws.Rows(R_FIRST), ws.Rows(lastRow)).RowHeight = 16

    ' ---------- window: rows 1-4 + columns A:D (names) frozen ----------
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Cells(R_FIRST, C_SHIFT).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    SetTabColor ws, NAVY                                ' v25.1
End Sub


' --------------------------------------------------------------------
'  HOME UNIT (v21) - column B, a drop-down of the eight units.
'  SetupHomeUnit (Alt+F8, once): writes the header and drop-down and fills
'  in everyone who has a code in exactly ONE unit. People with codes in
'  more than one unit are left blank for Graeme to fill in by hand
'  (SystemCheck and Needs Attention list them). Nothing else is written:
'  a backup copy is saved first and names / hours are checked after.
' --------------------------------------------------------------------
Private Sub HomeUnitValidation(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < R_FIRST Then Exit Sub
    Dim lst As String: lst = Join(DeptList(), ",")
    With ws.Range(ws.Cells(R_FIRST, C_HOME), ws.Cells(lastRow, C_HOME)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=lst
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "Home Unit"
        .ErrorMessage = "Pick one of the eight units from the drop-down, or leave it blank."
        .ShowError = True
    End With
End Sub

Public Sub SetupHomeUnit()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then MsgBox "No employees found on the roster.", vbExclamation, "Call Out Roster": Exit Sub

    ' column B must be free (it has held nothing since the OT Log migration)
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, UnitColLast())).Value
    Dim i As Long, stray As String, hdr As String
    hdr = UCase$(Trim$(SafeStr(ws.Cells(2, C_HOME).Value)))
    For i = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(i, C_HOME))) <> "" Then
            If InStr(1, "," & Join(DeptList(), ",") & ",", "," & Trim$(SafeStr(v(i, C_HOME))) & ",", vbTextCompare) = 0 Then
                stray = stray & "  B" & (R_FIRST + i - 1) & ": " & SafeStr(v(i, C_HOME)) & vbNewLine
            End If
        End If
    Next i
    If stray <> "" Or (hdr <> "" And hdr <> "HOME UNIT") Then
        MsgBox "Nothing was changed. Column B isn't free - it holds:" & vbNewLine & _
               IIf(hdr <> "" And hdr <> "HOME UNIT", "  B2: " & ws.Cells(2, C_HOME).Value & vbNewLine, "") & stray & vbNewLine & _
               "Clear those cells (or send " & PRIMARY_ADMIN_NAME & " a screenshot), then run SetupHomeUnit again.", vbCritical, "Home Unit"
        Exit Sub
    End If

    ' what would be filled in
    Dim fills As Long, blanks As String, nBlank As Long, u As String, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" And Trim$(SafeStr(v(i, C_HOME))) = "" Then
            u = SingleUnitFor(v, i)
            If u <> "" Then
                fills = fills + 1
            Else
                nBlank = nBlank + 1
                If nBlank <= 12 Then blanks = blanks & "  " & nm & vbNewLine
            End If
        End If
    Next i
    If nBlank > 12 Then blanks = blanks & "  ...and " & (nBlank - 12) & " more" & vbNewLine
    If MsgBox("Set up the Home Unit column (B)?" & vbNewLine & vbNewLine & _
              "- header and drop-down of the eight units" & vbNewLine & _
              "- " & fills & " people with a code in exactly one unit are filled in" & vbNewLine & _
              "- " & nBlank & " with codes in more than one unit (or none) are left blank for you:" & vbNewLine & blanks & vbNewLine & _
              "A backup copy is saved first. Nothing else on the roster is changed.", _
              vbOKCancel + vbQuestion, "Home Unit") <> vbOK Then Exit Sub

    ' ---- backup ----
    Dim bk As String
    bk = EnsureSlash(ThisWorkbook.Path) & "Backup before Home Unit " & Format$(Now, "yyyy-mm-dd hhmm") & " - " & ThisWorkbook.Name
    On Error Resume Next
    SaveRestrictedCopy bk                          ' v25 (I2)
    If Err.Number <> 0 Then
        Dim e As String: e = Err.Description
        On Error GoTo 0
        MsgBox "Couldn't save the backup copy, so nothing was changed:" & vbNewLine & e, vbCritical, "Call Out Roster"
        Exit Sub
    End If
    On Error GoTo 0

    ' ---- write (names, hours and codes untouched) ----
    Dim beforeNames As Variant, beforeTot As Variant
    beforeNames = ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)).Value2
    beforeTot = ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, C_TOTAL)).Value2
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Failed
    UnlockSheet ws
    ws.Cells(2, C_HOME).Value = "HOME UNIT"
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" And Trim$(SafeStr(v(i, C_HOME))) = "" Then
            u = SingleUnitFor(v, i)
            If u <> "" Then ws.Cells(R_FIRST + i - 1, C_HOME).Value = u
        End If
    Next i
    HomeUnitValidation ws, lastRow
    With ws.Range(ws.Cells(R_FIRST, C_HOME), ws.Cells(lastRow, C_HOME))
        .Font.Name = "Segoe UI": .Font.Size = 9: .Font.Color = INK
        .HorizontalAlignment = xlCenter
        .Locked = True                             ' v26.5: nothing on the roster is typed into (Employee > Edit)
    End With
    With ws.Cells(2, C_HOME)
        .Interior.Color = NAVY: .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 8
        .WrapText = True: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Columns(C_HOME).ColumnWidth = 14
    ws.Columns(C_HOME).Hidden = False
    Application.Calculate

    ' ---- verify ----
    Dim afterNames As Variant, afterTot As Variant, mism As String
    afterNames = ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)).Value2
    afterTot = ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, C_TOTAL)).Value2
    For i = 1 To UBound(v, 1)
        If SafeStr(beforeNames(i, 1)) <> SafeStr(afterNames(i, 1)) Then mism = mism & "  row " & (R_FIRST + i - 1) & ": name changed" & vbNewLine
        If Abs(NumOr0(beforeTot(i, 1)) - NumOr0(afterTot(i, 1))) > 0.0001 Then mism = mism & "  row " & (R_FIRST + i - 1) & ": Total Hours changed" & vbNewLine
    Next i
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
    If mism <> "" Then
        MarkNoSave
        MsgBox "STOP - the roster changed in a way it shouldn't have:" & vbNewLine & mism & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Your backup is at:" & vbNewLine & bk, vbCritical, "Call Out Roster"
        Exit Sub
    End If
    On Error Resume Next
    RefreshGradeSnapshot                           ' v21: the grade-change log keys on names + codes
    MarkCalDirty                                   ' Needs Attention lists the blanks
    On Error GoTo 0
    SaveIfWritable
    MsgBox "Home Unit set up: " & fills & " filled in, " & nBlank & " left blank for you." & vbNewLine & vbNewLine & _
           "Set the blank ones with Employee > Edit (double-click the row; Needs Attention lists them)." & vbNewLine & _
           "Backup saved as:" & vbNewLine & bk, vbInformation, "Home Unit"
    Exit Sub

Failed:
    Dim em As String: em = Err.Description
    LogError "SetupHomeUnit", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
    MarkNoSave
    MsgBox "Home Unit setup stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Your backup is at:" & vbNewLine & bk, vbCritical, "Call Out Roster"
End Sub

' Jump to a person's Home Unit cell (from Needs Attention)
Public Sub GoToHomeUnit(ByVal empName As String)
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), Trim$(empName), vbTextCompare) = 0 Then
            ws.Activate
            On Error Resume Next
            If ws.Rows(r).Hidden Then                 ' hidden by a sort button's filter: show everyone
                UnlockSheet ws
                If ws.FilterMode Then ws.ShowAllData
                ws.Rows(r).Hidden = False
                ProtectSheet ws
            End If
            Application.Goto ws.Cells(r, C_HOME), False
            On Error GoTo 0
            If IsAdmin() Then OpenEmployeeEdit empName   ' v26.5: column B is locked; the Employee pop-up's Edit sets it
            Exit Sub
        End If
    Next r
End Sub


' --------------------------------------------------------------------
'  FORMS - called at the end of each form's UserForm_Initialize
' --------------------------------------------------------------------
Public Sub ApplyFormTheme(ByVal frm As Object)
    Dim ctl As Object
    On Error Resume Next
    frm.BackColor = RGB(250, 251, 253)
    For Each ctl In frm.Controls
        ctl.Font.Name = "Segoe UI"
        Select Case TypeName(ctl)
            Case "Label"
                ctl.BackStyle = 0                       ' transparent
                If ctl.ForeColor = &H80000012 Then ctl.ForeColor = INK
            Case "ListBox", "ComboBox", "TextBox"
                ctl.BorderStyle = 1
                ctl.BorderColor = RGB(200, 206, 215)
                ctl.SpecialEffect = 0
            Case "CommandButton"
                If ctl.Name = "btnLoad" Or ctl.Name = "btnSubmit" Then
                    ctl.BackColor = NAVY
                    ctl.ForeColor = vbWhite
                End If
        End Select
    Next ctl
    On Error GoTo 0
End Sub

' v25.5 (briefs/v25.5.md item 2, A2): every pop-up's layout x UI_SCALE, once, as the LAST line of its Initialize:
' each control's position, size and point size (the designer controls too), a list's column widths, and the
' form itself. A size a form writes later (Load Posting's short layout, Add a Need's height) goes through UiPx.
' The call-out pop-up scales itself (Px in its paste-in) and does not call this.
' v27 (item 20, B14): the scale is PER LOGIN - a hidden name UiScale_<login> (the Desk's "Larger type" button writes
' it; 1 or UI_LARGER_SCALE); UI_SCALE, the constant, is the default for a login with none. Every ScaleForm / UiPx /
' Px reads it here, once per pop-up. formHeight > 0 = the form's height at scale 1: the scale is CLAMPED so the
' form stays on the screen (review MED-10: 615 pt x 1.2 is off a 1080p screen at 125 %). (UI_LARGER_SCALE sits in
' the declarations block at the top.)
Public Function UiScale(Optional ByVal formHeight As Double = 0) As Double
    UiScale = UI_SCALE
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(LoginNameKey("UiScale"))
    Dim v As Double: v = UI_SCALE
    If Not nmObj Is Nothing Then v = Val(Replace(Replace(nmObj.RefersTo, "=", ""), ",", "."))
    If v < 0.8 Or v > 2 Then v = UI_SCALE
    If formHeight > 0 Then
        Dim room As Double: room = Application.UsableHeight - 40
        If room > 200 And formHeight * v > room Then v = room / formHeight
        If v < 1 Then v = 1
    End If
    UiScale = v
End Function

' The status cell's current view text (the part after "View: "), "All" when none
Public Function CurrentViewText() As String
    CurrentViewText = "All"
    On Error Resume Next                            ' never raises
    Dim cur As String: cur = SafeStr(RosterSheet().Range(STATUS_CELL).Value)
    Dim p As Long: p = InStr(1, cur, "View: ")
    If p > 0 Then CurrentViewText = Trim$(Mid$(cur, p + 6))
End Function

' The Desk's "Larger type": this login's pop-ups open at UI_LARGER_SCALE (or back to the default); the next one shows it
Public Sub SetUiScaleForMe(ByVal larger As Boolean)
    On Error Resume Next                            ' never raises
    Dim nmKey As String: nmKey = LoginNameKey("UiScale")
    Dim v As Double: v = IIf(larger, UI_LARGER_SCALE, UI_SCALE)
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(nmKey)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=nmKey, RefersTo:="=" & Replace(CStr(v), ",", "."), Visible:=False
    Else
        nmObj.RefersTo = "=" & Replace(CStr(v), ",", ".")
    End If
End Sub

Public Sub ScaleForm(ByVal frm As Object)
    Dim k As Double: k = UiScale(frm.Height)         ' clamped to the screen for this form's height
    mFormScale = k                                   ' (review M6) the sizes written after this one use the same factor
    If k = 1 Then Exit Sub
    Dim ctl As Object
    On Error Resume Next
    For Each ctl In frm.Controls
        ctl.Left = ctl.Left * k
        ctl.Top = ctl.Top * k
        ctl.Width = ctl.Width * k
        ctl.Height = ctl.Height * k
        ctl.Font.Size = ctl.Font.Size * k
        If TypeName(ctl) = "ListBox" Or TypeName(ctl) = "ComboBox" Then ctl.ColumnWidths = ScaleWidths(CStr(ctl.ColumnWidths))
    Next ctl
    frm.Width = frm.Width * k
    frm.Height = frm.Height * k
    On Error GoTo 0
End Sub

' "22 pt;182 pt;-1 pt" (as a list reads its ColumnWidths back) -> each width x UI_SCALE; -1 (automatic) and "" stay
Public Function ScaleWidths(ByVal widths As String) As String
    If Trim$(widths) = "" Then Exit Function
    Dim parts As Variant, i As Long, v As Double
    parts = Split(widths, ";")
    For i = LBound(parts) To UBound(parts)
        v = Val(Replace(Trim$(CStr(parts(i))), ",", "."))      ' a comma-decimal locale reads back "22,5 pt" (review)
        If v > 0 Then parts(i) = CStr(CLng(v * FormScale())) & " pt" Else parts(i) = Trim$(CStr(parts(i)))   ' whole points: no decimal separator, whatever the locale (re-review)
    Next i
    ScaleWidths = Join(parts, ";")
End Function

' a size written AFTER ScaleForm ran goes through this
Public Function UiPx(ByVal v As Double) As Double
    UiPx = v * FormScale()
End Function

' (review M6) the factor of the form ScaleForm last scaled (one pop-up at a time), else this login's scale
Private Function FormScale() As Double
    If mFormScale >= 1 Then FormScale = mFormScale Else FormScale = UiScale()
End Function

' --------------------------------------------------------------------
'  KEYBOARD SHORTCUTS (v25.5, A5): Ctrl+Shift+S sick call, Ctrl+Shift+C call-out, Ctrl+Shift+K calendar,
'  Ctrl+Shift+D the Desk. On while this workbook is active (open / activate), off on deactivate and close -
'  OnKey is Excel-wide, so another workbook must never see them. Each key runs the macro its button runs
'  (the gates inside refuse as they do for a click). Listed on the Desk's footer row and the Help tab.
' --------------------------------------------------------------------
Public Sub SetShortcuts(ByVal on_ As Boolean)
    On Error Resume Next                            ' never raises
    Dim keys As Variant, macs As Variant, i As Long
    keys = Array("^+s", "^+S", "^+c", "^+C", "^+k", "^+K", "^+d", "^+D")
    macs = Array("modUI.KeySick", "modUI.KeySick", "modUI.KeyCallout", "modUI.KeyCallout", _
                 "modUI.KeyCalendar", "modUI.KeyCalendar", "modUI.KeyDesk", "modUI.KeyDesk")
    Dim wbq As String: wbq = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"   ' this workbook's macro, not a practice copy's (review)
    For i = LBound(keys) To UBound(keys)
        If on_ Then
            Application.OnKey CStr(keys(i)), wbq & CStr(macs(i))
        Else
            Application.OnKey CStr(keys(i))
        End If
    Next i
End Sub

Public Sub KeySick()
    On Error GoTo Oops
    ShowSickCallForm
    Exit Sub
Oops:
    ReportError "KeySick"
End Sub

Public Sub KeyCallout()
    On Error GoTo Oops
    ShowOvertimeForm
    Exit Sub
Oops:
    ReportError "KeyCallout"
End Sub

Public Sub KeyCalendar()
    On Error GoTo Oops
    ShowCalendar
    Exit Sub
Oops:
    ReportError "KeyCalendar"
End Sub

' the Desk: the "< Desk" switch for an admin on the full layout; the Desk sheet for a Desk admin or a
' supervisor; nothing for a plain user (no Desk)
Public Sub KeyDesk()
    On Error GoTo Oops
    If gAdminUnlocked And Not gDeskMode Then
        ShowDesk
    ElseIf SheetExists(DESK_SHEET) Then
        If ThisWorkbook.Worksheets(DESK_SHEET).Visible = xlSheetVisible Then ShowDeskSheet
    End If
    Exit Sub
Oops:
    ReportError "KeyDesk"
End Sub

' --------------------------------------------------------------------
'  SHARED FORM HELPERS (v27.1): the control builders the ten code-built
'  forms used to copy. A form keeps a one-line Private wrapper that adds
'  its own name prefix ("cv_", "sc_" ...) - or, in the four forms that
'  file controls into groups, the bookkeeping - so no call site changes.
'  The bodies are the forms' copies with Me -> frm, property writes in
'  the same order. The two designer forms keep their own helpers.
' --------------------------------------------------------------------

' a control on frm named ctlName (the full name, prefix included), placed and sized
Public Function FormAddCtl(ByVal frm As Object, ByVal progId As String, ByVal ctlName As String, ByVal L As Single, _
                           ByVal T As Single, ByVal W As Single, ByVal H As Single) As Object
    Dim o As Object
    Set o = frm.Controls.Add(progId, ctlName, True)
    o.Left = L: o.Top = T: o.Width = W: o.Height = H
    Set FormAddCtl = o
End Function

' a label on frm: FormAddCtl at height H (15 in most forms, 16 in frmHistory / frmTableEdit), then its caption
Public Function FormAddLabel(ByVal frm As Object, ByVal ctlName As String, ByVal cap As String, ByVal L As Single, _
                             ByVal T As Single, ByVal W As Single, Optional ByVal H As Single = 15) As Object
    Dim o As Object
    Set o = FormAddCtl(frm, "Forms.Label.1", ctlName, L, T, W, H)
    o.Caption = cap
    Set FormAddLabel = o
End Function

' a combo (or list) box emptied and filled with items, each as text; anything but an array leaves it empty
Public Sub FormFillCombo(ByVal cb As Object, ByVal items As Variant)
    Dim i As Long
    cb.Clear
    If Not IsArray(items) Then Exit Sub
    For i = LBound(items) To UBound(items)
        cb.AddItem CStr(items(i))
    Next i
End Sub

' the date typed in a form's date box, through the one parser (modDates.ParseTypedDate)
Public Function FormPickedDate(ByVal tb As Object, ByRef d As Date) As Boolean
    FormPickedDate = ParseTypedDate(tb.Text, d)
End Function
