Attribute VB_Name = "modCalendar"
Option Explicit

' ====================================================================
'  COVERAGE CALENDAR  (replaces the monthly vacation-calendar workbooks)
'
'  Absences  (tblAbsence)  one row per absence RANGE: person, unit, type,
'            start, end (blank = open-ended, e.g. STD), hours per shift
'            (blank = full shift), times note ("7-3"). The rotation decides
'            which days / nights it lands on.
'  Coverage  (tblCoverage) one row per date + shift line: who covers an
'            absence, or a need with nobody off (PHA, support op ...).
'            Postings, call-outs and shift switches add lines automatically.
'  Calendar  sheet drawn from the two tables: month + unit picker,
'            yellow = days, blue = nights, Print / PDF. Never typed into.
'
'  Privacy: STD, parental and LOA always show as "Leave" on the calendar.
'
'  v27.1 (the refactor, moved unchanged): the two-week board is in modBoard, minimum staffing in modMinStaff,
'  the old-calendar import in modCalImport, the Needs Attention tab in modNeeds, the line menus in
'  modCalMenus. This module keeps the tables, who works when, the slots, the writes, the calendar sheet and its
'  drawing (DrawCalendar and its steps), the line details, change tracking, the theme and the event routing.
' ====================================================================

Public Const AB_SHEET As String = "Absences"
Public Const AB_TABLE As String = "tblAbsence"
Public Const CV_SHEET As String = "Coverage"
Public Const CV_TABLE As String = "tblCoverage"
Public Const CAL_SHEET As String = "Calendar"
Public Const MS_CALLNOTE As Long = 5          ' v27 (review of item 25, M2): the Min Staff table's Call-out note column (added on first write)
Public Const BOARD_SHEET As String = "Board"        ' v26: the two-week board (BoardPrint), output only
Public Const BOARD_DAYS As Long = 14                 ' v26: today .. today + 13
Private Const MAP_SHEET As String = "CalMap"
Public Const FULL_SHIFT As Double = 12
Public Const CAL_VERSION As String = "Calendar v29.0.2 (Oct 28)"
' v25.1.2: draw-time measurement (the v25.2 plan's step 1): seconds spent inside SlotsFor during the current
' draw, read by DrawCalendar / DrawNeeds for their breakdown lines (v25.5: the notes share is always 0 - no
' hover notes are made; the figure stays in the line so the numbers read the same)
Public mTimeSlots As Double
Private mTimeNotes As Double
Private mHomeMap As Object               ' v25.2 (item 4): name -> Home Unit, filled by LoadCache
Public Const NA_SHEET As String = "Needs Attention"
Public Const NA_KEY As Long = 26           ' hidden key column on the Needs Attention tab
Public Const CAL_FONT As String = "Segoe UI"
Public Const CAL_FONT_B As String = "Segoe UI Semibold"
Private Const COL_W As Double = 22          ' width of each half of a day
Private Const CAL_HINT_ROW_H As Double = 40 ' v25.5: row 3 holds the hint or a clicked line's details, wrapped (was 28)

' tblAbsence
Public Const AB_ID As Long = 1
Public Const AB_EMP As Long = 2
Public Const AB_UNIT As Long = 3
Public Const AB_TYPE As Long = 4
Public Const AB_START As Long = 5
Public Const AB_END As Long = 6
Public Const AB_HOURS As Long = 7
Public Const AB_TIMES As Long = 8
Public Const AB_NOTE As Long = 9
Public Const AB_STATUS As Long = 10
Public Const AB_USER As Long = 11
Public Const AB_STAMP As Long = 12
Public Const AB_VOIDNOTE As Long = 13
Public Const AB_HOME As Long = 14       ' v22: Home Unit locked in when the booking is made (leave top-up peers)
Public Const AB_GROUP As Long = 15      ' v22: LV-#### linking a leave's going and return
Public Const AB_SHIFT As Long = 16      ' v22: Days / Nights / Both (sick calls; converted imports)
Public Const AB_N As Long = 16

' tblCoverage
Public Const CV_ID As Long = 1
Public Const CV_DATE As Long = 2
Public Const CV_SHIFT As Long = 3
Public Const CV_UNIT As Long = 4
Public Const CV_ABS As Long = 5
Public Const CV_NEED As Long = 6
Public Const CV_TYPE As Long = 7
Public Const CV_BY As Long = 8
Public Const CV_HOURS As Long = 9
Public Const CV_SOURCE As Long = 10
Public Const CV_STATUS As Long = 11
Public Const CV_USER As Long = 12
Public Const CV_STAMP As Long = 13
Public Const CV_NOTE As Long = 14
Public Const CV_N As Long = 14

Public Type CalSlot
    Key As String           ' "A:AB-0001" (absence) or "C:CV-0001" (coverage line)
    LeftText As String
    RightText As String
    TypeCode As String
    IsOpen As Boolean
End Type

' set by Post Line so the new posting links to that calendar line
Public gCalPostKey As String
Public gCalPostDate As Date
Public gCalPostShift As String

' calendar menu / redraw state
Public mPostV As Variant           ' tblPostings snapshot for the details text and the attention box
Public mDrawing As Boolean
Private mPublicDraw As Boolean      ' v26.5: True while the calendar is drawn as members may see it (the printed month: CalPrintPdf)
Private mEventsSeen As Boolean      ' True once ThisWorkbook's tab-switch / change events have fired
Private mDetailStamp As String      ' v25.5: the change counter + draw context the caches were loaded under for a click's details


' cache for one operation (tables + roster shifts)
Public mAbs As Variant
Public mCov As Variant
Private mRst As Variant                 ' v27 (item 21): the Restrictions table (Empty when none)
Public Const RS_ID As Long = 1          ' v27 (item 21): tblRestrictions - RS-nnnn, Employee, From, To (blank = open), Reason, Status, Entered By, Entered At, Void Note
Public Const RS_EMP As Long = 2
Public Const RS_FROM As Long = 3
Public Const RS_TO As Long = 4
Public Const RS_REASON As Long = 5
Public Const RS_STATUS As Long = 6
Public Const RS_USER As Long = 7
Public Const RS_STAMP As Long = 8
Public Const RS_VOIDNOTE As Long = 9
Public Const RS_N As Long = 9
Public mShiftMap As Object
Private mOt As Object               ' v26.5.5 (item 9): "name|yyyymmdd|D/N" -> OT hours accepted / awarded (non-void log rows)
Private mPendOff(1 To 2) As String  ' v27.5 (item 1): "name|yyyymmdd|Days/Nights" x 2 - the shifts a switch being saved gives up (SetPendingGiveUp)
Private mTestMode As Boolean        ' v21: SelfTest feeds mAbs from memory - the tables are never read

' ====================================================================
'  LISTS
' ====================================================================
Public Function AbsenceTypes() As Variant
    AbsenceTypes = Array("VAC - Vacation", "PER - Personal day (unpaid)", "TRN - Training", "LOA - Leave of absence", _
                         "STD - Short-term disability", "PAR - Parental leave", _
                         "WSIB - Workplace injury leave", "LTD - Long-term disability", _
                         "SIC - Sick day (past)", "OTH - Other")
End Function

' Leave of any kind: masked as "Leave" everywhere outside the Absences sheet (v22 adds WSIB, LTD)
Public Function IsLeaveCode(ByVal code As String) As Boolean
    Select Case TypeCode(code)                 ' takes a code or the full "STD - ..." text
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": IsLeaveCode = True
    End Select
End Function

' Leave types the Leave pop-up offers (v22). "OLV" = other leave (not the
' Book Time Off "OTH", which is not a leave and gets no top-up).
Public Function LeaveTypes() As Variant
    LeaveTypes = Array("PAR - Parental leave", "STD - Short-term disability", "LOA - Leave of absence", _
                       "WSIB - Workplace injury leave", "LTD - Long-term disability", "OLV - Other leave")
End Function

Public Function NeedTypes() As Variant
    NeedTypes = Array("PHA - PHA coverage", "SUP - Support / extra operator", "TRN - Training", "OTH - Other")
End Function

' Calendars: "Name|unit,unit"  (* = every unit). Edit to suit.
Public Function CalGroups() As Variant
    CalGroups = Array("Phosphine & Derivatives|Phosphine,Derivatives", _
                      "Cypure|Cypure", _
                      "Fumigation (Material Handler)|Material Handler", _
                      "QC Lab|QC Derivatives", _
                      "QC Gases|QC Gases", _
                      "Utilities|Utilities", _
                      "SBP|SBP", _
                      "All units|*")
End Function

Public Function CalGroupNames() As Variant
    Dim g As Variant, a() As String, i As Long
    g = CalGroups()
    ReDim a(0 To UBound(g))
    For i = 0 To UBound(g)
        a(i) = Split(g(i), "|")(0)
    Next i
    CalGroupNames = a
End Function

Public Function UnitInGroup(ByVal unit As String, ByVal groupName As String) As Boolean
    Dim g As Variant, i As Long, parts As Variant, u As Variant
    If groupName = "" Then UnitInGroup = True: Exit Function
    g = CalGroups()
    For i = 0 To UBound(g)
        parts = Split(g(i), "|")
        If StrComp(parts(0), groupName, vbTextCompare) = 0 Then
            If parts(1) = "*" Then UnitInGroup = True: Exit Function
            For Each u In Split(parts(1), ",")
                If StrComp(Trim$(u), Trim$(unit), vbTextCompare) = 0 Then UnitInGroup = True: Exit Function
            Next u
            Exit Function
        End If
    Next i
    ' not a group name - treat as a single unit
    UnitInGroup = (StrComp(unit, groupName, vbTextCompare) = 0)
End Function

' The calendar a unit belongs to (first group that lists it)
Public Function GroupOfUnit(ByVal unit As String) As String
    Dim g As Variant, i As Long, parts As Variant
    g = CalGroups()
    For i = 0 To UBound(g)
        parts = Split(g(i), "|")
        If parts(1) <> "*" Then
            If UnitInGroup(unit, CStr(parts(0))) Then GroupOfUnit = parts(0): Exit Function
        End If
    Next i
    GroupOfUnit = unit
End Function

' "WSIB - ..." -> "WSIB"; everything else is its first three letters
Public Function TypeCode(ByVal s As String) As String
    s = UCase$(Trim$(s))
    If Left$(s, 4) = "WSIB" Then TypeCode = "WSIB" Else TypeCode = Left$(s, 3)
End Function

' What the calendar shows (never "STD" / "PAR" / "WSIB" / "LTD"; sick calls say "Sick")
Public Function TypeTag(ByVal code As String) As String
    Select Case UCase$(code)
        Case "VAC": TypeTag = "Vac"
        Case "TRN": TypeTag = "Trn"
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": TypeTag = "Leave"
        Case "SIC": TypeTag = "Sick"
        Case "PER": TypeTag = "Personal"
        Case "SS": TypeTag = "S/S"
        Case "PHA": TypeTag = "PHA"
        Case "SUP": TypeTag = "Support"
        Case Else: TypeTag = "Off"
    End Select
End Function

Public Function TypeColor(ByVal code As String) As Long
    Select Case UCase$(code)
        Case "VAC": TypeColor = RGB(21, 101, 192)              ' blue
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV", "LV": TypeColor = RGB(106, 27, 154) ' purple ("LV" = PublicType's leave, v26.5)
        Case "SIC": TypeColor = RGB(230, 81, 0)                ' orange
        Case "PER": TypeColor = RGB(69, 90, 100)               ' blue-grey (v22.1 personal day)
        Case "TRN": TypeColor = RGB(0, 121, 107)               ' teal
        Case "PHA": TypeColor = RGB(46, 125, 50)               ' green
        Case "SUP": TypeColor = RGB(93, 64, 55)                ' brown
        Case "SS": TypeColor = RGB(194, 24, 91)                ' magenta
        Case Else: TypeColor = RGB(97, 97, 97)
    End Select
End Function

' "Campbell, Fred" -> "Fred C"
Public Function CalName(ByVal nm As String) As String
    Dim p As Long: p = InStr(nm, ",")
    If p > 0 Then
        CalName = Trim$(Mid$(nm, p + 1)) & " " & UCase$(Left$(Trim$(nm), 1))
    Else
        CalName = Trim$(nm)
    End If
End Function

' Number picker that works with long prompts (Application.InputBox is
' limited to 255 characters). Returns -1 if cancelled / not valid.
Public Function AskNumber(ByVal prompt As String, ByVal title As String, ByVal dflt As Long, _
                           ByVal lo As Long, ByVal hi As Long) As Long
    If Len(prompt) > 1000 Then prompt = Left$(prompt, 990) & vbNewLine & "..."
    Dim s As String: s = Trim$(VBA.InputBox(prompt, title, CStr(dflt)))
    AskNumber = -1
    If s = "" Or Not IsNumeric(s) Then Exit Function
    If CDbl(s) <> Int(CDbl(s)) Then Exit Function
    If CLng(s) < lo Or CLng(s) > hi Then
        MsgBox "Pick a number from " & lo & " to " & hi & ".", vbExclamation, "Call Out Roster"
        Exit Function
    End If
    AskNumber = CLng(s)
End Function

Public Function Clip(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then Clip = Left$(s, n - 1) & "~" Else Clip = s
End Function

Public Function HrsTag(ByVal h As Variant) As String
    If IsNumeric(h) And Not IsEmpty(h) Then
        If CDbl(h) > 0 And CDbl(h) < FULL_SHIFT Then HrsTag = " " & CDbl(h) & "h"
    End If
End Function

' ====================================================================
'  TABLES
' ====================================================================
Public Function AbsTable(Optional ByVal create As Boolean = False) As ListObject
    Set AbsTable = GetOrMakeTable(AB_SHEET, AB_TABLE, create, Array("Absence ID", "Employee", "Unit", "Type", _
        "Start", "End", "Hours / Shift", "Times", "Note", "Status", "Logged By", "Timestamp", "Void Note", _
        "Home Unit", "Leave Group", "Shift"), _
        Array(10, 24, 14, 8, 13, 13, 8, 9, 30, 8, 11, 18, 30, 14, 10, 8), RGB(31, 56, 100))
End Function

' True when tblAbsence has the v22 columns (Home Unit, Leave Group, Shift)
Public Function AbsenceColumnsCurrent() As Boolean
    Dim lo As ListObject: Set lo = AbsTable(False)
    If lo Is Nothing Then AbsenceColumnsCurrent = True: Exit Function    ' created fresh with all columns
    AbsenceColumnsCurrent = (lo.ListColumns.Count >= AB_N)
End Function

' Alt+F8 > UpgradeAbsenceColumns (once, v22): appends the three columns.
' Nothing moves - they go on the END - but it still follows the verified
' pattern: backup copy, change, check row count and every roster total.
Public Sub UpgradeAbsenceColumns()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim lo As ListObject: Set lo = AbsTable(False)
    If lo Is Nothing Then
        MsgBox "There is no Absences table yet - it will be created with every column the first time time off is booked. Nothing to do.", vbInformation, "Call Out Roster"
        Exit Sub
    End If
    If lo.ListColumns.Count >= AB_N Then MsgBox "The Absences table already has the v22 columns.", vbInformation, "Call Out Roster": Exit Sub
    If MsgBox("Add the Home Unit, Leave Group and Shift columns to the Absences table?" & vbNewLine & vbNewLine & _
              "They go on the end, so nothing moves. A backup copy is saved first and the row count and roster totals are checked after.", _
              vbOKCancel + vbQuestion, "Absences upgrade") <> vbOK Then Exit Sub
    Dim bk As String
    bk = EnsureSlash(ThisWorkbook.Path) & "Backup before Absences upgrade " & Format$(Now, "yyyy-mm-dd hhmm") & " - " & ThisWorkbook.Name
    On Error Resume Next
    SaveRestrictedCopy bk                          ' v25 (I2)
    If Err.Number <> 0 Then
        Dim e As String: e = Err.Description
        On Error GoTo 0
        MsgBox "Couldn't save the backup copy, so nothing was changed:" & vbNewLine & e, vbCritical, "Call Out Roster"
        Exit Sub
    End If
    On Error GoTo 0
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim before As Variant, nRows As Long
    Application.Calculate
    If lastRow >= R_FIRST Then before = rs.Range(rs.Cells(R_FIRST, C_TOTAL), rs.Cells(lastRow, C_PERIODX)).Value2   ' G:AT, always 2-D
    nRows = lo.ListRows.Count
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim names As Variant: names = Array("Home Unit", "Leave Group", "Shift")
    Dim i As Long, em As String
    On Error GoTo Failed
    UnlockSheet ws
    For i = lo.ListColumns.Count + 1 To AB_N
        lo.ListColumns.Add.Name = names(i - AB_VOIDNOTE - 1)
    Next i
    ws.Columns(AB_HOME).ColumnWidth = 14: ws.Columns(AB_GROUP).ColumnWidth = 10: ws.Columns(AB_SHIFT).ColumnWidth = 8
    ProtectSheet ws
    Application.Calculate
    Dim problems As String
    If lo.ListRows.Count <> nRows Then problems = problems & "- absence rows: " & nRows & " before, " & lo.ListRows.Count & " after" & vbNewLine
    If lo.ListColumns.Count <> AB_N Then problems = problems & "- column count is " & lo.ListColumns.Count & ", expected " & AB_N & vbNewLine
    If lastRow >= R_FIRST Then
        Dim after As Variant, j As Long: after = rs.Range(rs.Cells(R_FIRST, C_TOTAL), rs.Cells(lastRow, C_PERIODX)).Value2
        For i = 1 To UBound(before, 1)
            For j = 1 To UBound(before, 2)
                If SafeStr(after(i, j)) <> SafeStr(before(i, j)) Then
                    problems = problems & "- roster value changed on row " & (R_FIRST + i - 1) & ", column " & (C_TOTAL + j - 1) & vbNewLine
                    Exit For
                End If
            Next j
            If problems <> "" Then Exit For
        Next i
    End If
    If problems <> "" Then
        MarkNoSave
        MsgBox "The upgrade did NOT verify:" & vbNewLine & problems & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Call Out Roster"
        Exit Sub
    End If
    ResetCalCacheAll
    SaveIfWritable
    MsgBox "Absences table upgraded: " & lo.ListColumns.Count & " columns, " & nRows & " rows, roster totals unchanged." & vbNewLine & "Backup: " & bk, vbInformation, "Call Out Roster"
    Exit Sub
Failed:
    em = Err.Description
    LogError "UpgradeAbsenceColumns", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MarkNoSave
    MsgBox "The upgrade stopped part-way: " & em & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Call Out Roster"
End Sub

Public Sub ResetCalCacheAll()
    If mTestMode Then Exit Sub              ' SelfTest's fixtures stay until it clears them
    mAbs = Empty: mCov = Empty: Set mShiftMap = Nothing: Set mHomeMap = Nothing     ' v25.2: and the home-unit map
    Set mOt = Nothing                                                                ' v26.5.5: and the OT-worked map
End Sub

Public Function CovTable(Optional ByVal create As Boolean = False) As ListObject
    Set CovTable = GetOrMakeTable(CV_SHEET, CV_TABLE, create, Array("Line ID", "Date", "Shift", "Unit", _
        "Absence ID", "Off / Need", "Type", "Covered By", "Hours", "Source", "Status", "Logged By", "Timestamp", "Note"), _
        Array(10, 14, 7, 14, 10, 26, 6, 22, 6, 14, 8, 11, 18, 26), RGB(0, 121, 107))
End Function

Private Function GetOrMakeTable(ByVal sh As String, ByVal tn As String, ByVal create As Boolean, _
                                ByVal hdr As Variant, ByVal widths As Variant, ByVal tabCol As Long) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(sh) Then
        Set ws = ThisWorkbook.Worksheets(sh)
        On Error Resume Next
        Set lo = ws.ListObjects(tn)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sh
            RelockStructure
            SetTabColor ws, tabCol                      ' v25.1
        End If
        UnlockSheet ws
        Dim n As Long: n = UBound(hdr) + 1
        ws.Range("A1").Resize(1, n).Value = hdr
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, n), , xlYes)
        lo.Name = tn
        lo.TableStyle = "TableStyleMedium2"
        Dim i As Long
        For i = 1 To n
            ws.Columns(i).ColumnWidth = widths(i - 1)
        Next i
        ws.Rows(1).WrapText = True
        Dim fc As FormatCondition
        Set fc = ws.Range("A2").Resize(20000, n).FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=$" & ColLetter(IIf(tn = AB_TABLE, AB_STATUS, CV_STATUS)) & "2=""VOID""")
        fc.Font.Strikethrough = True
        fc.Font.Color = RGB(150, 150, 150)
        ProtectSheet ws
    End If
    Set GetOrMakeTable = lo
End Function

Public Function TableData(ByVal lo As ListObject) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    TableData = RangeTo2D(lo.DataBodyRange)
End Function

' The Absences table as a 2-D array (v22.1: modAttendance reads it)
Public Function AbsTableData() As Variant
    AbsTableData = TableData(AbsTable())
End Function

Public Function NextId(ByVal lo As ListObject, ByVal prefix As String) As String
    Dim v As Variant: v = TableData(lo)
    Dim i As Long, k As Long, mx As Long, s As String
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            s = UCase$(Trim$(SafeStr(v(i, 1))))
            If Left$(s, Len(prefix) + 1) = prefix & "-" Then
                k = CLng(Val(Mid$(s, Len(prefix) + 2)))
                If k > mx Then mx = k
            End If
        Next i
    End If
    NextId = prefix & "-" & Format$(mx + 1, "0000")
End Function

Public Function AppendTo(ByVal lo As ListObject, ByVal data As Variant, ByVal nCols As Long) As Boolean
    Dim n As Long: n = UBound(data, 1)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim em As String
    On Error GoTo Fail
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, nCols).Value = data
    ProtectSheet ws
    AppendTo = True
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendTo", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Could not write to " & lo.Name & ": " & em, vbCritical, "Call Out Roster"
End Function

' row number (in the table body) of an ID, 0 if not found
Public Function RowOf(ByVal v As Variant, ByVal id As String) As Long
    Dim i As Long
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), id, vbTextCompare) = 0 Then RowOf = i: Exit Function
    Next i
End Function

' SelfTest (v21): absence rows in the tblAbsence layout (or Empty) replace
' the tables until ClearTestFixtures. Nothing on any sheet is read or written.
Public Sub SetTestFixtures(ByVal absRows As Variant)
    mTestMode = True
    mAbs = absRows
    mCov = Empty
    Set mShiftMap = CreateObject("Scripting.Dictionary")
    mShiftMap.CompareMode = 1
End Sub

Public Sub ClearTestFixtures()
    mTestMode = False
    mAbs = Empty
    mCov = Empty
    Set mShiftMap = Nothing
End Sub

' M1: the members' publish loads the cache once, then calls FilterNotFree
' and SlotsFor with useCache:=True for the whole sweep.
Public Sub RefreshCalCache()
    LoadCache
End Sub

Public Sub LoadCache()
    If mTestMode Then Exit Sub
    mAbs = TableData(AbsTable())
    mCov = TableData(CovTable())
    mRst = Empty                                             ' v27 (item 21): the restrictions, when the table exists
    Dim rlo As ListObject: Set rlo = RstTable(False)
    If Not rlo Is Nothing Then mRst = TableData(rlo)
    Set mOt = Nothing                                        ' v26.5.5 (item 9): the OT-worked map is read lazily (OtWorkedOn) from the log
    Set mShiftMap = CreateObject("Scripting.Dictionary")
    mShiftMap.CompareMode = 1
    Set mHomeMap = CreateObject("Scripting.Dictionary")      ' v25.2 (item 4, C7): name -> Home Unit, one read
    mHomeMap.CompareMode = 1
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_HOME), ws.Cells(lastRow, C_SHIFT)).Value
    Dim i As Long, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME - C_HOME + 1)))
        If nm <> "" Then
            mShiftMap(nm) = UCase$(Trim$(SafeStr(v(i, C_SHIFT - C_HOME + 1))))
            If Not mHomeMap.Exists(nm) Then mHomeMap(nm) = Trim$(SafeStr(v(i, 1)))     ' the first row wins, as HomeUnitOf (review)
        End If
    Next i
End Sub

' v25.2 (item 4, C7): the Home Unit from the cache LoadCache filled (HomeUnitOf in modLeave scans the roster
' cell by cell: Needs Attention called it per name and per planned move); a name not on the roster asks HomeUnitOf
Public Function HomeUnitCached(ByVal emp As String) As String
    If mHomeMap Is Nothing Then LoadCache
    If Not mHomeMap Is Nothing Then
        If mHomeMap.Exists(Trim$(emp)) Then HomeUnitCached = mHomeMap(Trim$(emp)): Exit Function
    End If
    HomeUnitCached = HomeUnitOf(emp)
End Function

' ====================================================================
'  WHO WORKS WHEN / WHO IS OFF
' ====================================================================
' "Days", "Nights" or "" (not scheduled). Day workers: Mon-Fri days.
' Unknown crews show on Days so an absence is never hidden.
Public Function EmpShiftOn(ByVal empName As String, ByVal d As Date) As String
    If mShiftMap Is Nothing Then LoadCache
    Dim s As String
    If mShiftMap.Exists(Trim$(empName)) Then s = mShiftMap(Trim$(empName))
    s = CrewOn(empName, d, s)                   ' v24: the crew they were on that day (crew history)
    EmpShiftOn = CrewShiftOn(s, d)              ' A-D rotation, A/B or C/D days only, DAYS = Mon-Fri
    If EmpShiftOn = "?" Then EmpShiftOn = "Days" ' unknown crew: show on Days so it's never hidden
End Function

Public Function AbsActiveOn(ByVal r As Long, ByVal d As Date) As Boolean
    If UCase$(SafeStr(mAbs(r, AB_STATUS))) = "VOID" Then Exit Function
    If Not IsDate(mAbs(r, AB_START)) Then Exit Function
    If Int(CDate(mAbs(r, AB_START))) > Int(d) Then Exit Function
    If IsDate(mAbs(r, AB_END)) Then
        If Int(CDate(mAbs(r, AB_END))) < Int(d) Then Exit Function
    End If
    AbsActiveOn = True
End Function

' ====================================================================
'  v27 (item 21, 1.1): OT RESTRICTED - tblRestrictions on the admin-only Restrictions sheet (made on first use).
'  A restricted person is NOT LISTED for call-outs and postings between From and To (open = until ended) and
'  is never charged (decided Oct 21: never charged while restricted); the calendar and the staffing count
'  ignore it (they still work their scheduled shifts); nothing of it is published (members see a shorter list).
'  RestrictOT / RestrictOTFor (admin: the Employee pop-up's Edit, Alt+F8) write a row; EndRestriction voids one.
' ====================================================================
Public Function RstTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(RS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(RS_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(RS_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = RS_SHEET
            RelockStructure
            SetTabColor ws, RGB(176, 0, 32)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, RS_N).Value = Array("Restriction ID", "Employee", "From", "To", "Reason", "Status", "Entered By", "Entered At", "Void Note")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, RS_N), , xlYes)
        lo.Name = RS_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(3).NumberFormat = "ddd mmm d, yyyy": ws.Columns(4).NumberFormat = "ddd mmm d, yyyy"
        ws.Columns(8).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ws.Columns(2).ColumnWidth = 24: ws.Columns(5).ColumnWidth = 40
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set RstTable = lo
End Function

' True while an active restriction covers d; why = its reason and dates
Public Function OtRestrictedOn(ByVal empName As String, ByVal d As Date, Optional ByRef why As String) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    why = ""
    If Not IsArray(mRst) Then Exit Function
    Dim r As Long
    For r = 1 To UBound(mRst, 1)
        If StrComp(Trim$(SafeStr(mRst(r, RS_EMP))), Trim$(empName), vbTextCompare) = 0 Then
            If UCase$(SafeStr(mRst(r, RS_STATUS))) <> "VOID" And IsDate(mRst(r, RS_FROM)) Then
                If Int(CDate(mRst(r, RS_FROM))) <= Int(d) Then
                    If Not IsDate(mRst(r, RS_TO)) Or Int(CDate(mRst(r, RS_TO))) >= Int(d) Then
                        why = SafeStr(mRst(r, RS_ID)) & ": " & SafeStr(mRst(r, RS_REASON)) & " (from " & Format$(CDate(mRst(r, RS_FROM)), "mmm d") & _
                              IIf(IsDate(mRst(r, RS_TO)), " to " & Format$(CDate(mRst(r, RS_TO)), "mmm d"), ", open") & ")"
                        OtRestrictedOn = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next r
End Function

' The active restrictions as of today, one per line (Needs Attention; "" = none). count = how many.
Public Function ActiveRestrictionsText(ByRef count As Long) As String
    count = 0
    If mShiftMap Is Nothing Then LoadCache
    If Not IsArray(mRst) Then Exit Function
    Dim r As Long, s As String
    For r = 1 To UBound(mRst, 1)
        If UCase$(SafeStr(mRst(r, RS_STATUS))) <> "VOID" And IsDate(mRst(r, RS_FROM)) Then
            If Not IsDate(mRst(r, RS_TO)) Or Int(CDate(mRst(r, RS_TO))) >= Int(Date) Then
                count = count + 1
                s = s & IIf(s = "", "", vbNewLine) & SafeStr(mRst(r, RS_ID)) & "  " & CalName(SafeStr(mRst(r, RS_EMP))) & "  " & _
                    IIf(Int(CDate(mRst(r, RS_FROM))) > Int(Date), "from " & Format$(CDate(mRst(r, RS_FROM)), "ddd mmm d") & " ", "") & _
                    IIf(IsDate(mRst(r, RS_TO)), "until " & Format$(CDate(mRst(r, RS_TO)), "ddd mmm d"), "open-ended") & "  -  " & SafeStr(mRst(r, RS_REASON))
            End If
        End If
    Next r
    ActiveRestrictionsText = s
End Function

' Writes one restriction (admin). dTo Empty = open-ended. Returns the id ("" = not written).
Public Function SaveRestriction(ByVal emp As String, ByVal dFrom As Date, ByVal dTo As Variant, ByVal reason As String) As String
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = RstTable(True)
    If lo Is Nothing Then Exit Function
    Dim id As String: id = NextId(lo, "RS")
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To RS_N)
    rec(1, RS_ID) = id: rec(1, RS_EMP) = emp: rec(1, RS_FROM) = Int(dFrom)
    If IsDate(dTo) Then rec(1, RS_TO) = Int(CDate(dTo)) Else rec(1, RS_TO) = Empty
    rec(1, RS_REASON) = "'" & Left$(reason, 200): rec(1, RS_STATUS) = "Active"
    rec(1, RS_USER) = CurrentUser(): rec(1, RS_STAMP) = Now: rec(1, RS_VOIDNOTE) = ""
    If Not AppendTo(lo, rec, RS_N) Then Exit Function
    RefreshCalCache
    MarkCalDirty
    SaveIfWritable
    SaveRestriction = id
End Function

' Alt+F8 (admin): a restriction for a name typed
Public Sub RestrictOT()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim q As Variant
    q = Application.InputBox(Prompt:="Who is restricted from overtime? Part of the name (the last name is enough):", Title:="OT restriction", Type:=2)
    If VarType(q) = vbBoolean Then Exit Sub
    Dim nm As String: nm = ResolvePersonName(Trim$(CStr(q)), "OT restriction")
    If nm = "" Then Exit Sub
    RestrictOTFor nm
    Exit Sub
Oops:
    ReportError "RestrictOT"
End Sub

' The Employee pop-up's Edit (admin): From / To / the reason asked, the row written, a receipt shown
Public Sub RestrictOTFor(ByVal emp As String)
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim why As String
    If OtRestrictedOn(emp, Date, why) Then
        If MsgBox(emp & " is already restricted: " & why & vbNewLine & vbNewLine & "Add another restriction anyway? (No = end the current one through Alt+F8 > EndRestriction.)", vbYesNo + vbQuestion + vbDefaultButton2, "OT restriction") = vbNo Then Exit Sub
    End If
    Dim a As Variant, dFrom As Date, dTo As Variant, t As String
    a = Application.InputBox(Prompt:="Restricted from (date, e.g. Oct 26):", Title:="OT restriction", Default:=DateText(Date), Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    If Not ParseTypedDate(Trim$(CStr(a)), dFrom) Then MsgBox "That isn't a date I can read.", vbExclamation, "OT restriction": Exit Sub
    a = Application.InputBox(Prompt:="Restricted until (date, included). Leave blank for open-ended (ended by hand later):", Title:="OT restriction", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    t = Trim$(CStr(a))
    dTo = Empty
    If t <> "" Then
        Dim d2 As Date
        If Not ParseTypedDate(t, d2) Then MsgBox "That isn't a date I can read.", vbExclamation, "OT restriction": Exit Sub
        If d2 < dFrom Then MsgBox "The end is before the start.", vbExclamation, "OT restriction": Exit Sub
        dTo = d2
    End If
    Dim reason As String: reason = AskRequired("The reason (kept on the Restrictions sheet, admin-only; never published):", "OT restriction")
    If reason = "" Then Exit Sub
    Dim id As String: id = SaveRestriction(emp, dFrom, dTo, reason)
    If id = "" Then MsgBox "The restriction was not written.", vbExclamation, "OT restriction": Exit Sub
    ShowReceiptNow "OT restriction " & id & " saved", _
                   "  " & id & "  " & emp & "  from " & Format$(dFrom, "ddd mmm d") & IIf(IsDate(dTo), " to " & Format$(CDate(dTo), "ddd mmm d"), " (open-ended)") & "  (a Restrictions row; no OT Log row)" & vbNewLine, _
                   "  Nobody's Total Hours changed. " & emp & " is off every call-out and posting list while restricted and is never charged." & vbNewLine, _
                   "  Needs Attention lists it; the call-out pop-up's Not listed names it; the members' file shows nothing of it." & vbNewLine, _
                   "  Alt+F8 > EndRestriction ends it early (the row is voided with a note)." & vbNewLine     ' (Oct 22: a stray ) here was a syntax error)
    Exit Sub
Oops:
    ReportError "RestrictOTFor"
End Sub

' Alt+F8 (admin): voids one active restriction by its id, with a note
Public Sub EndRestriction()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim n As Long, txt As String: txt = ActiveRestrictionsText(n)
    If n = 0 Then MsgBox "No active OT restrictions.", vbInformation, "OT restriction": Exit Sub
    Dim a As Variant
    a = Application.InputBox(Prompt:="Active restrictions:" & vbNewLine & txt & vbNewLine & vbNewLine & "Which one ends now? Type its id (RS-0001):", Title:="OT restriction", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim id As String: id = UCase$(Trim$(CStr(a)))
    Dim lo As ListObject: Set lo = RstTable(False)
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then MsgBox id & " is not a restriction id.", vbExclamation, "OT restriction": Exit Sub
    If UCase$(SafeStr(v(r, RS_STATUS))) = "VOID" Then MsgBox id & " is already ended.", vbInformation, "OT restriction": Exit Sub
    Dim note As String: note = AskRequired("Why does it end now? (kept on the row)", "OT restriction")
    If note = "" Then Exit Sub
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(r, RS_STATUS).Value = "VOID"
    lo.DataBodyRange.Cells(r, RS_VOIDNOTE).Value = "'" & note & " - " & CurrentUser() & " " & Format$(Now, "mmm d, yyyy h:mm AM/PM")
    ProtectSheet lo.Parent
    RefreshCalCache
    MarkCalDirty
    SaveIfWritable
    Tell id & " ended: " & SafeStr(v(r, RS_EMP)) & " is back on the call lists from the next load.", vbInformation, "OT restriction"
    Exit Sub
Oops:
    ReportError "EndRestriction"
End Sub

' v27 (item 21): the OVERTIME hours already held in the 7 days before an offer - accepted call-outs and awards from
' the log (OtWorkedOn), both shifts of each day, and the other shift of the offer's own day when it comes first.
' Scheduled shifts are NOT counted (review: a 4-crew rotation alone holds 48-60 h in many 7-day windows, so the
' cap would fire on most off-day acceptances; the brief's "Worked Periods + provisional" is premium hours). Whether
' Graeme's 48 means this, a calendar week, or all hours is an open question in briefs/v27.md. The cap's "hoursBefore".
Public Function WeekHoursBefore(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal shiftLbl As String) As Double
    Dim k As Long, t As Double
    For k = 1 To 7
        t = t + OtWorkedOn(empName, d - k, False) + OtWorkedOn(empName, d - k, True)
    Next k
    If IsNightShift(shiftLbl) Then t = t + OtWorkedOn(empName, d, False)   ' the day shift of the same day comes before a night
    WeekHoursBefore = t
End Function

' "" when the offer fits the week; otherwise the sentence the pop-ups show before asking for a reason
Public Function WeekCapProblem(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal shiftLbl As String, ByVal offer As Double) As String
    If WeekCap() <= 0 Then Exit Function             ' off (the default): no log reads at all
    Dim before As Double: before = WeekHoursBefore(empName, crewCode, d, shiftLbl)
    Dim over As Double: over = WeekCapOver(before, offer, WeekCap())
    If over <= 0 Then Exit Function
    WeekCapProblem = CalName(empName) & ": " & HoursText(before) & " h of overtime in the 7 days before this shift + " & HoursText(offer) & " h on offer = " & _
                     HoursText(before + offer) & " h, over the " & HoursText(WeekCap()) & " h week by " & HoursText(over) & " h."
End Function

' Is this person booked off (any absence type) on date d?
Public Function AbsentOn(ByVal empName As String, ByVal d As Date, Optional ByRef typ As String) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    If Not IsArray(mAbs) Then Exit Function
    Dim r As Long
    For r = 1 To UBound(mAbs, 1)
        If StrComp(Trim$(SafeStr(mAbs(r, AB_EMP))), Trim$(empName), vbTextCompare) = 0 Then
            If AbsActiveOn(r, d) Then typ = TypeTag(SafeStr(mAbs(r, AB_TYPE))): AbsentOn = True: Exit Function
        End If
    Next r
End Function


' v25 (D-L1): BookedOff for ONE shift - a sick call blocks the shifts inside its 24 h window,
' other absences the whole date (AbsentForShift). The switch paths use it wherever the shift
' is known, so a person inside a sick window may neither work nor give up that shift.
Public Function BookedOffShift(ByVal empName As String, ByVal d As Date, ByVal shiftLbl As String, _
                               Optional ByRef typ As String) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    If Not mTestMode Then mAbs = TableData(AbsTable())
    If shiftLbl = "" Then
        BookedOffShift = AbsentOn(empName, d, typ)
    Else
        BookedOffShift = AbsentForShift(empName, d, shiftLbl, typ)
    End If
End Function


' The same rule when the person's crew can differ from one day to the next
' (v24, crew history): crewPrev / crewOnD / crewNext = their crew on d - 1, d, d + 1.
Public Function RotationFreeReason3(ByVal crewPrev As String, ByVal crewOnD As String, ByVal crewNext As String, _
                                    ByVal d As Date, ByVal shiftLbl As String) As String
    Dim night As Boolean: night = IsNightShift(shiftLbl)
    Dim onD As String: onD = CrewShiftOn(crewOnD, d)
    If onD = "Days" Or onD = "Nights" Then RotationFreeReason3 = "scheduled": Exit Function
    If onD <> "?" Then
        If Not night And CrewShiftOn(crewPrev, d - 1) = "Nights" Then RotationFreeReason3 = "rest": Exit Function
        If night And CrewShiftOn(crewNext, d + 1) = "Days" Then RotationFreeReason3 = "rest": Exit Function
    End If
End Function

' Why someone can't take overtime on date d, shift shiftLbl ("" = free to sign up):
'   booked off (Vac / Leave / Trn ...), scheduled that day (rotation, A/B / C/D
'   days, DAYS Mon-Fri), the rest rule (coming off nights for a day shift,
'   starting days the next morning for a night shift), working a shift
'   switch that day, or gave up a shift that day in a switch.
' Booked off for THIS shift: a sick call blocks every shift that starts within
' 24 h after the sick shift ends (v22, SickBlocks in modLeave); other
' absences block the whole date.
Public Function AbsentForShift(ByVal empName As String, ByVal d As Date, ByVal shiftLbl As String, _
                               Optional ByRef typ As String) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    If Not IsArray(mAbs) Then Exit Function
    Dim r As Long, code As String
    For r = 1 To UBound(mAbs, 1)
        If StrComp(Trim$(SafeStr(mAbs(r, AB_EMP))), Trim$(empName), vbTextCompare) = 0 Then
            If UCase$(SafeStr(mAbs(r, AB_STATUS))) <> "VOID" And IsDate(mAbs(r, AB_START)) Then
                code = UCase$(SafeStr(mAbs(r, AB_TYPE)))
                If code = "SIC" And UBound(mAbs, 2) >= AB_SHIFT Then
                    If SickBlocks(CDate(mAbs(r, AB_START)), IsNightShift(SafeStr(mAbs(r, AB_SHIFT))), d, IsNightShift(shiftLbl)) Then
                        typ = TypeTag(code): AbsentForShift = True: Exit Function
                    End If
                ElseIf AbsActiveOn(r, d) Then
                    typ = TypeTag(code): AbsentForShift = True: Exit Function
                End If
            End If
        End If
    Next r
End Function

Public Function NotFreeReason(ByVal empName As String, ByVal crewCode As String, _
                              ByVal d As Date, ByVal shiftLbl As String, _
                              Optional ByVal hoursOffered As Double = 0, Optional ByRef maxHours As Double) As String
    ' v26.5.5 (item 9): hoursOffered > 0 = the whole offer must fit the 16 h rule; maxHours comes back for the listed
    If mShiftMap Is Nothing Then LoadCache
    maxHours = FATIGUE_MAX_HOURS
    Dim night As Boolean: night = IsNightShift(shiftLbl)
    Dim t As String
    If AbsentForShift(empName, d, shiftLbl, t) Then NotFreeReason = t: Exit Function
    ' (v27 item 21, review H1: an OT restriction is NOT tested here - NotFreeReason is also the switch rule, and a
    ' restricted person still works and switches their own shifts. FilterNotFree and the two posting sites ask
    ' OtRestrictedOn themselves.)
    ' v24: the crew on each of the three days comes from the crew history; crewCode
    ' (the roster's crew) is only the answer for someone with no history rows
    NotFreeReason = RotationFreeReason3(CrewOn(empName, d - 1, crewCode), CrewOn(empName, d, crewCode), _
                                        CrewOn(empName, d + 1, crewCode), d, shiftLbl)
    ' v25 (D-C3): a shift GIVEN UP in a switch was not worked, so it brings no rest rule with it
    If NotFreeReason = "rest" Then
        If Not night And SwitchOffOn(empName, d - 1, "Nights") <> "" Then NotFreeReason = ""
        If night And SwitchOffOn(empName, d + 1, "Days") <> "" Then NotFreeReason = ""
    End If
    If NotFreeReason <> "" Then Exit Function
    If SwitchWorkingOn(empName, d, "") <> "" Then NotFreeReason = "S/S": Exit Function
    If SwitchOffOn(empName, d) <> "" Then NotFreeReason = "S/S off": Exit Function
    If Not night And SwitchWorkingOn(empName, d - 1, "Nights") <> "" Then NotFreeReason = "rest": Exit Function
    If night And SwitchWorkingOn(empName, d + 1, "Days") <> "" Then NotFreeReason = "rest": Exit Function
    ' v26.5.5 (review): OT already accepted or awarded on THIS shift - a second call-out or posting for the same
    ' shift (two sick calls on one shift) must not list them again
    If OtWorkedOn(empName, d, night) > 0 Then NotFreeReason = "on OT": Exit Function
    ' v26.5.5 (item 9): the 16 h rule. The rules above keep a SCHEDULED shift on either side off the list;
    ' this one sees OT already accepted or awarded (and a switch) on the shift touching the one on offer.
    ' With hours given, not free unless the whole of them fits; without, not free only when nothing fits
    ' (the lists keep a person who can still take a partial, and the pop-ups show "max N h").
    maxHours = FatigueMaxFor(empName, crewCode, d, shiftLbl)
    If maxHours < IIf(hoursOffered > 0, hoursOffered, 0.5) - 0.001 Then NotFreeReason = "16 h rule"
End Function

' ====================================================================
'  v26.5.5 (item 9): THE 16 H RULE - accepted OT, awards and switches are WORKED shifts
'  The plant's rule (Graeme, Oct 19-21, 2026): at most FATIGUE_MAX_HOURS of consecutive work and at
'  least FATIGUE_MIN_REST before the next shift. Only the work BEFORE the offer counts toward the 16 h
'  (the shift touching it, and the one before that when the touching shift was worked whole); the work
'  AFTER it is governed by the rest gap: from the offer's start to the next work's start, less the rest.
'  A partial OT on a touching shift is counted whole (its clock time is not recorded) and joins the shift
'  before it whenever both were worked, and the next work is taken to start a whole shift later, so the rule
'  errs on the safe side; the worked block uses the shift's actual length (a DST night is 13 / 11 h), the
'  gap counts 12 h steps. OT already held on the shift on offer itself is "on OT" (NotFreeReason).
'  PURE (SelfTest "fatigue" / "fatblock"): FatigueMaxHours / FatigueProblem / FatigueBlockCore. Readers:
'  WorkedHoursOn (one shift), WorkedBlockBefore, HoursToNextWork, FatigueMaxFor ("max N h" on the pop-ups).
'  The OT-worked map (mOt) is read from the log on its first use and dropped by ResetOtWorked whenever
'  the log changes (AppendLogRows, VoidRef) or the caches reset.
' ====================================================================
Public Function FatigueMaxHours(ByVal workedBefore As Double, ByVal hoursToNext As Double) As Double
    Dim m As Double: m = FATIGUE_MAX_HOURS - workedBefore
    If hoursToNext - FATIGUE_MIN_REST < m Then m = hoursToNext - FATIGUE_MIN_REST
    If m < 0 Then m = 0
    FatigueMaxHours = m
End Function

Public Function FatigueProblem(ByVal workedBefore As Double, ByVal hoursOffered As Double, ByVal hoursToNext As Double) As String
    If hoursOffered > FatigueMaxHours(workedBefore, hoursToNext) + 0.001 Then FatigueProblem = "16 h rule"
End Function

' PURE (SelfTest "fatblock"): from the hours worked on the four shifts around the offer - the one before the
' touching shift, the touching shift, the next shift, the one after - the consecutive block that ends when the
' offer starts (the touching shift plus the one before it whenever both were worked) and the hours from the
' offer's start to the next work (12 h steps; 99 = none)
Public Sub FatigueBlockCore(ByVal hPrevPrev As Double, ByVal hPrev As Double, ByVal hNext1 As Double, ByVal hNext2 As Double, _
                            ByRef workedBefore As Double, ByRef hoursToNext As Double)
    workedBefore = 0
    If hPrev > 0 Then workedBefore = hPrev + hPrevPrev
    hoursToNext = 99
    If hNext1 > 0 Then
        hoursToNext = USUAL_SHIFT_HOURS
    ElseIf hNext2 > 0 Then
        hoursToNext = 2 * USUAL_SHIFT_HOURS
    End If
End Sub

Public Sub ResetOtWorked()
    Set mOt = Nothing
End Sub

' The OT each person has ACCEPTED (a call-out) or been AWARDED (a posting), from the non-void log rows:
' the row's Hours (an Accepted row carries the hours taken, an Awarded row the hours awarded)
Private Sub LoadOtWorked()
    Set mOt = CreateObject("Scripting.Dictionary")
    mOt.CompareMode = 1
    If mTestMode Then Exit Sub                              ' SelfTest's fixtures hold no OT
    Dim v As Variant: v = LogRowsNow()
    If Not IsArray(v) Then Exit Sub
    Dim i As Long, k As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" And IsDate(v(i, L_DATE)) Then
            Select Case SafeStr(v(i, L_OUTCOME))
                Case "Accepted", "Awarded"
                    k = Trim$(SafeStr(v(i, L_EMP))) & "|" & Format$(CDate(v(i, L_DATE)), "yyyymmdd") & "|" & _
                        IIf(IsNightShift(SafeStr(v(i, L_SHIFTLBL))), "N", "D")
                    mOt(k) = NumOr0(mOt(k)) + NumOr0(v(i, L_BASEHRS))
            End Select
        End If
    Next i
End Sub

' OT this person accepted / was awarded on one shift (0 = none)
Public Function OtWorkedOn(ByVal empName As String, ByVal d As Date, ByVal isNight As Boolean) As Double
    If mOt Is Nothing Then LoadOtWorked
    Dim k As String: k = Trim$(empName) & "|" & Format$(d, "yyyymmdd") & "|" & IIf(isNight, "N", "D")
    If mOt.Exists(k) Then OtWorkedOn = mOt(k)
End Function

' Hours this person WORKS on one shift: OT accepted / awarded, a switch worked, or scheduled by the
' rotation (unless booked off for it or given up in a switch). 0 = not working it.
Public Function WorkedHoursOn(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal isNight As Boolean) As Double
    Dim lbl As String: lbl = IIf(isNight, "Nights", "Days")
    Dim t As String
    If AbsentForShift(empName, d, lbl, t) Then Exit Function
    Dim ot As Double: ot = OtWorkedOn(empName, d, isNight)
    If ot > 0 Then WorkedHoursOn = ot: Exit Function
    Dim full As Double: full = DefaultShiftHours(d, isNight, USUAL_SHIFT_HOURS)
    If SwitchWorkingOn(empName, d, lbl) <> "" Then WorkedHoursOn = full: Exit Function
    If SwitchOffOn(empName, d, lbl) <> "" Then Exit Function
    If PendingGiveUp(empName, d, lbl) Then Exit Function          ' v27.5 (item 1): given up in the switch being saved
    If CrewShiftOn(CrewOn(empName, d, crewCode), d) = lbl Then WorkedHoursOn = full
End Function

' v27.5 (item 1): a switch being saved is not in the table yet, so the shift its giver GIVES UP still reads as worked by
' the rotation (and a same-day sister swap gives up two). SwitchSaveStops names them here while it reads the 16 h facts;
' WorkedHoursOn then gives each 0 h. Cleared on every exit of SwitchSaveStops - never left set.
Public Sub SetPendingGiveUp(ByVal slot As Long, ByVal empName As String, ByVal d As Date, ByVal shiftLbl As String)
    If slot < 1 Or slot > 2 Then Exit Sub
    If shiftLbl = "" Then mPendOff(slot) = "" Else mPendOff(slot) = Trim$(empName) & "|" & Format$(d, "yyyymmdd") & "|" & shiftLbl
End Sub

Public Sub ClearPendingGiveUp()
    mPendOff(1) = "": mPendOff(2) = ""
End Sub

Private Function PendingGiveUp(ByVal empName As String, ByVal d As Date, ByVal lbl As String) As Boolean
    Dim k As String: k = Trim$(empName) & "|" & Format$(d, "yyyymmdd") & "|" & lbl
    PendingGiveUp = (StrComp(mPendOff(1), k, vbTextCompare) = 0 Or StrComp(mPendOff(2), k, vbTextCompare) = 0)
End Function

' v27.5 (item 1): the 16 h rule for a WHOLE shift someone would work in a switch - the arithmetic alone, so a scheduled
' shift beside it counts the same as OT accepted or a switch (NotFreeReason stops at "scheduled" / "rest" first, which
' is right for an OT list and wrong for a swap, where the person IS scheduled that day). "on OT" = OT already held on
' the shift; "16 h rule" = the whole shift does not fit (FatigueMaxFor); "" = fits, or no shift named.
Public Function FatigueShiftReason(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal shiftLbl As String) As String
    If shiftLbl = "" Then Exit Function
    Dim night As Boolean: night = IsNightShift(shiftLbl)
    If OtWorkedOn(empName, d, night) > 0 Then FatigueShiftReason = "on OT": Exit Function
    If FatigueMaxFor(empName, crewCode, d, shiftLbl) < DefaultShiftHours(d, night, USUAL_SHIFT_HOURS) - 0.001 Then FatigueShiftReason = "16 h rule"
End Function

' The four shifts around the offer and the hours worked on each, then the pure core
Private Sub FatigueInputs(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal isNight As Boolean, _
                          ByRef workedBefore As Double, ByRef hoursToNext As Double)
    Dim hPP As Double, hP As Double, hN1 As Double, hN2 As Double
    If isNight Then
        hP = WorkedHoursOn(empName, crewCode, d, False)          ' a night offer: that date's day shift touches it ...
        If hP > 0 Then hPP = WorkedHoursOn(empName, crewCode, d - 1, True)      ' ... then the night before that
        hN1 = WorkedHoursOn(empName, crewCode, d + 1, False)
        If hN1 = 0 Then hN2 = WorkedHoursOn(empName, crewCode, d + 1, True)
    Else
        hP = WorkedHoursOn(empName, crewCode, d - 1, True)       ' a day offer: the night before touches it ...
        If hP > 0 Then hPP = WorkedHoursOn(empName, crewCode, d - 1, False)     ' ... then that night's day shift
        hN1 = WorkedHoursOn(empName, crewCode, d, True)
        If hN1 = 0 Then hN2 = WorkedHoursOn(empName, crewCode, d + 1, False)
    End If
    FatigueBlockCore hPP, hP, hN1, hN2, workedBefore, hoursToNext
End Sub

' The consecutive hours of work that END when the OT on offer starts
Public Function WorkedBlockBefore(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal isNight As Boolean) As Double
    Dim wb As Double, htn As Double
    FatigueInputs empName, crewCode, d, isNight, wb, htn
    WorkedBlockBefore = wb
End Function

' Hours from the start of the OT on offer to the start of this person's next work (99 = none in the next two shifts)
Public Function HoursToNextWork(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal isNight As Boolean) As Double
    Dim wb As Double, htn As Double
    FatigueInputs empName, crewCode, d, isNight, wb, htn
    HoursToNextWork = htn
End Function

' The most OT this person may take on d / shiftLbl under the 16 h rule (FATIGUE_MAX_HOURS when nothing touches it)
Public Function FatigueMaxFor(ByVal empName As String, ByVal crewCode As String, ByVal d As Date, ByVal shiftLbl As String) As Double
    If mShiftMap Is Nothing Then LoadCache
    Dim wb As Double, htn As Double
    FatigueInputs empName, crewCode, d, IsNightShift(shiftLbl), wb, htn
    FatigueMaxFor = FatigueMaxHours(wb, htn)
End Function

' Names for the Shift Switch pop-up drop-downs, alphabetical. Everyone listed
' is signed off in the unit (a code in its column on the OT Roster).
'   role "OFF"   - scheduled that date (on shiftLbl if given), not booked off,
'                  not already on a switch that day
'   role "WORKS" - free that shift (see NotFreeReason; having given up their
'                  own shift that day is fine). Same-day sister swap: scheduled
'                  on the OTHER shift that day instead.
' With no date yet: everyone signed off in the unit.
Public Function SwitchPickList(ByVal unit As String, ByVal d As Variant, ByVal role As String, _
                               ByVal shiftLbl As String, ByVal sisterSwap As Boolean) As Variant
    Dim recs() As EmpRec, n As Long, i As Long, j As Long, k As Long
    If DeptCol(unit) = 0 Then Exit Function
    Dim hasDate As Boolean, dd As Date
    If IsDate(d) Then hasDate = True: dd = CDate(d)
    If hasDate Then
        n = LoadEligible(unit, "ALL", recs, 999, True, dd)      ' v24: crews as of that date
    Else
        n = LoadEligible(unit, "ALL", recs, 999, True)
    End If
    If n = 0 Then Exit Function
    LoadCache
    Dim names() As String: ReDim names(1 To n)
    Dim keep As Boolean, onD As String, t As String, why As String
    For i = 1 To n
        keep = True
        If hasDate Then
            onD = CrewShiftOn(recs(i).EmpShift, dd)
            If role = "OFF" Then
                keep = (onD = "Days" Or onD = "Nights" Or onD = "?")
                If keep And shiftLbl <> "" And onD <> "?" Then keep = (onD = shiftLbl)
                ' v25 (D-L1): booked off for the shift they would give up (a sick window is per shift)
                If keep Then keep = Not AbsentForShift(recs(i).EmpName, dd, IIf(onD = "?", shiftLbl, onD), t)
                If keep Then keep = (SwitchOffOn(recs(i).EmpName, dd) = "")
                If keep Then keep = (SwitchWorkingOn(recs(i).EmpName, dd, "") = "")
            ElseIf sisterSwap Then
                keep = ((onD = "Days" Or onD = "Nights") And onD <> shiftLbl)
                If keep Then keep = Not AbsentForShift(recs(i).EmpName, dd, onD, t)
                If keep Then keep = Not AbsentForShift(recs(i).EmpName, dd, shiftLbl, t)      ' v25 review (D-L1): nor for the shift they would WORK
            Else
                If shiftLbl = "" Then
                    why = NotFreeReason(recs(i).EmpName, recs(i).EmpShift, dd, "Days")
                    If why <> "" And why <> "S/S off" Then why = NotFreeReason(recs(i).EmpName, recs(i).EmpShift, dd, "Nights")
                Else
                    why = NotFreeReason(recs(i).EmpName, recs(i).EmpShift, dd, shiftLbl)
                End If
                keep = (why = "" Or why = "S/S off")
            End If
        End If
        If keep Then
            k = k + 1
            names(k) = recs(i).EmpName
        End If
    Next i
    If k = 0 Then Exit Function
    ' alphabetical
    Dim tmp As String
    For i = 2 To k
        tmp = names(i)
        j = i - 1
        Do While j >= 1
            If StrComp(names(j), tmp, vbTextCompare) <= 0 Then Exit Do
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = tmp
    Next i
    Dim outA() As String: ReDim outA(0 To k - 1)
    For i = 1 To k
        outA(i - 1) = names(i)
    Next i
    SwitchPickList = outA
End Function

' Keeps only people free to take overtime on d / shiftLbl (see NotFreeReason).
' removed gets a short summary, e.g. "14 scheduled that day, 4 rest rule;
' booked off: Fred C (Vac); shift switch: Paul M". Returns the new count.
' v27 (item 18, B6): detail = every person left off, by reason, one group per line ("Scheduled that day (3): Ann B,
' Bob C, Cal D" / "Rest rule ..." / "Booked off: Dee E (Vac), ..." / "Shift switch: ..." / "Already on OT that
' shift: ..." / "16 h rule: ...") - the "Not listed (N)..." box on the call-out pop-up. Reads only.
Public Function FilterNotFree(ByRef recs() As EmpRec, ByVal n As Long, ByVal d As Date, ByVal shiftLbl As String, _
                              ByRef removed As String, Optional ByVal useCache As Boolean = False, _
                              Optional ByRef detail As String = "", Optional ByVal reasons As Object = Nothing) As Long     ' v27 (M4 item 37): reasons(name) = why
    If Not useCache Then LoadCache
    removed = "": detail = ""
    Dim i As Long, k As Long, why As String, mx As Double
    Dim nSched As Long, nRest As Long, offTxt As String, ssTxt As String, fatTxt As String, otTxt As String
    Dim schedTxt As String, restTxt As String, rstTxt As String
    ' v26.5.5 (review): the 16 h readers look at four shifts per person - one read of the switch table for the list;
    ' the batch is ended on every path (a batch left on would feed stale switch rows to the rest of the session)
    Dim batched As Boolean
    If Not SwitchBatchActive() Then BeginSwitchBatch: batched = True
    On Error GoTo Failed
    For i = 1 To n
        ' v27 (item 21, review H1): the restriction is a LIST rule - tested here, never inside NotFreeReason
        If OtRestrictedOn(recs(i).EmpName, d) Then
            why = "restricted": mx = 0
        Else
            why = NotFreeReason(recs(i).EmpName, recs(i).EmpShift, d, shiftLbl, 0, mx)
        End If
        If why <> "" And Not reasons Is Nothing Then reasons(recs(i).EmpName) = why     ' (before the list closes up)
        Select Case why
            Case ""
                k = k + 1
                If k <> i Then recs(k) = recs(i)
                recs(k).MaxHours = mx                      ' v26.5.5 (item 9): the pop-ups mark anyone whose max is under the offer
            Case "scheduled": nSched = nSched + 1: schedTxt = schedTxt & IIf(schedTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case "rest": nRest = nRest + 1: restTxt = restTxt & IIf(restTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case "S/S", "S/S off": ssTxt = ssTxt & IIf(ssTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case "16 h rule": fatTxt = fatTxt & IIf(fatTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case "on OT": otTxt = otTxt & IIf(otTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case "restricted": rstTxt = rstTxt & IIf(rstTxt = "", "", ", ") & CalName(recs(i).EmpName)   ' v27 (item 21)
            Case Else: offTxt = offTxt & IIf(offTxt = "", "", ", ") & CalName(recs(i).EmpName) & " (" & why & ")"
        End Select
    Next i
    If batched Then EndSwitchBatch
    If nSched > 0 Then removed = nSched & " scheduled that day"
    If nRest > 0 Then removed = removed & IIf(removed = "", "", ", ") & nRest & " rest rule (off nights / days next morning)"
    If offTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "booked off: " & offTxt
    If ssTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "shift switch: " & ssTxt
    If otTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "already on OT that shift: " & otTxt
    If fatTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "16 h rule (OT accepted / awarded or a switch touching the shift): " & fatTxt
    If rstTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "OT restricted: " & rstTxt
    ' v27 (item 18): the same people, every name, one line per reason (each group capped so the box stays under
    ' MsgBox's 1,024 characters - review LOW-13)
    schedTxt = CapNames(schedTxt, 25): restTxt = CapNames(restTxt, 25): offTxt = CapNames(offTxt, 25)
    ssTxt = CapNames(ssTxt, 25): otTxt = CapNames(otTxt, 25): fatTxt = CapNames(fatTxt, 25): rstTxt = CapNames(rstTxt, 25)
    If nSched > 0 Then detail = "Scheduled that day (" & nSched & "): " & schedTxt
    If nRest > 0 Then detail = detail & IIf(detail = "", "", vbNewLine) & "Rest rule - off nights that morning / on days next morning (" & nRest & "): " & restTxt
    If offTxt <> "" Then detail = detail & IIf(detail = "", "", vbNewLine) & "Booked off, sick, or inside the 24 h after a sick shift: " & offTxt
    If ssTxt <> "" Then detail = detail & IIf(detail = "", "", vbNewLine) & "Shift switch that day (working one, or gave the shift up): " & ssTxt
    If otTxt <> "" Then detail = detail & IIf(detail = "", "", vbNewLine) & "Already on OT that shift: " & otTxt
    If fatTxt <> "" Then detail = detail & IIf(detail = "", "", vbNewLine) & "16 h rule (OT accepted / awarded or a switch touching the shift; nothing fits): " & fatTxt
    If rstTxt <> "" Then detail = detail & IIf(detail = "", "", vbNewLine) & "OT restricted by an admin (not listed, never charged): " & rstTxt
    FilterNotFree = k
    Exit Function
Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    If batched Then EndSwitchBatch
    On Error GoTo 0
    Err.Raise en, "FilterNotFree", em
End Function

' A ", "-joined list cut to its first n names, "+N more" after
Private Function CapNames(ByVal txt As String, ByVal n As Long) As String
    If txt = "" Then Exit Function
    Dim parts As Variant: parts = Split(txt, ", ")
    If UBound(parts) + 1 <= n Then CapNames = txt: Exit Function
    Dim i As Long, s As String
    For i = 0 To n - 1
        s = s & IIf(s = "", "", ", ") & parts(i)
    Next i
    CapNames = s & " +" & (UBound(parts) + 1 - n) & " more"
End Function


' ====================================================================
'  SLOTS (what the calendar shows for one date + shift)
' ====================================================================
Public Function SlotsFor(ByVal d As Date, ByVal shiftLbl As String, ByVal groupName As String, _
                         ByRef slots() As CalSlot, Optional ByVal useCache As Boolean = False) As Long
    If Not useCache Then LoadCache
    Dim tS As Double: tS = Timer                     ' v25.1.2: measured
    Dim n As Long, r As Long, c As Long, used As String, rt As String, lt As String
    ReDim slots(1 To 20)

    If IsArray(mAbs) Then
        For r = 1 To UBound(mAbs, 1)
            If AbsActiveOn(r, d) Then
                If UnitInGroup(SafeStr(mAbs(r, AB_UNIT)), groupName) Then
                    ' a sick call sits on the shift it was booked for; other absences on the crew's scheduled shift
                    If AbsenceOnShift(r, d, shiftLbl) Then
                        lt = CalName(SafeStr(mAbs(r, AB_EMP))) & " (" & TypeTag(SafeStr(mAbs(r, AB_TYPE)))
                        If SafeStr(mAbs(r, AB_TIMES)) <> "" Then lt = lt & " " & SafeStr(mAbs(r, AB_TIMES))
                        lt = lt & HrsTag(mAbs(r, AB_HOURS)) & ")"
                        rt = ""
                        If IsArray(mCov) Then
                            For c = 1 To UBound(mCov, 1)
                                If CovLive(c, d, shiftLbl) Then
                                    If StrComp(SafeStr(mCov(c, CV_ABS)), SafeStr(mAbs(r, AB_ID)), vbTextCompare) = 0 Then
                                        If SafeStr(mCov(c, CV_BY)) <> "" Then
                                            rt = rt & IIf(rt = "", "", "; ") & SafeStr(mCov(c, CV_BY)) & HrsTag(mCov(c, CV_HOURS))
                                        End If
                                        used = used & "|" & SafeStr(mCov(c, CV_ID))
                                    End If
                                End If
                            Next c
                        End If
                        n = n + 1
                        If n > UBound(slots) Then ReDim Preserve slots(1 To n + 20)
                        slots(n).Key = "A:" & SafeStr(mAbs(r, AB_ID))
                        slots(n).LeftText = lt
                        slots(n).RightText = rt
                        slots(n).TypeCode = TypeCode(SafeStr(mAbs(r, AB_TYPE)))
                        slots(n).IsOpen = (rt = "")
                    End If
                End If
            End If
        Next r
    End If

    If IsArray(mCov) Then
        For c = 1 To UBound(mCov, 1)
            If CovLive(c, d, shiftLbl) Then
                If UnitInGroup(SafeStr(mCov(c, CV_UNIT)), groupName) And _
                   InStr(used & "|", "|" & SafeStr(mCov(c, CV_ID)) & "|") = 0 Then
                    n = n + 1
                    If n > UBound(slots) Then ReDim Preserve slots(1 To n + 20)
                    slots(n).Key = "C:" & SafeStr(mCov(c, CV_ID))
                    slots(n).LeftText = SafeStr(mCov(c, CV_NEED))
                    slots(n).RightText = SafeStr(mCov(c, CV_BY))
                    If slots(n).RightText <> "" Then slots(n).RightText = slots(n).RightText & HrsTag(mCov(c, CV_HOURS))
                    slots(n).TypeCode = UCase$(SafeStr(mCov(c, CV_TYPE)))
                    slots(n).IsOpen = (SafeStr(mCov(c, CV_BY)) = "")
                End If
            End If
        Next c
    End If
    SlotsFor = n
    mTimeSlots = mTimeSlots + Secs(tS, Timer)
End Function

' Which shift an absence row shows on: sick calls on their booked shift (v22),
' everything else on the crew's scheduled shift that day
Public Function AbsenceOnShift(ByVal r As Long, ByVal d As Date, ByVal shiftLbl As String) As Boolean
    If UCase$(SafeStr(mAbs(r, AB_TYPE))) = "SIC" And UBound(mAbs, 2) >= AB_SHIFT Then
        If SafeStr(mAbs(r, AB_SHIFT)) <> "" Then
            AbsenceOnShift = (StrComp(SafeStr(mAbs(r, AB_SHIFT)), shiftLbl, vbTextCompare) = 0)
            Exit Function
        End If
    End If
    AbsenceOnShift = (EmpShiftOn(SafeStr(mAbs(r, AB_EMP)), d) = shiftLbl)
End Function

Private Function CovLive(ByVal c As Long, ByVal d As Date, ByVal shiftLbl As String) As Boolean
    If UCase$(SafeStr(mCov(c, CV_STATUS))) = "VOID" Then Exit Function
    If Not IsDate(mCov(c, CV_DATE)) Then Exit Function
    If Int(CDate(mCov(c, CV_DATE))) <> Int(d) Then Exit Function
    CovLive = (StrComp(SafeStr(mCov(c, CV_SHIFT)), shiftLbl, vbTextCompare) = 0)
End Function

' ====================================================================
'  WRITE: absences / needs / coverage
' ====================================================================
Public Function SaveAbsence(ByVal emp As String, ByVal unit As String, ByVal typ As String, _
                            ByVal dStart As Date, ByVal dEnd As Variant, ByVal hrs As Variant, _
                            ByVal times As String, ByVal note As String, _
                            Optional ByVal homeUnit As String = "", Optional ByVal leaveGroup As String = "", _
                            Optional ByVal shiftLbl As String = "", Optional ByVal idPrefix As String = "AB", _
                            Optional ByVal quiet As Boolean = False, Optional ByVal skipOnClash As Boolean = False) As String
    ' quiet: no overlap prompt (saves anyway) and no save-to-disk; skipOnClash: an
    ' overlapping booking wins and nothing is saved (Convert imports only)
    If Not RequireSupervisor() Then Exit Function     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = AbsTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < AB_N Then
        MsgBox "The Absences table needs its v22 upgrade first: an admin runs Alt+F8 > UpgradeAbsenceColumns once.", vbExclamation, "Time off"
        Exit Function
    End If

    ' overlap warning
    LoadCache
    Dim r As Long, clash As String, e2 As Date
    If IsArray(mAbs) Then
        For r = 1 To UBound(mAbs, 1)
            If StrComp(Trim$(SafeStr(mAbs(r, AB_EMP))), emp, vbTextCompare) = 0 And _
               UCase$(SafeStr(mAbs(r, AB_STATUS))) <> "VOID" And IsDate(mAbs(r, AB_START)) Then
                If IsDate(mAbs(r, AB_END)) Then e2 = CDate(mAbs(r, AB_END)) Else e2 = DateSerial(2999, 12, 31)
                If CDate(mAbs(r, AB_START)) <= IIf(IsDate(dEnd), dEnd, DateSerial(2999, 12, 31)) And e2 >= dStart Then
                    clash = clash & "  " & SafeStr(mAbs(r, AB_ID)) & " " & SafeStr(mAbs(r, AB_TYPE)) & " " & _
                            Format$(mAbs(r, AB_START), "mmm d") & " - " & _
                            IIf(IsDate(mAbs(r, AB_END)), Format$(mAbs(r, AB_END), "mmm d"), "open") & vbNewLine
                End If
            End If
        Next r
    End If
    If clash <> "" Then
        If skipOnClash Then Exit Function
        If Not quiet Then
            If MsgBox(emp & " already has time off in that range:" & vbNewLine & clash & vbNewLine & _
                      "Save this one too?", vbYesNo + vbExclamation, "Overlap") = vbNo Then Exit Function
        End If
    End If

    Dim id As String: id = NextId(lo, idPrefix)
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To AB_N)
    rec(1, AB_ID) = id
    rec(1, AB_EMP) = emp
    rec(1, AB_UNIT) = unit
    rec(1, AB_TYPE) = TypeCode(typ)
    rec(1, AB_START) = Int(dStart)
    If IsDate(dEnd) Then rec(1, AB_END) = Int(CDate(dEnd))
    If IsNumeric(hrs) And Not IsEmpty(hrs) Then
        If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then rec(1, AB_HOURS) = CDbl(hrs)
    End If
    rec(1, AB_TIMES) = times
    rec(1, AB_NOTE) = note
    rec(1, AB_STATUS) = "Active"
    rec(1, AB_USER) = CurrentUser()
    rec(1, AB_STAMP) = Now
    rec(1, AB_HOME) = homeUnit
    rec(1, AB_GROUP) = leaveGroup
    rec(1, AB_SHIFT) = shiftLbl
    If Not AppendTo(lo, rec, AB_N) Then Exit Function
    UnlockSheet lo.Parent
    lo.ListColumns(AB_START).DataBodyRange.NumberFormat = "ddd mmm d, yyyy"
    lo.ListColumns(AB_END).DataBodyRange.NumberFormat = "ddd mmm d, yyyy"
    ProtectSheet lo.Parent
    If Not quiet Then SaveIfWritable
    SaveAbsence = id
End Function

' Sets an absence's end date without asking (v22: leave return, sick call)
' the previous End for a Note: a block If, never CDate on a text cell (review)
Private Function OldEndText(ByVal x As Variant) As String
    If IsDate(x) Then OldEndText = Format$(CDate(x), "mmm d, yyyy") Else OldEndText = "open-ended"
End Function

Public Function SetAbsenceEndTo(ByVal id As String, ByVal dEnd As Date) As Boolean
    Dim lo As ListObject: Set lo = AbsTable()
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Function
    UnlockSheet lo.Parent
    ' v27 (item 13): the previous end stays in the Note here too
    lo.DataBodyRange.Cells(r, AB_NOTE).Value = Trim$(SafeStr(v(r, AB_NOTE)) & IIf(Trim$(SafeStr(v(r, AB_NOTE))) = "", "", " | ") & _
                                               "End was " & OldEndText(v(r, AB_END)) & " (changed " & Format$(Now, "mmm d") & ")")
    lo.DataBodyRange.Cells(r, AB_END).Value = Int(dEnd)
    lo.DataBodyRange.Cells(r, AB_END).NumberFormat = "ddd mmm d, yyyy"
    ProtectSheet lo.Parent
    SetAbsenceEndTo = True
End Function

' One absence row's fields (v22). False if not found.
Public Function AbsenceInfo(ByVal id As String, ByRef emp As String, ByRef unit As String, ByRef code As String, _
                            ByRef dStart As Date, ByRef dEnd As Variant, ByRef homeUnit As String, _
                            ByRef leaveGroup As String, ByRef status As String) As Boolean
    Dim v As Variant: v = TableData(AbsTable())
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Function
    emp = SafeStr(v(r, AB_EMP)): unit = SafeStr(v(r, AB_UNIT)): code = UCase$(SafeStr(v(r, AB_TYPE)))
    If IsDate(v(r, AB_START)) Then dStart = CDate(v(r, AB_START))
    If IsDate(v(r, AB_END)) Then dEnd = CDate(v(r, AB_END)) Else dEnd = Empty
    If UBound(v, 2) >= AB_GROUP Then homeUnit = SafeStr(v(r, AB_HOME)): leaveGroup = SafeStr(v(r, AB_GROUP))
    status = UCase$(SafeStr(v(r, AB_STATUS)))
    AbsenceInfo = True
End Function

' Open leave bookings (leave type, non-void, no end or end >= today) for the
' Returning pop-up: display strings; ids() gets the Absence IDs.
Public Function OpenLeaves(ByRef ids As Variant) As Variant
    ids = Empty
    Dim v As Variant: v = TableData(AbsTable())
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long, d() As String, idArr() As String, eTxt As String, grpTxt As String, homeTxt As String
    Dim hasV22 As Boolean: hasV22 = (UBound(v, 2) >= AB_SHIFT)
    ReDim d(0 To UBound(v, 1) - 1): ReDim idArr(0 To UBound(v, 1) - 1)
    For i = 1 To UBound(v, 1)
        If IsLeaveCode(SafeStr(v(i, AB_TYPE))) And UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" And IsDate(v(i, AB_START)) Then
            grpTxt = "": homeTxt = ""
            If hasV22 Then
                grpTxt = SafeStr(v(i, AB_GROUP))
                homeTxt = "   (" & SafeStr(v(i, AB_HOME)) & ")"
            End If
            ' open, or closed in the last 60 days and not yet topped up (a voided LV row lets it be redone)
            If (Not IsDate(v(i, AB_END)) Or (IsDate(v(i, AB_END)) And CDate(v(i, AB_END)) >= Date - 60)) And Not LeaveToppedUp(grpTxt) Then
                If IsDate(v(i, AB_END)) Then eTxt = "to " & Format$(v(i, AB_END), "mmm d") Else eTxt = "open-ended"
                idArr(n) = SafeStr(v(i, AB_ID))
                d(n) = idArr(n) & "   " & SafeStr(v(i, AB_EMP)) & "   Leave from " & Format$(v(i, AB_START), "mmm d, yyyy") & "   " & eTxt & homeTxt
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve d(0 To n - 1): ReDim Preserve idArr(0 To n - 1)
    ids = idArr
    OpenLeaves = d
End Function

' True when the leave group already has a non-void top-up row in the OT Log
Public Function LeaveToppedUp(ByVal leaveGroup As String) As Boolean
    If Trim$(leaveGroup) = "" Then Exit Function
    LeaveToppedUp = HasActiveLogRows(leaveGroup)
    If Not LeaveToppedUp Then LeaveToppedUp = ArchiveHasActiveRef(leaveGroup)     ' v23: topped up before the rollover
End Function

' Leave-type absence rows for the top-up peer test: True if emp has one
' overlapping dFrom..dTo (non-void)
Public Function HasLeaveOverlapping(ByVal emp As String, ByVal dFrom As Date, ByVal dTo As Date) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    If Not IsArray(mAbs) Then Exit Function
    Dim r As Long, e2 As Date
    For r = 1 To UBound(mAbs, 1)
        If StrComp(Trim$(SafeStr(mAbs(r, AB_EMP))), Trim$(emp), vbTextCompare) = 0 And _
           IsLeaveCode(SafeStr(mAbs(r, AB_TYPE))) And UCase$(SafeStr(mAbs(r, AB_STATUS))) <> "VOID" And IsDate(mAbs(r, AB_START)) Then
            If IsDate(mAbs(r, AB_END)) Then e2 = CDate(mAbs(r, AB_END)) Else e2 = DateSerial(2999, 12, 31)
            If CDate(mAbs(r, AB_START)) <= dTo And e2 >= dFrom Then HasLeaveOverlapping = True: Exit Function
        End If
    Next r
End Function

' Shifts a person would have worked in a range (for the pop-up)
' v25 (item 1c, from frmTimeOff): the scheduled shifts of a booking in one call - how many
' (days / nights) and the first and last day with one. 0 = none in the range (dFirst / dLast
' unset). The 400-day cap is ShiftCount's.
Public Function ScheduledSpan(ByVal emp As String, ByVal dStart As Date, ByVal dEnd As Date, _
                              ByRef dFirst As Date, ByRef dLast As Date, ByRef nDays As Long, ByRef nNights As Long) As Long
    LoadCache
    Dim d As Date, s As String, found As Boolean
    nDays = 0: nNights = 0
    If dEnd < dStart Then Exit Function
    If dEnd - dStart > 400 Then dEnd = dStart + 400
    For d = Int(dStart) To Int(dEnd)
        s = EmpShiftOn(emp, d)
        If s = "Days" Then nDays = nDays + 1
        If s = "Nights" Then nNights = nNights + 1
        If s <> "" Then
            If Not found Then dFirst = d
            dLast = d
            found = True
        End If
    Next d
    ScheduledSpan = nDays + nNights
End Function

' v25: which end(s) of a vacation / personal booking are days off and can be trimmed to the
' scheduled shifts inside it. "" = nothing to trim; "start", "end", "both". Pure (SelfTest trimkind).
Public Function TrimKind(ByVal ds As Date, ByVal de As Date, ByVal dFirst As Date, ByVal dLast As Date) As String
    Dim a As Boolean, b As Boolean
    a = (Int(dFirst) > Int(ds))
    b = (Int(dLast) < Int(de))
    If a And b Then
        TrimKind = "both"
    ElseIf a Then
        TrimKind = "start"
    ElseIf b Then
        TrimKind = "end"
    End If
End Function

Public Function ShiftCount(ByVal emp As String, ByVal dStart As Date, ByVal dEnd As Date, _
                           ByRef nDays As Long, ByRef nNights As Long) As Long
    LoadCache
    Dim d As Date, s As String
    nDays = 0: nNights = 0
    If dEnd < dStart Then Exit Function
    If dEnd - dStart > 400 Then dEnd = dStart + 400
    For d = dStart To dEnd
        s = EmpShiftOn(emp, d)
        If s = "Days" Then nDays = nDays + 1
        If s = "Nights" Then nNights = nNights + 1
    Next d
    ShiftCount = nDays + nNights
End Function

' A need with nobody off (PHA, support op ...) for each date in a range
Public Function SaveNeed(ByVal unit As String, ByVal needText As String, ByVal typ As String, _
                         ByVal dStart As Date, ByVal dEnd As Date, ByVal shiftSel As String, _
                         ByVal hrs As Variant) As Long
    If Not RequireAdmin() Then Exit Function          ' (v25.3 review: only Book Time Off reaches this - admin)
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = CovTable(True)
    If lo Is Nothing Then Exit Function
    If dEnd < dStart Then dEnd = dStart
    Dim shifts As Variant
    If shiftSel = "Both" Then shifts = Array("Days", "Nights") Else shifts = Array(shiftSel)
    Dim nRows As Long: nRows = (Int(dEnd) - Int(dStart) + 1) * (UBound(shifts) + 1)
    If nRows > 400 Then MsgBox "That's more than 400 lines - shorten the range.", vbExclamation, "Call Out Roster": Exit Function

    Dim data() As Variant: ReDim data(1 To nRows, 1 To CV_N)
    Dim base As Long: base = CLng(Val(Mid$(NextId(lo, "CV"), 4)))
    Dim d As Date, s As Variant, k As Long, user As String, stamp As Date
    user = CurrentUser(): stamp = Now
    For d = Int(dStart) To Int(dEnd)
        For Each s In shifts
            k = k + 1
            data(k, CV_ID) = "CV-" & Format$(base + k - 1, "0000")
            data(k, CV_DATE) = d
            data(k, CV_SHIFT) = s
            data(k, CV_UNIT) = unit
            data(k, CV_NEED) = needText
            data(k, CV_TYPE) = TypeCode(typ)
            If IsNumeric(hrs) And Not IsEmpty(hrs) Then
                If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then data(k, CV_HOURS) = CDbl(hrs)
            End If
            data(k, CV_SOURCE) = "Manual"
            data(k, CV_STATUS) = "Active"
            data(k, CV_USER) = user
            data(k, CV_STAMP) = stamp
        Next s
    Next d
    If Not AppendTo(lo, data, CV_N) Then Exit Function
    FormatCovDates lo
    SaveIfWritable
    SaveNeed = nRows
End Function

Public Sub FormatCovDates(ByVal lo As ListObject)
    On Error Resume Next                            ' never raises
    UnlockSheet lo.Parent
    lo.ListColumns(CV_DATE).DataBodyRange.NumberFormat = "ddd mmm d, yyyy"
    ProtectSheet lo.Parent
End Sub

' Adds one coverage line. slotKey: "A:AB-x" (covers that absence),
' "C:CV-x" (fills that need line if it's open / POSTED), or "NEW".
Public Function AddCoverage(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, _
                            ByVal slotKey As String, ByVal coveredBy As String, ByVal hrs As Variant, _
                            ByVal source As String, ByVal typ As String, ByVal needText As String) As String
    Dim lo As ListObject: Set lo = CovTable(True)
    If lo Is Nothing Then Exit Function
    LoadCache
    Dim v As Variant: v = mCov

    ' fill an open need line in place
    Dim r As Long
    ' (re-review B2) an absence whose posting closed with NO TAKERS: that live line is the one to fill, not a second one under it
    If Left$(slotKey, 2) = "A:" And IsArray(v) Then
        Dim ai As Long
        For ai = 1 To UBound(v, 1)
            If StrComp(SafeStr(v(ai, CV_ABS)), Mid$(slotKey, 3), vbTextCompare) = 0 And CovLive(ai, d, shiftLbl) Then
                If UCase$(SafeStr(v(ai, CV_BY))) Like "NO TAKERS*" Then slotKey = "C:" & SafeStr(v(ai, CV_ID)): Exit For
            End If
        Next ai
    End If
    If Left$(slotKey, 2) = "C:" Then
        r = RowOf(v, Mid$(slotKey, 3))
        If r > 0 Then
            Dim cur As String: cur = UCase$(SafeStr(v(r, CV_BY)))
            Dim curSrc As String: curSrc = SafeStr(v(r, CV_SOURCE))
            ' fill in place if empty, or if it's this same posting's POSTED line, or (v27, item 15) a NO TAKERS line -
            ' the posting closed with nobody, so the call-out that fills the shift takes the line over;
            ' a POSTED line from another source keeps its posting (new line below)
            If cur = "" Or (cur Like "POSTED*" And StrComp(curSrc, source, vbTextCompare) = 0) Or (cur Like "NO TAKERS*") Then
                UnlockSheet lo.Parent
                lo.DataBodyRange.Cells(r, CV_BY).Value = coveredBy
                If IsNumeric(hrs) And Not IsEmpty(hrs) Then
                    If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then
                        lo.DataBodyRange.Cells(r, CV_HOURS).Value = CDbl(hrs)
                    Else
                        lo.DataBodyRange.Cells(r, CV_HOURS).ClearContents      ' (re-review B4) a whole shift: no stale partial hours left on the line
                    End If
                End If
                If source <> "" Then
                    ' (re-review B3) a NO TAKERS line taken over by a call-out keeps its posting's id in the Note
                    If cur Like "NO TAKERS*" And curSrc <> "" And StrComp(curSrc, source, vbTextCompare) <> 0 Then
                        lo.DataBodyRange.Cells(r, CV_NOTE).Value = Trim$(SafeStr(v(r, CV_NOTE)) & " was " & curSrc & " (no takers)")
                    End If
                    lo.DataBodyRange.Cells(r, CV_SOURCE).Value = source
                End If
                ProtectSheet lo.Parent
                AddCoverage = SafeStr(v(r, CV_ID))
                Exit Function
            End If
            needText = SafeStr(v(r, CV_NEED))       ' already covered: add a second line under it
            typ = SafeStr(v(r, CV_TYPE))
        End If
    End If

    Dim absId As String
    If Left$(slotKey, 2) = "A:" Then
        absId = Mid$(slotKey, 3)
        Dim ar As Long: ar = RowOf(mAbs, absId)
        If ar > 0 Then
            If needText = "" Then needText = CalName(SafeStr(mAbs(ar, AB_EMP))) & " (" & TypeTag(SafeStr(mAbs(ar, AB_TYPE))) & ")"
            If typ = "" Then typ = SafeStr(mAbs(ar, AB_TYPE))
        End If
    End If
    ' coverage lines never carry the medical / parental / WSIB / LTD code - just "leave"
    If IsLeaveCode(typ) Then typ = "LOA"

    Dim id As String: id = NextId(lo, "CV")
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To CV_N)
    rec(1, CV_ID) = id
    rec(1, CV_DATE) = Int(d)
    rec(1, CV_SHIFT) = shiftLbl
    rec(1, CV_UNIT) = unit
    rec(1, CV_ABS) = absId
    rec(1, CV_NEED) = needText
    rec(1, CV_TYPE) = TypeCode(typ)
    rec(1, CV_BY) = coveredBy
    If IsNumeric(hrs) And Not IsEmpty(hrs) Then
        If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then rec(1, CV_HOURS) = CDbl(hrs)
    End If
    rec(1, CV_SOURCE) = source
    rec(1, CV_STATUS) = "Active"
    rec(1, CV_USER) = CurrentUser()
    rec(1, CV_STAMP) = Now
    If Not AppendTo(lo, rec, CV_N) Then Exit Function
    FormatCovDates lo
    AddCoverage = id
End Function

' Updates every live line from one source (posting award / no takers / cancel).
' Returns how many calendar lines it changed (0 = the source isn't on the calendar).
Public Function UpdateCoverageBySource(ByVal source As String, ByVal coveredBy As String, Optional ByVal hrs As Variant, _
                                       Optional ByVal clearHours As Boolean = False) As Long
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then Exit Function
    Dim v As Variant: v = TableData(lo)
    If Not IsArray(v) Then Exit Function
    Dim i As Long, hit As Boolean
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            If Not hit Then UnlockSheet lo.Parent: hit = True
            UpdateCoverageBySource = UpdateCoverageBySource + 1
            lo.DataBodyRange.Cells(i, CV_BY).Value = coveredBy
            If clearHours Then
                lo.DataBodyRange.Cells(i, CV_HOURS).ClearContents      ' v24: partial awards carry their hours in the text
            ElseIf Not IsMissing(hrs) Then
                If IsNumeric(hrs) And Not IsEmpty(hrs) Then
                    If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then lo.DataBodyRange.Cells(i, CV_HOURS).Value = CDbl(hrs)
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lo.Parent
    If hit Then MarkCalDirty
End Function

' A void in the OT Log: take empName (or everyone, if blank) off the calendar
' lines that came from source (S-#### call-out / P-#### posting).
' The line stays (it shows OPEN again). Returns how many lines changed.
Public Function ReleaseCoverage(ByVal source As String, Optional ByVal empName As String = "") As Long
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then Exit Function
    Dim v As Variant: v = TableData(lo)
    If Not IsArray(v) Then Exit Function
    Dim i As Long, hit As Boolean, by As String, newBy As String
    ' v27 (item 13, A11): a released line is never edited in place - it goes VOID and a live copy carries what is left;
    ' the copies are gathered and appended ONCE after the loop (review: AppendTo re-protects the sheet)
    Dim reps() As Variant, nRep As Long
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            by = SafeStr(v(i, CV_BY))
            If empName = "" Then newBy = "" Else newBy = DropName(by, CalName(empName))
            If newBy <> by Or empName = "" Then
                If Not hit Then UnlockSheet lo.Parent: hit = True
                lo.DataBodyRange.Cells(i, CV_STATUS).Value = "VOID"
                If Not (StrComp(SafeStr(v(i, CV_NEED)), "Call-out " & source, vbTextCompare) = 0 And newBy = "") Then
                    ' not a line the call-out itself made: a live copy keeps the need (OPEN again, or the one still on it),
                    ' its source kept so ReopenPosting can set it back to POSTED
                    lo.DataBodyRange.Cells(i, CV_NOTE).Value = Trim$(SafeStr(v(i, CV_NOTE)) & " released " & IIf(empName = "", "all", CalName(empName)) & " (" & source & " voided)")
                    nRep = nRep + 1
                    AddReplacement reps, nRep, v, i, newBy
                End If
                ReleaseCoverage = ReleaseCoverage + 1
            End If
        End If
    Next i
    If nRep > 0 Then
        Dim k As Long
        For k = 1 To nRep
            reps(k, CV_ID) = NextIdAfter(lo, "CV", k - 1)
        Next k
        If Not AppendTo(lo, reps, CV_N) Then
            ' (review) the copies did not land: the lines go back to Active, so no need vanishes from the calendar
            For k = 1 To UBound(v, 1)
                If StrComp(SafeStr(v(k, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(k, CV_STATUS))) <> "VOID" Then lo.DataBodyRange.Cells(k, CV_STATUS).Value = SafeStr(v(k, CV_STATUS))
            Next k
            ProtectSheet lo.Parent
            Err.Raise vbObjectError + 792, , "the released coverage lines could not be re-written"
        End If
        FormatCovDates lo
    End If
    If hit Then ProtectSheet lo.Parent
    If hit Then MarkCalDirty
End Function

' one replacement row into reps (every column as it was; Covered By = what is left; hours only when someone is still on it)
Private Sub AddReplacement(ByRef reps() As Variant, ByVal n As Long, ByRef v As Variant, ByVal i As Long, ByVal newBy As String)
    Dim tmp() As Variant, c As Long, k As Long
    If n = 1 Then
        ReDim reps(1 To 1, 1 To CV_N)
    Else
        ReDim tmp(1 To n, 1 To CV_N)
        For k = 1 To n - 1
            For c = 1 To CV_N
                tmp(k, c) = reps(k, c)
            Next c
        Next k
        reps = tmp
    End If
    For c = 1 To CV_N
        reps(n, c) = v(i, c)
    Next c
    reps(n, CV_BY) = newBy
    If newBy = "" Then reps(n, CV_HOURS) = Empty
    reps(n, CV_STATUS) = "Active"
    reps(n, CV_USER) = CurrentUser()
    reps(n, CV_STAMP) = Now
    reps(n, CV_NOTE) = "replaces " & SafeStr(v(i, CV_ID))
End Sub

' the k-th id after the table's next one ("CV-0007", "CV-0008", ...), for several rows appended in one go
Private Function NextIdAfter(ByVal lo As ListObject, ByVal prefix As String, ByVal offset As Long) As String
    Dim base As String: base = NextId(lo, prefix)
    Dim p As Long: p = InStrRev(base, "-")
    NextIdAfter = Left$(base, p) & Format$(Val(Mid$(base, p + 1)) + offset, String$(Len(base) - p, "0"))
End Function

' "Joe D / Fred C 8h" minus "Fred C" -> "Joe D"
Private Function DropName(ByVal by As String, ByVal nm As String) As String
    Dim parts As Variant, i As Long, t As String, core As String, outS As String
    parts = Split(Replace(by, ";", "/"), "/")
    For i = 0 To UBound(parts)
        t = Trim$(parts(i))
        core = t
        If core Like "* #h" Or core Like "* ##h" Or core Like "* #.#h" Or core Like "* ##.#h" Or core Like "* #.##h" Then
            core = Trim$(Left$(core, InStrRev(core, " ")))
        End If
        If StrComp(core, nm, vbTextCompare) <> 0 And t <> "" Then outS = outS & IIf(outS = "", "", " / ") & t
    Next i
    DropName = outS
End Function

' Put something on the calendar: asks which line it covers, then writes it.
' Returns e.g. "Fri Oct 23 Days - Brooke D (Leave)", or "" if left off.
Public Function CalendarPlace(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, _
                              ByVal what As String, ByVal coveredBy As String, ByVal hrs As Variant, _
                              ByVal source As String, ByVal typ As String, ByVal needText As String, _
                              Optional ByVal preferNew As Boolean = False, Optional ByVal slotKey As String = "", _
                              Optional ByVal autoPick As Boolean = False) As String
    Dim k As String: k = slotKey
    If k = "" And gSmokeRun Then k = "NEW"          ' v24.2: SmokeTest takes a new line without asking
    ' v27 (item 10, A5): a call-out acceptance or an award outcome covers the ONE open line of that group, date and
    ' shift without asking; the question (a typed number) stays for several open lines or none. A new posting
    ' never auto-picks: a support posting is usually its own line.
    If k <> "" And k <> "NEW" And autoPick Then
        If Not SlotKeyOpen(d, shiftLbl, unit, k) Then k = ""      ' (review L7) a key for a voided, covered or POSTED line: back to the pick
    End If
    If k = "" And autoPick Then k = OnlyOpenSlot(d, shiftLbl, unit)
    If k = "" Then k = PickSlot(d, shiftLbl, unit, what, preferNew)
    If k = "" Then Exit Function
    Dim lt As String, rt As String, lbl As String
    If k = "NEW" Then
        lbl = "new line" & IIf(needText <> "", " '" & needText & "'", "")
    ElseIf SlotTexts(k & "|" & CLng(Int(d)) & "|" & shiftLbl, lt, rt) Then
        lbl = lt
    Else
        lbl = "line"
    End If
    If AddCoverage(d, shiftLbl, unit, k, coveredBy, hrs, source, _
                   IIf(k = "NEW", typ, ""), IIf(k = "NEW", needText, "")) = "" Then Exit Function
    MarkCalDirty
    CalendarPlace = Format$(d, "ddd mmm d") & " " & shiftLbl & " - " & lbl
End Function

Public Sub VoidCoverageBySource(ByVal source As String)
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableData(lo)
    If Not IsArray(v) Then Exit Sub
    Dim i As Long, hit As Boolean
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            If Not hit Then UnlockSheet lo.Parent: hit = True
            lo.DataBodyRange.Cells(i, CV_STATUS).Value = "VOID"
        End If
    Next i
    If hit Then ProtectSheet lo.Parent
End Sub

' v27 (item 10): the key of the only OPEN line (nobody on it yet) of that UNIT (review M4: the group holds two units
' in one case, and a Phosphine acceptance must never land on Derivatives' line unasked) for the date and shift; ""
' when there are none or several. v29.0.3 (review 1.6): a NO TAKERS line counts as open too, as SlotKeyOpen reads it
' (the call-out fills it in place) - before, one NO TAKERS line and one OPEN line put the acceptance on the OPEN line
' unasked and left the NO TAKERS line red. Reads the cache SlotsFor loads; writes nothing.
Public Function OnlyOpenSlot(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String) As String
    Dim slots() As CalSlot, n As Long, i As Long, nOpen As Long, k As String
    n = SlotsFor(d, shiftLbl, GroupOfUnit(unit), slots)
    For i = 1 To n
        If slots(i).IsOpen Or (UCase$(slots(i).RightText) Like "NO TAKERS*") Then
            If StrComp(SlotUnit(slots(i).Key), unit, vbTextCompare) = 0 Then nOpen = nOpen + 1: k = slots(i).Key
        End If
    Next i
    If nOpen = 1 Then OnlyOpenSlot = k
End Function

' v27 (item 10, review L7): True when the key names a LIVE line of that date and shift that can take the person:
' OPEN (nobody on it) or NO TAKERS (the posting closed empty: the call-out fills it in place). A voided absence,
' a covered line or a POSTED line (its posting is still open: award it instead) says False.
Public Function SlotKeyOpen(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, ByVal key As String) As Boolean
    Dim slots() As CalSlot, n As Long, i As Long
    n = SlotsFor(d, shiftLbl, GroupOfUnit(unit), slots)
    For i = 1 To n
        If StrComp(slots(i).Key, key, vbTextCompare) = 0 Then
            SlotKeyOpen = slots(i).IsOpen Or (UCase$(slots(i).RightText) Like "NO TAKERS*")
            Exit Function
        End If
    Next i
End Function

' Pick which calendar line something covers. Returns the slot key,
' "NEW" for a new line, or "" to leave the calendar alone.
Public Function PickSlot(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, _
                         ByVal what As String, Optional ByVal preferNew As Boolean = False) As String
    Dim slots() As CalSlot, n As Long, i As Long, msg As String, pick As Long
    n = SlotsFor(d, shiftLbl, GroupOfUnit(unit), slots)
    If n > 12 Then n = 12
    msg = Clip(what, 90) & vbNewLine & Format$(d, "ddd mmm d") & " " & shiftLbl & " - " & GroupOfUnit(unit) & vbNewLine & vbNewLine
    For i = 1 To n
        msg = msg & i & ")  " & Clip(slots(i).LeftText, 30) & "  ->  " & _
              IIf(slots(i).RightText = "", "(open)", Clip(slots(i).RightText, 22)) & vbNewLine
    Next i
    msg = msg & "0)  new line" & vbNewLine & vbNewLine & "Type the number (Cancel = leave the calendar alone):"
    pick = AskNumber(msg, "Calendar", IIf(n > 0 And Not preferNew, 1, 0), 0, n)
    If pick < 0 Then Exit Function
    If pick = 0 Then PickSlot = "NEW": Exit Function
    PickSlot = slots(pick).Key
End Function

' Call-out: one accepted person onto the calendar
' v27 (item 10): slotKey = the line this call-out was started from (the sick call's absence, "A:SC-n"), so no question;
' with none, the one open line is taken; the question only when several or none. A Cancel is said in words a
' receipt can colour ("NOT ON THE CALENDAR").
Public Function CalendarAddAccepted(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String, _
                                    ByVal empName As String, ByVal hrs As Double, ByVal source As String, _
                                    Optional ByVal slotKey As String = "") As String
    On Error GoTo Failed
    Dim t As String
    t = CalendarPlace(d, shiftLbl, unit, "Call-out " & source & ": which calendar line does " & empName & " cover?", _
                      CalName(empName), hrs, source, "SUP", "Call-out " & source, False, slotKey, True)
    If t = "" Then
        CalendarAddAccepted = empName & ": NOT ON THE CALENDAR - open the Calendar, that shift's line > Set coverage"
    Else
        CalendarAddAccepted = empName & ": " & t
        SaveIfWritable
    End If
    RefreshCalendar
    Exit Function
Failed:
    CalendarAddAccepted = empName & ": NOT ON THE CALENDAR (" & Err.Description & ") - open the Calendar, that shift's line > Set coverage"
End Function

' ====================================================================
'  VOID / EDIT absences (Book Time Off pop-up)
' ====================================================================
Public Function PersonAbsences(ByVal emp As String, ByRef ids As Variant) As Variant
    ids = Empty
    Dim v As Variant: v = TableData(AbsTable())
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long, out() As String, idArr() As String
    ReDim out(0 To UBound(v, 1) - 1, 0 To 4): ReDim idArr(0 To UBound(v, 1) - 1)
    For i = UBound(v, 1) To 1 Step -1
        If StrComp(Trim$(SafeStr(v(i, AB_EMP))), emp, vbTextCompare) = 0 And UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" Then
            idArr(n) = SafeStr(v(i, AB_ID))
            out(n, 0) = idArr(n)
            out(n, 1) = SafeStr(v(i, AB_TYPE))
            out(n, 2) = Format$(v(i, AB_START), "mmm d, yyyy")
            out(n, 3) = IIf(IsDate(v(i, AB_END)), Format$(v(i, AB_END), "mmm d, yyyy"), "open")
            out(n, 4) = IIf(SafeStr(v(i, AB_HOURS)) = "", "full", SafeStr(v(i, AB_HOURS)) & "h") & _
                        IIf(SafeStr(v(i, AB_TIMES)) <> "", " " & SafeStr(v(i, AB_TIMES)), "")
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    Dim res() As String, j As Long, c As Long
    ReDim res(0 To n - 1, 0 To 4)
    For j = 0 To n - 1
        For c = 0 To 4
            res(j, c) = out(j, c)
        Next c
    Next j
    ReDim Preserve idArr(0 To n - 1)
    ids = idArr
    PersonAbsences = res
End Function

Public Function SetAbsenceEnd(ByVal id As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = AbsTable()
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Function
    Dim cur As String: cur = IIf(IsDate(v(r, AB_END)), DateText(CDate(v(r, AB_END))), "")
    Dim s As Variant
    s = Application.InputBox("New end date for " & id & " (" & SafeStr(v(r, AB_EMP)) & ", " & SafeStr(v(r, AB_TYPE)) & ")." & _
                             vbNewLine & "Leave blank for open-ended.", "End date", cur, Type:=2)
    If VarType(s) = vbBoolean Then Exit Function
    s = Trim$(CStr(s))
    Dim dEnd As Date                               ' v25: the one parser
    If s <> "" Then
        If Not ParseTypedDate(s, dEnd) Then MsgBox "That isn't a date (e.g. Oct 12).", vbExclamation, "Call Out Roster": Exit Function
        If dEnd < CDate(v(r, AB_START)) Then MsgBox "The end is before the start.", vbExclamation, "Call Out Roster": Exit Function
    End If
    UnlockSheet lo.Parent
    ' v27 (item 13, A11): the previous end stays in the Note (shown by the calendar's line details: no login in it)
    lo.DataBodyRange.Cells(r, AB_NOTE).Value = Trim$(SafeStr(v(r, AB_NOTE)) & IIf(Trim$(SafeStr(v(r, AB_NOTE))) = "", "", " | ") & _
                                               "End was " & IIf(cur = "", "open-ended", cur) & " (changed " & Format$(Now, "mmm d") & ")")   ' no login: the calendar shows this Note to every user
    If s = "" Then
        lo.DataBodyRange.Cells(r, AB_END).ClearContents
    Else
        lo.DataBodyRange.Cells(r, AB_END).Value = dEnd
    End If
    ProtectSheet lo.Parent
    SaveIfWritable
    SetAbsenceEnd = True
End Function

Public Function VoidAbsence(ByVal id As String, Optional ByVal reasonTxt As String = "") As Boolean
    ' reasonTxt (v24.2): given by code (SmokeTest); left out, the reason is asked for
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = AbsTable()
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Function
    Dim why As String: why = reasonTxt
    If Trim$(why) = "" Then why = InputBox("Why is " & id & " (" & SafeStr(v(r, AB_EMP)) & ") being removed?", "Void time off")
    If Trim$(why) = "" Then Exit Function
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(r, AB_STATUS).Value = "VOID"
    lo.DataBodyRange.Cells(r, AB_VOIDNOTE).Value = Format$(Now, "yyyy-mm-dd hh:nn") & " " & CurrentUser() & ": " & Trim$(why)
    ProtectSheet lo.Parent
    RequestBookingVoided id, why                    ' v29 (item 6): a member's request this booking came from reads "Booking voided"
    ' its coverage lines stay (they show as plain lines) so nothing awarded is lost
    SaveIfWritable
    VoidAbsence = True
End Function

' ====================================================================
'  CALENDAR SHEET
' ====================================================================
' Alt+F8 > CalendarCheck: shows which calendar code is installed, then redraws
Public Sub CalendarCheck()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    MarkCalDirty
    MsgBox "Installed: " & CAL_VERSION & vbNewLine & _
           "ThisWorkbook events: " & IIf(mEventsSeen, "working", _
               "not seen yet - switch tabs once, or re-paste ThisWorkbook.txt if this stays") & _
           vbNewLine & vbNewLine & "Click OK to redraw the calendar.", vbInformation, "Calendar"
    ShowCalendar
    Exit Sub
Oops:
    ReportError "CalendarCheck"
End Sub

Public Sub ShowCalendar()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = CalendarSheet()
    If gAdminUnlocked Then
        On Error Resume Next
        Dim tSync As Double: tSync = Timer
        SyncSwitchesToCalendar           ' switches logged before the calendar existed
        MapWrite "Z16", "'" & "Open path: the switch sync before the draw took " & Format$(Secs(tSync, Timer), "0.0") & " s"   ' v25.1.2 (review: outside the draw's own time)
        On Error GoTo Oops
    End If
    SetSheetVisible ws, xlSheetVisible
    ws.Activate                          ' the tab-switch check runs here (ThisWorkbook)
    If CalStale() Then DrawCalendar Else SetCalHint ws
    Exit Sub
Oops:
    ReportError "ShowCalendar"
End Sub

' Something changed: mark the calendar out of date. Only the tab you're
' looking at is redrawn now; the other one redraws when you next open it.
Public Sub RefreshCalendar()
    If Not SheetExists(CAL_SHEET) Then Exit Sub
    On Error Resume Next
    MarkCalDirty
    If ActiveSheet.Name = CAL_SHEET Then
        DrawCalendar
    ElseIf ActiveSheet.Name = NA_SHEET Then
        DrawNeeds
    ElseIf Not mEventsSeen Then
        ' ThisWorkbook's tab-switch redraw hasn't been seen this session (not
        ' pasted in?) - redraw now so the calendar is never left stale
        Dim was As Object: Set was = ActiveSheet
        DrawCalendar
        If Not was Is Nothing Then was.Activate
    End If
    On Error GoTo 0
End Sub

' ====================================================================
'  CHANGE TRACKING  - redraw only when something the calendar shows changed
'  CalMap Z7  = change counter (bumped whenever Absences, Coverage,
'               Postings or Shift Switches change - by a macro or by hand)
'  Z8 / Z9    = counter the Calendar / Needs Attention were drawn from
'  Z10 / Z11  = date + code version they were drawn under
'  Z12        = roster names + crews they were drawn from
' ====================================================================
Public Sub MarkCalDirty()
    On Error Resume Next                            ' never raises
    If Not SheetExists(MAP_SHEET) Then Exit Sub
    MapWrite "Z7", CLng(Val(MapSheet().Range("Z7").Value)) + 1     ' v25.1: through the pair
End Sub

' called from ThisWorkbook on every sheet change - cheap name check only
Public Sub CalSheetChanged(ByVal Sh As Object)
    mEventsSeen = True
    Select Case Sh.Name
        Case AB_SHEET, CV_SHEET, POSTINGS_SHEET, SS_SHEET
            MarkCalDirty
        Case HOL_SHEET                      ' an Observed on date was entered / changed (v20)
            ResetHolidayCache
            MarkCalDirty
        Case MINSTAFF_SHEET                 ' v26: a minimum was typed - the SHORT flags follow at the next draw
            MarkCalDirty
    End Select
End Sub

Public Function DrawContext() As String
    DrawContext = Format$(Date, "yyyymmdd") & "|" & CAL_VERSION
End Function

' order-independent fingerprint of roster name + crew (crew changes move time off)
Private Function RosterSig() As String
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_SHIFT)).Value
    Dim i As Long, j As Long, t As String, h As Double, tot As Double
    For i = 1 To UBound(v, 1)
        t = UCase$(Trim$(SafeStr(v(i, 1)))) & "=" & UCase$(Trim$(SafeStr(v(i, UBound(v, 2)))))
        h = 0
        For j = 1 To Len(t)
            h = h * 31 + AscW(Mid$(t, j, 1))
            h = h - Int(h / 1000003) * 1000003
        Next j
        tot = tot + h
    Next i
    RosterSig = CStr(UBound(v, 1)) & ":" & CStr(tot)
End Function

Private Function CalStale() As Boolean
    On Error GoTo Yes
    Dim mp As Worksheet: Set mp = MapSheet()
    If SafeStr(mp.Range("Z8").Value) <> SafeStr(mp.Range("Z7").Value) Then GoTo Yes
    If SafeStr(mp.Range("Z10").Value) <> DrawContext() Then GoTo Yes
    If SafeStr(mp.Range("Z12").Value) <> RosterSig() Then GoTo Yes
    Exit Function
Yes:
    CalStale = True
End Function

Public Function NAStale() As Boolean
    On Error GoTo Yes
    Dim mp As Worksheet: Set mp = MapSheet()
    If SafeStr(mp.Range("Z9").Value) <> SafeStr(mp.Range("Z7").Value) Then GoTo Yes
    If SafeStr(mp.Range("Z11").Value) <> DrawContext() & IIf(LiteShowing(), "|lite", "") Then GoTo Yes    ' v29: a lite login's order is its own context
    Exit Function
Yes:
    NAStale = True
End Function

' the toolbar hint depends on who's looking - cheap to fix without a redraw
' v25.1.2: the hint text, with the last draw's time on the end (CalMap Z13 holds the seconds)
Private Function CalHintText() As String
    ' v25.5 (bundle item 0): a click shows the line's details here, no hover boxes
    CalHintText = IIf(gAdminUnlocked, "Click a line: its details show here.  Double-click any line or empty space for options.", _
                      "Click a line: its details show here.  Double-click a line for options.  Print / PDF gives the copy to post.")
    Dim s As String: s = SafeStr(MapSheet().Range("Z13").Value)
    If s <> "" Then CalHintText = CalHintText & "  Drawn in " & s & " s."
End Function

' v25.1.2: seconds between two Timer readings (Timer restarts at midnight)
Public Function Secs(ByVal t0 As Double, ByVal t1 As Double) As Double
    Secs = t1 - t0
    If Secs < 0 Then Secs = Secs + 86400
End Function

' v25.1.2: Alt+F8 (admin): the last draw's time per phase for the Calendar and Needs Attention tabs
Public Sub ShowCalDrawTimes()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not SheetExists(MAP_SHEET) Then MsgBox "The calendar hasn't been drawn yet.", vbInformation, "Draw times": Exit Sub
    Dim a As String: a = SafeStr(MapSheet().Range("Z14").Value)
    Dim b As String: b = SafeStr(MapSheet().Range("Z15").Value)
    Dim c As String: c = SafeStr(MapSheet().Range("Z16").Value)
    If a = "" Then a = "Calendar: not drawn since v25.1.2 went in (open the Calendar tab, click Next month)."
    If b = "" Then b = "Needs Attention: not drawn since v25.1.2 went in."
    If c = "" Then c = "Open path: the Calendar button hasn't been clicked since v25.1.2 went in."
    MsgBox a & vbNewLine & vbNewLine & b & vbNewLine & vbNewLine & c & vbNewLine & vbNewLine & _
           "Send " & PRIMARY_ADMIN_NAME & " these lines (a photo is fine). The 'drawn in' number is the draw alone: a stopwatch from the click to the screen settling is the number you feel.  (" & CAL_VERSION & ")", _
           vbInformation, "Draw times"
    Exit Sub
Oops:
    ReportError "ShowCalDrawTimes"
End Sub

Private Sub SetCalHint(ByVal ws As Worksheet)
    On Error Resume Next                            ' never raises
    Dim t As String
    t = CalHintText()                                       ' v25.1.2: with the last draw's time
    If SafeStr(ws.Range("F3").Value) = t Then Exit Sub
    If Left$(SafeStr(ws.Range("F3").Value), 8) = "Calendar" Then Exit Sub   ' leave an error message showing
    UnlockSheet ws
    ws.Range("F3").Value = t
    ProtectSheet ws
End Sub

' v25.5 (bundle item 0, Graeme Oct 17): from ThisWorkbook's SheetSelectionChange. One click on a calendar
' line writes its LineDetails text into the hint row (F3:K3, wrapped) and onto the status bar; a click on an
' empty cell puts the hint back. Row 3 only - the click map and the grid are not touched, and a draw in
' progress, a failed draw's red line (SetCalHint's rule) and any other sheet are left alone. The caches are
' reloaded only when the calendar's change counter or the draw context moved since the last click, so a
' second click costs one map read and the text. Never raises.
Public Sub CalSelectionChanged(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Quit
    If mDrawing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Sh.Name <> CAL_SHEET Then Exit Sub
    If Not SheetExists(MAP_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = Sh
    If Target.Cells(1, 1).Row <= 3 Then Exit Sub        ' the title, legend and hint rows: the details stay readable (review)
    Dim keyTxt As String: keyTxt = SafeStr(MapSheet().Range(Target.Cells(1, 1).Address).Value)
    Dim txt As String
    If keyTxt <> "" Then
        Dim st As String: st = CalDirtyStamp() & "|" & DrawContext()
        If st <> mDetailStamp Or Not IsArray(mAbs) Then
            LoadCache
            LoadPostCache
            mDetailStamp = st
        End If
        txt = LineDetails(keyTxt, True)
    End If
    If txt = "" Then
        SetCalHint ws                                   ' the hint back (nothing written when it is showing already)
        ShowBarText ""
        Exit Sub
    End If
    txt = Replace(txt, vbLf, "   |   ")
    ShowBarText txt                                     ' the status bar first: it shows even when row 3 is kept (review)
    If Left$(SafeStr(ws.Range("F3").Value), 8) = "Calendar" Then Exit Sub   ' a failed draw's message stays
    If SafeStr(ws.Range("F3").Value) <> txt Then
        ' the write under Resume Next and the sheet protected again whatever happened (convention 17, review)
        UnlockSheet ws
        On Error Resume Next
        ws.Range("F3").Value = txt
        On Error GoTo Quit
        ProtectSheet ws
    End If
    Exit Sub
Quit:
    On Error Resume Next
    ProtectSheet ws
End Sub

' v25.5: the hint back in row 3 (the close: a clicked line's details are not saved for the next opener)
Public Sub CalHintReset()
    On Error Resume Next                            ' never raises
    If Not SheetExists(CAL_SHEET) Then Exit Sub
    If Not SheetExists(MAP_SHEET) Then Exit Sub
    SetCalHint ThisWorkbook.Worksheets(CAL_SHEET)
End Sub

' True if a live coverage line came from this source (SS- / P- / S- id)
Public Function CoverageHasSource(ByVal source As String) As Boolean
    Dim v As Variant: v = TableData(CovTable())
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            CoverageHasSource = True: Exit Function
        End If
    Next i
End Function

Public Function CalendarSheet() As Worksheet
    If Not SheetExists(CAL_SHEET) Then
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=RosterSheet())
        ws.Name = CAL_SHEET
        RelockStructure
        If SheetExists(MAP_SHEET) Then MapWrite "Z17", ""       ' v25.2 re-review: a fresh sheet needs its page setup again
    End If
    If Not SheetExists(MAP_SHEET) Then
        Dim m As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set m = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        m.Name = MAP_SHEET
        RelockStructure
        SetSheetVisible m, xlSheetVeryHidden
    End If
    Set CalendarSheet = ThisWorkbook.Worksheets(CAL_SHEET)
    SetTabColor CalendarSheet, thGreen()            ' v25.1: a structure change under v25's tab protection
End Function

Public Function MapSheet() As Worksheet
    Set MapSheet = ThisWorkbook.Worksheets(MAP_SHEET)
End Function

' v25.1: CalMap is protected like every sheet (v25), and a bare write to it raised "Application-defined
' or object-defined error" from the Calendar's unit list on Graeme's Oct 16 test. Every other sheet is
' written between UnlockSheet / ProtectSheet; the map now is too (DrawCalendar unlocks it for the whole
' draw). The map is relocked whatever happened and the error re-raised after (v25.1 review: MarkCalDirty,
' CalOnOpen and DrawNeeds never reach ReportError, so nothing else would relock it).
Public Sub MapWrite(ByVal addr As String, ByVal v As Variant)
    Dim mp As Worksheet: Set mp = MapSheet()
    Dim en As Long, em As String
    UnlockSheet mp
    On Error Resume Next
    mp.Range(addr).Value = v
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    ProtectSheet mp
    If en <> 0 Then Err.Raise en, , em
End Sub

' v25.1: the dirty counter as text, so SmokeTest can prove MarkCalDirty really wrote it; a roster where
' CalendarSheet() has never run has no map (MarkCalDirty does nothing there) and says so
Public Function CalDirtyStamp() As String
    CalDirtyStamp = "(no CalMap)"
    If Not SheetExists(MAP_SHEET) Then Exit Function
    On Error Resume Next                            ' never raises
    CalDirtyStamp = SafeStr(MapSheet().Range("Z7").Value)
End Function

Public Function CalMonth() As Date
    Dim v As Variant: v = MapSheet().Range("Z1").Value
    If IsDate(v) Then CalMonth = DateSerial(Year(v), Month(v), 1) Else CalMonth = DateSerial(Year(Date), Month(Date), 1)
End Function

Public Function CalGroup() As String
    CalGroup = SafeStr(MapSheet().Range("Z2").Value)
    If CalGroup = "" Then CalGroup = CalGroupNames()(0)
End Function

' the one unit of a calendar group, or "" when it holds several
Public Function SingleUnitOf(ByVal grp As String) As String
    Dim g As Variant, i As Long, parts As Variant
    g = CalGroups()
    For i = 0 To UBound(g)
        parts = Split(g(i), "|")
        If StrComp(parts(0), grp, vbTextCompare) = 0 Then
            If parts(1) <> "*" And InStr(parts(1), ",") = 0 Then SingleUnitOf = Trim$(parts(1))
            Exit Function
        End If
    Next i
End Function

Public Sub CalPrev()
    On Error GoTo Oops
    CalendarSheet
    MapWrite "Z1", DateAdd("m", -1, CalMonth())
    DrawCalendar
    Exit Sub
Oops:
    ReportError "CalPrev"
End Sub

Public Sub CalNext()
    On Error GoTo Oops
    CalendarSheet
    MapWrite "Z1", DateAdd("m", 1, CalMonth())
    DrawCalendar
    Exit Sub
Oops:
    ReportError "CalNext"
End Sub

Public Sub CalToday()
    On Error GoTo Oops
    CalendarSheet
    MapWrite "Z1", DateSerial(Year(Date), Month(Date), 1)
    DrawCalendar
    Exit Sub
Oops:
    ReportError "CalToday"
End Sub

Public Sub CalUnitChanged()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = CalendarSheet()
    Dim dd As Object
    On Error Resume Next
    Set dd = ws.DropDowns("cal_unit")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z2", dd.List(dd.ListIndex)
    DrawCalendar
    Exit Sub
Oops:
    ReportError "CalUnitChanged"
End Sub

Public Sub CalMonthChanged()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = CalendarSheet()
    Dim dd As Object
    On Error Resume Next
    Set dd = ws.DropDowns("cal_month")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z1", DateSerial(Year(Date) - 1, dd.ListIndex, 1)   ' list starts January last year
    DrawCalendar
    Exit Sub
Oops:
    ReportError "CalMonthChanged"
End Sub

' ====================================================================
'  THEME  - every calendar / needs-attention colour is set here
' ====================================================================
Public Function thBand() As Long          ' title band, borders around each day
    thBand = RGB(38, 50, 56)
End Function
Public Function thHead() As Long          ' weekday header, section bars
    thHead = RGB(55, 71, 79)
End Function
Public Function thNav() As Long           ' month / navigation buttons
    thNav = RGB(84, 110, 122)
End Function
Public Function thGreen() As Long         ' print / post / primary actions
    thGreen = RGB(46, 125, 50)
End Function
Private Function thDayFill() As Long
    thDayFill = RGB(255, 248, 225)
End Function
Private Function thDayTab() As Long
    thDayTab = RGB(255, 224, 130)
End Function
Private Function thDayText() As Long
    thDayText = RGB(150, 100, 0)
End Function
Private Function thNightFill() As Long
    thNightFill = RGB(232, 234, 246)
End Function
Private Function thNightTab() As Long
    thNightTab = RGB(197, 202, 233)
End Function
Private Function thNightText() As Long
    thNightText = RGB(40, 53, 147)
End Function
Public Function thToday() As Long
    thToday = RGB(255, 213, 79)
End Function
Private Function thWeekend() As Long
    thWeekend = RGB(236, 239, 241)
End Function
Private Function thOutFill() As Long       ' days outside the month
    thOutFill = RGB(245, 245, 245)
End Function
Public Function thGrid() As Long
    thGrid = RGB(207, 216, 220)
End Function
Public Function thOpen() As Long          ' OPEN / no takers
    thOpen = RGB(183, 28, 28)
End Function
Public Function thOpenFill() As Long
    thOpenFill = RGB(255, 205, 210)
End Function
Public Function thPosted() As Long        ' POSTED - waiting for sign-ups
    thPosted = RGB(191, 54, 12)
End Function
Public Function thPostedFill() As Long
    thPostedFill = RGB(255, 224, 178)
End Function
Private Function thInk() As Long           ' normal text / covered names
    thInk = RGB(33, 33, 33)
End Function
Public Function thMuted() As Long
    thMuted = RGB(117, 117, 117)
End Function

' a colour mixed toward white (f = 0.85 -> very light tint)
Public Function Tint(ByVal col As Long, ByVal f As Double) As Long
    Dim rr As Long, gg As Long, bb As Long
    rr = col Mod 256: gg = (col \ 256) Mod 256: bb = (col \ 65536) Mod 256
    Tint = RGB(rr + (255 - rr) * f, gg + (255 - gg) * f, bb + (255 - bb) * f)
End Function

' ====================================================================
'  DRAW THE CALENDAR
' ====================================================================
Public Sub DrawCalendar()
    Dim ws As Worksheet: Set ws = CalendarSheet()
    Dim mp As Worksheet: Set mp = MapSheet()
    Dim m0 As Date: m0 = CalMonth()
    Dim grp As String: grp = CalGroup()
    Dim drawnFrom As String: drawnFrom = SafeStr(mp.Range("Z7").Value)
    ' v25.1.2: draw-time measurement, one Timer reading per phase (kept in CalMap Z13 / Z14, shown in F3)
    Dim t0 As Double, tA As Double, tB As Double, tC As Double, tD As Double, tE As Double, tF As Double, tG As Double
    Dim nLines As Long                              ' the lines drawn, counted by CalDrawWeeks (for Z14)
    t0 = Timer
    mTimeSlots = 0: mTimeNotes = 0

    mDrawing = True
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim calcWas As Long: calcWas = Application.Calculation
    Application.EnableEvents = False            ' no change events for every cell we write
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Failed
    UnlockSheet ws
    UnlockSheet mp                                  ' v25.1: written all through the draw; relocked at Done (under the handler, review)
    mp.Range("Z1").Value = m0
    mp.Range("Z2").Value = grp
    LoadCache
    LoadPostCache
    StaffSweepBegin Date                            ' v26: the minimums and the names per unit, once per draw
    tA = Timer                                      ' caches loaded

    CalDrawWipe ws, mp                              ' v27.1: the grid and the click map wiped
    tB = Timer                                      ' grid wiped

    Dim r As Long
    CalDrawHeader ws, mp, r                         ' v27.1: rows 1-4; r = the first week row
    tC = Timer                                      ' header rows written
    mTimeSlots = 0: mTimeNotes = 0                  ' the weeks' share starts here

    CalDrawWeeks ws, mp, m0, grp, r, nLines         ' v27.1: the weeks; r moves past the last one
    tD = Timer                                      ' weeks written
    Dim wkSlots As Double, wkNotes As Double
    wkSlots = mTimeSlots: wkNotes = mTimeNotes      ' the weeks' share (OpenAhead's SlotsFor calls count under controls)

    CalDrawPage ws, mp, r                           ' v27.1: the page setup and the window (Resume Next inside)
    On Error GoTo Failed
    tE = Timer                                      ' page set up

    BuildCalControls ws, m0, grp
    tF = Timer                                      ' controls rebuilt
    mp.Range("Z8").Value = drawnFrom
    mp.Range("Z10").Value = "'" & DrawContext()
    mp.Range("Z12").Value = "'" & RosterSig()
    ' v25.1.2: the breakdown (the before numbers for the v25.2 rewrite); F3 says the total
    tG = Timer
    mp.Range("Z13").Value = "'" & Format$(Secs(t0, tG), "0.0")
    mp.Range("Z14").Value = "'" & "Calendar " & Format$(m0, "mmmm yyyy") & " (" & grp & ", " & nLines & " lines): drawn in " & _
        Format$(Secs(t0, tG), "0.0") & " s = caches " & Format$(Secs(t0, tA), "0.0") & ", wipe " & Format$(Secs(tA, tB), "0.0") & _
        ", header " & Format$(Secs(tB, tC), "0.0") & ", weeks " & Format$(Secs(tC, tD), "0.0") & " (slots " & Format$(wkSlots, "0.0") & _
        ", notes " & Format$(wkNotes, "0.0") & ", cells " & Format$(Secs(tC, tD) - wkSlots - wkNotes, "0.0") & "), page " & _
        Format$(Secs(tD, tE), "0.0") & ", controls " & Format$(Secs(tE, tF), "0.0") & "  [" & CAL_VERSION & "]"
    ws.Range("F3").Value = CalHintText()
    GoTo Done

Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description & " (error " & Err.Number & ")"
    Resume AfterFail                      ' leave the handler so the clean-up below can't be stopped by a 2nd error
AfterFail:
    On Error Resume Next
    tG = 0                                ' v25.1.2 re-review: no after-draw figure on a failed draw
    LogError "DrawCalendar", en, em       ' v24.3.1: the box below was the only record
    UnlockSheet ws
    ws.Cells(3, 6).Value = "Calendar couldn't finish drawing: " & em & "  - click This month to retry."
    ws.Cells(3, 6).Font.Color = thOpen()
    BuildCalControls ws, m0, grp          ' never leave the sheet without its buttons
    Application.ScreenUpdating = True
    MsgBox "The calendar couldn't finish drawing:" & vbNewLine & vbNewLine & em & vbNewLine & vbNewLine & _
           "Send a screenshot of this message.  (" & CAL_VERSION & ")", vbExclamation, "Calendar"
Done:
    On Error Resume Next
    Application.DisplayAlerts = True
    ProtectSheet ws
    ProtectSheet mp                                 ' v25.1
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
    If tG > 0 Then                                  ' v25.1.2 (review): the relock + recalculation after the draw, felt but outside "drawn in"
        MapWrite "Z14", "'" & SafeStr(mp.Range("Z14").Value) & "; then relock + recalc " & Format$(Secs(tG, Timer), "0.0")
    End If
    mDrawing = False
End Sub

' v27.1: DrawCalendar's steps. No handler of their own: an error rises to DrawCalendar's Failed, which
' restores everything. Called only while that handler is active.
Private Sub CalDrawWipe(ByVal ws As Worksheet, ByVal mp As Worksheet)
    ' ---- wipe the grid and the click map (controls are rebuilt at the end) ----
    ws.Rows("3:1000").Delete
    ws.Range("A1:P2").UnMerge
    ws.Range("A1:P2").Clear
    mp.Range("A1:P1000").ClearContents
    With ws.Cells.Font
        .Name = CAL_FONT
        .Size = 9
    End With
    ws.Columns(1).ColumnWidth = 4
    Dim k As Long
    For k = 0 To 13
        ws.Columns(2 + k).ColumnWidth = COL_W
    Next k
    ws.Columns(16).ColumnWidth = 1
End Sub

Private Sub CalDrawHeader(ByVal ws As Worksheet, ByVal mp As Worksheet, ByRef r As Long)
    Dim k As Long
    ' ---- row 1: title band (controls sit on it) ----
    ws.Rows(1).RowHeight = 38
    ws.Range("A1:P1").Interior.Color = thBand()
    With ws.Range("B1:C1")
        .Merge
        .Value = "COVERAGE CALENDAR" & IIf(mPublicDraw, "   -   members' copy", "")   ' v26.5: the printed month is masked
        .Font.Name = CAL_FONT_B: .Font.Size = 15: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With

    ' ---- row 2: legend chips ----
    ws.Rows(2).RowHeight = 24
    Dim lg As Variant, i As Long, parts() As String
    lg = Array("VAC|Vacation", "LOA|Leave", "TRN|Training", "PHA|PHA", "SUP|Support op", "SS|Shift switch")
    For i = 0 To UBound(lg)
        parts = Split(lg(i), "|")
        LegendChip ws.Cells(2, 2 + i), parts(1), TypeColor(parts(0)), Tint(TypeColor(parts(0)), 0.88)
    Next i
    LegendChip ws.Cells(2, 9), "OPEN  needs cover", thOpen(), thOpenFill()
    LegendChip ws.Cells(2, 10), "POSTED  posting out", thPosted(), thPostedFill()
    LegendChip ws.Cells(2, 11), "Name = covered", thInk(), RGB(245, 245, 245)
    LegendChip ws.Cells(2, 12), "NCN  no cover needed", thMuted(), RGB(245, 245, 245)

    ' ---- row 3: admin toolbar (buttons added in BuildCalControls); the hint / a clicked line's details ----
    ws.Rows(3).RowHeight = CAL_HINT_ROW_H               ' v25.5: taller, wrapped - a clicked line's details show here
    With ws.Range("F3:K3")
        .Merge
        .Value = CalHintText()                  ' v25.2: one text (rewritten with this draw's time at the end)
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = thMuted()
        .VerticalAlignment = xlTop                      ' v25.5 review: long details keep their first line
        .IndentLevel = 1
        .WrapText = True
    End With
    With ws.Range("M3:O3")
        .Merge
        .Value = CAL_VERSION
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        .Font.Size = 7: .Font.Color = RGB(176, 190, 197)
    End With
    mp.Range("Z4").Value = 3                         ' rows 3..Z4 are screen-only (hidden when printing)

    ' ---- row 4: weekday header ----
    r = 4
    ws.Rows(r).RowHeight = 20
    ws.Cells(r, 1).Interior.Color = thHead()
    Dim dn As Variant: dn = Array("SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY")
    For k = 0 To 6
        With ws.Range(ws.Cells(r, 2 + 2 * k), ws.Cells(r, 3 + 2 * k))
            .Merge
            .Value = dn(k)
            .Font.Name = CAL_FONT_B: .Font.Size = 9: .Font.Color = vbWhite
            .Interior.Color = thHead()
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next k
    r = r + 1
End Sub

Private Sub CalDrawWeeks(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal m0 As Date, ByVal grp As String, _
                         ByRef r As Long, ByRef nLines As Long)
    Dim k As Long, i As Long
    Dim shortTxt As String, nShortD As Long          ' v26: the date's short shifts (minimum staffing)
    ' ---- weeks ----
    ' v25.2 (item 3): each week block is built in two arrays (the texts for B:O and the click-map keys) and
    ' written with ONE .Value each; the formats go on by class (one Union per class per block) instead of
    ' per cell; the date cells' rich text stays per cell. (The v25.1.2 cell-by-cell writer it replaced was kept
    ' one release for SmokeTest's compare step, which passed on the real roster Oct 17; gone in v25.5.) The
    ' numbers: cells 0.7 s of 4.6.
    Dim mEnd As Date: mEnd = DateSerial(Year(m0), Month(m0) + 1, 0)
    Dim wk As Date: wk = m0 - (Weekday(m0, vbSunday) - 1)
    Dim slotsD() As CalSlot, slotsN() As CalSlot
    Dim d As Date, nD As Long, nN As Long, mxD As Long, mxN As Long, inMonth As Boolean
    Dim nOpen(0 To 6) As Long, sv As Long
    Dim vals() As Variant, keys() As Variant, nRows As Long, c As Long
    Dim numLen(0 To 6) As Long, flagLen(0 To 6) As Long, holLen(0 To 6) As Long, past(0 To 6) As Boolean
    Dim rgDay As Range, rgNight As Range, rgOut As Range, rgLine As Range, rgOpen As Range, rgPosted As Range
    Dim rgNcn As Range, rgCov As Range, rgDateA As Range, rgDateB As Range, rgHol As Range
    Dim rgBg(0 To 3) As Range                       ' the date row's fills: 0 out of month, 1 today, 2 weekend, 3 white
    Dim rgLeft As Object: Set rgLeft = CreateObject("Scripting.Dictionary")   ' type colour -> the left cells in it
    Dim colKey As Variant, t As String, numTxt As String, flagTxt As String, holTxt As String, bgIx As Long
    Dim blockTop As Long, blockBot As Long
    Do While wk <= mEnd
        mxD = 1: mxN = 1
        For k = 0 To 6
            d = wk + k
            nOpen(k) = 0
            If Month(d) = Month(m0) Then
                nD = SlotsFor(d, "Days", grp, slotsD, True)
                nN = SlotsFor(d, "Nights", grp, slotsN, True)
                If nD > mxD Then mxD = nD
                If nN > mxN Then mxN = nN
                nOpen(k) = OpenCount(slotsD, nD) + OpenCount(slotsN, nN)
            End If
        Next k
        nRows = 1 + mxD + mxN                       ' the date row, then the day slots, then the night slots
        ReDim vals(1 To nRows, 1 To 14): ReDim keys(1 To nRows, 1 To 14)
        ws.Rows(r).RowHeight = 38                   ' the heights first (review), then the block
        ws.Range(ws.Rows(r + 1), ws.Rows(r + nRows - 1)).RowHeight = 17
        Set rgDay = Nothing: Set rgNight = Nothing: Set rgOut = Nothing: Set rgLine = Nothing
        Set rgOpen = Nothing: Set rgPosted = Nothing: Set rgNcn = Nothing: Set rgCov = Nothing
        Set rgDateA = Nothing: Set rgDateB = Nothing: Set rgHol = Nothing
        For bgIx = 0 To 3
            Set rgBg(bgIx) = Nothing
        Next bgIx
        rgLeft.RemoveAll

        ' every day of the week into the arrays; the cells of each class collected on the way
        For k = 0 To 6
            d = wk + k
            inMonth = (Month(d) = Month(m0))
            c = 2 + 2 * k                           ' the day's left column (B, D, ... N)
            If Not inMonth Then
                bgIx = 0
            ElseIf Int(d) = Int(Date) Then
                bgIx = 1
            ElseIf k = 0 Or k = 6 Then
                bgIx = 2
            Else
                bgIx = 3
            End If
            AddToRange rgBg(bgIx), ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
            If inMonth Then
                sv = CLng(Int(d))
                numTxt = CStr(Day(d)): flagTxt = ""
                If nOpen(k) > 0 Then flagTxt = "   " & ChrW(9679) & " " & nOpen(k) & " open"
                If Int(d) >= Int(Date) And Int(d) < Int(Date) + BOARD_DAYS Then     ' v26: minimum staffing, the next 14 days
                    shortTxt = StaffShortText(d, "", nShortD)         ' v26.5: EVERY unit, whatever group is shown (Graeme, Oct 19); the cell keeps a short flag;
                    If nShortD > 0 Then flagTxt = flagTxt & "   " & ChrW(9650) & " " & nShortD & " short"   ' a click shows which (LineDetails)
                End If
                vals(1, 2 * k + 1) = numTxt & flagTxt & vbLf & "Days: " & DaysCrew(d) & " Shift"
                holTxt = HolidayLabel(d)
                t = "Nights: " & NightsCrew(d) & " Shift"
                If holTxt <> "" Then vals(1, 2 * k + 2) = holTxt & vbLf & t Else vals(1, 2 * k + 2) = t
                numLen(k) = Len(numTxt): flagLen(k) = Len(flagTxt): holLen(k) = Len(holTxt): past(k) = (Int(d) < Int(Date))
                keys(1, 2 * k + 1) = "D:DAY|" & sv & "|": keys(1, 2 * k + 2) = "D:DAY|" & sv & "|"
                AddToRange rgDateA, ws.Cells(r, c)
                AddToRange rgDateB, ws.Cells(r, c + 1)
                If holTxt <> "" Then AddToRange rgHol, ws.Cells(r, c + 1)
                nD = SlotsFor(d, "Days", grp, slotsD, True)
                nN = SlotsFor(d, "Nights", grp, slotsN, True)
                nLines = nLines + nD + nN
                AddToRange rgDay, ws.Range(ws.Cells(r + 1, c), ws.Cells(r + mxD, c + 1))
                AddToRange rgNight, ws.Range(ws.Cells(r + mxD + 1, c), ws.Cells(r + mxD + mxN, c + 1))
                For i = 1 To mxD
                    If i <= nD Then
                        SlotIntoArrays ws, vals, keys, 1 + i, k, c, r + i, slotsD(i), sv, "Days", _
                                       rgLine, rgOpen, rgPosted, rgNcn, rgCov, rgLeft
                    Else
                        keys(1 + i, 2 * k + 1) = "E:NEW|" & sv & "|Days": keys(1 + i, 2 * k + 2) = "E:NEW|" & sv & "|Days"
                    End If
                Next i
                For i = 1 To mxN
                    If i <= nN Then
                        SlotIntoArrays ws, vals, keys, 1 + mxD + i, k, c, r + mxD + i, slotsN(i), sv, "Nights", _
                                       rgLine, rgOpen, rgPosted, rgNcn, rgCov, rgLeft
                    Else
                        keys(1 + mxD + i, 2 * k + 1) = "E:NEW|" & sv & "|Nights": keys(1 + mxD + i, 2 * k + 2) = "E:NEW|" & sv & "|Nights"
                    End If
                Next i
            Else
                AddToRange rgOut, ws.Range(ws.Cells(r + 1, c), ws.Cells(r + mxD + mxN, c + 1))
            End If
        Next k

        ' one write for the texts, one for the click map
        ws.Range(ws.Cells(r, 2), ws.Cells(r + nRows - 1, 15)).Value = vals
        mp.Range(mp.Cells(r, 2), mp.Cells(r + nRows - 1, 15)).Value = keys

        ' fills by class (the date row, the day and night areas, days outside the month), then the lines' formats
        For bgIx = 0 To 3
            If Not rgBg(bgIx) Is Nothing Then rgBg(bgIx).Interior.Color = Choose(bgIx + 1, thOutFill(), thToday(), thWeekend(), vbWhite)
        Next bgIx
        If Not rgDay Is Nothing Then rgDay.Interior.Color = thDayFill()
        If Not rgNight Is Nothing Then rgNight.Interior.Color = thNightFill()
        If Not rgOut Is Nothing Then rgOut.Interior.Color = thOutFill()
        If Not rgLine Is Nothing Then
            With rgLine
                .IndentLevel = 1: .ShrinkToFit = True: .VerticalAlignment = xlCenter
            End With
        End If
        For Each colKey In rgLeft.Keys
            With rgLeft(colKey).Font
                .Name = CAL_FONT_B: .Color = CLng(colKey)
            End With
        Next colKey
        If Not rgOpen Is Nothing Then
            With rgOpen
                .Font.Name = CAL_FONT_B: .Font.Color = thOpen(): .Interior.Color = thOpenFill()
            End With
        End If
        If Not rgPosted Is Nothing Then
            With rgPosted
                .Font.Name = CAL_FONT_B: .Font.Color = thPosted(): .Interior.Color = thPostedFill()
            End With
        End If
        If Not rgNcn Is Nothing Then
            With rgNcn
                .Font.Italic = True: .Font.Color = thMuted()
            End With
        End If
        If Not rgCov Is Nothing Then
            With rgCov
                .Font.Name = CAL_FONT_B: .Font.Color = thInk()
            End With
        End If

        ' the date cells: the formats by class, then the rich text per cell (the big number, the red open flag, the holiday)
        If Not rgDateA Is Nothing Then
            With rgDateA
                .WrapText = True: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlBottom: .IndentLevel = 1
                .Font.Name = CAL_FONT_B: .Font.Size = 8: .Font.Bold = False: .Font.Color = thDayText()
            End With
        End If
        If Not rgDateB Is Nothing Then
            With rgDateB
                .HorizontalAlignment = xlRight: .VerticalAlignment = xlBottom: .IndentLevel = 1
                .Font.Name = CAL_FONT_B: .Font.Size = 8: .Font.Color = thNightText()
            End With
        End If
        If Not rgHol Is Nothing Then rgHol.WrapText = True
        For k = 0 To 6
            d = wk + k
            If Month(d) = Month(m0) Then DateRich ws.Cells(r, 2 + 2 * k), ws.Cells(r, 3 + 2 * k), numLen(k), flagLen(k), holLen(k), past(k)
        Next k
        SectionLabel ws, r + 1, mxD, "DAYS", thDayTab(), thDayText()
        SectionLabel ws, r + 1 + mxD, mxN, "NIGHTS", thNightTab(), thNightText()

        ' borders: thick around each day, thin day/night split, dotted inside the pair (as before)
        blockTop = r
        blockBot = r + nRows - 1
        With ws.Range(ws.Cells(blockTop, 1), ws.Cells(blockTop, 15)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlMedium: .Color = thBand()
        End With
        With ws.Range(ws.Cells(blockTop, 2), ws.Cells(blockTop, 15)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = thGrid()
        End With
        With ws.Range(ws.Cells(r + mxD, 1), ws.Cells(r + mxD, 15)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = thGrid()
        End With
        For k = 0 To 6
            With ws.Range(ws.Cells(blockTop, 2 + 2 * k), ws.Cells(blockBot, 2 + 2 * k)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous: .Weight = xlThick: .Color = thBand()
            End With
            With ws.Range(ws.Cells(blockTop + 1, 3 + 2 * k), ws.Cells(blockBot, 3 + 2 * k)).Borders(xlEdgeLeft)
                .LineStyle = xlDot: .Weight = xlThin: .Color = RGB(190, 190, 190)
            End With
        Next k
        With ws.Range(ws.Cells(blockTop, 15), ws.Cells(blockBot, 15)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = thBand()
        End With

        r = r + nRows
        wk = wk + 7
    Loop
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = thBand()
    End With
End Sub

' the page setup and the window: everything under On Error Resume Next, as it was inline (the caller's
' On Error GoTo Failed comes back when this returns)
Private Sub CalDrawPage(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal r As Long)
    ' ---- page ----
    On Error Resume Next
    ' v25.2 (item 2): the same page setup every draw cost 0.6 s of 4.6. The fixed properties are set once per
    ' release (CalMap Z17 remembers); the PRINT AREA is set on every draw regardless - the wipe's Rows.Delete
    ' shrinks the Print_Area name to rows 1-2, so a draw that skipped it would print two rows (review).
    Dim pageAddr As String: pageAddr = ws.Range(ws.Cells(1, 1), ws.Cells(r, 16)).Address
    If SafeStr(mp.Range("Z17").Value) <> CAL_VERSION Or ws.PageSetup.Orientation <> xlLandscape Then   ' re-review: a re-made or hand-changed sheet too
    Err.Clear
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = pageAddr
        .PrintTitleRows = ""
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterFooter = "&7OPEN = nobody yet   POSTED = waiting for sign-ups   NO TAKERS   NCN = no cover needed   S/S = switch   N short = below minimum"     ' v27 (item 27; review M1: short - a footer holds 255 characters in all)
    End With
    Application.PrintCommunication = True
    If Err.Number = 0 Then mp.Range("Z17").Value = "'" & CAL_VERSION     ' stamped only when every property set took (re-review)
    Else
        ws.PageSetup.PrintArea = pageAddr               ' the one property that must follow the rows
    End If
    If ActiveSheet Is ws Then
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .FreezePanes = False
            .ScrollRow = 1
            .SplitColumn = 0
            .SplitRow = 4                ' title, legend, toolbar and weekdays stay put
            .FreezePanes = True
        End With
    End If
    On Error GoTo 0                                 ' v27.1: ends this step's Resume Next (DrawCalendar's GoTo Failed follows)
End Sub

' lines on a date/shift that still need someone (OPEN or no takers)
Private Function OpenCount(ByRef slots() As CalSlot, ByVal n As Long) As Long
    Dim i As Long
    For i = 1 To n
        If slots(i).IsOpen Or UCase$(slots(i).RightText) Like "NO TAKERS*" Then OpenCount = OpenCount + 1
    Next i
End Function

' a legend "chip": tinted cell, coloured bold text, coloured left edge
Private Sub LegendChip(ByVal c As Range, ByVal txt As String, ByVal col As Long, ByVal fillCol As Long)
    c.Value = txt
    c.Font.Name = CAL_FONT_B: c.Font.Size = 8.5: c.Font.Color = col
    c.Interior.Color = fillCol
    c.HorizontalAlignment = xlCenter: c.VerticalAlignment = xlCenter
    c.ShrinkToFit = True
    With c.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = col
    End With
    With c.Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = vbWhite
    End With
End Sub

' v25.2 (item 3): the rich text of one date cell pair, after the class formats went on: the big number
' (grey when past), the red "n open" flag, the holiday name top-right. The texts were written in the block.
Private Sub DateRich(ByVal a As Range, ByVal b As Range, ByVal numLen As Long, ByVal flagLen As Long, _
                     ByVal holLen As Long, ByVal past As Boolean)
    With a.Characters(1, numLen).Font
        .Size = 15
        .Color = IIf(past, RGB(158, 158, 158), thBand())
    End With
    If flagLen > 0 Then
        With a.Characters(numLen + 1, flagLen).Font
            .Size = 8
            .Color = IIf(past, RGB(158, 158, 158), thOpen())
        End With
    End If
    If holLen > 0 Then
        With b.Characters(1, holLen).Font
            .Color = IIf(past, RGB(158, 158, 158), RGB(176, 0, 32))
        End With
    End If
End Sub

' v25.2 (item 3): a range grown by Union (Nothing to start with)
Public Sub AddToRange(ByRef rg As Range, ByVal part As Range)
    If rg Is Nothing Then Set rg = part Else Set rg = Union(rg, part)
End Sub

' v26.5: a line's texts as members may see them, only while the members' copy is drawn (CalPrintPdf): a sick
' call or a personal day reads "Off" (PublicLeft, as the members' file), and a text with a leave word in it
' is not printed as written. On screen (mPublicDraw off) the texts are as drawn for supervisors.
Private Function PublicSlotLeft(ByVal lt As String, ByVal code As String) As String
    If Not mPublicDraw Then PublicSlotLeft = lt: Exit Function
    PublicSlotLeft = PublicLeft(lt, code)
    If LeaveWordsIn(PublicSlotLeft) <> "" Then PublicSlotLeft = "(see the roster)"
End Function

Private Function PublicSlotRight(ByVal rt As String) As String
    PublicSlotRight = rt
    If Not mPublicDraw Then Exit Function
    If rt <> "" And LeaveWordsIn(rt) <> "" Then PublicSlotRight = "(see the roster)"
End Function

' v25.2 (item 3): one calendar line into the block's arrays (its texts and its click-map key) and its two
' cells into the class ranges the block formats in one go. v25.5 (bundle item 0, Graeme Oct 17): no hover
' note on any line - a click shows the line's details in row 3 (CalSelectionChanged); the v25.2 numbers put
' the OPEN / POSTED boxes at 1.6 of the 2.5 s.
Private Sub SlotIntoArrays(ByVal ws As Worksheet, ByRef vals() As Variant, ByRef keys() As Variant, ByVal ar As Long, _
                           ByVal k As Long, ByVal c As Long, ByVal sheetRow As Long, ByRef s As CalSlot, ByVal sv As Long, _
                           ByVal shiftLbl As String, ByRef rgLine As Range, ByRef rgOpen As Range, ByRef rgPosted As Range, _
                           ByRef rgNcn As Range, ByRef rgCov As Range, ByVal rgLeft As Object)
    Dim a As Range: Set a = ws.Cells(sheetRow, c)
    Dim b As Range: Set b = ws.Cells(sheetRow, c + 1)
    vals(ar, 2 * k + 1) = "'" & PublicSlotLeft(s.LeftText, s.TypeCode)     ' v26.5: masked while the members' copy is drawn
    Dim t As String: t = PublicSlotRight(s.RightText)
    If t = "" Or UCase$(t) Like "NO TAKERS*" Then
        vals(ar, 2 * k + 2) = IIf(t = "", "OPEN", "NO TAKERS")
        AddToRange rgOpen, b
    ElseIf UCase$(t) Like "POSTED*" Then
        vals(ar, 2 * k + 2) = "'" & t
        AddToRange rgPosted, b
    ElseIf UCase$(t) = "NCN" Then
        vals(ar, 2 * k + 2) = "'" & t
        AddToRange rgNcn, b
    Else
        vals(ar, 2 * k + 2) = "'" & t
        AddToRange rgCov, b
    End If
    AddToRange rgLine, ws.Range(a, b)
    Dim colr As Long                                         ' v26.5: the members' copy colours Off as Off, never as sick or personal
    If mPublicDraw Then colr = TypeColor(PublicType(s.TypeCode)) Else colr = TypeColor(s.TypeCode)
    If rgLeft.Exists(colr) Then
        Set rgLeft(colr) = Union(rgLeft(colr), a)
    Else
        rgLeft.Add colr, a
    End If
    Dim keyTxt As String: keyTxt = s.Key & "|" & sv & "|" & shiftLbl
    keys(ar, 2 * k + 1) = keyTxt: keys(ar, 2 * k + 2) = keyTxt
End Sub

' DAYS / NIGHTS tab in column A for each section of a week
Private Sub SectionLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal n As Long, ByVal txt As String, _
                         ByVal fillCol As Long, ByVal fontCol As Long)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r + n - 1, 1))
        If n > 1 Then .Merge
        .Interior.Color = fillCol
        .Font.Name = CAL_FONT_B: .Font.Size = 7: .Font.Color = fontCol
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        If n >= 3 Then
            .Orientation = 90
            .Cells(1, 1).Value = txt
        ElseIf n = 2 Then
            .Orientation = 90
            .Cells(1, 1).Value = IIf(txt = "DAYS", "DAY", "NGT")
        Else
            .Orientation = 0
            .Cells(1, 1).Value = Left$(txt, 1)
        End If
    End With
End Sub

Private Sub BuildCalControls(ByVal ws As Worksheet, ByVal m0 As Date, ByVal grp As String)
    On Error Resume Next                  ' one control failing must not stop the rest (never raises)
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "cal_" Then ws.Shapes(i).Delete
    Next i

    ' row 1:  [Needs attention]  < Sep | [October 2026 v] | Nov > | This month   [unit v]   Print / PDF
    Dim nOpen As Long: nOpen = OpenAhead(grp, 14)
    CalBtn ws, "cal_adm_na", IIf(nOpen > 0, nOpen & " need cover  >", "Needs attention  >"), _
           "modNeeds.ShowNeedsAttention", ws.Range("D1:E1"), IIf(nOpen > 0, thOpen(), thNav())
    CalBtn ws, "cal_prev", "<  " & Format$(DateAdd("m", -1, m0), "mmm"), "modCalendar.CalPrev", ws.Range("F1"), thNav()
    CalDrop ws, "cal_month", ws.Range("G1:H1")
    CalBtn ws, "cal_next", Format$(DateAdd("m", 1, m0), "mmm") & "  >", "modCalendar.CalNext", ws.Range("I1"), thNav()
    CalBtn ws, "cal_today", "This month", "modCalendar.CalToday", ws.Range("J1"), thNav()
    CalDrop ws, "cal_unit", ws.Range("L1:M1")
    CalBtn ws, "cal_sup_board", "Board", "modBoard.BoardPrint", ws.Range("K1"), thGreen()     ' v26: the two-week board (cal_sup_: supervisors and admins see it)
    CalBtn ws, "cal_pdf", "Print / PDF", "modCalendar.CalPrintPdf", ws.Range("N1:O1"), thGreen()

    ' month list: January last year .. December next year
    Dim dd As Object, y0 As Long, idx As Long
    Set dd = Nothing
    Set dd = ws.DropDowns("cal_month")
    If Not dd Is Nothing Then
        y0 = Year(Date) - 1
        For i = 0 To 35
            dd.AddItem Format$(DateSerial(y0, 1 + i, 1), "mmmm yyyy")
        Next i
        idx = (Year(m0) - y0) * 12 + Month(m0)
        If idx >= 1 And idx <= 36 Then dd.ListIndex = idx
        dd.OnAction = "modCalendar.CalMonthChanged"
    End If

    ' calendar (unit) list
    Set dd = Nothing
    Set dd = ws.DropDowns("cal_unit")
    If Not dd Is Nothing Then
        Dim g As Variant: g = CalGroupNames()
        idx = 0
        For i = 0 To UBound(g)
            dd.AddItem g(i)
            If StrComp(g(i), grp, vbTextCompare) = 0 Then idx = i + 1
        Next i
        If idx = 0 Then idx = 1
        dd.ListIndex = idx
        dd.OnAction = "modCalendar.CalUnitChanged"
    End If

    ' row 3: admin toolbar (hidden for everyone else - ApplyView toggles these too)
    CalBtn ws, "cal_adm_post", "Post Line", "modCalendar.CalPostLine", ws.Range("B3"), thGreen()
    CalBtn ws, "cal_adm_set", "Set Coverage", "modCalendar.CalSetCoverage", ws.Range("C3"), RGB(230, 81, 0)
    CalBtn ws, "cal_adm_book", "Book Time Off", "modCalendar.ShowTimeOffForm", ws.Range("D3"), TypeColor("VAC")
    CalBtn ws, "cal_adm_imp", "Import Old Cal", "modCalImport.ImportUnitCalendar", ws.Range("E3"), RGB(120, 144, 156)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 8) = "cal_adm_" Then ws.Shapes(i).Visible = IIf(gAdminUnlocked, msoTrue, msoFalse)
        If Left$(ws.Shapes(i).Name, 8) = "cal_sup_" Then ws.Shapes(i).Visible = IIf(gAdminUnlocked Or IsSupervisor(), msoTrue, msoFalse)   ' v26 review
    Next i
End Sub

' lines needing someone over the next n days (for the Needs attention button)
Private Function OpenAhead(ByVal grp As String, ByVal nDays As Long) As Long
    Dim i As Long, sh As Variant, slots() As CalSlot, m As Long
    For i = 0 To nDays - 1
        For Each sh In Array("Days", "Nights")
            m = SlotsFor(Date + i, CStr(sh), grp, slots, True)
            OpenAhead = OpenAhead + OpenCount(slots, m)
        Next sh
    Next i
End Function

' v25.3 (the Desk's tiles): open lines over the next n days in EVERY unit ("" is every unit to UnitInGroup),
' and the open postings whose closing moment has passed (the Needs Attention "READY TO AWARD" test)
Public Function OpenLinesAhead(ByVal nDays As Long) As Long
    LoadCache
    LoadPostCache
    OpenLinesAhead = OpenAhead("", nDays)
End Function

' v25.5 (A8, the Desk's cover tile): the first open line over the next n days as "Tue Oct 21 Nights, Phosphine"
' ("" when none). The same walk as OpenAhead, every unit, stopped at the first open or NO TAKERS slot.
Public Function NextOpenLine(ByVal nDays As Long) As String
    Dim d As Date, sh As String, unit As String
    If NextOpenSlot(nDays, d, sh, unit) Then NextOpenLine = Format$(d, "ddd mmm d") & " " & sh & IIf(unit <> "", ", " & UnitDisplayName(unit), "")   ' v27 (item 20, B12)
End Function

' v27 (item 15): the next uncovered shift (open, or NO TAKERS) within nDays as a date, shift and unit - the Desk's
' cover tile opens the Calendar ON it (CalShowDate). The one search NextOpenLine words.
Public Function NextOpenSlot(ByVal nDays As Long, ByRef d As Date, ByRef shiftLbl As String, ByRef unit As String) As Boolean
    LoadCache
    LoadPostCache
    Dim i As Long, sh As Variant, slots() As CalSlot, m As Long, k As Long
    For i = 0 To nDays - 1
        For Each sh In Array("Days", "Nights")
            m = SlotsFor(Date + i, CStr(sh), "", slots, True)
            For k = 1 To m
                If slots(k).IsOpen Or UCase$(slots(k).RightText) Like "NO TAKERS*" Then
                    d = Date + i: shiftLbl = CStr(sh): unit = SlotUnit(slots(k).Key)
                    NextOpenSlot = True
                    Exit Function
                End If
            Next k
        Next sh
    Next i
End Function

Public Function ReadyToAwardCount() As Long
    LoadPostCache
    If Not IsArray(mPostV) Then Exit Function
    Dim i As Long, cl As Date
    For i = 1 To UBound(mPostV, 1)
        If SafeStr(mPostV(i, PS_STATUS)) = "Open" Then
            cl = AsDate(mPostV(i, PS_CLOSES))
            If cl > 0 Then
                If PostingClosed(cl) Then ReadyToAwardCount = ReadyToAwardCount + 1
            End If
        End If
    Next i
End Function

' v27 (Desk v2): the closing date of the ready-to-award posting that has waited longest (0 when none) - the Desk's
' award tile names it; the same test as ReadyToAwardCount, never a rule of its own
Public Function ReadyToAwardOldest() As Date
    LoadPostCache
    If Not IsArray(mPostV) Then Exit Function
    Dim i As Long, cl As Date
    For i = 1 To UBound(mPostV, 1)
        If SafeStr(mPostV(i, PS_STATUS)) = "Open" Then
            cl = AsDate(mPostV(i, PS_CLOSES))
            If cl > 0 Then
                If PostingClosed(cl) Then
                    If ReadyToAwardOldest = 0 Or cl < ReadyToAwardOldest Then ReadyToAwardOldest = cl
                End If
            End If
        End If
    Next i
End Function

' button centred in its cell(s), fixed height so the text never clips
Public Sub CalBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                   ByVal cel As Range, ByVal col As Long)
    Dim h As Double: h = cel.Height - 10
    If h > 24 Then h = 24
    If h < 16 Then h = 16
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, cel.Left + 4, cel.Top + (cel.Height - h) / 2, cel.Width - 8, h)
    s.Name = nm
    s.Fill.ForeColor.RGB = col
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    s.OnAction = mac
    On Error Resume Next                            ' never raises
    s.DrawingObject.PrintObject = False
    s.Adjustments.Item(1) = 0.3
    With s.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Text = cap
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.ParagraphFormat.SpaceBefore = 0
        .TextRange.ParagraphFormat.SpaceAfter = 0
        .TextRange.Font.Name = CAL_FONT_B
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoFalse
        .TextRange.Font.Fill.ForeColor.RGB = vbWhite
    End With
End Sub

' drop-down centred in its cell(s)
Public Sub CalDrop(ByVal ws As Worksheet, ByVal nm As String, ByVal cel As Range)
    Dim dd As Object
    Set dd = ws.DropDowns.Add(cel.Left + 4, cel.Top + (cel.Height - 20) / 2, cel.Width - 8, 20)
    dd.Name = nm
    dd.PrintObject = False
End Sub

' Admin: click a calendar line, then Set Coverage
Public Sub CalSetCoverage()
    On Error GoTo Oops
    If Not ActiveSheet.Name = CAL_SHEET Then MsgBox "Click a line on the Calendar first.", vbInformation, "Call Out Roster": Exit Sub
    CalSetCoverageKey SafeStr(MapSheet().Range(ActiveCell.Address).Value)
    Exit Sub
Oops:
    ReportError "CalSetCoverage"
End Sub

Public Sub CalSetCoverageKey(ByVal keyTxt As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Left$(keyTxt, 2) <> "A:" And Left$(keyTxt, 2) <> "C:" Then MsgBox "Click a line on the calendar (who's off / the need), then Set Coverage." & vbNewLine & "For an empty spot, double-click it instead.", vbInformation, "Call Out Roster": Exit Sub
    Dim parts As Variant: parts = Split(keyTxt, "|")
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim shiftLbl As String: shiftLbl = parts(2)
    Dim leftTxt As String, rightTxt As String
    If Not SlotTexts(keyTxt, leftTxt, rightTxt) Then MsgBox "That line has changed - the calendar will redraw.", vbExclamation, "Call Out Roster": DrawCalendar: Exit Sub

    Dim ans As String
    ans = Trim$(VBA.InputBox("Who covers  " & Clip(leftTxt, 40) & "  on " & Format$(d, "ddd mmm d") & " " & shiftLbl & "?" & vbNewLine & vbNewLine & _
                             "Type a name (e.g. Joe D), NCN, or" & vbNewLine & _
                             "CLEAR to remove the coverage" & vbNewLine & _
                             "DELETE to remove this line" & IIf(Left$(slotKey, 2) = "A:", " (time off: void it in Book Time Off instead)", ""), _
                             "Set coverage"))
    If ans = "" Then Exit Sub

    Dim lo As ListObject: Set lo = CovTable(True)
    Dim unit As String
    LoadCache
    If Left$(slotKey, 2) = "A:" Then
        Dim ar As Long: ar = RowOf(mAbs, Mid$(slotKey, 3))
        If ar > 0 Then unit = SafeStr(mAbs(ar, AB_UNIT))
        If UCase$(ans) = "DELETE" Then MsgBox "Void the time off in Book Time Off instead.", vbInformation, "Call Out Roster": Exit Sub
        ' replace existing manual coverage for this absence / date / shift
        Dim v As Variant: v = mCov
        Dim i As Long, srcs As String
        If IsArray(v) Then
            For i = 1 To UBound(v, 1)
                If CovLive(i, d, shiftLbl) And StrComp(SafeStr(v(i, CV_ABS)), Mid$(slotKey, 3), vbTextCompare) = 0 Then
                    If SafeStr(v(i, CV_SOURCE)) <> "Manual" And SafeStr(v(i, CV_SOURCE)) <> "" Then srcs = srcs & " " & SafeStr(v(i, CV_SOURCE))
                End If
            Next i
            If srcs <> "" Then
                If MsgBox("This line is filled by" & srcs & ". Replace it anyway?", vbYesNo + vbExclamation, "Call Out Roster") = vbNo Then Exit Sub
            End If
            UnlockSheet lo.Parent
            For i = 1 To UBound(v, 1)
                If CovLive(i, d, shiftLbl) And StrComp(SafeStr(v(i, CV_ABS)), Mid$(slotKey, 3), vbTextCompare) = 0 Then
                    lo.DataBodyRange.Cells(i, CV_STATUS).Value = "VOID"
                End If
            Next i
            ProtectSheet lo.Parent
        End If
        If UCase$(ans) <> "CLEAR" Then AddCoverage d, shiftLbl, unit, slotKey, ans, Empty, "Manual", "", ""
    Else
        Dim cr As Long: cr = RowOf(mCov, Mid$(slotKey, 3))
        If cr = 0 Then Exit Sub
        UnlockSheet lo.Parent
        Select Case UCase$(ans)
            Case "DELETE": lo.DataBodyRange.Cells(cr, CV_STATUS).Value = "VOID"
            Case "CLEAR": lo.DataBodyRange.Cells(cr, CV_BY).ClearContents
            Case Else: lo.DataBodyRange.Cells(cr, CV_BY).Value = ans
        End Select
        ProtectSheet lo.Parent
    End If
    SaveIfWritable
    DrawCalendar
End Sub

' v24.3.2 (Graeme's Oct 13 test, C3): change the words on a need line. Only a "C:" line has free
' text of its own: a time-off line shows the name, the kind of time off and the times note from
' the booking. The new text goes through the publish's own leave-word screen (LeaveWordsIn), so a
' word that would stop Member View is refused here, with the word named. When the line came from
' a posting that is still Open, the posting's Reason changes with it (the members' file carries
' both); a closed posting's Reason is no longer published and is left alone.
Public Sub CalEditTextKey(ByVal keyTxt As String)
    Const TTL As String = "Change the text"
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Left$(keyTxt, 2) <> "C:" Then
        MsgBox "Only a need line has text of its own. A time-off line shows the name, the kind of time off and the " & _
               "times note from the booking: void it in Book Time Off and book it again.", vbInformation, TTL
        Exit Sub
    End If
    Dim parts As Variant: parts = Split(keyTxt, "|")
    Dim covId As String: covId = Mid$(CStr(parts(0)), 3)
    Dim d As Date: d = KeyDate(keyTxt)
    Dim leftTxt As String, rightTxt As String
    If Not SlotTexts(keyTxt, leftTxt, rightTxt) Then MsgBox "That line has changed - the calendar will redraw.", vbExclamation, TTL: DrawCalendar: Exit Sub
    Dim src As String: src = CovCellOf(covId, CV_SOURCE)
    Dim prompt As String
    prompt = "New text for the line  " & Clip(leftTxt, 40) & "  on " & Format$(d, "ddd mmm d") & " " & KeyShift(keyTxt) & "." & vbNewLine & vbNewLine & _
             "Members see this text on their calendar: never the kind of anyone's leave."
    Dim stat As String
    If UCase$(Left$(src, 2)) = "P-" Then
        stat = PostingStatus(src)
        If stat = "Open" Then
            If StrComp(Trim$(leftTxt), PostingReasonOf(src), vbTextCompare) = 0 Then
                prompt = prompt & vbNewLine & vbNewLine & "Posting " & src & " is open and these are its Reason's words: its Reason changes with it."
            Else
                prompt = prompt & vbNewLine & vbNewLine & "Posting " & src & " is open; its Reason has other words and stays (Change the Reason of posting changes it)."
            End If
        ElseIf stat <> "" Then
            prompt = prompt & vbNewLine & vbNewLine & "Posting " & src & " is " & LCase$(stat) & ": its own Reason stays as it was."
        End If
    End If
    Dim ans As String
    ans = Trim$(VBA.InputBox(prompt, TTL, leftTxt))
    If ans = "" Then Exit Sub
    If StrComp(ans, Trim$(leftTxt), vbBinaryCompare) = 0 Then Exit Sub
    Dim rep As String
    If Not SetLineText(covId, ans, rep) Then
        MsgBox rep, vbExclamation, TTL
        Exit Sub
    End If
    SaveIfWritable
    DrawCalendar
    MsgBox rep, vbInformation, TTL
End Sub

' v24.3.2: the work behind "Change the text of this line" (and SmokeTest step 9): with RenameEverywhere
' the only write to a line's need text after it is made. Writes the need
' text of coverage line covId and, when its source is a posting that is still Open, that posting's
' Reason. Refused, writing nothing, when the text is blank, holds a leave word (LeaveWordsIn: the
' publish would refuse it) or the line is gone. report says what was done, or why not.
Public Function SetLineText(ByVal covId As String, ByVal newText As String, ByRef report As String) As Boolean
    newText = Trim$(newText)
    report = TextRefusal(newText)
    If report <> "" Then Exit Function
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then report = "There is no Coverage table.": Exit Function
    Dim v As Variant: v = TableData(lo)
    Dim cr As Long: cr = RowOf(v, covId)
    If cr = 0 Then report = "Calendar line " & covId & " wasn't found: close the Calendar and open it again.": Exit Function
    If UCase$(SafeStr(v(cr, CV_STATUS))) = "VOID" Then report = "Calendar line " & covId & " has been removed.": Exit Function
    Dim src As String: src = Trim$(SafeStr(v(cr, CV_SOURCE)))
    Dim oldTxt As String: oldTxt = Trim$(SafeStr(v(cr, CV_NEED)))
    WriteNeedText lo, cr, newText
    MarkCalDirty
    MarkMembersDirty
    report = "The line now reads '" & newText & "'."
    If UCase$(Left$(src, 2)) = "P-" Then
        Dim stat As String: stat = PostingStatus(src)
        If stat = "Open" Then
            ' the posting's Reason follows only when the line's words WERE its Reason (a posting that
            ' filled an existing need line kept that line's own words; its Reason is changed on its own)
            If StrComp(oldTxt, PostingReasonOf(src), vbTextCompare) = 0 Then
                If SetPostingReason(src, newText) Then
                    report = report & " Posting " & src & "'s Reason too: print its sheet again if it is on the board."
                End If
            Else
                report = report & " Posting " & src & "'s own Reason ('" & PostingReasonOf(src) & "') has other words and was left: " & _
                         "Change the Reason of posting changes it."
            End If
        ElseIf stat <> "" Then
            report = report & " Posting " & src & " is " & LCase$(stat) & ": its own Reason was left as it was."
        End If
    End If
    SetLineText = True
End Function

' v24.3.2: why a text may not go on the calendar or a posting ("" = it may). The same two screens
' the publish runs on every free text (LeaveWordsIn, PrivateValueKind), so what is accepted here
' publishes.
Private Function TextRefusal(ByVal txt As String) As String
    If txt = "" Then TextRefusal = "The text can't be blank.": Exit Function
    Dim w As String: w = LeaveWordsIn(txt)
    If w <> "" Then
        TextRefusal = "Not changed: '" & w & "' would tell members the kind of someone's leave, and Member View would refuse to publish it. Use other words."
        Exit Function
    End If
    Select Case PrivateValueKind(txt)
        Case "email": TextRefusal = "Not changed: that looks like an email address, and Member View would refuse to publish it."
        Case "phone": TextRefusal = "Not changed: that holds a run of digits like a phone number or an ID, and Member View would refuse to publish it. " & _
                                    "A time range is written 0700-1900, never 700-1900."
    End Select
End Function

' v24.3.2: the protected write of a coverage line's need text. As text, whatever it starts with
' (an = + - or a date would be converted). Re-protects and re-raises if the write fails.
Private Sub WriteNeedText(ByVal lo As ListObject, ByVal rowIx As Long, ByVal txt As String)
    Dim en As Long, em As String
    On Error GoTo Failed
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(rowIx, CV_NEED).Value = "'" & txt
    ProtectSheet lo.Parent
    Exit Sub
Failed:
    en = Err.Number: em = Err.Description
    On Error Resume Next
    ProtectSheet lo.Parent
    On Error GoTo 0                                     ' (re-review) or the raise below would be swallowed by Resume Next
    Err.Raise en, , em
End Sub

' v24.3.2 (review): "Change the Reason of posting" on a line or a Needs Attention entry. For a
' posting whose words are not on a need line of its own (placed on a time-off line, filled an
' existing line, or not on the calendar), where Change the text of this line can't reach them.
Public Sub CalEditReasonId(ByVal id As String)
    Const TTL As String = "Change the Reason"
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim stat As String: stat = PostingStatus(id)
    If stat = "" Then MsgBox "Posting " & id & " isn't on the Postings sheet.", vbExclamation, TTL: Exit Sub
    If stat <> "Open" Then
        MsgBox "Posting " & id & " is " & LCase$(stat) & ". Only an open posting's Reason is changed: it is the one members see.", vbInformation, TTL
        Exit Sub
    End If
    Dim oldTxt As String: oldTxt = PostingReasonOf(id)
    Dim ans As String
    ans = Trim$(VBA.InputBox("New Reason for posting " & id & " (now '" & oldTxt & "')." & vbNewLine & vbNewLine & _
                             "Members see it on the posting and on their calendar: never the kind of anyone's leave." & vbNewLine & _
                             "A calendar line of this posting that carries these same words changes with it.", TTL, oldTxt))
    If ans = "" Then Exit Sub
    If StrComp(ans, oldTxt, vbBinaryCompare) = 0 Then Exit Sub
    Dim rep As String
    If Not SetReasonText(id, ans, rep) Then
        MsgBox rep, vbExclamation, TTL
        Exit Sub
    End If
    SaveIfWritable
    RefreshCalendar
    MsgBox rep, vbInformation, TTL
End Sub

' v24.3.2: the work behind "Change the Reason of posting" (and SmokeTest step 9). Screened as
' SetLineText; only an Open posting. Every live calendar line the posting wrote whose text IS the
' old Reason changes with it; a line with other words is left. (A write that fails part-way
' raises to the caller's handler with the Reason changed and that line not: the box says so.)
Public Function SetReasonText(ByVal id As String, ByVal newText As String, ByRef report As String) As Boolean
    newText = Trim$(newText)
    report = TextRefusal(newText)
    If report <> "" Then Exit Function
    Dim stat As String: stat = PostingStatus(id)
    If stat = "" Then report = "Posting " & id & " isn't on the Postings sheet.": Exit Function
    If stat <> "Open" Then report = "Posting " & id & " is " & LCase$(stat) & ": only an open posting's Reason is changed.": Exit Function
    Dim oldTxt As String: oldTxt = PostingReasonOf(id)
    If Not SetPostingReason(id, newText) Then report = "Posting " & id & "'s row couldn't be found.": Exit Function
    Dim lo As ListObject: Set lo = CovTable()
    Dim nLines As Long, i As Long, v As Variant
    If Not lo Is Nothing Then
        v = TableData(lo)
        If IsArray(v) Then
            For i = 1 To UBound(v, 1)
                If StrComp(Trim$(SafeStr(v(i, CV_SOURCE))), Trim$(id), vbTextCompare) = 0 Then
                    If UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
                        If StrComp(Trim$(SafeStr(v(i, CV_NEED))), oldTxt, vbTextCompare) = 0 Then
                            WriteNeedText lo, i, newText
                            nLines = nLines + 1
                        End If
                    End If
                End If
            Next i
        End If
    End If
    MarkCalDirty
    MarkMembersDirty
    report = "Posting " & id & "'s Reason now reads '" & newText & "'" & _
             IIf(nLines = 1, ", and so does its calendar line", IIf(nLines > 1, ", and so do its " & nLines & " calendar lines", "")) & _
             ". Print its sheet again if it is on the board."
    SetReasonText = True
End Function

' v24.3.2: one cell of one coverage line, as text ("" when the line isn't there)
Public Function CovCellOf(ByVal covId As String, ByVal col As Long) As String
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then Exit Function
    Dim v As Variant: v = TableData(lo)
    Dim cr As Long: cr = RowOf(v, covId)
    If cr > 0 Then CovCellOf = Trim$(SafeStr(v(cr, col)))
End Function

' v24.3.2: the first live coverage line that came from source (a P-#### posting, an SS-#### switch)
Public Function CovIdOfSource(ByVal source As String) As String
    Dim lo As ListObject: Set lo = CovTable()
    If lo Is Nothing Then Exit Function
    Dim v As Variant: v = TableData(lo)
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, CV_SOURCE))), Trim$(source), vbTextCompare) = 0 Then
            If UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
                CovIdOfSource = Trim$(SafeStr(v(i, CV_ID)))
                Exit Function
            End If
        End If
    Next i
End Function

' Admin: click an open line, then Post Line -> New Posting, pre-filled
Public Sub CalPostLine()
    On Error GoTo Oops
    If ActiveSheet.Name <> CAL_SHEET Then MsgBox "Click a line on the Calendar first.", vbInformation, "Call Out Roster": Exit Sub
    CalPostKey SafeStr(MapSheet().Range(ActiveCell.Address).Value)
    Exit Sub
Oops:
    ReportError "CalPostLine"
End Sub

Public Sub CalPostKey(ByVal keyTxt As String)
    If Not RequireSupervisor() Then Exit Sub            ' v27 (item 15): posting a shift is a supervisor's job (v25.5); was RequireAdmin
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Left$(keyTxt, 2) <> "A:" And Left$(keyTxt, 2) <> "C:" Then MsgBox "Click the line that needs coverage (who's off / the need), then Post Line." & vbNewLine & "For an empty spot, double-click it to post support.", vbInformation, "Call Out Roster": Exit Sub

    Dim parts As Variant: parts = Split(keyTxt, "|")
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim shiftLbl As String: shiftLbl = parts(2)
    Dim leftTxt As String, rightTxt As String     ' from the data, not the sheet (the sheet shows "OPEN")
    If Not SlotTexts(keyTxt, leftTxt, rightTxt) Then MsgBox "That line has changed - the calendar will redraw.", vbExclamation, "Call Out Roster": DrawCalendar: Exit Sub
    rightTxt = Trim$(rightTxt)

    If UCase$(rightTxt) Like "POSTED*" Then
        MsgBox "That line is already " & rightTxt & ".", vbInformation, "Call Out Roster": Exit Sub
    End If
    If rightTxt <> "" And rightTxt <> "-" Then
        If MsgBox(leftTxt & " is already covered by " & rightTxt & "." & vbNewLine & "Post it anyway?", _
                  vbYesNo + vbQuestion, "Post Line") = vbNo Then Exit Sub
    End If
    If Int(d) < Int(Date) Then
        If MsgBox("That date has passed. Post it anyway?", vbYesNo + vbQuestion, "Post Line") = vbNo Then Exit Sub
    End If

    ' unit / type / hours from the record behind the line
    LoadCache
    Dim unit As String, code As String, hrs As Double, r As Long
    hrs = FULL_SHIFT
    If Left$(slotKey, 2) = "A:" Then
        r = RowOf(mAbs, Mid$(slotKey, 3))
        If r > 0 Then
            unit = SafeStr(mAbs(r, AB_UNIT)): code = TypeCode(SafeStr(mAbs(r, AB_TYPE)))
            If IsNumeric(mAbs(r, AB_HOURS)) And Not IsEmpty(mAbs(r, AB_HOURS)) Then
                If CDbl(mAbs(r, AB_HOURS)) > 0 Then hrs = CDbl(mAbs(r, AB_HOURS))
            End If
        End If
    Else
        r = RowOf(mCov, Mid$(slotKey, 3))
        If r > 0 Then
            unit = SafeStr(mCov(r, CV_UNIT)): code = UCase$(SafeStr(mCov(r, CV_TYPE)))
            If IsNumeric(mCov(r, CV_HOURS)) And Not IsEmpty(mCov(r, CV_HOURS)) Then
                If CDbl(mCov(r, CV_HOURS)) > 0 Then hrs = CDbl(mCov(r, CV_HOURS))
            End If
        End If
    End If
    If r = 0 Then MsgBox "Couldn't find the record behind that line - redraw the calendar and try again.", vbExclamation, "Call Out Roster": Exit Sub

    Dim reason As String
    Select Case code
        Case "VAC": reason = "Vacation Coverage"
        Case "PHA": reason = "PHA"
        Case "TRN": reason = "Training Coverage"
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": reason = "Leave Coverage"
        Case "SUP": reason = "Unit Operator Support"
        Case Else: reason = "Other"
    End Select

    gCalPostKey = slotKey
    gCalPostDate = d
    gCalPostShift = shiftLbl
    frmNewPosting.Prefill unit, d, shiftLbl, reason, hrs, Format$(d, "ddd mmm d") & " " & shiftLbl & ": " & leftTxt
    frmNewPosting.Show
    gCalPostKey = ""                          ' cancelled or done - don't reuse
    RefreshCalendar
End Sub

Public Sub CalPrintPdf()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = CalendarSheet()
    Dim ans As VbMsgBoxResult
    ' v26.5.5 (B10): a question Yes / No answer, not invented button meanings (Cancel = neither)
    ans = MsgBox("The month is saved as a PDF (the members' copy: a sick call or a personal day reads Off)." & _
                 vbNewLine & vbNewLine & "Print it on paper as well?", vbYesNoCancel + vbQuestion, "Calendar")
    If ans = vbCancel Then Exit Sub
    Dim folder As String, subF As String
    folder = EnsureSlash(ReportFolder())
    If FolderExists(folder) Then
        subF = folder & "Calendars" & Application.PathSeparator
        If Not FolderExists(subF) Then
            On Error Resume Next
            MkDir subF
            On Error GoTo Oops
        End If
        If FolderExists(subF) Then folder = subF
    Else
        folder = EnsureSlash(ThisWorkbook.Path)
    End If
    Dim fName As String: fName = "Calendar " & Replace(Replace(CalGroup(), "/", "-"), "&", "and") & " " & Format$(CalMonth(), "yyyy-mm") & ".pdf"
    Dim msg As String
    ' v26.5: the printed month is the MEMBERS' copy - a sick call or a personal day reads "Off", a text with a leave
    ' word is not printed as written (decided Oct 18 after the v26 reviews; the board and the members' file do the
    ' same). Drawn masked, exported / printed, then drawn again as supervisors see it.
    mPublicDraw = True
    DrawCalendar
    ' the masked draw must have finished (DrawCalendar reports its own failure and never raises): otherwise the
    ' sheet may still hold the supervisors' draw - nothing is exported
    If InStr(SafeStr(ws.Range("B1").Value), "members' copy") = 0 Or Left$(SafeStr(ws.Range("F3").Value), 8) = "Calendar" Then
        mPublicDraw = False
        DrawCalendar
        MsgBox "The calendar couldn't be drawn for printing, so nothing was printed or saved.", vbExclamation, "Calendar"
        Exit Sub
    End If
    On Error Resume Next
    ' the admin toolbar and the needs-attention box are screen-only
    Dim hideTo As Long: hideTo = CLng(Val(MapSheet().Range("Z4").Value))
    If hideTo < 3 Then hideTo = 3
    UnlockSheet ws
    ws.Rows("3:" & hideTo).Hidden = True
    Dim ftWas As String: ftWas = ws.PageSetup.LeftFooter        ' v26.5.5 (5.2): the trust strip in the PDF's / print's footer
    Dim ctWas As String: ctWas = ws.PageSetup.CenterFooter      ' (review M1) the legend steps aside: the two together pass 255 characters
    ws.PageSetup.CenterFooter = ""
    ws.PageSetup.LeftFooter = "&8" & TrustStrip()
    Err.Clear
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & fName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then msg = "PDF NOT saved: " & Err.Description Else msg = "PDF saved: " & folder & fName
    Err.Clear
    If ans = vbYes Then
        ws.PrintOut Copies:=1
        If Err.Number <> 0 Then msg = msg & vbNewLine & "NOT printed: " & Err.Description Else msg = msg & vbNewLine & "Sent to printer."
    End If
    ws.PageSetup.LeftFooter = ftWas
    ws.PageSetup.CenterFooter = ctWas
    ws.Rows("3:" & hideTo).Hidden = False
    ProtectSheet ws
    On Error GoTo Oops
    mPublicDraw = False
    DrawCalendar
    MsgBox msg & vbNewLine & vbNewLine & "The printed copy is the members' copy: a sick call or a personal day reads Off.", vbInformation, "Calendar"
    Exit Sub
Oops:
    mPublicDraw = False                         ' never leave the masked draw as the working calendar
    ReportError "CalPrintPdf"
    DrawCalendar                                ' back to the supervisors' draw (it never raises)
End Sub

Public Sub ShowTimeOffForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If AbsTable(True) Is Nothing Then Exit Sub
    If CovTable(True) Is Nothing Then Exit Sub
    frmTimeOff.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "ShowTimeOffForm"
End Sub

' ====================================================================
'  LINE DETAILS  (a click's details in row 3 + Details in the menu)
' ====================================================================
Public Sub LoadPostCache()
    mPostV = Empty
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(PT_TABLE)
        If Not lo Is Nothing Then
            mPostV = TableData(lo)
            Exit Sub
        End If
    Next ws
End Sub

Public Function PostRow(ByVal id As String) As Long
    If IsArray(mPostV) Then PostRow = RowOf(mPostV, id)
End Function

Public Function AsDate(ByVal x As Variant) As Date
    If IsDate(x) Then AsDate = CDate(x)
End Function

Public Function TypeLong(ByVal code As String) As String      ' v27 (item 25): Public - the handover page shows leave as "Leave" the same way
    Select Case UCase$(code)
        Case "PER": TypeLong = "Personal day"
        Case "VAC": TypeLong = "Vacation"
        Case "TRN": TypeLong = "Training"
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": TypeLong = "Leave"
        Case "SIC": TypeLong = "Sick call"
        Case "PHA": TypeLong = "PHA"
        Case "SUP": TypeLong = "Unit operator support"
        Case "SS": TypeLong = "Shift switch"
        Case "OTH": TypeLong = "Other"
        Case Else: TypeLong = code
    End Select
End Function

Public Function StampTxt(ByVal who As Variant, ByVal stamp As Variant) As String
    StampTxt = Trim$(SafeStr(who))
    If IsDate(stamp) Then StampTxt = StampTxt & " on " & Format$(stamp, "mmm d, h:mm AM/PM")
End Function

' one line about a posting, from the cached tblPostings
Private Function PostingText(ByVal id As String) As String
    Dim r As Long: r = PostRow(id)
    If r = 0 Then PostingText = "Posting " & id: Exit Function
    Dim st As String: st = SafeStr(mPostV(r, PS_STATUS))
    Dim cl As Date: cl = AsDate(mPostV(r, PS_CLOSES))
    Dim t As String
    If st = "Open" Then
        If cl = 0 Then
            t = "Posting " & id & ": open (no closing date)"
        ElseIf PostingClosed(cl) Then                   ' v25 review: the one closing test
            t = "Posting " & id & ": closed - READY TO AWARD"
        ElseIf Int(cl) = Date Then
            t = "Posting " & id & ": closes TODAY " & POSTING_CLOSE_TIME
        Else
            t = "Posting " & id & ": open until " & Format$(cl, "ddd mmm d") & " " & POSTING_CLOSE_TIME
        End If
        t = t & vbLf & "   " & SafeStr(mPostV(r, PS_HOURS)) & "h, " & SafeStr(mPostV(r, PS_POS)) & " position(s), " & SafeStr(mPostV(r, PS_REASON))
    ElseIf SafeStr(mPostV(r, PS_AWARDTO)) <> "" Then
        t = "Posting " & id & ": awarded to " & SafeStr(mPostV(r, PS_AWARDTO))
        If IsDate(mPostV(r, PS_AWARDON)) Then t = t & " (" & Format$(mPostV(r, PS_AWARDON), "mmm d") & ")"
    Else
        t = "Posting " & id & ": " & st
    End If
    PostingText = t
End Function

Private Function SourceText(ByVal src As String) As String
    If src = "" Then Exit Function
    If UCase$(Left$(src, 2)) = "P-" Then
        SourceText = PostingText(src)
    ElseIf UCase$(Left$(src, 3)) = "SS-" Then
        SourceText = "Shift switch " & src
    ElseIf StrComp(src, "Manual", vbTextCompare) = 0 Then
        SourceText = "Set by hand"
    ElseIf UCase$(src) Like "IMPORT*" Then
        SourceText = "From the old calendar (" & src & ")"
    Else
        SourceText = "Call-out " & src
    End If
End Function

' who covers (one coverage row) + where it came from
Private Function CovDetails(ByVal c As Long) As String
    Dim by As String: by = SafeStr(mCov(c, CV_BY))
    Dim t As String
    If by <> "" Then
        If UCase$(by) Like "POSTED*" Then
            t = vbLf & "Posted - waiting for sign-ups"
        ElseIf UCase$(by) = "NCN" Then
            t = vbLf & "No coverage needed (NCN)"
        ElseIf UCase$(by) Like "NO TAKERS*" Then
            t = vbLf & "No takers - still needs cover"
        Else
            t = vbLf & "Covered by: " & by & HrsTag(mCov(c, CV_HOURS))
        End If
    End If
    Dim src As String: src = SourceText(SafeStr(mCov(c, CV_SOURCE)))
    If src <> "" Then t = t & vbLf & src
    CovDetails = t
End Function

' Everything known about one calendar line. keyTxt = "A:AB-0001|serial|Days".
' Leave types only ever show as "Leave"; their notes are never shown.
Public Function LineDetails(ByVal keyTxt As String, Optional ByVal useCache As Boolean = False) As String
    On Error GoTo Quit
    If Not useCache Then LoadCache: LoadPostCache
    Dim parts As Variant: parts = Split(keyTxt, "|")
    If UBound(parts) < 2 Then Exit Function
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim shiftLbl As String: shiftLbl = parts(2)
    Dim id As String: id = Mid$(slotKey, 3)
    Dim s As String, r As Long, c As Long, cov As String, code As String
    s = Format$(d, "dddd mmm d") & " - " & shiftLbl

    If Left$(slotKey, 2) = "D:" Then                    ' v26: the date cell - the crews on, and the short shifts behind its flag
        s = Format$(d, "dddd mmm d") & "  -  Days " & DaysCrew(d) & ", Nights " & NightsCrew(d)
        If Int(d) >= Int(Date) And Int(d) < Int(Date) + BOARD_DAYS Then
            Dim shortTxt As String: shortTxt = StaffShortText(d, "")          ' v26.5: every unit, whatever group is shown
            If shortTxt <> "" Then s = s & vbLf & "BELOW MINIMUM STAFFING  -  " & shortTxt
        End If
        LineDetails = s
        Exit Function
    End If

    If Left$(slotKey, 2) = "A:" Then
        r = RowOf(mAbs, id)
        If r = 0 Then LineDetails = s: Exit Function
        code = TypeCode(SafeStr(mAbs(r, AB_TYPE)))
        s = s & vbLf & CalName(SafeStr(mAbs(r, AB_EMP))) & " is off: " & TypeLong(code)
        s = s & vbLf & "From " & Format$(mAbs(r, AB_START), "mmm d") & " to " & _
            IIf(IsDate(mAbs(r, AB_END)), Format$(mAbs(r, AB_END), "mmm d"), "open-ended")
        If HrsTag(mAbs(r, AB_HOURS)) <> "" Then
            s = s & vbLf & "Off" & HrsTag(mAbs(r, AB_HOURS)) & " of the shift"
            If SafeStr(mAbs(r, AB_TIMES)) <> "" Then s = s & " (" & SafeStr(mAbs(r, AB_TIMES)) & ")"
        End If
        If IsArray(mCov) Then
            For c = 1 To UBound(mCov, 1)
                If CovLive(c, d, shiftLbl) Then
                    If StrComp(SafeStr(mCov(c, CV_ABS)), id, vbTextCompare) = 0 Then cov = cov & CovDetails(c)
                End If
            Next c
        End If
        If cov = "" Then cov = vbLf & "OPEN - nobody covering yet"
        s = s & cov
        If Not IsLeaveCode(code) And code <> "SIC" And SafeStr(mAbs(r, AB_NOTE)) <> "" Then
            s = s & vbLf & "Note: " & SafeStr(mAbs(r, AB_NOTE))
        End If
        s = s & vbLf & "Booked by " & StampTxt(mAbs(r, AB_USER), mAbs(r, AB_STAMP)) & "  [" & id & "]"
    Else
        r = RowOf(mCov, id)
        If r = 0 Then LineDetails = s: Exit Function
        code = UCase$(SafeStr(mCov(r, CV_TYPE)))
        s = s & vbLf & SafeStr(mCov(r, CV_NEED)) & "  (" & TypeLong(code) & ")"
        If SafeStr(mCov(r, CV_BY)) = "" Then s = s & vbLf & "OPEN - nobody covering yet"
        s = s & CovDetails(r)
        If SafeStr(mCov(r, CV_NOTE)) <> "" Then s = s & vbLf & "Note: " & SafeStr(mCov(r, CV_NOTE))
        s = s & vbLf & "Added by " & StampTxt(mCov(r, CV_USER), mCov(r, CV_STAMP)) & "  [" & id & "]"
    End If
    LineDetails = s
    Exit Function
Quit:
    LineDetails = s
End Function

Public Function PostingDetails(ByVal id As String) As String
    LoadPostCache
    Dim r As Long: r = PostRow(id)
    If r = 0 Then PostingDetails = "Posting " & id & " wasn't found.": Exit Function
    PostingDetails = SafeStr(mPostV(r, PS_UNIT)) & " - " & Format$(AsDate(mPostV(r, PS_DATE)), "dddd mmm d") & " " & _
                     SafeStr(mPostV(r, PS_SHIFT)) & "  (crew " & SafeStr(mPostV(r, PS_CREW)) & ")" & vbLf & _
                     PostingText(id) & vbLf & _
                     "Posted " & IIf(IsDate(mPostV(r, PS_POSTED)), Format$(mPostV(r, PS_POSTED), "mmm d"), "") & _
                     " by " & SafeStr(mPostV(r, PS_POSTEDBY))
End Function

' (v25.5: AddNote, the hover note on a cell, is gone - bundle items 0 and 1; a click shows the details in row 3)

' the line's left / right text, straight from the data
Public Function SlotTexts(ByVal keyTxt As String, ByRef lt As String, ByRef rt As String) As Boolean
    Dim parts As Variant: parts = Split(keyTxt, "|")
    If UBound(parts) < 2 Then Exit Function
    Dim slots() As CalSlot, n As Long, i As Long
    n = SlotsFor(CDate(CLng(parts(1))), CStr(parts(2)), "", slots)
    For i = 1 To n
        If StrComp(slots(i).Key, CStr(parts(0)), vbTextCompare) = 0 Then
            lt = slots(i).LeftText
            rt = slots(i).RightText
            SlotTexts = True
            Exit Function
        End If
    Next i
End Function

' P-#### behind a line (if a posting covers it)
Public Function PostIdOfLine(ByVal keyTxt As String) As String
    Dim parts As Variant: parts = Split(keyTxt, "|")
    If UBound(parts) < 2 Then Exit Function
    If Not IsArray(mCov) Then Exit Function
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim c As Long, src As String, hit As Boolean
    For c = 1 To UBound(mCov, 1)
        src = SafeStr(mCov(c, CV_SOURCE))
        If UCase$(Left$(src, 2)) = "P-" Then
            hit = False
            If Left$(slotKey, 2) = "C:" Then
                hit = (StrComp(SafeStr(mCov(c, CV_ID)), Mid$(slotKey, 3), vbTextCompare) = 0)
            ElseIf CovLive(c, d, CStr(parts(2))) Then
                hit = (StrComp(SafeStr(mCov(c, CV_ABS)), Mid$(slotKey, 3), vbTextCompare) = 0)
            End If
            If hit And UCase$(SafeStr(mCov(c, CV_STATUS))) <> "VOID" Then PostIdOfLine = src
        End If
    Next c
End Function

' unit behind a calendar line key
Public Function SlotUnit(ByVal keyTxt As String) As String
    Dim slotKey As String: slotKey = Split(keyTxt, "|")(0)
    Dim r As Long
    If Left$(slotKey, 2) = "A:" Then
        r = RowOf(mAbs, Mid$(slotKey, 3))
        If r > 0 Then SlotUnit = SafeStr(mAbs(r, AB_UNIT))
    ElseIf Left$(slotKey, 2) = "C:" Then
        r = RowOf(mCov, Mid$(slotKey, 3))
        If r > 0 Then SlotUnit = SafeStr(mCov(r, CV_UNIT))
    End If
End Function
' a live calendar line came from this source (uses the cached Coverage table)
Public Function CovHasSourceCached(ByVal source As String) As Boolean
    If Not IsArray(mCov) Then Exit Function
    Dim c As Long
    For c = 1 To UBound(mCov, 1)
        If StrComp(SafeStr(mCov(c, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(mCov(c, CV_STATUS))) <> "VOID" Then
            CovHasSourceCached = True
            Exit Function
        End If
    Next c
End Function
' a Variant that may be a date serial or a date -> Date (0 if neither)
Public Function Dt2(ByVal x As Variant) As Date
    If IsDate(x) Then
        Dt2 = CDate(x)
    ElseIf IsNumeric(x) And Not IsEmpty(x) Then
        Dt2 = CDate(CDbl(x))
    End If
End Function

Public Function KeyDate(ByVal keyTxt As String) As Date
    Dim p As Variant: p = Split(keyTxt, "|")
    If UBound(p) >= 1 Then
        If IsNumeric(p(1)) Then KeyDate = CDate(CLng(p(1)))
    End If
End Function

Public Function KeyShift(ByVal keyTxt As String) As String
    Dim p As Variant: p = Split(keyTxt, "|")
    If UBound(p) >= 2 Then KeyShift = CStr(p(2))
End Function
' ====================================================================
'  AUTOMATIC REDRAW + EVENT ROUTING  (called from ThisWorkbook)
' ====================================================================
' Opening the workbook: jump to this month and redraw.
' Opening the workbook: redraw only if the calendar is out of date
' (new day, something changed, or it was left on another month).
Public Sub CalOnOpen()
    mEventsSeen = True
    On Error Resume Next                            ' never raises
    If Not SheetExists(CAL_SHEET) Then Exit Sub
    CalendarSheet
    If gAdminUnlocked Then MinStaffTable True       ' v26: the Min Staff sheet made with the plant's numbers on an admin's first open
    Dim m0 As Date: m0 = DateSerial(Year(Date), Month(Date), 1)
    If CalMonth() <> m0 Then
        MapWrite "Z1", m0
        DrawCalendar
    ElseIf CalStale() Then
        DrawCalendar
    End If
End Sub

' Switching to the Calendar tab: redraw only if something changed since it was drawn.
Public Sub CalActivated()
    If mDrawing Then Exit Sub
    On Error Resume Next                            ' never raises
    If CalStale() Then DrawCalendar Else SetCalHint CalendarSheet()
End Sub

Public Sub NAActivated()
    If mDrawing Then Exit Sub
    If Not gAdminUnlocked Then Exit Sub
    On Error Resume Next                            ' never raises
    If NAStale() Then DrawNeeds
End Sub

Public Sub CalSheetActivated(ByVal Sh As Object)
    mEventsSeen = True
    On Error Resume Next                            ' never raises
    If Sh.Name = CAL_SHEET Then
        CalActivated
    ElseIf Sh.Name = NA_SHEET Then
        NAActivated
    End If
End Sub

Public Sub CalSheetDoubleClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    If Sh.Name = CAL_SHEET Then
        CalDoubleClick Target, Cancel
    ElseIf Sh.Name = NA_SHEET Then
        NADoubleClick Target, Cancel
    End If
End Sub

Public Sub CalSheetRightClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    If Sh.Name = CAL_SHEET Then
        CalRightClick Target, Cancel
    ElseIf Sh.Name = NA_SHEET Then
        NARightClick Target, Cancel
    End If
End Sub

' v27 (item 32): the last calendar draw's seconds, from the CalMap line "... drawn in N s" (Empty when never drawn). Never raises.
Public Function CalDrawSeconds() As Variant
    On Error Resume Next                            ' never raises
    CalDrawSeconds = Empty
    If Not SheetExists(MAP_SHEET) Then Exit Function
    Dim s As String: s = SafeStr(MapSheet().Range("Z14").Value)
    Dim k As Long: k = InStr(1, s, "drawn in ", vbTextCompare)
    If k > 0 Then CalDrawSeconds = Val(Mid$(s, k + 9))
End Function

' v27 (item 27, 5.5): ONE legend - the words, colours and marks the calendar and the log use - for the calendar's
' printed page, the Desk and (with Members M4) the members' file
Public Function CalLegendText() As String
    CalLegendText = "OPEN (red) = nobody on it yet   POSTED (orange) = posted, waiting for sign-ups   NO TAKERS = posted, nobody signed up   " & _
                    "NCN = no coverage needed   a name = covered   S/S = shift switch   N short = below minimum staffing"
End Function

Public Function OutcomeLegendText() As String
    OutcomeLegendText = "A accepted   R refused (charged x" & Replace(CStr(CHARGE_MULTIPLIER), ",", ".") & ")   L left message (charged)   " & _
                        "U unavailable (not charged)   E plant upset (exception, not charged)   Not offered = a partial acceptance closed the shift   " & _
                        "~ = hours after a planned move   * = provisional OT   >> = called, never charged"
End Function
