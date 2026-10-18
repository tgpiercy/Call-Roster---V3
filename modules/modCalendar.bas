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
' ====================================================================

Public Const AB_SHEET As String = "Absences"
Public Const AB_TABLE As String = "tblAbsence"
Public Const CV_SHEET As String = "Coverage"
Public Const CV_TABLE As String = "tblCoverage"
Public Const CAL_SHEET As String = "Calendar"
Private Const MAP_SHEET As String = "CalMap"
Private Const FULL_SHIFT As Double = 12
Public Const CAL_VERSION As String = "Calendar v25.2 (Oct 17)"
' v25.1.2: draw-time measurement (the v25.2 plan's step 1): seconds spent inside SlotsFor and AddNote
' during the current draw, read by DrawCalendar / DrawNeeds for their breakdown lines
Private mTimeSlots As Double
Private mTimeNotes As Double
Public Const NA_SHEET As String = "Needs Attention"
Private Const NA_KEY As Long = 26           ' hidden key column on the Needs Attention tab
Private Const CAL_FONT As String = "Segoe UI"
Private Const CAL_FONT_B As String = "Segoe UI Semibold"
Private Const COL_W As Double = 22          ' width of each half of a day

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
Private mPostV As Variant           ' tblPostings snapshot for notes and the attention box
Private mMenuKey As String          ' line (or "P:P-0001") the double-click menu is for
Private mDrawing As Boolean
Private mEventsSeen As Boolean      ' True once ThisWorkbook's tab-switch / change events have fired
Private Const MENU_NAME As String = "CalLineMenu"

' one row on the Needs Attention tab
Private Type NARow
    Cell(0 To 7) As String
    Key As String
    SortKey As Double
    Accent As Long
    StatusCol As Long
    StatusFill As Long
    WhenCol As Long
End Type

' cache for one operation (tables + roster shifts)
Private mAbs As Variant
Private mCov As Variant
Private mShiftMap As Object
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
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": TypeColor = RGB(106, 27, 154) ' purple
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
Private Function AskNumber(ByVal prompt As String, ByVal title As String, ByVal dflt As Long, _
                           ByVal lo As Long, ByVal hi As Long) As Long
    If Len(prompt) > 1000 Then prompt = Left$(prompt, 990) & vbNewLine & "..."
    Dim s As String: s = Trim$(VBA.InputBox(prompt, title, CStr(dflt)))
    AskNumber = -1
    If s = "" Or Not IsNumeric(s) Then Exit Function
    If CDbl(s) <> Int(CDbl(s)) Then Exit Function
    If CLng(s) < lo Or CLng(s) > hi Then
        MsgBox "Pick a number from " & lo & " to " & hi & ".", vbExclamation
        Exit Function
    End If
    AskNumber = CLng(s)
End Function

Private Function Clip(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then Clip = Left$(s, n - 1) & "~" Else Clip = s
End Function

Private Function HrsTag(ByVal h As Variant) As String
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
        MsgBox "There is no Absences table yet - it will be created with every column the first time time off is booked. Nothing to do.", vbInformation
        Exit Sub
    End If
    If lo.ListColumns.Count >= AB_N Then MsgBox "The Absences table already has the v22 columns.", vbInformation: Exit Sub
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
        MsgBox "Couldn't save the backup copy, so nothing was changed:" & vbNewLine & e, vbCritical
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
        MsgBox "The upgrade did NOT verify:" & vbNewLine & problems & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical
        Exit Sub
    End If
    ResetCalCacheAll
    SaveIfWritable
    MsgBox "Absences table upgraded: " & lo.ListColumns.Count & " columns, " & nRows & " rows, roster totals unchanged." & vbNewLine & "Backup: " & bk, vbInformation
    Exit Sub
Failed:
    em = Err.Description
    LogError "UpgradeAbsenceColumns", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MarkNoSave
    MsgBox "The upgrade stopped part-way: " & em & vbNewLine & "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical
End Sub

Public Sub ResetCalCacheAll()
    If mTestMode Then Exit Sub              ' SelfTest's fixtures stay until it clears them
    mAbs = Empty: mCov = Empty: Set mShiftMap = Nothing
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

Private Function TableData(ByVal lo As ListObject) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    TableData = RangeTo2D(lo.DataBodyRange)
End Function

' The Absences table as a 2-D array (v22.1: modAttendance reads it)
Public Function AbsTableData() As Variant
    AbsTableData = TableData(AbsTable())
End Function

Private Function NextId(ByVal lo As ListObject, ByVal prefix As String) As String
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

Private Function AppendTo(ByVal lo As ListObject, ByVal data As Variant, ByVal nCols As Long) As Boolean
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
    MsgBox "Could not write to " & lo.Name & ": " & em, vbCritical
End Function

' row number (in the table body) of an ID, 0 if not found
Private Function RowOf(ByVal v As Variant, ByVal id As String) As Long
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

Private Sub LoadCache()
    If mTestMode Then Exit Sub
    mAbs = TableData(AbsTable())
    mCov = TableData(CovTable())
    Set mShiftMap = CreateObject("Scripting.Dictionary")
    mShiftMap.CompareMode = 1
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_SHIFT)).Value
    Dim i As Long, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" Then mShiftMap(nm) = UCase$(Trim$(SafeStr(v(i, 2))))
    Next i
End Sub


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

Private Function AbsActiveOn(ByVal r As Long, ByVal d As Date) As Boolean
    If UCase$(SafeStr(mAbs(r, AB_STATUS))) = "VOID" Then Exit Function
    If Not IsDate(mAbs(r, AB_START)) Then Exit Function
    If Int(CDate(mAbs(r, AB_START))) > Int(d) Then Exit Function
    If IsDate(mAbs(r, AB_END)) Then
        If Int(CDate(mAbs(r, AB_END))) < Int(d) Then Exit Function
    End If
    AbsActiveOn = True
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

' Fresh check (reloads the tables): is this person booked off on d?
' typ gets the public label (Vac / Trn / Leave) - never STD / PAR.
Public Function BookedOff(ByVal empName As String, ByVal d As Date, Optional ByRef typ As String) As Boolean
    If mShiftMap Is Nothing Then LoadCache
    If Not mTestMode Then mAbs = TableData(AbsTable())   ' just the absences - cheap enough per keystroke
    BookedOff = AbsentOn(empName, d, typ)
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

' The rotation part of the rule, pure (v21 - mirrored in tests/):
'   "scheduled" on that shift's date, "rest" (coming off nights that
'   morning for day OT, starting days the next morning for night OT),
'   "" otherwise. A/B, C/D and DAYS codes are handled by CrewShiftOn,
'   including plant holidays for DAYS.
Public Function RotationFreeReason(ByVal crewCode As String, ByVal d As Date, ByVal shiftLbl As String) As String
    RotationFreeReason = RotationFreeReason3(crewCode, crewCode, crewCode, d, shiftLbl)
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
                              ByVal d As Date, ByVal shiftLbl As String) As String
    If mShiftMap Is Nothing Then LoadCache
    Dim night As Boolean: night = IsNightShift(shiftLbl)
    Dim t As String
    If AbsentForShift(empName, d, shiftLbl, t) Then NotFreeReason = t: Exit Function
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
Public Function FilterNotFree(ByRef recs() As EmpRec, ByVal n As Long, ByVal d As Date, ByVal shiftLbl As String, _
                              ByRef removed As String, Optional ByVal useCache As Boolean = False) As Long
    If Not useCache Then LoadCache
    removed = ""
    Dim i As Long, k As Long, why As String
    Dim nSched As Long, nRest As Long, offTxt As String, ssTxt As String
    For i = 1 To n
        why = NotFreeReason(recs(i).EmpName, recs(i).EmpShift, d, shiftLbl)
        Select Case why
            Case ""
                k = k + 1
                If k <> i Then recs(k) = recs(i)
            Case "scheduled": nSched = nSched + 1
            Case "rest": nRest = nRest + 1
            Case "S/S", "S/S off": ssTxt = ssTxt & IIf(ssTxt = "", "", ", ") & CalName(recs(i).EmpName)
            Case Else: offTxt = offTxt & IIf(offTxt = "", "", ", ") & CalName(recs(i).EmpName) & " (" & why & ")"
        End Select
    Next i
    If nSched > 0 Then removed = nSched & " scheduled that day"
    If nRest > 0 Then removed = removed & IIf(removed = "", "", ", ") & nRest & " rest rule (off nights / days next morning)"
    If offTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "booked off: " & offTxt
    If ssTxt <> "" Then removed = removed & IIf(removed = "", "", "; ") & "shift switch: " & ssTxt
    FilterNotFree = k
End Function

' Drops anyone booked off on date d from a call / posting list.
' removed gets "Name (Vac), Name (Leave)". Returns the new count.
Public Function FilterAbsent(ByRef recs() As EmpRec, ByVal n As Long, ByVal d As Date, ByRef removed As String) As Long
    LoadCache
    removed = ""
    Dim i As Long, k As Long, t As String
    For i = 1 To n
        t = ""
        If AbsentOn(recs(i).EmpName, d, t) Then
            removed = removed & IIf(removed = "", "", ", ") & recs(i).EmpName & " (" & t & ")"
        Else
            k = k + 1
            If k <> i Then recs(k) = recs(i)
        End If
    Next i
    FilterAbsent = k
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
Private Function AbsenceOnShift(ByVal r As Long, ByVal d As Date, ByVal shiftLbl As String) As Boolean
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
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = AbsTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < AB_N Then
        MsgBox "The Absences table needs its v22 columns first: run Alt+F8 > UpgradeAbsenceColumns.", vbExclamation
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
Public Function SetAbsenceEndTo(ByVal id As String, ByVal dEnd As Date) As Boolean
    Dim lo As ListObject: Set lo = AbsTable()
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Function
    UnlockSheet lo.Parent
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
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = CovTable(True)
    If lo Is Nothing Then Exit Function
    If dEnd < dStart Then dEnd = dStart
    Dim shifts As Variant
    If shiftSel = "Both" Then shifts = Array("Days", "Nights") Else shifts = Array(shiftSel)
    Dim nRows As Long: nRows = (Int(dEnd) - Int(dStart) + 1) * (UBound(shifts) + 1)
    If nRows > 400 Then MsgBox "That's more than 400 lines - shorten the range.", vbExclamation: Exit Function

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

Private Sub FormatCovDates(ByVal lo As ListObject)
    On Error Resume Next
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
    If Left$(slotKey, 2) = "C:" Then
        Dim r As Long: r = RowOf(v, Mid$(slotKey, 3))
        If r > 0 Then
            Dim cur As String: cur = UCase$(SafeStr(v(r, CV_BY)))
            Dim curSrc As String: curSrc = SafeStr(v(r, CV_SOURCE))
            ' fill in place if empty, or if it's this same posting's POSTED line;
            ' a POSTED line from another source keeps its posting (new line below)
            If cur = "" Or (cur Like "POSTED*" And StrComp(curSrc, source, vbTextCompare) = 0) Then
                UnlockSheet lo.Parent
                lo.DataBodyRange.Cells(r, CV_BY).Value = coveredBy
                If IsNumeric(hrs) And Not IsEmpty(hrs) Then
                    If CDbl(hrs) > 0 And CDbl(hrs) < FULL_SHIFT Then lo.DataBodyRange.Cells(r, CV_HOURS).Value = CDbl(hrs)
                End If
                If source <> "" Then lo.DataBodyRange.Cells(r, CV_SOURCE).Value = source
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
    For i = 1 To UBound(v, 1)
        If StrComp(SafeStr(v(i, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(v(i, CV_STATUS))) <> "VOID" Then
            by = SafeStr(v(i, CV_BY))
            If empName = "" Then newBy = "" Else newBy = DropName(by, CalName(empName))
            If newBy <> by Or empName = "" Then
                If Not hit Then UnlockSheet lo.Parent: hit = True
                If newBy <> "" Then
                    lo.DataBodyRange.Cells(i, CV_BY).Value = newBy
                ElseIf StrComp(SafeStr(v(i, CV_NEED)), "Call-out " & source, vbTextCompare) = 0 Then
                    lo.DataBodyRange.Cells(i, CV_STATUS).Value = "VOID"     ' the call-out made this line - remove it
                Else
                    ' keep the line and its source (ReopenPosting can set it back to POSTED)
                    lo.DataBodyRange.Cells(i, CV_BY).ClearContents
                    lo.DataBodyRange.Cells(i, CV_HOURS).ClearContents
                End If
                ReleaseCoverage = ReleaseCoverage + 1
            End If
        End If
    Next i
    If hit Then ProtectSheet lo.Parent
    If hit Then MarkCalDirty
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
                              Optional ByVal preferNew As Boolean = False, Optional ByVal slotKey As String = "") As String
    Dim k As String: k = slotKey
    If k = "" And gSmokeRun Then k = "NEW"          ' v24.2: SmokeTest takes a new line without asking
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
Public Function CalendarAddAccepted(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String, _
                                    ByVal empName As String, ByVal hrs As Double, ByVal source As String) As String
    On Error GoTo Failed
    Dim t As String
    t = CalendarPlace(d, shiftLbl, unit, "Call-out " & source & ": which calendar line does " & empName & " cover?", _
                      CalName(empName), hrs, source, "SUP", "Call-out " & source)
    If t = "" Then
        CalendarAddAccepted = empName & ": not put on the calendar"
    Else
        CalendarAddAccepted = empName & ": " & t
        SaveIfWritable
    End If
    RefreshCalendar
    Exit Function
Failed:
    CalendarAddAccepted = empName & ": calendar NOT updated (" & Err.Description & ")"
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
        If Not ParseTypedDate(s, dEnd) Then MsgBox "That isn't a date (e.g. Oct 12).", vbExclamation: Exit Function
        If dEnd < CDate(v(r, AB_START)) Then MsgBox "The end is before the start.", vbExclamation: Exit Function
    End If
    UnlockSheet lo.Parent
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
    On Error Resume Next
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
    End Select
End Sub

Private Function DrawContext() As String
    DrawContext = Format$(Date, "yyyymmdd") & "|" & CAL_VERSION
End Function

' order-independent fingerprint of roster name + crew (crew changes move time off)
Private Function RosterSig() As String
    On Error Resume Next
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

Private Function NAStale() As Boolean
    On Error GoTo Yes
    Dim mp As Worksheet: Set mp = MapSheet()
    If SafeStr(mp.Range("Z9").Value) <> SafeStr(mp.Range("Z7").Value) Then GoTo Yes
    If SafeStr(mp.Range("Z11").Value) <> DrawContext() Then GoTo Yes
    Exit Function
Yes:
    NAStale = True
End Function

' the toolbar hint depends on who's looking - cheap to fix without a redraw
' v25.1.2: the hint text, with the last draw's time on the end (CalMap Z13 holds the seconds)
Private Function CalHintText() As String
    CalHintText = IIf(gAdminUnlocked, "Double-click any line or empty space for options.  Hover an open or posted line for details.", _
                      "Double-click a line for its details; hover an open or posted line.  Print / PDF gives the copy to post.")
    Dim s As String: s = SafeStr(MapSheet().Range("Z13").Value)
    If s <> "" Then CalHintText = CalHintText & "  Drawn in " & s & " s."
End Function

' v25.1.2: seconds between two Timer readings (Timer restarts at midnight)
Private Function Secs(ByVal t0 As Double, ByVal t1 As Double) As Double
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
           "Send me these lines (a photo is fine). The 'drawn in' number is the draw alone: a stopwatch from the click to the screen settling is the number you feel.  (" & CAL_VERSION & ")", _
           vbInformation, "Draw times"
    Exit Sub
Oops:
    ReportError "ShowCalDrawTimes"
End Sub

Private Sub SetCalHint(ByVal ws As Worksheet)
    On Error Resume Next
    Dim t As String
    t = CalHintText()                                       ' v25.1.2: with the last draw's time
    If SafeStr(ws.Range("F3").Value) = t Then Exit Sub
    If Left$(SafeStr(ws.Range("F3").Value), 8) = "Calendar" Then Exit Sub   ' leave an error message showing
    UnlockSheet ws
    ws.Range("F3").Value = t
    ProtectSheet ws
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

Private Function CalendarSheet() As Worksheet
    If Not SheetExists(CAL_SHEET) Then
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=RosterSheet())
        ws.Name = CAL_SHEET
        RelockStructure
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

Private Function MapSheet() As Worksheet
    Set MapSheet = ThisWorkbook.Worksheets(MAP_SHEET)
End Function

' v25.1: CalMap is protected like every sheet (v25), and a bare write to it raised "Application-defined
' or object-defined error" from the Calendar's unit list on Graeme's Oct 16 test. Every other sheet is
' written between UnlockSheet / ProtectSheet; the map now is too (DrawCalendar unlocks it for the whole
' draw). The map is relocked whatever happened and the error re-raised after (v25.1 review: MarkCalDirty,
' CalOnOpen and DrawNeeds never reach ReportError, so nothing else would relock it).
Private Sub MapWrite(ByVal addr As String, ByVal v As Variant)
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
    On Error Resume Next
    CalDirtyStamp = SafeStr(MapSheet().Range("Z7").Value)
End Function

Private Function CalMonth() As Date
    Dim v As Variant: v = MapSheet().Range("Z1").Value
    If IsDate(v) Then CalMonth = DateSerial(Year(v), Month(v), 1) Else CalMonth = DateSerial(Year(Date), Month(Date), 1)
End Function

Private Function CalGroup() As String
    CalGroup = SafeStr(MapSheet().Range("Z2").Value)
    If CalGroup = "" Then CalGroup = CalGroupNames()(0)
End Function

' the one unit of a calendar group, or "" when it holds several
Private Function SingleUnitOf(ByVal grp As String) As String
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
Private Function thBand() As Long          ' title band, borders around each day
    thBand = RGB(38, 50, 56)
End Function
Private Function thHead() As Long          ' weekday header, section bars
    thHead = RGB(55, 71, 79)
End Function
Private Function thNav() As Long           ' month / navigation buttons
    thNav = RGB(84, 110, 122)
End Function
Private Function thGreen() As Long         ' print / post / primary actions
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
Private Function thToday() As Long
    thToday = RGB(255, 213, 79)
End Function
Private Function thWeekend() As Long
    thWeekend = RGB(236, 239, 241)
End Function
Private Function thOutFill() As Long       ' days outside the month
    thOutFill = RGB(245, 245, 245)
End Function
Private Function thGrid() As Long
    thGrid = RGB(207, 216, 220)
End Function
Private Function thOpen() As Long          ' OPEN / no takers
    thOpen = RGB(183, 28, 28)
End Function
Private Function thOpenFill() As Long
    thOpenFill = RGB(255, 205, 210)
End Function
Private Function thPosted() As Long        ' POSTED - waiting for sign-ups
    thPosted = RGB(191, 54, 12)
End Function
Private Function thPostedFill() As Long
    thPostedFill = RGB(255, 224, 178)
End Function
Private Function thInk() As Long           ' normal text / covered names
    thInk = RGB(33, 33, 33)
End Function
Private Function thMuted() As Long
    thMuted = RGB(117, 117, 117)
End Function

' a colour mixed toward white (f = 0.85 -> very light tint)
Private Function Tint(ByVal col As Long, ByVal f As Double) As Long
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
    Dim nLines As Long
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
    tA = Timer                                      ' caches loaded

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
    tB = Timer                                      ' grid wiped

    ' ---- row 1: title band (controls sit on it) ----
    ws.Rows(1).RowHeight = 38
    ws.Range("A1:P1").Interior.Color = thBand()
    With ws.Range("B1:C1")
        .Merge
        .Value = "COVERAGE CALENDAR"
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

    ' ---- row 3: admin toolbar (buttons added in BuildCalControls) ----
    ws.Rows(3).RowHeight = 28
    With ws.Range("F3:K3")
        .Merge
        .Value = CalHintText()                  ' v25.2: one text (rewritten with this draw's time at the end)
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = thMuted()
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    With ws.Range("M3:O3")
        .Merge
        .Value = CAL_VERSION
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        .Font.Size = 7: .Font.Color = RGB(176, 190, 197)
    End With
    mp.Range("Z4").Value = 3                         ' rows 3..Z4 are screen-only (hidden when printing)

    ' ---- row 4: weekday header ----
    Dim r As Long: r = 4
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
    tC = Timer                                      ' header rows written
    mTimeSlots = 0: mTimeNotes = 0                  ' the weeks' share starts here

    ' ---- weeks ----
    Dim mEnd As Date: mEnd = DateSerial(Year(m0), Month(m0) + 1, 0)
    Dim wk As Date: wk = m0 - (Weekday(m0, vbSunday) - 1)
    Dim slotsD() As CalSlot, slotsN() As CalSlot
    Dim d As Date, nD As Long, nN As Long, mxD As Long, mxN As Long, rr As Long, inMonth As Boolean
    Dim bg As Long, nOpen(0 To 6) As Long, sv As Long
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

        ' date row: number over "Days: B Shift" | "Nights: A Shift" in the pair column
        ws.Rows(r).RowHeight = 38
        For k = 0 To 6
            d = wk + k
            inMonth = (Month(d) = Month(m0))
            If Not inMonth Then
                bg = thOutFill()
            ElseIf Int(d) = Int(Date) Then
                bg = thToday()
            ElseIf k = 0 Or k = 6 Then
                bg = thWeekend()
            Else
                bg = vbWhite
            End If
            ws.Range(ws.Cells(r, 2 + 2 * k), ws.Cells(r, 3 + 2 * k)).Interior.Color = bg
            If inMonth Then
                DateCells ws.Cells(r, 2 + 2 * k), ws.Cells(r, 3 + 2 * k), d, nOpen(k)
                sv = CLng(Int(d))
                mp.Cells(r, 2 + 2 * k).Value = "D:DAY|" & sv & "|"
                mp.Cells(r, 3 + 2 * k).Value = "D:DAY|" & sv & "|"
            End If
        Next k
        r = r + 1

        ' day slots, then night slots (empty rows are clickable too)
        For k = 0 To 6
            d = wk + k
            If Month(d) = Month(m0) Then
                sv = CLng(Int(d))
                nD = SlotsFor(d, "Days", grp, slotsD, True)
                nN = SlotsFor(d, "Nights", grp, slotsN, True)
                nLines = nLines + nD + nN
                ws.Range(ws.Cells(r, 2 + 2 * k), ws.Cells(r + mxD - 1, 3 + 2 * k)).Interior.Color = thDayFill()
                ws.Range(ws.Cells(r + mxD, 2 + 2 * k), ws.Cells(r + mxD + mxN - 1, 3 + 2 * k)).Interior.Color = thNightFill()
                For i = 1 To mxD
                    If i <= nD Then
                        WriteSlot ws, mp, r + i - 1, k, slotsD(i), d, "Days"
                    Else
                        mp.Cells(r + i - 1, 2 + 2 * k).Value = "E:NEW|" & sv & "|Days"
                        mp.Cells(r + i - 1, 3 + 2 * k).Value = "E:NEW|" & sv & "|Days"
                    End If
                Next i
                For i = 1 To mxN
                    If i <= nN Then
                        WriteSlot ws, mp, r + mxD + i - 1, k, slotsN(i), d, "Nights"
                    Else
                        mp.Cells(r + mxD + i - 1, 2 + 2 * k).Value = "E:NEW|" & sv & "|Nights"
                        mp.Cells(r + mxD + i - 1, 3 + 2 * k).Value = "E:NEW|" & sv & "|Nights"
                    End If
                Next i
            Else
                ws.Range(ws.Cells(r, 2 + 2 * k), ws.Cells(r + mxD + mxN - 1, 3 + 2 * k)).Interior.Color = thOutFill()
            End If
        Next k
        For rr = r To r + mxD + mxN - 1
            ws.Rows(rr).RowHeight = 17
        Next rr
        SectionLabel ws, r, mxD, "DAYS", thDayTab(), thDayText()
        SectionLabel ws, r + mxD, mxN, "NIGHTS", thNightTab(), thNightText()

        ' borders: thick around each day, thin day/night split, dotted inside the pair
        Dim blockTop As Long: blockTop = r - 1
        Dim blockBot As Long: blockBot = r + mxD + mxN - 1
        With ws.Range(ws.Cells(blockTop, 1), ws.Cells(blockTop, 15)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlMedium: .Color = thBand()
        End With
        With ws.Range(ws.Cells(blockTop, 2), ws.Cells(blockTop, 15)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = thGrid()
        End With
        With ws.Range(ws.Cells(r + mxD - 1, 1), ws.Cells(r + mxD - 1, 15)).Borders(xlEdgeBottom)
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

        r = r + mxD + mxN
        wk = wk + 7
    Loop
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 15)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Weight = xlThick: .Color = thBand()
    End With
    tD = Timer                                      ' weeks written
    Dim wkSlots As Double, wkNotes As Double
    wkSlots = mTimeSlots: wkNotes = mTimeNotes      ' the weeks' share (OpenAhead's SlotsFor calls count under controls)

    ' ---- page ----
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 16)).Address
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
    End With
    Application.PrintCommunication = True
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

' date row of one day: big number (+ red "2 open" flag) with "Days: B Shift"
' under it on the left, "Nights: A Shift" bottom-right in the pair column
Private Sub DateCells(ByVal a As Range, ByVal b As Range, ByVal d As Date, ByVal nOpen As Long)
    Dim num As String: num = CStr(Day(d))
    Dim flag As String
    If nOpen > 0 Then flag = "   " & ChrW(9679) & " " & nOpen & " open"
    Dim t1 As String: t1 = "Days: " & DaysCrew(d) & " Shift"
    Dim past As Boolean: past = (Int(d) < Int(Date))
    a.Value = num & flag & vbLf & t1
    a.WrapText = True
    a.HorizontalAlignment = xlLeft: a.VerticalAlignment = xlBottom
    a.IndentLevel = 1
    a.Font.Name = CAL_FONT_B: a.Font.Size = 8: a.Font.Bold = False: a.Font.Color = thDayText()
    With a.Characters(1, Len(num)).Font
        .Size = 15
        .Color = IIf(past, RGB(158, 158, 158), thBand())
    End With
    If flag <> "" Then
        With a.Characters(Len(num) + 1, Len(flag)).Font
            .Size = 8
            .Color = IIf(past, RGB(158, 158, 158), thOpen())
        End With
    End If

    ' plant holiday name top-right (v20), "Nights: A Shift" under it
    Dim hol As String: hol = HolidayLabel(d)
    Dim t2 As String: t2 = "Nights: " & NightsCrew(d) & " Shift"
    If hol <> "" Then
        b.Value = hol & vbLf & t2
        b.WrapText = True
    Else
        b.Value = t2
    End If
    b.HorizontalAlignment = xlRight: b.VerticalAlignment = xlBottom
    b.IndentLevel = 1
    b.Font.Name = CAL_FONT_B: b.Font.Size = 8: b.Font.Color = thNightText()
    If hol <> "" Then
        With b.Characters(1, Len(hol)).Font
            .Color = IIf(past, RGB(158, 158, 158), RGB(176, 0, 32))
        End With
    End If
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

Private Sub WriteSlot(ByVal ws As Worksheet, ByVal mp As Worksheet, ByVal rr As Long, ByVal k As Long, _
                      ByRef s As CalSlot, ByVal d As Date, ByVal shiftLbl As String)
    Dim a As Range: Set a = ws.Cells(rr, 2 + 2 * k)
    Dim b As Range: Set b = ws.Cells(rr, 3 + 2 * k)
    a.Value = "'" & s.LeftText
    a.Font.Name = CAL_FONT_B: a.Font.Size = 9: a.Font.Color = TypeColor(s.TypeCode)
    a.IndentLevel = 1
    a.ShrinkToFit = True                  ' long names shrink instead of spilling into the next day
    a.VerticalAlignment = xlCenter
    Dim t As String: t = s.RightText
    Dim noteIt As Boolean                     ' v25.2 (item 1): a hover note only where action is needed
    b.Font.Name = CAL_FONT: b.Font.Size = 9: b.Font.Bold = False: b.Font.Italic = False
    b.IndentLevel = 1
    b.ShrinkToFit = True
    b.VerticalAlignment = xlCenter
    If t = "" Or UCase$(t) Like "NO TAKERS*" Then
        b.Value = IIf(t = "", "OPEN", "NO TAKERS")
        b.Font.Name = CAL_FONT_B: b.Font.Color = thOpen()
        b.Interior.Color = thOpenFill()
        noteIt = True
    ElseIf UCase$(t) Like "POSTED*" Then
        b.Value = "'" & t
        b.Font.Name = CAL_FONT_B: b.Font.Color = thPosted()
        b.Interior.Color = thPostedFill()
        noteIt = True
    ElseIf UCase$(t) = "NCN" Then
        b.Value = "'" & t: b.Font.Italic = True: b.Font.Color = thMuted()
    Else
        b.Value = "'" & t: b.Font.Name = CAL_FONT_B: b.Font.Color = thInk()
    End If
    Dim keyTxt As String: keyTxt = s.Key & "|" & CLng(Int(d)) & "|" & shiftLbl
    mp.Cells(rr, 2 + 2 * k).Value = keyTxt
    mp.Cells(rr, 3 + 2 * k).Value = keyTxt
    ' v25.2 (item 1, D-C4 revisited on the v25.1.2 numbers: 85 notes were 2.8 of 4.6 s): OPEN, NO TAKERS and
    ' POSTED lines keep the hover box; a covered line's details are on double-click > Details...
    If noteIt Then
        Dim tL As Double: tL = Timer          ' v25.1.2 review: the details text counts under notes, not cells
        Dim det As String: det = LineDetails(keyTxt, True)
        mTimeNotes = mTimeNotes + Secs(tL, Timer)
        AddNote b, det                        ' hover for the details
    End If
End Sub

Private Sub BuildCalControls(ByVal ws As Worksheet, ByVal m0 As Date, ByVal grp As String)
    On Error Resume Next                  ' one control failing must not stop the rest
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "cal_" Then ws.Shapes(i).Delete
    Next i

    ' row 1:  [Needs attention]  < Sep | [October 2026 v] | Nov > | This month   [unit v]   Print / PDF
    Dim nOpen As Long: nOpen = OpenAhead(grp, 14)
    CalBtn ws, "cal_adm_na", IIf(nOpen > 0, nOpen & " need cover  >", "Needs attention  >"), _
           "modCalendar.ShowNeedsAttention", ws.Range("D1:E1"), IIf(nOpen > 0, thOpen(), thNav())
    CalBtn ws, "cal_prev", "<  " & Format$(DateAdd("m", -1, m0), "mmm"), "modCalendar.CalPrev", ws.Range("F1"), thNav()
    CalDrop ws, "cal_month", ws.Range("G1:H1")
    CalBtn ws, "cal_next", Format$(DateAdd("m", 1, m0), "mmm") & "  >", "modCalendar.CalNext", ws.Range("I1"), thNav()
    CalBtn ws, "cal_today", "This month", "modCalendar.CalToday", ws.Range("J1"), thNav()
    CalDrop ws, "cal_unit", ws.Range("L1:M1")
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
    CalBtn ws, "cal_adm_imp", "Import Old Cal", "modCalendar.ImportUnitCalendar", ws.Range("E3"), RGB(120, 144, 156)
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 8) = "cal_adm_" Then ws.Shapes(i).Visible = IIf(gAdminUnlocked, msoTrue, msoFalse)
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

' button centred in its cell(s), fixed height so the text never clips
Private Sub CalBtn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
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
    On Error Resume Next
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
Private Sub CalDrop(ByVal ws As Worksheet, ByVal nm As String, ByVal cel As Range)
    Dim dd As Object
    Set dd = ws.DropDowns.Add(cel.Left + 4, cel.Top + (cel.Height - 20) / 2, cel.Width - 8, 20)
    dd.Name = nm
    dd.PrintObject = False
End Sub


' Admin: click a calendar line, then Set Coverage
Public Sub CalSetCoverage()
    On Error GoTo Oops
    If Not ActiveSheet.Name = CAL_SHEET Then MsgBox "Click a line on the Calendar first.", vbInformation: Exit Sub
    CalSetCoverageKey SafeStr(MapSheet().Range(ActiveCell.Address).Value)
    Exit Sub
Oops:
    ReportError "CalSetCoverage"
End Sub

Private Sub CalSetCoverageKey(ByVal keyTxt As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Left$(keyTxt, 2) <> "A:" And Left$(keyTxt, 2) <> "C:" Then MsgBox "Click a line on the calendar (who's off / the need), then Set Coverage." & vbNewLine & "For an empty spot, double-click it instead.", vbInformation: Exit Sub
    Dim parts As Variant: parts = Split(keyTxt, "|")
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim shiftLbl As String: shiftLbl = parts(2)
    Dim leftTxt As String, rightTxt As String
    If Not SlotTexts(keyTxt, leftTxt, rightTxt) Then MsgBox "That line has changed - the calendar will redraw.", vbExclamation: DrawCalendar: Exit Sub

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
        If UCase$(ans) = "DELETE" Then MsgBox "Void the time off in Book Time Off instead.", vbInformation: Exit Sub
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
                If MsgBox("This line is filled by" & srcs & ". Replace it anyway?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
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
Private Sub CalEditTextKey(ByVal keyTxt As String)
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
        Case "phone": TextRefusal = "Not changed: that holds a run of digits like a phone number or an ID, and Member View would refuse to publish it."
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
Private Sub CalEditReasonId(ByVal id As String)
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
    If ActiveSheet.Name <> CAL_SHEET Then MsgBox "Click a line on the Calendar first.", vbInformation: Exit Sub
    CalPostKey SafeStr(MapSheet().Range(ActiveCell.Address).Value)
    Exit Sub
Oops:
    ReportError "CalPostLine"
End Sub

Private Sub CalPostKey(ByVal keyTxt As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Left$(keyTxt, 2) <> "A:" And Left$(keyTxt, 2) <> "C:" Then MsgBox "Click the line that needs coverage (who's off / the need), then Post Line." & vbNewLine & "For an empty spot, double-click it to post support.", vbInformation: Exit Sub

    Dim parts As Variant: parts = Split(keyTxt, "|")
    Dim slotKey As String: slotKey = parts(0)
    Dim d As Date: d = CDate(CLng(parts(1)))
    Dim shiftLbl As String: shiftLbl = parts(2)
    Dim leftTxt As String, rightTxt As String     ' from the data, not the sheet (the sheet shows "OPEN")
    If Not SlotTexts(keyTxt, leftTxt, rightTxt) Then MsgBox "That line has changed - the calendar will redraw.", vbExclamation: DrawCalendar: Exit Sub
    rightTxt = Trim$(rightTxt)

    If UCase$(rightTxt) Like "POSTED*" Then
        MsgBox "That line is already " & rightTxt & ".", vbInformation: Exit Sub
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
    If r = 0 Then MsgBox "Couldn't find the record behind that line - redraw the calendar and try again.", vbExclamation: Exit Sub

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
    ans = MsgBox("Yes = print AND save a PDF" & vbNewLine & "No = save a PDF only", vbYesNoCancel + vbQuestion, "Calendar")
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
    On Error Resume Next
    ' the admin toolbar and the needs-attention box are screen-only
    Dim hideTo As Long: hideTo = CLng(Val(MapSheet().Range("Z4").Value))
    If hideTo < 3 Then hideTo = 3
    UnlockSheet ws
    ws.Rows("3:" & hideTo).Hidden = True
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=folder & fName, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then msg = "PDF NOT saved: " & Err.Description Else msg = "PDF saved: " & folder & fName
    Err.Clear
    If ans = vbYes Then
        ws.PrintOut Copies:=1
        If Err.Number <> 0 Then msg = msg & vbNewLine & "NOT printed: " & Err.Description Else msg = msg & vbNewLine & "Sent to printer."
    End If
    ws.Rows("3:" & hideTo).Hidden = False
    ProtectSheet ws
    On Error GoTo Oops
    MsgBox msg, vbInformation, "Calendar"
    Exit Sub
Oops:
    ReportError "CalPrintPdf"
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
'  ONE-TIME IMPORT of an old monthly calendar workbook
' ====================================================================
Public Sub ImportUnitCalendar()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub

    ' which unit
    Dim units As Variant: units = DeptList()
    Dim msg As String, i As Long
    For i = 0 To UBound(units)
        msg = msg & (i + 1) & ")  " & units(i) & vbNewLine
    Next i
    Dim pick As Long
    pick = AskNumber("Import an old calendar workbook into which unit?" & vbNewLine & vbNewLine & msg & vbNewLine & _
                     "(For the combined Phosphine / Derivatives calendar pick either - both show on that calendar.)", _
                     "Import calendar", 1, 1, UBound(units) + 1)
    If pick < 1 Then Exit Sub
    Dim unit As String: unit = units(pick - 1)

    Dim dIn As Variant
    dIn = Application.InputBox("Import entries from this date on:", "Import calendar", _
                               DateText(DateSerial(Year(Date), Month(Date), 1)), Type:=2)
    If VarType(dIn) = vbBoolean Then Exit Sub
    Dim dFrom As Date                              ' v25: the one parser
    If Not ParseTypedDate(CStr(dIn), dFrom) Then MsgBox "Not a date (e.g. Oct 1).", vbExclamation: Exit Sub

    Dim f As Variant
    f = Application.GetOpenFilename("Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Pick the " & unit & " calendar workbook")
    If VarType(f) = vbBoolean Then Exit Sub

    Dim lo As ListObject: Set lo = CovTable(True)
    Dim src As String: src = "Import " & Format$(Date, "yyyy-mm-dd")

    ' already imported for this unit / period? (avoid doubles)
    Dim ov As Variant: ov = TableData(lo)
    Dim prior As Long, j As Long
    If IsArray(ov) Then
        For j = 1 To UBound(ov, 1)
            If Left$(SafeStr(ov(j, CV_SOURCE)), 6) = "Import" And UCase$(SafeStr(ov(j, CV_STATUS))) <> "VOID" And _
               StrComp(SafeStr(ov(j, CV_UNIT)), unit, vbTextCompare) = 0 Then
                If IsDate(ov(j, CV_DATE)) Then
                    If CDate(ov(j, CV_DATE)) >= dFrom Then prior = prior + 1
                End If
            End If
        Next j
    End If
    If prior > 0 Then
        Select Case MsgBox(prior & " imported lines already exist for " & unit & " from " & Format$(dFrom, "mmm d") & "." & vbNewLine & vbNewLine & _
                           "Yes = remove those and import again" & vbNewLine & "No = keep them and add these too" & vbNewLine & _
                           "Cancel = stop", vbYesNoCancel + vbExclamation, "Import calendar")
            Case vbCancel: Exit Sub
            Case vbYes
                UnlockSheet lo.Parent
                For j = 1 To UBound(ov, 1)
                    If Left$(SafeStr(ov(j, CV_SOURCE)), 6) = "Import" And UCase$(SafeStr(ov(j, CV_STATUS))) <> "VOID" And _
                       StrComp(SafeStr(ov(j, CV_UNIT)), unit, vbTextCompare) = 0 Then
                        If IsDate(ov(j, CV_DATE)) Then
                            If CDate(ov(j, CV_DATE)) >= dFrom Then lo.DataBodyRange.Cells(j, CV_STATUS).Value = "VOID"
                        End If
                    End If
                Next j
                ProtectSheet lo.Parent
        End Select
    End If

    Application.ScreenUpdating = False
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=f, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo Oops
    If wb Is Nothing Then Application.ScreenUpdating = True: MsgBox "Couldn't open that file.", vbCritical: Exit Sub

    Dim buf() As Variant: ReDim buf(1 To 5000, 1 To CV_N)
    Dim n As Long, sh As Worksheet, nSheets As Long
    For Each sh In wb.Worksheets
        n = ParseMonthSheet(sh, unit, dFrom, src, buf, n)
        nSheets = nSheets + 1
    Next sh
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If n = 0 Then MsgBox "No entries found on or after " & Format$(dFrom, "mmm d, yyyy") & ".", vbInformation: Exit Sub
    If MsgBox("Found " & n & " calendar lines for " & unit & " from " & Format$(dFrom, "mmm d, yyyy") & "." & vbNewLine & vbNewLine & _
              "They come in as written (e.g. 'Paul M 3rd' -> 'posted') and show on the calendar." & vbNewLine & _
              "They don't remove anyone from call lists until you convert them (Convert imports button) or book the time off." & vbNewLine & vbNewLine & _
              "Import them?", vbOKCancel + vbQuestion, "Import calendar") <> vbOK Then Exit Sub

    Dim base As Long: base = CLng(Val(Mid$(NextId(lo, "CV"), 4)))
    Dim data() As Variant: ReDim data(1 To n, 1 To CV_N)
    Dim r As Long, c As Long, user As String, stamp As Date
    user = CurrentUser(): stamp = Now
    For r = 1 To n
        For c = 1 To CV_N
            data(r, c) = buf(r, c)
        Next c
        data(r, CV_ID) = "CV-" & Format$(base + r - 1, "0000")
        data(r, CV_USER) = user
        data(r, CV_STAMP) = stamp
    Next r
    If Not AppendTo(lo, data, CV_N) Then Exit Sub
    FormatCovDates lo
    SaveIfWritable
    MsgBox n & " lines imported for " & unit & ".", vbInformation
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "ImportUnitCalendar"
End Sub

' Reads one month tab: finds the SUNDAY..SATURDAY header, date rows, and
' the slot rows under each date (yellow = days, other fill = nights).
Private Function ParseMonthSheet(ByVal sh As Worksheet, ByVal unit As String, ByVal dFrom As Date, _
                                 ByVal src As String, ByRef buf() As Variant, ByVal n As Long) As Long
    ParseMonthSheet = n
    Dim hdrRow As Long, sunCol As Long, r As Long, c As Long
    For r = 1 To 10
        For c = 1 To 20
            If UCase$(Trim$(SafeStr(sh.Cells(r, c).Value))) = "SUNDAY" Then hdrRow = r: sunCol = c: Exit For
        Next c
        If hdrRow > 0 Then Exit For
    Next r
    If hdrRow = 0 Then Exit Function

    Dim dayCol(0 To 6) As Long, k As Long
    For k = 0 To 6
        dayCol(k) = sunCol + 2 * k
    Next k

    ' the month: a real date near the top (B1 / A1). The day cells only
    ' hold day numbers (=H4+1 ...), so dates are month + day number.
    Dim mStart As Date, found As Boolean, v As Variant
    For r = 1 To hdrRow
        For c = 1 To 8
            v = sh.Cells(r, c).Value2
            If Not IsError(v) Then
                If VarType(v) = vbDouble Then
                    If v > 40000 And v < 80000 Then mStart = DateSerial(Year(CDate(v)), Month(CDate(v)), 1): found = True
                End If
            End If
            If found Then Exit For
        Next c
        If found Then Exit For
    Next r
    If Not found Then Exit Function
    Dim mDays As Long: mDays = Day(DateSerial(Year(mStart), Month(mStart) + 1, 0))
    Dim lastRow As Long: lastRow = sh.UsedRange.Row + sh.UsedRange.Rows.Count
    If lastRow > 200 Then lastRow = 200

    ' date rows = any day column holding a date serial
    Dim isDateRow() As Boolean: ReDim isDateRow(1 To lastRow + 1)
    For r = hdrRow + 1 To lastRow
        For k = 0 To 6
            If DayNum(sh.Cells(r, dayCol(k)), mDays) > 0 Then isDateRow(r) = True: Exit For
        Next k
    Next r

    Dim cur(0 To 6) As Variant, rowD As Long, lt As String, rt As String, shiftLbl As String, colr As Long
    For r = hdrRow + 1 To lastRow
        If isDateRow(r) Then
            For k = 0 To 6
                If DayNum(sh.Cells(r, dayCol(k)), mDays) > 0 Then
                    cur(k) = DateSerial(Year(mStart), Month(mStart), DayNum(sh.Cells(r, dayCol(k)), mDays))
                Else
                    cur(k) = Empty
                End If
            Next k
        Else
            For k = 0 To 6
                If Not IsEmpty(cur(k)) Then
                    lt = Trim$(SafeStr(sh.Cells(r, dayCol(k)).Value))
                    rt = Trim$(SafeStr(sh.Cells(r, dayCol(k) + 1).Value))
                    If (lt <> "" Or rt <> "") And cur(k) >= dFrom Then
                        colr = sh.Cells(r, dayCol(k)).Interior.Color
                        If sh.Cells(r, dayCol(k)).Interior.ColorIndex = xlNone Then colr = sh.Cells(r, dayCol(k) + 1).Interior.Color
                        If colr = RGB(255, 255, 0) Then
                            shiftLbl = "Days"
                        ElseIf sh.Cells(r, dayCol(k)).Interior.ColorIndex = xlNone Or colr = RGB(255, 255, 255) Then
                            shiftLbl = "Days"
                        Else
                            shiftLbl = "Nights"
                        End If
                        n = n + 1
                        If n > UBound(buf, 1) Then n = n - 1: ParseMonthSheet = n: Exit Function
                        buf(n, CV_DATE) = Int(CDate(cur(k)))
                        buf(n, CV_SHIFT) = shiftLbl
                        buf(n, CV_UNIT) = unit
                        buf(n, CV_NEED) = IIf(lt = "", "(blank)", lt)
                        buf(n, CV_TYPE) = GuessType(lt & " " & rt)
                        If LCase$(rt) Like "posted*" Then rt = "POSTED" & Mid$(rt, 7)
                        buf(n, CV_BY) = rt
                        buf(n, CV_SOURCE) = src
                        buf(n, CV_STATUS) = "Active"
                        buf(n, CV_NOTE) = sh.Name
                    End If
                End If
            Next k
        End If
    Next r
    ParseMonthSheet = n
End Function

' Day-of-month held in a calendar date cell (1..mDays), 0 if none.
' Accepts a plain day number or a real date.
Private Function DayNum(ByVal c As Range, ByVal mDays As Long) As Long
    Dim v As Variant: v = c.Value2
    If IsError(v) Then Exit Function
    If VarType(v) <> vbDouble Then Exit Function
    If v <> Int(v) Then Exit Function
    If v > 40000 Then v = Day(CDate(v))
    If v >= 1 And v <= mDays Then DayNum = CLng(v)
End Function

Private Function GuessType(ByVal t As String) As String
    t = UCase$(t)
    If InStr(t, "S/S") > 0 Or InStr(t, " SS") > 0 Then
        GuessType = "SS"
    ElseIf InStr(t, "PHA") > 0 Then
        GuessType = "PHA"
    ElseIf InStr(t, "TRAIN") > 0 Then
        GuessType = "TRN"
    ElseIf InStr(t, "OPERATOR") > 0 Or InStr(t, " OP") > 0 Or InStr(t, "EXTRA") > 0 Or InStr(t, "FLOATER") > 0 Then
        GuessType = "SUP"
    Else
        GuessType = "VAC"
    End If
End Function


' ====================================================================
'  LINE DETAILS  (hover notes + Details in the menu)
' ====================================================================
Private Sub LoadPostCache()
    mPostV = Empty
    On Error Resume Next
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

Private Function PostRow(ByVal id As String) As Long
    If IsArray(mPostV) Then PostRow = RowOf(mPostV, id)
End Function

Private Function AsDate(ByVal x As Variant) As Date
    If IsDate(x) Then AsDate = CDate(x)
End Function

Private Function TypeLong(ByVal code As String) As String
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

Private Function StampTxt(ByVal who As Variant, ByVal stamp As Variant) As String
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
Private Function LineDetails(ByVal keyTxt As String, Optional ByVal useCache As Boolean = False) As String
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

Private Function PostingDetails(ByVal id As String) As String
    LoadPostCache
    Dim r As Long: r = PostRow(id)
    If r = 0 Then PostingDetails = "Posting " & id & " wasn't found.": Exit Function
    PostingDetails = SafeStr(mPostV(r, PS_UNIT)) & " - " & Format$(AsDate(mPostV(r, PS_DATE)), "dddd mmm d") & " " & _
                     SafeStr(mPostV(r, PS_SHIFT)) & "  (crew " & SafeStr(mPostV(r, PS_CREW)) & ")" & vbLf & _
                     PostingText(id) & vbLf & _
                     "Posted " & IIf(IsDate(mPostV(r, PS_POSTED)), Format$(mPostV(r, PS_POSTED), "mmm d"), "") & _
                     " by " & SafeStr(mPostV(r, PS_POSTEDBY))
End Function

' a hover note on a cell (screen only - notes don't print)
Private Sub AddNote(ByVal c As Range, ByVal txt As String)
    If txt = "" Then Exit Sub
    Dim tN As Double: tN = Timer                     ' v25.1.2: measured
    On Error Resume Next
    If Not c.Comment Is Nothing Then c.ClearComments      ' v25.2 (item 1): fresh rows have none
    c.AddComment txt
    With c.Comment.Shape
        With .TextFrame.Characters.Font
            .Name = CAL_FONT: .Size = 9: .Bold = False
        End With
        Dim p As Long: p = InStr(txt, vbLf)
        If p > 1 Then .TextFrame.Characters(1, p - 1).Font.Bold = True
        .TextFrame.AutoSize = True
        .Fill.ForeColor.RGB = RGB(255, 253, 240)
        .Line.ForeColor.RGB = RGB(150, 150, 150)
    End With
    mTimeNotes = mTimeNotes + Secs(tN, Timer)
End Sub

' the line's left / right text, straight from the data
Private Function SlotTexts(ByVal keyTxt As String, ByRef lt As String, ByRef rt As String) As Boolean
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
Private Function PostIdOfLine(ByVal keyTxt As String) As String
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



' ====================================================================
'  NEEDS ATTENTION TAB  (admin only - double-click any row for options)
' ====================================================================
Private Function NeedsSheet() As Worksheet
    If Not SheetExists(NA_SHEET) Then
        Dim ws As Worksheet
        Dim cs As Worksheet: Set cs = CalendarSheet()  ' v25.1 review: before the unlock (its SetTabColor relocks the tabs)
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=cs)
        ws.Name = NA_SHEET
        RelockStructure
    End If
    Set NeedsSheet = ThisWorkbook.Worksheets(NA_SHEET)
    SetTabColor NeedsSheet, thOpen()                ' v25.1
End Function

Public Sub ShowNeedsAttention()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    CalendarSheet
    Dim ws As Worksheet: Set ws = NeedsSheet()
    SetSheetVisible ws, xlSheetVisible
    ws.Activate                          ' the tab-switch check runs here (ThisWorkbook)
    If NAStale() Then DrawNeeds
    Exit Sub
Oops:
    ReportError "ShowNeedsAttention"
End Sub

Private Function NAGroup() As String
    NAGroup = SafeStr(MapSheet().Range("Z5").Value)
    If NAGroup = "" Then NAGroup = "All units"
End Function

Private Function NADays() As Long
    NADays = CLng(Val(MapSheet().Range("Z6").Value))
    If NADays <= 0 Then NADays = 14
End Function

Private Function NARanges() As Variant
    NARanges = Array(7, 14, 28, 56)
End Function

Public Sub NAUnitChanged()
    On Error GoTo Oops
    Dim dd As Object
    On Error Resume Next
    Set dd = NeedsSheet().DropDowns("na_unit")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z5", dd.List(dd.ListIndex)
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NAUnitChanged"
End Sub

Public Sub NARangeChanged()
    On Error GoTo Oops
    Dim dd As Object
    On Error Resume Next
    Set dd = NeedsSheet().DropDowns("na_range")
    On Error GoTo Oops
    If dd Is Nothing Then Exit Sub
    If dd.ListIndex < 1 Then Exit Sub
    MapWrite "Z6", NARanges()(dd.ListIndex - 1)
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NARangeChanged"
End Sub

Public Sub NARefresh()
    On Error GoTo Oops
    DrawNeeds
    Exit Sub
Oops:
    ReportError "NARefresh"
End Sub

Public Sub NAGoCalendar()
    On Error GoTo Oops
    ShowCalendar
    Exit Sub
Oops:
    ReportError "NAGoCalendar"
End Sub

' "Today", "Tomorrow", "in 5 days", "3 days ago"
Private Function WhenText(ByVal d As Date) As String
    Dim n As Long: n = Int(d) - Int(Date)
    Select Case n
        Case 0: WhenText = "Today"
        Case 1: WhenText = "Tomorrow"
        Case -1: WhenText = "Yesterday"
        Case Is > 1: WhenText = "in " & n & " days"
        Case Else: WhenText = -n & " days ago"
    End Select
End Function

Private Function WhenColor(ByVal d As Date) As Long
    Select Case Int(d) - Int(Date)
        Case Is <= 0: WhenColor = thOpen()
        Case 1: WhenColor = thPosted()
        Case Else: WhenColor = thMuted()
    End Select
End Function

' tblPostingList snapshot (for "n signed up")
Private Function PostListData() As Variant
    On Error Resume Next
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(PL_TABLE)
        If Not lo Is Nothing Then
            PostListData = TableData(lo)
            Exit Function
        End If
    Next ws
End Function

Private Function SignedCount(ByVal plV As Variant, ByVal id As String) As Long
    If Not IsArray(plV) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(plV, 1)
        If StrComp(SafeStr(plV(i, PL_ID)), id, vbTextCompare) = 0 Then
            If UCase$(SafeStr(plV(i, PL_SIGNED))) = "YES" Then SignedCount = SignedCount + 1
        End If
    Next i
End Function

' unit behind a calendar line key
Private Function SlotUnit(ByVal keyTxt As String) As String
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

Private Sub NAAdd(ByRef recs() As NARow, ByRef n As Long, ByVal c0 As String, ByVal c1 As String, ByVal c2 As String, _
                  ByVal c3 As String, ByVal c4 As String, ByVal c5 As String, ByVal c6 As String, ByVal c7 As String, _
                  ByVal keyTxt As String, ByVal sortKey As Double, ByVal accent As Long, ByVal stCol As Long, _
                  ByVal stFill As Long, ByVal whenCol As Long)
    n = n + 1
    If n > UBound(recs) Then ReDim Preserve recs(1 To n + 50)
    With recs(n)
        .Cell(0) = c0: .Cell(1) = c1: .Cell(2) = c2: .Cell(3) = c3
        .Cell(4) = c4: .Cell(5) = c5: .Cell(6) = c6: .Cell(7) = c7
        .Key = keyTxt
        .SortKey = sortKey
        .Accent = accent
        .StatusCol = stCol
        .StatusFill = stFill
        .WhenCol = whenCol
    End With
End Sub

Private Sub NASort(ByRef recs() As NARow, ByVal n As Long)
    Dim i As Long, j As Long, t As NARow
    For i = 2 To n
        t = recs(i)
        j = i - 1
        Do While j >= 1
            If recs(j).SortKey <= t.SortKey Then Exit Do
            recs(j + 1) = recs(j)
            j = j - 1
        Loop
        recs(j + 1) = t
    Next i
End Sub

Public Sub DrawNeeds()
    Dim ws As Worksheet: Set ws = NeedsSheet()
    Dim grp As String: grp = NAGroup()
    Dim nDays As Long: nDays = NADays()
    Dim drawnFrom As String: drawnFrom = SafeStr(MapSheet().Range("Z7").Value)
    Dim t0 As Double, tA As Double, tB As Double, tC As Double    ' v25.1.2: draw-time measurement (CalMap Z15)
    t0 = Timer
    mTimeSlots = 0
    mDrawing = True
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim calcWas As Long: calcWas = Application.Calculation
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    On Error GoTo Failed
    UnlockSheet ws
    LoadCache
    LoadPostCache
    Dim plV As Variant: plV = PostListData()

    ' ---- gather ----
    Dim pr() As NARow, nP As Long, lr() As NARow, nL As Long, sr() As NARow, nS As Long, br() As NARow, nB As Long
    ReDim pr(1 To 50): ReDim lr(1 To 50): ReDim sr(1 To 50): ReDim br(1 To 50)
    Dim i As Long, j As Long, m As Long, cl As Date, sd As Date, d As Date, sh As Variant, rt As String
    Dim nReady As Long, tier As Long, st As String, stC As Long, stF As Long, id As String, txt As String

    ' A. postings still open
    If IsArray(mPostV) Then
        For i = 1 To UBound(mPostV, 1)
            If SafeStr(mPostV(i, PS_STATUS)) = "Open" And UnitInGroup(SafeStr(mPostV(i, PS_UNIT)), grp) Then
                id = SafeStr(mPostV(i, PS_ID))
                cl = AsDate(mPostV(i, PS_CLOSES))
                sd = AsDate(mPostV(i, PS_DATE))
                If cl > 0 And PostingClosed(cl) Then          ' v25 review: the one closing test
                    tier = 0: st = "READY TO AWARD": stC = thOpen(): stF = thOpenFill(): nReady = nReady + 1
                ElseIf cl > 0 And Int(cl) = Date Then
                    tier = 1: st = "Closes today " & POSTING_CLOSE_TIME: stC = thPosted(): stF = thPostedFill()
                ElseIf cl > 0 Then
                    tier = 2: st = "Open until " & Format$(cl, "ddd mmm d"): stC = thMuted(): stF = -1
                Else
                    tier = 3: st = "Open (no closing date)": stC = thMuted(): stF = -1
                End If
                txt = id & "  " & SafeStr(mPostV(i, PS_REASON)) & "   (" & SafeStr(mPostV(i, PS_HOURS)) & "h, " & _
                      SafeStr(mPostV(i, PS_POS)) & " pos, " & SignedCount(plV, id) & " signed up)"
                If Not CovHasSourceCached(id) Then
                    txt = txt & "   - NOT ON CALENDAR"
                    If tier > 0 Then stC = thOpen()
                End If
                NAAdd pr, nP, Format$(sd, "ddd mmm d"), SafeStr(mPostV(i, PS_SHIFT)), SafeStr(mPostV(i, PS_CREW)), _
                      SafeStr(mPostV(i, PS_UNIT)), txt, TypeLong(ReasonType(SafeStr(mPostV(i, PS_REASON)))), st, _
                      IIf(cl > 0, WhenText(cl), ""), "P:" & id, tier * 1000000# + CDbl(Int(cl)), thPosted(), stC, stF, _
                      IIf(cl > 0, WhenColor(cl), thMuted())
            End If
        Next i
        NASort pr, nP
    End If

    ' B. open lines in the window
    Dim slots() As CalSlot, unit As String
    For i = 0 To nDays - 1
        d = Date + i
        For Each sh In Array("Days", "Nights")
            m = SlotsFor(d, CStr(sh), grp, slots, True)
            For j = 1 To m
                rt = UCase$(slots(j).RightText)
                If slots(j).IsOpen Or rt Like "NO TAKERS*" Then
                    unit = SlotUnit(slots(j).Key)
                    NAAdd lr, nL, Format$(d, "ddd mmm d"), CStr(sh), IIf(sh = "Days", DaysCrew(d), NightsCrew(d)), unit, _
                          slots(j).LeftText, TypeLong(slots(j).TypeCode), IIf(rt Like "NO TAKERS*", "NO TAKERS", "OPEN"), _
                          WhenText(d), slots(j).Key & "|" & CLng(Int(d)) & "|" & sh, 0, thOpen(), thOpen(), thOpenFill(), WhenColor(d)
                End If
            Next j
        Next sh
    Next i

    ' C. shift switch paybacks still owed
    Dim ow As Variant, owIds As Variant, swD As Date, swSh As String, swU As String, swOff As String, swWorks As String
    On Error Resume Next
    ow = OwedSwitches(owIds)
    On Error GoTo Failed
    If IsArray(owIds) Then
        For i = LBound(owIds) To UBound(owIds)
            id = CStr(owIds(i))
            If SwitchInfo(id, swD, swSh, swU, swOff, swWorks) Then
                If UnitInGroup(swU, grp) Then
                    NAAdd sr, nS, Format$(swD, "ddd mmm d"), swSh, IIf(swSh = "Days", DaysCrew(swD), NightsCrew(swD)), swU, _
                          swOff & " owes " & swWorks & " a shift   (" & id & ")", "Shift switch", "Payback owed", _
                          WhenText(swD), "S:" & id, CDbl(Int(swD)), TypeColor("SS"), TypeColor("SS"), Tint(TypeColor("SS"), 0.85), thMuted()
                End If
            End If
        Next i
    End If

    ' D. leave returning soon, and anything open-ended
    Dim code As String, eD As Date, nOpenEnded As Long, empNm As String
    If IsArray(mAbs) Then
        For i = 1 To UBound(mAbs, 1)
            If UCase$(SafeStr(mAbs(i, AB_STATUS))) <> "VOID" And IsDate(mAbs(i, AB_START)) Then
                If UnitInGroup(SafeStr(mAbs(i, AB_UNIT)), grp) Then
                    code = TypeCode(SafeStr(mAbs(i, AB_TYPE)))
                    sd = CDate(mAbs(i, AB_START))
                    st = ""
                    If Not IsDate(mAbs(i, AB_END)) Then
                        If Int(sd) <= Date + nDays - 1 Then
                            st = "Open-ended - no return date": eD = sd: nOpenEnded = nOpenEnded + 1
                        End If
                    ElseIf IsLeaveCode(code) Then
                        eD = CDate(mAbs(i, AB_END)) + 1
                        If Int(eD) >= Date And Int(eD) <= Date + nDays - 1 Then st = "Returns " & Format$(eD, "ddd mmm d")
                    End If
                    If st <> "" Then
                        empNm = SafeStr(mAbs(i, AB_EMP))
                        NAAdd br, nB, Format$(sd, "ddd mmm d"), "", EmpCrewOf(empNm), SafeStr(mAbs(i, AB_UNIT)), _
                              CalName(empNm), TypeLong(code), st, WhenText(eD), "B:" & SafeStr(mAbs(i, AB_ID)), _
                              IIf(Left$(st, 4) = "Open", 0, 1000000#) + CDbl(Int(eD)), TypeColor(code), TypeColor(code), _
                              Tint(TypeColor(code), 0.85), thMuted()
                    End If
                End If
            End If
        Next i
        NASort br, nB
    End If

    ' E. weekend holidays still waiting for the weekday the plant gives off (v20)
    Dim hr() As NARow, nH As Long, hv As Variant, hk As Variant, hd As Date, hLook As Long
    ReDim hr(1 To 20)
    hLook = nDays: If hLook < HOL_CONFIRM_DAYS Then hLook = HOL_CONFIRM_DAYS
    hv = HolidaysToConfirm(Date, Date + hLook - 1)
    If IsArray(hv) Then
        For Each hk In hv
            hd = DateSerial(CLng(Left$(hk, 4)), CLng(Mid$(hk, 6, 2)), CLng(Mid$(hk, 9, 2)))
            NAAdd hr, nH, Format$(hd, "ddd mmm d"), "", "DAYS", "All units", Mid$(hk, 12), "Holiday", _
                  "Set the weekday off", WhenText(hd), "H:" & Format$(hd, "yyyy-mm-dd"), CDbl(hd), _
                  RGB(176, 0, 32), RGB(176, 0, 32), Tint(RGB(176, 0, 32), 0.88), thMuted()
        Next hk
    End If

    ' H. imported vacation lines not yet booked (v22) - they don't keep people off call lists
    Dim ir() As NARow, nI As Long, iv As Variant
    ReDim ir(1 To 20)
    On Error Resume Next
    iv = ImportCandidates(Date, Date + IMPORT_WATCH_DAYS - 1)
    On Error GoTo Failed
    If IsArray(iv) Then
        For i = 1 To UBound(iv, 1)
            If UnitInGroup(SafeStr(iv(i, 4)), grp) Then
                NAAdd ir, nI, Format$(Dt2(iv(i, 2)), "ddd mmm d"), SafeStr(iv(i, 3)), "", SafeStr(iv(i, 4)), _
                      SafeStr(iv(i, 5)) & IIf(SafeStr(iv(i, 6)) = "", "   (no roster match)", IIf(SafeStr(iv(i, 6)) = "?", "   (several matches)", "   = " & SafeStr(iv(i, 6)))), _
                      "Imported", "Not booked", WhenText(Dt2(iv(i, 2))), "I:" & SafeStr(iv(i, 1)), CDbl(Int(Dt2(iv(i, 2)))), _
                      TypeColor("VAC"), thOpen(), thOpenFill(), WhenColor(Dt2(iv(i, 2)))
            End If
        Next i
        NASort ir, nI
    End If

    ' G. members' file not published lately (M1)
    Dim mr() As NARow, nM As Long, pubAt As Date, pubTxt As String
    ReDim mr(1 To 2)
    pubAt = MembersPublishedAt()
    If pubAt = 0 Then
        pubTxt = "never published - click Member View on the roster"
    ElseIf Int(Date) - Int(pubAt) >= MEMBERS_STALE_DAYS Then
        pubTxt = "last published " & Format$(pubAt, "ddd mmm d, h:mm AM/PM") & " (" & (Int(Date) - Int(pubAt)) & " days ago)"
    End If
    If pubTxt <> "" Then
        NAAdd mr, nM, Format$(Date, "ddd mmm d"), "", "", "All units", "Members file: " & pubTxt, "Publish", _
              "Publish again", "", "M:publish", 0, thPosted(), thOpen(), thOpenFill(), thMuted()
    End If

    ' E. errors nobody has looked at yet (v24.2)
    Dim er() As NARow, nE As Long, nErrLines As Long
    ReDim er(1 To 2)
    nErrLines = UnseenErrorCount()
    If nErrLines > 0 Then
        NAAdd er, nE, Format$(Date, "ddd mmm d"), "", "", "All units", nErrLines & " new line(s) in the Error Log", "Error", _
              "Open the log", "", "X:errors", 0, RGB(176, 0, 32), thOpen(), thOpenFill(), thMuted()
    End If

    ' F. people with no Home Unit (v21) - Graeme fills these in on the roster
    Dim ur() As NARow, nU As Long, uv As Variant, uk As Variant
    ReDim ur(1 To 20)
    On Error Resume Next
    uv = BlankHomeUnits()
    On Error GoTo Failed
    If IsArray(uv) Then
        For Each uk In uv
            NAAdd ur, nU, "", "", EmpCrewOf(CStr(uk)), "", CStr(uk), "Roster", "Home Unit blank", "", _
                  "U:" & CStr(uk), 0, thMuted(), thOpen(), thOpenFill(), thMuted()
        Next uk
    End If

    ' V. unit moves entered ahead of time (v24.3): waiting for their date, or due and not applied
    Dim vr() As NARow, nV As Long, pmv As Variant, pmD As Date, pmDue As Boolean, pmWho As String
    ReDim vr(1 To 10)
    On Error Resume Next
    pmv = PlannedRows()
    On Error GoTo Failed
    If IsArray(pmv) Then
        For i = 1 To UBound(pmv, 1)
            pmD = Dt2(pmv(i, PM_FROM))
            pmDue = (Int(pmD) <= Date)                  ' a date that can't be read (0) shows as a problem too
            pmWho = SafeStr(pmv(i, PM_EMP))
            NAAdd vr, nV, IIf(pmD = 0, "?", Format$(pmD, "ddd mmm d")), "", SafeStr(pmv(i, PM_CREW)), SafeStr(pmv(i, PM_UNIT)), _
                  CalName(pmWho) & "   " & HomeUnitOf(pmWho) & " -> " & SafeStr(pmv(i, PM_UNIT)) & " (" & SafeStr(pmv(i, PM_CODE)) & ")   " & _
                  SafeStr(pmv(i, PM_ID)), "Unit move", IIf(pmD = 0, "Date can't be read", IIf(pmDue, "DUE - not applied", "Planned")), _
                  IIf(pmD = 0, "", WhenText(pmD)), "V:" & SafeStr(pmv(i, PM_ID)), CDbl(Int(pmD)), thMuted(), _
                  IIf(pmDue, thOpen(), thMuted()), IIf(pmDue, thOpenFill(), -1), WhenColor(pmD)
        Next i
        NASort vr, nV
    End If

    ' T. attendance (v22.1): sick / personal days at or over the limit, by home unit
    Dim tr() As NARow, nT As Long, rn As Variant, ast As AttendanceStatus, lvl As Long, hu As String, nmT As String
    ReDim tr(1 To 20)
    On Error Resume Next
    rn = RosterNames()
    If IsArray(rn) Then
        BeginAttendanceBatch
        For i = LBound(rn) To UBound(rn)
            nmT = CStr(rn(i))
            hu = HomeUnitOf(nmT)
            If hu = "" Or UnitInGroup(hu, grp) Then
                If AttendanceFor(nmT, Date, ast) Then
                    lvl = AttendanceLevel(ast)
                    If lvl > 0 Then
                        NAAdd tr, nT, IIf(ast.LastSick > 0, Format$(ast.LastSick, "ddd mmm d"), ""), "", EmpCrewOf(nmT), hu, _
                              CalName(nmT) & "   " & ast.SickUsed & " sick / " & ast.PersonalUsed & " personal, " & ast.Occurrences & " occ.", _
                              "Attendance", AttendanceStatusText(ast), _
                              IIf(ast.WaitUntil > 0, "until " & Format$(ast.WaitUntil, "mmm d, yyyy"), ""), _
                              "T:" & nmT, CDbl(-lvl * 1000 - ast.Combined), TypeColor("SIC"), _
                              IIf(lvl = 2, thOpen(), thPosted()), IIf(lvl = 2, thOpenFill(), thPostedFill()), thMuted()
                    End If
                End If
            End If
        Next i
        EndAttendanceBatch
        NASort tr, nT
    End If
    On Error GoTo Failed

    tA = Timer                                      ' gathered
    ' ---- draw ----
    ws.Rows("1:3000").Delete
    With ws.Cells.Font
        .Name = CAL_FONT
        .Size = 9
    End With
    ws.Cells.Interior.Color = vbWhite
    ws.Columns(1).ColumnWidth = 1.5
    ws.Columns(2).ColumnWidth = 13
    ws.Columns(3).ColumnWidth = 8
    ws.Columns(4).ColumnWidth = 7
    ws.Columns(5).ColumnWidth = 17
    ws.Columns(6).ColumnWidth = 46
    ws.Columns(7).ColumnWidth = 16
    ws.Columns(8).ColumnWidth = 26
    ws.Columns(9).ColumnWidth = 13
    ws.Columns(10).ColumnWidth = 2
    ws.Range(ws.Columns(NA_KEY), ws.Columns(NA_KEY + 12)).Hidden = True

    ws.Rows(1).RowHeight = 38
    ws.Range("A1:J1").Interior.Color = thBand()
    With ws.Range("B1:E1")
        .Merge
        .Value = "NEEDS ATTENTION"
        .Font.Name = CAL_FONT_B: .Font.Size = 15: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 22
    With ws.Range("B2:I2")
        .Merge
        .Value = grp & "   " & ChrW(183) & "   " & Format$(Date, "ddd mmm d") & " to " & Format$(Date + nDays - 1, "ddd mmm d") & _
                 "   " & ChrW(183) & "   updated " & Format$(Now, "h:mm AM/PM") & "   " & ChrW(183) & "   double-click any row for options" & _
                 IIf(RolloverNote() <> "", "   " & ChrW(183) & "   " & RolloverNote(), "")
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = thMuted()
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(3).RowHeight = 8
    ws.Rows(4).RowHeight = 30
    ws.Rows(5).RowHeight = 16
    ws.Rows(6).RowHeight = 12

    Dim r As Long: r = 7
    Dim rowP As Long, rowL As Long, rowS As Long, rowB As Long
    rowL = r
    r = NASection(ws, r, "OPEN LINES  -  need cover in the next " & nDays & " days", nL, thOpen(), _
                  Array("Date", "Shift", "Crew", "Unit", "Who's off / need", "Type", "Status", "When"))
    r = NARows(ws, r, lr, nL, "Nothing open - every line in the window is covered, posted or NCN.")
    rowP = r
    r = NASection(ws, r, "POSTINGS  -  open postings, oldest closing first", nP, thPosted(), _
                  Array("Shift date", "Shift", "Crew", "Unit", "Posting", "Type", "Status", "Closes"))
    r = NARows(ws, r, pr, nP, "No open postings.")
    rowS = r
    r = NASection(ws, r, "SHIFT SWITCH PAYBACKS OWED", nS, TypeColor("SS"), _
                  Array("Switch date", "Shift", "Crew", "Unit", "Owed", "Type", "Status", "When"))
    r = NARows(ws, r, sr, nS, "No paybacks owed.")
    rowB = r
    r = NASection(ws, r, "LEAVE  -  returning soon or open-ended", nB, TypeColor("LOA"), _
                  Array("Started", "", "Crew", "Unit", "Who", "Type", "Status", "When"))
    r = NARows(ws, r, br, nB, "Nobody returning from leave in the window, and nothing open-ended.")
    If nH > 0 Then
        r = NASection(ws, r, "HOLIDAYS TO CONFIRM  -  weekend holiday, weekday off not entered (next " & hLook & " days)", _
                      nH, RGB(176, 0, 32), Array("Holiday date", "", "Crew", "Unit", "Holiday", "Type", "Status", "When"))
        r = NARows(ws, r, hr, nH, "")
    End If
    If nU > 0 Then
        r = NASection(ws, r, "HOME UNIT NOT SET  -  pick it from the drop-down in column B of the roster (double-click a name)", _
                      nU, thMuted(), Array("", "", "Crew", "", "Who", "Type", "Status", ""))
        r = NARows(ws, r, ur, nU, "")
    End If
    If nE > 0 Then
        r = NASection(ws, r, "ERRORS  -  something didn't finish (double-click to open the Error Log)", _
                      nE, RGB(176, 0, 32), Array("Today", "", "", "", "What", "Type", "Status", ""))
        r = NARows(ws, r, er, nE, "")
    End If
    If nM > 0 Then
        r = NASection(ws, r, "MEMBERS FILE  -  what the union sees is out of date (double-click to publish)", _
                      nM, thPosted(), Array("Today", "", "", "", "Members file", "Type", "Status", ""))
        r = NARows(ws, r, mr, nM, "")
    End If
    If nI > 0 Then
        r = NASection(ws, r, "IMPORTED, NOT BOOKED  -  old-calendar vacations that don't keep people off call lists yet (double-click to convert)", _
                      nI, TypeColor("VAC"), Array("Date", "Shift", "", "Unit", "As imported", "Type", "Status", "When"))
        r = NARows(ws, r, ir, nI, "")
    End If
    If nV > 0 Then
        r = NASection(ws, r, "PLANNED UNIT MOVES  -  applied when an admin opens the file on or after the date (double-click to cancel)", _
                      nV, thMuted(), Array("Effective", "", "New crew", "New unit", "Who", "Type", "Status", "When"))
        r = NARows(ws, r, vr, nV, "")
    End If
    If nT > 0 Then
        r = NASection(ws, r, "ATTENDANCE  -  sick / personal days at or over the limit (double-click for the Attendance sheet)", _
                      nT, TypeColor("SIC"), Array("Last sick", "", "Crew", "Home unit", "Who", "Type", "Status", "Until"))
        r = NARows(ws, r, tr, nT, "")
    End If

    tB = Timer                                      ' rows drawn
    ' ---- summary tiles (double-click one to jump to its list) ----
    NATile ws, 2, 3, nL, "OPEN LINES", thOpen(), rowL
    NATile ws, 4, 5, nReady, "READY TO AWARD", thPosted(), rowP
    NATile ws, 6, 6, nP, "OPEN POSTINGS", TypeColor("VAC"), rowP
    NATile ws, 7, 7, nS, "PAYBACKS OWED", TypeColor("SS"), rowS
    NATile ws, 8, 9, nB, "LEAVE TO WATCH", TypeColor("LOA"), rowB

    BuildNAControls ws, grp, nDays

    On Error Resume Next
    If ActiveSheet Is ws Then
        With ActiveWindow
            .DisplayGridlines = False
            .DisplayHeadings = False
            .FreezePanes = False
            .ScrollRow = 1
            .SplitColumn = 0
            .SplitRow = 2
            .FreezePanes = True
        End With
    End If
    On Error GoTo Failed
    MapWrite "Z9", drawnFrom                        ' v25.1: through the pair
    MapWrite "Z11", "'" & DrawContext()
    tC = Timer                                      ' v25.1.2: tiles drawn
    MapWrite "Z15", "'" & "Needs Attention " & grp & ", " & nDays & " days (" & (nP + nL + nS + nB + nH + nU + nE + nM + nI + nV + nT) & " rows): drawn in " & _
        Format$(Secs(t0, tC), "0.0") & " s = gather " & Format$(Secs(t0, tA), "0.0") & " (slots " & Format$(mTimeSlots, "0.0") & _
        "), rows " & Format$(Secs(tA, tB), "0.0") & ", tiles " & Format$(Secs(tB, tC), "0.0") & "  [" & CAL_VERSION & "]"
    GoTo Done

Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description & " (error " & Err.Number & ")"
    Resume AfterFail
AfterFail:
    On Error Resume Next
    LogError "DrawNeeds", en, em          ' v24.3.1
    UnlockSheet ws
    ws.Cells(3, 2).Value = "Couldn't finish: " & em
    ws.Cells(3, 2).Font.Color = thOpen()
    BuildNAControls ws, grp, nDays
    Application.ScreenUpdating = True
    MsgBox "The Needs Attention tab couldn't finish drawing:" & vbNewLine & vbNewLine & em & vbNewLine & vbNewLine & _
           "Send a screenshot of this message.  (" & CAL_VERSION & ")", vbExclamation, "Needs Attention"
Done:
    On Error Resume Next
    ProtectSheet ws
    Application.Calculation = calcWas
    Application.EnableEvents = evWas
    Application.ScreenUpdating = True
    mDrawing = False
End Sub

' a live calendar line came from this source (uses the cached Coverage table)
Private Function CovHasSourceCached(ByVal source As String) As Boolean
    If Not IsArray(mCov) Then Exit Function
    Dim c As Long
    For c = 1 To UBound(mCov, 1)
        If StrComp(SafeStr(mCov(c, CV_SOURCE)), source, vbTextCompare) = 0 And UCase$(SafeStr(mCov(c, CV_STATUS))) <> "VOID" Then
            CovHasSourceCached = True
            Exit Function
        End If
    Next c
End Function

Private Function EmpCrewOf(ByVal empName As String) As String
    If mShiftMap Is Nothing Then Exit Function
    If mShiftMap.Exists(Trim$(empName)) Then EmpCrewOf = CrewOn(empName, Date, CStr(mShiftMap(Trim$(empName))))   ' v24: today's crew
End Function

' coloured bar + column headings; returns the first row for data
Private Function NASection(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, ByVal n As Long, _
                           ByVal accent As Long, ByVal heads As Variant) As Long
    ws.Rows(r).RowHeight = 24
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Interior.Color = accent
    With ws.Cells(r, 2)
        .Value = title & "     " & n
        .Font.Name = CAL_FONT_B: .Font.Size = 10.5: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
    End With
    Dim i As Long
    ws.Rows(r + 1).RowHeight = 18
    For i = 0 To 7
        With ws.Cells(r + 1, 2 + i)
            .Value = heads(i)
            .Font.Name = CAL_FONT_B: .Font.Size = 8: .Font.Color = thMuted()
            .Interior.Color = RGB(245, 247, 248)
            .VerticalAlignment = xlCenter
            .IndentLevel = 1
        End With
    Next i
    ws.Cells(r + 1, 1).Interior.Color = RGB(245, 247, 248)
    With ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 1, 9)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = thGrid()
    End With
    NASection = r + 2
End Function

' the rows of one section (zebra, status pill, hover notes); returns the next free row
Private Function NARows(ByVal ws As Worksheet, ByVal r As Long, ByRef recs() As NARow, ByVal n As Long, _
                        ByVal emptyMsg As String) As Long
    Dim i As Long, c As Long
    If n = 0 Then
        ws.Rows(r).RowHeight = 20
        With ws.Cells(r, 2)
            .Value = emptyMsg
            .Font.Italic = True: .Font.Color = thGreen()
            .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
        ws.Rows(r + 1).RowHeight = 14
        NARows = r + 2
        Exit Function
    End If
    For i = 1 To n
        ws.Rows(r).RowHeight = 20
        If i Mod 2 = 0 Then ws.Range(ws.Cells(r, 2), ws.Cells(r, 9)).Interior.Color = RGB(250, 250, 251)
        ws.Cells(r, 1).Interior.Color = recs(i).Accent
        For c = 0 To 7
            With ws.Cells(r, 2 + c)
                .Value = "'" & recs(i).Cell(c)
                .VerticalAlignment = xlCenter
                .IndentLevel = 1
                .ShrinkToFit = True
            End With
        Next c
        ws.Cells(r, 6).Font.Name = CAL_FONT_B
        With ws.Cells(r, 8)
            .Font.Name = CAL_FONT_B: .Font.Color = recs(i).StatusCol
            If recs(i).StatusFill >= 0 Then .Interior.Color = recs(i).StatusFill
        End With
        With ws.Cells(r, 9)
            .Font.Color = recs(i).WhenCol
            .Font.Name = IIf(recs(i).WhenCol = thMuted(), CAL_FONT, CAL_FONT_B)
        End With
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = thGrid()
        End With
        ws.Cells(r, NA_KEY).Value = recs(i).Key
        If Left$(recs(i).Key, 2) = "A:" Or Left$(recs(i).Key, 2) = "C:" Then AddNote ws.Cells(r, 6), LineDetails(recs(i).Key, True)
        r = r + 1
    Next i
    ws.Rows(r).RowHeight = 14
    NARows = r + 1
End Function

' summary tile over columns c1..c2, rows 4-5
Private Sub NATile(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal n As Long, ByVal cap As String, _
                   ByVal col As Long, ByVal jumpRow As Long)
    With ws.Range(ws.Cells(4, c1), ws.Cells(4, c2))
        .Merge
        .Value = n
        .Font.Name = CAL_FONT_B: .Font.Size = 20: .Font.Color = IIf(n > 0, col, RGB(176, 190, 197))
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlBottom: .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(5, c1), ws.Cells(5, c2))
        .Merge
        .Value = cap
        .Font.Name = CAL_FONT_B: .Font.Size = 7.5: .Font.Color = thMuted()
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlTop: .IndentLevel = 1
    End With
    With ws.Range(ws.Cells(4, c1), ws.Cells(5, c2))
        .Interior.Color = Tint(col, 0.9)
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = col
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = vbWhite
        End With
    End With
    ' jump key for this tile (hidden columns to the right of the key column)
    ws.Cells(4, NA_KEY + c1).Value = "J:" & jumpRow
    ws.Cells(5, NA_KEY + c1).Value = "J:" & jumpRow
End Sub

Private Sub BuildNAControls(ByVal ws As Worksheet, ByVal grp As String, ByVal nDays As Long)
    On Error Resume Next
    UnlockSheet ws
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "na_" Then ws.Shapes(i).Delete
    Next i
    CalDrop ws, "na_unit", ws.Range("F1")
    CalDrop ws, "na_range", ws.Range("G1")
    CalBtn ws, "na_refresh", "Refresh", "modCalendar.NARefresh", ws.Range("H1"), thNav()
    CalBtn ws, "na_cal", "Calendar", "modCalendar.NAGoCalendar", ws.Range("I1"), thGreen()

    Dim dd As Object, g As Variant, idx As Long, rg As Variant
    Set dd = Nothing
    Set dd = ws.DropDowns("na_unit")
    If Not dd Is Nothing Then
        g = CalGroupNames()
        For i = 0 To UBound(g)
            dd.AddItem g(i)
            If StrComp(g(i), grp, vbTextCompare) = 0 Then idx = i + 1
        Next i
        If idx = 0 Then idx = UBound(g) + 1
        dd.ListIndex = idx
        dd.OnAction = "modCalendar.NAUnitChanged"
    End If
    Set dd = Nothing
    Set dd = ws.DropDowns("na_range")
    If Not dd Is Nothing Then
        rg = NARanges()
        idx = 2
        For i = 0 To UBound(rg)
            dd.AddItem "Next " & rg(i) & " days"
            If rg(i) = nDays Then idx = i + 1
        Next i
        dd.ListIndex = idx
        dd.OnAction = "modCalendar.NARangeChanged"
    End If
End Sub


' ====================================================================
'  MENUS  (double-click / right-click on the Calendar or Needs Attention)
'  Keys:  A:AB-0001|serial|Days   a time-off line
'         C:CV-0001|serial|Days   a need / coverage line
'         E:NEW|serial|Days       an empty spot under a date
'         D:DAY|serial|           a date header
'         P:P-0001                a posting
'         S:SS-0001               a shift switch owed a payback
'         B:AB-0001               a time-off booking
' ====================================================================
Public Sub CalDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    On Error Resume Next
    If Target.Worksheet.Name <> CAL_SHEET Then Exit Sub
    Cancel = True                                  ' never "this cell is protected"
    Dim keyTxt As String: keyTxt = SafeStr(MapSheet().Range(Target.Cells(1, 1).Address).Value)
    On Error GoTo 0
    If keyTxt <> "" Then ShowLineMenu keyTxt
End Sub

Public Sub CalRightClick(ByVal Target As Range, ByRef Cancel As Boolean)
    On Error Resume Next
    If Target.Worksheet.Name <> CAL_SHEET Then Exit Sub
    Dim keyTxt As String: keyTxt = SafeStr(MapSheet().Range(Target.Cells(1, 1).Address).Value)
    On Error GoTo 0
    If keyTxt = "" Then Exit Sub                  ' ordinary right-click elsewhere
    Cancel = True
    ShowLineMenu keyTxt
End Sub

Private Function NAKeyAt(ByVal Target As Range) As String
    On Error Resume Next
    Dim ws As Worksheet: Set ws = Target.Worksheet
    Dim r As Long: r = Target.Cells(1, 1).Row
    If r = 4 Or r = 5 Then
        ' summary tiles: find the tile's first column
        Dim c As Long
        For c = Target.Cells(1, 1).Column To 2 Step -1
            If SafeStr(ws.Cells(r, NA_KEY + c).Value) <> "" Then NAKeyAt = SafeStr(ws.Cells(r, NA_KEY + c).Value): Exit Function
        Next c
    End If
    NAKeyAt = SafeStr(ws.Cells(r, NA_KEY).Value)
End Function

Public Sub NADoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Cancel = True
    Dim keyTxt As String: keyTxt = NAKeyAt(Target)
    If keyTxt = "" Then Exit Sub
    If Left$(keyTxt, 2) = "J:" Then
        On Error Resume Next
        Application.Goto Target.Worksheet.Cells(CLng(Mid$(keyTxt, 3)), 1), True
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "H:" Then              ' holiday to confirm: open its row on the Holidays sheet
        GoToHoliday DateSerial(CLng(Mid$(keyTxt, 3, 4)), CLng(Mid$(keyTxt, 8, 2)), CLng(Mid$(keyTxt, 11, 2)))
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "U:" Then              ' v21: Home Unit blank - jump to that cell on the roster
        GoToHomeUnit Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "M:" Then              ' M1: publish the members' file
        If MsgBox("Publish the members' file now?", vbYesNo + vbQuestion, "Member View") = vbYes Then PublishMemberView
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "X:" Then              ' v24.2: errors - open the Error Log
        ShowErrorLog
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "V:" Then              ' v24.3: a planned unit move - cancel it, or apply it when due
        PlannedMoveMenu Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "I:" Then              ' v22: imported line - open the convert pop-up
        ShowConvertForm
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "T:" Then              ' v22.1: attendance - open the sheet
        ShowAttendance
        Exit Sub
    End If
    ShowLineMenu keyTxt
End Sub

Public Sub NARightClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Dim keyTxt As String: keyTxt = NAKeyAt(Target)
    If keyTxt = "" Or Left$(keyTxt, 2) = "J:" Then Exit Sub
    Cancel = True
    If Left$(keyTxt, 2) = "H:" Then
        GoToHoliday DateSerial(CLng(Mid$(keyTxt, 3, 4)), CLng(Mid$(keyTxt, 8, 2)), CLng(Mid$(keyTxt, 11, 2)))
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "U:" Then
        GoToHomeUnit Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "M:" Then
        If MsgBox("Publish the members' file now?", vbYesNo + vbQuestion, "Member View") = vbYes Then PublishMemberView
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "X:" Then
        ShowErrorLog
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "V:" Then
        PlannedMoveMenu Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "I:" Then
        ShowConvertForm
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "T:" Then
        ShowAttendance
        Exit Sub
    End If
    ShowLineMenu keyTxt
End Sub

' a Variant that may be a date serial or a date -> Date (0 if neither)
Private Function Dt2(ByVal x As Variant) As Date
    If IsDate(x) Then
        Dt2 = CDate(x)
    ElseIf IsNumeric(x) And Not IsEmpty(x) Then
        Dt2 = CDate(CDbl(x))
    End If
End Function

Private Function KeyDate(ByVal keyTxt As String) As Date
    Dim p As Variant: p = Split(keyTxt, "|")
    If UBound(p) >= 1 Then
        If IsNumeric(p(1)) Then KeyDate = CDate(CLng(p(1)))
    End If
End Function

Private Function KeyShift(ByVal keyTxt As String) As String
    Dim p As Variant: p = Split(keyTxt, "|")
    If UBound(p) >= 2 Then KeyShift = CStr(p(2))
End Function

Private Sub ShowLineMenu(ByVal keyTxt As String)
    mMenuKey = keyTxt
    LoadCache
    LoadPostCache
    Dim kind As String: kind = Left$(keyTxt, 2)

    ' everyone else just sees the details
    If Not gAdminUnlocked Then
        If kind = "E:" Or kind = "D:" Then CalMenuDayDetails Else CalMenuDetails
        Exit Sub
    End If

    Dim title As String, lt As String, rt As String, postId As String, d As Date, sh As String
    Dim caps() As String, macs() As String, grpBreak() As Boolean, n As Long
    ReDim caps(1 To 16): ReDim macs(1 To 16): ReDim grpBreak(1 To 16)
    d = KeyDate(keyTxt)
    sh = KeyShift(keyTxt)

    Select Case kind
        Case "A:", "C:"
            If Not SlotTexts(keyTxt, lt, rt) Then
                MsgBox "That line has changed - redrawing.", vbInformation, "Calendar"
                RefreshCalendar
                Exit Sub
            End If
            title = Format$(d, "ddd mmm d") & " " & sh & "  -  " & Clip(lt, 34)
            postId = PostIdOfLine(keyTxt)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If Not (UCase$(rt) Like "POSTED*") Then AddChoice caps, macs, grpBreak, n, "Post this line...", "CalMenuPost", True
            AddChoice caps, macs, grpBreak, n, "Set coverage...", "CalMenuSetCoverage", (UCase$(rt) Like "POSTED*")
            If kind = "C:" Then AddChoice caps, macs, grpBreak, n, "Change the text of this line...", "CalMenuEditText", False
            If postId <> "" Then
                If PostRow(postId) > 0 Then
                    If SafeStr(mPostV(PostRow(postId), PS_STATUS)) = "Open" Then
                        AddChoice caps, macs, grpBreak, n, "Award posting " & postId & "...", "CalMenuAward", True
                        AddChoice caps, macs, grpBreak, n, "Change the Reason of posting " & postId & "...", "CalMenuEditReason", False
                    End If
                End If
                AddChoice caps, macs, grpBreak, n, "View posting " & postId, "CalMenuViewPosting", False
            End If
            If kind = "A:" Then
                AddChoice caps, macs, grpBreak, n, "Change end date of this time off...", "CalMenuEndDate", True
                AddChoice caps, macs, grpBreak, n, "Void this time off...", "CalMenuVoid", False
            End If
            AddChoice caps, macs, grpBreak, n, "Post support on this shift...", "CalMenuPostSupport", True
            AddChoice caps, macs, grpBreak, n, "Add another need on this shift...", "CalMenuAddNeed", False
            If ActiveSheet.Name = NA_SHEET Then AddChoice caps, macs, grpBreak, n, "Show on the calendar", "CalMenuGoto", True

        Case "E:"
            title = Format$(d, "ddd mmm d") & " " & sh & "  (" & IIf(sh = "Days", DaysCrew(d), NightsCrew(d)) & " Shift)"
            AddChoice caps, macs, grpBreak, n, "Post unit operator support...", "CalMenuPostSupport", False
            AddChoice caps, macs, grpBreak, n, "Post PHA coverage...", "CalMenuPostPHA", False
            AddChoice caps, macs, grpBreak, n, "Add a need (no posting yet)...", "CalMenuAddNeed", True
            AddChoice caps, macs, grpBreak, n, "Book time off from this date...", "CalMenuBookFrom", False
            AddChoice caps, macs, grpBreak, n, "What's on this day", "CalMenuDayDetails", True

        Case "D:"
            title = Format$(d, "dddd mmm d") & "  (Days " & DaysCrew(d) & ", Nights " & NightsCrew(d) & ")"
            AddChoice caps, macs, grpBreak, n, "Post support - Days...", "CalMenuPostSupportD", False
            AddChoice caps, macs, grpBreak, n, "Post support - Nights...", "CalMenuPostSupportN", False
            AddChoice caps, macs, grpBreak, n, "Add a need (no posting yet)...", "CalMenuAddNeed", True
            AddChoice caps, macs, grpBreak, n, "Book time off from this date...", "CalMenuBookFrom", False
            AddChoice caps, macs, grpBreak, n, "What's on this day", "CalMenuDayDetails", True

        Case "P:"
            postId = Mid$(keyTxt, 3)
            title = "Posting " & postId
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If PostRow(postId) > 0 Then
                If SafeStr(mPostV(PostRow(postId), PS_STATUS)) = "Open" Then
                    AddChoice caps, macs, grpBreak, n, "Award posting " & postId & "...", "CalMenuAward", True
                    AddChoice caps, macs, grpBreak, n, "Change the Reason of posting " & postId & "...", "CalMenuEditReason", False
                End If
            End If
            AddChoice caps, macs, grpBreak, n, "View posting " & postId, "CalMenuViewPosting", False
            If Not CovHasSourceCached(postId) Then
                AddChoice caps, macs, grpBreak, n, "Put on the calendar...", "CalMenuLinkPosting", True
            End If

        Case "S:"
            title = "Shift switch " & Mid$(keyTxt, 3)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            AddChoice caps, macs, grpBreak, n, "Log the payback (Shift Switch)...", "CalMenuSwitch", True

        Case "B:"
            title = "Time off " & Mid$(keyTxt, 3)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If MenuAbsIsLeave() Then AddChoice caps, macs, grpBreak, n, "Return from leave now (top-up)...", "CalMenuLeaveReturn", True
            AddChoice caps, macs, grpBreak, n, "Change end date...", "CalMenuEndDate", Not MenuAbsIsLeave()
            AddChoice caps, macs, grpBreak, n, "Void this time off...", "CalMenuVoid", False

        Case Else
            Exit Sub
    End Select

    ' pop-up menu at the mouse
    Dim cb As Object, ctl As Object, i As Long
    On Error Resume Next
    Application.CommandBars(MENU_NAME).Delete
    On Error GoTo Fallback
    Set cb = Application.CommandBars.Add(Name:=MENU_NAME, Position:=5, Temporary:=True)   ' 5 = msoBarPopup
    Set ctl = cb.Controls.Add(Type:=1, Temporary:=True)                                    ' 1 = msoControlButton
    ctl.Caption = title
    ctl.Enabled = False
    For i = 1 To n
        Set ctl = cb.Controls.Add(Type:=1, Temporary:=True)
        ctl.Caption = caps(i)
        ctl.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macs(i)
        ctl.BeginGroup = (i = 1) Or grpBreak(i)
    Next i
    cb.ShowPopup
    Exit Sub

Fallback:
    ' no pop-up menus (Mac / restricted Excel) - numbered list instead
    Dim msg As String, pick As Long
    msg = title & vbNewLine & vbNewLine
    For i = 1 To n
        msg = msg & i & ")  " & caps(i) & vbNewLine
    Next i
    pick = AskNumber(msg & vbNewLine & "Type the number:", "Calendar", 1, 1, n)
    If pick >= 1 Then Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macs(pick)
End Sub

Private Sub AddChoice(ByRef caps() As String, ByRef macs() As String, ByRef grpBreak() As Boolean, _
                      ByRef n As Long, ByVal cap As String, ByVal mac As String, ByVal newGroup As Boolean)
    n = n + 1
    caps(n) = cap
    macs(n) = mac
    grpBreak(n) = newGroup
End Sub

Private Function MenuPostId() As String
    If Left$(mMenuKey, 2) = "P:" Then
        MenuPostId = Mid$(mMenuKey, 3)
    Else
        LoadCache
        MenuPostId = PostIdOfLine(mMenuKey)
    End If
End Function

Private Function MenuAbsId() As String
    If Left$(mMenuKey, 2) = "A:" Or Left$(mMenuKey, 2) = "B:" Then MenuAbsId = Mid$(Split(mMenuKey, "|")(0), 3)
End Function

' unit for a new posting / need started from the menu
Private Function MenuUnit() As String
    If Left$(mMenuKey, 2) = "A:" Or Left$(mMenuKey, 2) = "C:" Then
        LoadCache
        MenuUnit = SlotUnit(mMenuKey)
    Else
        MenuUnit = SingleUnitOf(CalGroup())
    End If
End Function

Public Sub CalMenuDetails()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    If mMenuKey = "" Then Exit Sub
    Dim t As String
    Select Case Left$(mMenuKey, 2)
        Case "P:": t = PostingDetails(Mid$(mMenuKey, 3))
        Case "S:": t = SwitchDetails(Mid$(mMenuKey, 3))
        Case "B:": t = BookingDetails(Mid$(mMenuKey, 3))
        Case Else: t = LineDetails(mMenuKey)
    End Select
    If t = "" Then t = "No details found - try Refresh."
    MsgBox Replace(t, vbLf, vbNewLine), vbInformation, "Details"
    Exit Sub
Oops:
    ReportError "CalMenuDetails"
End Sub

Public Sub CalMenuDayDetails()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim d As Date: d = KeyDate(mMenuKey)
    If d = 0 Then Exit Sub
    Dim grp As String: grp = CalGroup()
    Dim t As String, sh As Variant, slots() As CalSlot, n As Long, i As Long, rt As String
    t = Format$(d, "dddd mmm d") & "  -  " & grp
    For Each sh In Array("Days", "Nights")
        n = SlotsFor(d, CStr(sh), grp, slots)
        t = t & vbNewLine & vbNewLine & UCase$(sh) & "  (" & IIf(sh = "Days", DaysCrew(d), NightsCrew(d)) & " Shift)"
        If n = 0 Then t = t & vbNewLine & "   nothing booked"
        For i = 1 To n
            rt = slots(i).RightText
            If rt = "" Then rt = "OPEN"
            t = t & vbNewLine & "   " & slots(i).LeftText & "   ->   " & rt
        Next i
    Next sh
    MsgBox t, vbInformation, "Calendar"
    Exit Sub
Oops:
    ReportError "CalMenuDayDetails"
End Sub

Private Function SwitchDetails(ByVal id As String) As String
    Dim d As Date, sh As String, u As String, offN As String, worksN As String
    If Not SwitchInfo(id, d, sh, u, offN, worksN) Then SwitchDetails = "Shift switch " & id & " wasn't found.": Exit Function
    SwitchDetails = "Shift switch " & id & vbLf & _
                    worksN & " worked " & Format$(d, "ddd mmm d") & " " & sh & " (" & u & ") for " & offN & "." & vbLf & _
                    "Payback still owed: " & offN & " works a shift for " & worksN & "." & vbLf & vbLf & _
                    "Log it in Shift Switch - pick this switch in the owed list."
End Function

Private Function BookingDetails(ByVal id As String) As String
    LoadCache
    Dim r As Long: r = RowOf(mAbs, id)
    If r = 0 Then BookingDetails = "Time off " & id & " wasn't found.": Exit Function
    Dim code As String: code = TypeCode(SafeStr(mAbs(r, AB_TYPE)))
    Dim t As String
    t = CalName(SafeStr(mAbs(r, AB_EMP))) & "  -  " & TypeLong(code) & "  (" & SafeStr(mAbs(r, AB_UNIT)) & ")" & vbLf & _
        "From " & Format$(mAbs(r, AB_START), "ddd mmm d, yyyy") & " to " & _
        IIf(IsDate(mAbs(r, AB_END)), Format$(mAbs(r, AB_END), "ddd mmm d, yyyy"), "open-ended (no return date)")
    If HrsTag(mAbs(r, AB_HOURS)) <> "" Then t = t & vbLf & "Off" & HrsTag(mAbs(r, AB_HOURS)) & " per shift"
    If Not IsLeaveCode(code) And code <> "SIC" And SafeStr(mAbs(r, AB_NOTE)) <> "" Then t = t & vbLf & "Note: " & SafeStr(mAbs(r, AB_NOTE))
    t = t & vbLf & "Booked by " & StampTxt(mAbs(r, AB_USER), mAbs(r, AB_STAMP)) & "  [" & id & "]"
    BookingDetails = t
End Function

Public Sub CalMenuPost()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    CalPostKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuPost"
End Sub

Public Sub CalMenuSetCoverage()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    CalSetCoverageKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuSetCoverage"
End Sub

' v24.3.2: the words on a need line (the Reason a posting wrote, or a typed need)
Public Sub CalMenuEditText()
    On Error GoTo Oops
    CalEditTextKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuEditText"
End Sub

' v24.3.2: an open posting's Reason, from its line or its Needs Attention entry
Public Sub CalMenuEditReason()
    On Error GoTo Oops
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    CalEditReasonId id
    Exit Sub
Oops:
    ReportError "CalMenuEditReason"
End Sub

Public Sub CalMenuPostSupport()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", KeyShift(mMenuKey)
    Exit Sub
Oops:
    ReportError "CalMenuPostSupport"
End Sub

Public Sub CalMenuPostPHA()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "PHA", KeyShift(mMenuKey)
    Exit Sub
Oops:
    ReportError "CalMenuPostPHA"
End Sub

Public Sub CalMenuPostSupportD()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", "Days"
    Exit Sub
Oops:
    ReportError "CalMenuPostSupportD"
End Sub

Public Sub CalMenuPostSupportN()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", "Nights"
    Exit Sub
Oops:
    ReportError "CalMenuPostSupportN"
End Sub

' New Posting for a NEW calendar line (support / PHA with nobody off)
Private Sub PostNewLine(ByVal reason As String, ByVal shiftLbl As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Dim d As Date: d = KeyDate(mMenuKey)
    If d = 0 Then Exit Sub
    If shiftLbl = "" Then shiftLbl = "Days"
    If Int(d) < Int(Date) Then
        If MsgBox("That date has passed. Post it anyway?", vbYesNo + vbQuestion, "Post") = vbNo Then Exit Sub
    End If
    gCalPostKey = "NEW"
    gCalPostDate = d
    gCalPostShift = shiftLbl
    frmNewPosting.Prefill MenuUnit(), d, shiftLbl, reason, FULL_SHIFT, Format$(d, "ddd mmm d") & " " & shiftLbl & ": " & reason
    frmNewPosting.Show
    gCalPostKey = ""
    RefreshCalendar
End Sub

Public Sub CalMenuAddNeed()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim sh As String: sh = KeyShift(mMenuKey)
    If sh = "" Then sh = "Both"
    OpenTimeOff True, KeyDate(mMenuKey), sh
    Exit Sub
Oops:
    ReportError "CalMenuAddNeed"
End Sub

Public Sub CalMenuBookFrom()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    OpenTimeOff False, KeyDate(mMenuKey), ""
    Exit Sub
Oops:
    ReportError "CalMenuBookFrom"
End Sub

' Book Time Off, already filled in with the date / shift / unit
Private Sub OpenTimeOff(ByVal needMode As Boolean, ByVal d As Date, ByVal shiftLbl As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If AbsTable(True) Is Nothing Then Exit Sub
    If CovTable(True) Is Nothing Then Exit Sub
    Dim unit As String: unit = MenuUnit()
    Load frmTimeOff
    On Error Resume Next
    With frmTimeOff
        If needMode Then .Controls("to_optN").Value = True
        If d > 0 Then
            .Controls("to_txtStart").Text = Format$(d, "mm/dd/yyyy")
            If needMode Then .Controls("to_txtEnd").Text = Format$(d, "mm/dd/yyyy")
        End If
        If needMode And shiftLbl <> "" Then .Controls("to_cboShift").Value = shiftLbl
        If unit <> "" Then .Controls("to_cboUnit").Value = unit
    End With
    On Error GoTo 0
    frmTimeOff.Show
    RefreshCalendar
End Sub

Public Sub CalMenuViewPosting()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    Dim rep As String
    If Not PostingOutput(id, True, False, False, rep) Then MsgBox "Couldn't show posting " & id & ". " & rep, vbExclamation
    Exit Sub
Oops:
    ReportError "CalMenuViewPosting"
End Sub

' Award pop-up with this posting already picked - just click LOAD
Public Sub CalMenuAward()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim ids As Variant, ch As Variant, i As Long, idx As Long
    idx = -1
    ch = OpenPostingChoices(ids)
    If IsArray(ids) Then
        For i = LBound(ids) To UBound(ids)
            If StrComp(CStr(ids(i)), id, vbTextCompare) = 0 Then idx = i - LBound(ids)
        Next i
    End If
    Load frmPostingEntry
    If idx >= 0 Then
        On Error Resume Next
        frmPostingEntry.Controls("cmbPosting").ListIndex = idx
        On Error GoTo 0
    End If
    frmPostingEntry.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuAward"
End Sub

' A posting that never made it onto the calendar: pick its line now
Public Sub CalMenuLinkPosting()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim p As PostRec, lst() As PListRec
    LoadPosting id, p, lst
    If p.ID = "" Then MsgBox "Posting " & id & " wasn't found.", vbExclamation: Exit Sub
    LoadPostCache
    Dim byTxt As String, r As Long, aw As Variant, i As Long
    Select Case p.Status
        Case "Open": byTxt = "POSTED " & p.ID
        Case "Cancelled": byTxt = "NCN"
        Case "No takers": byTxt = "No takers"
        Case Else
            r = PostRow(p.ID)
            If r > 0 Then
                aw = Split(Replace(SafeStr(mPostV(r, PS_AWARDTO)), " (out of dept)", ""), ";")
                For i = 0 To UBound(aw)
                    If Trim$(aw(i)) <> "" Then byTxt = byTxt & IIf(byTxt = "", "", " / ") & CalName(Trim$(aw(i)))
                Next i
            End If
            If byTxt = "" Then byTxt = "POSTED " & p.ID
    End Select
    Dim t As String
    t = CalendarPlace(p.ShiftDate, p.ShiftLbl, p.Unit, "Posting " & p.ID & " - which calendar line does it cover?", _
                      byTxt, p.Hours, p.ID, ReasonType(p.Reason), p.Reason, IsSupportReason(p.Reason))
    If t <> "" Then
        SaveIfWritable
        MsgBox p.ID & " is on the calendar: " & t, vbInformation, "Calendar"
    End If
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuLinkPosting"
End Sub

Public Sub CalMenuEndDate()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If SetAbsenceEnd(id) Then RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuEndDate"
End Sub

' v22: is the booking under the menu a leave (so "Return from leave" applies)?
Private Function MenuAbsIsLeave() As Boolean
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Function
    Dim r As Long: r = RowOf(mAbs, id)
    If r > 0 Then
        MenuAbsIsLeave = IsLeaveCode(SafeStr(mAbs(r, AB_TYPE)))
        If MenuAbsIsLeave And UBound(mAbs, 2) >= AB_GROUP Then
            If LeaveToppedUp(SafeStr(mAbs(r, AB_GROUP))) Then MenuAbsIsLeave = False     ' already done
        End If
    End If
End Function

Public Sub CalMenuLeaveReturn()
    On Error GoTo Oops                         ' v24.3.1: the second door into the Leave pop-up (ShowLeaveForm has one)
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not AbsenceColumnsCurrent() Then
        MsgBox "Run Alt+F8 > UpgradeAbsenceColumns first (the Absences table needs its v22 columns).", vbExclamation
        Exit Sub
    End If
    Load frmLeave
    frmLeave.StartReturn id
    frmLeave.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuLeaveReturn"
End Sub

Public Sub CalMenuVoid()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If VoidAbsence(id) Then RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuVoid"
End Sub

Public Sub CalMenuSwitch()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    ShowShiftSwitchForm
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuSwitch"
End Sub

Public Sub CalMenuBook()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    ShowTimeOffForm
    Exit Sub
Oops:
    ReportError "CalMenuBook"
End Sub

' from Needs Attention: open the calendar on that month with the line selected
Public Sub CalMenuGoto()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim keyTxt As String: keyTxt = mMenuKey
    Dim d As Date: d = KeyDate(keyTxt)
    If d = 0 Then Exit Sub
    CalendarSheet
    MapWrite "Z1", DateSerial(Year(d), Month(d), 1)
    LoadCache
    Dim u As String: u = SlotUnit(keyTxt)
    If u <> "" Then
        If Not UnitInGroup(u, CalGroup()) Then MapWrite "Z2", GroupOfUnit(u)
    End If
    ShowCalendar
    On Error Resume Next
    Dim hit As Range
    Set hit = MapSheet().Range("A1:P1000").Find(What:=keyTxt, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not hit Is Nothing Then
        Application.Goto CalendarSheet().Range(hit.Address), False
    End If
    Exit Sub
Oops:
    ReportError "CalMenuGoto"
End Sub


' ====================================================================
'  AUTOMATIC REDRAW + EVENT ROUTING  (called from ThisWorkbook)
' ====================================================================
' Opening the workbook: jump to this month and redraw.
' Opening the workbook: redraw only if the calendar is out of date
' (new day, something changed, or it was left on another month).
Public Sub CalOnOpen()
    mEventsSeen = True
    On Error Resume Next
    If Not SheetExists(CAL_SHEET) Then Exit Sub
    CalendarSheet
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
    On Error Resume Next
    If CalStale() Then DrawCalendar Else SetCalHint CalendarSheet()
End Sub

Public Sub NAActivated()
    If mDrawing Then Exit Sub
    If Not gAdminUnlocked Then Exit Sub
    On Error Resume Next
    If NAStale() Then DrawNeeds
End Sub

Public Sub CalSheetActivated(ByVal Sh As Object)
    mEventsSeen = True
    On Error Resume Next
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
