Attribute VB_Name = "modHolidays"
Option Explicit

' ====================================================================
'  PLANT HOLIDAYS (v20)
'  The same 11 holidays every year, worked out for any year:
'    New Year's Day   Jan 1               Canada Day        Jul 1
'    Family Day       3rd Mon in Feb      Civic Holiday     1st Mon in Aug
'    Good Friday      2 days before Easter Labour Day       1st Mon in Sep
'    Victoria Day     Mon on/before May 24 Thanksgiving     2nd Mon in Oct
'    Remembrance Day  Nov 11              Christmas Day     Dec 25
'                                         Boxing Day        Dec 26
'  The 24-hour floater is booked as vacation - nothing here.
'
'  Who it affects: DAYS (Mon-Fri) workers are OFF on the holiday, so they
'  are free for overtime that day. A/B, C/D and the A-D rotation work as
'  scheduled. modShiftCycle.CrewShiftOn asks DaysHolidayOff().
'
'  Weekend holidays: the plant announces the weekday given off. Enter it
'  in "Observed on" on the Holidays sheet (admin only). Until it's entered,
'  Needs Attention lists the holiday under HOLIDAYS TO CONFIRM.
'  An Observed on date also overrides a weekday holiday if ever needed.
' ====================================================================

Public Const HOL_VERSION As String = "Holidays v29 (Oct 28)"
Public Const HOL_SHEET As String = "Holidays"
Public Const HOL_TABLE As String = "tblHoliday"
Public Const HOL_CONFIRM_DAYS As Long = 90     ' look-ahead for weekend holidays still to confirm

Private Const HC_YEAR As Long = 1
Private Const HC_NAME As Long = 2
Private Const HC_DATE As Long = 3
Private Const HC_DAY As Long = 4
Private Const HC_OBS As Long = 5
Private Const HC_NOTE As Long = 6
Private Const HC_N As Long = 6

Private mOff As Object        ' Long(date) -> holiday name: DAYS workers are off that date
Private mOn As Object         ' Long(date) -> holiday name on its actual date
Private mObs As Object        ' Long(date) -> holiday name, observed on this (different) date
Private mYears As Object      ' years already worked out
Private mSheet As Variant     ' Holidays table contents (read once)
Private mSheetRead As Boolean
Private mTestMode As Boolean  ' v21: SelfTest supplies the "sheet" from memory


' ====================================================================
'  RULES
' ====================================================================
Public Function HolidayNames() As Variant
    HolidayNames = Array("New Year's Day", "Family Day", "Good Friday", "Victoria Day", "Canada Day", _
                         "Civic Holiday", "Labour Day", "Thanksgiving", "Remembrance Day", _
                         "Christmas Day", "Boxing Day")
End Function

Public Function HolidayDate(ByVal holName As String, ByVal y As Long) As Date
    Select Case holName
        Case "New Year's Day": HolidayDate = DateSerial(y, 1, 1)
        Case "Family Day": HolidayDate = NthMonday(y, 2, 3)
        Case "Good Friday": HolidayDate = EasterSunday(y) - 2
        Case "Victoria Day"
            Dim d As Date: d = DateSerial(y, 5, 24)
            HolidayDate = d - (Weekday(d, vbMonday) - 1)
        Case "Canada Day": HolidayDate = DateSerial(y, 7, 1)
        Case "Civic Holiday": HolidayDate = NthMonday(y, 8, 1)
        Case "Labour Day": HolidayDate = NthMonday(y, 9, 1)
        Case "Thanksgiving": HolidayDate = NthMonday(y, 10, 2)
        Case "Remembrance Day": HolidayDate = DateSerial(y, 11, 11)
        Case "Christmas Day": HolidayDate = DateSerial(y, 12, 25)
        Case "Boxing Day": HolidayDate = DateSerial(y, 12, 26)
    End Select
End Function

Private Function NthMonday(ByVal y As Long, ByVal m As Long, ByVal n As Long) As Date
    Dim d As Date: d = DateSerial(y, m, 1)
    d = d + ((9 - Weekday(d, vbSunday)) Mod 7)        ' first Monday
    NthMonday = d + 7 * (n - 1)
End Function

' Western (Gregorian) Easter Sunday
Private Function EasterSunday(ByVal y As Long) As Date
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long
    Dim h As Long, i As Long, k As Long, L As Long, m As Long
    a = y Mod 19: b = y \ 100: c = y Mod 100
    d = b \ 4: e = b Mod 4
    f = (b + 8) \ 25: g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4: k = c Mod 4
    L = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * L) \ 451
    EasterSunday = DateSerial(y, (h + L - 7 * m + 114) \ 31, ((h + L - 7 * m + 114) Mod 31) + 1)
End Function

Private Function IsWeekend(ByVal d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) >= 6)
End Function


' ====================================================================
'  LOOKUPS (fast - cached)
' ====================================================================
' True when DAYS (Mon-Fri) workers are off for a holiday on date d
Public Function DaysHolidayOff(ByVal d As Date) As Boolean
    EnsureYears Year(d)
    DaysHolidayOff = mOff.Exists(CLng(Int(d)))
End Function

' Label for the calendar: "Thanksgiving" on the holiday itself,
' "Boxing Day (observed)" on the weekday given for a weekend holiday.
Public Function HolidayLabel(ByVal d As Date) As String
    EnsureYears Year(d)
    Dim k As Long: k = CLng(Int(d))
    If mOn.Exists(k) Then
        HolidayLabel = mOn(k)
    ElseIf mObs.Exists(k) Then
        HolidayLabel = mObs(k) & " (observed)"
    End If
End Function

' Forget the cached dates (the Holidays sheet changed)
Public Sub ResetHolidayCache()
    Set mOff = Nothing: Set mOn = Nothing: Set mObs = Nothing: Set mYears = Nothing
    If Not mTestMode Then mSheet = Empty: mSheetRead = False
End Sub

' SelfTest (v21): rows in the Holidays table layout (Year, Holiday, Date,
' Day, Observed on, Note) stand in for the sheet until ClearTestHolidays.
Public Sub SetTestHolidays(ByVal rows As Variant)
    Set mOff = Nothing: Set mOn = Nothing: Set mObs = Nothing: Set mYears = Nothing
    mTestMode = True
    mSheet = rows
    mSheetRead = True
End Sub

Public Sub ClearTestHolidays()
    mTestMode = False
    Set mOff = Nothing: Set mOn = Nothing: Set mObs = Nothing: Set mYears = Nothing
    mSheet = Empty: mSheetRead = False
End Sub

' Works out year y and the years either side (an observed day can cross
' New Year, e.g. a Saturday Jan 1 observed on Friday Dec 31).
Private Sub EnsureYears(ByVal y As Long)
    If mYears Is Nothing Then
        Set mOff = CreateObject("Scripting.Dictionary")
        Set mOn = CreateObject("Scripting.Dictionary")
        Set mObs = CreateObject("Scripting.Dictionary")
        Set mYears = CreateObject("Scripting.Dictionary")
    End If
    Dim yy As Long
    For yy = y - 1 To y + 1
        If Not mYears.Exists(yy) Then LoadYear yy
    Next yy
End Sub

Private Sub LoadYear(ByVal y As Long)
    mYears(y) = True
    If Not mSheetRead Then ReadSheet
    Dim nm As Variant, d As Date, obs As Variant
    For Each nm In HolidayNames()
        d = HolidayDate(CStr(nm), y)
        mOn(CLng(d)) = CStr(nm)
        obs = SheetObserved(y, CStr(nm))
        If IsDate(obs) Then
            mOff(CLng(Int(CDate(obs)))) = CStr(nm)
            If Int(CDate(obs)) <> d Then mObs(CLng(Int(CDate(obs)))) = CStr(nm)
        ElseIf Not IsWeekend(d) Then
            mOff(CLng(d)) = CStr(nm)
        End If
    Next nm
End Sub

Private Sub ReadSheet()
    mSheetRead = True
    If mTestMode Then Exit Sub
    mSheet = Empty
    Dim lo As ListObject: Set lo = HolidayTable(False)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Sub
    mSheet = RangeTo2D(lo.DataBodyRange)
End Sub

' "Observed on" entered on the sheet for this holiday (Empty if none)
Private Function SheetObserved(ByVal y As Long, ByVal holName As String) As Variant
    SheetObserved = Empty
    If Not IsArray(mSheet) Then Exit Function
    Dim r As Long
    For r = 1 To UBound(mSheet, 1)
        If Val(SafeStr(mSheet(r, HC_YEAR))) = y And _
           StrComp(Trim$(SafeStr(mSheet(r, HC_NAME))), holName, vbTextCompare) = 0 Then
            If IsDate(mSheet(r, HC_OBS)) Then SheetObserved = CDate(mSheet(r, HC_OBS))
            Exit Function
        End If
    Next r
End Function

' Weekend holidays from fromD to toD with no weekday entered yet.
' Returns an array of "yyyy-mm-dd|Holiday" strings, or Empty.
Public Function HolidaysToConfirm(ByVal fromD As Date, ByVal toD As Date) As Variant
    If Not mSheetRead Then ReadSheet
    Dim out() As String, n As Long, y As Long, nm As Variant, d As Date
    For y = Year(fromD) To Year(toD)
        For Each nm In HolidayNames()
            d = HolidayDate(CStr(nm), y)
            If d >= Int(fromD) And d <= Int(toD) And IsWeekend(d) Then
                If Not IsDate(SheetObserved(y, CStr(nm))) Then
                    n = n + 1
                    ReDim Preserve out(1 To n)
                    out(n) = Format$(d, "yyyy-mm-dd") & "|" & nm
                End If
            End If
        Next nm
    Next y
    If n > 0 Then HolidaysToConfirm = out
End Function


' ====================================================================
'  HOLIDAYS SHEET (admin only)
' ====================================================================
' v25 (R4): the "Observed on" column was the one cell block an admin typed into; v27 (item 33): none, adminView unused
Public Sub LockHolidayCells(ByVal adminView As Boolean)
    Dim lo As ListObject: Set lo = HolidayTable()
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    ws.Cells.Locked = True                          ' v27 (item 33): every cell - the Holidays pop-up writes Observed on (SetObservedOn)
End Sub

Public Function HolidayTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(HOL_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(HOL_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(HOL_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = HOL_SHEET
            RelockStructure
            SetTabColor ws, RGB(176, 0, 32)             ' v25.1
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, HC_N).Value = Array("Year", "Holiday", "Date", "Day", "Observed on", "Note")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, HC_N), , xlYes)
        lo.Name = HOL_TABLE
        lo.TableStyle = "TableStyleMedium3"
        Dim w As Variant, i As Long
        w = Array(7, 18, 13, 6, 13, 52)
        For i = 1 To HC_N
            ws.Columns(i).ColumnWidth = w(i - 1)
        Next i
        lo.ListColumns(HC_DATE).DataBodyRange.NumberFormat = "mmm d, yyyy"
        lo.ListColumns(HC_OBS).DataBodyRange.NumberFormat = "ddd mmm d, yyyy"
        ProtectSheet ws
    End If
    Set HolidayTable = lo
End Function

' Adds any missing rows for years y1..y2. Admin + writable only; quiet.
' Returns the number of rows added.
Public Function EnsureHolidayYears(ByVal y1 As Long, ByVal y2 As Long) As Long
    If ThisWorkbook.ReadOnly Or Not IsAdmin() Then Exit Function
    Dim lo As ListObject: Set lo = HolidayTable(True)
    If lo Is Nothing Then Exit Function

    ' what's there already
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1
    Dim v As Variant, r As Long, blankFirst As Boolean
    If Not lo.DataBodyRange Is Nothing Then
        v = RangeTo2D(lo.DataBodyRange)
        For r = 1 To UBound(v, 1)
            If SafeStr(v(r, HC_NAME)) <> "" Then have(SafeStr(v(r, HC_YEAR)) & "|" & Trim$(SafeStr(v(r, HC_NAME)))) = True
        Next r
        blankFirst = (lo.ListRows.Count = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0)
    End If

    Dim addRows() As Variant, n As Long, y As Long, nm As Variant, d As Date
    ReDim addRows(1 To 11 * (y2 - y1 + 1), 1 To HC_N)
    For y = y1 To y2
        For Each nm In HolidayNames()
            If Not have.Exists(CStr(y) & "|" & nm) Then
                d = HolidayDate(CStr(nm), y)
                n = n + 1
                addRows(n, HC_YEAR) = y
                addRows(n, HC_NAME) = CStr(nm)
                addRows(n, HC_DATE) = d
                addRows(n, HC_DAY) = Format$(d, "ddd")
                If IsWeekend(d) Then
                    addRows(n, HC_OBS) = Empty
                    addRows(n, HC_NOTE) = "Falls on a " & Format$(d, "dddd") & " - enter the weekday the plant gives off"
                Else
                    addRows(n, HC_OBS) = d
                    addRows(n, HC_NOTE) = ""
                End If
            End If
        Next nm
    Next y
    If n = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim evOn As Boolean: evOn = Application.EnableEvents
    On Error GoTo Fail
    Application.EnableEvents = False
    UnlockSheet ws
    Dim cnt As Long: cnt = lo.ListRows.Count
    If blankFirst Then cnt = 0
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    Dim outA() As Variant, i As Long, j As Long
    ReDim outA(1 To n, 1 To HC_N)
    For i = 1 To n
        For j = 1 To HC_N
            outA(i, j) = addRows(i, j)
        Next j
    Next i
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, HC_N).Value = outA
    lo.ListColumns(HC_DATE).DataBodyRange.NumberFormat = "mmm d, yyyy"
    lo.ListColumns(HC_OBS).DataBodyRange.NumberFormat = "ddd mmm d, yyyy"
    ' newest year last, in date order
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(HC_DATE).DataBodyRange, Order:=xlAscending
        .Header = xlYes
        .Apply
    End With
    LockHolidayCells gAdminUnlocked               ' v25 review: the new rows' "Observed on" cells open unlocked for an admin
    ProtectSheet ws
    Application.EnableEvents = evOn
    ResetHolidayCache
    EnsureHolidayYears = n
    Exit Function
Fail:
    Dim em As String: em = Err.Description
    LogError "EnsureHolidayYears", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evOn
    Debug.Print "EnsureHolidayYears: " & em
End Function

' Workbook_Open: make sure this year and next are listed (admins only)
Public Sub HolidaysOnOpen()
    On Error Resume Next                            ' never raises
    If EnsureHolidayYears(Year(Date), Year(Date) + 1) > 0 Then MarkCalDirty
End Sub

' Alt+F8 > ShowHolidays (admin): opens the Holidays sheet
Public Sub ShowHolidays()
    If Not RequireAdmin() Then Exit Sub
    On Error GoTo Fail
    EnsureHolidayYears Year(Date), Year(Date) + 1
    If HolidayTable(False) Is Nothing Then
        MsgBox "The Holidays sheet couldn't be created (is the file read-only?).", vbExclamation, "Call Out Roster"
        Exit Sub
    End If
    SetSheetVisible ThisWorkbook.Worksheets(HOL_SHEET), xlSheetVisible
    ThisWorkbook.Worksheets(HOL_SHEET).Activate
    Exit Sub
Fail:
    ReportError "ShowHolidays"                     ' v25 (I5): logged, state restored, plain message
End Sub

' Opens the sheet on a given holiday's row (from Needs Attention); v27 (item 33): the pop-up on that row instead
Public Sub GoToHoliday(ByVal d As Date)
    If Not RequireAdmin() Then Exit Sub
    On Error GoTo Fail
    OpenHolidayEditor d
    Exit Sub
Fail:
    Dim em As String: em = Err.Description
    LogError "GoToHoliday", Err.Number, em          ' v25 (I5): an error here used to be a bare box, never logged
    MsgBox "The Holidays pop-up couldn't be opened: " & em, vbExclamation, "Call Out Roster"     ' v27: em (LogError clears Err)
End Sub

' ====================================================================
'  THE HOLIDAYS POP-UP (v27 item 33)
' ====================================================================
' Alt+F8 / the admin Desk (Holidays / minimum staffing): the pop-up on the Holidays table
Public Sub ShowHolidayEditor()
    On Error GoTo Oops
    OpenHolidayEditor 0
    Exit Sub
Oops:
    ReportError "ShowHolidayEditor"
End Sub

' the pop-up with one holiday's row picked (0 = none); Needs Attention's HOLIDAYS TO CONFIRM lines open it so
Public Sub OpenHolidayEditor(ByVal pick As Date)
    If Not RequireAdmin() Then Exit Sub
    EnsureHolidayYears Year(Date), Year(Date) + 1
    If pick > 0 Then EnsureHolidayYears Year(pick), Year(pick)
    If HolidayTable(False) Is Nothing Then MsgBox "The Holidays sheet couldn't be created (is the file read-only?).", vbExclamation, "Holidays": Exit Sub
    frmTableEdit.Setup "holidays", pick
    frmTableEdit.Show
    Unload frmTableEdit                             ' (review L6) the form hides itself on Close
End Sub

' the pop-up's rows: the table as it stands (Year, Holiday, Date, Day, Observed on, Note), or Empty
Public Function HolidayEditRows() As Variant
    Dim lo As ListObject: Set lo = HolidayTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    HolidayEditRows = RangeTo2D(lo.DataBodyRange)
End Function

' The ONE write to an "Observed on" cell: row r of the table (1-based); obs 0 = cleared. Admin, writable. A weekday
' within a week of the holiday. The holiday cache and the calendar follow. False = refused (why).
Public Function SetObservedOn(ByVal r As Long, ByVal obs As Date, ByRef why As String) As Boolean
    why = ""
    If Not RequireAdmin() Then why = "admins only": Exit Function
    If Not RequireWritable() Then why = "this copy can't be saved": Exit Function
    Dim lo As ListObject: Set lo = HolidayTable(False)
    If lo Is Nothing Then why = "there is no Holidays table": Exit Function
    If lo.DataBodyRange Is Nothing Then why = "the Holidays table is empty": Exit Function
    If r < 1 Or r > lo.ListRows.Count Then why = "that row is gone - close and open the pop-up again": Exit Function
    Dim hd As Variant: hd = lo.DataBodyRange.Cells(r, HC_DATE).Value
    If obs > 0 Then
        If Weekday(obs, vbMonday) > 5 Then why = Format$(obs, "ddd mmm d") & " is a weekend day": Exit Function
        If IsDate(hd) Then
            If Abs(Int(obs) - Int(CDate(hd))) > 7 Then why = "the day off must be within a week of the holiday (" & Format$(CDate(hd), "ddd mmm d") & ")": Exit Function
        End If
    End If
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet ws
    With lo.DataBodyRange.Cells(r, HC_OBS)
        If obs > 0 Then
            .NumberFormat = "ddd mmm d, yyyy"
            .Value = Int(obs)
        Else
            .ClearContents
        End If
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    ResetHolidayCache
    MarkCalDirty
    SetObservedOn = True
End Function
