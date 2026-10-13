Attribute VB_Name = "modShiftSwitch"
Option Explicit

' ====================================================================
'  SHIFT SWITCHES (S/S)
'  One record = one date: "Works" covers "Gives up"'s shift.
'  The reciprocal payback is its own record on its own date, linked
'  back with "Payback For" (and the original shows "Paid Back By").
'  No hours / charging impact - scheduling only. Call-out lists flag
'  anyone who is working a switch on that date / shift.
' ====================================================================

Public Const SS_VERSION As String = "Shift switches v24.3 (Jan 5)"
Public Const SS_SHEET As String = "Shift Switches"
Public Const SS_TABLE As String = "tblShiftSwitch"

Public Const SW_ID As Long = 1
Public Const SW_DATE As Long = 2
Public Const SW_SHIFT As Long = 3
Public Const SW_UNIT As Long = 4
Public Const SW_OFF As Long = 5
Public Const SW_OFFCREW As Long = 6
Public Const SW_WORKS As Long = 7
Public Const SW_WORKSCREW As Long = 8
Public Const SW_PAYFOR As Long = 9
Public Const SW_PAIDBY As Long = 10
Public Const SW_NOTE As Long = 11
Public Const SW_STATUS As Long = 12
Public Const SW_USER As Long = 13
Public Const SW_STAMP As Long = 14
Public Const SW_VOIDNOTE As Long = 15
Public Const SW_N As Long = 15

Private mTestMode As Boolean        ' v21: SelfTest feeds SwitchData from memory
Private mTestRows As Variant
Private mBatch As Boolean           ' M1: one read of the table reused for a whole publish sweep
Private mBatchRows As Variant


' ---------------- launcher (roster button) ----------------
Public Sub ShowShiftSwitchForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If SwitchTable(True) Is Nothing Then Exit Sub
    SyncSwitchesToCalendar
    frmShiftSwitch.Show
    Exit Sub
Oops:
    ReportError "ShowShiftSwitchForm"
End Sub

' Any active switch without a calendar line (e.g. logged before the
' calendar existed) gets one. Returns how many were added.
Public Function SyncSwitchesToCalendar() As Long
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long, id As String
    On Error GoTo Quit
    For i = 1 To UBound(v, 1)
        id = Trim$(SafeStr(v(i, SW_ID)))
        If id <> "" And UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" And IsDate(v(i, SW_DATE)) Then
            If Not CoverageHasSource(id) Then
                If AddCoverage(CDate(v(i, SW_DATE)), SafeStr(v(i, SW_SHIFT)), SafeStr(v(i, SW_UNIT)), "NEW", _
                               CalName(SafeStr(v(i, SW_WORKS))) & " S/S", Empty, id, "SS", _
                               CalName(SafeStr(v(i, SW_OFF))) & " S/S") <> "" Then n = n + 1
            End If
        End If
    Next i
Quit:
    If n > 0 Then SaveIfWritable
    SyncSwitchesToCalendar = n
End Function


' ---------------- table ----------------
Public Function SwitchTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(SS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(SS_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(SS_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then Set lo = CreateSwitchTable()
    Set SwitchTable = lo
End Function

Private Function CreateSwitchTable() As ListObject
    Dim ws As Worksheet
    If SheetExists(SS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(SS_SHEET)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SS_SHEET
        ws.Tab.Color = RGB(180, 100, 0)
    End If
    ws.Unprotect Password:=SHEET_PASSWORD
    ws.Range("A1").Resize(1, SW_N).Value = Array("SS ID", "Shift Date", "Shift", "Unit", "Gives Up", "Off Crew", _
        "Works", "Works Crew", "Payback For", "Paid Back By", "Note", "Status", "Logged By", "Timestamp", "Void Note")
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, SW_N), , xlYes)
    lo.Name = SS_TABLE
    lo.TableStyle = "TableStyleMedium3"
    Dim w As Variant, i As Long
    w = Array(9, 15, 8, 14, 24, 7, 24, 7, 10, 11, 30, 8, 11, 18, 30)
    For i = 1 To SW_N
        ws.Columns(i).ColumnWidth = w(i - 1)
    Next i
    ws.Columns(SW_DATE).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(SW_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ws.Rows(1).WrapText = True
    Dim fc As FormatCondition
    Set fc = ws.Range("A2:" & ColLetter(SW_N) & "20000").FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=$" & ColLetter(SW_STATUS) & "2=""VOID""")
    fc.Font.Strikethrough = True
    fc.Font.Color = RGB(150, 150, 150)
    ProtectSheet ws
    Set CreateSwitchTable = lo
End Function

' SelfTest (v21): switch rows in the tblShiftSwitch layout (or Empty)
' replace the table until ClearTestSwitches. The sheet is never touched.
Public Sub SetTestSwitches(ByVal rows As Variant)
    mTestMode = True
    mTestRows = rows
End Sub

Public Sub ClearTestSwitches()
    mTestMode = False
    mTestRows = Empty
End Sub

' M1: the members' publish asks "who is free" hundreds of times; read the
' switch table once for the whole sweep. Always paired with EndSwitchBatch.
Public Sub BeginSwitchBatch()
    mBatch = False
    mBatchRows = SwitchData()
    mBatch = True
End Sub

Public Sub EndSwitchBatch()
    mBatch = False
    mBatchRows = Empty
End Sub

Private Function SwitchData() As Variant
    If mTestMode Then
        If IsArray(mTestRows) Then SwitchData = mTestRows
        Exit Function
    End If
    If mBatch Then
        If IsArray(mBatchRows) Then SwitchData = mBatchRows
        Exit Function
    End If
    Dim lo As ListObject: Set lo = SwitchTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    SwitchData = RangeTo2D(lo.DataBodyRange)
End Function

Private Function NextSwitchID() As String
    Dim v As Variant: v = SwitchData()
    Dim i As Long, k As Long, mx As Long, s As String
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            s = UCase$(Trim$(SafeStr(v(i, SW_ID))))
            If Left$(s, 3) = "SS-" Then
                k = CLng(Val(Mid$(s, 4)))
                If k > mx Then mx = k
            End If
        Next i
    End If
    NextSwitchID = "SS-" & Format$(mx + 1, "0000")
End Function


' ---------------- roster helpers ----------------
' All roster names, A-Z
Public Function RosterNames() As Variant
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = RangeTo2D(ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)))   ' one name = one cell
    Dim a() As String, n As Long, i As Long, j As Long, t As String
    ReDim a(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        t = Trim$(SafeStr(v(i, 1)))
        If t <> "" Then n = n + 1: a(n) = t
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve a(1 To n)
    For i = 2 To n
        t = a(i): j = i - 1
        Do While j >= 1
            If StrComp(a(j), t, vbTextCompare) > 0 Then
                a(j + 1) = a(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        a(j + 1) = t
    Next i
    RosterNames = a
End Function

' Rotation crew (A-D); "" for DAYS / unknown. With a date (v24): the crew
' they are on that day (crew history); without: the roster's crew column.
Public Function EmpCrew(ByVal empName As String, Optional ByVal asOf As Variant) As String
    Dim s As String
    If IsMissing(asOf) Then s = EmpCrewCode(empName) Else s = EmpCrewCode(empName, asOf)
    If Len(s) = 1 And InStr("ABCD", s) > 0 Then EmpCrew = s
End Function

' Crew code, upper case, no spaces: "A".."D", "A/B", "C/D", "DAYS" ...
' With a date (v24): the crew they are on that day (crew history).
' Without: the roster's crew column as typed.
Public Function EmpCrewCode(ByVal empName As String, Optional ByVal asOf As Variant) As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), Trim$(empName), vbTextCompare) = 0 Then
            EmpCrewCode = UCase$(Replace(Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value)), " ", ""))
            Exit For
        End If
    Next r
    If Not IsMissing(asOf) Then
        If IsDate(asOf) Then EmpCrewCode = UCase$(Replace(CrewOn(empName, CDate(asOf), EmpCrewCode), " ", ""))
    End If
End Function

Private Function IsPairCrew(ByVal code As String) As Boolean
    IsPairCrew = (code Like "[ABCD]/[ABCD]")
End Function

' What the rotation says about a person on date d, e.g. "Crew C - off" / "Crew A - on days"
Public Function CrewStatus(ByVal empName As String, ByVal d As Date) As String
    Dim code As String: code = EmpCrewCode(empName, d)
    If IsPairCrew(code) Then
        If CrewShiftOn(code, d) = "Days" Then
            CrewStatus = code & " days - on DAYS"
        Else
            CrewStatus = code & " days - off"
        End If
        Exit Function
    End If
    If code = "DAYS" Then
        If CrewShiftOn(code, d) = "Days" Then
            CrewStatus = "DAYS - on DAYS"
        ElseIf Weekday(d, vbMonday) <= 5 Then
            CrewStatus = "DAYS - off (" & HolidayLabel(d) & ")"
        Else
            CrewStatus = "DAYS - off (weekend)"
        End If
        Exit Function
    End If
    Dim c As String: c = EmpCrew(empName, d)
    If c = "" Then CrewStatus = "not on the 4-crew rotation": Exit Function
    If c = DaysCrew(d) Then
        CrewStatus = "Crew " & c & " - on DAYS"
    ElseIf c = NightsCrew(d) Then
        CrewStatus = "Crew " & c & " - on NIGHTS"
    Else
        CrewStatus = "Crew " & c & " - off"
    End If
End Function

' Problems with "works" covering the shift (blank = looks fine)
Public Function WorksConflict(ByVal empName As String, ByVal d As Date, ByVal isNight As Boolean) As String
    ' v24: a crew change can fall between d - 1, d and d + 1, so each day has its own crew
    Dim code As String: code = EmpCrewCode(empName, d)
    Dim codePrev As String: codePrev = EmpCrewCode(empName, d - 1)
    Dim codeNext As String: codeNext = EmpCrewCode(empName, d + 1)
    If IsPairCrew(code) Or code = "DAYS" Then   ' DAYS: Mon-Fri days (v20)
        If CrewShiftOn(code, d) = "Days" Then
            WorksConflict = code & " is already working days that day"
        ElseIf isNight And CrewShiftOn(codeNext, d + 1) = "Days" Then
            WorksConflict = codeNext & " starts days the next morning"
        ElseIf Not isNight And CrewShiftOn(codePrev, d - 1) = "Nights" Then
            WorksConflict = "crew " & codePrev & " is coming off nights that morning"     ' was on a rotation crew the day before
        End If
        Exit Function
    End If
    Dim c As String: c = EmpCrew(empName, d)
    If c = "" Then Exit Function
    If c = DaysCrew(d) Or c = NightsCrew(d) Then
        WorksConflict = "crew " & c & " is already working that day"
    ElseIf Not isNight And CrewShiftOn(codePrev, d - 1) = "Nights" Then
        WorksConflict = "crew " & codePrev & " is coming off nights that morning"
    ElseIf isNight And CrewShiftOn(codeNext, d + 1) = "Days" Then
        WorksConflict = "crew " & codeNext & " starts days the next morning"
    End If
End Function

' Problems with "gives up" (blank = they're scheduled on that shift)
Public Function OffConflict(ByVal empName As String, ByVal d As Date, ByVal isNight As Boolean) As String
    Dim code As String: code = EmpCrewCode(empName, d)
    If IsPairCrew(code) Or code = "DAYS" Then   ' DAYS: Mon-Fri days (v20)
        If isNight Then
            OffConflict = code & " works days only"
        ElseIf CrewShiftOn(code, d) <> "Days" Then
            OffConflict = code & " isn't on days that day"
        End If
        Exit Function
    End If
    Dim c As String: c = EmpCrew(empName, d)
    If c = "" Then Exit Function
    If isNight And c <> NightsCrew(d) Then OffConflict = "crew " & c & " isn't on nights that day"
    If Not isNight And c <> DaysCrew(d) Then OffConflict = "crew " & c & " isn't on days that day"
End Function

' Days / Nights the person is scheduled for on d ("" if off)
Public Function ScheduledShift(ByVal empName As String, ByVal d As Date) As String
    Dim code As String: code = EmpCrewCode(empName, d)
    If Len(code) = 1 Or IsPairCrew(code) Or code = "DAYS" Then
        ScheduledShift = CrewShiftOn(code, d)
        If ScheduledShift = "?" Then ScheduledShift = ""
    End If
End Function


' ---------------- lookups ----------------
' SS ID if empName WORKS a switch on d / shiftLbl ("" if not)
Public Function SwitchWorkingOn(ByVal empName As String, ByVal d As Date, ByVal shiftLbl As String) As String
    If mBatch Then
        SwitchWorkingOn = MatchSwitch(mBatchRows, SW_WORKS, empName, d, shiftLbl)   ' no copy of the rows per call
    Else
        Dim v As Variant: v = SwitchData()
        SwitchWorkingOn = MatchSwitch(v, SW_WORKS, empName, d, shiftLbl)
    End If
End Function

' SS ID if empName GAVE UP a shift (any shift) on date d ("" if not)
Public Function SwitchOffOn(ByVal empName As String, ByVal d As Date) As String
    If mBatch Then
        SwitchOffOn = MatchSwitch(mBatchRows, SW_OFF, empName, d, "")
    Else
        Dim v As Variant: v = SwitchData()
        SwitchOffOn = MatchSwitch(v, SW_OFF, empName, d, "")
    End If
End Function

' first non-void row where column nameCol is empName on date d (and shiftLbl, if given)
Private Function MatchSwitch(ByRef v As Variant, ByVal nameCol As Long, ByVal empName As String, _
                             ByVal d As Date, ByVal shiftLbl As String) As String
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" Then
            If StrComp(Trim$(SafeStr(v(i, nameCol))), Trim$(empName), vbTextCompare) = 0 Then
                If IsDate(v(i, SW_DATE)) Then
                    If Int(CDate(v(i, SW_DATE))) = Int(d) And _
                       (shiftLbl = "" Or StrComp(SafeStr(v(i, SW_SHIFT)), shiftLbl, vbTextCompare) = 0) Then
                        MatchSwitch = SafeStr(v(i, SW_ID)): Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' Hard stops for one switch (blank = OK):
'   the person giving up the shift must not be booked off (vac / training /
'   leave) or already on a switch that day; the person working it must not
'   be booked off or already working a switch that day.
Public Function EligibilityProblems(ByVal offName As String, ByVal worksName As String, ByVal d As Date) As String
    Dim s As String, t As String, id As String
    If BookedOff(offName, d, t) Then s = s & "- " & offName & " is booked off (" & t & ") on " & Format$(d, "mmm d") & vbNewLine
    id = SwitchOffOn(offName, d)
    If id <> "" Then s = s & "- " & offName & " already switched off that day (" & id & ")" & vbNewLine
    id = SwitchWorkingOn(offName, d, "")
    If id <> "" Then s = s & "- " & offName & " is already working a switch that day (" & id & ")" & vbNewLine
    If BookedOff(worksName, d, t) Then s = s & "- " & worksName & " is booked off (" & t & ") on " & Format$(d, "mmm d") & vbNewLine
    id = SwitchWorkingOn(worksName, d, "")
    If id <> "" Then s = s & "- " & worksName & " is already working a switch that day (" & id & ")" & vbNewLine
    EligibilityProblems = s
End Function

' Short status for the pop-up hint ("booked off (Vac)", "on SS-0003") or "".
' kind (optional) says which: "OFF" booked off, "GAVE_UP" already gave up a
' shift that day, "WORKING" already working a switch, "" none. The form
' checks kind, never the wording.
Public Function PersonFlag(ByVal empName As String, ByVal d As Date, Optional ByRef kind As String) As String
    Dim t As String, id As String
    kind = ""
    If BookedOff(empName, d, t) Then PersonFlag = "booked off (" & t & ")": kind = "OFF": Exit Function
    id = SwitchOffOn(empName, d)
    If id <> "" Then PersonFlag = "already switched off (" & id & ")": kind = "GAVE_UP": Exit Function
    id = SwitchWorkingOn(empName, d, "")
    If id <> "" Then PersonFlag = "already working " & id: kind = "WORKING"
End Function

' Unpaid switches this pair could be paying back: earlier records where
' the person now WORKING gave up a shift and the person now OFF worked it.
' Returns display strings; ids() gets the SS IDs.
Public Function OpenPaybacks(ByVal offName As String, ByVal worksName As String, ByRef ids As Variant) As Variant
    ids = Empty
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long, d() As String, idArr() As String
    ReDim d(0 To UBound(v, 1) - 1): ReDim idArr(0 To UBound(v, 1) - 1)
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" And Trim$(SafeStr(v(i, SW_PAIDBY))) = "" _
           And Trim$(SafeStr(v(i, SW_PAYFOR))) = "" Then
            If StrComp(Trim$(SafeStr(v(i, SW_OFF))), Trim$(worksName), vbTextCompare) = 0 And _
               StrComp(Trim$(SafeStr(v(i, SW_WORKS))), Trim$(offName), vbTextCompare) = 0 Then
                idArr(n) = SafeStr(v(i, SW_ID))
                d(n) = idArr(n) & "  " & Format$(v(i, SW_DATE), "ddd mmm d") & " " & SafeStr(v(i, SW_SHIFT)) & _
                       " - " & SafeStr(v(i, SW_WORKS)) & " worked for " & SafeStr(v(i, SW_OFF))
                n = n + 1
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve d(0 To n - 1): ReDim Preserve idArr(0 To n - 1)
    ids = idArr
    OpenPaybacks = d
End Function

' Paybacks owed for more than minDays days (IntegrityCheck, v23). Returns how
' many; txt lists the first five ("SS-0004 Oct 2; SS-0009 Oct 20").
Public Function OwedOlderThan(ByVal minDays As Long, ByRef txt As String) As Long
    txt = ""
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(v, 1)
        If IsOwed(v, i) And IsDate(v(i, SW_DATE)) Then
            If Int(Date) - Int(CDate(v(i, SW_DATE))) > minDays Then
                n = n + 1
                If n <= 5 Then txt = txt & IIf(txt = "", "", "; ") & SafeStr(v(i, SW_ID)) & " " & Format$(CDate(v(i, SW_DATE)), "mmm d")
            End If
        End If
    Next i
    If n > 5 Then txt = txt & " ... and " & (n - 5) & " more"
    OwedOlderThan = n
End Function

' Switches still OWED a payback (original switches with no payback logged),
' oldest first. 2-D array (rows x 6) for the pop-up list; ids() gets the IDs.
Public Function OwedSwitches(ByRef ids As Variant) As Variant
    ids = Empty
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, n As Long
    For i = 1 To UBound(v, 1)
        If IsOwed(v, i) Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    Dim out() As String, idArr() As String, k As Long
    ReDim out(0 To n - 1, 0 To 5): ReDim idArr(0 To n - 1)
    For i = 1 To UBound(v, 1)
        If IsOwed(v, i) Then
            idArr(k) = SafeStr(v(i, SW_ID))
            out(k, 0) = idArr(k)
            out(k, 1) = Format$(v(i, SW_DATE), "ddd mmm d")
            out(k, 2) = Left$(SafeStr(v(i, SW_SHIFT)), 1)
            out(k, 3) = SafeStr(v(i, SW_OFF))
            out(k, 4) = SafeStr(v(i, SW_WORKS))
            out(k, 5) = SafeStr(v(i, SW_UNIT))
            k = k + 1
        End If
    Next i
    ids = idArr
    OwedSwitches = out
End Function

Private Function IsOwed(ByVal v As Variant, ByVal i As Long) As Boolean
    IsOwed = (UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" And Trim$(SafeStr(v(i, SW_PAIDBY))) = "" _
              And Trim$(SafeStr(v(i, SW_PAYFOR))) = "" And Trim$(SafeStr(v(i, SW_ID))) <> "")
End Function

' Details of one switch (for filling the payback). False if not found.
Public Function SwitchInfo(ByVal id As String, ByRef d As Date, ByRef shiftLbl As String, ByRef unit As String, _
                           ByRef offName As String, ByRef worksName As String) As Boolean
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, SW_ID))), id, vbTextCompare) = 0 Then
            If IsDate(v(i, SW_DATE)) Then d = CDate(v(i, SW_DATE))
            shiftLbl = SafeStr(v(i, SW_SHIFT))
            unit = SafeStr(v(i, SW_UNIT))
            offName = SafeStr(v(i, SW_OFF))
            worksName = SafeStr(v(i, SW_WORKS))
            SwitchInfo = True
            Exit Function
        End If
    Next i
End Function

' Most recent switches for the pop-up list: 2-D array (rows x 6) and ids()
Public Function RecentSwitches(ByVal maxRows As Long, ByRef ids As Variant) As Variant
    ids = Empty
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim n As Long: n = UBound(v, 1)
    If n > maxRows Then n = maxRows
    Dim out() As String, idArr() As String, i As Long, r As Long, pay As String
    ReDim out(0 To n - 1, 0 To 5): ReDim idArr(0 To n - 1)
    For i = 0 To n - 1
        r = UBound(v, 1) - i                       ' newest first
        idArr(i) = SafeStr(v(r, SW_ID))
        out(i, 0) = idArr(i)
        out(i, 1) = Format$(v(r, SW_DATE), "ddd mmm d")
        out(i, 2) = Left$(SafeStr(v(r, SW_SHIFT)), 1)
        out(i, 3) = SafeStr(v(r, SW_OFF))
        out(i, 4) = SafeStr(v(r, SW_WORKS))
        If UCase$(SafeStr(v(r, SW_STATUS))) = "VOID" Then
            pay = "VOID"
        ElseIf SafeStr(v(r, SW_PAYFOR)) <> "" Then
            pay = "pays " & SafeStr(v(r, SW_PAYFOR))
        ElseIf SafeStr(v(r, SW_PAIDBY)) <> "" Then
            pay = "paid " & SafeStr(v(r, SW_PAIDBY))
        Else
            pay = "owed"
        End If
        out(i, 5) = pay
    Next i
    ids = idArr
    RecentSwitches = out
End Function


' ---------------- write ----------------
Public Function SaveSwitch(ByVal d As Date, ByVal shiftLbl As String, ByVal unit As String, _
                           ByVal offName As String, ByVal worksName As String, _
                           ByVal payFor As String, ByVal note As String) As String
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = SwitchTable(True)
    If lo Is Nothing Then Exit Function

    Dim id As String: id = NextSwitchID()
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To SW_N)
    rec(1, SW_ID) = id
    rec(1, SW_DATE) = Int(d)
    rec(1, SW_SHIFT) = shiftLbl
    rec(1, SW_UNIT) = unit
    rec(1, SW_OFF) = offName
    rec(1, SW_OFFCREW) = EmpCrewCode(offName, d)
    rec(1, SW_WORKS) = worksName
    rec(1, SW_WORKSCREW) = EmpCrewCode(worksName, d)
    rec(1, SW_PAYFOR) = payFor
    rec(1, SW_NOTE) = note
    rec(1, SW_STATUS) = "Active"
    rec(1, SW_USER) = CurrentUser()
    rec(1, SW_STAMP) = Now

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim em As String
    On Error GoTo Fail
    ws.Unprotect Password:=SHEET_PASSWORD
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, SW_N).Value = rec
    If payFor <> "" Then SetPaidBy lo, payFor, id
    ProtectSheet ws
    ' calendar line: "Paul M S/S" -> "Josh L S/S"
    Dim calErr As String
    On Error Resume Next
    If AddCoverage(d, shiftLbl, unit, "NEW", CalName(worksName) & " S/S", Empty, id, "SS", CalName(offName) & " S/S") = "" Then
        calErr = IIf(Err.Number <> 0, Err.Description, "the line wasn't written")
    End If
    Err.Clear
    On Error GoTo 0
    SaveIfWritable
    If calErr <> "" Then
        MsgBox id & " is saved, but it couldn't be added to the calendar: " & calErr & vbNewLine & _
               "Opening the Shift Switch pop-up again retries it.", vbExclamation
    End If
    RefreshCalendar
    SaveSwitch = id
    Exit Function
Fail:
    em = Err.Description
    LogError "SaveSwitch", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Could not save the shift switch: " & em, vbCritical
End Function

Private Sub SetPaidBy(ByVal lo As ListObject, ByVal origId As String, ByVal newId As String)
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, SW_ID))), origId, vbTextCompare) = 0 Then
            lo.DataBodyRange.Cells(i, SW_PAIDBY).Value = newId
            Exit Sub
        End If
    Next i
End Sub

Public Function VoidSwitch(ByVal id As String, Optional ByVal reasonTxt As String = "") As Boolean
    ' reasonTxt (v24.2): given by code (SmokeTest); left out, the reason is asked for
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    Dim lo As ListObject: Set lo = SwitchTable()
    If lo Is Nothing Then Exit Function
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, r As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, SW_ID))), id, vbTextCompare) = 0 Then r = i: Exit For
    Next i
    If r = 0 Then Exit Function
    If UCase$(SafeStr(v(r, SW_STATUS))) = "VOID" Then MsgBox id & " is already void.", vbInformation: Exit Function

    Dim why As String: why = reasonTxt
    If Trim$(why) = "" Then
        why = InputBox("Why is " & id & " being voided?" & vbNewLine & _
                       Format$(v(r, SW_DATE), "ddd mmm d") & " " & SafeStr(v(r, SW_SHIFT)) & ": " & _
                       SafeStr(v(r, SW_WORKS)) & " for " & SafeStr(v(r, SW_OFF)), "Void shift switch")
    End If
    If Trim$(why) = "" Then Exit Function

    Dim ws As Worksheet: Set ws = lo.Parent
    ws.Unprotect Password:=SHEET_PASSWORD
    lo.DataBodyRange.Cells(r, SW_STATUS).Value = "VOID"
    lo.DataBodyRange.Cells(r, SW_VOIDNOTE).Value = Format$(Now, "yyyy-mm-dd hh:nn") & " " & CurrentUser() & ": " & Trim$(why)
    ' a voided payback no longer pays the original back
    Dim payFor As String: payFor = Trim$(SafeStr(v(r, SW_PAYFOR)))
    If payFor <> "" Then
        For i = 1 To UBound(v, 1)
            If StrComp(Trim$(SafeStr(v(i, SW_ID))), payFor, vbTextCompare) = 0 Then
                If StrComp(Trim$(SafeStr(v(i, SW_PAIDBY))), id, vbTextCompare) = 0 Then
                    lo.DataBodyRange.Cells(i, SW_PAIDBY).ClearContents
                End If
            End If
        Next i
    End If
    ProtectSheet ws
    On Error Resume Next
    VoidCoverageBySource id
    On Error GoTo 0
    SaveIfWritable
    RefreshCalendar
    VoidSwitch = True
End Function
