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

Public Const SS_VERSION As String = "Shift switches v27.5 (Oct 25)"
Public Const SS_SHEET As String = "Shift Switches"
Public Const SS_TABLE As String = "tblShiftSwitch"
Public Const LEDGER_SHEET As String = "Paybacks"     ' v27 (item 25): the payback ledger, output only (supervisors and admins)

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

' v27.1: the facts SwitchSaveStops reads, by position, for the pure SwitchStopsCore (suite switchstop). A booked-off fact is
' "" or "+" & the booked type; a switch fact the switch's ID or ""; a 16 h fact FatigueShiftReason's answer (v27.5; NotFreeReason's
' before); an "on" fact "+".
Public Const SF_BKOFF As Long = 0           ' the one giving up: booked off on the shift
Public Const SF_GAVEOFF As Long = 1         '   gave up a shift that day (ID)
Public Const SF_WORKOFF As Long = 2         '   works a switch that day (ID)
Public Const SF_BKWORKS As Long = 3         ' the one working it: booked off on the shift
Public Const SF_WORKWORKS As Long = 4       '   works a switch that day (ID)
Public Const SF_FAT As Long = 5             '   NotFreeReason on the shift (the 16 h stop)
Public Const SF_P_BKWORKS As Long = 6       ' two-way, another day: the payback - the one working the switch now gives up
Public Const SF_P_GAVEWORKS As Long = 7
Public Const SF_P_WORKWORKS As Long = 8
Public Const SF_P_BKOFF As Long = 9         '   and the one who gave up works it
Public Const SF_P_WORKOFF As Long = 10
Public Const SF_P_FAT As Long = 11
Public Const SF_S_BKOFF As Long = 12        ' two-way, the same day: the one giving up booked off on the shift
Public Const SF_S_BKWORKS As Long = 13      '   the one working it booked off on the other shift
Public Const SF_S_BKOFF2 As Long = 14       '   the one giving up booked off on the other shift (the one they work)
Public Const SF_S_BKWORKS1 As Long = 15     '   the one working it booked off on the shift (the one they work)
Public Const SF_S_OFFON As Long = 16        '   the one giving up is on any switch that day
Public Const SF_S_WORKSON As Long = 17      '   the one working it is on any switch that day
Public Const SF_S_FATWORKS As Long = 18     ' v27.5 (item 1): the one working it - the 16 h rule on the shift (sh2 given up)
Public Const SF_S_FATOFF As Long = 19       '   the one giving up - the 16 h rule on the other shift (the one they work; shiftLbl given up)
Public Const SF_N As Long = 19

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
    If Err.Number <> 0 Then LogError "SyncSwitchesToCalendar", Err.Number, Err.Description      ' v25 (I5): not swallowed any more
    On Error Resume Next
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
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SS_SHEET
        RelockStructure
        SetTabColor ws, RGB(180, 100, 0)                ' v25.1
    End If
    UnlockSheet ws
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

' v26: True while a batch is on, so a sweep inside another sweep neither re-reads nor ends the outer batch
Public Function SwitchBatchActive() As Boolean
    SwitchBatchActive = mBatch
End Function

' v26 (the staffing sweep): the names on any non-void switch dated d1 .. d2 (the one off and the one
' working), as a Dictionary (name -> True; case-insensitive). One read of the table.
Public Function SwitchNamesBetween(ByVal d1 As Date, ByVal d2 As Date) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set SwitchNamesBetween = d
    Dim v As Variant: v = SwitchData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, nm As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID" And IsDate(v(i, SW_DATE)) Then
            If Int(CDate(v(i, SW_DATE))) >= Int(d1) And Int(CDate(v(i, SW_DATE))) <= Int(d2) Then
                nm = Trim$(SafeStr(v(i, SW_OFF))): If nm <> "" Then d(nm) = True
                nm = Trim$(SafeStr(v(i, SW_WORKS))): If nm <> "" Then d(nm) = True
            End If
        End If
    Next i
End Function

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

' Problems with "works" covering the shift (blank = looks fine). v25: the one rule
' (NotFreeReason: the rotation on d - 1, d, d + 1 from the crew history, a given-up shift
' relieving the rest rule); only "scheduled" and "rest" are warnings here - a booking or a
' switch that day is a hard stop in EligibilityProblems.
Public Function WorksConflict(ByVal empName As String, ByVal d As Date, ByVal isNight As Boolean) As String
    Dim shiftLbl As String: shiftLbl = IIf(isNight, "Nights", "Days")
    Dim code As String: code = EmpCrewCode(empName, d)
    ' v26.5.5 (item 9): a switch is the whole shift, so the whole shift must fit the 16 h rule
    Select Case NotFreeReason(empName, code, d, shiftLbl, DefaultShiftHours(d, isNight, USUAL_SHIFT_HOURS))
        Case "scheduled"
            WorksConflict = "is already working " & LCase$(CrewShiftOn(CrewOn(empName, d, code), d)) & " that day"
        Case "rest"
            If isNight Then WorksConflict = "starts days the next morning" Else WorksConflict = "is coming off nights that morning"
        Case "16 h rule"
            WorksConflict = "can't work it under the 16 h rule (with the shifts beside it - scheduled / OT / a switch - it would run past " & FATIGUE_MAX_HOURS & " h or leave under " & FATIGUE_MIN_REST & " h rest)"
        Case "on OT"
            WorksConflict = "already has OT accepted on that shift"
    End Select
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

' SS ID if empName GAVE UP a shift on date d ("" if not); shiftLbl "" = any shift (v25: or that one)
Public Function SwitchOffOn(ByVal empName As String, ByVal d As Date, Optional ByVal shiftLbl As String = "") As String
    If mBatch Then
        SwitchOffOn = MatchSwitch(mBatchRows, SW_OFF, empName, d, shiftLbl)
    Else
        Dim v As Variant: v = SwitchData()
        SwitchOffOn = MatchSwitch(v, SW_OFF, empName, d, shiftLbl)
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
' v25 (D-L1): shiftLbl = the shift switched; a sick window blocks the shifts inside it only.
' v27.1: the reads stay here, in the same order (t carries over between the two bookings, as it always did); the
' words are the pure EligibilityText.
Public Function EligibilityProblems(ByVal offName As String, ByVal worksName As String, ByVal d As Date, _
                                    Optional ByVal shiftLbl As String = "") As String
    Dim t As String, bkOff As String, gaveOff As String, workOff As String, bkWorks As String, workWorks As String
    If BookedOffShift(offName, d, shiftLbl, t) Then bkOff = "+" & t
    gaveOff = SwitchOffOn(offName, d)
    workOff = SwitchWorkingOn(offName, d, "")
    If BookedOffShift(worksName, d, shiftLbl, t) Then bkWorks = "+" & t
    workWorks = SwitchWorkingOn(worksName, d, "")
    EligibilityProblems = EligibilityText(offName, worksName, Format$(d, "mmm d"), bkOff, gaveOff, workOff, bkWorks, workWorks)
End Function

' v27.1 (pure; suite switchstop): EligibilityProblems' stops from what it read - bk* = "" or "+" & the booked type,
' gaveOff / workOff / workWorks = the switch's ID or ""; dTxt = the date as "mmm d"
Public Function EligibilityText(ByVal offName As String, ByVal worksName As String, ByVal dTxt As String, _
                                ByVal bkOff As String, ByVal gaveOff As String, ByVal workOff As String, _
                                ByVal bkWorks As String, ByVal workWorks As String) As String
    Dim s As String
    If bkOff <> "" Then s = s & "- " & offName & " is booked off (" & Mid$(bkOff, 2) & ") on " & dTxt & vbNewLine
    If gaveOff <> "" Then s = s & "- " & offName & " already switched off that day (" & gaveOff & ")" & vbNewLine
    If workOff <> "" Then s = s & "- " & offName & " is already working a switch that day (" & workOff & ")" & vbNewLine
    If bkWorks <> "" Then s = s & "- " & worksName & " is booked off (" & Mid$(bkWorks, 2) & ") on " & dTxt & vbNewLine
    If workWorks <> "" Then s = s & "- " & worksName & " is already working a switch that day (" & workWorks & ")" & vbNewLine
    EligibilityText = s
End Function

' Short status for the pop-up hint ("booked off (Vac)", "on SS-0003") or "".
' kind (optional) says which: "OFF" booked off, "GAVE_UP" already gave up a
' shift that day, "WORKING" already working a switch, "" none. The form
' checks kind, never the wording.
Public Function PersonFlag(ByVal empName As String, ByVal d As Date, Optional ByRef kind As String, _
                           Optional ByVal shiftLbl As String = "") As String
    Dim t As String, id As String
    kind = ""
    If BookedOffShift(empName, d, shiftLbl, t) Then PersonFlag = "booked off (" & t & ")": kind = "OFF": Exit Function
    id = SwitchOffOn(empName, d)
    If id <> "" Then PersonFlag = "already switched off (" & id & ")": kind = "GAVE_UP": Exit Function
    id = SwitchWorkingOn(empName, d, "")
    If id <> "" Then PersonFlag = "already working " & id: kind = "WORKING"
End Function

' ---------------- v25: the switch pop-up's rules, out of the form ----------------
' "Fred C" for "Campbell, Fred" (the pop-up's short form)
Public Function SwitchShortName(ByVal nm As String) As String
    Dim p As Long: p = InStr(nm, ",")
    If p > 0 Then SwitchShortName = Trim$(Mid$(nm, p + 1)) & " " & Left$(nm, 1) Else SwitchShortName = nm
End Function

' Can the payback physically happen on d2? "" = yes.
'  - the one being paid back (today's "works") must be scheduled that day (sh2)
'  - the one paying back (today's "gives up") must be free for that shift:
'    not scheduled, not under the rest rule, not booked off, not on another switch
' A same-day sister swap is exempt from the "scheduled" check (they trade shifts).
' v25 review (D-C3): shiftLbl = the shift given up on d in THIS save; when the payback is the adjoining
' shift (nights given up, days paid back next morning, or the mirror image) the rest rule is relieved
' here too - the switch that relieves it isn't saved yet, so NotFreeReason can't see it.
Public Function PaybackProblem(ByVal offName As String, ByVal worksName As String, ByVal d As Date, _
                               ByVal d2 As Date, ByVal sh2 As String, Optional ByVal shiftLbl As String = "") As String
    If sh2 = "" Then PaybackProblem = SwitchShortName(worksName) & " isn't scheduled that day - no shift to pay back": Exit Function
    Dim who As String: who = SwitchShortName(offName)
    Dim t As String
    If Int(d2) = Int(d) Then
        If BookedOffShift(offName, d2, sh2, t) Then PaybackProblem = who & " is booked off (" & t & ")"
        Exit Function
    End If
    Dim code As String: code = EmpCrewCode(offName, d2)          ' v24: their crew on the payback date
    ' v26.5.5: the payback is a whole shift, so it must fit the 16 h rule
    Dim why As String: why = NotFreeReason(offName, code, d2, sh2, DefaultShiftHours(d2, IsNightShift(sh2), USUAL_SHIFT_HOURS))
    If why = "rest" Then
        If sh2 = "Days" And shiftLbl = "Nights" And Int(d2) = Int(d) + 1 Then why = ""
        If sh2 = "Nights" And shiftLbl = "Days" And Int(d2) = Int(d) - 1 Then why = ""
    End If
    Select Case why
        Case "", "S/S off"
        Case "scheduled": PaybackProblem = who & " is scheduled to work " & LCase$(CrewShiftOn(code, d2)) & " that day"
        Case "rest"
            If sh2 = "Days" Then
                PaybackProblem = who & " is coming off nights that morning"
            Else
                PaybackProblem = who & " starts days the next morning"
            End If
        Case "S/S": PaybackProblem = who & " is already working a switch that day"
        Case "16 h rule": PaybackProblem = who & " can't work that shift under the 16 h rule (with the shifts beside it - scheduled / OT / a switch - it would run past " & FATIGUE_MAX_HOURS & " h or leave under " & FATIGUE_MIN_REST & " h rest)"
        Case "on OT": PaybackProblem = who & " already has OT accepted on that shift"
        Case Else: PaybackProblem = who & " is booked off (" & why & ")"
    End Select
End Function

' v26.5.5 (review): a switch is a whole shift, so the one WORKING it must fit the 16 h rule - a hard stop
' ("" = fine). v27.5 (item 1, Graeme Oct 25): the rule's own arithmetic (modCalendar.FatigueShiftReason), so a
' SCHEDULED shift beside the one worked counts like OT accepted or a switch - nights last night then days today is a
' stop, no longer the "coming off nights" warning alone. A same-day sister swap reads it too but SwitchStopsCore keeps
' only "on OT" from it (the swap of both shifts is exempt from the arithmetic: Graeme, Oct 25). The shift(s) the
' switch gives up are set aside by SwitchSaveStops (SetPendingGiveUp) before this reads, so a two-way payback on the
' adjoining shift is not stopped by the shift it gives up.
' v27.1: what the rule reads - the crew code, then the reason ("" when no shift is named); FatigueStopText words it.
Private Function FatigueReason(ByVal worksName As String, ByVal d As Date, ByVal shiftLbl As String) As String
    If shiftLbl = "" Then Exit Function
    Dim code As String: code = EmpCrewCode(worksName, d)
    FatigueReason = FatigueShiftReason(worksName, code, d, shiftLbl)
End Function

' v27.1 (pure; suite switchstop): the 16 h stop from FatigueShiftReason's answer ("" = fine; NotFreeReason's before v27.5) -
' only the 16 h rule and OT already held on the shift stop it; dTxt = the date as "mmm d"
Public Function FatigueStopText(ByVal worksName As String, ByVal dTxt As String, ByVal shiftLbl As String, ByVal reason As String) As String
    If shiftLbl = "" Then Exit Function
    Select Case reason
        Case "16 h rule"      ' v27.5 (item 1): a scheduled shift beside it counts too, so the words no longer say "OT accepted"
            FatigueStopText = "- " & worksName & " can't work " & dTxt & " " & LCase$(shiftLbl) & _
                              " under the 16 h rule (with the shifts beside it - scheduled / OT / a switch - it would run past " & _
                              FATIGUE_MAX_HOURS & " h or leave under " & FATIGUE_MIN_REST & " h rest)" & vbNewLine
        Case "on OT"
            FatigueStopText = "- " & worksName & " already has OT accepted on " & dTxt & " " & LCase$(shiftLbl) & vbNewLine
    End Select
End Function

' The hard stops before a switch is saved ("" = none): booked off / already on a switch, and the 16 h rule for the
' one working it, for the switch itself and, two-way, for its payback (a same-day sister swap: each is the other's
' switch, so bookings and other switches that day, and - v27.5 - OT already held on the shift each one works; the 16 h
' arithmetic is not applied to that swap: Graeme, Oct 25).
' v27.1: every fact is read here in the order the stops always read them (so an error comes from the same read; a
' same-day swap still reads the one-way facts first, then sets them aside), and the pure SwitchStopsCore words them.
' v27.5 (item 1): the shift(s) this switch gives up are not worked - named to modCalendar (SetPendingGiveUp) before the
' 16 h facts are read, so neither counts against the one who gives it up; cleared on every exit (the handler too).
Public Function SwitchSaveStops(ByVal offName As String, ByVal worksName As String, ByVal d As Date, ByVal shiftLbl As String, _
                                ByVal twoWay As Boolean, ByVal d2 As Date, ByVal sh2 As String) As String
    Dim f(0 To SF_N) As String, t As String, t2 As String, tB As String, sameDay As Boolean
    On Error GoTo Oops
    ClearPendingGiveUp
    SetPendingGiveUp 1, offName, d, shiftLbl
    If twoWay Then SetPendingGiveUp 2, worksName, d2, sh2
    If BookedOffShift(offName, d, shiftLbl, t) Then f(SF_BKOFF) = "+" & t          ' (EligibilityProblems' reads, its t carried over)
    f(SF_GAVEOFF) = SwitchOffOn(offName, d)
    f(SF_WORKOFF) = SwitchWorkingOn(offName, d, "")
    If BookedOffShift(worksName, d, shiftLbl, t) Then f(SF_BKWORKS) = "+" & t
    f(SF_WORKWORKS) = SwitchWorkingOn(worksName, d, "")
    f(SF_FAT) = FatigueReason(worksName, d, shiftLbl)                               ' v26.5.5: the whole shift must fit the 16 h rule
    If twoWay Then
        sameDay = (Int(d2) = Int(d))
        If sameDay Then
            If BookedOffShift(offName, d, shiftLbl, t2) Then f(SF_S_BKOFF) = "+" & t2
            If BookedOffShift(worksName, d, sh2, t2) Then f(SF_S_BKWORKS) = "+" & t2
            ' v25 review (D-L1): the shift each one WORKS in the swap is inside no sick window either
            If BookedOffShift(offName, d, sh2, t2) Then f(SF_S_BKOFF2) = "+" & t2
            If BookedOffShift(worksName, d, shiftLbl, t2) Then f(SF_S_BKWORKS1) = "+" & t2
            If SwitchOffOn(offName, d) <> "" Or SwitchWorkingOn(offName, d, "") <> "" Then f(SF_S_OFFON) = "+"
            If SwitchOffOn(worksName, d) <> "" Or SwitchWorkingOn(worksName, d, "") <> "" Then f(SF_S_WORKSON) = "+"
            f(SF_S_FATWORKS) = FatigueReason(worksName, d, shiftLbl)                ' v27.5 (item 1): each one on the shift they WORK,
            f(SF_S_FATOFF) = FatigueReason(offName, d, sh2)                         '   the shift each gives up set aside
        Else
            If BookedOffShift(worksName, d2, sh2, tB) Then f(SF_P_BKWORKS) = "+" & tB   ' (EligibilityProblems(works, off, d2, sh2)'s reads)
            f(SF_P_GAVEWORKS) = SwitchOffOn(worksName, d2)
            f(SF_P_WORKWORKS) = SwitchWorkingOn(worksName, d2, "")
            If BookedOffShift(offName, d2, sh2, tB) Then f(SF_P_BKOFF) = "+" & tB
            f(SF_P_WORKOFF) = SwitchWorkingOn(offName, d2, "")
            f(SF_P_FAT) = FatigueReason(offName, d2, sh2)                           ' v26.5.5: the payback shift too
        End If
    End If
    ClearPendingGiveUp
    SwitchSaveStops = SwitchStopsCore(offName, worksName, Format$(d, "mmm d"), shiftLbl, twoWay, sameDay, Format$(d2, "mmm d"), sh2, f)
    Exit Function
Oops:
    Dim en As Long, em As String: en = Err.Number: em = Err.Description
    ClearPendingGiveUp
    Err.Raise en, , em
End Function

' v27.1 (pure; suite switchstop): the hard stops from the facts SwitchSaveStops read (positions SF_*). One-way: the
' eligibility stops, then the 16 h stop for the one working it. Two-way on another day: the same again for the payback,
' the roles swapped. A same-day sister swap: the one-way stops are set aside (each one IS scheduled that day - on the
' shift they give up); bookings and other switches that day, then (v27.5, item 1) only "on OT" from the two 16 h facts
' (SF_S_FATWORKS / SF_S_FATOFF): the 16 h arithmetic is NOT applied to the swap of both shifts (Graeme, Oct 25).
Public Function SwitchStopsCore(ByVal offName As String, ByVal worksName As String, ByVal dTxt As String, ByVal shiftLbl As String, _
                                ByVal twoWay As Boolean, ByVal sameDay As Boolean, ByVal d2Txt As String, ByVal sh2 As String, _
                                ByRef f() As String) As String
    Dim stops As String
    stops = EligibilityText(offName, worksName, dTxt, f(SF_BKOFF), f(SF_GAVEOFF), f(SF_WORKOFF), f(SF_BKWORKS), f(SF_WORKWORKS))
    stops = stops & FatigueStopText(worksName, dTxt, shiftLbl, f(SF_FAT))
    If twoWay Then
        If sameDay Then
            stops = ""
            If f(SF_S_BKOFF) <> "" Then stops = stops & "- " & offName & " is booked off (" & Mid$(f(SF_S_BKOFF), 2) & ")" & vbNewLine
            If f(SF_S_BKWORKS) <> "" Then stops = stops & "- " & worksName & " is booked off (" & Mid$(f(SF_S_BKWORKS), 2) & ")" & vbNewLine
            If f(SF_S_BKOFF2) <> "" Then stops = stops & "- " & offName & " can't work the " & LCase$(sh2) & " shift (" & Mid$(f(SF_S_BKOFF2), 2) & ")" & vbNewLine
            If f(SF_S_BKWORKS1) <> "" Then stops = stops & "- " & worksName & " can't work the " & LCase$(shiftLbl) & " shift (" & Mid$(f(SF_S_BKWORKS1), 2) & ")" & vbNewLine
            If f(SF_S_OFFON) <> "" Then stops = stops & "- " & offName & " is already on a switch that day" & vbNewLine
            If f(SF_S_WORKSON) <> "" Then stops = stops & "- " & worksName & " is already on a switch that day" & vbNewLine
            ' v27.5 (item 1; Graeme Oct 25 after review): the swap of BOTH shifts is the one way an A/B pair can swap, and under
            ' the rotation one of the two always has the adjoining shift - so the 16 h arithmetic is set aside here; OT already
            ' held (or an award) on the shift each one would work still stops it
            If f(SF_S_FATWORKS) = "on OT" Then stops = stops & FatigueStopText(worksName, dTxt, shiftLbl, f(SF_S_FATWORKS))
            If f(SF_S_FATOFF) = "on OT" Then stops = stops & FatigueStopText(offName, dTxt, sh2, f(SF_S_FATOFF))
        Else
            stops = stops & EligibilityText(worksName, offName, d2Txt, f(SF_P_BKWORKS), f(SF_P_GAVEWORKS), f(SF_P_WORKWORKS), f(SF_P_BKOFF), f(SF_P_WORKOFF))
            stops = stops & FatigueStopText(offName, d2Txt, sh2, f(SF_P_FAT))
        End If
    End If
    SwitchStopsCore = stops
End Function

' The rotation warnings shown in the confirm ("" = none): the one giving up isn't on that shift;
' the one working is scheduled or under the rest rule (not for a same-day sister swap).
Public Function SwitchWarnings(ByVal offName As String, ByVal worksName As String, ByVal d As Date, _
                               ByVal isNight As Boolean, ByVal sameDaySwap As Boolean) As String
    Dim warn As String, s As String
    s = OffConflict(offName, d, isNight)
    If s <> "" Then warn = warn & "- " & offName & ": " & s & vbNewLine
    If Not sameDaySwap Then
        s = WorksConflict(worksName, d, isNight)
        If s <> "" Then warn = warn & "- " & worksName & ": " & s & vbNewLine
    End If
    SwitchWarnings = warn
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

' v26 (payback aging): how many owed paybacks are older than PAYBACK_WARN_DAYS (the Desk tile)
Public Function OwedOverdueCount() As Long
    Dim txt As String
    OwedOverdueCount = OwedOlderThan(PaybackWarnDays(), txt)
End Function

' v26 (payback aging): the Needs Attention status for a payback owed for `days` days.
' Pure; mirrored (tests/rules.py owed_status_text, suite "payback").
Public Function OwedStatusText(ByVal days As Long, ByVal warnDays As Long) As String
    If days > warnDays Then
        OwedStatusText = "Payback owed " & days & IIf(days = 1, " day", " days")
    Else
        OwedStatusText = "Payback owed"
    End If
End Function

' Switches still OWED a payback (original switches with no payback logged),
' in table order (Needs Attention sorts them oldest first, v26). 2-D array
' (rows x 6) for the pop-up list; ids() gets the IDs.
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
    UnlockSheet ws
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
               "Opening the Shift Switch pop-up again retries it.", vbExclamation, "Call Out Roster"
    End If
    RefreshCalendar
    SaveSwitch = id
    Exit Function
Fail:
    em = Err.Description
    LogError "SaveSwitch", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Could not save the shift switch: " & em, vbCritical, "Call Out Roster"
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
    If UCase$(SafeStr(v(r, SW_STATUS))) = "VOID" Then MsgBox id & " is already void.", vbInformation, "Call Out Roster": Exit Function

    Dim why As String: why = reasonTxt
    If Trim$(why) = "" Then
        why = InputBox("Why is " & id & " being voided?" & vbNewLine & _
                       Format$(v(r, SW_DATE), "ddd mmm d") & " " & SafeStr(v(r, SW_SHIFT)) & ": " & _
                       SafeStr(v(r, SW_WORKS)) & " for " & SafeStr(v(r, SW_OFF)), "Void shift switch")
    End If
    If Trim$(why) = "" Then Exit Function

    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
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

' ====================================================================
'  THE PAYBACK LEDGER (v27 item 25, 1.4)
' ====================================================================
' Every payback still owed (an original switch with no payback logged): who gave up a shift owes the one who
' worked it. Everyone = one table by the person who owes, oldest first, with its age and the
' over-PAYBACK_WARN_DAYS status in amber; a name typed = that person's statement (what they owe, what they are
' owed), with a PDF beside the roster on request. Supervisors and admins; changes no data; never published.
Public Sub SwitchLedger()
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Oops
    If Not IsSupervisor() Then MsgBox "The payback ledger is for supervisors and admins.", vbInformation, "Paybacks owed": Exit Sub
    Dim a As Variant
    a = Application.InputBox(Prompt:="One person's statement: part of the name." & vbNewLine & "Leave it blank for everyone.", Title:="Paybacks owed", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim who As String
    If Trim$(CStr(a)) <> "" Then
        who = ResolvePersonName(Trim$(CStr(a)), "Paybacks owed")
        If who = "" Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim v As Variant: v = SwitchData()
    Dim idx() As Long, n As Long, i As Long, j As Long, t As Long
    If IsArray(v) Then
        ReDim idx(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            If IsOwed(v, i) And IsDate(v(i, SW_DATE)) Then
                If who = "" Or StrComp(Trim$(SafeStr(v(i, SW_OFF))), who, vbTextCompare) = 0 _
                   Or StrComp(Trim$(SafeStr(v(i, SW_WORKS))), who, vbTextCompare) = 0 Then
                    n = n + 1: idx(n) = i
                End If
            End If
        Next i
        For i = 2 To n                              ' by the person who owes, then oldest first
            t = idx(i): j = i - 1
            Do While j >= 1
                If LedgerKey(v, idx(j)) <= LedgerKey(v, t) Then Exit Do
                idx(j + 1) = idx(j): j = j - 1
            Loop
            idx(j + 1) = t
        Next i
    End If
    Dim ws As Worksheet: Set ws = ReportSheetNamed(LEDGER_SHEET)
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    Dim warn As Long: warn = PaybackWarnDays()
    ws.Range("A1").Value = IIf(who = "", "PAYBACKS OWED  -  everyone", "PAYBACK STATEMENT  -  " & who)
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "'As of " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & ".  Over " & warn & " days in amber.  Whoever gave up a shift owes the one who worked it."
    Dim r As Long: r = 4
    If who = "" Then
        LedgerPart ws, r, "OWED PAYBACKS (by the person who owes)", v, idx, n, "", warn
    Else
        LedgerPart ws, r, who & " OWES (gave up a shift)", v, idx, n, "off:" & who, warn
        r = r + 1
        LedgerPart ws, r, who & " IS OWED (worked someone's shift)", v, idx, n, "works:" & who, warn
    End If
    ws.Columns(1).ColumnWidth = 24
    ws.Columns(2).ColumnWidth = 24
    For i = 3 To 8
        ws.Columns(i).ColumnWidth = 12
    Next i
    ws.Columns(7).ColumnWidth = 22
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = "Paybacks owed - " & Format$(Now, "ddd mmm d, yyyy") & " - page &P of &N"
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    ws.Activate
    If Ask(n & " payback(s) owed" & IIf(who = "", "", " involving " & who) & ". The Paybacks tab lists them." & vbNewLine & vbNewLine & _
           "Save a PDF beside the roster as well?", vbYesNo + vbQuestion, "Paybacks owed", vbNo) = vbYes Then
        ReportPdf ws, IIf(who = "", "Paybacks owed", "Paybacks " & CalName(who)), "Paybacks owed", True
    End If
    Exit Sub
Oops:
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    ReportError "SwitchLedger"
End Sub

' the sort key: the person who owes, then the switch date
Private Function LedgerKey(ByVal v As Variant, ByVal i As Long) As String
    LedgerKey = LCase$(Trim$(SafeStr(v(i, SW_OFF)))) & "|" & Format$(CDate(v(i, SW_DATE)), "yyyymmdd")
End Function

' one block of the ledger; only "off:<name>" / "works:<name>" rows when a filter is given
Private Sub LedgerPart(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String, ByVal v As Variant, ByRef idx() As Long, _
                       ByVal n As Long, ByVal filt As String, ByVal warn As Long)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(30, 56, 100)
    r = r + 1
    Dim heads As Variant: heads = Array("Owes (gave up)", "Owed to (worked it)", "Switch date", "Shift", "Unit", "Days", "Status", "Switch")
    Dim j As Long
    For j = 0 To 7
        ws.Cells(r, 1 + j).Value = heads(j)
    Next j
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Font.Bold = True: .Font.Size = 9
        .Interior.Color = RGB(222, 229, 240)
    End With
    r = r + 1
    Dim k As Long, i As Long, days As Long, shown As Long, keep As Boolean, nm As String
    For k = 1 To n
        i = idx(k)
        keep = True
        If Left$(filt, 4) = "off:" Then
            nm = Mid$(filt, 5)
            keep = (StrComp(Trim$(SafeStr(v(i, SW_OFF))), nm, vbTextCompare) = 0)
        ElseIf Left$(filt, 6) = "works:" Then
            nm = Mid$(filt, 7)
            keep = (StrComp(Trim$(SafeStr(v(i, SW_WORKS))), nm, vbTextCompare) = 0)
        End If
        If keep Then
            days = Int(Date) - Int(CDate(v(i, SW_DATE)))
            ws.Cells(r, 1).Value = "'" & SafeStr(v(i, SW_OFF))
            ws.Cells(r, 2).Value = "'" & SafeStr(v(i, SW_WORKS))
            ws.Cells(r, 3).Value = "'" & Format$(CDate(v(i, SW_DATE)), "ddd mmm d, yyyy")
            ws.Cells(r, 4).Value = "'" & SafeStr(v(i, SW_SHIFT))
            ws.Cells(r, 5).Value = "'" & SafeStr(v(i, SW_UNIT))
            ws.Cells(r, 6).Value = days
            ws.Cells(r, 7).Value = "'" & OwedStatusText(days, warn)
            ws.Cells(r, 8).Value = "'" & SafeStr(v(i, SW_ID))
            If days > warn Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 8)).Interior.Color = RGB(255, 235, 156)
            r = r + 1
            shown = shown + 1
        End If
    Next k
    If shown = 0 Then
        ws.Cells(r, 1).Value = "None."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If
End Sub
