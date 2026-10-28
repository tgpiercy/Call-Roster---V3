Attribute VB_Name = "modMockData"
Option Explicit

' ====================================================================
'  TEST-ONLY - THE ALPHA'S MOCK PLANT (P9, Nov 1-14, 2026). NEVER SHIPS. NEVER ON THE REAL DATA.
' ====================================================================
' It lives in tests/mock/, not modules/: build_zip packs it as test-only/ (outside modules\, so the Install panel never
' installs it) and SystemCheck never counts it. It calls only Public procedures that exist in v29. Nothing in it writes
' a roster formula by hand or edits an OT Log row: the people go in through AddRosterRow, the log rows through
' AppendLogRows, every other table through its ListObject with the roster's own append idiom.
'
' LoadMockData   - on a PRACTICE copy (COPY / DRY RUN in the name), or on the live file after the word MOCK is typed:
'                  a dated "Backup before MOCK LOAD ..." copy first; then every DATA table and the roster's people rows
'                  are emptied and the mock plant written (69 people, the OT Log of periods 21-22, the bookings, the
'                  switch, the postings, the planned move, the restriction, the Hours History, the ID marks); Worked (H)
'                  is the Oct 18 snapshot's figure and KronosDataThrough is Oct 17 - the Day-0 Kronos import brings
'                  both to Oct 30. VERIFIED before the save: every person's G / I / period sum against expected_load,
'                  the log's charged total, RunIntegrity 0 PROBLEM; a mismatch blocks the save (MarkNoSave) and names
'                  the first person that differs. The two carried-in sign-up files for P-0003 are written last.
' ResetRosterData - the wipe alone (the same list, the roster rows too), verified empty. For the rehearsal of the loader
'                  on a practice copy. The LIVE reset at the close-out is a file swap (docs/alpha-data-reset.md).
' MockSignupFile / MockRequestFile - write one sign-up / request file the way the members' file does (member/modMember.bas
'                  SendSignup / WriteRequestFile), for a plan step with no member tester at hand.
'
' The data between the MOCK DATA markers is generated from tests/mock/*.csv by tests/mock/gen_mock_loader.py: run it
' after changing a CSV; tests/test_mock.py checks the two agree. Settings sheets (Min Staff, Holidays, Rules,
' Settings) are never touched.
' ====================================================================

Public Const MOCK_VERSION As String = "Mock plant P9 (Oct 28, 2026) for v29"
Private Const MOCK_WORD As String = "MOCK"
Private Const TTL As String = "Mock plant"
Private Const SEP As String = "|"
Private Const CHECK_PEOPLE As Long = 69
Private mRows As Collection

' ====================================================================
'  LOAD
' ====================================================================
Public Sub LoadMockData()
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long, em As String, bk As String
    If Not MockGuard("LoadMockData") Then Exit Sub
    If Not PeopleBackup("MOCK LOAD", bk) Then Exit Sub
    QuietOn evWas, scrWas, calcWas
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ClearAll
    WritePeople
    WriteLog
    WriteTable CrewTable(True), "crew", "T,T,D,T,S"
    WriteTable AbsTable(True), "absences", "T,T,T,T,D,D,N,T,T,T,T,S,T,T,T,T"
    WriteTable CovTable(True), "coverage", "T,D,T,T,T,T,T,T,N,T,T,T,S,T"
    WriteTable SwitchTable(True), "switches", "T,D,T,T,T,T,T,T,T,T,T,T,T,S,T"
    WriteTable GetTable(PT_TABLE), "postings", "T,T,D,D,D,T,T,T,N,N,R,T,D,T,T"
    WriteTable GetTable(PL_TABLE), "postinglist", "T,N,T,T,T,N,T,T,N,N"
    WriteTable PlannedTable(True), "planned", "T,T,T,T,T,T,T,D,T,T,S,D,T"
    WriteTable RstTable(True), "restrictions", "T,T,D,D,R,T,T,S,T"
    WriteTable HoursHistoryTable(True), "hourshistory", "T,S,D,T,N,T,D,N"
    WriteMarks
    SetDateName PROV_NAME, DateSerial(2026, 10, 17)      ' the Oct 18 snapshot's Data Through; the Day-0 import moves it
    EnsureRefusedThrough
    EnsureKeyColumns
    FillPayrollKeys True, True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    RefreshGradeSnapshot
    AfterPeopleChange
    ResetHolidayCache
    RefreshCalCache
    MarkMembersDirty
    Dim problem As String: problem = VerifyLoad()
    QuietOff evWas, scrWas, calcWas
    If problem <> "" Then
        MarkNoSave
        MsgBox "The mock plant does NOT match its expected figures - saving is blocked. Close WITHOUT saving, then look at:" & vbNewLine & vbNewLine & _
               problem, vbCritical, TTL
        Exit Sub
    End If
    Dim nFiles As Long: nFiles = WriteCarriedSignups()
    SaveIfWritable
    MsgBox "Mock plant loaded (" & MOCK_VERSION & ")." & vbNewLine & vbNewLine & _
           "People " & (RosterLastRow() - R_FIRST + 1) & ", OT Log rows " & LogRowCount() & ", every figure matches expected_load, IntegrityCheck 0 PROBLEM." & vbNewLine & _
           "Backup copy: " & bk & vbNewLine & _
           "Sign-up files written for P-0003: " & nFiles & vbNewLine & vbNewLine & _
           "Next (the plan, Day 0): Import Kronos (the Day-0 cumulative report), SystemCheck, IntegrityCheck, SelfTest, Update the members' file.", vbInformation, TTL
    Exit Sub
Failed:
    em = Err.Description
    LogError "LoadMockData", Err.Number, em
    On Error Resume Next
    QuietOff evWas, scrWas, calcWas
    ProtectSheet RosterSheet()
    MarkNoSave
    MsgBox "LoadMockData stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving, reopen, and run it again.", vbCritical, TTL
End Sub

' ====================================================================
'  RESET (the wipe alone)
' ====================================================================
Public Sub ResetRosterData()
    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long, em As String, bk As String
    If Not MockGuard("ResetRosterData") Then Exit Sub
    If Not PeopleBackup("MOCK RESET", bk) Then Exit Sub
    QuietOn evWas, scrWas, calcWas
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ClearAll
    RefreshGradeSnapshot
    AfterPeopleChange
    RefreshCalCache
    Application.Calculate
    Dim leftTxt As String
    If RosterLastRow() >= R_FIRST Then leftTxt = leftTxt & "roster rows " & (RosterLastRow() - R_FIRST + 1) & "; "
    If LogRowCount() > 0 Then leftTxt = leftTxt & "OT Log rows " & LogRowCount() & "; "
    QuietOff evWas, scrWas, calcWas
    If leftTxt <> "" Then
        MarkNoSave
        MsgBox "ResetRosterData left data behind (" & leftTxt & ") - saving is blocked. Close WITHOUT saving.", vbCritical, TTL
        Exit Sub
    End If
    SaveIfWritable
    MsgBox "Every data table and the roster's people rows are empty; the Min Staff, Holidays, Rules and Settings sheets are untouched." & vbNewLine & _
           "Backup copy: " & bk, vbInformation, TTL
    Exit Sub
Failed:
    em = Err.Description
    LogError "ResetRosterData", Err.Number, em
    On Error Resume Next
    QuietOff evWas, scrWas, calcWas
    ProtectSheet RosterSheet()
    MarkNoSave
    MsgBox "ResetRosterData stopped part-way: " & em & vbNewLine & vbNewLine & "Saving is blocked. Close WITHOUT saving.", vbCritical, TTL
End Sub

' ====================================================================
'  GUARDS
' ====================================================================
' A practice copy (COPY / DRY RUN in the name) runs after one Yes; the live file only after the word MOCK is typed.
Private Function MockGuard(ByVal what As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    If Not PostingTablesReady() Then Exit Function
    If NoSave() Then
        MsgBox "Saving is blocked in this session (an earlier check failed). Close WITHOUT saving, reopen, then run " & what & ".", vbExclamation, TTL
        Exit Function
    End If
    If IsBackupCopyName(ThisWorkbook.Name) Then
        MsgBox ThisWorkbook.Name & " is a BACKUP copy (the name holds Backup or FINAL): " & what & " never runs on one. " & _
               "Open the live roster or a practice copy (COPY in the name).", vbExclamation, TTL
        Exit Function
    End If
    If Not IsDryRunName(ThisWorkbook.Name) Then
        Dim typed As String
        typed = Trim$(InputBox("THIS IS NOT A PRACTICE COPY: " & ThisWorkbook.Name & vbNewLine & vbNewLine & _
                               what & " empties every data table and every person on this roster and writes the mock plant." & vbNewLine & _
                               "The real roster must already be set aside (docs/alpha-data-reset.md, page 1)." & vbNewLine & vbNewLine & _
                               "Type the word " & MOCK_WORD & " to go on:", TTL))
        If StrComp(typed, MOCK_WORD, vbTextCompare) <> 0 Then
            MsgBox "Nothing was changed.", vbInformation, TTL
            Exit Function
        End If
    Else
        If MsgBox(what & " on " & ThisWorkbook.Name & " (a practice copy): every data table and every person on it is replaced by the mock plant." & vbNewLine & vbNewLine & _
                  "Go ahead?", vbYesNo + vbQuestion + vbDefaultButton2, TTL) <> vbYes Then Exit Function
    End If
    MockGuard = True
End Function

' ====================================================================
'  THE WIPE
' ====================================================================
Private Sub ClearAll()
    ClearTable LogTable()
    ClearTable AbsTable(True)
    ClearTable CovTable(True)
    ClearTable GetTable(PT_TABLE)
    ClearTable GetTable(PL_TABLE)
    ClearTable SwitchTable(True)
    ClearTable CrewTable(True)
    ClearTable GradeLogTable(True)
    ClearTable PlannedTable(True)
    ClearTable FormerTable(True)
    ClearTable RstTable(True)
    ClearTable RequestsTable(True)
    ClearTable FeedbackTable(True)
    ClearTable EvidenceTable(True)
    ClearTable ClosedPeriodsTable(True)
    ClearTable OpensTable(True)
    ClearTable DraftTable(True)
    ClearTable HoursHistoryTable(True)
    ClearTable WorkedPeriodsTable(True)
    ClearTable WorkedChangesTable(True)
    ClearTable KronosFilesTable(True)
    ClearTable ErrorLogTable(True)
    ClearKronosTables                               ' Kronos Totals + Worked Days, and the KronosCumThrough / KronosLedgerFrom names
    ClearAttendance
    ClearRosterRows
    ClearNames
End Sub

' The rollover's idiom: ClearContents, then one empty row as on a brand-new table
Private Sub ClearTable(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    On Error Resume Next
    lo.AutoFilter.ShowAllData
    On Error GoTo 0
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents
    lo.Resize lo.HeaderRowRange.Resize(2)
    ProtectSheet ws
End Sub

Private Sub ClearAttendance()
    If Not SheetExists(AT_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(AT_SHEET)
    UnlockSheet ws
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells.Clear
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ProtectSheet ws
End Sub

' Row R_FIRST is kept (emptied, not deleted) as the format template: AddRosterRow copies formats and validation from the
' row above (modPeople.AddRosterRow), so a deleted first row would leave every mock row without the roster's layout
Private Sub ClearRosterRows()
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    UnlockSheet ws
    If ws.FilterMode Then ws.ShowAllData
    If lastRow > R_FIRST Then ws.Rows((R_FIRST + 1) & ":" & lastRow).Delete Shift:=xlUp
    If lastRow >= R_FIRST Then
        With ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(R_FIRST, C_LAST))
            .ClearContents
            .ClearComments
        End With
    End If
    EnsureRosterFilter ws
    ProtectSheet ws
End Sub

' The hidden names that carry data: the ID marks, the Kronos stamps, the members' publish stamp, the Lite Desk's daily
' marks. RefusedThrough and KronosDataThrough are RE-SET, never deleted (the roster's formulas read them).
Private Sub ClearNames()
    Dim nm As Name, i As Long
    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If Left$(nm.Name, 7) = "IdMark_" Then nm.Delete
    Next i
    DeleteName "KronosImportedAt"
    DeleteName "KronosFileDate"
    DeleteName PUB_NAME
    DeleteName "DailyKronos"
    DeleteName "DailyNeeds"
    DeleteName "DailyPublish"
    ClearKronosAwait
    SetDateName PROV_NAME, DateSerial(1899, 12, 30)
End Sub

Private Sub DeleteName(ByVal s As String)
    On Error Resume Next
    ThisWorkbook.Names(s).Delete
    On Error GoTo 0
End Sub

' ====================================================================
'  THE WRITES
' ====================================================================
' Each person through AddRosterRow (the roster's own formulas in C, G, I and the period columns); Worked (H) = the
' Oct 18 snapshot's figure (the 17th field).
Private Sub WritePeople()
    Dim rows As Collection: Set rows = MockRows("people")
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim s As Variant, f() As String, codes() As String, r As Long, u As Long
    Dim units As Variant: units = Array("Phosphine", "Derivatives", "Cypure", "Material Handler", "QC Gases", "QC Derivatives", "Utilities", "SBP")
    For Each s In rows
        f = Split(CStr(s), SEP)
        ReDim codes(0 To UnitColLast() - UnitColFirst())
        For u = 0 To 7
            codes(DeptCol(CStr(units(u))) - UnitColFirst()) = f(8 + u)
        Next u
        r = AddRosterRow(ws, f(2), f(3), f(1), MockDate(f(6)), codes, f(4), f(7), f(0))
        ws.Cells(r, C_WORKED).Value = Val(f(16))
    Next s
    EnsureRosterFilter ws
    ProtectSheet ws
End Sub

' The OT Log rows, oldest first, in one AppendLogRows (it numbers the Entry IDs 1..n in that order and sorts newest first)
Private Sub WriteLog()
    Dim rows As Collection: Set rows = MockRows("log")
    Dim n As Long: n = rows.Count
    If n = 0 Then Exit Sub
    Dim data() As Variant: ReDim data(1 To n, 1 To L_NCOLS)
    Dim i As Long, c As Long, f() As String
    Dim types() As String: types = Split("T,T,D,T,T,T,T,N,N,T,T,N,T,T,N,D,N,T,S,T,T,T", ",")
    For i = 1 To n
        f = Split(rows(i), SEP)
        For c = 1 To L_NCOLS
            data(i, c) = MockValue(f(c - 1), types(c - 1))
        Next c
        data(i, L_ENTRY) = Empty                    ' AppendLogRows numbers them
    Next i
    If Not AppendLogRows(data) Then Err.Raise vbObjectError + 901, , "the OT Log rows could not be written"
End Sub

' One table from its rows: the roster's append idiom (UnlockSheet, Resize, Value, ProtectSheet); dates formatted
Private Sub WriteTable(ByVal lo As ListObject, ByVal key As String, ByVal typeList As String)
    If lo Is Nothing Then Err.Raise vbObjectError + 902, , "the table for " & key & " is missing"
    Dim rows As Collection: Set rows = MockRows(key)
    Dim n As Long: n = rows.Count
    If n = 0 Then Err.Raise vbObjectError + 904, , "no data rows for " & key & " (the generated block is incomplete)"
    Dim types() As String: types = Split(typeList, ",")
    Dim nCols As Long: nCols = UBound(types) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To nCols)
    Dim i As Long, c As Long, f() As String
    For i = 1 To n
        f = Split(rows(i), SEP)
        If UBound(f) + 1 <> nCols Then Err.Raise vbObjectError + 903, , key & " row " & i & " has " & (UBound(f) + 1) & " fields, not " & nCols
        For c = 1 To nCols
            data(i, c) = MockValue(f(c - 1), types(c - 1))
        Next c
    Next i
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, nCols)
        For c = 1 To nCols
            If types(c - 1) = "T" Or types(c - 1) = "R" Then .Columns(c).NumberFormat = "@"
        Next c
        .Value = data
        For c = 1 To nCols
            If types(c - 1) = "D" Then .Columns(c).NumberFormat = "ddd mmm d, yyyy"
            If types(c - 1) = "S" Then .Columns(c).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
        Next c
    End With
    ProtectSheet ws
End Sub

' The ID marks the log-based prefixes read (S, P, ADJ, NH, UM, LV, OB, MAN, ENTRY, RQ); the table-based prefixes
' (AB, SC, CV, SS, PM, RS, KR) count from their table and need none
Private Sub WriteMarks()
    Dim rows As Collection: Set rows = MockRows("idmarks")
    Dim s As Variant, f() As String
    For Each s In rows
        f = Split(CStr(s), SEP)
        SetIdMark f(0), CLng(Val(f(1)))
    Next s
    SetIdMark "ENTRY", LogRowCount()
End Sub

' T text, R a Reason (text; the real writers prefix an apostrophe into a General cell, the "@" format here does the same job),
' D date, S date and time, N number
Private Function MockValue(ByVal s As String, ByVal t As String) As Variant
    Select Case t
        Case "D", "S"
            If s = "" Then MockValue = Empty Else MockValue = MockDate(s)
        Case "N"
            If s = "" Then MockValue = Empty Else MockValue = Val(s)
        Case "R"
            MockValue = s                               ' the column is formatted "@" first: the text stays text with no prefix character
        Case Else
            MockValue = s
    End Select
End Function

Private Function MockDate(ByVal s As String) As Date
    If Len(s) < 10 Then Exit Function
    MockDate = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
    If Len(s) >= 19 Then MockDate = MockDate + TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
End Function

Private Function LogRowCount() As Long
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    End If
    LogRowCount = lo.ListRows.Count
End Function

' ====================================================================
'  THE VERIFY (before the save)
' ====================================================================
' RosterTotalsMap gives "G|I|sum of the period cells" per name; expected_load carries G, I and that sum, and the
' log's charged total. "" = everything matches and RunIntegrity found no PROBLEM.
Private Function VerifyLoad() As String
    Dim tm As Object: Set tm = RosterTotalsMap()
    Dim rows As Collection: Set rows = MockRows("expected")
    Dim s As Variant, f() As String, parts() As String, nPeople As Long
    For Each s In rows
        f = Split(CStr(s), SEP)
        If f(0) = "(log total Hours Charged)" Then
            If Abs(LogChargedTotal() - Val(f(1))) > 0.005 Then
                VerifyLoad = "the OT Log's total Hours Charged is " & Format$(LogChargedTotal(), "0.00") & ", expected " & f(1)
                Exit Function
            End If
        Else
            nPeople = nPeople + 1
            If Not tm.Exists(f(0)) Then VerifyLoad = f(0) & " is not on the roster": Exit Function
            parts = Split(CStr(tm(f(0))), "|")
            If UBound(parts) < 2 Then VerifyLoad = f(0) & ": the totals could not be read (" & tm(f(0)) & ")": Exit Function
            If Abs(Val(parts(0)) - Val(f(1))) > 0.005 Then VerifyLoad = f(0) & ": Total Hours G is " & parts(0) & ", expected " & f(1): Exit Function
            If Abs(Val(parts(1)) - Val(f(2))) > 0.005 Then VerifyLoad = f(0) & ": Refused I is " & parts(1) & ", expected " & f(2): Exit Function
            If Abs(Val(parts(2)) - Val(f(3))) > 0.005 Then VerifyLoad = f(0) & ": the period cells sum to " & parts(2) & ", expected " & f(3): Exit Function
        End If
    Next s
    If nPeople <> CHECK_PEOPLE Or tm.Count <> CHECK_PEOPLE Then
        VerifyLoad = "the roster holds " & tm.Count & " people, expected " & CHECK_PEOPLE
        Exit Function
    End If
    ' the home unit and the eight codes per person (RosterCodesMap: "Home|L|M|N|O|P|Q|R|S")
    Dim cm As Object: Set cm = RosterCodesMap()
    Dim units As Variant: units = Array("Phosphine", "Derivatives", "Cypure", "Material Handler", "QC Gases", "QC Derivatives", "Utilities", "SBP")
    Dim want() As String, u As Long, c As Long
    For Each s In MockRows("people")
        f = Split(CStr(s), SEP)
        ReDim want(0 To 8)
        want(0) = f(1)
        For u = 0 To 7
            c = DeptCol(CStr(units(u))) - UnitColFirst() + 1
            want(c) = f(8 + u)
        Next u
        If Not cm.Exists(f(2)) Then VerifyLoad = f(2) & " is not on the roster (codes)": Exit Function
        If StrComp(CStr(cm(f(2))), Join(want, "|"), vbTextCompare) <> 0 Then
            VerifyLoad = f(2) & ": home unit / codes are " & cm(f(2)) & ", expected " & Join(want, "|")
            Exit Function
        End If
    Next s
    Dim nErr As Long, nWarn As Long
    RunIntegrity nErr, nWarn
    If nErr > 0 Then VerifyLoad = "IntegrityCheck found " & nErr & " PROBLEM(S) - run IntegrityCheck to read them"
End Function

' ====================================================================
'  THE MEMBERS' FILES (member/modMember.bas SendSignup / WriteRequestFile, line for line)
' ====================================================================
' The two sign-up files P-0003 carries into Day 0: Harper Greenlaw (the whole shift) and Devon Stonebridge (6 h)
Private Function WriteCarriedSignups() As Long
    If MockSignupFile("P-0003", "Harper Greenlaw", 0, "signup", "campeauc") Then WriteCarriedSignups = WriteCarriedSignups + 1
    If MockSignupFile("P-0003", "Devon Stonebridge", 6, "signup", "sheltond") Then WriteCarriedSignups = WriteCarriedSignups + 1
End Function

' A sign-up / withdraw file in Signups\ (stamped now, so SignupSuspect never flags it)
Public Function MockSignupFile(ByVal postId As String, ByVal who As String, ByVal hrs As Double, ByVal action As String, _
                               ByVal login As String) As Boolean
    If Not EnsureSignupFolder() Then Exit Function
    Dim folder As String: folder = SignupFolderPath()
    Dim pc As String: pc = MockSafeName(Environ$("COMPUTERNAME"))
    Dim stamp As Date: stamp = Now
    Dim path As String, n As Long, ff As Integer
    path = folder & UCase$(postId) & "__" & MockSafeName(login) & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & ".txt"
    Do While Len(Dir$(path)) > 0 And n < 20
        n = n + 1
        path = folder & UCase$(postId) & "__" & MockSafeName(login) & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & "-" & n & ".txt"
    Loop
    On Error GoTo Failed
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "posting=" & UCase$(postId)
    Print #ff, "employee=" & who
    Print #ff, "hours=" & IIf(hrs > 0, Replace(CStr(hrs), ",", "."), "")
    Print #ff, "action=" & action
    Print #ff, "login=" & login
    Print #ff, "machine=" & Environ$("COMPUTERNAME")
    Print #ff, "stamp=" & Format$(stamp, "yyyy-mm-dd hh:nn:ss")
    Close #ff
    MockSignupFile = True
    Exit Function
Failed:
    On Error Resume Next
    Close #ff
    MsgBox "The sign-up file could not be written: " & Err.Description, vbExclamation, TTL
End Function

' A request file in Requests\: action request / agree / refuse / withdraw; kind "Time off" or "Switch"; rid = the
' request id (blank = a new one from the login and the stamp, as the members' file makes it)
Public Function MockRequestFile(ByVal rid As String, ByVal action As String, ByVal kind As String, ByVal who As String, _
                                ByVal dFrom As Date, ByVal dTo As Date, ByVal shiftLbl As String, ByVal partner As String, _
                                ByVal note As String, ByVal login As String) As String
    If Not EnsureRequestFolder() Then Exit Function
    Dim folder As String: folder = RequestFolderPath()
    Dim pc As String: pc = MockSafeName(Environ$("COMPUTERNAME"))
    Dim stamp As Date: stamp = Now
    If rid = "" Then rid = MockSafeName(login) & "-" & Format$(stamp, "yyyymmddhhnnss")
    Dim path As String, n As Long, ff As Integer
    path = folder & "REQ__" & rid & "__" & action & "__" & MockSafeName(login) & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & ".txt"
    Do While Len(Dir$(path)) > 0 And n < 20
        n = n + 1
        path = folder & "REQ__" & rid & "__" & action & "__" & MockSafeName(login) & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & "-" & n & ".txt"
    Loop
    On Error GoTo Failed
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "request=" & rid
    Print #ff, "action=" & action
    Print #ff, "kind=" & kind
    Print #ff, "employee=" & who
    Print #ff, "from=" & Format$(dFrom, "yyyy-mm-dd")
    Print #ff, "to=" & Format$(IIf(dTo > 0, dTo, dFrom), "yyyy-mm-dd")
    Print #ff, "shift=" & shiftLbl
    Print #ff, "partner=" & partner
    Print #ff, "note=" & Replace(Replace(note, vbCr, " "), vbLf, " ")
    Print #ff, "login=" & login
    Print #ff, "machine=" & Environ$("COMPUTERNAME")
    Print #ff, "stamp=" & Format$(stamp, "yyyy-mm-dd hh:nn:ss")
    Close #ff
    MockRequestFile = rid
    Exit Function
Failed:
    On Error Resume Next
    Close #ff
    MsgBox "The request file could not be written: " & Err.Description, vbExclamation, TTL
End Function

' Alt+F8: one request file from four boxes (the fallback when no member tester is at hand); the plan names the values
Public Sub MockRequestFromBox()
    On Error GoTo Oops
    Dim kind As String, who As String, dF As String, sh As String, partner As String, login As String, action As String, rid As String
    kind = Trim$(InputBox("Kind: Time off or Switch", TTL, "Time off"))
    If kind = "" Then Exit Sub
    who = Trim$(InputBox("Employee (as on the roster)", TTL))
    If who = "" Then Exit Sub
    dF = Trim$(InputBox("Date yyyy-mm-dd (a switch: the shift's date)", TTL))
    If Not (dF Like "####-##-##") Then MsgBox "The date must be typed as yyyy-mm-dd.", vbExclamation, TTL: Exit Sub
    If StrComp(kind, "Switch", vbTextCompare) = 0 Then
        sh = Trim$(InputBox("Shift: Days or Nights", TTL, "Days"))
        partner = Trim$(InputBox("Partner (the one who works it)", TTL))
    End If
    login = Trim$(InputBox("Login the file is written as (the member tester's Windows login)", TTL, Environ$("USERNAME")))
    action = Trim$(InputBox("Action: request, agree, refuse or withdraw", TTL, "request"))
    If action <> "request" Then rid = Trim$(InputBox("The request id the file answers (REQ__<id>__... in Requests\)", TTL))
    rid = MockRequestFile(rid, action, kind, who, MockDate(dF), MockDate(dF), sh, partner, "mock plan step", login)
    If rid <> "" Then MsgBox "Written to " & RequestFolderPath() & " as request " & rid & ". Alt+F8 > ImportRequests reads it.", vbInformation, TTL
    Exit Sub
Oops:
    ReportError "MockRequestFromBox"
End Sub

' member/modMember.bas SafeName: letters, digits, . _ - only; "user" when nothing is left
Private Function MockSafeName(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then out = out & ch
    Next i
    If out = "" Then out = "user"
    MockSafeName = out
End Function

' ====================================================================
'  THE DATA (generated)
' ====================================================================
Private Sub MockRow(ByVal s As String)
    mRows.Add s
End Sub

Private Function MockRows(ByVal key As String) As Collection
    Set mRows = New Collection
    LoadRows key
    Set MockRows = mRows
End Function

' >>> MOCK DATA (generated from tests/mock/*.csv by tests/mock/gen_mock_loader.py - do not edit by hand)
Private Sub LoadRows(ByVal key As String)
    Select Case key
        Case "people": Data_people_1
        Case "log": Data_log_1
        Case "crew": Data_crew_1
        Case "absences": Data_absences_1
        Case "coverage": Data_coverage_1
        Case "switches": Data_switches_1
        Case "postings": Data_postings_1
        Case "postinglist": Data_postinglist_1
        Case "planned": Data_planned_1
        Case "restrictions": Data_restrictions_1
        Case "hourshistory": Data_hourshistory_1: Data_hourshistory_2
        Case "idmarks": Data_idmarks_1
        Case "expected": Data_expected_1
    End Select
End Sub

Private Sub Data_people_1()
    MockRow "90100001|Phosphine|Avery Ashgrove|A|905-555-0100|23|2008-04-14|avery.ashgrove@example.test|P1||||||||16"
    MockRow "90100002|Phosphine|Blair Bellweather|A|905-555-0101|48|2015-09-21|blair.bellweather@example.test|P1||||||||38"
    MockRow "90100003|Phosphine|Casey Cardwell|A|905-555-0102|60|2019-02-04|casey.cardwell@example.test|P2||||||||47"
    MockRow "90100004|Phosphine|Dana Dunmore|A|905-555-0103|81|2023-06-12|dana.dunmore@example.test|P3||||||||65"
    MockRow "90100005|Derivatives|Ellis Everhart|A|905-555-0104|75|2006-11-06|ellis.everhart@example.test||D1|||||||56"
    MockRow "90100006|Derivatives|Finley Fairbanks|A|905-555-0105|42|2011-05-02|finley.fairbanks@example.test||D1|||||||20"
    MockRow "90100007|Derivatives|Harper Greenlaw|A|905-555-0106|58|2017-03-13|harper.greenlaw@example.test||D1.5|||||||33"
    MockRow "90100008|Derivatives|Jordan Hollins|A|905-555-0107|58|2020-10-05|jordan.hollins@example.test||D2|||||||49"
    MockRow "90100009|Derivatives|Kendall Ironside|A|905-555-0108|75|2024-01-15|kendall.ironside@example.test||D3|||||||63"
    MockRow "90100010|Cypure|Logan Juniper|A|905-555-0109|56|2012-08-20|logan.juniper@example.test|||G1||||||41"
    MockRow "90100011|Cypure|Morgan Kestrel|A|905-555-0110|40|2021-04-19|morgan.kestrel@example.test|||G2||||||22"
    MockRow "90100012|Material Handler|Quinn Lockridge|A|905-555-0111|52|2009-07-27|quinn.lockridge@example.test||||M1|||||31"
    MockRow "90100013|Material Handler|Reese Marsden|A|905-555-0112|75|2022-02-14|reese.marsden@example.test||||M2|||||51"
    MockRow "90100014|QC Derivatives|Riley Northway|A|905-555-0113|54|2013-10-07|riley.northway@example.test||||||Q1|||46"
    MockRow "90100015|Phosphine|Rowan Oakhurst|B|905-555-0114|34|2007-03-19|rowan.oakhurst@example.test|P1||||||||23"
    MockRow "90100016|Phosphine|Sawyer Pemberton|B|905-555-0115|58|2016-01-11|sawyer.pemberton@example.test|P1.5||||||||44"
    MockRow "90100017|Phosphine|Skyler Quarrie|B|905-555-0116|83|2018-06-25|skyler.quarrie@example.test|P2||||||||66"
    MockRow "90100018|Phosphine|Taylor Redfern|B|905-555-0117|77|2020-03-02|taylor.redfern@example.test|P2||||||||57"
    MockRow "90100019|Derivatives|Devon Stonebridge|B|905-555-0118|46|2011-05-02|devon.stonebridge@example.test||D1|||||||23"
    MockRow "90100020|Derivatives|Emery Thornbury|B|905-555-0119|58|2010-09-13|emery.thornbury@example.test||D1|||||||51"
    MockRow "90100021|Derivatives|Hayden Underhill|B|905-555-0120|52|2019-11-18|hayden.underhill@example.test||D2|||||||42"
    MockRow "90100022|Derivatives|Jesse Vantage|B|905-555-0121|83|2022-08-29|jesse.vantage@example.test||D2|||||||70"
    MockRow "90100023|Derivatives|Lane Wexford|B|905-555-0122|40|2024-05-06|lane.wexford@example.test||D3|||||||24"
    MockRow "90100024|Cypure|Marlow Yardley|B|905-555-0123|66|2014-02-17|marlow.yardley@example.test|||G1|M2|||||47"
    MockRow "90100025|Cypure|Oakley Ashby|B|905-555-0124|58|2023-09-11|oakley.ashby@example.test|||G3||||||36"
    MockRow "90100026|Material Handler|Parker Brightwater|B|905-555-0125|83|2005-06-20|parker.brightwater@example.test||||M1|||||58"
    MockRow "90100027|Material Handler|Remy Coldstream|B|905-555-0126|34|2021-01-25|remy.coldstream@example.test||||M3|||||25"
    MockRow "90100028|QC Derivatives|Shay Dovetail|B|905-555-0127|40|2012-04-30|shay.dovetail@example.test||||||Q2|||28"
    MockRow "90100029|Phosphine|Tatum Ellery|C|905-555-0128|64|2014-03-03|tatum.ellery@example.test|P1||||||||49"
    MockRow "90100030|Phosphine|Wren Fenwick|C|905-555-0129|77|2017-07-10|wren.fenwick@example.test|P1||||||||59"
    MockRow "90100031|Phosphine|Arden Glassford|C|905-555-0130|31|2019-09-16|arden.glassford@example.test|P2||||||||10"
    MockRow "90100032|Phosphine|Bryn Hartwell|C|905-555-0131|40|2024-03-25|bryn.hartwell@example.test|P3||||||||16"
    MockRow "90100033|Derivatives|Corin Islington|C|905-555-0132|52|2009-01-19|corin.islington@example.test||D1|||||||44"
    MockRow "90100034|Derivatives|Darcy Kingsley|C|905-555-0133|89|2015-05-18|darcy.kingsley@example.test||D1.5|||||||78"
    MockRow "90100035|Derivatives|Eden Larkspur|C|905-555-0134|75|2018-10-29|eden.larkspur@example.test||D2|||||||61"
    MockRow "90100036|Derivatives|Flynn Moorcroft|C|905-555-0135|34|2021-08-16|flynn.moorcroft@example.test||D2|||||||17"
    MockRow "90100037|Derivatives|Greer Nettleton|C|905-555-0136|58|2023-11-27|greer.nettleton@example.test||D3|||||||38"
    MockRow "90100038|Cypure|Hollis Orchard|C|905-555-0137|74|2013-06-03|hollis.orchard@example.test|||G1||||||51"
    MockRow "90100039|Cypure|Indy Pennyworth|C|905-555-0138|75|2020-07-20|indy.pennyworth@example.test|||G2||||||68"
    MockRow "90100040|Material Handler|Jules Ravenscroft|C|905-555-0139|40|2008-12-08|jules.ravenscroft@example.test||||M1|||||30"
    MockRow "90100041|Material Handler|Kit Saltmarsh|C|905-555-0140|40|2022-11-21|kit.saltmarsh@example.test||||M2|||||27"
    MockRow "90100042|QC Derivatives|Linden Tillingham|C|905-555-0141|60|2016-04-11|linden.tillingham@example.test||||||Q1|||44"
    MockRow "90100043|Phosphine|Merritt Vauxhall|D|905-555-0142|75|2016-08-15|merritt.vauxhall@example.test|P1||||||||56"
    MockRow "90100044|Phosphine|Nico Whitlock|D|905-555-0143|46|2018-02-26|nico.whitlock@example.test|P1.5||||||||24"
    MockRow "90100045|Phosphine|Odell Amberley|D|905-555-0144|42|2020-05-11|odell.amberley@example.test|P2||||||||17"
    MockRow "90100046|Phosphine|Perry Birchwood|D|905-555-0145|66|2023-03-06|perry.birchwood@example.test|P2||||||||57"
    MockRow "90100047|Derivatives|Rafe Crossley|D|905-555-0146|75|2007-10-15|rafe.crossley@example.test||D1|||||||63"
    MockRow "90100048|Derivatives|Sidney Dalrymple|D|905-555-0147|69|2012-12-10|sidney.dalrymple@example.test||D1|||||||54"
    MockRow "90100049|Derivatives|Tobin Eastwick|D|905-555-0148|46|2017-09-04|tobin.eastwick@example.test||D2|||||||28"
    MockRow "90100050|Derivatives|Vale Fairweather|D|905-555-0149|66|2021-06-14|vale.fairweather@example.test||D2.5|||||||45"
    MockRow "90100051|Derivatives|Wynn Goldfinch|D|905-555-0150|52|2024-08-19|wynn.goldfinch@example.test||D3|||||||28"
    MockRow "90100052|Cypure|Ash Hawthorne|D|905-555-0151|83|2015-01-12|ash.hawthorne@example.test|||G1||||||75"
    MockRow "90100053|Cypure|Blake Ingleby|D|905-555-0152|40|2022-05-23|blake.ingleby@example.test|||G2||||||29"
    MockRow "90100054|Material Handler|Cary Jarrow|D|905-555-0153|66|2010-02-22|cary.jarrow@example.test||||M1|||||52"
    MockRow "90100055|Material Handler|Dale Kelsall|D|905-555-0154|58|2019-04-08|dale.kelsall@example.test||||M2|||||41"
    MockRow "90100056|QC Derivatives|Evan Lindhurst|D|905-555-0155|75|2014-09-29|evan.lindhurst@example.test||||||Q3|||55"
    MockRow "90100057|QC Gases|Frankie Mallory|DAYS|905-555-0156|34|2011-11-14|frankie.mallory@example.test|||||QG1||||11"
    MockRow "90100058|QC Gases|Gale Norwood|DAYS|905-555-0157|48|2018-08-13|gale.norwood@example.test|||||QG2||||41"
    MockRow "90100059|QC Gases|Haven Overton|DAYS|905-555-0158|72|2023-01-09|haven.overton@example.test|||||QG3||||62"
    MockRow "90100060|Utilities|Ira Prentice|DAYS|905-555-0159|69|2009-05-25|ira.prentice@example.test|||||||U1||56"
    MockRow "90100061|Utilities|Jamie Rushmore|DAYS|905-555-0160|23|2016-10-17|jamie.rushmore@example.test|||||||U2||7"
    MockRow "90100062|Utilities|Kai Silverton|DAYS|905-555-0161|48|2022-09-26|kai.silverton@example.test|||||||U3||29"
    MockRow "90100063|SBP|Lee Tredwell|DAYS|905-555-0162|52|2010-06-28|lee.tredwell@example.test||||||||S1|30"
    MockRow "90100064|SBP|Max Upton|DAYS|905-555-0163|81|2017-12-04|max.upton@example.test||||||||S2|56"
    MockRow "90100065|SBP|Noel Wainwright|DAYS|905-555-0164|75|2021-03-15|noel.wainwright@example.test||||||||S3|66"
    MockRow "90100066|QC Derivatives|Pat Yeoman|DAYS|905-555-0165|50|2013-03-25|pat.yeoman@example.test||||||Q1|||38"
    MockRow "90100067|QC Derivatives|Robin Ashworth|DAYS|905-555-0166|58|2019-07-29|robin.ashworth@example.test||||||Q2|||43"
    MockRow "90100068|Phosphine|Sam Bramble|DAYS|905-555-0167|58|2012-02-06|sam.bramble@example.test|P1||||||||40"
    MockRow "90100069|Derivatives|Teal Copperfield|DAYS|905-555-0168|75|2015-11-23|teal.copperfield@example.test||D1||||Q2|||54"
End Sub

Private Sub Data_log_1()
    MockRow "OB-0001|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Avery Ashgrove|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|201|2025-12-28|1|piercyg|2026-01-04 09:00:00|||90100001"
    MockRow "OB-0002|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Blair Bellweather|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|113|2025-12-28|2|piercyg|2026-01-04 09:00:00|||90100002"
    MockRow "OB-0003|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Casey Cardwell|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|238.25|2025-12-28|3|piercyg|2026-01-04 09:00:00|||90100003"
    MockRow "OB-0004|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Dana Dunmore|P3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|269|2025-12-28|4|piercyg|2026-01-04 09:00:00|||90100004"
    MockRow "OB-0005|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Ellis Everhart|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|72|2025-12-28|5|piercyg|2026-01-04 09:00:00|||90100005"
    MockRow "OB-0006|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Finley Fairbanks|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|140|2025-12-28|6|piercyg|2026-01-04 09:00:00|||90100006"
    MockRow "OB-0007|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Harper Greenlaw|D1.5|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|185.25|2025-12-28|7|piercyg|2026-01-04 09:00:00|||90100007"
    MockRow "OB-0008|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Jordan Hollins|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|252.5|2025-12-28|8|piercyg|2026-01-04 09:00:00|||90100008"
    MockRow "OB-0009|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Kendall Ironside|D3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|327|2025-12-28|9|piercyg|2026-01-04 09:00:00|||90100009"
    MockRow "OB-0010|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Logan Juniper|G1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|120.5|2025-12-28|10|piercyg|2026-01-04 09:00:00|||90100010"
    MockRow "OB-0011|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Morgan Kestrel|G2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|248|2025-12-28|11|piercyg|2026-01-04 09:00:00|||90100011"
    MockRow "OB-0012|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Quinn Lockridge|M1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|81|2025-12-28|12|piercyg|2026-01-04 09:00:00|||90100012"
    MockRow "OB-0013|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Reese Marsden|M2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|135|2025-12-28|13|piercyg|2026-01-04 09:00:00|||90100013"
    MockRow "OB-0014|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Riley Northway|Q1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|213|2025-12-28|14|piercyg|2026-01-04 09:00:00|||90100014"
    MockRow "OB-0015|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Rowan Oakhurst|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|150.25|2025-12-28|15|piercyg|2026-01-04 09:00:00|||90100015"
    MockRow "OB-0016|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Sawyer Pemberton|P1.5|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|84|2025-12-28|16|piercyg|2026-01-04 09:00:00|||90100016"
    MockRow "OB-0017|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Skyler Quarrie|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|162|2025-12-28|17|piercyg|2026-01-04 09:00:00|||90100017"
    MockRow "OB-0018|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Taylor Redfern|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|242|2025-12-28|18|piercyg|2026-01-04 09:00:00|||90100018"
    MockRow "OB-0019|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Devon Stonebridge|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|136|2025-12-28|19|piercyg|2026-01-04 09:00:00|||90100019"
    MockRow "OB-0020|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Emery Thornbury|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|123.75|2025-12-28|20|piercyg|2026-01-04 09:00:00|||90100020"
    MockRow "OB-0021|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Hayden Underhill|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|206.5|2025-12-28|21|piercyg|2026-01-04 09:00:00|||90100021"
    MockRow "OB-0022|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Jesse Vantage|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|218|2025-12-28|22|piercyg|2026-01-04 09:00:00|||90100022"
    MockRow "OB-0023|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Lane Wexford|D3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|348|2025-12-28|23|piercyg|2026-01-04 09:00:00|||90100023"
    MockRow "OB-0024|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Marlow Yardley|G1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|109|2025-12-28|24|piercyg|2026-01-04 09:00:00|||90100024"
    MockRow "OB-0025|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Oakley Ashby|G3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|217.5|2025-12-28|25|piercyg|2026-01-04 09:00:00|||90100025"
    MockRow "OB-0026|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Parker Brightwater|M1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|45|2025-12-28|26|piercyg|2026-01-04 09:00:00|||90100026"
    MockRow "OB-0027|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Remy Coldstream|M3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|206|2025-12-28|27|piercyg|2026-01-04 09:00:00|||90100027"
    MockRow "OB-0028|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Shay Dovetail|Q2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|172|2025-12-28|28|piercyg|2026-01-04 09:00:00|||90100028"
    MockRow "OB-0029|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Tatum Ellery|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|86.5|2025-12-28|29|piercyg|2026-01-04 09:00:00|||90100029"
    MockRow "OB-0030|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Wren Fenwick|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|94.5|2025-12-28|30|piercyg|2026-01-04 09:00:00|||90100030"
    MockRow "OB-0031|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Arden Glassford|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|241|2025-12-28|31|piercyg|2026-01-04 09:00:00|||90100031"
    MockRow "OB-0032|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Bryn Hartwell|P3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|298.5|2025-12-28|32|piercyg|2026-01-04 09:00:00|||90100032"
    MockRow "OB-0033|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Corin Islington|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|104|2025-12-28|33|piercyg|2026-01-04 09:00:00|||90100033"
    MockRow "OB-0034|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Darcy Kingsley|D1.5|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|110.25|2025-12-28|34|piercyg|2026-01-04 09:00:00|||90100034"
    MockRow "OB-0035|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Eden Larkspur|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|176|2025-12-28|35|piercyg|2026-01-04 09:00:00|||90100035"
    MockRow "OB-0036|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Flynn Moorcroft|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|260.75|2025-12-28|36|piercyg|2026-01-04 09:00:00|||90100036"
    MockRow "OB-0037|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Greer Nettleton|D3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|307|2025-12-28|37|piercyg|2026-01-04 09:00:00|||90100037"
    MockRow "OB-0038|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Hollis Orchard|G1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|68|2025-12-28|38|piercyg|2026-01-04 09:00:00|||90100038"
    MockRow "OB-0039|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Indy Pennyworth|G2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|208|2025-12-28|39|piercyg|2026-01-04 09:00:00|||90100039"
    MockRow "OB-0040|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Jules Ravenscroft|M1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|118.25|2025-12-28|40|piercyg|2026-01-04 09:00:00|||90100040"
    MockRow "OB-0041|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Kit Saltmarsh|M2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|181|2025-12-28|41|piercyg|2026-01-04 09:00:00|||90100041"
    MockRow "OB-0042|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Linden Tillingham|Q1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|187.5|2025-12-28|42|piercyg|2026-01-04 09:00:00|||90100042"
    MockRow "OB-0043|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Merritt Vauxhall|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|107.5|2025-12-28|43|piercyg|2026-01-04 09:00:00|||90100043"
    MockRow "OB-0044|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Nico Whitlock|P1.5|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|147|2025-12-28|44|piercyg|2026-01-04 09:00:00|||90100044"
    MockRow "OB-0045|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Odell Amberley|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|220.25|2025-12-28|45|piercyg|2026-01-04 09:00:00|||90100045"
    MockRow "OB-0046|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Perry Birchwood|P2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|261|2025-12-28|46|piercyg|2026-01-04 09:00:00|||90100046"
    MockRow "OB-0047|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Rafe Crossley|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|73|2025-12-28|47|piercyg|2026-01-04 09:00:00|||90100047"
    MockRow "OB-0048|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Sidney Dalrymple|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|138.5|2025-12-28|48|piercyg|2026-01-04 09:00:00|||90100048"
    MockRow "OB-0049|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Tobin Eastwick|D2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|190|2025-12-28|49|piercyg|2026-01-04 09:00:00|||90100049"
    MockRow "OB-0050|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Vale Fairweather|D2.5|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|223|2025-12-28|50|piercyg|2026-01-04 09:00:00|||90100050"
    MockRow "OB-0051|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Wynn Goldfinch|D3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|327.5|2025-12-28|51|piercyg|2026-01-04 09:00:00|||90100051"
    MockRow "OB-0052|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Ash Hawthorne|G1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|99|2025-12-28|52|piercyg|2026-01-04 09:00:00|||90100052"
    MockRow "OB-0053|Opening Balance|2026-01-04|Cypure|||Opening balance|0||Blake Ingleby|G2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|256|2025-12-28|53|piercyg|2026-01-04 09:00:00|||90100053"
    MockRow "OB-0054|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Cary Jarrow|M1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|81|2025-12-28|54|piercyg|2026-01-04 09:00:00|||90100054"
    MockRow "OB-0055|Opening Balance|2026-01-04|Material Handler|||Opening balance|0||Dale Kelsall|M2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|171.75|2025-12-28|55|piercyg|2026-01-04 09:00:00|||90100055"
    MockRow "OB-0056|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Evan Lindhurst|Q3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|179|2025-12-28|56|piercyg|2026-01-04 09:00:00|||90100056"
    MockRow "OB-0057|Opening Balance|2026-01-04|QC Gases|||Opening balance|0||Frankie Mallory|QG1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|99|2025-12-28|57|piercyg|2026-01-04 09:00:00|||90100057"
    MockRow "OB-0058|Opening Balance|2026-01-04|QC Gases|||Opening balance|0||Gale Norwood|QG2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|140|2025-12-28|58|piercyg|2026-01-04 09:00:00|||90100058"
    MockRow "OB-0059|Opening Balance|2026-01-04|QC Gases|||Opening balance|0||Haven Overton|QG3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|169.5|2025-12-28|59|piercyg|2026-01-04 09:00:00|||90100059"
    MockRow "OB-0060|Opening Balance|2026-01-04|Utilities|||Opening balance|0||Ira Prentice|U1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|36|2025-12-28|60|piercyg|2026-01-04 09:00:00|||90100060"
    MockRow "OB-0061|Opening Balance|2026-01-04|Utilities|||Opening balance|0||Jamie Rushmore|U2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|151|2025-12-28|61|piercyg|2026-01-04 09:00:00|||90100061"
    MockRow "OB-0062|Opening Balance|2026-01-04|Utilities|||Opening balance|0||Kai Silverton|U3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|171.5|2025-12-28|62|piercyg|2026-01-04 09:00:00|||90100062"
    MockRow "OB-0063|Opening Balance|2026-01-04|SBP|||Opening balance|0||Lee Tredwell|S1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|87|2025-12-28|63|piercyg|2026-01-04 09:00:00|||90100063"
    MockRow "OB-0064|Opening Balance|2026-01-04|SBP|||Opening balance|0||Max Upton|S2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|114.5|2025-12-28|64|piercyg|2026-01-04 09:00:00|||90100064"
    MockRow "OB-0065|Opening Balance|2026-01-04|SBP|||Opening balance|0||Noel Wainwright|S3|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|153|2025-12-28|65|piercyg|2026-01-04 09:00:00|||90100065"
    MockRow "OB-0066|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Pat Yeoman|Q1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|101|2025-12-28|66|piercyg|2026-01-04 09:00:00|||90100066"
    MockRow "OB-0067|Opening Balance|2026-01-04|QC Derivatives|||Opening balance|0||Robin Ashworth|Q2|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|179|2025-12-28|67|piercyg|2026-01-04 09:00:00|||90100067"
    MockRow "OB-0068|Opening Balance|2026-01-04|Phosphine|||Opening balance|0||Sam Bramble|P1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|85.5|2025-12-28|68|piercyg|2026-01-04 09:00:00|||90100068"
    MockRow "OB-0069|Opening Balance|2026-01-04|Derivatives|||Opening balance|0||Teal Copperfield|D1|0|Adjustment|Opening balance: charges carried into 2026 before the roster went live|131|2025-12-28|69|piercyg|2026-01-04 09:00:00|||90100069"
    MockRow "S-0001|Callout|2026-10-06|Phosphine|Nights|All not scheduled||12|1|Rowan Oakhurst|P1|184.25|Refused||21|2026-10-04|70|sheltond|2026-10-06 17:35:00|||90100015"
    MockRow "S-0001|Callout|2026-10-06|Phosphine|Nights|All not scheduled||12|2|Sawyer Pemberton|P1.5|142|Unavailable|Out of town|0|2026-10-04|71|sheltond|2026-10-06 17:35:00|||90100016"
    MockRow "S-0001|Callout|2026-10-06|Phosphine|Nights|All not scheduled||12|3|Skyler Quarrie|P2|245|Refused||21|2026-10-04|72|sheltond|2026-10-06 17:35:00|||90100017"
    MockRow "S-0001|Callout|2026-10-06|Phosphine|Nights|All not scheduled||12|4|Taylor Redfern|P2|319|Accepted||0|2026-10-04|73|sheltond|2026-10-06 17:35:00|||90100018"
    MockRow "P-0001|Posting|2026-10-10|Derivatives|Days|All not scheduled|Vacation Coverage|12|1|Ellis Everhart|D1|147|Refused|Posting - didn't sign up (#1, awarded #3)|21|2026-10-04|74|campeauc|2026-09-28 09:10:00|||90100005"
    MockRow "P-0001|Posting|2026-10-10|Derivatives|Days|All not scheduled|Vacation Coverage|12|2|Devon Stonebridge|D1|182|Refused|Posting - didn't sign up (#2, awarded #3)|21|2026-10-04|75|campeauc|2026-09-28 09:10:00|||90100019"
    MockRow "P-0001|Posting|2026-10-10|Derivatives|Days|All not scheduled|Vacation Coverage|12|3|Finley Fairbanks|D1|182|Awarded||0|2026-10-04|76|campeauc|2026-09-28 09:10:00|||90100006"
    MockRow "P-0001|Posting|2026-10-10|Derivatives|Days|All not scheduled|Vacation Coverage|12|5|Harper Greenlaw|D1.5|243.25|Signed Up||0|2026-10-04|77|campeauc|2026-09-28 09:10:00|||90100007"
    MockRow "S-0002|Callout|2026-10-13|Cypure|Days|All not scheduled||12|1|Hollis Orchard|G1|142|Left Message||21|2026-10-04|78|mihaychukc|2026-10-13 05:50:00|||90100038"
    MockRow "S-0002|Callout|2026-10-13|Cypure|Days|All not scheduled||12|2|Ash Hawthorne|G1|182|Accepted||0|2026-10-04|79|mihaychukc|2026-10-13 05:50:00|||90100052"
    MockRow "S-0003|Callout|2026-10-20|Derivatives|Nights|All not scheduled||12|1|Ellis Everhart|D1|168|Refused||21|2026-10-18|80|sheltond|2026-10-20 16:20:00|||90100005"
    MockRow "S-0003|Callout|2026-10-20|Derivatives|Nights|All not scheduled||12|2|Finley Fairbanks|D1|182|Refused||21|2026-10-18|81|sheltond|2026-10-20 16:20:00|||90100006"
    MockRow "S-0003|Callout|2026-10-20|Derivatives|Nights|All not scheduled||12|3|Harper Greenlaw|D1.5|243.25|Exception|Plant upset - called in for the unit, not an OT refusal|0|2026-10-18|82|sheltond|2026-10-20 16:20:00|||90100007"
    MockRow "S-0003|Callout|2026-10-20|Derivatives|Nights|All not scheduled||12|4|Jordan Hollins|D2|310.5|Accepted||0|2026-10-18|83|sheltond|2026-10-20 16:20:00|||90100008"
    MockRow "P-0002|Posting|2026-10-24|Material Handler|Days|All not scheduled|Training Coverage|12|1|Parker Brightwater|M1|128|Refused|Posting - no takers|21|2026-10-18|84|piercyg|2026-10-16 08:05:00|||90100026"
    MockRow "P-0002|Posting|2026-10-24|Material Handler|Days|All not scheduled|Training Coverage|12|2|Quinn Lockridge|M1|133|Refused|Posting - no takers|21|2026-10-18|85|piercyg|2026-10-16 08:05:00|||90100012"
    MockRow "P-0002|Posting|2026-10-24|Material Handler|Days|All not scheduled|Training Coverage|12|3|Marlow Yardley|M2|175|Refused|Posting - no takers|21|2026-10-18|86|piercyg|2026-10-16 08:05:00|||90100024"
    MockRow "P-0002|Posting|2026-10-24|Material Handler|Days|All not scheduled|Training Coverage|12|4|Reese Marsden|M2|210|Refused|Posting - no takers|21|2026-10-18|87|piercyg|2026-10-16 08:05:00|||90100013"
    MockRow "S-0004|Callout|2026-10-24|Utilities|Days|All not scheduled||12|1|Ira Prentice|U1|105|Refused||21|2026-10-18|88|campeauc|2026-10-24 05:30:00|||90100060"
    MockRow "S-0004|Callout|2026-10-24|Utilities|Days|All not scheduled||12|2|Jamie Rushmore|U2|174|Accepted||0|2026-10-18|89|campeauc|2026-10-24 05:30:00|||90100061"
    MockRow "S-0005|Callout|2026-10-30|Phosphine|Nights|All not scheduled||12|1|Tatum Ellery|P1|150.5|Left Message||21|2026-10-18|90|mihaychukc|2026-10-30 17:10:00|||90100029"
    MockRow "S-0005|Callout|2026-10-30|Phosphine|Nights|All not scheduled||12|2|Wren Fenwick|P1|171.5|Accepted||0|2026-10-18|91|mihaychukc|2026-10-30 17:10:00|||90100030"
    MockRow "S-0006|Callout|2026-10-31|Derivatives|Days|All not scheduled||12|1|Rafe Crossley|D1|148|Refused||21|2026-10-18|92|sheltond|2026-10-31 05:40:00|||90100047"
    MockRow "S-0006|Callout|2026-10-31|Derivatives|Days|All not scheduled||12|2|Corin Islington|D1|156|Left Message||21|2026-10-18|93|sheltond|2026-10-31 05:40:00|||90100033"
    MockRow "S-0006|Callout|2026-10-31|Derivatives|Days|All not scheduled||12|3|Teal Copperfield|D1|206|Accepted||0|2026-10-18|94|sheltond|2026-10-31 05:40:00|||90100069"
End Sub

Private Sub Data_crew_1()
    MockRow "Avery Ashgrove|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Blair Bellweather|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Casey Cardwell|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Dana Dunmore|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Ellis Everhart|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Finley Fairbanks|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Harper Greenlaw|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Jordan Hollins|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Kendall Ironside|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Logan Juniper|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Morgan Kestrel|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Quinn Lockridge|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Reese Marsden|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Riley Northway|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Rowan Oakhurst|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Sawyer Pemberton|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Skyler Quarrie|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Taylor Redfern|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Devon Stonebridge|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Emery Thornbury|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Hayden Underhill|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Jesse Vantage|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Lane Wexford|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Marlow Yardley|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Oakley Ashby|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Parker Brightwater|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Remy Coldstream|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Shay Dovetail|B|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Tatum Ellery|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Wren Fenwick|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Arden Glassford|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Bryn Hartwell|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Corin Islington|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Darcy Kingsley|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Eden Larkspur|A|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Eden Larkspur|C|2026-10-20|piercyg|2026-10-19 15:10:00"
    MockRow "Flynn Moorcroft|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Greer Nettleton|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Hollis Orchard|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Indy Pennyworth|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Jules Ravenscroft|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Kit Saltmarsh|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Linden Tillingham|C|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Merritt Vauxhall|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Nico Whitlock|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Odell Amberley|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Perry Birchwood|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Rafe Crossley|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Sidney Dalrymple|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Tobin Eastwick|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Vale Fairweather|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Wynn Goldfinch|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Ash Hawthorne|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Blake Ingleby|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Cary Jarrow|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Dale Kelsall|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Evan Lindhurst|D|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Frankie Mallory|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Gale Norwood|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Haven Overton|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Ira Prentice|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Jamie Rushmore|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Kai Silverton|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Lee Tredwell|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Max Upton|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Noel Wainwright|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Pat Yeoman|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Robin Ashworth|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Sam Bramble|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
    MockRow "Teal Copperfield|DAYS|2025-12-28|piercyg|2025-12-28 08:00:00"
End Sub

Private Sub Data_absences_1()
    MockRow "SC-0001|Blair Bellweather|Phosphine|SIC|2026-08-21|2026-08-21||||Active|sheltond|2026-08-21 06:20:00||Phosphine||Days"
    MockRow "SC-0002|Blair Bellweather|Phosphine|SIC|2026-09-23|2026-09-23||||Active|campeauc|2026-09-23 17:45:00||Phosphine||Nights"
    MockRow "AB-0001|Vale Fairweather|Derivatives|VAC|2026-11-02|2026-11-05||||Active|piercyg|2026-09-30 09:12:00||Derivatives||"
    MockRow "AB-0002|Emery Thornbury|Derivatives|STD|2026-10-05|||||Active|piercyg|2026-10-05 08:30:00||Derivatives||"
    MockRow "AB-0003|Skyler Quarrie|Phosphine|PER|2026-11-10|2026-11-10||||Active|piercyg|2026-10-28 14:05:00||Phosphine||"
    MockRow "AB-0004|Bryn Hartwell|Phosphine|VAC|2026-09-14|2026-09-17||||Active|piercyg|2026-08-20 10:00:00||Phosphine||"
End Sub

Private Sub Data_coverage_1()
    MockRow "CV-0001|2026-11-02|Days|Derivatives|AB-0001|Vale Fairweather (Vac)|VAC|||AB-0001|Active|piercyg|2026-11-02 08:00:00|"
    MockRow "CV-0002|2026-11-03|Days|Derivatives||Vale Fairweather (Vac)|VAC|POSTED P-0003||P-0003|Active|piercyg|2026-11-03 08:00:00|"
    MockRow "CV-0003|2026-11-10|Days|Phosphine|AB-0003|Skyler Quarrie (Personal)|PER|||AB-0003|Active|piercyg|2026-11-10 08:00:00|"
    MockRow "CV-0004|2026-09-12|Days|Cypure||Hollis Orchard S/S|SS|Logan Juniper S/S||SS-0001|Active|piercyg|2026-09-12 08:00:00|"
    MockRow "CV-0005|2026-10-06|Nights|Phosphine||Call-out S-0001|SUP|Taylor Redfern||S-0001|Active|piercyg|2026-10-06 08:00:00|"
    MockRow "CV-0006|2026-10-10|Days|Derivatives||Vacation Coverage|VAC|Finley Fairbanks||P-0001|Active|piercyg|2026-10-10 08:00:00|"
    MockRow "CV-0007|2026-10-13|Days|Cypure||Call-out S-0002|SUP|Ash Hawthorne||S-0002|Active|piercyg|2026-10-13 08:00:00|"
    MockRow "CV-0008|2026-10-20|Nights|Derivatives||Call-out S-0003|SUP|Jordan Hollins||S-0003|Active|piercyg|2026-10-20 08:00:00|"
    MockRow "CV-0009|2026-10-24|Days|Utilities||Call-out S-0004|SUP|Jamie Rushmore||S-0004|Active|piercyg|2026-10-24 08:00:00|"
    MockRow "CV-0010|2026-10-30|Nights|Phosphine||Call-out S-0005|SUP|Wren Fenwick||S-0005|Active|piercyg|2026-10-30 08:00:00|"
    MockRow "CV-0011|2026-10-31|Days|Derivatives||Call-out S-0006|SUP|Teal Copperfield||S-0006|Active|piercyg|2026-10-31 08:00:00|"
    MockRow "CV-0012|2026-10-24|Days|Material Handler||Training Coverage|TRN|NO TAKERS||P-0002|Active|piercyg|2026-10-24 08:00:00|"
    MockRow "CV-0013|2026-11-05|Days|Material Handler||Training Coverage|TRN|POSTED P-0004||P-0004|Active|piercyg|2026-11-05 08:00:00|"
    MockRow "CV-0014|2026-11-20|Days|Phosphine||Unit Operator Support|SUP|POSTED P-0005||P-0005|Active|piercyg|2026-11-20 08:00:00|"
End Sub

Private Sub Data_switches_1()
    MockRow "SS-0001|2026-09-12|Days|Cypure|Hollis Orchard|C|Logan Juniper|A|||payback to be arranged|Active|mihaychukc|2026-09-10 13:40:00|"
End Sub

Private Sub Data_postings_1()
    MockRow "P-0001|Awarded|2026-09-24|2026-09-26|2026-10-10|Derivatives|Days|All not scheduled|12|1|Vacation Coverage|Finley Fairbanks|2026-09-28|campeauc|campeauc"
    MockRow "P-0002|No takers|2026-10-08|2026-10-10|2026-10-24|Material Handler|Days|All not scheduled|12|1|Training Coverage||2026-10-16|piercyg|piercyg"
    MockRow "P-0003|Open|2026-10-20|2026-11-02|2026-11-03|Derivatives|Days|All not scheduled|12|1|Vacation Coverage||||campeauc"
    MockRow "P-0004|Open|2026-10-24|2026-11-01|2026-11-05|Material Handler|Days|All not scheduled|12|1|Training Coverage||||sheltond"
    MockRow "P-0005|Open|2026-10-28|2026-11-06|2026-11-20|Phosphine|Days|All not scheduled|12|1|Unit Operator Support||||piercyg"
End Sub

Private Sub Data_postinglist_1()
    MockRow "P-0001|1|Ellis Everhart|A|D1|147||Charged|21|"
    MockRow "P-0001|2|Devon Stonebridge|B|D1|182||Charged|21|"
    MockRow "P-0001|3|Finley Fairbanks|A|D1|182|YES|Awarded||"
    MockRow "P-0001|4|Teal Copperfield|DAYS|D1|206||||"
    MockRow "P-0001|5|Harper Greenlaw|A|D1.5|243.25|YES|Signed up||"
    MockRow "P-0001|6|Eden Larkspur|A|D2|251||||"
    MockRow "P-0001|7|Hayden Underhill|B|D2|258.5||||"
    MockRow "P-0001|8|Jesse Vantage|B|D2|301||||"
    MockRow "P-0001|9|Jordan Hollins|A|D2|310.5||||"
    MockRow "P-0001|10|Lane Wexford|B|D3|388||||"
    MockRow "P-0001|11|Kendall Ironside|A|D3|402||||"
    MockRow "P-0002|1|Parker Brightwater|B|M1|128||Charged|21|"
    MockRow "P-0002|2|Quinn Lockridge|A|M1|133||Charged|21|"
    MockRow "P-0002|3|Marlow Yardley|B|M2|175||Charged|21|"
    MockRow "P-0002|4|Reese Marsden|A|M2|210||Charged|21|"
    MockRow "P-0002|5|Remy Coldstream|B|M3|240||||"
    MockRow "P-0003|1|Ellis Everhart|A|D1|189||||"
    MockRow "P-0003|2|Devon Stonebridge|B|D1|203||||"
    MockRow "P-0003|3|Finley Fairbanks|A|D1|203||||"
    MockRow "P-0003|4|Harper Greenlaw|A|D1.5|243.25||||"
    MockRow "P-0003|5|Hayden Underhill|B|D2|258.5||||"
    MockRow "P-0003|6|Jesse Vantage|B|D2|301||||"
    MockRow "P-0003|7|Jordan Hollins|A|D2|310.5||||"
    MockRow "P-0003|8|Lane Wexford|B|D3|388||||"
    MockRow "P-0003|9|Kendall Ironside|A|D3|402||||"
    MockRow "P-0004|1|Cary Jarrow|D|M1|147||||"
    MockRow "P-0004|2|Jules Ravenscroft|C|M1|158.25||||"
    MockRow "P-0004|3|Kit Saltmarsh|C|M2|221||||"
    MockRow "P-0004|4|Dale Kelsall|D|M2|229.75||||"
    MockRow "P-0005|1|Rowan Oakhurst|B|P1|205.25||||"
    MockRow "P-0005|2|Sawyer Pemberton|B|P1.5|142||||"
    MockRow "P-0005|3|Skyler Quarrie|B|P2|266||||"
    MockRow "P-0005|4|Taylor Redfern|B|P2|319||||"
End Sub

Private Sub Data_planned_1()
    MockRow "PM-0001|Blake Ingleby|Material Handler|M2|Cypure||D|2026-11-08|Planned|piercyg|2026-10-27 11:30:00||"
End Sub

Private Sub Data_restrictions_1()
    MockRow "RS-0001|Ellis Everhart|2026-11-09|2026-11-13|medical - no overtime this week|Active|piercyg|2026-10-29 16:00:00|"
End Sub

Private Sub Data_hourshistory_1()
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Avery Ashgrove|11|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Avery Ashgrove|16|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Blair Bellweather|28|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Blair Bellweather|38|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Casey Cardwell|32|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Casey Cardwell|47|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Dana Dunmore|45|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Dana Dunmore|65|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Ellis Everhart|31|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Ellis Everhart|56|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Finley Fairbanks|13|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Finley Fairbanks|20|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Harper Greenlaw|21|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Harper Greenlaw|33|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Jordan Hollins|32|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Jordan Hollins|49|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Kendall Ironside|41|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Kendall Ironside|63|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Logan Juniper|14|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Logan Juniper|41|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Morgan Kestrel|13|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Morgan Kestrel|22|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Quinn Lockridge|17|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Quinn Lockridge|31|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Reese Marsden|32|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Reese Marsden|51|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Riley Northway|22|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Riley Northway|46|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Rowan Oakhurst|17|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Rowan Oakhurst|23|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Sawyer Pemberton|33|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Sawyer Pemberton|44|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Skyler Quarrie|50|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Skyler Quarrie|66|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Taylor Redfern|36|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Taylor Redfern|57|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Devon Stonebridge|0|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Devon Stonebridge|23|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Emery Thornbury|43|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Emery Thornbury|51|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Hayden Underhill|29|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Hayden Underhill|42|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Jesse Vantage|52|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Jesse Vantage|70|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Lane Wexford|1|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Lane Wexford|24|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Marlow Yardley|42|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Marlow Yardley|47|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Oakley Ashby|26|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Oakley Ashby|36|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Parker Brightwater|43|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Parker Brightwater|58|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Remy Coldstream|5|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Remy Coldstream|25|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Shay Dovetail|3|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Shay Dovetail|28|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Tatum Ellery|42|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Tatum Ellery|49|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Wren Fenwick|47|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Wren Fenwick|59|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Arden Glassford|0|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Arden Glassford|10|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Bryn Hartwell|0|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Bryn Hartwell|16|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Corin Islington|17|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Corin Islington|44|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Darcy Kingsley|69|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Darcy Kingsley|78|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Eden Larkspur|47|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Eden Larkspur|61|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Flynn Moorcroft|0|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Flynn Moorcroft|17|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Greer Nettleton|14|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Greer Nettleton|38|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Hollis Orchard|45|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Hollis Orchard|51|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Indy Pennyworth|57|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Indy Pennyworth|68|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Jules Ravenscroft|14|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Jules Ravenscroft|30|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Kit Saltmarsh|6|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Kit Saltmarsh|27|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Linden Tillingham|18|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Linden Tillingham|44|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Merritt Vauxhall|48|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Merritt Vauxhall|56|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Nico Whitlock|11|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Nico Whitlock|24|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Odell Amberley|0|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Odell Amberley|17|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Perry Birchwood|34|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Perry Birchwood|57|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Rafe Crossley|58|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Rafe Crossley|63|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Sidney Dalrymple|44|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Sidney Dalrymple|54|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Tobin Eastwick|13|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Tobin Eastwick|28|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Vale Fairweather|25|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Vale Fairweather|45|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Wynn Goldfinch|3|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Wynn Goldfinch|28|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Ash Hawthorne|68|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Ash Hawthorne|75|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Blake Ingleby|17|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Blake Ingleby|29|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Cary Jarrow|35|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Cary Jarrow|52|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Dale Kelsall|19|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Dale Kelsall|41|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Evan Lindhurst|28|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Evan Lindhurst|55|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Frankie Mallory|2|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Frankie Mallory|11|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Gale Norwood|27|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Gale Norwood|41|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Haven Overton|43|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Haven Overton|62|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Ira Prentice|32|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Ira Prentice|56|piercyg|2026-10-17|2026"
End Sub

Private Sub Data_hourshistory_2()
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Jamie Rushmore|1|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Jamie Rushmore|7|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Kai Silverton|18|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Kai Silverton|29|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Lee Tredwell|14|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Lee Tredwell|30|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Max Upton|35|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Max Upton|56|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Noel Wainwright|40|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Noel Wainwright|66|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Pat Yeoman|30|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Pat Yeoman|38|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Robin Ashworth|30|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Robin Ashworth|43|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Sam Bramble|22|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Sam Bramble|40|piercyg|2026-10-17|2026"
    MockRow "KR-0021|2026-10-04 07:40:00|2026-10-04|Teal Copperfield|31|piercyg|2026-10-03|2026"
    MockRow "KR-0022|2026-10-18 07:35:00|2026-10-18|Teal Copperfield|54|piercyg|2026-10-17|2026"
End Sub

Private Sub Data_idmarks_1()
    MockRow "OB|69"
    MockRow "S|6"
    MockRow "P|5"
    MockRow "ADJ|0"
    MockRow "RQ|0"
    MockRow "LV|0"
    MockRow "UM|0"
    MockRow "NH|0"
    MockRow "MAN|0"
End Sub

Private Sub Data_expected_1()
    MockRow "Avery Ashgrove|217.00|201.00|201.00|16|0"
    MockRow "Blair Bellweather|151.00|113.00|113.00|38|0"
    MockRow "Casey Cardwell|285.25|238.25|238.25|47|0"
    MockRow "Dana Dunmore|334.00|269.00|269.00|65|0"
    MockRow "Ellis Everhart|170.00|114.00|114.00|56|0"
    MockRow "Finley Fairbanks|181.00|161.00|161.00|20|0"
    MockRow "Harper Greenlaw|218.25|185.25|185.25|33|0"
    MockRow "Jordan Hollins|313.50|252.50|252.50|49|12"
    MockRow "Kendall Ironside|390.00|327.00|327.00|63|0"
    MockRow "Logan Juniper|161.50|120.50|120.50|41|0"
    MockRow "Morgan Kestrel|270.00|248.00|248.00|22|0"
    MockRow "Quinn Lockridge|133.00|102.00|102.00|31|0"
    MockRow "Reese Marsden|207.00|156.00|156.00|51|0"
    MockRow "Riley Northway|259.00|213.00|213.00|46|0"
    MockRow "Rowan Oakhurst|194.25|171.25|171.25|23|0"
    MockRow "Sawyer Pemberton|128.00|84.00|84.00|44|0"
    MockRow "Skyler Quarrie|249.00|183.00|183.00|66|0"
    MockRow "Taylor Redfern|299.00|242.00|242.00|57|0"
    MockRow "Devon Stonebridge|180.00|157.00|157.00|23|0"
    MockRow "Emery Thornbury|174.75|123.75|123.75|51|0"
    MockRow "Hayden Underhill|248.50|206.50|206.50|42|0"
    MockRow "Jesse Vantage|288.00|218.00|218.00|70|0"
    MockRow "Lane Wexford|372.00|348.00|348.00|24|0"
    MockRow "Marlow Yardley|177.00|130.00|130.00|47|0"
    MockRow "Oakley Ashby|253.50|217.50|217.50|36|0"
    MockRow "Parker Brightwater|124.00|66.00|66.00|58|0"
    MockRow "Remy Coldstream|231.00|206.00|206.00|25|0"
    MockRow "Shay Dovetail|200.00|172.00|172.00|28|0"
    MockRow "Tatum Ellery|156.50|107.50|107.50|49|0"
    MockRow "Wren Fenwick|165.50|94.50|94.50|59|12"
    MockRow "Arden Glassford|251.00|241.00|241.00|10|0"
    MockRow "Bryn Hartwell|314.50|298.50|298.50|16|0"
    MockRow "Corin Islington|169.00|125.00|125.00|44|0"
    MockRow "Darcy Kingsley|188.25|110.25|110.25|78|0"
    MockRow "Eden Larkspur|237.00|176.00|176.00|61|0"
    MockRow "Flynn Moorcroft|277.75|260.75|260.75|17|0"
    MockRow "Greer Nettleton|345.00|307.00|307.00|38|0"
    MockRow "Hollis Orchard|140.00|89.00|89.00|51|0"
    MockRow "Indy Pennyworth|276.00|208.00|208.00|68|0"
    MockRow "Jules Ravenscroft|148.25|118.25|118.25|30|0"
    MockRow "Kit Saltmarsh|208.00|181.00|181.00|27|0"
    MockRow "Linden Tillingham|231.50|187.50|187.50|44|0"
    MockRow "Merritt Vauxhall|163.50|107.50|107.50|56|0"
    MockRow "Nico Whitlock|171.00|147.00|147.00|24|0"
    MockRow "Odell Amberley|237.25|220.25|220.25|17|0"
    MockRow "Perry Birchwood|318.00|261.00|261.00|57|0"
    MockRow "Rafe Crossley|157.00|94.00|94.00|63|0"
    MockRow "Sidney Dalrymple|192.50|138.50|138.50|54|0"
    MockRow "Tobin Eastwick|218.00|190.00|190.00|28|0"
    MockRow "Vale Fairweather|268.00|223.00|223.00|45|0"
    MockRow "Wynn Goldfinch|355.50|327.50|327.50|28|0"
    MockRow "Ash Hawthorne|174.00|99.00|99.00|75|0"
    MockRow "Blake Ingleby|285.00|256.00|256.00|29|0"
    MockRow "Cary Jarrow|133.00|81.00|81.00|52|0"
    MockRow "Dale Kelsall|212.75|171.75|171.75|41|0"
    MockRow "Evan Lindhurst|234.00|179.00|179.00|55|0"
    MockRow "Frankie Mallory|110.00|99.00|99.00|11|0"
    MockRow "Gale Norwood|181.00|140.00|140.00|41|0"
    MockRow "Haven Overton|231.50|169.50|169.50|62|0"
    MockRow "Ira Prentice|113.00|57.00|57.00|56|0"
    MockRow "Jamie Rushmore|170.00|151.00|151.00|7|12"
    MockRow "Kai Silverton|200.50|171.50|171.50|29|0"
    MockRow "Lee Tredwell|117.00|87.00|87.00|30|0"
    MockRow "Max Upton|170.50|114.50|114.50|56|0"
    MockRow "Noel Wainwright|219.00|153.00|153.00|66|0"
    MockRow "Pat Yeoman|139.00|101.00|101.00|38|0"
    MockRow "Robin Ashworth|222.00|179.00|179.00|43|0"
    MockRow "Sam Bramble|125.50|85.50|85.50|40|0"
    MockRow "Teal Copperfield|197.00|131.00|131.00|54|12"
    MockRow "(log total Hours Charged)|11864.25||||"
End Sub
' <<< MOCK DATA
