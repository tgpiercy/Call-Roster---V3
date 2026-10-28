Attribute VB_Name = "modFormer"
Option Explicit
' v27.1 (the refactor, Oct 2026): the Former sheet, Remove and Bring back, from modPeople; moved unchanged (tests/movecheck.py proves it)
Public Const FORMER_VERSION As String = "Former v29.0.1 (Oct 28)"

' tblFormer (v24): one row per person taken off the roster, as values
Public Const FM_SHEET As String = "Former"
Public Const FM_TABLE As String = "tblFormer"
Public Const FM_EMP As Long = 1
Public Const FM_LEFT As Long = 2
Public Const FM_CREW As Long = 3
Public Const FM_HOME As Long = 4
Public Const FM_SENIORITY As Long = 5
Public Const FM_CODE1 As Long = 6           ' the eight unit codes, in roster column order (L:S)
Public Const FM_TOTAL As Long = 14
Public Const FM_WORKED As Long = 15
Public Const FM_REFUSED As Long = 16
Public Const FM_PROV As Long = 17
Public Const FM_YEAR As Long = 18           ' the roster year they were removed in
Public Const FM_USER As Long = 19
Public Const FM_STAMP As Long = 20
Public Const FM_STATUS As Long = 21         ' Former / Returned
Public Const FM_BACK As Long = 22           ' Returned On
Public Const FM_PAYROLL As Long = 23        ' v24.2: Payroll # (added on the end, so the columns before it keep their numbers)
Public Const FM_N As Long = 23
' ====================================================================
'  FORMER EMPLOYEES (tblFormer)
' ====================================================================
Public Function FormerTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(FM_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(FM_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(FM_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = FM_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        Dim hdr() As Variant: ReDim hdr(1 To FM_N)
        Dim c As Long
        hdr(FM_EMP) = "Employee": hdr(FM_LEFT) = "Left On": hdr(FM_CREW) = "Crew": hdr(FM_HOME) = "Home Unit"
        hdr(FM_SENIORITY) = "Plant Seniority"
        For c = UnitColFirst() To UnitColLast()
            hdr(FM_CODE1 + c - UnitColFirst()) = DeptOfCol(c)
        Next c
        hdr(FM_TOTAL) = "Total Hours": hdr(FM_WORKED) = "Worked": hdr(FM_REFUSED) = "Refused": hdr(FM_PROV) = "Provisional"
        hdr(FM_YEAR) = "Roster Year": hdr(FM_USER) = "Removed By": hdr(FM_STAMP) = "Timestamp"
        hdr(FM_STATUS) = "Status": hdr(FM_BACK) = "Returned On": hdr(FM_PAYROLL) = "Payroll #"
        ws.Range("A1").Resize(1, FM_N).Value = hdr
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, FM_N), , xlYes)
        lo.Name = FM_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(FM_EMP).ColumnWidth = 26
        ws.Range(ws.Columns(FM_LEFT), ws.Columns(FM_N)).ColumnWidth = 12
        ws.Columns(FM_HOME).ColumnWidth = 16
        ws.Columns(FM_STAMP).ColumnWidth = 20
        ws.Rows(1).WrapText = True
        FormatFormer ws
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    If Not lo Is Nothing And create Then EnsureFormerColumns lo      ' never from a read (IntegrityCheck, SystemCheck)
    Set FormerTable = lo
End Function

' A v24 Former table has 22 columns: Payroll # (v24.2) goes on the end, once
Private Sub EnsureFormerColumns(ByVal lo As ListObject)
    If lo.ListColumns.Count >= FM_N Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    lo.ListColumns.Add.Name = "Payroll #"
    ws.Columns(FM_PAYROLL).ColumnWidth = 12
    ws.Columns(FM_PAYROLL).NumberFormat = "@"
    ProtectSheet ws
    On Error GoTo 0
End Sub

Private Sub FormatFormer(ByVal ws As Worksheet)
    ws.Columns(FM_LEFT).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_BACK).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(FM_SENIORITY).NumberFormat = "mmm d, yyyy"
    ws.Columns(FM_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ws.Columns(FM_CREW).NumberFormat = "@"
    ws.Columns(FM_PAYROLL).NumberFormat = "@"
End Sub

' Row in tblFormer of a person who is still "Former" (0 = none)
Private Function FormerRowOf(ByVal emp As String, ByRef v As Variant) As Long
    v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = UBound(v, 1) To 1 Step -1
        If StrComp(Trim$(SafeStr(v(i, FM_EMP))), Trim$(emp), vbTextCompare) = 0 And _
           StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then FormerRowOf = i: Exit Function
    Next i
End Function

' True for someone moved to the Former sheet and not brought back
Public Function IsFormer(ByVal emp As String) As Boolean
    Dim v As Variant
    IsFormer = (FormerRowOf(emp, v) > 0)
End Function

' The day a person still on the Former sheet left (0 = not there)
Public Function FormerLeftOn(ByVal emp As String) As Date
    Dim v As Variant, fr As Long
    fr = FormerRowOf(emp, v)
    If fr = 0 Then Exit Function
    If IsDate(v(fr, FM_LEFT)) Then FormerLeftOn = Int(CDate(v(fr, FM_LEFT)))
End Function

' Everyone on the Former sheet (Former or Returned): name -> True.
' IntegrityCheck and the Kronos import use it: these names are known.
Public Function FormerNameSet() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set FormerNameSet = d
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim i As Long, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, FM_EMP)))
        If nm <> "" Then d(nm) = True
    Next i
End Function

' Names still "Former", A-Z, for the Bring back list (Empty if none)
Public Function FormerNames() As Variant
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    Dim a() As String: ReDim a(0 To UBound(v, 1) - 1)
    Dim i As Long, j As Long, n As Long, t As String
    For i = 1 To UBound(v, 1)
        t = Trim$(SafeStr(v(i, FM_EMP)))
        If t <> "" And StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then
            a(n) = t & "   (left " & IIf(IsDate(v(i, FM_LEFT)), Format$(v(i, FM_LEFT), "mmm d, yyyy"), "?") & ")"
            n = n + 1
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve a(0 To n - 1)
    For i = 1 To n - 1
        t = a(i): j = i - 1
        Do While j >= 0
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    FormerNames = a
End Function

' Rows of the Former sheet removed in roster year y, for the rollover's Year
' End record: out(1..n, 1..6) = name, home unit, total, worked, refused, provisional
Public Function FormerFinals(ByVal y As Long, ByRef out() As Variant) As Long
    Dim v As Variant: v = BodyOf(FormerTable(False))
    If Not IsArray(v) Then Exit Function
    ReDim out(1 To UBound(v, 1), 1 To 6)
    Dim i As Long, n As Long
    For i = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(i, FM_EMP))) <> "" And CLng(NumOr0(v(i, FM_YEAR))) = y And _
           StrComp(SafeStr(v(i, FM_STATUS)), "Former", vbTextCompare) = 0 Then
            n = n + 1
            out(n, 1) = Trim$(SafeStr(v(i, FM_EMP)))
            out(n, 2) = SafeStr(v(i, FM_HOME))
            out(n, 3) = NumOr0(v(i, FM_TOTAL))
            out(n, 4) = NumOr0(v(i, FM_WORKED))
            out(n, 5) = NumOr0(v(i, FM_REFUSED))
            out(n, 6) = NumOr0(v(i, FM_PROV))
        End If
    Next i
    FormerFinals = n
End Function

Private Function AppendFormerRow(ByRef rec() As Variant) As Boolean
    Dim lo As ListObject: Set lo = FormerTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < FM_N Then Exit Function      ' its Payroll # column couldn't be added (read-only?)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    FormatFormer ws
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, FM_N).Value = rec
    ProtectSheet ws
    AppendFormerRow = True
End Function


' ====================================================================
'  REMOVE / BRING BACK  (the Former sheet)
' ====================================================================
' Moves the person's roster row to the Former sheet as values. Their OT Log,
' absence, switch and posting rows stay. A backup is saved first; every other
' row's totals and the OT Log total are checked after.
Public Function RemoveEmployee(ByVal emp As String, ByVal leftOn As Date, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim nRows As Long, r As Long
    r = RosterRowOf(emp, nRows)
    If r = 0 Then MsgBox emp & " isn't on the roster.", vbExclamation, "Remove": Exit Function
    If nRows > 1 Then
        MsgBox emp & " is on the roster " & nRows & " times. Delete the extra row by hand first (IntegrityCheck shows which)." & ByHandHint(RosterSheet().Name), vbExclamation, "Remove"
        Exit Function
    End If
    If Int(leftOn) > Date Then
        MsgBox "Remove takes the person off every list for every date, so it is done on or after their last day." & vbNewLine & _
               "Until then, book them off (Book Time Off).", vbExclamation, "Remove"
        Exit Function
    End If
    emp = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
    Dim bk As String
    If Not PeopleBackup("remove", bk) Then Exit Function

    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, after As Object, logBefore As Double, em As String, diff As String, postTxt As String
    Dim rec() As Variant: ReDim rec(1 To 1, 1 To FM_N)
    Dim c As Long, logHrs As Double
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    logHrs = LogTotalOnRoster(emp)
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    ' v24.3: a unit move still waiting is cancelled first, while the row is here
    ' (it may have dated a crew change, and the Former row keeps the crew)
    Dim planId As String, planTxt As String
    planId = PlannedFor(emp)
    If planId <> "" Then
        If Not CancelPlannedCore(planId, "removed from the roster, left " & Format$(leftOn, "mmm d, yyyy") & _
                                 " (" & CurrentUser() & ", " & Format$(Date, "mmm d") & ")", planTxt, False) Then
            Err.Raise vbObjectError + 732, , planTxt
        End If
        planTxt = vbNewLine & "Their planned move " & planId & " was cancelled."
    End If
    rec(1, FM_EMP) = emp
    rec(1, FM_LEFT) = Int(leftOn)
    rec(1, FM_CREW) = Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value))
    rec(1, FM_HOME) = Trim$(SafeStr(ws.Cells(r, C_HOME).Value))
    rec(1, FM_SENIORITY) = ws.Cells(r, C_SENIORITY).Value
    For c = UnitColFirst() To UnitColLast()
        rec(1, FM_CODE1 + c - UnitColFirst()) = Trim$(SafeStr(ws.Cells(r, c).Value))
    Next c
    rec(1, FM_TOTAL) = NumOr0(ws.Cells(r, C_TOTAL).Value2)
    rec(1, FM_WORKED) = NumOr0(ws.Cells(r, C_WORKED).Value2)
    rec(1, FM_REFUSED) = NumOr0(ws.Cells(r, C_REFUSED).Value2)
    rec(1, FM_PROV) = NumOr0(ws.Cells(r, C_PROV).Value2)
    rec(1, FM_YEAR) = RosterYear()
    rec(1, FM_USER) = CurrentUser()
    rec(1, FM_STAMP) = Now
    rec(1, FM_STATUS) = "Former"
    rec(1, FM_PAYROLL) = PayrollText(ws.Cells(r, C_PAYROLL).Value)
    If Not AppendFormerRow(rec) Then Err.Raise vbObjectError + 731, , "the Former sheet couldn't be written"

    UnlockSheet ws
    If ws.FilterMode Then ws.ShowAllData
    ws.Rows(r).Delete Shift:=xlUp
    EnsureRosterFilter ws
    ProtectSheet ws
    Application.Calculate
    Set after = RosterTotalsMap()
    diff = TotalsDifference(before, after, emp)
    If after.Exists(emp) Then diff = diff & "  " & emp & " is still on the roster" & vbNewLine
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    If Not IsFormer(emp) Then diff = diff & "  " & emp & " isn't on the Former sheet" & vbNewLine
    If diff <> "" Then
        QuietOff evWas, scrWas, calcWas
        MarkNoSave
        MsgBox "Removing " & emp & " did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Remove"
        Exit Function
    End If
    postTxt = ExemptFromOpenPostings(emp, Int(leftOn), "", "left " & Format$(leftOn, "mmm d"))
    QuietOff evWas, scrWas, calcWas
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = emp & " is on the Former sheet (left " & Format$(leftOn, "mmm d, yyyy") & ") and off the roster." & vbNewLine & _
             "Final figures kept: Total " & HoursText(rec(1, FM_TOTAL)) & ", Worked " & HoursText(rec(1, FM_WORKED)) & _
             ", Refused " & HoursText(rec(1, FM_REFUSED)) & ". Their " & HoursText(logHrs) & " h in the OT Log stay there." & _
             IIf(postTxt <> "", vbNewLine & "Not chargeable on open posting(s): " & postTxt, "") & planTxt & vbNewLine & _
             "Everyone else's totals are unchanged. Publish the members' file again." & vbNewLine & "Backup: " & bk
    RemoveEmployee = True
    Exit Function
Failed:
    em = Err.Description
    LogError "RemoveEmployee", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "Removing " & emp & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Remove"
End Function

' Puts a former person back on the roster with the same formulas as everyone
' else, so their OT Log hours for this year show again.
Public Function BringBack(ByVal emp As String, ByRef report As String) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If LogTable() Is Nothing Then Exit Function
    Dim fv As Variant, fr As Long
    fr = FormerRowOf(emp, fv)
    If fr = 0 Then MsgBox emp & " isn't on the Former sheet.", vbExclamation, "Bring back": Exit Function
    emp = Trim$(SafeStr(fv(fr, FM_EMP)))
    If RosterRowOf(emp) > 0 Then MsgBox emp & " is already on the roster.", vbExclamation, "Bring back": Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim bk As String
    If Not PeopleBackup("bring back", bk) Then Exit Function

    Dim evWas As Boolean, scrWas As Boolean, calcWas As Long
    QuietOn evWas, scrWas, calcWas
    Dim before As Object, logBefore As Double, em As String, diff As String, r As Long, c As Long
    Dim codes() As String: ReDim codes(0 To UnitColLast() - UnitColFirst())
    Dim lo As ListObject, found As Boolean, crewTxt As String, newest As String, postTxt As String
    For c = UnitColFirst() To UnitColLast()
        codes(c - UnitColFirst()) = Trim$(SafeStr(fv(fr, FM_CODE1 + c - UnitColFirst())))
    Next c
    crewTxt = NormCrew(SafeStr(fv(fr, FM_CREW)))
    Application.Calculate
    Set before = RosterTotalsMap()
    logBefore = LogChargedTotal()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mBusy = True
    Dim payNo As String
    If UBound(fv, 2) >= FM_PAYROLL Then payNo = PayrollText(fv(fr, FM_PAYROLL))
    r = AddRosterRow(ws, emp, crewTxt, SafeStr(fv(fr, FM_HOME)), fv(fr, FM_SENIORITY), codes, "", "", payNo)
    EnsureRosterFilter ws
    ProtectSheet ws
    Set lo = FormerTable(False)
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(fr, FM_STATUS).Value = "Returned"
    lo.DataBodyRange.Cells(fr, FM_BACK).Value = Date
    ProtectSheet lo.Parent
    newest = NewestCrew(emp, found)
    If Not found Or newest <> crewTxt Then
        If Not AddCrewRow(emp, crewTxt, Date) Then Err.Raise vbObjectError + 741, , "the crew history row couldn't be written"
    End If
    postTxt = ClearLeftExemptions(emp)                   ' chargeable again on open postings
    Application.Calculate
    diff = TotalsDifference(before, RosterTotalsMap(), emp)
    If Abs(LogChargedTotal() - logBefore) > 0.0001 Then diff = diff & "  OT Log total Hours Charged changed" & vbNewLine
    If Abs(NumOr0(ws.Cells(r, C_REFUSED).Value2) - LogTotalOnRoster(emp, True)) > 0.0001 Then   ' v26.5: up to the current period, as I counts
        diff = diff & "  " & emp & ": Refused shows " & NumOr0(ws.Cells(r, C_REFUSED).Value2) & ", the OT Log has " & LogTotalOnRoster(emp, True) & " up to the current pay period" & vbNewLine
    End If
    QuietOff evWas, scrWas, calcWas
    If diff <> "" Then
        MarkNoSave
        MsgBox "Bringing " & emp & " back did NOT verify:" & vbNewLine & diff & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Bring back"
        Exit Function
    End If
    RefreshGradeSnapshot
    AfterPeopleChange
    SaveIfWritable
    report = emp & " is back on the roster (row " & r & ") with " & HoursText(NumOr0(ws.Cells(r, C_REFUSED).Value2)) & _
             " h charged this year from the OT Log." & vbNewLine & _
             "Type their phone and email in again (they aren't kept on the Former sheet). Worked hours fill in at the next Kronos import." & _
             IIf(postTxt <> "", vbNewLine & "Chargeable again on open posting(s): " & postTxt, "") & _
             vbNewLine & "Backup: " & bk
    BringBack = True
    Exit Function
Failed:
    em = Err.Description
    LogError "BringBack", Err.Number, em
    MarkNoSave                                  ' first: nothing below can stop the save block going on
    QuietOff evWas, scrWas, calcWas
    On Error Resume Next
    ProtectAll
    MsgBox "Bringing " & emp & " back stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving. Backup: " & bk, vbCritical, "Bring back"
End Function


