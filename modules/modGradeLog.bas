Attribute VB_Name = "modGradeLog"
Option Explicit

' ====================================================================
'  GRADE CHANGE LOG (v21)
'  When a unit-code cell on the roster (L:S) changes, e.g. D2 -> D1, a
'  row goes into tblGradeChange on the hidden admin sheet "Grade Log":
'     Timestamp | Employee | Unit | Old Code | New Code | Windows Login
'  The OLD value comes from a snapshot (very hidden sheet "GradeSnap",
'  one row per employee: name + the eight codes). The snapshot is keyed
'  on the NAME, not the row, so sorting the roster moves rows but changes
'  no name->code pair and writes nothing. Filters fire no change event.
'  Macros that write codes set gSuppressGradeLog (or turn events off).
'  On open, the roster is reconciled against the snapshot so a change
'  made with macros off is still recorded (login marked "on open").
' ====================================================================

Public Const GL_VERSION As String = "Grade log v27.1 (Oct 22)"
Public Const GL_SHEET As String = "Grade Log"
Public Const GL_TABLE As String = "tblGradeChange"
Public Const GS_SHEET As String = "GradeSnap"
Private Const GL_N As Long = 6
Private Const GL_NF As Long = 7             ' v27 (item 13, review A11): + Field - a person's detail changed in place (seniority, home unit, crew, phone, email)

Public gSuppressGradeLog As Boolean     ' set by macros that write unit codes on purpose
Private mBusy As Boolean                ' re-entry guard (our own writes fire change events)


' ---------------- event entry (from ThisWorkbook.Workbook_SheetChange) ----------------
Public Sub GradeSheetChanged(ByVal Sh As Object, ByVal Target As Range)
    If mBusy Or gSuppressGradeLog Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    If Target Is Nothing Then Exit Sub
    On Error GoTo Quit
    Dim ws As Worksheet: Set ws = RosterSheet()
    If Not Sh Is ws Then Exit Sub
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    ' a Home Unit filled in by hand: Needs Attention's list is out of date
    If Not Application.Intersect(Target, ws.Range(ws.Cells(R_FIRST, C_HOME), ws.Cells(lastRow, C_HOME))) Is Nothing Then MarkCalDirty
    Dim hit As Range
    Set hit = Application.Intersect(Target, ws.Range(ws.Cells(R_FIRST, UnitColFirst()), ws.Cells(lastRow, UnitColLast())))
    If hit Is Nothing Then Exit Sub

    mBusy = True
    ' rows touched (a sort hands us the whole block - every row is checked by name)
    Dim r1 As Long, r2 As Long, area As Range
    r1 = lastRow: r2 = R_FIRST
    For Each area In hit.Areas
        If area.Row < r1 Then r1 = area.Row
        If area.Row + area.Rows.Count - 1 > r2 Then r2 = area.Row + area.Rows.Count - 1
    Next area
    RecordChanges ws, r1, r2, CurrentUser()
Quit:
    mBusy = False
End Sub


' ---------------- the comparison ----------------
' Compares roster rows r1..r2 with the snapshot (by name) and logs every
' code that differs. Names not in the snapshot are added silently.
' Returns the number of rows logged.
Private Function RecordChanges(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal who As String) As Long
    Dim snap As Object: Set snap = SnapshotDict()
    If snap Is Nothing Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, UnitColLast())).Value
    Dim i As Long, c As Long, nm As String, oldCodes As Variant, newCode As String, oldCode As String
    Dim rows() As Variant: ReDim rows(1 To (r2 - r1 + 1) * 8, 1 To GL_N)
    Dim n As Long, stamp As Date: stamp = Now
    Dim needRefresh As Boolean
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" Then
            If snap.Exists(nm) Then
                oldCodes = Split(snap(nm), "|")
                For c = UnitColFirst() To UnitColLast()
                    newCode = Trim$(SafeStr(v(i, c)))
                    oldCode = CStr(oldCodes(c - UnitColFirst()))
                    If StrComp(newCode, oldCode, vbBinaryCompare) <> 0 Then
                        n = n + 1
                        rows(n, 1) = stamp
                        rows(n, 2) = nm
                        rows(n, 3) = DeptOfCol(c)
                        rows(n, 4) = oldCode
                        rows(n, 5) = newCode
                        rows(n, 6) = who
                    End If
                Next c
            Else
                needRefresh = True                 ' new (or renamed) person: snapshot them, no entry
            End If
        End If
    Next i
    Dim written As Boolean
    If n > 0 Then
        Dim outA() As Variant: ReDim outA(1 To n, 1 To GL_N)
        For i = 1 To n
            For c = 1 To GL_N
                outA(i, c) = rows(i, c)
            Next c
        Next i
        written = AppendGradeRows(outA, n)
    End If
    ' old values for next time - but if the rows couldn't be written, keep the
    ' old snapshot so the change is still caught (on open, or next edit)
    If written Or (n = 0 And needRefresh) Then RefreshGradeSnapshot
    RecordChanges = n
End Function


' The Employee pop-up (v24) writes unit codes with events off, then asks
' for the same comparison a typed change gets.
Public Sub RecordGradeRows(ByVal r1 As Long, ByVal r2 As Long, ByVal who As String)
    If mBusy Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    On Error GoTo Quit
    mBusy = True
    RecordChanges RosterSheet(), r1, r2, who
Quit:
    mBusy = False
End Sub

' A rename (v24) leaves a line here, so a name in a closed year's archive
' can be traced to the person's name today: Unit "(renamed)", Old / New = the names.
Public Sub LogRename(ByVal oldName As String, ByVal newName As String)
    On Error Resume Next                            ' never raises
    Dim outA() As Variant: ReDim outA(1 To 1, 1 To GL_N)
    outA(1, 1) = Now
    outA(1, 2) = newName
    outA(1, 3) = "(renamed)"
    outA(1, 4) = oldName
    outA(1, 5) = newName
    outA(1, 6) = CurrentUser()
    AppendGradeRows outA, 1
End Sub


' ---------------- snapshot ----------------
' name -> "L|M|N|O|P|Q|R|S" from the GradeSnap sheet (Nothing if there is none yet)
Private Function SnapshotDict() As Object
    If Not SheetExists(GS_SHEET) Then Exit Function
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(GS_SHEET)
    Dim last As Long: last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    If last >= 2 Then
        Dim v As Variant: v = ws.Range(ws.Cells(2, 1), ws.Cells(last, 9)).Value
        Dim i As Long, c As Long, k As String, nm As String
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, 1)))
            If nm <> "" Then
                k = ""
                For c = 2 To 9
                    k = k & IIf(c = 2, "", "|") & Trim$(SafeStr(v(i, c)))
                Next c
                d(nm) = k
            End If
        Next i
    End If
    Set SnapshotDict = d
End Function

' Rewrites the snapshot from the roster (every name + its eight codes).
' Called after each recorded change, by SetupHomeUnit, and on open when
' there is no snapshot yet. Quiet: never stops the caller.
Public Sub RefreshGradeSnapshot()
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim evWas As Boolean: evWas = Application.EnableEvents    ' captured first: Quit restores it
    Dim ws As Worksheet
    On Error GoTo Quit
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Set ws = SnapSheet(True)
    If ws Is Nothing Then Exit Sub
    Application.EnableEvents = False
    UnlockSheet ws
    ws.Cells.ClearContents
    ws.Range("A1:I1").Value = Array("Employee", "QC Derivatives", "Derivatives", "Cypure", "Phosphine", _
                                   "Material Handler", "QC Gases", "Utilities", "SBP")
    ws.Range("K1").Value = "Snapshot of roster unit codes for the grade change log - do not edit"
    If lastRow >= R_FIRST Then
        Dim v As Variant: v = rs.Range(rs.Cells(R_FIRST, 1), rs.Cells(lastRow, UnitColLast())).Value
        Dim outA() As Variant: ReDim outA(1 To UBound(v, 1), 1 To 9)
        Dim i As Long, c As Long, n As Long, nm As String
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, C_NAME)))
            If nm <> "" Then
                n = n + 1
                outA(n, 1) = nm
                For c = UnitColFirst() To UnitColLast()
                    outA(n, c - UnitColFirst() + 2) = Trim$(SafeStr(v(i, c)))
                Next c
            End If
        Next i
        If n > 0 Then ws.Range("A2").Resize(n, 9).Value = outA
    End If
    ProtectSheet ws
    SetSheetVisible ws, xlSheetVeryHidden
    Application.EnableEvents = evWas
    Exit Sub
Quit:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
End Sub

Private Function SnapSheet(ByVal create As Boolean) As Worksheet
    If SheetExists(GS_SHEET) Then
        Set SnapSheet = ThisWorkbook.Worksheets(GS_SHEET)
    ElseIf create Then
        Dim ws As Worksheet, was As Object
        Set was = ActiveSheet                      ' adding a sheet activates it - put the user back
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = GS_SHEET
        RelockStructure
        SetSheetVisible ws, xlSheetVeryHidden
        On Error Resume Next
        If Not was Is Nothing Then was.Activate
        On Error GoTo 0
        Set SnapSheet = ws
    End If
End Function

' Workbook_Open (admins, writable): build the snapshot the first time, or
' log anything that changed while macros were off. Quiet.
Public Sub GradeLogOnOpen()
    If ThisWorkbook.ReadOnly Or Not IsAdmin() Then Exit Sub
    On Error Resume Next                            ' never raises
    If Not SheetExists(GS_SHEET) Then
        RefreshGradeSnapshot
    Else
        mBusy = True
        Dim rs As Worksheet: Set rs = RosterSheet()
        Dim lastRow As Long: lastRow = RosterLastRow()
        If lastRow >= R_FIRST Then RecordChanges rs, R_FIRST, lastRow, CurrentUser() & " (on open)"
        mBusy = False
    End If
End Sub


' ---------------- the log table ----------------
Public Function GradeLogTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(GL_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(GL_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(GL_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet       ' adding a sheet activates it - put the user back
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = GL_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, GL_NF).Value = Array("Timestamp", "Employee", "Unit", "Old Code", "New Code", "Windows Login", "Field")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, GL_NF), , xlYes)
        lo.Name = GL_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns("A").ColumnWidth = 19
        ws.Columns("B").ColumnWidth = 24
        ws.Columns("C").ColumnWidth = 16
        ws.Columns("D:E").ColumnWidth = 9
        ws.Columns("F").ColumnWidth = 14
        ws.Columns("A").NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    If Not lo Is Nothing Then EnsureGradeLogField lo
    Set GradeLogTable = lo
End Function

' v27: a table made before v27 has 6 columns; the Field column goes on the end once (as Hours History's Data Through did)
Private Sub EnsureGradeLogField(ByVal lo As ListObject)
    If lo.ListColumns.Count >= GL_NF Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    lo.ListColumns.Add.Name = "Field"
    ws.Columns(GL_NF).ColumnWidth = 12
    ProtectSheet ws
    On Error GoTo 0
End Sub

' v27 (item 13): a detail changed in place through Employee > Edit - one row: Unit "(details)", Old / New = the values,
' Field = which detail (seniority, home unit, crew, phone, email). A code change keeps its own rows (RecordGradeRows).
Public Sub LogFieldChange(ByVal emp As String, ByVal fld As String, ByVal oldV As String, ByVal newV As String)
    On Error Resume Next                            ' never raises
    Dim outA() As Variant: ReDim outA(1 To 1, 1 To GL_NF)
    outA(1, 1) = Now
    outA(1, 2) = emp
    outA(1, 3) = "(details)"
    outA(1, 4) = oldV
    outA(1, 5) = newV
    outA(1, 6) = CurrentUser()
    outA(1, 7) = fld
    AppendGradeRows outA, 1
End Sub

' True when the rows are in the table
Private Function AppendGradeRows(ByRef outA() As Variant, ByVal n As Long) As Boolean
    Dim lo As ListObject: Set lo = GradeLogTable(True)
    If lo Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim em As String
    On Error GoTo Fail
    Application.EnableEvents = False
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, UBound(outA, 2))      ' v27: 6 or 7 columns, as the caller built them
        .Value = outA
        .Columns(1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    AppendGradeRows = True
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendGradeRows", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    Application.EnableEvents = evWas
    MsgBox "The grade change couldn't be written to the Grade Log: " & em & vbNewLine & _
           "The roster change itself is kept. Send " & PRIMARY_ADMIN_NAME & " this message.", vbExclamation, "Grade Log"
End Function

' Alt+F8 > ShowGradeLog (admin): opens the Grade Log sheet
Public Sub ShowGradeLog()
    On Error GoTo Oops                         ' v24.3.1
    If Not RequireAdmin() Then Exit Sub
    If GradeLogTable(True) Is Nothing Then
        MsgBox "The Grade Log couldn't be created (is the file read-only?).", vbExclamation, "Call Out Roster"
        Exit Sub
    End If
    SetSheetVisible ThisWorkbook.Worksheets(GL_SHEET), xlSheetVisible
    ThisWorkbook.Worksheets(GL_SHEET).Activate
    Exit Sub
Oops:
    ReportError "ShowGradeLog"
End Sub
