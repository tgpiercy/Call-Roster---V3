Attribute VB_Name = "modKeys"
Option Explicit

' ====================================================================
'  v28 (item 2): THE PAYROLL # AS THE KEY - carried on every people table, the rename and the Kronos match by it.
'  Graeme, Oct 25, 2026: every people table gets a "Payroll #" column (added on the END, so no column number moves);
'  writers leave it blank and FillPayrollKeys stamps it from the roster's column A (then the Former sheet) before
'  every save, after a Kronos import and after Add / Edit; the roster's pay-period SUMIFS stay by name; the members'
'  file never sees the number (PrivateDataIn refuses a 7+ digit run). AssignPayrollKeys (admin, run once at the
'  v28 install) adds the columns and fills every row, verified. A new hire's number is typed at Add (frmEmployee).
'  Readers still match by name in v28; the key makes the rename safe (RenameEverywhere) and names the mismatches
'  (IntegrityCheck 16 "Keys").
' ====================================================================
Public Const KEY_VERSION As String = "Keys v29.0.3 (Oct 28)"
Public Const KEY_HEADER As String = "Payroll #"
Public Const KEY_HEADER_OFF As String = "Payroll # (gives up)"   ' the Shift Switches table carries two people
Public Const KEY_HEADER_WORKS As String = "Payroll # (works)"
Private Const KEY_MAX_DIGITS As Long = 10

' ---- PURE (suite payroll): a typed Payroll #: "" when fine (blank allowed), else why not ----
Public Function PayrollValid(ByVal txt As String) As String
    txt = Trim$(txt)
    If txt = "" Then Exit Function
    If Len(txt) > KEY_MAX_DIGITS Then PayrollValid = "at most " & KEY_MAX_DIGITS & " digits": Exit Function
    Dim i As Long
    For i = 1 To Len(txt)
        If Mid$(txt, i, 1) < "0" Or Mid$(txt, i, 1) > "9" Then PayrollValid = "digits only": Exit Function
    Next i
End Function

' ---- PURE (suite keymatch): two rows are the same person when both carry a key and the keys agree; with a key
'      missing on either side, when the names agree (trimmed, any case) ----
Public Function KeyMatchCore(ByVal keyA As String, ByVal keyB As String, ByVal nameA As String, ByVal nameB As String) As Boolean
    keyA = Trim$(keyA): keyB = Trim$(keyB)
    If keyA <> "" And keyB <> "" Then
        KeyMatchCore = (keyA = keyB)
    Else
        KeyMatchCore = (StrComp(Trim$(nameA), Trim$(nameB), vbTextCompare) = 0) And Trim$(nameA) <> ""
    End If
End Function

' ---- the roster's numbers: name -> Payroll # (the roster's column A, then the Former sheet) ----
Public Function PayrollMap() As Object
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary"): map.CompareMode = 1
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim r As Long, nm As String, k As String
    If lastRow >= R_FIRST Then
        Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_PAYROLL), ws.Cells(lastRow, C_NAME)).Value
        For r = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(r, C_NAME - C_PAYROLL + 1)))
            k = PayrollText(v(r, 1))
            If nm <> "" And k <> "" Then
                If Not map.Exists(nm) Then map.Add nm, k
            End If
        Next r
    End If
    Dim lo As ListObject: Set lo = FormerTable(False)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing And lo.ListColumns.Count >= FM_PAYROLL Then
            Dim fv As Variant: fv = RangeTo2D(lo.DataBodyRange)
            For r = 1 To UBound(fv, 1)
                nm = Trim$(SafeStr(fv(r, FM_EMP)))
                k = PayrollText(fv(r, FM_PAYROLL))
                If nm <> "" And k <> "" Then
                    If Not map.Exists(nm) Then map.Add nm, k
                End If
            Next r
        End If
    End If
    Set PayrollMap = map
End Function

' True when another roster row, or a Former person under another name, already carries this number (review MED-3: a
' Former number taken by a new hire would give the Former person's old rows to the new hire at the next rename)
Public Function PayrollTaken(ByVal no As String, ByVal exceptName As String) As Boolean
    no = Trim$(no)
    If no = "" Then Exit Function
    Dim map As Object: Set map = PayrollMap()
    Dim kk As Variant
    For Each kk In map.Keys
        If CStr(map(kk)) = no Then
            If StrComp(Trim$(CStr(kk)), Trim$(exceptName), vbTextCompare) <> 0 Then PayrollTaken = True: Exit Function
        End If
    Next kk
End Function

' v28 (review MED-1): a person's number corrected (or cleared) in the Employee pop-up - the rows that carry the OLD number
' under this name take the new one (blank = cleared, and FillPayrollKeys leaves it so until a number is known)
Public Sub RekeyRows(ByVal nm As String, ByVal oldKey As String, ByVal newKey As String)
    On Error GoTo Oops
    oldKey = Trim$(oldKey): newKey = Trim$(newKey): nm = Trim$(nm)
    If oldKey = "" Or oldKey = newKey Then Exit Sub
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim n As Long: n = KeyTables(los, nameCols, heads, labels)
    Dim i As Long, kc As Long, r As Long, names As Variant, keys As Variant, touched As Boolean, cur As Worksheet
    For i = 1 To n
        Set cur = Nothing
        kc = KeyColumnOf(los(i), heads(i))
        If kc > 0 And Not los(i).DataBodyRange Is Nothing Then
            names = RangeTo2D(los(i).ListColumns(nameCols(i)).DataBodyRange)
            keys = RangeTo2D(los(i).ListColumns(kc).DataBodyRange)
            touched = False
            For r = 1 To UBound(names, 1)
                If PayrollText(keys(r, 1)) = oldKey And StrComp(Trim$(SafeStr(names(r, 1))), nm, vbTextCompare) = 0 Then
                    keys(r, 1) = newKey: touched = True
                End If
            Next r
            If touched Then
                Set cur = los(i).Parent
                UnlockSheet cur
                los(i).ListColumns(kc).DataBodyRange.NumberFormat = "@"
                los(i).ListColumns(kc).DataBodyRange.Value = keys
                ProtectSheet cur
                Set cur = Nothing
            End If
        End If
    Next i
    Exit Sub
Oops:
    Dim en As Long, em As String: en = Err.Number: em = Err.Description
    On Error Resume Next                            ' never raises
    If Not cur Is Nothing Then ProtectSheet cur
    LogError "RekeyRows", en, em
End Sub

' v28 (review MED-2): the key columns on the sheets a supervisor or a plain user sees (the OT Log, the posting lists,
' Shift Switches) are hidden in those views, as the roster's column A is; the admin-only sheets are hidden whole
Public Sub HideKeyColumns(ByVal hide As Boolean)
    On Error Resume Next                            ' never raises (a view switch must finish)
    Dim adminOnly As Object: Set adminOnly = CreateObject("Scripting.Dictionary"): adminOnly.CompareMode = 1
    Dim nm As Variant
    For Each nm In AdminOnlySheets()
        adminOnly(CStr(nm)) = True
    Next nm
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim n As Long: n = KeyTables(los, nameCols, heads, labels)
    Dim i As Long, kc As Long
    For i = 1 To n
        If Not adminOnly.Exists(los(i).Parent.Name) Then
            kc = KeyColumnOf(los(i), heads(i))
            If kc > 0 Then
                UnlockSheet los(i).Parent
                los(i).ListColumns(kc).Range.EntireColumn.Hidden = hide
                ProtectSheet los(i).Parent
            End If
        End If
    Next i
End Sub

' ---- the people tables: their name column(s) and the key header each carries ----
Private Function KeyTables(ByRef los() As ListObject, ByRef nameCols() As Long, ByRef heads() As String, ByRef labels() As String) As Long
    ReDim los(1 To 14): ReDim nameCols(1 To 14): ReDim heads(1 To 14): ReDim labels(1 To 14)
    Dim n As Long
    AddKeyTable los, nameCols, heads, labels, n, SheetTable(LOG_SHEET, LOG_TABLE), L_EMP, KEY_HEADER, "OT Log"   ' quiet: LogTable() shows boxes (re-review); a 21-column log has no key column yet and is skipped here
    AddKeyTable los, nameCols, heads, labels, n, SheetTable(POSTINGS_SHEET, PL_TABLE), PL_EMP, KEY_HEADER, "Posting lists"
    AddKeyTable los, nameCols, heads, labels, n, AbsTable(False), AB_EMP, KEY_HEADER, "Absences"
    AddKeyTable los, nameCols, heads, labels, n, SwitchTable(False), SW_OFF, KEY_HEADER_OFF, "Shift switches (gives up)"
    AddKeyTable los, nameCols, heads, labels, n, SwitchTable(False), SW_WORKS, KEY_HEADER_WORKS, "Shift switches (works)"
    AddKeyTable los, nameCols, heads, labels, n, GradeLogTable(False), 2, KEY_HEADER, "Grade Log"
    AddKeyTable los, nameCols, heads, labels, n, HoursHistoryTable(False), 4, KEY_HEADER, "Hours History"
    AddKeyTable los, nameCols, heads, labels, n, CrewTable(False), CH_EMP, KEY_HEADER, "Crew History"
    AddKeyTable los, nameCols, heads, labels, n, PlannedTable(False), PM_EMP, KEY_HEADER, "Planned moves"
    AddKeyTable los, nameCols, heads, labels, n, RstTable(False), RS_EMP, KEY_HEADER, "OT restrictions"
    AddKeyTable los, nameCols, heads, labels, n, EvidenceTable(False), 3, KEY_HEADER, "Evidence"
    AddKeyTable los, nameCols, heads, labels, n, RequestsTable(False), RQ_EMP, KEY_HEADER, "Requests"
    KeyTables = n
End Function

Private Sub AddKeyTable(ByRef los() As ListObject, ByRef nameCols() As Long, ByRef heads() As String, ByRef labels() As String, _
                        ByRef n As Long, ByVal lo As ListObject, ByVal nameCol As Long, ByVal head As String, ByVal label As String)
    If lo Is Nothing Then Exit Sub
    If nameCol > lo.ListColumns.Count Then Exit Sub
    n = n + 1
    Set los(n) = lo
    nameCols(n) = nameCol
    heads(n) = head
    labels(n) = label
End Sub

Private Function SheetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    If Not SheetExists(sheetName) Then Exit Function
    On Error Resume Next                            ' never raises
    Set SheetTable = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

' the key column's index (0 = the table has none yet)
Public Function KeyColumnOf(ByVal lo As ListObject, ByVal head As String) As Long
    Dim c As Long
    For c = lo.ListColumns.Count To 1 Step -1       ' the key sits at the end: start there
        If StrComp(lo.ListColumns(c).Name, head, vbTextCompare) = 0 Then KeyColumnOf = c: Exit Function
    Next c
End Function

' adds the key column on the END of a table that lacks it (a writable workbook only); returns its index, 0 when it
' could not be added. The OT Log's is L_PAYROLL = L_NCOLS: LogLayoutCurrent expects it once added.
Public Function EnsureKeyColumn(ByVal lo As ListObject, ByVal head As String) As Long
    If lo Is Nothing Then Exit Function
    EnsureKeyColumn = KeyColumnOf(lo, head)
    If EnsureKeyColumn > 0 Then Exit Function
    If ThisWorkbook.ReadOnly Then Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim en As Long
    On Error Resume Next
    UnlockSheet ws
    lo.ListColumns.Add.Name = head
    en = Err.Number
    If en = 0 Then
        lo.ListColumns(lo.ListColumns.Count).DataBodyRange.NumberFormat = "@"
        ws.Columns(lo.ListColumns(lo.ListColumns.Count).Range.Column).ColumnWidth = 12
    End If
    ProtectSheet ws
    On Error GoTo 0
    If en = 0 Then EnsureKeyColumn = KeyColumnOf(lo, head)
End Function

' every people table gets its key column (the Install panel's run-once, and the writers' guard)
Public Sub EnsureKeyColumns()
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim n As Long: n = KeyTables(los, nameCols, heads, labels)
    Dim i As Long
    For i = 1 To n
        EnsureKeyColumn los(i), heads(i)
    Next i
End Sub

' ---- STAMP the blank keys from the names (before every save, after an import, after Add / Edit). Returns how
'      many cells were stamped; quiet = no box, errors to the Error Log. Never raises. ----
Public Function FillPayrollKeys(ByVal quiet As Boolean, Optional ByVal addColumns As Boolean = False) As Long
    ' addColumns (review LOW-8): only the verified run-once (AssignPayrollKeys) adds a table's key column; a save, an
    ' import or an Add / Edit stamps the columns that exist (the OT Log's is added by LogTable itself)
    On Error GoTo Oops
    If ThisWorkbook.ReadOnly Then Exit Function
    Dim map As Object: Set map = PayrollMap()
    If map.Count = 0 Then Exit Function
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim n As Long: n = KeyTables(los, nameCols, heads, labels)
    Dim i As Long, kc As Long, r As Long, nm As String, k As String, nStamped As Long, touched As Boolean
    Dim names As Variant, keys As Variant, cur As Worksheet
    For i = 1 To n
        Set cur = Nothing
        If addColumns Then kc = EnsureKeyColumn(los(i), heads(i)) Else kc = KeyColumnOf(los(i), heads(i))
        If kc > 0 And Not los(i).DataBodyRange Is Nothing Then
            names = RangeTo2D(los(i).ListColumns(nameCols(i)).DataBodyRange)
            keys = RangeTo2D(los(i).ListColumns(kc).DataBodyRange)
            touched = False
            For r = 1 To UBound(names, 1)
                If PayrollText(keys(r, 1)) = "" Then
                    nm = Trim$(SafeStr(names(r, 1)))
                    If nm <> "" Then
                        If map.Exists(nm) Then keys(r, 1) = CStr(map(nm)): touched = True: nStamped = nStamped + 1
                    End If
                End If
            Next r
            If touched Then
                Set cur = los(i).Parent
                UnlockSheet cur
                los(i).ListColumns(kc).DataBodyRange.NumberFormat = "@"
                los(i).ListColumns(kc).DataBodyRange.Value = keys
                ProtectSheet cur
                Set cur = Nothing
            End If
        End If
    Next i
    FillPayrollKeys = nStamped
    If Not quiet Then Tell nStamped & " key cell(s) stamped from the roster's Payroll # column.", vbInformation, "Payroll # keys"
    Exit Function
Oops:
    Dim en As Long, em As String: en = Err.Number: em = Err.Description
    On Error Resume Next                            ' never raises
    If Not cur Is Nothing Then ProtectSheet cur     ' (re-review) a sheet unlocked for the write is never left open
    LogError "FillPayrollKeys", en, em
End Function

' the live roster's keys kept current before a save (SaveIfWritable); never raises, never blocks the save
Public Sub KeysBeforeSave()
    On Error Resume Next                            ' never raises
    If gSmokeRun Then Exit Sub
    If Not IsLiveRoster() Then Exit Sub
    FillPayrollKeys True
End Sub

' ---- ADMIN, RUN ONCE at the v28 install (Alt+F8 / the Install panel): the key columns + every row stamped, verified ----
Public Sub AssignPayrollKeys()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If MsgBox("Add the Payroll # key to every people table and fill it from the roster?" & vbNewLine & vbNewLine & _
              "A 'Payroll #' column goes on the END of the OT Log, the posting lists, Absences, Shift Switches (two), the Grade Log, " & _
              "Hours History, Crew History, Planned Moves, Restrictions, Evidence and Requests; every row whose name the roster or the " & _
              "Former sheet knows is stamped. Nothing else moves. The roster's totals are checked before and after. Save a backup first.", _
              vbOKCancel + vbQuestion, "Payroll # keys") <> vbOK Then Exit Sub
    Application.Calculate
    Dim before As Object: Set before = RosterTotalsMap()
    Dim logBefore As Double: logBefore = LogChargedTotal()
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Application.ScreenUpdating = False
    EnsureKeyColumns
    Dim nStamped As Long: nStamped = FillPayrollKeys(True, True)
    Application.ScreenUpdating = scrWas
    Application.Calculate
    Dim diff As String: diff = TotalsDifference(before, RosterTotalsMap())
    If Abs(LogChargedTotal() - logBefore) > 0.005 Then diff = diff & "  the log's charged total moved" & vbNewLine
    If diff <> "" Then
        MarkNoSave
        MsgBox "Something moved that must not:" & vbNewLine & diff & vbNewLine & "SAVING IS BLOCKED. Close WITHOUT saving and reopen. Tell " & PRIMARY_ADMIN_NAME & ".", vbCritical, "Payroll # keys"
        Exit Sub
    End If
    Dim unmatched As String: unmatched = KeylessReport()
    SaveIfWritable
    ShowReceiptNow "Payroll # keys assigned", _
                   "  " & nStamped & " key cell(s) stamped; a Payroll # column on every people table." & vbNewLine, _
                   "  Nobody's Total Hours changed (checked)." & vbNewLine, _
                   "  Every people table (admin); IntegrityCheck 16 'Keys' names what is left." & vbNewLine & IIf(unmatched = "", "", unmatched), _
                   "  Nothing to undo: the columns stay; a wrong number is corrected through Employee > Edit (Payroll #)." & vbNewLine
    Exit Sub
Oops:
    ReportError "AssignPayrollKeys"
End Sub

' the rows still without a key, per table ("" = none)
Private Function KeylessReport() As String
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim n As Long: n = KeyTables(los, nameCols, heads, labels)
    Dim i As Long, kc As Long, r As Long, cnt As Long, names As Variant, keys As Variant
    For i = 1 To n
        kc = KeyColumnOf(los(i), heads(i))
        cnt = 0
        If kc > 0 And Not los(i).DataBodyRange Is Nothing Then
            names = RangeTo2D(los(i).ListColumns(nameCols(i)).DataBodyRange)
            keys = RangeTo2D(los(i).ListColumns(kc).DataBodyRange)
            For r = 1 To UBound(names, 1)
                If Trim$(SafeStr(names(r, 1))) <> "" And PayrollText(keys(r, 1)) = "" Then cnt = cnt + 1
            Next r
        End If
        If cnt > 0 Then KeylessReport = KeylessReport & "  " & labels(i) & ": " & cnt & " row(s) without a key (the name is on neither the roster nor Former with a number)" & vbNewLine
    Next i
End Function

' ---- IntegrityCheck 16 "Keys": rows of (severity, text, fix); sev 2 = PROBLEM, 1 = Check, 0 = Note. Reads only. ----
Public Function KeyProblemList() As Variant
    Dim out() As Variant: ReDim out(1 To 40, 1 To 3)
    Dim n As Long
    ' a number on two roster rows
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String, nm As String
    For r = R_FIRST To RosterLastRow()
        k = PayrollText(ws.Cells(r, C_PAYROLL).Value)
        nm = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If k <> "" And nm <> "" Then
            If seen.Exists(k) Then
                n = n + 1
                If n <= 40 Then out(n, 1) = 2: out(n, 2) = "Payroll # " & k & " is on two roster rows: " & CStr(seen(k)) & " and " & nm: out(n, 3) = "Correct one of them (Employee > Edit, Payroll #)"
            Else
                seen.Add k, nm
            End If
        End If
    Next r
    ' keys in the tables the roster and Former do not hold; rows without a key
    Dim map As Object: Set map = PayrollMap()
    Dim known As Object: Set known = CreateObject("Scripting.Dictionary")
    Dim kk As Variant
    For Each kk In map.Keys
        known(CStr(map(kk))) = True
    Next kk
    Dim los() As ListObject, nameCols() As Long, heads() As String, labels() As String
    Dim nT As Long: nT = KeyTables(los, nameCols, heads, labels)
    ' a roster number the Former sheet holds under another name (review MED-3)
    Dim lo As ListObject: Set lo = FormerTable(False)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing And lo.ListColumns.Count >= FM_PAYROLL Then
            Dim fv As Variant: fv = RangeTo2D(lo.DataBodyRange)
            For r = 1 To UBound(fv, 1)
                k = PayrollText(fv(r, FM_PAYROLL)): nm = Trim$(SafeStr(fv(r, FM_EMP)))
                If k <> "" And nm <> "" Then
                    If seen.Exists(k) Then
                        If StrComp(CStr(seen(k)), nm, vbTextCompare) <> 0 Then
                            n = n + 1
                            If n <= 40 Then out(n, 1) = 1: out(n, 2) = "Payroll # " & k & " is " & CStr(seen(k)) & "'s on the roster and " & nm & "'s on the Former sheet": out(n, 3) = "One of them is wrong: a rename by key would mix their rows (Employee > Edit, Payroll #)"
                        End If
                    End If
                End If
            Next r
        End If
    End If
    Dim i As Long, kc As Long, names As Variant, keys As Variant, nUnknown As Long, nBlank As Long, firstUnknown As String, noCol As String
    For i = 1 To nT
        kc = KeyColumnOf(los(i), heads(i))
        nUnknown = 0: nBlank = 0: firstUnknown = ""
        If kc = 0 Then
            noCol = noCol & IIf(noCol = "", "", ", ") & labels(i)
        ElseIf Not los(i).DataBodyRange Is Nothing Then
            names = RangeTo2D(los(i).ListColumns(nameCols(i)).DataBodyRange)
            keys = RangeTo2D(los(i).ListColumns(kc).DataBodyRange)
            For r = 1 To UBound(names, 1)
                k = PayrollText(keys(r, 1))
                If Trim$(SafeStr(names(r, 1))) <> "" Then
                    If k = "" Then
                        nBlank = nBlank + 1
                    ElseIf Not known.Exists(k) Then
                        nUnknown = nUnknown + 1
                        If firstUnknown = "" Then firstUnknown = k & " (" & Trim$(SafeStr(names(r, 1))) & ")"
                    End If
                End If
            Next r
            If nUnknown > 0 Then
                n = n + 1
                If n <= 40 Then out(n, 1) = 1: out(n, 2) = labels(i) & ": " & nUnknown & " row(s) carry a Payroll # neither the roster nor Former holds, e.g. " & firstUnknown: out(n, 3) = "A renamed or removed person whose number changed: check column A and the Former sheet"
            End If
            If nBlank > 0 Then
                n = n + 1
                If n <= 40 Then out(n, 1) = 0: out(n, 2) = labels(i) & ": " & nBlank & " row(s) without a key (the name has no number on the roster or Former)": out(n, 3) = "Fills itself once the person's Payroll # is known (Add / Edit, or the Kronos import)"
            End If
        End If
    Next i
    If noCol <> "" Then
        n = n + 1
        If n <= 40 Then out(n, 1) = 1: out(n, 2) = "No Payroll # column yet on: " & noCol: out(n, 3) = "Run AssignPayrollKeys (Alt+F8) once after the v28 install"
    End If
    If n = 0 Then Exit Function
    If n > 40 Then n = 40
    Dim res() As Variant: ReDim res(1 To n, 1 To 3)
    Dim j As Long
    For i = 1 To n
        For j = 1 To 3
            res(i, j) = out(i, j)
        Next j
    Next i
    KeyProblemList = res
End Function
