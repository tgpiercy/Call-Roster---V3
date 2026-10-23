Attribute VB_Name = "modEvidence"
Option Explicit

' ====================================================================
'  EVIDENCE (v27, item 23 - briefs/v27.md 1.3): what the parallel run and a
'  grievance need kept outside the live workbook.
'
'  PERIOD CLOSE (ClosePeriod, admin, the Desk's FINISH band): one pay
'  period's OT Log rows, roster totals and postings to a CSV in
'  CLOSED_FOLDER (.\Closed), its SHA-256 worked out by Windows' certutil
'  (through WScript.Shell; "(not available)" when the plant blocks
'  it), a Period Close PDF beside it with the hash in its footer, and one
'  row in tblClosedPeriods (period, when, who, rows, charged, the files,
'  the hash, the PREVIOUS period's hash - so the chain can be checked).
'  Nothing in the workbook changes.
'  SCAN LINK (AddScan, from the award pop-up): a chosen image / PDF copied
'  to SCAN_FOLDER (.\Scans) as <posting>.<ext>, one Evidence row; the FULL
'  history report shows it.
'  DISPUTES (RaiseDispute / CloseDispute, admin): an Evidence row per
'  dispute - the ref, who, the date, the outcome, the note - a SEPARATE
'  table, never a column on the OT Log (VOID logic untouched); Needs
'  Attention lists the open ones; the FULL history report shows them.
'  Nothing here is published.
' ====================================================================

Public Const EVID_VERSION As String = "Evidence v27.1 (Oct 22)"
Public Const EV_SHEET As String = "Evidence"
Public Const EV_TABLE As String = "tblEvidence"
Public Const CP_SHEET As String = "Closed Periods"
Public Const CP_TABLE As String = "tblClosedPeriods"
Private Const EV_N As Long = 10                 ' Ref ID, Kind, Who, Date, Outcome, Note, Path, Entered By, Entered At, Status
Private Const EV_REF As Long = 1
Private Const EV_KIND As Long = 2
Private Const EV_WHO As Long = 3
Private Const EV_DATE As Long = 4
Private Const EV_OUTCOME As Long = 5
Private Const EV_NOTE As Long = 6
Private Const EV_PATH As Long = 7
Private Const EV_USER As Long = 8
Private Const EV_STAMP As Long = 9
Private Const EV_STATUS As Long = 10
Private Const CP_N As Long = 9                  ' Period, Closed At, By, Log Rows, Charged, CSV, PDF, SHA-256, Previous Hash
Public Const EXPORT_KEEP_DAYS As Long = 30      ' v27 (item 25): the daily CSV exports of every table are kept this many days
Private Const EXPORT_DONE As String = "export complete.txt"   ' (review L2) written last: a folder without it is redone
Private mExportLogged As Boolean               ' (review L1) one Error Log row per session, not one per open

' ---- the two tables (made on first use, admin-only, never published) ----
Private Function MakeTable(ByVal sheetName As String, ByVal tableName As String, ByRef heads As Variant, ByVal create As Boolean) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sheetName
            RelockStructure
            SetTabColor ws, RGB(176, 0, 32)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        Dim n As Long: n = UBound(heads) - LBound(heads) + 1
        ws.Range("A1").Resize(1, n).Value = heads
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, n), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleLight9"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set MakeTable = lo
End Function

Public Function EvidenceTable(Optional ByVal create As Boolean = False) As ListObject
    Set EvidenceTable = MakeTable(EV_SHEET, EV_TABLE, Array("Ref ID", "Kind", "Who", "Date", "Outcome", "Note", "Path", "Entered By", "Entered At", "Status"), create)
    If Not EvidenceTable Is Nothing And create Then
        On Error Resume Next
        EvidenceTable.Parent.Columns(4).NumberFormat = "ddd mmm d, yyyy"
        EvidenceTable.Parent.Columns(9).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        EvidenceTable.Parent.Columns(6).ColumnWidth = 50: EvidenceTable.Parent.Columns(7).ColumnWidth = 50
        On Error GoTo 0
    End If
End Function

Public Function ClosedPeriodsTable(Optional ByVal create As Boolean = False) As ListObject
    Set ClosedPeriodsTable = MakeTable(CP_SHEET, CP_TABLE, Array("Period", "Closed At", "By", "Log Rows", "Charged", "CSV", "PDF", "SHA-256", "Previous Hash"), create)
    If Not ClosedPeriodsTable Is Nothing And create Then
        On Error Resume Next
        ClosedPeriodsTable.Parent.Columns(1).NumberFormat = "ddd mmm d, yyyy"
        ClosedPeriodsTable.Parent.Columns(2).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ClosedPeriodsTable.Parent.Columns(6).ColumnWidth = 44: ClosedPeriodsTable.Parent.Columns(7).ColumnWidth = 44
        ClosedPeriodsTable.Parent.Columns(8).ColumnWidth = 66: ClosedPeriodsTable.Parent.Columns(9).ColumnWidth = 66
        On Error GoTo 0
    End If
End Function

' one row appended (plain cells, events off); True when it landed
Private Function AppendRow(ByVal lo As ListObject, ByRef rec As Variant, ByVal nCols As Long) As Boolean
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim ws As Worksheet, cnt As Long
    On Error GoTo Quit
    Application.EnableEvents = False
    Set ws = lo.Parent
    cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, nCols).Value = rec
    ProtectSheet ws
    AppendRow = True
Quit:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
End Function

Private Function EvidenceRows() As Variant
    Dim lo As ListObject: Set lo = EvidenceTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    EvidenceRows = RangeTo2D(lo.DataBodyRange)
End Function

' ---- folders ----
Private Function EnsureFolderPath(ByVal setting As String) As String
    Dim p As String: p = EnsureSlash(ResolveFolder(setting))
    On Error Resume Next
    If Not FolderExists(p) Then MkDir Left$(p, Len(p) - 1)
    On Error GoTo 0
    If FolderExists(p) Then EnsureFolderPath = p
End Function

' ====================================================================
'  THE SCAN LINK
' ====================================================================
' From the award pop-up (a loaded posting): pick the scanned sheet, copy it beside the roster as <posting>.<ext>,
' one Evidence row. Supervisors and admins (the award is theirs).
Public Sub AddScan(ByVal postingId As String)
    On Error GoTo Oops
    If Not IsSupervisor() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    postingId = UCase$(Trim$(postingId))
    If postingId = "" Then Exit Sub
    Dim f As Variant
    f = Application.GetOpenFilename("Scans (*.pdf;*.jpg;*.jpeg;*.png),*.pdf;*.jpg;*.jpeg;*.png", , "The scanned sign-up sheet for " & postingId)
    If VarType(f) = vbBoolean Then Exit Sub
    Dim src As String: src = CStr(f)
    Dim baseName As String: baseName = Mid$(src, InStrRev(src, "\") + 1)     ' (review M4) the extension of the FILE name
    Dim ext As String
    If InStrRev(baseName, ".") > 0 Then ext = LCase$(Mid$(baseName, InStrRev(baseName, ".")))
    Dim folder As String: folder = EnsureFolderPath(SCAN_FOLDER)
    If folder = "" Then MsgBox "The scans folder (" & SCAN_FOLDER & ") could not be made beside the roster.", vbExclamation, "Scan": Exit Sub
    Dim target As String: target = folder & postingId & ext
    Dim n As Long
    Do While Len(Dir$(target)) > 0                 ' a second scan of the same posting keeps the first
        n = n + 1
        target = folder & postingId & "-" & n & ext
    Loop
    FileCopy src, target
    Dim lo As ListObject: Set lo = EvidenceTable(True)
    If lo Is Nothing Then Exit Sub
    Dim rec(1 To 1, 1 To EV_N) As Variant
    rec(1, EV_REF) = postingId: rec(1, EV_KIND) = "scan": rec(1, EV_WHO) = "": rec(1, EV_DATE) = Int(Date)
    rec(1, EV_OUTCOME) = "": rec(1, EV_NOTE) = "'" & "sign-up sheet": rec(1, EV_PATH) = target
    rec(1, EV_USER) = CurrentUser(): rec(1, EV_STAMP) = Now: rec(1, EV_STATUS) = "Filed"
    If Not AppendRow(lo, rec, EV_N) Then MsgBox "The scan was copied to " & target & " but the Evidence row was not written.", vbExclamation, "Scan": Exit Sub
    SaveIfWritable
    Tell "Filed: " & target & vbNewLine & "The FULL employee history shows it beside " & postingId & "'s rows.", vbInformation, "Scan"
    Exit Sub
Oops:
    ' (review M4) called from inside the award pop-up: its own box, never ReportError (which unloads the pop-up
    ' and throws away the supervisor's unsaved ticks)
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    LogError "AddScan", en, em
    MsgBox "The scan was not filed: " & em & vbNewLine & vbNewLine & "The award pop-up is as you left it. Try again, or file it later.", vbExclamation, "Scan"
End Sub

' ====================================================================
'  DISPUTES
' ====================================================================
Public Sub RaiseDispute()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim a As Variant
    a = Application.InputBox(Prompt:="The log entry disputed - its Ref ID (S-0012, P-0041, ADJ-0003):", Title:="Dispute", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim ref As String: ref = UCase$(Trim$(CStr(a)))
    If ref = "" Then Exit Sub
    If Not RefInLog(ref) Then                       ' (review L14) a typo is said, not filed
        If MsgBox(ref & " is not a Ref ID in the OT Log. Record the dispute against it anyway?", vbYesNo + vbQuestion + vbDefaultButton2, "Dispute") = vbNo Then Exit Sub
    End If
    a = Application.InputBox(Prompt:="Who disputes it? Part of the name (the last name is enough):", Title:="Dispute", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim who As String: who = ResolvePersonName(Trim$(CStr(a)), "Dispute")
    If who = "" Then Exit Sub
    Dim note As String: note = AskRequired("What is disputed (one or two sentences; kept on the Evidence sheet, admin-only):", "Dispute")
    If note = "" Then Exit Sub
    Dim lo As ListObject: Set lo = EvidenceTable(True)
    If lo Is Nothing Then Exit Sub
    Dim rec(1 To 1, 1 To EV_N) As Variant
    rec(1, EV_REF) = ref: rec(1, EV_KIND) = "dispute": rec(1, EV_WHO) = who: rec(1, EV_DATE) = Int(Date)
    rec(1, EV_OUTCOME) = "": rec(1, EV_NOTE) = "'" & Left$(note, 300): rec(1, EV_PATH) = ""
    rec(1, EV_USER) = CurrentUser(): rec(1, EV_STAMP) = Now: rec(1, EV_STATUS) = "Open"
    If Not AppendRow(lo, rec, EV_N) Then MsgBox "The dispute was not written.", vbExclamation, "Dispute": Exit Sub
    MarkCalDirty
    SaveIfWritable
    Tell "Dispute recorded against " & ref & " for " & who & ". Needs Attention lists it until CloseDispute records the outcome.", vbInformation, "Dispute"
    Exit Sub
Oops:
    ReportError "RaiseDispute"
End Sub

Public Sub CloseDispute()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim n As Long, txt As String: txt = OpenDisputesText(n)
    If n = 0 Then MsgBox "No open disputes.", vbInformation, "Dispute": Exit Sub
    Dim a As Variant
    a = Application.InputBox(Prompt:="Open disputes:" & vbNewLine & txt & vbNewLine & vbNewLine & "Which one is settled? Type its Ref ID:", Title:="Dispute", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim ref As String: ref = UCase$(Trim$(CStr(a)))
    Dim lo As ListObject: Set lo = EvidenceTable(False)
    Dim v As Variant: v = EvidenceRows()
    If Not IsArray(v) Then Exit Sub
    Dim i As Long, r As Long
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, EV_REF)))) = ref And LCase$(SafeStr(v(i, EV_KIND))) = "dispute" And UCase$(SafeStr(v(i, EV_STATUS))) = "OPEN" Then r = i: Exit For
    Next i
    If r = 0 Then MsgBox "No open dispute on " & ref & ".", vbExclamation, "Dispute": Exit Sub
    Dim outcome As String: outcome = AskRequired("The outcome (upheld / not upheld / withdrawn, and a word on why):", "Dispute")
    If outcome = "" Then Exit Sub
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(r, EV_OUTCOME).Value = "'" & Left$(outcome, 300)
    lo.DataBodyRange.Cells(r, EV_STATUS).Value = "Closed " & Format$(Date, "mmm d") & " by " & CurrentUser()
    ProtectSheet lo.Parent
    Application.EnableEvents = evWas
    MarkCalDirty
    SaveIfWritable
    Tell "Dispute on " & ref & " closed: " & outcome, vbInformation, "Dispute"
    Exit Sub
Oops:
    Application.EnableEvents = True
    ReportError "CloseDispute"
End Sub

' The open disputes, one per line (Needs Attention); count = how many
Public Function OpenDisputesText(ByRef count As Long) As String
    count = 0
    On Error Resume Next                            ' never raises
    Dim v As Variant: v = EvidenceRows()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, s As String
    For i = 1 To UBound(v, 1)
        If LCase$(SafeStr(v(i, EV_KIND))) = "dispute" And UCase$(SafeStr(v(i, EV_STATUS))) = "OPEN" Then
            count = count + 1
            s = s & IIf(s = "", "", vbNewLine) & SafeStr(v(i, EV_REF)) & "  " & CalName(SafeStr(v(i, EV_WHO))) & "  " & _
                IIf(IsDate(v(i, EV_DATE)), Format$(CDate(v(i, EV_DATE)), "mmm d"), "") & "  -  " & Left$(SafeStr(v(i, EV_NOTE)), 80)
        End If
    Next i
    OpenDisputesText = s
End Function

' The FULL history report's Evidence column: "scan: <file>; dispute (Open): ..." for a ref ("" = none)
Public Function EvidenceFor(ByVal ref As String) As String
    On Error Resume Next                            ' never raises
    Dim v As Variant: v = EvidenceRows()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, s As String
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, EV_REF))), Trim$(ref), vbTextCompare) = 0 Then
            If LCase$(SafeStr(v(i, EV_KIND))) = "scan" Then
                s = s & IIf(s = "", "", "; ") & "scan: " & Mid$(SafeStr(v(i, EV_PATH)), InStrRev(SafeStr(v(i, EV_PATH)), "\") + 1)
            Else
                s = s & IIf(s = "", "", "; ") & "dispute (" & SafeStr(v(i, EV_STATUS)) & ") " & CalName(SafeStr(v(i, EV_WHO))) & ": " & Left$(SafeStr(v(i, EV_NOTE)), 60) & _
                    IIf(SafeStr(v(i, EV_OUTCOME)) <> "", " -> " & Left$(SafeStr(v(i, EV_OUTCOME)), 60), "")
            End If
        End If
    Next i
    EvidenceFor = s
End Function

' ====================================================================
'  THE PERIOD CLOSE
' ====================================================================
' A CSV field: quoted, quotes doubled, line breaks to spaces
Private Function Csv(ByVal x As Variant) As String
    Dim s As String
    If IsDate(x) And VarType(x) = vbDate Then
        s = Format$(CDate(x), "yyyy-mm-dd hh:nn")
    Else
        s = SafeStr(x)
    End If
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    Csv = """" & Replace(s, """", """""") & """"
End Function

' SHA-256 of a file through Windows' certutil ("" when it cannot run or the plant blocks it). Late-bound WScript.Shell;
' the output's second line is the hex (older certutil writes it with spaces).
Public Function FileSha256(ByVal path As String) As String
    On Error GoTo Quit
    Dim sh As Object: Set sh = CreateObject("WScript.Shell")
    Dim ex As Object: Set ex = sh.Exec("certutil -hashfile """ & path & """ SHA256")
    Dim out As String: out = ex.StdOut.ReadAll
    Dim lines As Variant: lines = Split(Replace(out, vbCr, ""), vbLf)
    Dim i As Long, t As String
    For i = LBound(lines) To UBound(lines)
        t = Replace(Trim$(CStr(lines(i))), " ", "")
        If Len(t) = 64 And Not t Like "*[!0-9a-fA-F]*" Then FileSha256 = LCase$(t): Exit Function      ' (review L13) every character hex
    Next i
Quit:
End Function

' (review L14) the Ref ID is on some OT Log row
Private Function RefInLog(ByVal ref As String) As Boolean
    Dim lg As Variant: lg = LogRowsNow()
    If Not IsArray(lg) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(lg, 1)
        If StrComp(Trim$(SafeStr(lg(i, L_REF))), ref, vbTextCompare) = 0 Then RefInLog = True: Exit Function
    Next i
End Function

' The newest closed period's hash ("" when none): the next close writes it beside its own
Private Function LastClosedHash(ByRef lastPeriod As Date) As String
    lastPeriod = 0
    On Error Resume Next                            ' never raises
    Dim lo As ListObject: Set lo = ClosedPeriodsTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = UBound(v, 1) To 1 Step -1
        If IsDate(v(i, 1)) Then
            lastPeriod = CDate(v(i, 1))
            LastClosedHash = SafeStr(v(i, 8))
            Exit Function
        End If
    Next i
End Function

' Admin (the Desk's FINISH band, Alt+F8): closes the LAST COMPLETED pay period (asked; any period start may be typed)
Public Sub ClosePeriod()
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub          ' (review L12) a read-only copy would leave files with no Closed Periods row
    Dim ps As Date: ps = PeriodStartFor(Date) - 14
    Dim a As Variant
    a = Application.InputBox(Prompt:="Close which pay period? Its START date (the last completed one is offered):", Title:="Close the pay period", Default:=DateText(ps), Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    If Not ParseTypedDate(Trim$(CStr(a)), ps) Then MsgBox "That isn't a date I can read.", vbExclamation, "Close the pay period": Exit Sub
    ps = PeriodStartFor(ps)
    If ps + 13 >= Date Then
        If MsgBox("The period starting " & Format$(ps, "ddd mmm d") & " is not over yet. Close it anyway (a later close of the same period writes a second row)?", vbYesNo + vbQuestion + vbDefaultButton2, "Close the pay period") = vbNo Then Exit Sub
    End If
    Dim folder As String: folder = EnsureFolderPath(CLOSED_FOLDER)
    If folder = "" Then MsgBox "The folder " & CLOSED_FOLDER & " could not be made beside the roster.", vbExclamation, "Close the pay period": Exit Sub
    Dim prevHash As String, prevPeriod As Date: prevHash = LastClosedHash(prevPeriod)

    Application.ScreenUpdating = False
    Application.Calculate
    ' ---- the CSV: three sections ----
    Dim tag As String: tag = "Period " & Format$(ps, "yyyy-mm-dd")
    Dim csvPath As String: csvPath = folder & tag & " closed " & Format$(Now, "yyyy-mm-dd hhnn") & ".csv"
    Dim ff As Integer: ff = FreeFile
    Open csvPath For Output As #ff
    Print #ff, Csv("Call Out Roster period close") & "," & Csv(ThisWorkbook.Name) & "," & Csv(EXPECTED_RELEASE) & "," & Csv("period start") & "," & Csv(Format$(ps, "yyyy-mm-dd")) & "," & Csv("closed") & "," & Csv(Format$(Now, "yyyy-mm-dd hh:nn")) & "," & Csv("by") & "," & Csv(CurrentUser())
    Print #ff, Csv("previous period") & "," & Csv(IIf(prevPeriod > 0, Format$(prevPeriod, "yyyy-mm-dd"), "")) & "," & Csv("previous hash") & "," & Csv(prevHash)
    ' 1. the OT Log rows of the period (every column; void rows included, marked)
    Print #ff, ""
    Print #ff, Csv("OT LOG ROWS WITH PERIOD START " & Format$(ps, "yyyy-mm-dd"))
    Dim heads As Variant: heads = LogHeaders()
    Dim h As String, i As Long, j As Long
    For j = LBound(heads) To UBound(heads)
        h = h & IIf(h = "", "", ",") & Csv(heads(j))
    Next j
    Print #ff, h
    Dim lg As Variant: lg = LogRowsNow()
    Dim nRows As Long, charged As Double, ln As String
    If IsArray(lg) Then
        For i = UBound(lg, 1) To 1 Step -1              ' oldest first
            If IsDate(lg(i, L_PERIOD)) Then
                If Int(CDate(lg(i, L_PERIOD))) = Int(ps) Then
                    ln = ""
                    For j = 1 To L_NCOLS
                        ln = ln & IIf(j = 1, "", ",") & Csv(lg(i, j))
                    Next j
                    Print #ff, ln
                    nRows = nRows + 1
                    If UCase$(Trim$(SafeStr(lg(i, L_STATUS)))) <> "VOID" Then charged = charged + NumOr0(lg(i, L_CHARGED))
                End If
            End If
        Next i
    End If
    ' 2. the roster as it stands (Total, Refused, the period cells)
    Print #ff, ""
    Print #ff, Csv("ROSTER TOTALS NOW (Total Hours | Refused | the sum of the period cells)")
    Dim tm As Object: Set tm = RosterTotalsMap()
    Dim k As Variant, parts As Variant
    For Each k In tm.Keys
        parts = Split(CStr(tm(k)), "|")
        Print #ff, Csv(k) & "," & Csv(parts(0)) & "," & Csv(parts(1)) & "," & Csv(IIf(UBound(parts) >= 2, parts(2), ""))
    Next k
    ' 3. the postings with a shift in the period
    Print #ff, ""
    Print #ff, Csv("POSTINGS WITH A SHIFT IN THE PERIOD")
    Dim nPost As Long
    If SheetExists(POSTINGS_SHEET) Then             ' (review H2) tblPostings lives on Postings, not the print sheet
        Dim plo As ListObject
        On Error Resume Next
        Set plo = ThisWorkbook.Worksheets(POSTINGS_SHEET).ListObjects(PT_TABLE)
        On Error GoTo Oops
        If Not plo Is Nothing Then
            If Not plo.DataBodyRange Is Nothing Then
                Dim pv As Variant: pv = RangeTo2D(plo.DataBodyRange)
                ln = ""
                For j = 1 To plo.ListColumns.Count
                    ln = ln & IIf(j = 1, "", ",") & Csv(plo.ListColumns(j).Name)
                Next j
                Print #ff, ln
                For i = 1 To UBound(pv, 1)
                    If IsDate(pv(i, PS_DATE)) Then
                        If Int(CDate(pv(i, PS_DATE))) >= Int(ps) And Int(CDate(pv(i, PS_DATE))) <= Int(ps) + 13 Then
                            ln = ""
                            For j = 1 To UBound(pv, 2)
                                ln = ln & IIf(j = 1, "", ",") & Csv(pv(i, j))
                            Next j
                            Print #ff, ln
                            nPost = nPost + 1
                        End If
                    End If
                Next i
            End If
        End If
    End If
    Close #ff
    ff = 0

    ' ---- the hash, the PDF, the row ----
    Dim hash As String: hash = FileSha256(csvPath)
    Dim ws As Worksheet: Set ws = ReportSheetNamed(CP_SHEET & " print")
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Range("A1").Value = "PAY PERIOD CLOSE  -  " & Format$(ps, "ddd mmm d") & " to " & Format$(ps + 13, "ddd mmm d, yyyy")
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "Closed " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & " by " & CurrentUser() & "  (" & ThisWorkbook.Name & ", " & EXPECTED_RELEASE & ")"
    ws.Range("A4").Value = "OT Log rows with this period start: " & nRows & "   Charged hours (non-void): " & HoursText(charged) & "   Postings with a shift in the period: " & nPost
    ws.Range("A5").Value = "People on the roster: " & tm.Count
    ws.Range("A7").Value = "CSV: " & csvPath
    ws.Range("A8").Value = "SHA-256 of the CSV: " & IIf(hash = "", "(hash not available on this PC - certutil did not answer)", hash)
    ws.Range("A9").Value = "Previous period: " & IIf(prevPeriod > 0, Format$(prevPeriod, "ddd mmm d, yyyy") & "  hash " & prevHash, "none closed before this one")
    ws.Range("A11").Value = "How to check later: certutil -hashfile ""<the CSV>"" SHA256 must print the hash above; the next period's close carries this hash as its Previous Hash."
    ws.Columns(1).ColumnWidth = 120
    ws.Range("A8").Font.Bold = True
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .CenterFooter = "SHA-256 " & IIf(hash = "", "(not available)", hash)
    End With
    ProtectSheet ws
    Dim pdfBase As String: pdfBase = tag & " close " & Format$(Now, "hhnn")      ' (review L12) a same-day re-close keeps the first PDF
    ReportPdf ws, pdfBase, "Close the pay period", False, folder
    Dim pdfPath As String: pdfPath = folder & pdfBase & " " & Format$(Date, "yyyy-mm-dd") & ".pdf"
    SetSheetVisible ws, xlSheetVeryHidden

    Dim lo As ListObject: Set lo = ClosedPeriodsTable(True)
    Dim rec(1 To 1, 1 To CP_N) As Variant
    rec(1, 1) = Int(ps): rec(1, 2) = Now: rec(1, 3) = CurrentUser(): rec(1, 4) = nRows: rec(1, 5) = charged
    rec(1, 6) = csvPath: rec(1, 7) = pdfPath: rec(1, 8) = IIf(hash = "", "(not available)", hash): rec(1, 9) = prevHash
    If Not AppendRow(lo, rec, CP_N) Then MsgBox "The files were written but the Closed Periods row was not.", vbExclamation, "Close the pay period"
    SaveIfWritable
    Application.ScreenUpdating = scrWas
    Application.EnableEvents = evWas
    MsgBox "Period " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d") & " closed." & vbNewLine & vbNewLine & _
           nRows & " log rows, " & HoursText(charged) & " h charged, " & nPost & " postings." & vbNewLine & _
           "CSV + PDF in " & folder & vbNewLine & "SHA-256: " & IIf(hash = "", "(not available on this PC)", hash), vbInformation, "Close the pay period"
    Exit Sub
Oops:
    ReportError "ClosePeriod"                      ' (review M3) first, while Err still holds the error
    On Error Resume Next
    If ff <> 0 Then Close #ff
    SetSheetVisible ThisWorkbook.Worksheets(CP_SHEET & " print"), xlSheetVeryHidden   ' (review L12) never left showing
    Application.ScreenUpdating = scrWas
    Application.EnableEvents = evWas
End Sub

' ====================================================================
'  THE DAILY TABLE EXPORT (v27 item 25, 1.4)
' ====================================================================
' Workbook_Open step 12: the first ADMIN open of the live roster each day writes the tables a SUPERVISOR's view
' shows to CSV in <backup folder>\csv\<yyyy-mm-dd>\ (one file per table, every column, every row whatever the view),
' then prunes the day folders older than EXPORT_KEEP_DAYS. Plain text a person can open without Excel: what a
' later SharePoint List would be loaded from. (Review of item 25, HIGH: the folder sits beside the roster where
' every user can list it, so the admin-only tables - absences with their leave types, the error log, feedback,
' evidence, restrictions, the Kronos tables - and the very hidden ones are NOT written; they live in the daily
' .xlsm backup.) A practice, backup or read-only copy does nothing. Quiet: a failure goes to the Error Log once a
' session, never a box.
Public Sub ExportTablesOnOpen()
    On Error GoTo Quit
    If Not IsAdmin() Then Exit Sub                  ' (review H1) the brief's "admin, daily on open"
    If ThisWorkbook.ReadOnly Then Exit Sub
    If IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name) Then Exit Sub
    If Trim$(BACKUP_FOLDER) = "" Then Exit Sub
    Dim root As String: root = EnsureSlash(BackupFolderPath())
    If Not FolderExists(root) Then Exit Sub             ' the backup step already said so
    Dim csvRoot As String: csvRoot = root & "csv\"
    If Not FolderExists(csvRoot) Then MkDir Left$(csvRoot, Len(csvRoot) - 1)
    Dim dayDir As String: dayDir = csvRoot & Format$(Date, "yyyy-mm-dd") & "\"
    If Len(Dir$(dayDir & EXPORT_DONE)) > 0 Then Exit Sub    ' already taken today (review L2: a half-written folder is redone)
    If Not FolderExists(dayDir) Then MkDir Left$(dayDir, Len(dayDir) - 1)
    ExportTablesTo dayDir
    MarkExportDone dayDir
    PruneCsvExports csvRoot
    Exit Sub
Quit:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    If Not mExportLogged Then mExportLogged = True: LogError "ExportTablesOnOpen", en, em
End Sub

' the marker that says the day's folder is complete
Private Sub MarkExportDone(ByVal folder As String)
    Dim ff As Integer: ff = FreeFile
    Open folder & EXPORT_DONE For Output As #ff
    Print #ff, "Call Out Roster CSV export " & Format$(Now, "yyyy-mm-dd hh:nn") & " by " & CurrentUser() & " (" & EXPECTED_RELEASE & ")"
    Close #ff
End Sub

' Alt+F8 (admin): the same export now, into a folder of its own (<date> <hhnn>), and says where
Public Sub ExportTables()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim root As String: root = EnsureFolderPath(BACKUP_FOLDER)
    If root = "" Then MsgBox "The backup folder (" & BACKUP_FOLDER & ") can't be reached.", vbExclamation, "Export tables": Exit Sub
    Dim csvRoot As String: csvRoot = root & "csv\"
    If Not FolderExists(csvRoot) Then MkDir Left$(csvRoot, Len(csvRoot) - 1)
    Dim dayDir As String: dayDir = csvRoot & Format$(Now, "yyyy-mm-dd hhnn") & "\"
    If Not FolderExists(dayDir) Then MkDir Left$(dayDir, Len(dayDir) - 1)
    Dim n As Long: n = ExportTablesTo(dayDir)
    MarkExportDone dayDir
    Tell n & " tables written as CSV to:" & vbNewLine & dayDir & vbNewLine & vbNewLine & "The admin-only tables are not written beside the roster: they are in the daily backup copy.", vbInformation, "Export tables"
    Exit Sub
Oops:
    ReportError "ExportTables"
End Sub

' Every ListObject on a sheet a supervisor's view shows -> <folder><table name>.csv (an admin-only sheet, a very
' hidden one and an archived year are skipped). Returns how many were written. An error closes the open file and
' goes up to the caller.
Private Function ExportTablesTo(ByVal folder As String) As Long
    Dim ws As Worksheet, lo As ListObject, ff As Integer, v As Variant, i As Long, j As Long, ln As String
    On Error GoTo Fail
    Dim skip As Object: Set skip = CreateObject("Scripting.Dictionary")
    skip.CompareMode = 1
    Dim nm As Variant
    For Each nm In AdminOnlySheets()
        skip(CStr(nm)) = True
    Next nm
    For Each ws In ThisWorkbook.Worksheets
        If skip.Exists(ws.Name) Or ws.Visible = xlSheetVeryHidden Or IsYearEndSheet(ws.Name) Then GoTo NextSheet
        For Each lo In ws.ListObjects
            ff = FreeFile
            Open folder & lo.Name & ".csv" For Output As #ff
            ln = ""
            For j = 1 To lo.ListColumns.Count
                ln = ln & IIf(j = 1, "", ",") & Csv(lo.ListColumns(j).Name)
            Next j
            Print #ff, ln
            If Not lo.DataBodyRange Is Nothing Then
                v = RangeTo2D(lo.DataBodyRange)
                For i = 1 To UBound(v, 1)
                    ln = ""
                    For j = 1 To UBound(v, 2)
                        ln = ln & IIf(j = 1, "", ",") & Csv(v(i, j))
                    Next j
                    Print #ff, ln
                Next i
            End If
            Close #ff
            ff = 0
            ExportTablesTo = ExportTablesTo + 1
        Next lo
NextSheet:
    Next ws
    Exit Function
Fail:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    If ff <> 0 Then Close #ff
    Err.Raise en, "ExportTablesTo", em
End Function

' The day folders (named yyyy-mm-dd ...) older than EXPORT_KEEP_DAYS: their CSV files deleted, then the folder.
' Only this export's own folders and files are touched. Never raises.
Private Sub PruneCsvExports(ByVal csvRoot As String)
    On Error Resume Next                            ' never raises
    Dim dirs As Collection: Set dirs = New Collection
    Dim nm As String: nm = Dir$(csvRoot & "*", vbDirectory)
    Do While nm <> ""
        If nm Like "####-##-##" Or nm Like "####-##-## ####" Then dirs.Add nm      ' (review L3) the export's own names only
        nm = Dir$()
    Loop
    Dim k As Variant, d As Date, files As Collection, f As Variant
    For Each k In dirs
        d = 0
        d = DateSerial(CLng(Left$(k, 4)), CLng(Mid$(k, 6, 2)), CLng(Mid$(k, 9, 2)))
        If d > 0 And Date - d > EXPORT_KEEP_DAYS Then
            Set files = New Collection
            nm = Dir$(csvRoot & k & "\*.csv")
            Do While nm <> ""
                If LCase$(Left$(nm, 3)) = "tbl" Then files.Add nm     ' (review L3) a table's file, nothing else
                nm = Dir$()
            Loop
            For Each f In files
                Kill csvRoot & k & "\" & f
            Next f
            If Len(Dir$(csvRoot & k & "\" & EXPORT_DONE)) > 0 Then Kill csvRoot & k & "\" & EXPORT_DONE
            RmDir csvRoot & k                       ' fails quietly when something else was put in it
        End If
    Next k
End Sub
