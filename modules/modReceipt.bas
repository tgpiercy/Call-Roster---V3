Attribute VB_Name = "modReceipt"
Option Explicit

' ====================================================================
'  RECEIPTS (v27, item 7 - briefs/v27.md; review 5.1)
'
'  One card after every save, one layout on every form (Graeme, Oct 20):
'    WHAT WAS RECORDED   one line per row written (Ref ID, type, person, hours)
'    WHO IT CHANGED      "Joe Smith: 120 -> 141 h" from the roster's Total Hours
'                        before and after (RosterTotalsMap), or "nobody"
'    WHERE TO SEE IT     the calendar line, the OT Log rows, the sheet, the
'                        members' file after the next update
'    IF THIS IS WRONG    Undo last / Void Entry for an admin; "ask <admin>" for
'                        a supervisor
'  and the trust strip at the foot. Print receipt saves a PDF beside the
'  roster (the Receipt output sheet, redrawn each time); Email me this opens
'  a mail draft with the text. The card is frmReceipt (code-built); it reads
'  gReceipt. Under SmokeTest the text goes to gReceiptLog instead of the
'  card, so the steps can assert its lines. Holds no rule: every figure
'  comes from the save that just happened.
' ====================================================================

Public Const RCPT_VERSION As String = "Receipt v29.0.1 (Oct 28)"     ' stamped with the release at release prep (v27)
Public Const RECEIPT_SHEET As String = "Receipt"                     ' output only: drawn for a print, hidden again
' v27 (item 9): "Was this right?" - every answer given on a receipt card lands here (admin-only, never published)
Public Const FB_SHEET As String = "Feedback"
Public Const FB_TABLE As String = "tblFeedback"
Private Const FB_N As Long = 7              ' When, Login, Release, Ref ID, Answer, Note, Seen (the columns an answer writes)
Private Const FB_SEEN As Long = 7
Private Const FB_RESBY As Long = 8          ' v29 (item 3a / 4): Resolved By - who fixed the entry (Fix an entry), on the END
Private Const FB_RESREF As Long = 9         ' v29: Resolved Ref - the Ref ID the re-entry got ("void" when voided and not re-entered)
Private Const FB_NFULL As Long = 9          ' every column since v29; an older table gets the two added once (EnsureFeedbackColumns)
Private Const FB_MAX_NOTE As Long = 300

Public Type Receipt
    Title As String
    Recorded As String
    Changed As String
    WhereTxt As String
    WrongTxt As String
    Trust As String
End Type

Public gReceipt As Receipt
Public gReceiptLog As String                ' SmokeTest: every receipt's text, appended (the card is not shown)

' v27 (item 8): the before / after marks on the roster. ReceiptChanges keeps every name whose Total Hours
' moved ("before|delta"); MarkChangedTotals paints those G cells after the card; ClearTotalMarks takes
' them off (ApplyView, the next save, the first click on the roster MARK_SECONDS after the paint).
Private mMarks As Object                    ' name -> "before|delta" from the last ReceiptChanges
Private mMarksFresh As Boolean              ' True between a ReceiptChanges and the paint that used it
Private mMarkedAt As Date                   ' when the marks were painted (0 = none on the roster)
Public Const MARK_SECONDS As Long = 60      ' the marks come off at the first click on the roster after this
Private Const MARK_TAG As String = "Was "   ' every mark's note starts with it: how ClearTotalMarks finds them
Private Const MARK_FILL As Long = 13104127  ' RGB(255, 243, 199): a light amber on the moved Total Hours cells (the banding's
                                            ' conditional format wins on every second visible row - review M3 - so the FONT marks too)
Private Const MARK_INK As Long = 23470      ' RGB(174, 91, 0): the amber ink on a marked Total Hours cell

' The roster's Total Hours before a save, for ReceiptChanges (name -> "Total|Refused|periods")
Public Function ReceiptBefore() As Object
    ClearTotalMarks                               ' v27 (item 8): the last save's marks come off before this save's go on
    Set mMarks = Nothing: mMarksFresh = False      ' (review M2) and a map a failed save left behind is forgotten
    On Error Resume Next                          ' (review M4) a receipt must never stop a save: an empty map when the roster cannot be read (never raises)
    Set ReceiptBefore = CreateObject("Scripting.Dictionary")
    Set ReceiptBefore = RosterTotalsMap()
End Function

' Who the save changed: every roster name whose Total Hours moved, "name: before -> after h (+d)"; new or gone names
' said so; "Nobody's Total Hours changed." when nothing moved. Never raises (a receipt must not stop a save).
Public Function ReceiptChanges(ByVal before As Object) As String
    On Error GoTo Quiet
    If before Is Nothing Then ReceiptChanges = "  (the before figures could not be read)" & vbNewLine: Exit Function
    If before.Count = 0 Then ReceiptChanges = "  (the before figures could not be read)" & vbNewLine: Exit Function   ' (review) never 'everyone is new'
    Application.Calculate
    Set mMarks = CreateObject("Scripting.Dictionary")     ' v27 (item 8): the names this save moved, for the roster marks
    mMarks.CompareMode = 1
    mMarksFresh = True
    Dim after As Object: Set after = RosterTotalsMap()
    Dim k As Variant, s As String, tb As Double, ta As Double, n As Long
    For Each k In after.Keys
        ta = TotalOf(CStr(after(k)))
        If before.Exists(k) Then
            tb = TotalOf(CStr(before(k)))
            If Abs(ta - tb) > 0.001 Then
                n = n + 1
                mMarks(k) = HoursText(tb) & "|" & HoursText(ta - tb)
                If n <= 12 Then s = s & "  " & k & ": " & HoursText(tb) & " -> " & HoursText(ta) & " h (" & IIf(ta > tb, "+", "") & HoursText(ta - tb) & ")" & vbNewLine
            End If
        Else
            n = n + 1
            If n <= 12 Then s = s & "  " & k & ": new on the roster, " & HoursText(ta) & " h" & vbNewLine
        End If
    Next k
    For Each k In before.Keys
        If Not after.Exists(k) Then
            n = n + 1
            If n <= 12 Then s = s & "  " & k & ": no longer on the roster" & vbNewLine
        End If
    Next k
    If n > 12 Then s = s & "  ... and " & (n - 12) & " more" & vbNewLine
    If s = "" Then s = "  Nobody's Total Hours changed." & vbNewLine
    ReceiptChanges = s
    Exit Function
Quiet:
    ReceiptChanges = "  (the before / after figures could not be read: " & Err.Description & ")" & vbNewLine
End Function

Private Function TotalOf(ByVal packed As String) As Double
    Dim p As Long: p = InStr(packed, "|")
    If p > 0 Then packed = Left$(packed, p - 1)
    TotalOf = Val(Replace(packed, ",", "."))
End Function

' The "if this is wrong" line for whoever is logged in
Public Function WrongText() As String
    If IsAdmin() Then
        WrongText = "  Undo my last log entry (the Desk) voids your newest OT Log entry of today; Void Entry on the bar voids any Ref ID." & vbNewLine
    Else
        WrongText = "  Ask " & PRIMARY_ADMIN_NAME & " to void it (Void Entry); the rows stay in the log, marked VOID." & vbNewLine
    End If
End Function

' Fills the receipt (every part a block of lines, "" allowed) and shows it
Public Sub ShowReceiptNow(ByVal title As String, ByVal recorded As String, ByVal changed As String, ByVal whereTxt As String, ByVal wrongTxt As String)
    gReceipt.Title = title
    gReceipt.Recorded = recorded
    gReceipt.Changed = changed
    gReceipt.WhereTxt = whereTxt
    gReceipt.WrongTxt = wrongTxt
    ShowReceipt
End Sub

' Shows gReceipt: the card, or (SmokeTest) the text into gReceiptLog; a card that cannot open falls back to a box
Public Sub ShowReceipt()
    gReceipt.Trust = TrustStrip()
    MarkChangedTotals                            ' v27 (item 8): the moved Total Hours cells marked before the card (SmokeTest reads the marks too)
    ' (review M7) a save that failed, or saving blocked: the card says so first - it must never read as a saved entry
    If Not gSmokeRun And (gSaveFailedAt > 0 Or NoSave()) Then
        If Left$(gReceipt.Title, 9) <> "NOT SAVED" Then gReceipt.Title = "NOT SAVED - " & gReceipt.Title
        gReceipt.WhereTxt = "  " & SaveBlockText() & vbNewLine & gReceipt.WhereTxt
    End If
    If gSmokeRun Then
        gReceiptLog = gReceiptLog & vbNewLine & "---- " & ReceiptText(gReceipt)
        Exit Sub
    End If
    On Error GoTo Fallback
    frmReceipt.Show
    RestartMarkClock                             ' (review M4) the MARK_SECONDS run from the card's close, not its paint
    Exit Sub
Fallback:
    LogError "ShowReceipt", Err.Number, Err.Description
    MsgBox ReceiptText(gReceipt), vbInformation, IIf(gReceipt.Title = "", "Saved", gReceipt.Title)
    RestartMarkClock
End Sub

' ====================================================================
'  v27 (item 8): BEFORE / AFTER ON THE ROSTER
'  After a save, every Total Hours cell the save moved carries a light fill
'  and a note "Was 120, +21 from S-0012 (3:10 AM)". The marks come off at
'  ApplyView (every view change, every saved copy), the next save
'  (ReceiptBefore) or the first click on the roster MARK_SECONDS after
'  the paint (RosterSelectionChanged asks TotalMarksExpired). No OnTime
'  timer: it would reopen a closed roster. Nothing here raises.
' ====================================================================

' Paints the marks from the last ReceiptChanges (once per receipt: a card without a ReceiptChanges of its
' own - a cancelled posting - repaints nothing)
Public Sub MarkChangedTotals()
    Dim ws As Worksheet, n As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents    ' (review M1) captured BEFORE On Error: Quiet restores it on every path
    On Error GoTo Quiet
    If Not mMarksFresh Then Exit Sub
    mMarksFresh = False
    If mMarks Is Nothing Then Exit Sub
    If mMarks.Count = 0 Then Exit Sub
    Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    Dim refTxt As String: refTxt = FirstRefIn(gReceipt.Recorded)
    If refTxt = "" Then refTxt = FirstRefIn(gReceipt.Title)
    If refTxt = "" Then refTxt = "this save"
    Dim markTime As String: markTime = Format$(Now, "h:mm AM/PM")     ' v27.1: markTime (stampTxt would hide modCalendar.StampTxt)
    Dim names As Variant: names = RangeTo2D(ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)))   ' a one-person roster is still 2-D
    Application.EnableEvents = False
    UnlockSheet ws
    Dim i As Long, nm As String, parts() As String
    For i = 1 To UBound(names, 1)
        nm = Trim$(SafeStr(names(i, 1)))
        If nm <> "" Then
            If mMarks.Exists(nm) Then
                parts = Split(CStr(mMarks(nm)), "|")
                With ws.Cells(R_FIRST + i - 1, C_TOTAL)
                    ' (review L6) the note first, then the colours: a cell whose note is not ours (or holds a threaded
                    ' comment, which AddComment refuses) is skipped whole, so no colour is ever left without the note that clears it
                    If MarkableCell(ws.Cells(R_FIRST + i - 1, C_TOTAL)) Then
                        If Not .Comment Is Nothing Then .Comment.Delete
                        On Error Resume Next                 ' (re-review L2) a threaded comment (Range.Comment is Nothing for it) makes AddComment raise: skip the cell, go on
                        .AddComment MARK_TAG & parts(0) & ", " & IIf(Left$(parts(1), 1) = "-", "", "+") & parts(1) & " from " & refTxt & " (" & markTime & ")"
                        Dim added As Boolean: added = (Err.Number = 0)
                        Err.Clear
                        On Error GoTo Quiet
                        If added Then
                            .Comment.Visible = False
                            .Interior.Color = MARK_FILL
                            .Font.Color = MARK_INK
                            n = n + 1
                        End If
                    End If
                End With
            End If
        End If
    Next i
    ProtectSheet ws
    Application.EnableEvents = evWas
    If n > 0 Then mMarkedAt = Now
    Exit Sub
Quiet:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    If n > 0 Then mMarkedAt = Now                  ' (re-review L3) the cells painted before the error still expire on a click
    LogError "MarkChangedTotals", en, em
End Sub

' Takes every mark off the roster (the fill and the note; a Total Hours note that is not ours is left alone).
' Safe to call any time, on any sheet, in any view; a roster that cannot be read is left as it is.
Public Sub ClearTotalMarks()
    Dim ws As Worksheet
    Dim evWas As Boolean: evWas = Application.EnableEvents    ' (review M1) captured before On Error
    On Error GoTo Quiet
    mMarkedAt = 0
    Set ws = RosterSheet()
    If ws.Comments.Count = 0 Then Exit Sub         ' the note is the mark: a colour is never left behind without one
    Application.EnableEvents = False
    UnlockSheet ws
    Dim i As Long, c As Comment
    For i = ws.Comments.Count To 1 Step -1
        Set c = ws.Comments(i)
        If c.Parent.Column = C_TOTAL And c.Parent.Row >= R_FIRST Then
            If Left$(c.Text, Len(MARK_TAG)) = MARK_TAG Then
                c.Parent.Interior.ColorIndex = xlNone
                c.Parent.Font.Color = 3355443        ' the roster's ink, RGB(51, 51, 51) (modUI.INK; re-review L4 - xlAutomatic is black)
                c.Delete
            End If
        End If
    Next i
    ProtectSheet ws
    Application.EnableEvents = evWas
    Exit Sub
Quiet:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    LogError "ClearTotalMarks", en, em
End Sub

' True when the cell can take our note: no note, or ours (a stranger's note, or a threaded comment, is left alone)
Private Function MarkableCell(ByVal cell As Range) As Boolean
    On Error GoTo No
    If cell.Comment Is Nothing Then MarkableCell = True: Exit Function
    MarkableCell = (Left$(cell.Comment.Text, Len(MARK_TAG)) = MARK_TAG)
    Exit Function
No:
    MarkableCell = False
End Function

' (review M4) the MARK_SECONDS count from the moment the card closes: the reader may sit on it for a while
Public Sub RestartMarkClock()
    If mMarkedAt > 0 Then mMarkedAt = Now
End Sub

' True once marks have sat on the roster for MARK_SECONDS (the next click takes them off)
Public Function TotalMarksExpired() As Boolean
    If mMarkedAt = 0 Then Exit Function
    TotalMarksExpired = ((Now - mMarkedAt) * 86400 >= MARK_SECONDS)
End Function

' The mark on one person's Total Hours cell ("" = none): SmokeTest step 4 reads it
Public Function TotalMarkText(ByVal emp As String) As String
    On Error Resume Next                            ' never raises
    Dim r As Long: r = RosterRowOf(emp)
    If r < R_FIRST Then Exit Function
    Dim c As Comment: Set c = RosterSheet().Cells(r, C_TOTAL).Comment
    If c Is Nothing Then Exit Function
    If Left$(c.Text, Len(MARK_TAG)) = MARK_TAG Then TotalMarkText = c.Text
End Function

' The first Ref ID in a block of text ("S-0012", "ADJ-0003", "P-0041"): one to four capital letters, a dash,
' one or more digits, standing on its own (no letter or digit touching either end - so a date, "Test-1" or
' "LIEU OT - DT" is not one); "" when there is none. Pure: suite firstref (tests/rules.py first_ref).
Public Function FirstRefIn(ByVal txt As String) As String
    Dim n As Long, p As Long, a As Long, b As Long, ch As String
    n = Len(txt)
    p = InStr(1, txt, "-")
    Do While p > 0
        a = p - 1                                   ' walks back over the capital letters
        Do While a >= 1
            ch = Mid$(txt, a, 1)
            If ch < "A" Or ch > "Z" Then Exit Do
            a = a - 1
        Loop
        b = p + 1                                   ' walks forward over the digits
        Do While b <= n
            ch = Mid$(txt, b, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            b = b + 1
        Loop
        If p - 1 - a >= 1 And p - 1 - a <= 4 And b - p - 1 >= 1 Then
            If RefEdgeFree(txt, a) And RefEdgeFree(txt, b) Then
                FirstRefIn = Mid$(txt, a + 1, b - a - 1)
                Exit Function
            End If
        End If
        p = InStr(p + 1, txt, "-")
    Loop
End Function

' True when position i is off the text's ends or holds something that is not a letter or a digit
Private Function RefEdgeFree(ByVal txt As String, ByVal i As Long) As Boolean
    If i < 1 Or i > Len(txt) Then RefEdgeFree = True: Exit Function
    RefEdgeFree = Not (Mid$(txt, i, 1) Like "[A-Za-z0-9]")
End Function

' ====================================================================
'  v27 (item 9): "WAS THIS RIGHT?" - the three answers on the card
'  Yes / Not sure / No, with an optional line, go into tblFeedback on the
'  admin-only Feedback sheet (When, Login, Release, Ref ID, Answer, Note,
'  Seen), modelled on the Error Log: made on first use, never published,
'  pruned by hand. Needs Attention lists the unread ones (UnseenFeedbackCount);
'  ShowFeedback (admin, Alt+F8 or that line) opens the sheet and marks them seen.
'  Nothing here raises.
' ====================================================================

' The table, or Nothing. create:=True adds the sheet and the table.
Public Function FeedbackTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(FB_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(FB_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(FB_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet     ' adding a sheet activates it - put the user back
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = FB_SHEET
            RelockStructure
            SetTabColor ws, RGB(30, 56, 100)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, FB_NFULL).Value = Array("When", "Login", "Release", "Ref ID", "Answer", "Note", "Seen", "Resolved By", "Resolved Ref")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, FB_NFULL), , xlYes)
        lo.Name = FB_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 20
        ws.Columns(2).ColumnWidth = 12
        ws.Columns(3).ColumnWidth = 9
        ws.Columns(4).ColumnWidth = 10
        ws.Columns(5).ColumnWidth = 9
        ws.Columns(6).ColumnWidth = 60
        ws.Columns(7).ColumnWidth = 7
        ws.Columns(FB_RESBY).ColumnWidth = 12
        ws.Columns(FB_RESREF).ColumnWidth = 12
        ws.Columns(1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    If Not lo Is Nothing Then EnsureFeedbackColumns lo
    Set FeedbackTable = lo
End Function

' v29: a table made before v29 has 7 columns; Resolved By / Resolved Ref go on the END once (as the Grade Log's Field did)
Private Sub EnsureFeedbackColumns(ByVal lo As ListObject)
    If lo.ListColumns.Count >= FB_NFULL Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    If lo.ListColumns.Count < FB_RESBY Then lo.ListColumns.Add.Name = "Resolved By"
    If lo.ListColumns.Count < FB_RESREF Then lo.ListColumns.Add.Name = "Resolved Ref"
    ws.Columns(FB_RESBY).ColumnWidth = 12
    ws.Columns(FB_RESREF).ColumnWidth = 12
    ProtectSheet ws
    On Error GoTo 0
End Sub

' v29 (the Lite Desk's "Corrections asked for" card): the No / Not sure answers nobody has resolved yet (Resolved Ref
' blank). Never raises.
Public Function UnresolvedNoCount() As Long
    Dim v As Variant: v = UnresolvedNoRefs()
    If IsArray(v) Then UnresolvedNoCount = UBound(v, 1)
End Function

' The unresolved No / Not sure answers, oldest first: a 2-D array (1 To n, 1 To 4) of Ref ID, Answer, Note, When
' (Fix an entry's pick list). Empty when there are none. Never raises.
Public Function UnresolvedNoRefs() As Variant
    Dim lo As ListObject, v As Variant, i As Long, n As Long
    On Error GoTo Quit
    Set lo = FeedbackTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = RangeTo2D(lo.DataBodyRange)
    Dim hasRes As Boolean: hasRes = (UBound(v, 2) >= FB_RESREF)
    Dim out() As Variant: ReDim out(1 To UBound(v, 1), 1 To 4)
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, 1)) <> "" And SafeStr(v(i, 4)) <> "" Then
            If UCase$(SafeStr(v(i, 5))) <> "YES" Then
                If Not hasRes Or SafeStr(v(i, FB_RESREF)) = "" Then
                    n = n + 1
                    out(n, 1) = SafeStr(v(i, 4)): out(n, 2) = SafeStr(v(i, 5)): out(n, 3) = SafeStr(v(i, 6)): out(n, 4) = v(i, 1)
                End If
            End If
        End If
    Next i
    If n = 0 Then Exit Function
    Dim res() As Variant: ReDim res(1 To n, 1 To 4)
    Dim k As Long
    For i = 1 To n
        For k = 1 To 4
            res(i, k) = out(i, k)
        Next k
    Next i
    UnresolvedNoRefs = res
Quit:
End Function

' v29 (Fix an entry): every unresolved No / Not sure row for oldRef is marked resolved by this login with newRef (the
' re-entry's Ref ID, or "void" when the entry was voided and not re-entered). The count of rows marked; 0 when none or
' the copy is read-only (the caller's VoidRef has already refused a NoSave copy; SmokeTest, which never saves, runs this
' on its practice copy). Never raises.
Public Function ResolveFeedback(ByVal oldRef As String, ByVal newRef As String) As Long
    ' v29 (re-review LOW-9): a row already marked "checked" / "already void" (looked at, left) is resolved again by a real
    ' outcome - a void or a new Ref - so the card's later void records what happened
    Dim lo As ListObject, v As Variant, i As Long, ws As Worksheet
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Quit
    If ThisWorkbook.ReadOnly Then Exit Function
    Set lo = FeedbackTable(True)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If lo.ListColumns.Count < FB_NFULL Then Exit Function
    Set ws = lo.Parent
    v = RangeTo2D(lo.DataBodyRange)
    Dim by As Variant: by = RangeTo2D(lo.ListColumns(FB_RESBY).DataBodyRange)
    Dim rf As Variant: rf = RangeTo2D(lo.ListColumns(FB_RESREF).DataBodyRange)
    oldRef = UCase$(Trim$(oldRef))
    Dim realOutcome As Boolean: realOutcome = (LCase$(Trim$(newRef)) <> "checked" And LCase$(Trim$(newRef)) <> "already void")
    Dim prev As String
    For i = 1 To UBound(v, 1)
        prev = LCase$(Trim$(SafeStr(rf(i, 1))))
        If UCase$(Trim$(SafeStr(v(i, 4)))) = oldRef And UCase$(SafeStr(v(i, 5))) <> "YES" Then
            If prev = "" Or (realOutcome And (prev = "checked" Or prev = "already void")) Then
                by(i, 1) = CurrentUser(): rf(i, 1) = Trim$(newRef)
                ResolveFeedback = ResolveFeedback + 1
            End If
        End If
    Next i
    If ResolveFeedback = 0 Then Exit Function
    Application.EnableEvents = False
    UnlockSheet ws
    lo.ListColumns(FB_RESBY).DataBodyRange.Value = by
    lo.ListColumns(FB_RESREF).DataBodyRange.Value = rf
    ProtectSheet ws
    Application.EnableEvents = evWas
    MarkCalDirty
    Exit Function
Quit:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    If en <> 0 Then LogError "ResolveFeedback", en, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    ResolveFeedback = 0
End Function

' One answer (Yes / Not sure / No) with its optional line, against the receipt on screen (its first Ref ID).
' True when the row landed. Never raises, never shows anything (the card says thanks or that it could not be kept).
Public Function SaveFeedback(ByVal answer As String, ByVal note As String) As Boolean
    ' (review L8) a copy that will not be saved: the answer would be thrown away at close - the card says so.
    ' (re-review H1) the refusal sits HERE, not in the write: SmokeTest runs with saving blocked from its first
    ' line and tests both - this refusal, and the row through WriteFeedbackRow.
    If ThisWorkbook.ReadOnly Or NoSave() Then Exit Function
    SaveFeedback = WriteFeedbackRow(answer, note)
End Function

' The row itself (the card's answer, or SmokeTest's). True when it landed. Never raises.
Public Function WriteFeedbackRow(ByVal answer As String, ByVal note As String) As Boolean
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim lo As ListObject, ws As Worksheet, cnt As Long
    On Error GoTo Quit
    Application.EnableEvents = False               ' before the sheet is made: adding it must not start a redraw
    Set lo = FeedbackTable(True)
    If lo Is Nothing Then GoTo Quit
    Set ws = lo.Parent
    cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim refTxt As String: refTxt = FirstRefIn(gReceipt.Recorded)
    If refTxt = "" Then refTxt = FirstRefIn(gReceipt.Title)
    Dim rec(1 To 1, 1 To FB_N) As Variant
    rec(1, 1) = Now
    rec(1, 2) = CurrentUser()
    rec(1, 3) = EXPECTED_RELEASE
    rec(1, 4) = IIf(refTxt = "", Left$(gReceipt.Title, 60), refTxt)
    rec(1, 5) = answer
    rec(1, 6) = "'" & Left$(Trim$(Replace(Replace(note, vbCr, " "), vbLf, " ")), FB_MAX_NOTE)   ' kept as text whatever it starts with
    rec(1, FB_SEEN) = ""
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, FB_N).Value = rec
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, 1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ProtectSheet ws
    WriteFeedbackRow = True
Quit:
    If Not WriteFeedbackRow Then
        Dim en As Long: en = Err.Number
        Dim em As String: em = Err.Description
        If en <> 0 Then LogError "WriteFeedbackRow", en, em
    End If
    On Error Resume Next
    RelockStructure                                ' (review L7) a failure while the sheet was being made
    If Not ws Is Nothing Then ProtectSheet ws       ' a failed write leaves the sheet protected
    Application.EnableEvents = evWas
    MarkCalDirty                                   ' Needs Attention's feedback line
End Function

' The answers nobody has read yet (the Seen column blank). Never raises.
Public Function UnseenFeedbackCount() As Long
    Dim lo As ListObject, v As Variant, i As Long
    On Error GoTo Quit
    Set lo = FeedbackTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, 1)) <> "" And SafeStr(v(i, FB_SEEN)) = "" Then UnseenFeedbackCount = UnseenFeedbackCount + 1
    Next i
Quit:
End Function

' Alt+F8 / the Needs Attention line (admin): shows the Feedback sheet and marks every unread answer seen
Public Sub ShowFeedback()
    If Not RequireAdmin() Then Exit Sub
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, n As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Failed
    Set lo = FeedbackTable(True)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    n = UnseenFeedbackCount()
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    If n > 0 And Not lo.DataBodyRange Is Nothing Then
        v = RangeTo2D(lo.ListColumns(FB_SEEN).DataBodyRange)
        Dim w As Variant: w = RangeTo2D(lo.ListColumns(1).DataBodyRange)
        For i = 1 To UBound(v, 1)
            If SafeStr(w(i, 1)) <> "" And SafeStr(v(i, 1)) = "" Then v(i, 1) = "Yes"
        Next i
        Application.EnableEvents = False
        UnlockSheet ws
        lo.ListColumns(FB_SEEN).DataBodyRange.Value = v
        ProtectSheet ws
        Application.EnableEvents = evWas
        MarkCalDirty
    End If
    MsgBox IIf(n = 0, "No new answers.", n & " new answer(s) to ""Was this right?"", now marked as seen.") & vbNewLine & vbNewLine & _
           "Newest is at the bottom: who answered, which Ref ID, Yes / Not sure / No and their line. A No or a Not sure is worth a look at that log entry. Prune old rows by hand (UnlockForRepair).", vbInformation, "Was this right?"
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    LogError "ShowFeedback", Err.Number, em        ' (review L10) convention 15: right after the capture
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    MsgBox "The Feedback sheet couldn't be opened: " & em, vbExclamation, "Was this right?"
End Sub

' The receipt as plain text (the print, the email, the SmokeTest log)
Public Function ReceiptText(ByRef r As Receipt) As String
    ReceiptText = r.Title & vbNewLine & _
                  Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & "  -  " & CurrentUser() & vbNewLine & vbNewLine & _
                  "WHAT WAS RECORDED" & vbNewLine & IIf(r.Recorded = "", "  (nothing written to the OT Log)" & vbNewLine, r.Recorded) & vbNewLine & _
                  "WHO IT CHANGED" & vbNewLine & IIf(r.Changed = "", "  Nobody's Total Hours changed." & vbNewLine, r.Changed) & vbNewLine & _
                  "WHERE TO SEE IT" & vbNewLine & r.WhereTxt & vbNewLine & _
                  "IF THIS IS WRONG" & vbNewLine & r.WrongTxt & vbNewLine & _
                  r.Trust
End Function

' One line per OT Log row written: "  S-0012  Refused  Joe Smith  12 h  charged 21" (data as AppendLogRows takes it)
Public Function RowsText(ByRef data As Variant, ByVal nRows As Long) As String
    Dim i As Long, s As String
    For i = 1 To nRows
        s = s & "  " & SafeStr(data(i, L_REF)) & "  " & SafeStr(data(i, L_OUTCOME)) & IIf(SafeStr(data(i, L_OUTCOME)) = "", SafeStr(data(i, L_TASK)), "") & _
            "  " & SafeStr(data(i, L_EMP)) & "  " & HoursText(NumOr0(data(i, L_BASEHRS))) & " h" & _
            IIf(NumOr0(data(i, L_CHARGED)) <> 0, "  charged " & HoursText(NumOr0(data(i, L_CHARGED))), "") & vbNewLine
    Next i
    RowsText = s & PaperLines(data, nRows)          ' v27 (item 6): the rules paper never had, each named once
End Function

' v27 (item 6, 5.3): one "On paper this would have been different" line per rule the paper roster never had that
' these rows used - read from the rows themselves: a partial acceptance that closed the shift (CH-2), an offer the
' 16 h rule capped (EL-4), a charge on a DST night at its actual hours (TH-3). Each rule named once.
Public Function PaperLines(ByRef data As Variant, ByVal nRows As Long) As String
    On Error GoTo Quit
    Dim i As Long, note As String, ids As String, d As Variant
    For i = 1 To nRows
        note = SafeStr(data(i, L_REASON))
        If Left$(note, 5) = "Took " Or InStr(1, note, "a partial acceptance closed the shift", vbTextCompare) > 0 Then
            If InStr(ids, "|CH-2|") = 0 Then ids = ids & "|CH-2|"
        End If
        If InStr(1, note, "(the 16 h rule", vbTextCompare) > 0 Or InStr(1, note, "the 16 h rule allowed", vbTextCompare) > 0 Then     ' (review) the cap's own words, not a typed reason
            If InStr(ids, "|EL-4|") = 0 Then ids = ids & "|EL-4|"
        End If
        d = data(i, L_DATE)
        If IsDate(d) And NumOr0(data(i, L_CHARGED)) > 0 And IsNightShift(SafeStr(data(i, L_SHIFTLBL))) Then
            If DstNightHours(CDate(d)) <> 12 And Abs(NumOr0(data(i, L_BASEHRS)) - DstNightHours(CDate(d))) < 0.001 Then     ' (review) charged AT the night's actual hours
                If InStr(ids, "|TH-3|") = 0 Then ids = ids & "|TH-3|"
            End If
        End If
    Next i
    Dim k As Variant
    For Each k In Split(Replace(ids, "||", "|"), "|")
        If CStr(k) <> "" Then PaperLines = PaperLines & PaperRuleLine(CStr(k))
    Next k
Quit:
End Function

' "Print receipt": the text onto the Receipt sheet, a PDF beside the roster, the sheet hidden again
Public Sub ReceiptPdf()
    On Error GoTo Oops
    Dim ws As Worksheet
    If Not SheetExists(RECEIPT_SHEET) Then
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RECEIPT_SHEET
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)
    End If
    Set ws = ThisWorkbook.Worksheets(RECEIPT_SHEET)
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    Dim lines As Variant: lines = Split(ReceiptText(gReceipt), vbNewLine)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i + 1, 1).NumberFormat = "@"
        ws.Cells(i + 1, 1).Value = CStr(lines(i))
        If i = 0 Then ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
        If CStr(lines(i)) = "WHAT WAS RECORDED" Or CStr(lines(i)) = "WHO IT CHANGED" Or CStr(lines(i)) = "WHERE TO SEE IT" Or CStr(lines(i)) = "IF THIS IS WRONG" Then ws.Cells(i + 1, 1).Font.Bold = True
    Next i
    ws.Columns(1).ColumnWidth = 110
    ws.Columns(1).WrapText = True
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(UBound(lines) + 1, 1)).Address
        .Orientation = xlPortrait: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = "Call Out Roster receipt"
    End With
    On Error GoTo Oops
    ProtectSheet ws
    Dim nm As String: nm = gReceipt.Title
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        nm = Replace(nm, CStr(bad), " ")
    Next bad
    ReportPdf ws, "Receipt " & Left$(Trim$(nm), 60) & " " & Format$(Now, "hhmm"), "Receipt", True
    SetSheetVisible ws, xlSheetVeryHidden
    Exit Sub
Oops:
    ReportError "ReceiptPdf"                     ' first: an On Error statement would clear the error it reports
    On Error Resume Next
    SetSheetVisible ThisWorkbook.Worksheets(RECEIPT_SHEET), xlSheetVeryHidden   ' (review M6) never left as a tab
End Sub

' "Email me this": a mail draft with the receipt's text (the address is typed in the mail app)
Public Sub ReceiptMail()
    On Error GoTo Oops
    If Not OpenMailTo("", "Call Out Roster receipt: " & gReceipt.Title, ReceiptText(gReceipt)) Then
        MsgBox "No mail app could be opened. Print receipt saves the same text as a PDF beside the roster.", vbExclamation, "Receipt"
    End If
    Exit Sub
Oops:
    ReportError "ReceiptMail"
End Sub
