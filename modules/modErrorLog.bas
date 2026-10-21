Attribute VB_Name = "modErrorLog"
Option Explicit

' ====================================================================
'  ERROR LOG (v24.2)
'
'  A runtime error used to leave a supervisor looking at the VBA debug
'  box. Now the entry points (button macros, the macros that open the
'  pop-ups, the workbook events) catch it:
'
'      Sub Something()
'          On Error GoTo Oops
'          ...
'          Exit Sub
'      Oops:
'          ReportError "Something"
'      End Sub
'
'  ReportError writes one row to the Error Log sheet (tblErrorLog:
'  When, User, Procedure, Error, Detail, Release, Seen) and shows a
'  plain message with the line number. LogError writes the row with no
'  message, for handlers that already tell the user what happened.
'  Needs Attention shows a line while there are rows nobody has looked
'  at; ShowErrorLog (Alt+F8, admin) opens the sheet and marks them seen.
'
'  Neither procedure can raise an error or change who is on a list.
'  The sheet is admin-only (hidden in every saved copy) and is never
'  published.
' ====================================================================

Public Const EL_VERSION As String = "Error log v26.5.5 (Oct 21)"
Public Const EL_SHEET As String = "Error Log"
Public Const EL_TABLE As String = "tblErrorLog"
Private Const EL_N As Long = 7
Private Const EL_SEEN As Long = 7
Private Const EL_MAX_TEXT As Long = 500
Private Const EL_FILE As String = "Roster error log.txt"   ' the same lines, beside the workbook (see LogError)

' True only while ReportError is closing the pop-ups an error left open: their
' "close without saving?" questions stay quiet (frmOvertimeEntry, frmPostingEntry)
Public gErrorClosing As Boolean


' The table, or Nothing. create:=True adds the sheet and the table.
Public Function ErrorLogTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(EL_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(EL_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(EL_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet     ' adding a sheet activates it - put the user back
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = EL_SHEET
            RelockStructure
            SetTabColor ws, RGB(176, 0, 32)             ' v25.1
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, EL_N).Value = Array("When", "User", "Procedure", "Error", "Detail", "Release", "Seen")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, EL_N), , xlYes)
        lo.Name = EL_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 20
        ws.Columns(2).ColumnWidth = 12
        ws.Columns(3).ColumnWidth = 30
        ws.Columns(4).ColumnWidth = 60
        ws.Columns(5).ColumnWidth = 40
        ws.Columns(6).ColumnWidth = 9
        ws.Columns(7).ColumnWidth = 7
        ws.Columns(1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set ErrorLogTable = lo
End Function

' Appends one row. Returns its line number in the log (0 = it couldn't be
' written). Never raises, never shows anything.
' The same line also goes into a small text file beside the workbook: several
' handlers block saving after they log ("close WITHOUT saving"), and the row on
' the sheet goes with the unsaved workbook. The text file keeps it.
Public Function LogError(ByVal proc As String, ByVal errNumber As Long, ByVal errText As String, _
                         Optional ByVal detail As String = "") As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim lo As ListObject, ws As Worksheet, cnt As Long
    On Error GoTo Quit
    LogErrorToFile proc, errNumber, errText, detail
    Application.EnableEvents = False               ' before the sheet is made: adding it must not start a redraw
    Set lo = ErrorLogTable(True)
    If lo Is Nothing Then GoTo Quit
    Set ws = lo.Parent
    cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim rec(1 To 1, 1 To EL_N) As Variant
    rec(1, 1) = Now
    rec(1, 2) = CurrentUser()
    rec(1, 3) = AsCellText(Left$(proc, 120))
    rec(1, 4) = AsCellText(IIf(errNumber <> 0, "#" & errNumber & "  ", "") & Left$(errText, EL_MAX_TEXT))
    rec(1, 5) = AsCellText(Left$(detail, EL_MAX_TEXT))
    rec(1, 6) = EXPECTED_RELEASE
    rec(1, EL_SEEN) = ""
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, EL_N).Value = rec
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, 1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ProtectSheet ws
    LogError = cnt + 1
Quit:
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws       ' v25 review: a failed write leaves the Error Log protected
    Application.EnableEvents = evWas
    MarkCalDirty                                   ' Needs Attention's error line
End Function

' Text that starts like a formula ("=", "+", "-", "@") gets a leading apostrophe, so Excel keeps it as text
Private Function AsCellText(ByVal s As String) As String
    If s <> "" Then
        If InStr("=+-@", Left$(s, 1)) > 0 Then s = "'" & s
    End If
    AsCellText = s
End Function

' One line appended to "Roster error log.txt" beside the workbook. Quiet on any failure.
Private Sub LogErrorToFile(ByVal proc As String, ByVal errNumber As Long, ByVal errText As String, ByVal detail As String)
    Dim ff As Integer, p As String
    On Error GoTo Quit
    p = ThisWorkbook.Path
    If p = "" Or LCase$(Left$(p, 4)) = "http" Then Exit Sub
    ff = FreeFile
    Open EnsureSlash(p) & EL_FILE For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & EXPECTED_RELEASE & vbTab & CurrentUser() & vbTab & proc & vbTab & _
               IIf(errNumber <> 0, "#" & errNumber & " ", "") & Replace(Replace(Left$(errText, EL_MAX_TEXT), vbCr, " "), vbLf, " ") & vbTab & _
               Replace(Replace(Left$(detail, EL_MAX_TEXT), vbCr, " "), vbLf, " ")
    Close #ff
    Exit Sub
Quit:
    On Error Resume Next
    Close #ff
End Sub

' For an entry point's handler: logs the error that is active right now,
' gives Excel back to the user and says what to do. Call it as the FIRST
' statement after the label.
Public Sub ReportError(ByVal proc As String, Optional ByVal detail As String = "")
    Dim en As Long: en = Err.Number                ' read before anything can clear it
    Dim em As String: em = Err.Description
    On Error Resume Next
    If en = 18 Then em = "Stopped by Esc / Ctrl+Break: " & em      ' v26.5.5 (A3): a cancelled write is an error like any other
    Application.EnableCancelKey = xlInterrupt       ' v26.5.5 (A3): a writer that made Esc trappable hands the key back here
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    ' v24.3.1: everything a bulk routine turns off (convention 6). Automatic is the roster's only
    ' normal state; a routine that captured and restores its own mode never reaches this line.
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.DisplayAlerts = True
    ProtectAll                                     ' v25 review: EVERY sheet (a table write that raised between UnlockSheet and
                                                   ' ProtectSheet) and the tabs; ProtectSheet keeps an admin's unlocked cells
    Dim k As Long                                  ' a pop-up the error happened in is still loaded: close it
    gErrorClosing = True                           ' ...without its "close without saving?" question
    For k = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(k)
    Next k
    gErrorClosing = False
    gCalPostKey = ""                               ' a calendar line picked for a posting that never got made
    Dim lineNo As Long: lineNo = LogError(proc, en, em, detail)
    MsgBox "This didn't finish." & vbNewLine & vbNewLine & _
           "Anything it did before the error is still in the workbook" & _
           IIf(NoSave() Or ThisWorkbook.ReadOnly, " (this copy can't be saved).", " and is saved when you close.") & vbNewLine & _
           "Check what you were doing (run it again if it was an import), and tell " & PRIMARY_ADMIN_NAME & ":" & vbNewLine & _
           "Error Log line " & IIf(lineNo > 0, CStr(lineNo), "(not written)") & "  -  " & proc & vbNewLine & vbNewLine & _
           Left$(em, 200), vbExclamation, "Call Out Roster"
End Sub

' v25.3 (the Desk's "Something went wrong"): the newest row as one line of text (rows are appended)
Public Function NewestErrorLine() As String
    On Error Resume Next
    NewestErrorLine = "(the Error Log is empty)"
    Dim lo As ListObject: Set lo = ErrorLogTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = UBound(v, 1) To 1 Step -1
        If SafeStr(v(i, 1)) <> "" Then
            NewestErrorLine = "Line " & i & ": " & Format$(v(i, 1), "mmm d, yyyy h:mm AM/PM") & " - " & SafeStr(v(i, 3)) & _
                              " - " & SafeStr(v(i, 4)) & IIf(SafeStr(v(i, 5)) <> "", " - " & Left$(SafeStr(v(i, 5)), 200), "")
            Exit Function
        End If
    Next i
End Function

' Rows nobody has looked at yet
Public Function UnseenErrorCount() As Long
    Dim lo As ListObject, v As Variant, i As Long
    On Error GoTo Quit
    Set lo = ErrorLogTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, 1)) <> "" And SafeStr(v(i, EL_SEEN)) = "" Then UnseenErrorCount = UnseenErrorCount + 1
    Next i
Quit:
End Function

' Alt+F8 (admin): opens the sheet and marks every row seen
Public Sub ShowErrorLog()
    If Not RequireAdmin() Then Exit Sub
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, n As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Failed
    Set lo = ErrorLogTable(True)
    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent
    n = UnseenErrorCount()
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    If n > 0 And Not lo.DataBodyRange Is Nothing Then
        v = RangeTo2D(lo.ListColumns(EL_SEEN).DataBodyRange)
        Dim w As Variant: w = RangeTo2D(lo.ListColumns(1).DataBodyRange)
        For i = 1 To UBound(v, 1)
            If SafeStr(w(i, 1)) <> "" And SafeStr(v(i, 1)) = "" Then v(i, 1) = "Yes"
        Next i
        Application.EnableEvents = False
        UnlockSheet ws
        lo.ListColumns(EL_SEEN).DataBodyRange.Value = v
        ProtectSheet ws
        Application.EnableEvents = evWas
        MarkCalDirty
    End If
    MsgBox IIf(n = 0, "No new errors.", n & " new error line(s), now marked as seen.") & vbNewLine & vbNewLine & _
           "Newest is at the bottom. Send " & PRIMARY_ADMIN_NAME & " the Procedure and Error text of any line you want looked at.", vbInformation, "Error Log"
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws       ' v25 review
    Application.EnableEvents = evWas
    MsgBox "The Error Log couldn't be opened: " & em, vbExclamation, "Error Log"
End Sub

' Alt+F8 (admin): raises an error on purpose, to see the log working after an install
Public Sub ErrorLogTest()
    If Not RequireAdmin() Then Exit Sub
    On Error GoTo Oops
    Err.Raise vbObjectError + 999, , "a test error, raised on purpose by ErrorLogTest"
    Exit Sub
Oops:
    ReportError "ErrorLogTest", "nothing is wrong"
End Sub
