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

Public Const EL_VERSION As String = "Error log v24.2 (Jan 5)"
Public Const EL_SHEET As String = "Error Log"
Public Const EL_TABLE As String = "tblErrorLog"
Private Const EL_N As Long = 7
Private Const EL_SEEN As Long = 7
Private Const EL_MAX_TEXT As Long = 500


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
            Dim was As Object: Set was = ActiveSheet       ' adding a sheet activates it - put the user back
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = EL_SHEET
            ws.Tab.Color = RGB(176, 0, 32)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        ws.Unprotect Password:=SHEET_PASSWORD
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
        ws.Visible = IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set ErrorLogTable = lo
End Function

' Appends one row. Returns its line number in the log (0 = it couldn't be
' written). Never raises, never shows anything.
Public Function LogError(ByVal proc As String, ByVal errNumber As Long, ByVal errText As String, _
                         Optional ByVal detail As String = "") As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim lo As ListObject, ws As Worksheet, cnt As Long
    On Error GoTo Quit
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
    rec(1, 3) = Left$(proc, 120)
    rec(1, 4) = IIf(errNumber <> 0, "#" & errNumber & "  ", "") & Left$(errText, EL_MAX_TEXT)
    rec(1, 5) = Left$(detail, EL_MAX_TEXT)
    rec(1, 6) = EXPECTED_RELEASE
    rec(1, EL_SEEN) = ""
    Application.EnableEvents = False
    ws.Unprotect Password:=SHEET_PASSWORD
    lo.Resize lo.HeaderRowRange.Resize(cnt + 2)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, EL_N).Value = rec
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(1, 1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ProtectSheet ws
    LogError = cnt + 1
Quit:
    On Error Resume Next
    Application.EnableEvents = evWas
    MarkCalDirty                                   ' Needs Attention's error line
End Function

' For an entry point's handler: logs the error that is active right now,
' gives Excel back to the user and says what to do. Call it as the FIRST
' statement after the label.
Public Sub ReportError(ByVal proc As String, Optional ByVal detail As String = "")
    Dim en As Long: en = Err.Number                ' read before anything can clear it
    Dim em As String: em = Err.Description
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Dim k As Long                                  ' a pop-up the error happened in is still loaded: close it
    For k = VBA.UserForms.Count - 1 To 0 Step -1
        Unload VBA.UserForms(k)
    Next k
    Dim lineNo As Long: lineNo = LogError(proc, en, em, detail)
    MsgBox "This didn't finish." & vbNewLine & vbNewLine & _
           "Nothing was saved unless a message before this one said so." & vbNewLine & _
           "Tell Graeme: Error Log line " & IIf(lineNo > 0, CStr(lineNo), "(not written)") & "  -  " & proc & vbNewLine & vbNewLine & _
           Left$(em, 200), vbExclamation, "Call Out Roster"
End Sub

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
    ws.Visible = xlSheetVisible
    ws.Activate
    If n > 0 And Not lo.DataBodyRange Is Nothing Then
        v = RangeTo2D(lo.ListColumns(EL_SEEN).DataBodyRange)
        Dim w As Variant: w = RangeTo2D(lo.ListColumns(1).DataBodyRange)
        For i = 1 To UBound(v, 1)
            If SafeStr(w(i, 1)) <> "" And SafeStr(v(i, 1)) = "" Then v(i, 1) = "Yes"
        Next i
        Application.EnableEvents = False
        ws.Unprotect Password:=SHEET_PASSWORD
        lo.ListColumns(EL_SEEN).DataBodyRange.Value = v
        ProtectSheet ws
        Application.EnableEvents = evWas
        MarkCalDirty
    End If
    MsgBox IIf(n = 0, "No new errors.", n & " new error line(s), now marked as seen.") & vbNewLine & vbNewLine & _
           "Newest is at the bottom. Send me the Procedure and Error text of any line you want looked at.", vbInformation, "Error Log"
    Exit Sub
Failed:
    Dim em As String: em = Err.Description
    On Error Resume Next
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
