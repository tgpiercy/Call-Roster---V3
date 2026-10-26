Attribute VB_Name = "modLogLayout"
Option Explicit

' ====================================================================
'  OT LOG LAYOUT
'  UpgradeOTLogLayout - one time: re-orders the OT Log to match the
'     Postings words and adds the Shift (Days / Nights) column:
'       Ref ID | Type | Shift Date | Unit | Shift | Crew | Reason | Hours
'       | Call Order | Employee | Dept Code | Total Hrs Then | Outcome
'       | Note | Hours Charged | Period Start
'       | Entry ID | Logged By | Timestamp | Status | Void Note  (audit, collapsible)
'     Columns are moved / renamed in place, so the roster formulas keep
'     pointing at the same data. Roster totals are checked before/after.
'  FormatOTLog - widths, formats, colours, event separators, audit group.
' ====================================================================

Public Sub UpgradeOTLogLayout()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not SheetExists(LOG_SHEET) Then MsgBox "Sheet '" & LOG_SHEET & "' not found.", vbCritical, "Call Out Roster": Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(LOG_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then MsgBox "Table " & LOG_TABLE & " not found.", vbCritical, "Call Out Roster": Exit Sub

    If LogLayoutCurrent(lo) Then
        FormatOTLog
        MsgBox "The OT Log is already in the new layout - formatting refreshed.", vbInformation, "Call Out Roster"
        Exit Sub
    End If

    ' ---- expect the old 20-column layout exactly ----
    Dim oldH As Variant
    oldH = Array("Entry ID", "Ref ID", "Type", "Shift Date", "Department", "Shift", "Task", _
                 "Base Hours", "Call Order", "Employee", "Dept Code", "Employee Total Hrs", _
                 "Outcome", "Reason / Note", "Hours Charged", "Period Start", "Logged By", _
                 "Timestamp", "Status", "Void Note")
    Dim i As Long
    If lo.ListColumns.Count <> 20 Then
        MsgBox "The OT Log has " & lo.ListColumns.Count & " columns - expected 20. Nothing was changed.", vbCritical, "Call Out Roster"
        Exit Sub
    End If
    For i = 0 To 19
        If lo.ListColumns(i + 1).Name <> oldH(i) Then
            MsgBox "OT Log column " & (i + 1) & " is '" & lo.ListColumns(i + 1).Name & "', expected '" & _
                   oldH(i) & "'. Nothing was changed.", vbCritical, "Call Out Roster"
            Exit Sub
        End If
    Next i

    If MsgBox("Upgrade the OT Log layout?" & vbNewLine & vbNewLine & _
              "- columns re-ordered to match the Postings sheet" & vbNewLine & _
              "- new Shift (Days / Nights) column; Shift becomes Crew" & vbNewLine & _
              "- audit columns (Entry ID, Logged By, ...) moved to the end and grouped" & vbNewLine & vbNewLine & _
              "Roster totals are checked before and after. Save a backup copy first.", _
              vbOKCancel + vbQuestion, "Upgrade OT Log") <> vbOK Then Exit Sub

    ' ---- before: roster pay-period values + log totals ----
    Application.Calculate
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim before As Variant
    If lastRow >= R_FIRST Then before = rs.Range(rs.Cells(R_FIRST, C_PERIOD1), rs.Cells(lastRow, C_PERIODX)).Value2
    Dim nRows As Long, sumCharged As Double
    If Not lo.DataBodyRange Is Nothing Then
        nRows = lo.ListRows.Count
        sumCharged = Application.WorksheetFunction.Sum(lo.ListColumns("Hours Charged").DataBodyRange)
    End If

    Dim calcMode As XlCalculation: calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Dim changed As Boolean                  ' False until the first column is touched
    On Error GoTo Failed
    UnlockSheet ws
    On Error Resume Next
    ws.Cells.ClearOutline                   ' errors if there's no outline - that's fine
    Err.Clear
    On Error GoTo Failed

    ' ---- rename + insert (old col numbers) ----
    changed = True
    lo.ListColumns("Shift").Name = "Crew"
    lo.ListColumns("Department").Name = "Unit"
    lo.ListColumns("Task").Name = "Reason"
    lo.ListColumns("Base Hours").Name = "Hours"
    lo.ListColumns("Employee Total Hrs").Name = "Total Hrs Then"
    lo.ListColumns("Reason / Note").Name = "Note"
    lo.ListColumns.Add(Position:=6).Name = "Shift"            ' between Unit and Crew

    ' Entry ID: copy to the audit block, then remove the old first column
    Dim newCol As ListColumn
    Set newCol = lo.ListColumns.Add(Position:=lo.ListColumns("Logged By").Index)
    newCol.Name = "EntryID_new"
    If Not lo.DataBodyRange Is Nothing Then
        newCol.DataBodyRange.Value = lo.ListColumns("Entry ID").DataBodyRange.Value
        newCol.DataBodyRange.NumberFormat = "0"
    End If
    lo.ListColumns("Entry ID").Delete
    lo.ListColumns("EntryID_new").Name = "Entry ID"

    If lo.ListColumns.Count = L_NCOLS - 1 Then lo.ListColumns.Add.Name = KEY_HEADER     ' v28: the key on the end

    ' ---- final check of the column order ----
    Dim newH As Variant: newH = LogHeaders()
    For i = 0 To L_NCOLS - 1
        If lo.ListColumns(i + 1).Name <> newH(i) Then
            Err.Raise vbObjectError + 1, , "column " & (i + 1) & " is '" & lo.ListColumns(i + 1).Name & _
                                           "', expected '" & newH(i) & "'"
        End If
    Next i

    ' ---- Shift (Days / Nights) for existing posting rows ----
    If Not lo.DataBodyRange Is Nothing Then
        Dim map As Object: Set map = PostingShiftMap()
        Dim v As Variant, sh As Variant, nFill As Long
        v = RangeTo2D(lo.DataBodyRange)
        ReDim sh(1 To UBound(v, 1), 1 To 1)
        For i = 1 To UBound(v, 1)
            sh(i, 1) = Empty
            If SafeStr(v(i, L_TYPE)) = "Posting" Then
                If map.Exists(Trim$(SafeStr(v(i, L_REF)))) Then
                    sh(i, 1) = map(Trim$(SafeStr(v(i, L_REF))))
                    nFill = nFill + 1
                End If
            End If
        Next i
        lo.ListColumns(L_SHIFTLBL).DataBodyRange.Value = sh
    End If

    ' ---- after: same totals? ----
    Application.Calculation = calcMode
    Application.CalculateFull
    Dim problems As String
    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.Count <> nRows Then problems = problems & "- log rows: " & nRows & " before, " & lo.ListRows.Count & " after" & vbNewLine
        If Abs(Application.WorksheetFunction.Sum(lo.ListColumns(L_CHARGED).DataBodyRange) - sumCharged) > 0.001 Then
            problems = problems & "- total Hours Charged changed" & vbNewLine
        End If
    End If
    If problems = "" And lastRow >= R_FIRST Then
        If RosterDiffers(rs, lastRow, before) Then
            ' re-point the roster formulas by name and check again
            Dim r As Long
            UnlockSheet rs
            For r = R_FIRST To lastRow
                If Trim$(SafeStr(rs.Cells(r, C_NAME).Value)) <> "" Then WritePeriodRow rs, r
            Next r
            ProtectSheet rs
            Application.CalculateFull
            If RosterDiffers(rs, lastRow, before) Then problems = problems & "- roster pay-period totals changed" & vbNewLine
        End If
    End If

    If problems <> "" Then
        On Error Resume Next
        ProtectSheet ws
        Application.ScreenUpdating = True
        MarkNoSave
        MsgBox "The upgrade did NOT verify:" & vbNewLine & problems & vbNewLine & _
               "Saving is blocked. Close WITHOUT saving and send " & PRIMARY_ADMIN_NAME & " a screenshot.", vbCritical, "Call Out Roster"
        Exit Sub
    End If

    On Error GoTo 0
    FormatOTLog
    Application.ScreenUpdating = True
    SetRosterStatus
    SaveIfWritable
    MsgBox "OT Log upgraded and verified." & vbNewLine & vbNewLine & _
           "- " & nRows & " rows, roster totals unchanged" & vbNewLine & _
           "- Shift (Days / Nights) filled for " & nFill & " posting row(s); older call-outs stay blank" & vbNewLine & _
           "- audit columns are grouped: click + above the columns to show them", vbInformation, "Call Out Roster"
    Exit Sub

Failed:
    Dim em As String: em = Err.Description
    LogError "UpgradeOTLogLayout", Err.Number, em
    On Error Resume Next
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    If Not changed Then
        ProtectSheet ws
        MsgBox "The upgrade couldn't start: " & em & vbNewLine & vbNewLine & _
               "Nothing was changed - the OT Log is as it was.", vbExclamation, "Call Out Roster"
        Exit Sub
    End If
    MarkNoSave
    MsgBox "The upgrade stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving and reopen your file.", vbCritical, "Call Out Roster"
End Sub

Private Function RosterDiffers(ByVal rs As Worksheet, ByVal lastRow As Long, ByVal before As Variant) As Boolean
    Dim after As Variant, i As Long, j As Long
    after = rs.Range(rs.Cells(R_FIRST, C_PERIOD1), rs.Cells(lastRow, C_PERIODX)).Value2
    For i = 1 To UBound(before, 1)
        For j = 1 To UBound(before, 2)
            If IsError(after(i, j)) Then RosterDiffers = True: Exit Function
            If IsError(before(i, j)) Then
                ' was already an error - ignore
            ElseIf Abs(NumOr0(after(i, j)) - NumOr0(before(i, j))) > 0.001 Then
                RosterDiffers = True: Exit Function
            End If
        Next j
    Next i
End Function


' --------------------------------------------------------------------
'  FORMAT (safe to run any time)
' --------------------------------------------------------------------
Public Sub FormatOTLog()
    If Not SheetExists(LOG_SHEET) Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(LOG_TABLE)
    On Error GoTo 0
    If Not LogLayoutCurrent(lo) Then Exit Sub

    Dim wasActive As Worksheet: Set wasActive = ActiveSheet
    Application.ScreenUpdating = False
    UnlockSheet ws

    Dim lastCol As String: lastCol = ColLetter(L_NCOLS)
    Dim widths As Variant, i As Long
    widths = Array(10, 10, 11, 14, 8, 7, 18, 7, 7, 24, 9, 10, 14, 36, 10, 11, 8, 11, 18, 8, 30, 12)     ' v28: + Payroll #
    For i = 1 To L_NCOLS
        ws.Columns(i).ColumnWidth = widths(i - 1)
    Next i
    ws.Columns(L_DATE).NumberFormat = "ddd mmm d, yyyy"
    ws.Columns(L_PERIOD).NumberFormat = "mmm d, yyyy"
    ws.Columns(L_STAMP).NumberFormat = "mmm d, yyyy h:mm AM/PM"
    ws.Columns(L_ENTRY).NumberFormat = "0"
    ws.Columns(L_CHARGED).NumberFormat = "General;-General;"
    ws.Columns(L_EMPHRS).NumberFormat = "#,##0"
    Dim c As Variant
    For Each c In Array(L_SHIFTLBL, L_CREW, L_BASEHRS, L_ORDER, L_CODE, L_STATUS)
        ws.Columns(c).HorizontalAlignment = xlCenter
    Next c
    lo.HeaderRowRange.WrapText = True
    lo.HeaderRowRange.VerticalAlignment = xlCenter
    ws.Rows(1).RowHeight = 30

    ' ---- conditional formats ----
    ws.Cells.FormatConditions.Delete
    Dim body As Range: Set body = ws.Range("A2:" & lastCol & "50000")
    Dim oc As String: oc = ColLetter(L_OUTCOME)
    Dim outRng As Range: Set outRng = ws.Range(oc & "2:" & oc & "50000")
    AddValueCF outRng, "Accepted", RGB(198, 239, 206), RGB(0, 97, 0)
    AddValueCF outRng, "Awarded", RGB(198, 239, 206), RGB(0, 97, 0)
    AddValueCF outRng, "Refused", RGB(255, 199, 206), RGB(156, 0, 6)
    AddValueCF outRng, "Left Message", RGB(255, 235, 156), RGB(156, 101, 0)
    AddValueCF outRng, "Signed Up", RGB(255, 235, 156), RGB(156, 101, 0)
    AddValueCF outRng, "Unavailable", RGB(217, 225, 242), RGB(50, 65, 90)
    AddValueCF outRng, "Exception", RGB(230, 224, 240), RGB(80, 50, 120)      ' v27 (item 24, review L9)

    Dim fc As FormatCondition
    ' CF formulas are relative to the top-left cell of the range (A2)
    Set fc = body.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=$" & ColLetter(L_STATUS) & "2=""VOID""")
    fc.Font.Strikethrough = True
    fc.Font.Color = RGB(150, 150, 150)
    ' a line between events (rows with a different Ref ID)
    Set fc = body.FormatConditions.Add(Type:=xlExpression, _
                Formula1:="=AND($" & ColLetter(L_REF) & "2<>"""",$" & ColLetter(L_REF) & "2<>$" & ColLetter(L_REF) & "1)")
    With fc.Borders(xlTop)
        .LineStyle = xlContinuous
        .Color = RGB(30, 56, 100)
    End With

    ' ---- audit columns grouped (collapsed) ----
    On Error Resume Next
    ws.Cells.ClearOutline
    ws.Range(ws.Cells(1, L_ENTRY), ws.Cells(1, L_VOIDNOTE)).EntireColumn.Group
    ws.Outline.SummaryColumn = xlSummaryOnLeft
    ws.Outline.ShowLevels ColumnLevels:=1
    On Error GoTo 0

    ' ---- freeze header + Ref ID ----
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False

    SortLogNewestFirst lo
    ProtectSheet ws
    If Not wasActive Is Nothing Then wasActive.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub AddValueCF(ByVal rng As Range, ByVal txt As String, ByVal fillCol As Long, ByVal fontCol As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & txt & """")
    fc.Interior.Color = fillCol
    fc.Font.Color = fontCol
End Sub
