Attribute VB_Name = "modCalImport"
Option Explicit
' v27.1 (the refactor, Oct 2026): the one-time import of an old monthly calendar workbook, from modCalendar; moved unchanged (tests/movecheck.py proves it)
Public Const CALIMP_VERSION As String = "Calendar import v27.1 (Oct 22)"

' ====================================================================
'  ONE-TIME IMPORT of an old monthly calendar workbook
' ====================================================================
Public Sub ImportUnitCalendar()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub

    ' which unit
    Dim units As Variant: units = DeptList()
    Dim msg As String, i As Long
    For i = 0 To UBound(units)
        msg = msg & (i + 1) & ")  " & units(i) & vbNewLine
    Next i
    Dim pick As Long
    pick = AskNumber("Import an old calendar workbook into which unit?" & vbNewLine & vbNewLine & msg & vbNewLine & _
                     "(For the combined Phosphine / Derivatives calendar pick either - both show on that calendar.)", _
                     "Import calendar", 1, 1, UBound(units) + 1)
    If pick < 1 Then Exit Sub
    Dim unit As String: unit = units(pick - 1)

    Dim dIn As Variant
    dIn = Application.InputBox("Import entries from this date on:", "Import calendar", _
                               DateText(DateSerial(Year(Date), Month(Date), 1)), Type:=2)
    If VarType(dIn) = vbBoolean Then Exit Sub
    Dim dFrom As Date                              ' v25: the one parser
    If Not ParseTypedDate(CStr(dIn), dFrom) Then MsgBox "Not a date (e.g. Oct 1).", vbExclamation, "Call Out Roster": Exit Sub

    Dim f As Variant
    f = Application.GetOpenFilename("Excel files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Pick the " & unit & " calendar workbook")
    If VarType(f) = vbBoolean Then Exit Sub

    Dim lo As ListObject: Set lo = CovTable(True)
    Dim src As String: src = "Import " & Format$(Date, "yyyy-mm-dd")

    ' already imported for this unit / period? (avoid doubles)
    Dim ov As Variant: ov = TableData(lo)
    Dim prior As Long, j As Long
    If IsArray(ov) Then
        For j = 1 To UBound(ov, 1)
            If Left$(SafeStr(ov(j, CV_SOURCE)), 6) = "Import" And UCase$(SafeStr(ov(j, CV_STATUS))) <> "VOID" And _
               StrComp(SafeStr(ov(j, CV_UNIT)), unit, vbTextCompare) = 0 Then
                If IsDate(ov(j, CV_DATE)) Then
                    If CDate(ov(j, CV_DATE)) >= dFrom Then prior = prior + 1
                End If
            End If
        Next j
    End If
    If prior > 0 Then
        Select Case MsgBox(prior & " imported lines already exist for " & unit & " from " & Format$(dFrom, "mmm d") & "." & vbNewLine & vbNewLine & _
                           "Yes = remove those and import again" & vbNewLine & "No = keep them and add these too" & vbNewLine & _
                           "Cancel = stop", vbYesNoCancel + vbExclamation, "Import calendar")
            Case vbCancel: Exit Sub
            Case vbYes
                UnlockSheet lo.Parent
                For j = 1 To UBound(ov, 1)
                    If Left$(SafeStr(ov(j, CV_SOURCE)), 6) = "Import" And UCase$(SafeStr(ov(j, CV_STATUS))) <> "VOID" And _
                       StrComp(SafeStr(ov(j, CV_UNIT)), unit, vbTextCompare) = 0 Then
                        If IsDate(ov(j, CV_DATE)) Then
                            If CDate(ov(j, CV_DATE)) >= dFrom Then lo.DataBodyRange.Cells(j, CV_STATUS).Value = "VOID"
                        End If
                    End If
                Next j
                ProtectSheet lo.Parent
        End Select
    End If

    Application.ScreenUpdating = False
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=f, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo Oops
    If wb Is Nothing Then Application.ScreenUpdating = True: MsgBox "Couldn't open that file.", vbCritical, "Call Out Roster": Exit Sub

    Dim buf() As Variant: ReDim buf(1 To 5000, 1 To CV_N)
    Dim n As Long, sh As Worksheet, nSheets As Long
    For Each sh In wb.Worksheets
        n = ParseMonthSheet(sh, unit, dFrom, src, buf, n)
        nSheets = nSheets + 1
    Next sh
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    If n = 0 Then MsgBox "No entries found on or after " & Format$(dFrom, "mmm d, yyyy") & ".", vbInformation, "Call Out Roster": Exit Sub
    If MsgBox("Found " & n & " calendar lines for " & unit & " from " & Format$(dFrom, "mmm d, yyyy") & "." & vbNewLine & vbNewLine & _
              "They come in as written (e.g. 'Paul M 3rd' -> 'posted') and show on the calendar." & vbNewLine & _
              "They don't remove anyone from call lists until you convert them (Convert imports button) or book the time off." & vbNewLine & vbNewLine & _
              "Import them?", vbOKCancel + vbQuestion, "Import calendar") <> vbOK Then Exit Sub

    Dim base As Long: base = CLng(Val(Mid$(NextId(lo, "CV"), 4)))
    Dim data() As Variant: ReDim data(1 To n, 1 To CV_N)
    Dim r As Long, c As Long, user As String, stamp As Date
    user = CurrentUser(): stamp = Now
    For r = 1 To n
        For c = 1 To CV_N
            data(r, c) = buf(r, c)
        Next c
        data(r, CV_ID) = "CV-" & Format$(base + r - 1, "0000")
        data(r, CV_USER) = user
        data(r, CV_STAMP) = stamp
    Next r
    If Not AppendTo(lo, data, CV_N) Then Exit Sub
    FormatCovDates lo
    SaveIfWritable
    MsgBox n & " lines imported for " & unit & ".", vbInformation, "Call Out Roster"
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "ImportUnitCalendar"
End Sub

' Reads one month tab: finds the SUNDAY..SATURDAY header, date rows, and
' the slot rows under each date (yellow = days, other fill = nights).
Private Function ParseMonthSheet(ByVal sh As Worksheet, ByVal unit As String, ByVal dFrom As Date, _
                                 ByVal src As String, ByRef buf() As Variant, ByVal n As Long) As Long
    ParseMonthSheet = n
    Dim hdrRow As Long, sunCol As Long, r As Long, c As Long
    For r = 1 To 10
        For c = 1 To 20
            If UCase$(Trim$(SafeStr(sh.Cells(r, c).Value))) = "SUNDAY" Then hdrRow = r: sunCol = c: Exit For
        Next c
        If hdrRow > 0 Then Exit For
    Next r
    If hdrRow = 0 Then Exit Function

    Dim dayCol(0 To 6) As Long, k As Long
    For k = 0 To 6
        dayCol(k) = sunCol + 2 * k
    Next k

    ' the month: a real date near the top (B1 / A1). The day cells only
    ' hold day numbers (=H4+1 ...), so dates are month + day number.
    Dim mStart As Date, found As Boolean, v As Variant
    For r = 1 To hdrRow
        For c = 1 To 8
            v = sh.Cells(r, c).Value2
            If Not IsError(v) Then
                If VarType(v) = vbDouble Then
                    If v > 40000 And v < 80000 Then mStart = DateSerial(Year(CDate(v)), Month(CDate(v)), 1): found = True
                End If
            End If
            If found Then Exit For
        Next c
        If found Then Exit For
    Next r
    If Not found Then Exit Function
    Dim mDays As Long: mDays = Day(DateSerial(Year(mStart), Month(mStart) + 1, 0))
    Dim lastRow As Long: lastRow = sh.UsedRange.Row + sh.UsedRange.Rows.Count
    If lastRow > 200 Then lastRow = 200

    ' date rows = any day column holding a date serial
    Dim isDateRow() As Boolean: ReDim isDateRow(1 To lastRow + 1)
    For r = hdrRow + 1 To lastRow
        For k = 0 To 6
            If DayNum(sh.Cells(r, dayCol(k)), mDays) > 0 Then isDateRow(r) = True: Exit For
        Next k
    Next r

    Dim cur(0 To 6) As Variant, rowD As Long, lt As String, rt As String, shiftLbl As String, colr As Long
    For r = hdrRow + 1 To lastRow
        If isDateRow(r) Then
            For k = 0 To 6
                If DayNum(sh.Cells(r, dayCol(k)), mDays) > 0 Then
                    cur(k) = DateSerial(Year(mStart), Month(mStart), DayNum(sh.Cells(r, dayCol(k)), mDays))
                Else
                    cur(k) = Empty
                End If
            Next k
        Else
            For k = 0 To 6
                If Not IsEmpty(cur(k)) Then
                    lt = Trim$(SafeStr(sh.Cells(r, dayCol(k)).Value))
                    rt = Trim$(SafeStr(sh.Cells(r, dayCol(k) + 1).Value))
                    If (lt <> "" Or rt <> "") And cur(k) >= dFrom Then
                        colr = sh.Cells(r, dayCol(k)).Interior.Color
                        If sh.Cells(r, dayCol(k)).Interior.ColorIndex = xlNone Then colr = sh.Cells(r, dayCol(k) + 1).Interior.Color
                        If colr = RGB(255, 255, 0) Then
                            shiftLbl = "Days"
                        ElseIf sh.Cells(r, dayCol(k)).Interior.ColorIndex = xlNone Or colr = RGB(255, 255, 255) Then
                            shiftLbl = "Days"
                        Else
                            shiftLbl = "Nights"
                        End If
                        n = n + 1
                        If n > UBound(buf, 1) Then n = n - 1: ParseMonthSheet = n: Exit Function
                        buf(n, CV_DATE) = Int(CDate(cur(k)))
                        buf(n, CV_SHIFT) = shiftLbl
                        buf(n, CV_UNIT) = unit
                        buf(n, CV_NEED) = IIf(lt = "", "(blank)", lt)
                        buf(n, CV_TYPE) = GuessType(lt & " " & rt)
                        If LCase$(rt) Like "posted*" Then rt = "POSTED" & Mid$(rt, 7)
                        buf(n, CV_BY) = rt
                        buf(n, CV_SOURCE) = src
                        buf(n, CV_STATUS) = "Active"
                        buf(n, CV_NOTE) = sh.Name
                    End If
                End If
            Next k
        End If
    Next r
    ParseMonthSheet = n
End Function

' Day-of-month held in a calendar date cell (1..mDays), 0 if none.
' Accepts a plain day number or a real date.
Private Function DayNum(ByVal c As Range, ByVal mDays As Long) As Long
    Dim v As Variant: v = c.Value2
    If IsError(v) Then Exit Function
    If VarType(v) <> vbDouble Then Exit Function
    If v <> Int(v) Then Exit Function
    If v > 40000 Then v = Day(CDate(v))
    If v >= 1 And v <= mDays Then DayNum = CLng(v)
End Function

Private Function GuessType(ByVal t As String) As String
    t = UCase$(t)
    If InStr(t, "S/S") > 0 Or InStr(t, " SS") > 0 Then
        GuessType = "SS"
    ElseIf InStr(t, "PHA") > 0 Then
        GuessType = "PHA"
    ElseIf InStr(t, "TRAIN") > 0 Then
        GuessType = "TRN"
    ElseIf InStr(t, "OPERATOR") > 0 Or InStr(t, " OP") > 0 Or InStr(t, "EXTRA") > 0 Or InStr(t, "FLOATER") > 0 Then
        GuessType = "SUP"
    Else
        GuessType = "VAC"
    End If
End Function

