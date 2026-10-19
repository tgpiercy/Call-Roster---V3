Attribute VB_Name = "ImportKronosData"
Option Explicit

' ====================================================================
'  KRONOS -> Worked Hours (column H)
'  Reads "Sorted Data Here" (name in B, grand total in M) and fills H.
'  Names match ignoring spaces/case; a Kronos name with a trailing
'  middle initial (e.g. "Dobbin,Stephen A") also matches.
'  Matched names turn green; unmatched names are listed at the end.
'
'  v20 (Oct 9):
'   - Remembers when the import ran and the Kronos file's saved date, for
'     the status line and the stale-data warning (KronosStatusText).
'   - Every import adds a dated snapshot of each person's worked hours to
'     the hidden Hours History sheet (needed later for the leave return
'     top-up and provisional accepted OT).
'   - Also lists Kronos names that aren't on the roster.
'  v21 (Oct 11):
'   - Each snapshot row records the "Data Through" date: the Kronos data
'     runs through the day BEFORE the file was saved. The latest value is
'     kept as the hidden name KronosDataThrough, which the roster's
'     provisional-OT formulas (column C) read: accepted / awarded shifts
'     that have ended by that date are in Kronos, so they stop counting.
' ====================================================================

Public Const KR_VERSION As String = "Kronos import v26 (Oct 18)"
Public Const HH_SHEET As String = "Hours History"
Public Const HH_TABLE As String = "tblHoursHistory"
Private Const HH_NCOLS As Long = 8          ' v21: + Data Through; v23: + Year (the roster year the hours belong to)

' v25: the by-heading rule on its own (D-I4): "" when 'Employee Name' and 'Grand Total' are both found
' on one heading row within the first 10 rows of src (a 2-D array of the sheet), else what is missing.
Public Function KronosHeadingProblem(ByRef src As Variant, ByRef nameCol As Long, ByRef totCol As Long, _
                                     ByRef hdrName As Long, ByRef hdrTot As Long) As String
    Dim missingHdr As String
    nameCol = HeaderColumn(src, hdrName, "Employee Name")
    totCol = HeaderColumn(src, hdrTot, "Grand Total")
    If nameCol = 0 Then missingHdr = "'Employee Name'"
    If totCol = 0 Then missingHdr = missingHdr & IIf(missingHdr = "", "", " and ") & "'Grand Total'"
    If missingHdr = "" And hdrName <> hdrTot Then missingHdr = "'Employee Name' and 'Grand Total' on the same heading row"
    KronosHeadingProblem = missingHdr
End Function

Sub SyncKronosData()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub

    Dim wsDest As Worksheet: Set wsDest = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub

    ' ---- open (or reuse) the Kronos workbook ----
    Dim fileName As String: fileName = Mid$(KRONOS_PATH, InStrRev(KRONOS_PATH, "\") + 1)
    Dim wbSource As Workbook, openedHere As Boolean
    On Error Resume Next
    Set wbSource = Workbooks(fileName)
    On Error GoTo Oops
    If wbSource Is Nothing Then
        Application.EnableEvents = False          ' don't run the Kronos file's own macros
        On Error Resume Next
        Set wbSource = Workbooks.Open(Filename:=KRONOS_PATH, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo Oops
        Application.EnableEvents = True
        If wbSource Is Nothing Then
            MsgBox "Could not open the Kronos file:" & vbNewLine & KRONOS_PATH, vbCritical
            Exit Sub
        End If
        openedHere = True
    End If

    Dim wsSource As Worksheet
    On Error Resume Next
    Set wsSource = wbSource.Worksheets(KRONOS_SHEET)
    On Error GoTo Oops
    If wsSource Is Nothing Then
        If openedHere Then wbSource.Close SaveChanges:=False
        MsgBox "Sheet '" & KRONOS_SHEET & "' not found in " & fileName, vbCritical
        Exit Sub
    End If

    ' ---- Kronos names -> grand total ----
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim rawOf As Object: Set rawOf = CreateObject("Scripting.Dictionary")   ' key -> name as Kronos spells it
    Dim used As Object: Set used = CreateObject("Scripting.Dictionary")    ' keys matched to a roster name
    Dim alt As Object: Set alt = CreateObject("Scripting.Dictionary")   ' "Last,First X" -> "last,first"
    ' v24.3.1 (review I3, decided D-I4): the sheet is read whole, and the name and total columns are
    ' found by their headings, never by position. Before this, column 2 and column 13 were read for
    ' every row: a report with one pay-code column more or fewer would have put a pay-code column
    ' into everyone's Worked hours (or read no Grand Total at all) and still said "complete".
    Dim lastSrc As Long, lastCol As Long, lastCell As Range
    ' the last cell that holds anything, by rows and by columns (UsedRange can run to the sheet's end
    ' when a column carries a whole-column format); nothing found = an empty sheet, read 2 x 2
    ' (LookIn:=xlFormulas, not xlValues: a hidden or filtered-out last row or column still counts)
    Set lastCell = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastSrc = 2 Else lastSrc = lastCell.Row
    Set lastCell = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastCol = 2 Else lastCol = lastCell.Column
    If lastSrc < 2 Then lastSrc = 2                   ' always a 2-D array below
    If lastCol < 2 Then lastCol = 2
    Dim src As Variant: src = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastSrc, lastCol)).Value
    Dim i As Long, k As String, raw As String, a As String
    Dim nameCol As Long, totCol As Long, hdrName As Long, hdrTot As Long, missingHdr As String
    missingHdr = KronosHeadingProblem(src, nameCol, totCol, hdrName, hdrTot)     ' v25: the rule on its own (SmokeTest runs it)
    If missingHdr <> "" Then
        If openedHere Then wbSource.Close SaveChanges:=False
        ThisWorkbook.Activate
        MsgBox "The Kronos sheet doesn't have the heading(s) the import needs: " & missingHdr & "." & vbNewLine & _
               "Nothing was changed." & vbNewLine & vbNewLine & _
               "Open " & fileName & ", sheet '" & KRONOS_SHEET & "', and check its heading row (it must be within the first 10 rows). " & _
               "If the report's layout has changed, send me a screenshot of its first rows.", vbExclamation, "Kronos import"
        Exit Sub
    End If
    ' v24.2: the Payroll # column, found by its header (not there = the import works as before)
    Dim payOf As Object: Set payOf = CreateObject("Scripting.Dictionary")        ' name key -> Payroll #
    Dim keyOfNum As Object: Set keyOfNum = CreateObject("Scripting.Dictionary")  ' Payroll # -> name key ("" = two people share it)
    Dim pCol As Long, hdrRow As Long, pn As String
    hdrRow = hdrName
    pCol = PayrollColumn(src, hdrRow)                 ' v24.3.1: on the heading row only (a title cell above it isn't the column)
    For i = hdrRow + 1 To UBound(src, 1)
        raw = Trim$(SafeStr(src(i, nameCol)))
        k = NameKey(raw)
        If k <> "" Then
            If pCol > 0 And Not dict.Exists(k) Then
                pn = PayrollText(src(i, pCol))
                If pn <> "" Then
                    payOf(k) = pn
                    If keyOfNum.Exists(pn) Then keyOfNum(pn) = "" Else keyOfNum.Add pn, k
                End If
            End If
            If Not dict.Exists(k) Then dict.Add k, src(i, totCol): rawOf.Add k, raw
            If raw Like "* [A-Za-z]" Or raw Like "* [A-Za-z]." Then
                a = NameKey(Left$(raw, InStrRev(raw, " ") - 1))
                If alt.Exists(a) Then
                    alt(a) = ""                 ' two people share it - don't guess
                Else
                    alt.Add a, k
                End If
            End If
        End If
    Next i
    Dim fileDate As Date
    On Error Resume Next
    fileDate = FileDateTime(KRONOS_PATH)
    If fileDate = 0 Then fileDate = FileDateTime(wbSource.FullName)
    On Error GoTo Oops
    If openedHere Then wbSource.Close SaveChanges:=False
    ThisWorkbook.Activate

    ' ---- v23: after a rollover, Worked (H) stays 0 until the export covers the new year ----
    Dim awaitFrom As Date: awaitFrom = KronosAwaitFrom()
    If awaitFrom > 0 Then
        Dim fdChk As Date: fdChk = fileDate
        If fdChk = 0 Then fdChk = Now
        If DataThroughFor(fdChk) < awaitFrom Then
            MsgBox "This Kronos file was saved " & Format$(fdChk, "ddd mmm d, yyyy") & ", so it still holds last year's hours." & vbNewLine & vbNewLine & _
                   "Worked hours stay at 0 until you import an export saved after " & Format$(awaitFrom, "mmm d, yyyy") & _
                   " (the first one of the new year). Nothing was changed.", vbExclamation, "Kronos import"
            Exit Sub
        End If
    End If

    ' ---- fill column H ----
    Application.ScreenUpdating = False
    UnlockSheet wsDest
    wsDest.Range(wsDest.Cells(R_FIRST, C_NAME), wsDest.Cells(lastRow, C_NAME)).Interior.ColorIndex = xlNone

    Dim r As Long, nm As String, hit As String, found As Long, missing As String, kVal As Variant
    Dim rosterNum As String, numKey As String, nameClash As String, nFilled As Long, noNum As String, nNoNum As Long, payDiff As String, payShared As String
    If pCol > 0 Then                               ' the column keeps text, so a leading zero survives
        wsDest.Range(wsDest.Cells(R_FIRST, C_PAYROLL), wsDest.Cells(lastRow, C_PAYROLL)).NumberFormat = "@"
        If Trim$(SafeStr(wsDest.Cells(2, C_PAYROLL).Value)) = "" Then
            With wsDest.Cells(2, C_PAYROLL)
                .Value = "PAYROLL #"
                .Interior.Color = RGB(31, 56, 100): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 8
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            End With
        End If
    End If
    Dim histN As Long, histName() As String, histHrs() As Double
    ReDim histName(1 To lastRow - R_FIRST + 1): ReDim histHrs(1 To lastRow - R_FIRST + 1)
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(wsDest.Cells(r, C_NAME).Value))
        If nm <> "" Then
            hit = FindKey(dict, alt, NameKey(nm))
            ' v24.2: a row that has its Payroll # is matched by the number first
            rosterNum = "": numKey = ""
            If pCol > 0 Then
                rosterNum = PayrollText(wsDest.Cells(r, C_PAYROLL).Value)
                If rosterNum <> "" Then
                    If keyOfNum.Exists(rosterNum) Then numKey = CStr(keyOfNum(rosterNum))
                End If
            End If
            Select Case PayrollDecision(rosterNum, numKey <> "", (numKey <> "" And numKey = hit))
                Case "number"
                    hit = numKey
                Case "mismatch"                    ' same number, another name: a rename in Kronos, or a typing slip
                    used(numKey) = True
                    nameClash = nameClash & "  " & nm & "   (Kronos has that Payroll # as " & rawOf(numKey) & ")" & vbNewLine
                    hit = "~leave~"
            End Select
            If hit = "~leave~" Then
                ' nothing is changed for this row
            ElseIf hit <> "" Then
                If pCol > 0 And payOf.Exists(hit) Then
                    If rosterNum = "" Then
                        If CStr(keyOfNum(CStr(payOf(hit)))) = "" Then
                            payShared = payShared & "  " & nm & vbNewLine          ' Kronos has that number under two names: don't guess
                        Else
                            wsDest.Cells(r, C_PAYROLL).Value = CStr(payOf(hit))
                            nFilled = nFilled + 1
                        End If
                    ElseIf rosterNum <> CStr(payOf(hit)) Then
                        payDiff = payDiff & "  " & nm & vbNewLine                  ' matched by name; the roster's number isn't Kronos's
                    End If
                End If
                used(hit) = True
                kVal = dict(hit)
                If Not IsError(kVal) And IsNumeric(kVal) And Not IsEmpty(kVal) Then
                    wsDest.Cells(r, C_WORKED).Value = CDbl(kVal)
                    wsDest.Cells(r, C_NAME).Interior.Color = RGB(226, 239, 218)
                    found = found + 1
                    histN = histN + 1: histName(histN) = nm: histHrs(histN) = CDbl(kVal)
                Else
                    missing = missing & "  " & nm & " (no number in Kronos)" & vbNewLine
                End If
            Else
                missing = missing & "  " & nm & vbNewLine
            End If
        End If
    Next r
    If pCol > 0 Then                               ' who still has no Payroll #
        For r = R_FIRST To lastRow
            If Trim$(SafeStr(wsDest.Cells(r, C_NAME).Value)) <> "" And PayrollText(wsDest.Cells(r, C_PAYROLL).Value) = "" Then
                nNoNum = nNoNum + 1
                If nNoNum <= 10 Then noNum = noNum & "  " & Trim$(SafeStr(wsDest.Cells(r, C_NAME).Value)) & vbNewLine
            End If
        Next r
        If nNoNum > 10 Then noNum = noNum & "  ... and " & (nNoNum - 10) & " more" & vbNewLine
    End If
    wsDest.Rows(R_FIRST & ":" & lastRow).AutoFit
    ProtectSheet wsDest

    ' ---- v20: remember when, and keep a dated snapshot ----
    Dim stamp As Date: stamp = Now
    If fileDate = 0 Then fileDate = stamp
    SetDateName "KronosImportedAt", stamp
    SetDateName "KronosFileDate", fileDate
    ' v21: the data covers everything through the day before the file was saved
    Dim through As Date: through = DataThroughFor(fileDate)
    SetDateName PROV_NAME, through
    If awaitFrom > 0 Then ClearKronosAwait      ' v23: the new year's hours are in
    Dim histId As String
    If histN > 0 Then histId = AppendHoursHistory(stamp, fileDate, through, histName, histHrs, histN)
    Application.CalculateFull                 ' provisional (C) and Total (G) pick up the new date
    Application.ScreenUpdating = True
    wsDest.Activate
    On Error Resume Next
    SetRosterStatus
    On Error GoTo Oops

    ' Kronos names that matched nobody on the roster
    Dim extra As String, nExtra As Long, kk As Variant
    Dim formerKeys As Object: Set formerKeys = CreateObject("Scripting.Dictionary")     ' v24: people on the Former sheet are known
    On Error Resume Next
    Dim fk As String
    For Each kk In FormerNameSet().Keys
        fk = FindKey(dict, alt, NameKey(CStr(kk)))         ' the same match a roster name gets (middle initial in Kronos)
        If fk <> "" Then formerKeys(fk) = True
    Next kk
    On Error GoTo Oops
    For Each kk In dict.Keys
        If Not used.Exists(kk) And Not formerKeys.Exists(kk) Then
            nExtra = nExtra + 1
            If nExtra <= 15 Then extra = extra & "  " & rawOf(kk) & vbNewLine
        End If
    Next kk
    If nExtra > 15 Then extra = extra & "  ... and " & (nExtra - 15) & " more" & vbNewLine

    Dim msg As String
    msg = "Kronos import complete. " & found & " names matched (green)." & vbNewLine & _
          "Kronos file saved " & Format$(fileDate, "ddd mmm d, h:mm AM/PM") & ", data through " & _
          Format$(through, "ddd mmm d") & "." & vbNewLine & _
          "Accepted / awarded shifts that ended by then now count from Kronos instead of as provisional." & _
          IIf(histId <> "", vbNewLine & "Snapshot " & histId & " added to Hours History.", "") & _
          IIf(awaitFrom > 0, vbNewLine & vbNewLine & "FIRST IMPORT OF THE NEW YEAR: check two or three people's Worked hours against Kronos. " & _
                             "They should be this year's hours only, not last year's total.", "")
    If pCol = 0 Then
        msg = msg & vbNewLine & "No 'Payroll #' column was found in the Kronos sheet, so names were matched as before."
    ElseIf nFilled > 0 Then
        msg = msg & vbNewLine & "Payroll # filled in for " & nFilled & " people (roster column A, admin view)."
    End If
    If nameClash <> "" Then msg = msg & vbNewLine & vbNewLine & _
        "SAME Payroll #, DIFFERENT NAME - nothing changed for these (rename them with Employee > Rename, or fix column A):" & vbNewLine & nameClash
    If missing <> "" Then msg = msg & vbNewLine & vbNewLine & _
        "On the roster but not found in Kronos - Worked Hours NOT updated:" & vbNewLine & missing
    If extra <> "" Then msg = msg & vbNewLine & _
        "In Kronos but not on the roster (spelling different, or not on this roster):" & vbNewLine & extra
    If Len(msg) > 1000 Then msg = Left$(msg, 980) & vbNewLine & "... (more)"
    MsgBox msg, IIf(missing = "" And extra = "" And nameClash = "", vbInformation, vbExclamation), "Kronos import"
    ' v24.2: the Payroll # lists get a box of their own (names only, never a number)
    Dim payMsg As String
    If payDiff <> "" Then payMsg = payMsg & "Matched by name, but column A doesn't hold the Payroll # Kronos has - check column A:" & vbNewLine & payDiff & vbNewLine
    If payShared <> "" Then payMsg = payMsg & "Kronos has one Payroll # under two names, so column A was left blank for:" & vbNewLine & payShared & vbNewLine
    If noNum <> "" Then payMsg = payMsg & "Roster rows with no Payroll # yet (" & nNoNum & "):" & vbNewLine & noNum
    If payMsg <> "" Then
        If Len(payMsg) > 1000 Then payMsg = Left$(payMsg, 980) & vbNewLine & "... (more)"
        MsgBox payMsg, IIf(payDiff = "" And payShared = "", vbInformation, vbExclamation), "Kronos import - Payroll #"
    End If
    Exit Sub
Oops:
    ReportError "SyncKronosData"
    KronosCleanUp wsDest, wbSource, openedHere
End Sub

' After an error part-way through the import: the roster is protected again and
' the Kronos file is closed if the import opened it. Quiet.
Private Sub KronosCleanUp(ByVal wsDest As Worksheet, ByVal wbSource As Workbook, ByVal openedHere As Boolean)
    On Error Resume Next
    If Not wsDest Is Nothing Then ProtectSheet wsDest
    If openedHere And Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    ThisWorkbook.Activate
End Sub


' ====================================================================
'  IMPORT DATE / STALE CHECK (v20)
' ====================================================================
Private Sub SetDateName(ByVal nmName As String, ByVal d As Date)
    Dim ref As String: ref = "=" & Replace(Format$(CDbl(d), "0.00000"), ",", ".")
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(nmName)
    On Error GoTo 0
    On Error Resume Next
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=nmName, RefersTo:=ref, Visible:=False
    Else
        nmObj.RefersTo = ref                  ' in place: a formula that reads the name never sees it missing
        nmObj.Visible = False
    End If
    On Error GoTo 0
End Sub

Private Function GetDateName(ByVal nmName As String) As Date
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(nmName)
    On Error GoTo 0
    If nmObj Is Nothing Then Exit Function
    Dim x As Double: x = Val(Mid$(nmObj.RefersTo, 2))
    If x > 0 Then GetDateName = CDate(x)
End Function

' ---- waiting for the new year's Kronos data (v23) ----
' Set by RolloverYear to Jan 1 of the new year: Worked (H) is 0 and an
' import is refused until the file's data runs through that date.
Public Function KronosAwaitFrom() As Date
    KronosAwaitFrom = GetDateName(AWAIT_NAME)
End Function

Public Sub SetKronosAwait(ByVal d As Date)
    SetDateName AWAIT_NAME, d
End Sub

Public Sub ClearKronosAwait(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    On Error Resume Next
    ThisWorkbook.Names(AWAIT_NAME).Delete
    On Error GoTo 0
End Sub

' When the Kronos data was saved (falls back to when it was imported); 0 = never
Public Function KronosDataDate() As Date
    KronosDataDate = GetDateName("KronosFileDate")
    If KronosDataDate = 0 Then KronosDataDate = GetDateName("KronosImportedAt")
End Function

Public Function KronosImportedAt() As Date
    KronosImportedAt = GetDateName("KronosImportedAt")
End Function

' ---- data-through date (v21) ----
' The Kronos data runs through the day before the file was saved.
Public Function DataThroughFor(ByVal fileDate As Date) As Date
    DataThroughFor = Int(fileDate) - 1
End Function

' Last date the imported Kronos data covers (0 = not known)
Public Function KronosDataThrough() As Date
    KronosDataThrough = GetDateName(PROV_NAME)
End Function

' Makes sure the hidden name the provisional formulas read exists, and
' returns its value. Before the first v21 import it is worked out from the
' v20 Kronos file date; with no import at all, yesterday (so nothing that
' has already ended is counted twice). Safe to call any time.
Public Function EnsureDataThrough() As Date
    Dim d As Date: d = KronosDataThrough()
    If d = 0 Then
        If KronosDataDate() > 0 Then
            d = DataThroughFor(KronosDataDate())
        Else
            d = Date - 1
        End If
        If Not ThisWorkbook.ReadOnly Then SetDateName PROV_NAME, d
    End If
    EnsureDataThrough = d
End Function

' Status-line / SystemCheck text, e.g. "Provisional OT: Kronos data through Oct 9"
Public Function DataThroughText() As String
    Dim d As Date: d = KronosDataThrough()
    If d = 0 Then
        DataThroughText = "Provisional OT: data-through date not set (run Import Kronos)"
    Else
        DataThroughText = "Provisional OT: Kronos data through " & Format$(d, "mmm d")
    End If
End Function

' Days since the Kronos data was saved; -1 = never imported (v20)
Public Function KronosAgeDays() As Long
    Dim d As Date: d = KronosDataDate()
    If d = 0 Then KronosAgeDays = -1 Else KronosAgeDays = Int(Date) - Int(d)
End Function

Public Function KronosIsStale() As Boolean
    Dim aw As Date: aw = KronosAwaitFrom()       ' v23: after a rollover the clock starts on Jan 1
    If aw > 0 Then KronosIsStale = (Int(Date) > Int(aw) + KRONOS_STALE_DAYS): Exit Function
    Dim a As Long: a = KronosAgeDays()
    KronosIsStale = (a < 0 Or a > KRONOS_STALE_DAYS)
End Function

' Status line text, e.g. "Kronos: Oct 7 (2 days ago)"
Public Function KronosStatusText() As String
    Dim aw As Date: aw = KronosAwaitFrom()
    If aw > 0 Then KronosStatusText = "Kronos: awaiting " & Year(aw) & " data (Worked = 0)": Exit Function
    Dim a As Long: a = KronosAgeDays()
    If a < 0 Then KronosStatusText = "Kronos: not imported since v20": Exit Function
    Dim t As String
    Select Case a
        Case 0: t = "today"
        Case 1: t = "yesterday"
        Case Else: t = a & " days ago"
    End Select
    KronosStatusText = "Kronos: " & Format$(KronosDataDate(), "mmm d") & " (" & t & ")" & _
                       IIf(a > KRONOS_STALE_DAYS, " - import", "")
End Function

' Call-out / posting warning. True = go ahead.
Public Function KronosOkToContinue(ByVal what As String) As Boolean
    KronosOkToContinue = True
    If Not KronosIsStale() Then Exit Function
    Dim a As Long: a = KronosAgeDays()
    Dim msg As String
    If a < 0 Then
        msg = "Worked hours haven't been imported from Kronos since this update was installed, " & _
              "so there's no record of how current they are."
    Else
        msg = "Worked hours were last imported from Kronos data saved " & Format$(KronosDataDate(), "ddd mmm d") & _
              " (" & a & " days ago)."
    End If
    msg = msg & vbNewLine & vbNewLine & "Call order is only as current as that import. Continue with the " & what & " anyway?" & _
          vbNewLine & vbNewLine & "(No = stop here and run Import Kronos first.)"
    KronosOkToContinue = (MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, "Kronos data is old") = vbYes)
End Function


' ====================================================================
'  HOURS HISTORY (v20) - one dated row per person per import
' ====================================================================
Public Function HoursHistoryTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(HH_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(HH_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(HH_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = HH_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
        End If
        UnlockSheet ws
        ws.Range("A1:H1").Value = Array("Import ID", "Imported At", "Kronos File Date", "Employee", "Worked Hours", "Imported By", "Data Through", "Year")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H2"), , xlYes)
        lo.Name = HH_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns("A").ColumnWidth = 10
        ws.Columns("B:C").ColumnWidth = 17
        ws.Columns("D").ColumnWidth = 24
        ws.Columns("E").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 13
        ws.Columns("H").ColumnWidth = 7
        ProtectSheet ws
    End If
    If Not lo Is Nothing Then EnsureHoursHistoryColumns lo
    Set HoursHistoryTable = lo
End Function

' A v20 table has 6 columns: "Data Through" (v21) and "Year" (v23) go on the end, once
Private Sub EnsureHoursHistoryColumns(ByVal lo As ListObject)
    If lo.ListColumns.Count >= HH_NCOLS Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    If lo.ListColumns.Count < 7 Then
        lo.ListColumns.Add.Name = "Data Through"
        ws.Columns(7).ColumnWidth = 13
    End If
    If lo.ListColumns.Count < 8 Then
        lo.ListColumns.Add.Name = "Year"
        ws.Columns(8).ColumnWidth = 7
    End If
    ProtectSheet ws
    On Error GoTo 0
End Sub

' Gives every snapshot that has no Year yet the year y (the rollover stamps
' the year it is closing, so a late-December snapshot is never read as the
' new year's hours). Returns how many rows were stamped.
Public Function StampHoursHistoryYear(ByVal y As Long) As Long
    Dim lo As ListObject: Set lo = HoursHistoryTable(False)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < HH_NCOLS Then Err.Raise vbObjectError + 750, , "the Hours History sheet couldn't be given its Year column"
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim yv As Variant: yv = RangeTo2D(lo.ListColumns(HH_NCOLS).DataBodyRange)
    Dim idv As Variant: idv = RangeTo2D(lo.ListColumns(1).DataBodyRange)
    Dim i As Long, n As Long
    For i = 1 To UBound(yv, 1)
        If Trim$(SafeStr(yv(i, 1))) = "" And Trim$(SafeStr(idv(i, 1))) <> "" Then
            yv(i, 1) = y
            n = n + 1
        End If
    Next i
    If n > 0 Then
        UnlockSheet lo.Parent
        lo.ListColumns(HH_NCOLS).DataBodyRange.Value = yv
        ProtectSheet lo.Parent
    End If
    StampHoursHistoryYear = n
End Function

' Appends the snapshot; returns its Import ID ("KR-0007") or "" on failure
Private Function AppendHoursHistory(ByVal stamp As Date, ByVal fileDate As Date, ByVal through As Date, _
                                    histName() As String, histHrs() As Double, ByVal n As Long) As String
    Dim lo As ListObject: Set lo = HoursHistoryTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < HH_NCOLS Then
        MsgBox "Worked hours were imported, but the Hours History sheet couldn't be given its Data Through / Year columns." & _
               vbNewLine & "Is the file read-only?", vbExclamation
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = lo.Parent

    ' next Import ID
    Dim mx As Long, v As Variant, i As Long, cnt As Long, blankFirst As Boolean
    cnt = lo.ListRows.Count
    If Not lo.DataBodyRange Is Nothing Then
        If cnt = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
            blankFirst = True: cnt = 0
        Else
            v = RangeTo2D(lo.ListColumns(1).DataBodyRange)
            For i = 1 To UBound(v, 1)
                If Left$(SafeStr(v(i, 1)), 3) = "KR-" Then
                    If Val(Mid$(SafeStr(v(i, 1)), 4)) > mx Then mx = Val(Mid$(SafeStr(v(i, 1)), 4))
                End If
            Next i
        End If
    End If
    Dim id As String: id = "KR-" & Format$(mx + 1, "0000")

    Dim outA() As Variant: ReDim outA(1 To n, 1 To HH_NCOLS)
    Dim yrNow As Long
    On Error Resume Next
    yrNow = RosterYear()
    On Error GoTo 0
    For i = 1 To n
        outA(i, 1) = id
        outA(i, 2) = stamp
        outA(i, 3) = fileDate
        outA(i, 4) = histName(i)
        outA(i, 5) = histHrs(i)
        outA(i, 6) = CurrentUser()
        outA(i, 7) = through
        If yrNow > 0 Then outA(i, 8) = yrNow
    Next i

    Dim em As String
    On Error GoTo Fail
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, HH_NCOLS)
        .Value = outA
        .Columns(2).NumberFormat = "mmm d, yyyy h:mm"
        .Columns(3).NumberFormat = "mmm d, yyyy h:mm"
        .Columns(7).NumberFormat = "mmm d, yyyy"
    End With
    ProtectSheet ws
    AppendHoursHistory = id
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendHoursHistory", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Worked hours were imported, but the Hours History snapshot couldn't be saved:" & vbNewLine & em, vbExclamation
End Function

' v24.3.1: the column whose heading (first 10 rows) says `want` by its letters and digits
' (IsHeaderMatch in modOvertimeEntry: "Employee  Name" and "EMPLOYEE NAME" match, "Name" doesn't).
' 0 = none; hdrRow = the row it was found on.
Private Function HeaderColumn(ByRef src As Variant, ByRef hdrRow As Long, ByVal want As String) As Long
    Dim r As Long, c As Long, rMax As Long
    hdrRow = 0
    rMax = UBound(src, 1)
    If rMax > 10 Then rMax = 10
    For r = 1 To rMax
        For c = 1 To UBound(src, 2)
            If VarType(src(r, c)) = vbString Then
                If IsHeaderMatch(CStr(src(r, c)), want) Then
                    HeaderColumn = c
                    hdrRow = r
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

' The column of the Kronos sheet whose heading says Payroll # (0 = none), on the heading row hdrRow
' (v24.3.1: that row only - before, the first match in the first 10 rows).
Private Function PayrollColumn(ByRef src As Variant, ByVal hdrRow As Long) As Long
    Dim c As Long
    If hdrRow < 1 Or hdrRow > UBound(src, 1) Then Exit Function
    For c = 1 To UBound(src, 2)
        If VarType(src(hdrRow, c)) = vbString Then
            If IsPayrollHeader(CStr(src(hdrRow, c))) Then
                PayrollColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Public Function NameKey(ByVal s As String) As String
    NameKey = LCase$(Replace(Replace(Trim$(s), " ", ""), ".", ""))
End Function

' exact match, or the roster name + a middle initial in Kronos ("Dobbin,Stephen A")
Private Function FindKey(ByVal dict As Object, ByVal alt As Object, ByVal k As String) As String
    If k = "" Then Exit Function
    If dict.Exists(k) Then FindKey = k: Exit Function
    If alt.Exists(k) Then FindKey = CStr(alt(k))
End Function
