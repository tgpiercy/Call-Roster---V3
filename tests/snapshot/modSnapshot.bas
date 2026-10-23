Attribute VB_Name = "modSnapshot"
Option Explicit
' ====================================================================
'  TEST-ONLY - THE v27.1 REFACTOR PROOF. NEVER SHIPS. NEVER ON THE REAL ROSTER.
' ====================================================================
' It lives in tests/snapshot/, not modules/: build_zip never packs it and SystemCheck never counts it.
'
' TakeSnapshot draws the pages a supervisor sees and writes, for each one, a hash of every row
' (each cell's value, fill, font colour, bold, italic, size, number format, alignment, wrap,
' borders, merge; the row's height and hidden flag), of the column widths, and of every shape
' (name, type, place, size, fill, text, the macro it runs) to snapshot-<release>.txt beside the
' workbook. Only hashes and sheet / shape names are written: no names, no phones.
'     Calendar   this month, every named group (All units is the sum of them)
'     Needs Attention, the two-week Board, the newest posting's sheet, the printed call-out
'     roster, the Desk
' Run it on a practice copy on v27, install v27.1 on that copy, run it again, then
' CompareSnapshots: the two files must be identical. Both runs the same day (the Board, the Desk
' and Needs Attention count from today), with nothing logged in between.
'
' Two boxes ask to print (the call-out roster, the Board): answer No. The Board also saves its
' PDF beside the calendars, as it always does.
' Text that carries a clock - "Drawn ...", "printed ...", "updated ...", "N days old / ago" -
' has its digits masked before hashing; on the Desk every digit is masked (its counts move with
' each open: "since you were last here"). A macro's module name is dropped before hashing (it is
' MEANT to change: modCalendar.BoardPrint -> modBoard.BoardPrint; callcheck's MACROSTR proves the
' new name resolves). It calls only Public procedures that keep their names through the split, so
' the one file compiles on both releases.

Private mOut As String
Private mItems As Long

Public Sub TakeSnapshot()
    On Error GoTo Oops
    If IsLiveRoster() Then
        MsgBox "TakeSnapshot runs on a PRACTICE copy only (this is the live roster).", vbExclamation, "Snapshot"
        Exit Sub
    End If
    If MsgBox("TakeSnapshot draws the Calendar (every group), Needs Attention, the Board, the newest posting's sheet, " & _
              "the call-out roster and the Desk, and writes a hash of each beside this workbook." & vbNewLine & vbNewLine & _
              "Two boxes will ask to print: answer No." & vbNewLine & vbNewLine & "Go ahead?", _
              vbYesNo + vbQuestion, "Snapshot") <> vbYes Then Exit Sub
    Dim t0 As Double: t0 = Timer
    mOut = "# snapshot " & EXPECTED_RELEASE & vbNewLine & _
           "# taken " & Format$(Now, "yyyy-mm-dd hh:nn") & " by " & Environ$("USERNAME") & vbNewLine & _
           "# day " & Format$(Date, "yyyy-mm-dd") & vbNewLine & _
           "# workbook " & ThisWorkbook.Name & vbNewLine
    mItems = 0
    Dim g As Variant, parts As Variant, unit As String
    For Each g In CalGroups()
        parts = Split(g, "|")
        If parts(1) <> "*" Then
            unit = Split(parts(1), ",")(0)
            CalShowDate Date, unit
            DrawCalendar
            SnapSheet ThisWorkbook.Worksheets(CAL_SHEET), "Calendar[" & parts(0) & "]", False
        End If
    Next g
    DrawNeeds
    SnapSheet ThisWorkbook.Worksheets(NA_SHEET), "Needs Attention", False
    BoardPrint
    SnapSheet ThisWorkbook.Worksheets(BOARD_SHEET), "Board", False
    Dim pid As String, rep As String: pid = NewestPostingId()
    If pid <> "" Then
        If PostingOutput(pid, False, False, False, rep) Then SnapSheet ThisWorkbook.Worksheets(POSTING_SHEET), "Posting " & pid, False
    Else
        mOut = mOut & "# no posting to draw" & vbNewLine
    End If
    Call_Out
    SnapSheet ThisWorkbook.Worksheets(CALLOUT_SHEET), "Call Out Roster", False
    DrawDesk
    SnapSheet ThisWorkbook.Worksheets(DESK_SHEET), "Desk", True
    Dim path As String: path = ThisWorkbook.Path & Application.PathSeparator & "snapshot-" & EXPECTED_RELEASE & ".txt"
    WriteText path, mOut
    MsgBox "Snapshot written: " & mItems & " items in " & Format$(Timer - t0, "0") & " s." & vbNewLine & path & vbNewLine & vbNewLine & _
           "Next: install the new release on THIS copy, run TakeSnapshot again, then CompareSnapshots.", vbInformation, "Snapshot"
    Exit Sub
Oops:
    Dim em As String: em = Err.Description
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "The snapshot stopped: " & em & vbNewLine & "(nothing was written)", vbExclamation, "Snapshot"
End Sub

' Compares the two newest snapshot-*.txt files beside the workbook (or the two named). Lists every item that differs
' in a box and in snapshot-compare.txt.
Public Sub CompareSnapshots(Optional ByVal fileA As String = "", Optional ByVal fileB As String = "")
    On Error GoTo Oops
    Dim folder As String: folder = ThisWorkbook.Path & Application.PathSeparator
    If fileA = "" Or fileB = "" Then
        If Not TwoNewest(folder, fileA, fileB) Then
            MsgBox "Two snapshot-*.txt files are needed beside this workbook.", vbExclamation, "Snapshot"
            Exit Sub
        End If
    End If
    Dim a As Object, b As Object, ha As String, hb As String
    Set a = ReadSnap(folder & fileA, ha)
    Set b = ReadSnap(folder & fileB, hb)
    Dim diffs As String, nDiff As Long, k As Variant
    For Each k In a.Keys
        If Not b.Exists(k) Then
            nDiff = nDiff + 1: diffs = diffs & "only in " & fileA & ": " & k & vbNewLine
        ElseIf a(k) <> b(k) Then
            nDiff = nDiff + 1: diffs = diffs & "differs: " & k & vbNewLine
        End If
    Next k
    For Each k In b.Keys
        If Not a.Exists(k) Then nDiff = nDiff + 1: diffs = diffs & "only in " & fileB & ": " & k & vbNewLine
    Next k
    Dim dayA As String, dayB As String
    dayA = HeaderValue(ha, "# day "): dayB = HeaderValue(hb, "# day ")
    Dim note As String
    If dayA <> dayB Then note = vbNewLine & "WARNING: taken on different days (" & dayA & " / " & dayB & "): the Board, the Desk and Needs Attention count from today - retake both on one day."
    Dim res As String
    res = fileA & " (" & a.Count & " items)" & vbNewLine & fileB & " (" & b.Count & " items)" & vbNewLine & vbNewLine
    If nDiff = 0 Then
        res = res & "IDENTICAL." & note
    Else
        res = res & "DIFFERENT: " & nDiff & " item(s)." & note & vbNewLine & vbNewLine & diffs
    End If
    WriteText folder & "snapshot-compare.txt", res
    If Len(res) > 900 Then res = Left$(res, 900) & vbNewLine & "... (the whole list is in snapshot-compare.txt)"
    MsgBox res, IIf(nDiff = 0, vbInformation, vbExclamation), "Snapshot"
    Exit Sub
Oops:
    MsgBox "The compare stopped: " & Err.Description, vbExclamation, "Snapshot"
End Sub

' ---- one sheet: its rows, its column widths, its shapes ----
Private Sub SnapSheet(ByVal ws As Worksheet, ByVal label As String, ByVal maskAll As Boolean)
    Dim lastR As Long, lastC As Long, r As Long, c As Long, s As String, cl As Range
    LastUsed ws, lastR, lastC
    For r = 1 To lastR
        s = "h=" & ws.Rows(r).RowHeight & ";x=" & ws.Rows(r).Hidden
        For c = 1 To lastC
            Set cl = ws.Cells(r, c)
            s = s & "|" & CellText(cl, maskAll)
        Next c
        AddItem label & "|row " & r, s
    Next r
    s = ""
    For c = 1 To lastC + 2
        s = s & "|" & ws.Columns(c).ColumnWidth & IIf(ws.Columns(c).Hidden, "x", "")
    Next c
    AddItem label & "|columns", s
    Dim shp As Shape, i As Long
    For Each shp In ws.Shapes
        i = i + 1
        AddItem label & "|shape " & shp.Name & " #" & i, ShapeText(shp, maskAll)
    Next shp
End Sub

Private Function CellText(ByVal cl As Range, ByVal maskAll As Boolean) As String
    Dim s As String, t As String, e As Variant
    On Error Resume Next                           ' never raises: a property a cell can't give is read as blank, the same on both runs
    t = CStr(cl.Value2)
    If maskAll Or IsClockText(t) Then t = MaskDigits(t)
    s = t & "~" & cl.Interior.Color & "~" & cl.Interior.Pattern & "~" & cl.Font.Color & "~" & cl.Font.Bold & "~" & cl.Font.Italic & _
        "~" & cl.Font.Size & "~" & cl.Font.Name & "~" & cl.NumberFormat & "~" & cl.HorizontalAlignment & "~" & cl.VerticalAlignment & _
        "~" & cl.WrapText & "~" & cl.IndentLevel
    If cl.MergeCells Then s = s & "~m" & cl.MergeArea.Address
    For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With cl.Borders(e)
            s = s & "~" & .LineStyle & "/" & .Weight & "/" & .Color
        End With
    Next e
    CellText = s
End Function

Private Function ShapeText(ByVal shp As Shape, ByVal maskAll As Boolean) As String
    Dim s As String, t As String, mac As String, fv As Boolean
    On Error Resume Next                           ' never raises: a shape without text / fill / a macro reads as blank
    s = "t=" & shp.Type & ";v=" & shp.Visible & ";at=" & Format$(shp.Left, "0.0") & "," & Format$(shp.Top, "0.0") & _
        ";sz=" & Format$(shp.Width, "0.0") & "x" & Format$(shp.Height, "0.0")
    fv = False: fv = shp.Fill.Visible              ' read first: a raising test inside an If runs its Then part
    If fv Then s = s & ";fill=" & shp.Fill.ForeColor.RGB
    s = s & ";line=" & shp.Line.Visible & "/" & shp.Line.ForeColor.RGB
    t = ""
    t = shp.TextFrame2.TextRange.Text
    If maskAll Or IsClockText(t) Then t = MaskDigits(t)
    s = s & ";text=" & t
    s = s & ";font=" & shp.TextFrame2.TextRange.Font.Size & "/" & shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB & "/" & shp.TextFrame2.TextRange.Font.Bold
    mac = shp.OnAction
    If InStr(mac, "!") > 0 Then mac = Mid$(mac, InStrRev(mac, "!") + 1)
    If InStr(mac, ".") > 0 Then mac = Mid$(mac, InStrRev(mac, ".") + 1)
    s = s & ";macro=" & mac
    If shp.Type = msoFormControl Then s = s & ";list=" & shp.ControlFormat.ListCount & ";val=" & shp.ControlFormat.Value
    ShapeText = s
End Function

Private Function IsClockText(ByVal t As String) As Boolean
    Dim l As String: l = LCase$(t)
    IsClockText = (InStr(l, "drawn") > 0 Or InStr(l, "printed") > 0 Or InStr(l, "updated") > 0 Or _
                   InStr(l, " days old") > 0 Or InStr(l, " days ago") > 0 Or InStr(l, " day ago") > 0)
End Function

Private Function MaskDigits(ByVal t As String) As String
    Dim i As Long
    For i = 0 To 9
        t = Replace(t, CStr(i), "#")
    Next i
    MaskDigits = t
End Function

Private Sub LastUsed(ByVal ws As Worksheet, ByRef lastR As Long, ByRef lastC As Long)
    Dim f As Range
    lastR = 0: lastC = 0
    Set f = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then lastR = f.Row
    Set f = ws.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then lastC = f.Column
    lastR = lastR + 2                              ' a coloured row or a border just below the text counts too
    lastC = lastC + 1
End Sub

Private Sub AddItem(ByVal key As String, ByVal text As String)
    mItems = mItems + 1
    mOut = mOut & key & vbTab & HashText(text) & vbNewLine
End Sub

' Two independent 31-bit polynomial hashes (exact in a Double): a change anywhere in the text changes both.
Private Function HashText(ByVal s As String) As String
    Dim h1 As Double, h2 As Double, i As Long, c As Long
    h1 = 7: h2 = 11
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        h1 = h1 * 131 + c: h1 = h1 - Int(h1 / 2147483647#) * 2147483647#
        h2 = h2 * 137 + c: h2 = h2 - Int(h2 / 2147483629#) * 2147483629#
    Next i
    HashText = Format$(h1, "0") & "-" & Format$(h2, "0") & "-" & Len(s)
End Function

Private Function NewestPostingId() As String
    Dim lo As ListObject, ws As Worksheet, v As Variant, i As Long, best As Long, n As Long, id As String
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = PT_TABLE Then
                If Not lo.DataBodyRange Is Nothing Then
                    v = lo.DataBodyRange.Columns(PS_ID).Value
                    If IsArray(v) Then
                        For i = 1 To UBound(v, 1)
                            id = CStr(v(i, 1))
                            If Left$(id, 2) = "P-" And IsNumeric(Mid$(id, 3)) Then
                                n = CLng(Mid$(id, 3))
                                If n > best Then best = n: NewestPostingId = id
                            End If
                        Next i
                    Else
                        NewestPostingId = CStr(v)
                    End If
                End If
            End If
        Next lo
    Next ws
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim ff As Integer: ff = FreeFile
    Open path For Output As #ff
    Print #ff, text;
    Close #ff
End Sub

Private Function ReadSnap(ByVal path As String, ByRef header As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ff As Integer, ln As String, tabAt As Long
    ff = FreeFile
    Open path For Input As #ff
    Do While Not EOF(ff)
        Line Input #ff, ln
        If Left$(ln, 1) = "#" Then
            header = header & ln & vbNewLine
        Else
            tabAt = InStr(ln, vbTab)
            If tabAt > 0 Then d(Left$(ln, tabAt - 1)) = Mid$(ln, tabAt + 1)
        End If
    Loop
    Close #ff
    Set ReadSnap = d
End Function

Private Function HeaderValue(ByVal header As String, ByVal tag As String) As String
    Dim p As Long: p = InStr(header, tag)
    If p = 0 Then Exit Function
    HeaderValue = Mid$(header, p + Len(tag))
    p = InStr(HeaderValue, vbNewLine)
    If p > 0 Then HeaderValue = Left$(HeaderValue, p - 1)
End Function

' The two newest snapshot-*.txt beside the workbook, older first.
Private Function TwoNewest(ByVal folder As String, ByRef fileA As String, ByRef fileB As String) As Boolean
    Dim f As String, t As Date, t1 As Date, t2 As Date, n1 As String, n2 As String
    f = Dir$(folder & "snapshot-*.txt")
    Do While f <> ""
        If LCase$(f) <> "snapshot-compare.txt" Then
            t = FileDateTime(folder & f)
            If t > t1 Then
                t2 = t1: n2 = n1: t1 = t: n1 = f
            ElseIf t > t2 Then
                t2 = t: n2 = f
            End If
        End If
        f = Dir$()
    Loop
    If n1 = "" Or n2 = "" Then Exit Function
    fileA = n2: fileB = n1
    TwoNewest = True
End Function
