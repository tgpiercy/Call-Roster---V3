Attribute VB_Name = "modMinStaff"
Option Explicit

Private mStaffMin As Object         ' v26: "unit|Days" -> minimum, read once per sweep (StaffSweepBegin)
Private mStaffNames As Object       ' v26: unit (or unit|date) -> the names holding a code, one LoadEligible each per sweep
Private mStaffPerDate As Boolean    ' v26: a planned unit move inside the window: the names are per unit AND date
Private mStaffBusy As Object        ' v26: names with any absence row or a switch near the window - the only ones the off / switch lookups run for
' ====================================================================
'  MINIMUM STAFFING (v26, kit item 2; the numbers from Graeme, Oct 18)
' ====================================================================
' The Min Staff sheet (admin-typed: Unit, Shift, Minimum, Note) holds the
' plant's minimums. For today .. today + 13 a unit's shift is SHORT when
'   staffed = scheduled - off + covered  <  minimum
' scheduled = the people on the crew that is on that shift holding a code
' in the unit (LoadEligible as of the date, CrewOn / CrewShiftOn: a DAYS
' crew off on a plant holiday is not scheduled); off = booked off or sick
' on that shift (OffForStaffing) or off in a switch; covered = the coverage
' lines a named person covers - a switch's own S/S line counts its worker
' (OPEN, NO TAKERS, POSTED and NCN add nobody). Shown on the calendar's
' date cell and in Needs Attention. Reads only; no rule changes.

' PURE (SelfTest "minstaff"): "" (fine, or no minimum) / "SHORT staffed/minimum"
Public Function StaffShort(ByVal scheduled As Long, ByVal off As Long, ByVal covered As Long, ByVal minimum As Long) As String
    If minimum <= 0 Then Exit Function
    Dim staffed As Long: staffed = scheduled - off + covered
    If staffed < minimum Then StaffShort = "SHORT " & staffed & "/" & minimum
End Function

' The plant's minimums on first use (Graeme, Oct 18, 2026): unit, days, nights (0 = none), note
Private Function MinStaffDefaults() As Variant
    MinStaffDefaults = Array( _
        Array("Phosphine", 2, 2, "1 per train (train 1 + train 2); the roster cannot tell the trains apart, so 2 is the check"), _
        Array("Derivatives", 3, 3, ""), _
        Array("Cypure", 1, 1, ""), _
        Array("Material Handler", 1, 1, "fumigation"), _
        Array("QC Derivatives", 1, 1, "the QC lab; blank the Nights minimum if the lab runs no nights"), _
        Array("QC Gases", 1, 0, "1 on days and 1 on afternoons; the roster has no afternoon shift, so only days are checked"), _
        Array("Utilities", 1, 0, "days only"), _
        Array("SBP", 1, 0, "days only"))
End Function

' The settings table, made on first use with the defaults (an admin corrects the Minimum cells)
Public Function MinStaffTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(MINSTAFF_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(MINSTAFF_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(MINSTAFF_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ActiveSheet
            Dim evWas As Boolean: evWas = Application.EnableEvents
            Dim em As String, en As Long
            Application.EnableEvents = False                 ' review: made at open - no SheetActivate for the Add and the Activate back
            On Error GoTo AddFailed
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = MINSTAFF_SHEET
            RelockStructure
            SetTabColor ws, RGB(176, 0, 32)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
            Application.EnableEvents = evWas
        End If
        UnlockSheet ws
        Dim defs As Variant: defs = MinStaffDefaults()
        Dim n As Long: n = 2 * (UBound(defs) - LBound(defs) + 1)
        Dim rows() As Variant: ReDim rows(1 To n, 1 To 4)
        Dim i As Long, k As Long, row As Variant
        For i = LBound(defs) To UBound(defs)
            row = defs(i)
            k = k + 1: rows(k, 1) = row(0): rows(k, 2) = "Days": rows(k, 4) = row(3)
            If CLng(row(1)) > 0 Then rows(k, 3) = CLng(row(1))
            k = k + 1: rows(k, 1) = row(0): rows(k, 2) = "Nights": rows(k, 4) = row(3)
            If CLng(row(2)) > 0 Then rows(k, 3) = CLng(row(2))
        Next i
        ws.Range("A1").Resize(1, 4).Value = Array("Unit", "Shift", "Minimum", "Note")
        ws.Range("A2").Resize(n, 4).Value = rows
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(n + 1, 4), , xlYes)
        lo.Name = MINSTAFF_TABLE
        lo.TableStyle = "TableStyleMedium3"
        ws.Columns(1).ColumnWidth = 18: ws.Columns(2).ColumnWidth = 8: ws.Columns(3).ColumnWidth = 10: ws.Columns(4).ColumnWidth = 90
        ws.Range("A" & (n + 3)).Value = "Minimum operators per unit and shift (admin). Blank = no minimum on that shift. The calendar's date cells and " & _
                                        "Needs Attention flag a shift below its minimum for the next " & BOARD_DAYS & " days. Change them through the admin Desk's Holidays / minimum staffing."
        ws.Range("A" & (n + 3)).Font.Italic = True: ws.Range("A" & (n + 3)).Font.Size = 8.5
        LockMinStaffCells gAdminUnlocked
        ProtectSheet ws
        If gDeskMode Then SetSheetVisible ws, xlSheetVeryHidden      ' review: an admin's Desk view hides the admin sheets (ApplyViewCore ran before)
    End If
    Set MinStaffTable = lo
    Exit Function
AddFailed:                                           ' re-review: events back on whatever failed, then the error goes up
    en = Err.Number: em = Err.Description
    On Error Resume Next
    Application.EnableEvents = evWas
    RelockStructure
    Err.Raise en, , em
End Function

' v26 (R4): the cells an admin may type into on the Min Staff sheet - Minimum and Note; v27 (item 33): none (adminView unused).
' Called by ApplyViewCore (the sheet is re-protected by ProtectAll after it) and at the table's making.
Public Sub LockMinStaffCells(ByVal adminView As Boolean)
    Dim lo As ListObject: Set lo = MinStaffTable(False)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    ws.Cells.Locked = True                          ' v27 (item 33): every cell - the Min Staff pop-up writes Minimum and Note (SetMinStaff)
    ProtectSheet ws
End Sub

' "unit|Days" -> minimum (only > 0), from the table; empty when the sheet isn't made yet
Public Function MinStaffMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set MinStaffMap = d
    Dim lo As ListObject: Set lo = MinStaffTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, u As String, sh As String, m As Long
    For i = 1 To UBound(v, 1)
        u = Trim$(SafeStr(v(i, 1))): sh = Trim$(SafeStr(v(i, 2)))
        m = CLng(NumOr0(v(i, 3)))
        If u <> "" And m > 0 Then
            If UCase$(Left$(sh, 1)) = "N" Then sh = "Nights" Else sh = "Days"
            d(u & "|" & sh) = m
        End If
    Next i
End Function

' v27 (item 25, 1.4): the admin-typed CALL-OUT NOTE on the Min Staff sheet for a unit and shift ("" when none) - the
' 5th column (review M2: the Note column holds the minimums' own explanations, "days only" ..., which were never
' meant for the call-out pop-up). The pop-up shows it at LOAD. Never raises.
Public Function MinStaffNote(ByVal unit As String, ByVal shiftLbl As String) As String
    On Error Resume Next                            ' never raises
    Dim lo As ListObject: Set lo = MinStaffTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    If UBound(v, 2) < MS_CALLNOTE Then Exit Function      ' an older table: no call-out note yet
    Dim i As Long, sh As String, want As String
    want = IIf(IsNightShift(shiftLbl), "N", "D")
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), Trim$(unit), vbTextCompare) = 0 Then
            sh = UCase$(Left$(Trim$(SafeStr(v(i, 2))), 1))
            If sh <> "N" Then sh = "D"
            If sh = want And Trim$(SafeStr(v(i, MS_CALLNOTE))) <> "" Then MinStaffNote = Trim$(SafeStr(v(i, MS_CALLNOTE))): Exit Function
        End If
    Next i
End Function

' One sweep's caches (the start of DrawCalendar / DrawNeeds): the minimums, the names per unit, and the
' "busy" names - anyone with an absence row on file or on a switch dated near the window. The off / switch
' lookups scan whole tables per name, so they run only for those (review: the sweep's cost).
Public Sub StaffSweepBegin(ByVal d0 As Date)
    Set mStaffMin = MinStaffMap()
    Set mStaffNames = CreateObject("Scripting.Dictionary")
    mStaffNames.CompareMode = 1
    mStaffPerDate = PlannedWithin(d0, d0 + BOARD_DAYS - 1)
    Set mStaffBusy = CreateObject("Scripting.Dictionary")
    mStaffBusy.CompareMode = 1
    If mShiftMap Is Nothing Then LoadCache
    Dim i As Long, nm As String, kk As Variant
    If IsArray(mAbs) Then
        For i = 1 To UBound(mAbs, 1)
            If UCase$(SafeStr(mAbs(i, AB_STATUS))) <> "VOID" Then
                nm = Trim$(SafeStr(mAbs(i, AB_EMP)))
                If nm <> "" Then mStaffBusy(nm) = True
            End If
        Next i
    End If
    Dim sw As Object: Set sw = SwitchNamesBetween(d0 - 1, d0 + BOARD_DAYS)
    For Each kk In sw.Keys
        mStaffBusy(CStr(kk)) = True
    Next kk
End Sub

Public Function MinStaffCount() As Long
    If mStaffMin Is Nothing Then Exit Function
    MinStaffCount = mStaffMin.Count
End Function

' Everyone holding a code in the unit (any grade, any crew; a planned move counts from its date), as names
Private Function StaffNames(ByVal unit As String, ByVal d As Date) As Variant
    If mStaffNames Is Nothing Then StaffSweepBegin Date
    Dim key As String: key = unit & IIf(mStaffPerDate, "|" & CLng(Int(d)), "")
    If mStaffNames.Exists(key) Then StaffNames = mStaffNames(key): Exit Function
    Dim recs() As EmpRec, n As Long, i As Long, names() As String
    If mStaffPerDate Then
        n = LoadEligible(unit, "ALL", recs, 999, True, d)
    Else
        n = LoadEligible(unit, "ALL", recs, 999, True)
    End If
    If n > 0 Then
        ReDim names(1 To n)
        For i = 1 To n
            names(i) = recs(i).EmpName
        Next i
        mStaffNames(key) = names
        StaffNames = names
    Else
        mStaffNames(key) = Empty
    End If
End Function

' The counts behind one unit / date / shift (reads the caches only; switch reads batched by the caller)
Private Sub StaffCounts(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String, _
                        ByRef scheduled As Long, ByRef off As Long, ByRef covered As Long)
    scheduled = 0: off = 0: covered = 0
    Dim names As Variant: names = StaffNames(unit, d)
    Dim i As Long, nm As String, onShift As Boolean, fb As String, busy As Boolean
    Dim people As Variant, tokens() As String           ' v27.1: one token per name for StaffCountsCore
    If IsArray(names) Then
        ReDim tokens(LBound(names) To UBound(names))
        For i = LBound(names) To UBound(names)
            nm = CStr(names(i))
            fb = ""                                         ' the roster crew as the fallback (no cell scan in CrewOn)
            If mShiftMap.Exists(nm) Then fb = CStr(mShiftMap(nm))
            onShift = (CrewShiftOn(CrewOn(nm, d, fb), d) = shiftLbl)
            If mStaffBusy Is Nothing Then busy = True Else busy = mStaffBusy.Exists(nm)    ' re-review: a sweep that failed half-way
            ' (review) the one WORKING a switch is not added here: every switch has its own S/S calendar line, which
            ' counts them below as a covered line; adding them here too counted every switch twice
            If onShift Then
                tokens(i) = "S"
                If busy Then
                    If OffForStaffing(nm, d, shiftLbl) Then
                        tokens(i) = "SA"
                    ElseIf SwitchOffOn(nm, d, shiftLbl) <> "" Then
                        tokens(i) = "SW"
                    End If
                End If
            End If
        Next i
        people = tokens
    End If
    Dim slots() As CalSlot, n As Long, rts() As String, texts As Variant
    n = SlotsFor(d, shiftLbl, unit, slots, True)
    If n > 0 Then
        ReDim rts(1 To n)
        For i = 1 To n
            rts(i) = Trim$(slots(i).RightText)
        Next i
        texts = rts
    End If
    StaffCountsCore people, texts, scheduled, off, covered
End Sub

' v27.1 (R4): PURE (SelfTest "staffcount") - the counts behind one unit / date / shift from what StaffCounts gathered.
' people = one token per code-holder: "S" on the shift by rotation, "SA" also booked off or sick for it, "SW" also
' gave it up in a switch ("" = not on the shift). rightTexts = the Covered-by texts of the unit's lines on that
' shift. Not an array = none. scheduled = the tokens with S; off = those with A or W, once each; covered = the people
' each text stands for (CoveredCount), NO TAKERS / POSTED / NCN and a blank text adding nobody.
Public Sub StaffCountsCore(ByRef people As Variant, ByRef rightTexts As Variant, _
                           ByRef scheduled As Long, ByRef off As Long, ByRef covered As Long)
    scheduled = 0: off = 0: covered = 0
    Dim i As Long, tok As String, rt As String
    If IsArray(people) Then
        For i = LBound(people) To UBound(people)
            tok = CStr(people(i))
            If InStr(tok, "S") > 0 Then
                scheduled = scheduled + 1
                If InStr(tok, "A") > 0 Or InStr(tok, "W") > 0 Then off = off + 1
            End If
        Next i
    End If
    If IsArray(rightTexts) Then
        For i = LBound(rightTexts) To UBound(rightTexts)
            rt = Trim$(CStr(rightTexts(i)))
            If rt <> "" Then
                ' an awarded multi-position posting writes its names into one line, "A B / C D" (review): count each; a
                ' partial award "Joe D 6h / Sam K 6h" (re-review) counts its hours as a share of a 12 h shift, rounded up
                If Not (UCase$(rt) Like "NO TAKERS*" Or UCase$(rt) Like "POSTED*" Or UCase$(rt) = "NCN") Then covered = covered + CoveredCount(rt)
            End If
        Next i
    End If
End Sub

' v26 (re-review): how many people a covered-by text stands for: one per " / " part; a part ending in "Nh"
' (a partial award) counts N / 12 of a person; the sum rounded up, at least 1
Private Function CoveredCount(ByVal rt As String) As Long
    Dim parts As Variant: parts = Split(rt, " / ")
    Dim i As Long, part As String, p As Long, hrs As Double, total As Double
    For i = LBound(parts) To UBound(parts)
        part = Trim$(CStr(parts(i)))
        hrs = 0
        If LCase$(Right$(part, 1)) = "h" Then
            p = InStrRev(part, " ")
            If p > 0 Then
                If IsNumeric(Mid$(part, p + 1, Len(part) - p - 1)) Then hrs = Val(Mid$(part, p + 1))
            End If
        End If
        If hrs > 0 And hrs < USUAL_SHIFT_HOURS Then total = total + hrs / USUAL_SHIFT_HOURS Else total = total + 1
    Next i
    CoveredCount = CLng(-Int(-total))
    If CoveredCount < 1 Then CoveredCount = 1
End Function

' v26 (review): booked off or sick FOR THIS SHIFT, as the calendar draws it - a sick call sits on the shift it
' was booked for only; the 24 h window after it is an overtime-list rule (AbsentForShift), not a day off work
Private Function OffForStaffing(ByVal nm As String, ByVal d As Date, ByVal shiftLbl As String) As Boolean
    If Not IsArray(mAbs) Then Exit Function
    Dim r As Long
    For r = 1 To UBound(mAbs, 1)
        If StrComp(Trim$(SafeStr(mAbs(r, AB_EMP))), nm, vbTextCompare) = 0 Then
            If AbsActiveOn(r, d) Then
                If AbsenceOnShift(r, d, shiftLbl) Then OffForStaffing = True: Exit Function
            End If
        End If
    Next r
End Function

' The short shifts of one date inside a calendar group, for the date cell:
' "Days: Phosphine 1/2, Cypure 0/1   Nights: Derivatives 2/3" ("" = none, or no minimums)
Public Function StaffShortText(ByVal d As Date, ByVal groupName As String, Optional ByRef nShort As Long) As String
    nShort = 0
    If mStaffMin Is Nothing Then StaffSweepBegin Date
    If mStaffMin.Count = 0 Then Exit Function
    Dim mine As Boolean
    If Not SwitchBatchActive() Then BeginSwitchBatch: mine = True
    On Error GoTo Done
    Dim sh As Variant, kk As Variant, u As String, s As Long, o As Long, c As Long, t As String, part As String
    For Each sh In Array("Days", "Nights")
        part = ""
        For Each kk In mStaffMin.Keys
            If Right$(CStr(kk), Len(CStr(sh)) + 1) = "|" & CStr(sh) Then
                u = Left$(CStr(kk), Len(CStr(kk)) - Len(CStr(sh)) - 1)
                If UnitInGroup(u, groupName) Then
                    StaffCounts u, d, CStr(sh), s, o, c
                    t = StaffShort(s, o, c, CLng(mStaffMin(kk)))
                    If t <> "" Then part = part & IIf(part = "", "", ", ") & u & " " & Mid$(t, 7): nShort = nShort + 1
                End If
            End If
        Next kk
        If part <> "" Then StaffShortText = StaffShortText & IIf(StaffShortText = "", "", "   ") & CStr(sh) & ": " & part
    Next sh
Done:
    If mine Then EndSwitchBatch
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

' Every short unit / shift from d0 for nDays inside a group, for Needs Attention: a 2-D array
' (1..n, 1..5) = unit, date serial, shift, "staffed 1 of 2 (3 on, 2 off, 0 covered)", "SHORT 1/2"; Empty = none
Public Function StaffShortList(ByVal d0 As Date, ByVal nDays As Long, ByVal groupName As String) As Variant
    If mStaffMin Is Nothing Then StaffSweepBegin d0
    If mStaffMin.Count = 0 Then Exit Function
    Dim mine As Boolean
    If Not SwitchBatchActive() Then BeginSwitchBatch: mine = True
    On Error GoTo Done
    Dim out() As Variant, n As Long, cap As Long: cap = 2 * nDays * mStaffMin.Count + 1
    ReDim out(1 To cap, 1 To 5)
    Dim dd As Long, d As Date, kk As Variant, u As String, sh As String, s As Long, o As Long, c As Long, t As String
    For dd = 0 To nDays - 1
        d = d0 + dd
        For Each kk In mStaffMin.Keys
            u = Left$(CStr(kk), InStr(CStr(kk), "|") - 1)
            sh = Mid$(CStr(kk), InStr(CStr(kk), "|") + 1)
            If UnitInGroup(u, groupName) Then
                StaffCounts u, d, sh, s, o, c
                t = StaffShort(s, o, c, CLng(mStaffMin(kk)))
                If t <> "" Then
                    n = n + 1
                    out(n, 1) = u: out(n, 2) = CDbl(Int(d)): out(n, 3) = sh
                    out(n, 4) = "staffed " & (s - o + c) & " of " & CLng(mStaffMin(kk)) & "  (" & s & " on, " & o & " off, " & c & " covered)"
                    out(n, 5) = t
                End If
            End If
        Next kk
    Next dd
    If n > 0 Then
        Dim res() As Variant, i As Long, j As Long
        ReDim res(1 To n, 1 To 5)
        For i = 1 To n
            For j = 1 To 5
                res(i, j) = out(i, j)
            Next j
        Next i
        StaffShortList = res
    End If
Done:
    If mine Then EndSwitchBatch
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

' ====================================================================
'  THE MIN STAFF POP-UP (v27 item 33)
' ====================================================================
' Alt+F8 / the admin Desk (Holidays / minimum staffing): the pop-up on the Min Staff table
Public Sub ShowMinStaffEditor()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If MinStaffTable(True) Is Nothing Then MsgBox "The Min Staff sheet couldn't be made (is the file read-only?).", vbExclamation, "Minimum staffing": Exit Sub
    frmTableEdit.Setup "minstaff"
    frmTableEdit.Show
    Unload frmTableEdit                             ' (review L6) the form hides itself on Close
    Exit Sub
Oops:
    ReportError "ShowMinStaffEditor"
End Sub

' the pop-up's rows: Unit, Shift, Minimum, Note, or Empty
Public Function MinStaffEditRows() As Variant
    Dim lo As ListObject: Set lo = MinStaffTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    MinStaffEditRows = RangeTo2D(lo.DataBodyRange)
End Function

' The ONE write to the Min Staff table: row r (1-based), the minimum as typed (blank = none, else a whole number 0
' to 20; 0 = none) and the note (blank = none, 250 characters kept). Admin, writable. False = refused (why).
Public Function SetMinStaff(ByVal r As Long, ByVal minTxt As String, ByVal note As String, ByRef why As String, _
                            Optional ByVal callNote As String = "") As Boolean     ' (review M2) + the call-out note, column 5
    why = ""
    If Not RequireAdmin() Then why = "admins only": Exit Function
    If Not RequireWritable() Then why = "this copy can't be saved": Exit Function
    minTxt = Trim$(minTxt): note = Trim$(Replace(Replace(note, vbCr, " "), vbLf, " "))
    callNote = Trim$(Replace(Replace(callNote, vbCr, " "), vbLf, " "))
    Dim m As Long
    If minTxt <> "" Then
        If Not IsNumeric(minTxt) Then why = "the minimum is a whole number (blank = no minimum)": Exit Function
        If Val(minTxt) <> Int(Val(minTxt)) Or Val(minTxt) < 0 Or Val(minTxt) > 20 Then why = "the minimum is a whole number from 0 to 20": Exit Function
        m = CLng(Val(minTxt))
    End If
    Dim lo As ListObject: Set lo = MinStaffTable(False)
    If lo Is Nothing Then why = "there is no Min Staff table": Exit Function
    If lo.DataBodyRange Is Nothing Then why = "the Min Staff table is empty": Exit Function
    If r < 1 Or r > lo.ListRows.Count Then why = "that row is gone - close and open the pop-up again": Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet ws
    If lo.ListColumns.Count < MS_CALLNOTE Then      ' an older table: the column is added once
        lo.ListColumns.Add
        lo.HeaderRowRange.Cells(1, MS_CALLNOTE).Value = "Call-out note"
        ws.Columns(lo.Range.Column + MS_CALLNOTE - 1).ColumnWidth = 50
    End If
    If callNote <> "" Then lo.DataBodyRange.Cells(r, MS_CALLNOTE).Value = "'" & Left$(callNote, 200) Else lo.DataBodyRange.Cells(r, MS_CALLNOTE).ClearContents
    If m > 0 Then lo.DataBodyRange.Cells(r, 3).Value = m Else lo.DataBodyRange.Cells(r, 3).ClearContents
    If note <> "" Then lo.DataBodyRange.Cells(r, 4).Value = "'" & Left$(note, 250) Else lo.DataBodyRange.Cells(r, 4).ClearContents
    ProtectSheet ws
    Application.EnableEvents = evWas
    MarkCalDirty
    SetMinStaff = True
End Function

