Attribute VB_Name = "modReports"
Option Explicit

' ====================================================================
'  REFUSED HOURS BY KIND (v24.1)
'
'  Refused (roster column I) stays ONE figure: it drives Total Hours and
'  call order. This module only shows what it is made of:
'     call-out   charges on OT Log rows of Type "Callout"
'     posting    charges on rows of Type "Posting"
'     other      everything else (adjustments: ADJ-, LV-, NH-, UM-;
'                opening balance; values moved in from typed cells)
'  Same test as the roster's own formulas: not VOID, Period Start inside
'  this roster year. call-out + posting + other = Refused, always.
'
'  Where it shows (decided Oct 12, 2026: no new roster columns):
'   - Excel's status bar (bottom left) for the roster row that is selected
'   - the members' file (modPublish adds the three figures to dRoster)
'   - RefusalsReport: a printable sheet, per person and per unit
'
'  Pure (mirrored in tests/rules.py, checked by SelfTest): RefusedBucket,
'  SplitTotals.
' ====================================================================

Public Const REP_VERSION As String = "Reports v26.5.5 (Oct 21)"

' what SplitTotals keeps per person: a Variant array with these indexes
Public Const SP_CALLOUT As Long = 0     ' hours
Public Const SP_POSTING As Long = 1
Public Const SP_OTHER As Long = 2
Public Const SP_NCALLOUT As Long = 3    ' how many charged rows
Public Const SP_NPOSTING As Long = 4
Public Const HANDOVER_SHEET As String = "Handover"   ' v27 (item 25): the shift handover page, output only (supervisors and admins)
Public Const HANDOVER_HOURS As Long = 12

Private mSplit As Object                ' name -> the array above, for the status bar
Private mSplitStamp As String
Private mBarShown As Boolean            ' True while the status bar holds our text


' ====================================================================
'  PURE RULES
' ====================================================================
' Which part of Refused a log row's charge belongs to
Public Function RefusedBucket(ByVal typ As String) As String
    Select Case LCase$(Trim$(typ))
        Case "callout": RefusedBucket = "callout"
        Case "posting": RefusedBucket = "posting"
        Case Else: RefusedBucket = "other"
    End Select
End Function

' lg = rows in the OT Log layout. firstStart / lastStart = the first and
' last Period Start of the roster year, as whole numbers. Returns
' name -> Array(call-out h, posting h, other h, call-out rows, posting rows).
Public Function SplitTotals(ByRef lg As Variant, ByVal firstStart As Long, ByVal lastStart As Long) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set SplitTotals = d
    If Not IsArray(lg) Then Exit Function
    Dim i As Long, emp As String, h As Double, per As Long, a As Variant
    For i = 1 To UBound(lg, 1)
        emp = Trim$(SafeStr(lg(i, L_EMP)))
        If emp <> "" And UCase$(Trim$(SafeStr(lg(i, L_STATUS)))) <> "VOID" Then
            per = PeriodNumber(lg(i, L_PERIOD))
            h = NumOr0(lg(i, L_CHARGED))
            If h <> 0 And per >= firstStart And per <= lastStart Then
                If d.Exists(emp) Then
                    a = d(emp)
                Else
                    a = Array(0#, 0#, 0#, 0&, 0&)
                End If
                Select Case RefusedBucket(SafeStr(lg(i, L_TYPE)))
                    Case "callout"
                        a(SP_CALLOUT) = a(SP_CALLOUT) + h
                        If h > 0 Then a(SP_NCALLOUT) = a(SP_NCALLOUT) + 1
                    Case "posting"
                        a(SP_POSTING) = a(SP_POSTING) + h
                        If h > 0 Then a(SP_NPOSTING) = a(SP_NPOSTING) + 1
                    Case Else
                        a(SP_OTHER) = a(SP_OTHER) + h
                End Select
                d(emp) = a
            End If
        End If
    Next i
End Function

' v26 PURE (SelfTest "chargewhy"): 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 11-13 -> "th", 21 -> "21st"
Public Function OrdinalText(ByVal n As Long) As String
    Dim sfx As String
    If (n Mod 100) >= 11 And (n Mod 100) <= 13 Then
        sfx = "th"
    Else
        Select Case n Mod 10
            Case 1: sfx = "st"
            Case 2: sfx = "nd"
            Case 3: sfx = "rd"
            Case Else: sfx = "th"
        End Select
    End If
    OrdinalText = n & sfx
End Function

' v26 PURE (SelfTest "chargewhy"): "why was I charged" from an OT Log row's own fields - no new string
' in the log, nothing parsed out of the Note (the caller appends the Note where it adds something).
'   typ        Callout / Posting / anything else (Adjustment, the opening balance ...)
'   outcome    the Outcome column (Refused, Left Message, Accepted, Unavailable, Awarded, Signed Up ...);
'              for a row that is not a call-out or posting, the type held in its Reason column
'   callOrder  the Call Order column (0 = none)
'   hours      the Hours column: the hours offered (call-out) or the posting's / awarded hours
'   charged    the Hours Charged column
'   code       the Dept Code column
' Charging itself lives in CalloutCharge / ComputeAward ...; this only says what the row shows.
Public Function ChargeWhy(ByVal typ As String, ByVal outcome As String, ByVal callOrder As Long, _
                          ByVal hours As Double, ByVal charged As Double, ByVal code As String) As String
    Dim t As String: t = LCase$(Trim$(typ))
    Dim o As String: o = LCase$(Trim$(outcome))
    Dim gradeTxt As String: gradeTxt = IIf(Trim$(code) = "", "no code", "grade " & Trim$(code))
    Dim posTxt As String
    Dim aboveTxt As String: aboveTxt = IIf(Trim$(code) = "", "no code", gradeTxt & " is above " & HoursText(CHARGE_MAX_GRADE))
    If t = "callout" Then
        posTxt = IIf(callOrder > 0, OrdinalText(callOrder) & " on the call list, ", "")     ' the list position (a skipped person has no row - review)
        If o = "exception" Then ChargeWhy = "Not charged: " & posTxt & "plant upset (exception)": Exit Function   ' v27 (item 24)
        If charged <> 0 Then
            ChargeWhy = "Charged " & HoursText(charged) & " h: " & posTxt & Trim$(outcome) & ", " & HoursText(hours) & _
                        " h offered x " & HoursText(CHARGE_MULTIPLIER) & " (" & gradeTxt & ")"
        ElseIf o = "accepted" Then
            ChargeWhy = "Not charged: " & posTxt & "accepted " & HoursText(hours) & " h"
        ElseIf (o = "refused" Or o = "left message") And DeptCodeValue(code) > CHARGE_MAX_GRADE Then
            ChargeWhy = "Not charged: " & posTxt & Trim$(outcome) & ", " & aboveTxt
        Else
            ChargeWhy = "Not charged: " & posTxt & IIf(Trim$(outcome) = "", "no outcome", Trim$(outcome))
        End If
    ElseIf t = "posting" Then
        posTxt = IIf(callOrder > 0, "#" & callOrder & " in call order, ", "")
        If charged <> 0 Then
            ChargeWhy = "Charged " & HoursText(charged) & " h: " & posTxt & "didn't sign up, posting hours " & HoursText(hours) & _
                        " x " & HoursText(CHARGE_MULTIPLIER) & " (" & gradeTxt & ")"
        ElseIf o = "awarded" Then
            ChargeWhy = "Not charged: " & posTxt & "awarded " & HoursText(hours) & " h"
        ElseIf o = "signed up" Then
            ChargeWhy = "Not charged: " & posTxt & "signed up, not awarded"
        ElseIf o = "refused" And DeptCodeValue(code) > CHARGE_MAX_GRADE Then
            ChargeWhy = "Not charged: " & posTxt & "didn't sign up, " & aboveTxt
        Else
            ChargeWhy = "Not charged: " & posTxt & IIf(Trim$(outcome) = "", "no outcome", Trim$(outcome))
        End If
    Else
        ChargeWhy = IIf(Trim$(typ) = "", "Entry", Trim$(typ)) & ": " & IIf(charged >= 0, "+", "") & HoursText(charged) & " h" & _
                    IIf(Trim$(outcome) = "", "", " (" & Trim$(outcome) & ")")
    End If
End Function

' v26: the sentence for row r of a log array in the OT Log layout (RangeTo2D of the table or of one
' row). withNote appends the row's Note where it says more than the fields (a posting's "didn't sign
' up (#3, awarded #5)", a call-out's partial texts), with the electronic-sign-up login cut
' (PublicNote). A VOID row says so first.
' codeNotesOnly (v26 review): append only the Notes the code itself writes (the posting and partial texts),
' never a typed one - what the members' file carries (PublicWhy); the admin views pass False.
Public Function WhyFromRow(ByRef lg As Variant, ByVal r As Long, ByVal withNote As Boolean, _
                           Optional ByVal codeNotesOnly As Boolean = False) As String
    Dim typ As String: typ = SafeStr(lg(r, L_TYPE))
    Dim outcome As String: outcome = SafeStr(lg(r, L_OUTCOME))
    Dim t As String: t = LCase$(Trim$(typ))
    If t <> "callout" And t <> "posting" Then outcome = SafeStr(lg(r, L_TASK))     ' an adjustment's type lives in Reason
    WhyFromRow = ChargeWhy(typ, outcome, CLng(NumOr0(lg(r, L_ORDER))), NumOr0(lg(r, L_BASEHRS)), _
                           NumOr0(lg(r, L_CHARGED)), SafeStr(lg(r, L_CODE)))
    If withNote Then
        Dim note As String: note = PublicNote(SafeStr(lg(r, L_REASON)))
        If note <> "" And (t = "callout" Or t = "posting") Then
            If Not codeNotesOnly Or IsCodeNote(note) Then WhyFromRow = WhyFromRow & "  (" & note & ")"
        End If
    End If
    If UCase$(Trim$(SafeStr(lg(r, L_STATUS)))) = "VOID" Then WhyFromRow = "VOID - " & WhyFromRow
End Function

' v26: a Note the code wrote (WriteOutcome / SubmitCalloutSession), by its fixed opening words
Private Function IsCodeNote(ByVal note As String) As Boolean
    Dim p As Variant
    For Each p In Array("Posting - ", "Partial award: ", "Offered ", "Took ", "Not offered: ", "Signed up for ", "Out of department")
        If StrComp(Left$(note, Len(CStr(p))), CStr(p), vbTextCompare) = 0 Then IsCodeNote = True: Exit Function
    Next p
End Function

' v26: a log Note as anyone may see it: the "[electronic sign-up <login> <time>]" marker the Award
' pop-up adds (v24.1) is cut off - a login is never shown (was Private in modPublish)
Public Function PublicNote(ByVal note As String) As String
    Dim p As Long: p = InStr(1, note, "[electronic sign-up", vbTextCompare)
    If p > 0 Then note = Trim$(Left$(note, p - 1))
    PublicNote = note
End Function

' A Period Start cell as a whole number (0 = not a date)
Private Function PeriodNumber(ByVal v As Variant) As Long
    If IsError(v) Then Exit Function
    If IsDate(v) Then
        PeriodNumber = CLng(Int(CDbl(CDate(v))))
    ElseIf IsNumeric(v) And Not IsEmpty(v) Then
        PeriodNumber = CLng(Int(CDbl(v)))
    End If
End Function


' ====================================================================
'  THE LIVE FIGURES
' ====================================================================
' The OT Log table without any message (Nothing if it isn't there)
Private Function QuietLog() As ListObject
    If Not SheetExists(LOG_SHEET) Then Exit Function
    On Error Resume Next
    Set QuietLog = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo 0
End Function

' v26: the OT Log's rows as a 2-D array in the log layout (Empty when there is no table, it is
' empty, or its layout isn't current) - the one reader the split, the report and the status bar use
Public Function LogRowsNow() As Variant
    Dim lo As ListObject: Set lo = QuietLog()
    If lo Is Nothing Then Exit Function
    If Not LogLayoutCurrent(lo) Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    LogRowsNow = RangeTo2D(lo.DataBodyRange)
End Function

' Everyone's split for this roster year, straight from the OT Log
Public Function SplitNow() As Object
    Dim lg As Variant: lg = LogRowsNow()
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim sRow As Long: sRow = PStartRow()
    ' v26.5: the split adds up to Refused (I), which counts the periods up to the current one (RefusedThrough) and
    ' never past this roster's last period (from Dec 27 the current period is next year's first)
    Dim lastStart As Long: lastStart = CLng(ws.Cells(sRow, PeriodLastCol()).Value2)
    Dim thr As Long: thr = CLng(Int(CDbl(RefusedThrough())))
    If thr > 0 And thr < lastStart Then lastStart = thr
    Set SplitNow = SplitTotals(lg, CLng(ws.Cells(sRow, C_PERIOD1).Value2), lastStart)
End Function

' v24.2: the same split for part of the year (period start numbers, both ends included)
Public Function SplitBetween(ByVal firstStart As Long, ByVal lastStart As Long) As Object
    Dim lg As Variant: lg = LogRowsNow()
    Set SplitBetween = SplitTotals(lg, firstStart, lastStart)
End Function

' v24.2: which pay periods the report covers. The box opens on "1-<current period>"
' (year to date). False = cancelled.
Private Function AskReportRange(ByRef firstStart As Long, ByRef lastStart As Long, ByRef rangeTxt As String, _
                                ByRef wholeYear As Boolean) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim nPer As Long: nPer = PeriodLastCol() - C_PERIOD1 + 1
    Dim base As Long: base = CLng(ws.Cells(PStartRow(), C_PERIOD1).Value2)
    Dim cur As Long: cur = PeriodColumn(PeriodStartFor(Date)) - C_PERIOD1 + 1
    If cur < 1 Or cur > nPer Then cur = nPer
    Dim a As Variant, t As String, dash As Long, pFrom As Long, pTo As Long, s1 As String, s2 As String
    Do
        a = Application.InputBox("Pay periods to include, as  from-to  (1 to " & nPer & ")." & vbNewLine & vbNewLine & _
                                 "Period " & cur & " is the current one, so  1-" & cur & "  is the year to date." & vbNewLine & _
                                 "For the whole year enter  1-" & nPer & ".", "Refusals Report", "1-" & cur, Type:=2)
        If VarType(a) = vbBoolean Then Exit Function
        t = Replace(Trim$(CStr(a)), " ", "")
        pFrom = 0: pTo = 0
        dash = InStr(t, "-")
        If dash > 1 Then
            s1 = Left$(t, dash - 1): s2 = Mid$(t, dash + 1)
        Else
            s1 = t: s2 = t
        End If
        If Len(s1) >= 1 And Len(s1) <= 2 And Len(s2) >= 1 And Len(s2) <= 2 Then
            If s1 Like String$(Len(s1), "#") And s2 Like String$(Len(s2), "#") Then pFrom = CLng(s1): pTo = CLng(s2)
        End If
        If pFrom >= 1 And pTo >= pFrom And pTo <= nPer Then Exit Do
        MsgBox "Enter two period numbers with a dash between them, for example  1-" & cur & ".  They run from 1 to " & nPer & ".", _
               vbExclamation, "Refusals Report"
    Loop
    firstStart = base + 14 * (pFrom - 1)
    lastStart = base + 14 * (pTo - 1)
    wholeYear = (pFrom = 1 And pTo = nPer)
    rangeTxt = "pay periods " & pFrom & " to " & pTo & " (" & Format$(CDate(firstStart), "mmm d, yyyy") & " to " & _
               Format$(CDate(lastStart + 13), "mmm d, yyyy") & ")"
    AskReportRange = True
End Function

' One person's split. False when they have nothing charged this year.
Public Function RefusedSplitOf(ByVal splitMap As Object, ByVal emp As String, ByRef callout As Double, _
                               ByRef posting As Double, ByRef other As Double) As Boolean
    callout = 0: posting = 0: other = 0
    If splitMap Is Nothing Then Exit Function
    If Not splitMap.Exists(Trim$(emp)) Then Exit Function
    Dim a As Variant: a = splitMap(Trim$(emp))
    callout = a(SP_CALLOUT): posting = a(SP_POSTING): other = a(SP_OTHER)
    RefusedSplitOf = True
End Function

' "call-out 42 / posting 21" (+ " / other 12.3" only when it isn't zero)
Public Function SplitText(ByVal callout As Double, ByVal posting As Double, ByVal other As Double) As String
    SplitText = "call-out " & HoursText(callout) & " / posting " & HoursText(posting)
    If Abs(other) > 0.0001 Then SplitText = SplitText & " / other " & HoursText(other)
End Function

' The cached map for the status bar: rebuilt when the log changes (row
' count, total charged or number of voids) or the roster year moves.
Private Function CachedSplit() As Object
    Dim lo As ListObject: Set lo = QuietLog()
    If lo Is Nothing Then Exit Function
    Dim stamp As String
    If Not lo.DataBodyRange Is Nothing Then
        stamp = lo.ListRows.Count & "|" & _
                Application.WorksheetFunction.Sum(lo.ListColumns(L_CHARGED).DataBodyRange) & "|" & _
                Application.WorksheetFunction.CountIf(lo.ListColumns(L_STATUS).DataBodyRange, "VOID")
    End If
    stamp = stamp & "|" & CStr(RosterSheet().Cells(PStartRow(), C_PERIOD1).Value2) & "|" & CStr(CDbl(RefusedThrough()))   ' v26.5: the cap moves the split
    If mSplit Is Nothing Or stamp <> mSplitStamp Then
        Set mSplit = SplitNow()
        mSplitStamp = stamp
    End If
    Set CachedSplit = mSplit
End Function

' A rename changes names without changing any hours: forget the cache
Public Sub ResetSplitCache(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    Set mSplit = Nothing
    mSplitStamp = ""
End Sub


' ====================================================================
'  STATUS BAR: the split for the selected roster row
' ====================================================================
' From ThisWorkbook.Workbook_SheetSelectionChange. Writes nothing to the
' workbook (a cell write on every click would clear Undo), only Excel's
' status bar. Must never raise or slow a click.
Public Sub RosterSelectionChanged(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Quit
    If TotalMarksExpired() Then ClearTotalMarks       ' v27 (item 8): the before / after marks come off at the first click MARK_SECONDS after a save (the one write this handler makes; the save had cleared Undo already)
    Dim txt As String, nm As String, c As Double, p As Double, o As Double
    If Not Target Is Nothing Then
        If Sh.Name = RosterSheet().Name And Target.Rows.Count = 1 And Target.Row >= R_FIRST Then
            nm = Trim$(SafeStr(Sh.Cells(Target.Row, C_NAME).Value))
            If nm <> "" Then
                RefusedSplitOf CachedSplit(), nm, c, p, o
                txt = nm & ":   Refused " & HoursText(c + p + o) & " h  =  " & SplitText(c, p, o)
            End If
        ElseIf Sh.Name = LOG_SHEET And Target.Rows.Count = 1 Then
            txt = LogRowWhy(Target.Row)                 ' v26: "why was I charged" for the selected OT Log row
        End If
    End If
    If txt <> "" Then
        Application.StatusBar = Left$(txt, 250)          ' the bar takes 255 at most (review: a long Note raised, and nothing showed)
        mBarShown = True
    ElseIf mBarShown Then
        Application.StatusBar = False
        mBarShown = False
    End If
    Exit Sub
Quit:
    On Error Resume Next
    If mBarShown Then Application.StatusBar = False: mBarShown = False
End Sub

' ====================================================================
'  v27 (item 8, review 5.2): the Total Hours breakdown - a double-click on a G cell, any login
' ====================================================================
' From ThisWorkbook's double-click through modPeople.PeopleRosterDoubleClick. A box with the figure's
' parts (Worked, Refused with its split, provisional) and the charged hours per pay period, the same
' figures the history report draws. Reads only; no rule here (the roster's formulas and RefusedBucket).
Public Sub ShowTotalBreakdown(ByVal nm As String)
    On Error GoTo Oops
    MsgBox TotalBreakdownText(nm), vbInformation, "Total Hours - " & nm
    Exit Sub
Oops:
    ReportError "ShowTotalBreakdown"
End Sub

Public Function TotalBreakdownText(ByVal nm As String) As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long: r = RosterRowOf(nm)
    If r < R_FIRST Then TotalBreakdownText = nm & " is not on the roster.": Exit Function
    Dim g As Double, h As Double, rf As Double, pv As Double
    g = NumOr0(ws.Cells(r, C_TOTAL).Value): h = NumOr0(ws.Cells(r, C_WORKED).Value)
    rf = NumOr0(ws.Cells(r, C_REFUSED).Value): pv = NumOr0(ws.Cells(r, C_PROV).Value)
    Dim co As Double, po As Double, ot As Double
    RefusedSplitOf CachedSplit(), nm, co, po, ot
    Dim dt As Date: dt = KronosDataThrough()
    Dim capAt As Date: capAt = RefusedThrough()
    Dim s As String
    s = "Total Hours " & HoursText(g) & " h  =" & vbNewLine & _
        "  Worked (Kronos" & IIf(dt > 0, " through " & Format$(dt, "mmm d"), ": nothing imported yet") & "):  " & HoursText(h) & " h" & vbNewLine & _
        "  Refused (charged, pay periods up to " & Format$(capAt, "mmm d") & "):  " & HoursText(rf) & " h  =  " & SplitText(co, po, ot) & vbNewLine & _
        "  Provisional accepted OT (logged, not yet in Kronos):  " & HoursText(pv) & " h" & vbNewLine
    If Abs(g - (h + rf + pv)) > 0.01 Then s = s & "  (the parts add to " & HoursText(h + rf + pv) & ": run IntegrityCheck)" & vbNewLine

    ' the charged hours per pay period from the log (non-void rows for this person), newest first
    Dim lg As Variant: lg = LogRowsNow()
    Dim tot As Object: Set tot = CreateObject("Scripting.Dictionary")
    Dim tc As Object: Set tc = CreateObject("Scripting.Dictionary")
    Dim tp As Object: Set tp = CreateObject("Scripting.Dictionary")
    Dim i As Long, ps As Long, ch As Double
    If IsArray(lg) Then
        For i = 1 To UBound(lg, 1)
            If StrComp(Trim$(SafeStr(lg(i, L_EMP))), nm, vbTextCompare) = 0 And IsDate(lg(i, L_PERIOD)) Then
                If UCase$(Trim$(SafeStr(lg(i, L_STATUS)))) <> "VOID" Then
                    ps = CLng(Int(CDbl(CDate(lg(i, L_PERIOD)))))
                    ch = NumOr0(lg(i, L_CHARGED))
                    If Abs(ch) < 0.0001 Then GoTo NextRow      ' (review L5) only rows that charged: an Accepted / Awarded row adds nothing here
                    If tot.Exists(ps) Then tot(ps) = tot(ps) + ch Else tot.Add ps, ch
                    Select Case RefusedBucket(SafeStr(lg(i, L_TYPE)))
                        Case "callout": If tc.Exists(ps) Then tc(ps) = tc(ps) + ch Else tc.Add ps, ch
                        Case "posting": If tp.Exists(ps) Then tp(ps) = tp(ps) + ch Else tp.Add ps, ch
                    End Select
                End If
            End If
NextRow:
        Next i
    End If
    Dim worked As Object: Set worked = WorkedPeriodsFor(nm)
    Dim ks() As Long, nk As Long, k As Variant, t As Long, j As Long
    If tot.Count > 0 Then
        ReDim ks(1 To tot.Count)
        For Each k In tot.Keys
            nk = nk + 1: ks(nk) = CLng(k)
        Next k
        For i = 2 To nk                                     ' newest first
            t = ks(i): j = i - 1
            Do While j >= 1
                If ks(j) >= t Then Exit Do
                ks(j + 1) = ks(j): j = j - 1
            Loop
            ks(j + 1) = t
        Next i
    End If
    s = s & vbNewLine & "Charged by pay period" & IIf(nk = 0, ": nothing this year.", " (newest first):") & vbNewLine
    Dim shown As Long, cc As Double, pp As Double, wTxt As String
    For i = 1 To nk
        If shown >= 6 Then s = s & "  ... and " & (nk - shown) & " more: see Employee history" & vbNewLine: Exit For   ' (review L5) MsgBox shows 1,024 characters at most
        cc = 0: pp = 0
        If tc.Exists(ks(i)) Then cc = tc(ks(i))
        If tp.Exists(ks(i)) Then pp = tp(ks(i))
        wTxt = ""
        If worked.Exists(ks(i)) Then wTxt = "  worked " & HoursText(worked(ks(i)))
        s = s & "  " & Format$(CDate(ks(i)), "mmm d") & " - " & Format$(CDate(ks(i)) + 13, "mmm d") & ":  " & HoursText(tot(ks(i))) & " h  (" & SplitText(cc, pp, tot(ks(i)) - cc - pp) & ")" & _
            IIf(CDate(ks(i)) > capAt, "  - counts from " & Format$(CDate(ks(i)), "mmm d"), "") & wTxt & vbNewLine
        shown = shown + 1
    Next i
    s = s & vbNewLine & "A charge counts in Total Hours once its shift's pay period begins. Employee history (the Desk) draws every row."
    If Len(s) > 1000 Then s = Left$(s, 990) & " ..."     ' (review L5) never cut off without saying so
    TotalBreakdownText = s
End Function

' v26: the status-bar text for one OT Log sheet row ("" when the row isn't a table row): the person,
' the ref, the shift and WhyFromRow with the Note. Reads one table row; writes nothing.
Private Function LogRowWhy(ByVal sheetRow As Long) As String
    Dim lo As ListObject: Set lo = QuietLog()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Not LogLayoutCurrent(lo) Then Exit Function
    Dim r As Long: r = sheetRow - lo.DataBodyRange.Row + 1
    If r < 1 Or r > lo.ListRows.Count Then Exit Function
    Dim lg As Variant: lg = RangeTo2D(lo.DataBodyRange.Rows(r))
    Dim nm As String: nm = Trim$(SafeStr(lg(1, L_EMP)))
    If nm = "" Then Exit Function
    Dim d As String
    If IsDate(lg(1, L_DATE)) Then d = Format$(CDate(lg(1, L_DATE)), "ddd mmm d") & " "
    LogRowWhy = nm & ":   " & SafeStr(lg(1, L_REF)) & "  " & d & SafeStr(lg(1, L_SHIFTLBL)) & ", " & SafeStr(lg(1, L_DEPT)) & _
                "  -  " & WhyFromRow(lg, 1, True)
End Function

' v25.5 (A10) "Find a person" on the bar (supervisors and admins): part of a name -> the roster row is selected
' (a unit filter that hides it is cleared first) and the status bar says where the person stands in their home
' unit's call order today - LoadEligible, the one call order, everyone holding a code in the unit. No rule here.
Public Sub FindPerson()
    On Error GoTo Oops
    If Not IsSupervisor() Then Exit Sub                 ' review: the button is theirs; Alt+F8 by a plain user does nothing
    Dim q As Variant
    q = Application.InputBox(Prompt:="Part of the name (the last name is enough):", Title:="Find a person", Type:=2)
    If VarType(q) = vbBoolean Then Exit Sub
    Dim t As String: t = Trim$(CStr(q))
    If t = "" Then Exit Sub
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim r As Long, lastRow As Long: lastRow = RosterLastRow()
    Dim hitRow As Long, nHits As Long, nm As String, firstNm As String, others As String
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(rs.Cells(r, C_NAME).Value))
        If nm <> "" Then
            If InStr(1, nm, t, vbTextCompare) > 0 Then
                nHits = nHits + 1
                If hitRow = 0 Then
                    hitRow = r: firstNm = nm
                ElseIf nHits <= 7 Then
                    others = others & IIf(others = "", "", ", ") & CalName(nm)     ' v27 (item 20, B16): the other matches, named
                End If
            End If
        End If
    Next r
    If hitRow = 0 Then
        MsgBox "Nobody on the roster matches """ & t & """.", vbInformation, "Find a person"
        Exit Sub
    End If
    If rs.Rows(hitRow).Hidden Then
        ShowAllRows                                       ' a unit filter was hiding the row - it also sorts A-Z (review)
        hitRow = 0
        For r = R_FIRST To RosterLastRow()
            If StrComp(Trim$(SafeStr(rs.Cells(r, C_NAME).Value)), firstNm, vbTextCompare) = 0 Then hitRow = r: Exit For
        Next r
        If hitRow = 0 Then MsgBox firstNm & " moved while the filter cleared - search again.", vbInformation, "Find a person": Exit Sub
    End If
    rs.Activate
    Application.Goto rs.Cells(hitRow, C_NAME), True
    Dim txt As String: txt = firstNm
    Dim hu As String: hu = HomeUnitOf(firstNm)
    If hu <> "" Then
        Dim recs() As EmpRec, n As Long, i As Long, pos As Long
        n = LoadEligible(hu, CrewFilter(CREW_ALL), recs, 999, CALLOUT_INCLUDE_DAYS, Date)
        For i = 1 To n
            If StrComp(recs(i).EmpName, firstNm, vbTextCompare) = 0 Then pos = i: Exit For
        Next i
        If pos > 0 Then
            txt = txt & ":   " & pos & " of " & n & " in the " & hu & " call order today, " & TotalText(recs(pos).TotalHrs, KronosDataThrough())   ' v26.5.5 (5.2)
        Else
            txt = txt & ":   no " & hu & " code on the roster today"
        End If
    Else
        txt = txt & ":   no home unit on the roster (column B)"
    End If
    If nHits > 1 Then txt = txt & "   (" & nHits & " names match - the first is selected; also: " & others & IIf(nHits > 7, ", ...", "") & ")"
    ShowBarText txt
    ' v27 (item 20, B16): the status cell carries it too, after the view it already shows (review MED-5: a unit filter's
    ' name stays), so it stays on screen after the next click; a sort or a view change rewrites the line
    On Error Resume Next
    Dim viewNow As String: viewNow = CurrentViewText()
    Dim pf As Long: pf = InStr(1, viewNow, " - found ", vbTextCompare)
    If pf > 0 Then viewNow = Left$(viewNow, pf - 1)
    SetRosterStatus viewNow & " - found " & CalName(firstNm) & IIf(nHits > 1, " (+" & (nHits - 1) & ")", "")
    On Error GoTo Oops
    Exit Sub
Oops:
    ReportError "FindPerson"
End Sub

' v27 (item 20, B15): "Signed in as <login>: admin" / "supervisor" / "call-outs and the calendar (ask <admin> to be
' added as a supervisor)" - the three gating levels of convention 10, in the status cell
Public Function SignedInText() As String
    Dim role As String
    If IsAdmin() Then
        role = "admin"
    ElseIf IsSupervisor() Then
        role = "supervisor"
    Else
        role = "call-outs + calendar"                     ' (review MED-6) short: the first-open box says how to be added
    End If
    SignedInText = "Signed in: " & CurrentUser() & " (" & role & ")"
End Function

' Gives Excel's status bar back if it holds our text. Called when the workbook
' loses focus or closes, and by macros that switch tabs with events off.
Public Sub ReleaseStatusBar(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    On Error Resume Next                            ' never raises
    If mBarShown Then Application.StatusBar = False: mBarShown = False
End Sub

' v25.5: a selection's text on the status bar from another module (the Calendar's click details), under the
' same flag, so it is given back the same way; "" gives it back
Public Sub ShowBarText(ByVal txt As String)
    On Error Resume Next                            ' never raises
    If txt <> "" Then
        Application.StatusBar = Left$(txt, 250)
        mBarShown = True
    ElseIf mBarShown Then
        Application.StatusBar = False
        mBarShown = False
    End If
End Sub

' Leaving a tab: give the status bar back (the text was about a selection on the tab just left; v25.5: the
' Calendar's too)
Public Sub ReportsSheetActivated(ByVal Sh As Object)
    On Error Resume Next                            ' never raises
    If Not mBarShown Then Exit Sub
    Application.StatusBar = False: mBarShown = False
End Sub


' ====================================================================
'  REFUSALS REPORT (PRINT / PUBLISH > Refusals Report)
' ====================================================================
' One row per person with charged hours this roster year: call-out and
' posting refusals (hours and how many), other, total; a sub-total per
' Home Unit. Drawn on the output-only "Refusal Report" sheet. No notes,
' no names of who logged anything.
Public Sub RefusalsReport()
    If Not RequireAdmin() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Dim unit As String, allUnits As Boolean
    If Not AskReportUnit(unit, allUnits) Then Exit Sub
    Dim firstStart As Long, lastStart As Long, rangeTxt As String, wholeYear As Boolean
    If Not AskReportRange(firstStart, lastStart, rangeTxt, wholeYear) Then Exit Sub

    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim ws As Worksheet, em As String
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- who, and their home unit ----
    Dim splitMap As Object: Set splitMap = SplitBetween(firstStart, lastStart)
    Dim homeOf As Object: Set homeOf = CreateObject("Scripting.Dictionary")
    homeOf.CompareMode = 1
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim v As Variant, i As Long, nm As String
    If lastRow >= R_FIRST Then
        v = rs.Range(rs.Cells(R_FIRST, 1), rs.Cells(lastRow, C_NAME)).Value        ' A:D, always 2-D
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, C_NAME)))
            If nm <> "" Then homeOf(nm) = Trim$(SafeStr(v(i, C_HOME)))
        Next i
    End If
    Dim n As Long, k As Variant, hu As String
    Dim names() As String, units() As String, keys() As String
    ReDim names(1 To splitMap.Count + 1): ReDim units(1 To splitMap.Count + 1): ReDim keys(1 To splitMap.Count + 1)
    For Each k In splitMap.Keys
        If homeOf.Exists(k) Then
            hu = homeOf(k)
            If hu = "" Then hu = "(no home unit)"
        Else
            hu = "(not on the roster)"
        End If
        If allUnits Or StrComp(hu, unit, vbTextCompare) = 0 Then
            n = n + 1
            names(n) = CStr(k): units(n) = hu
            keys(n) = LCase$(hu) & "|" & LCase$(CStr(k))
        End If
    Next k
    ' sort by unit, then name
    Dim j As Long, tK As String, tN As String, tU As String
    For i = 2 To n
        tK = keys(i): tN = names(i): tU = units(i): j = i - 1
        Do While j >= 1
            If keys(j) <= tK Then Exit Do
            keys(j + 1) = keys(j): names(j + 1) = names(j): units(j + 1) = units(j)
            j = j - 1
        Loop
        keys(j + 1) = tK: names(j + 1) = tN: units(j + 1) = tU
    Next i

    ' ---- the sheet ----
    Set ws = ReportSheet()
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    Const NC As Long = 8
    ws.Range("A1").Value = "REFUSED HOURS BY KIND  -  " & IIf(allUnits, "all units", unit) & "  -  " & RosterYear()
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = "Charged hours in the OT Log for " & rangeTxt & ", not void.  " & _
                           IIf(wholeYear, "Call-out + Posting + Other = Refused on the roster.  ", _
                               "Part of the year: the totals are less than Refused on the roster for anyone charged outside these periods.  ") & _
                           "Drawn " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & "."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    ws.Range("A4").Resize(1, NC).Value = Array("Home Unit", "Employee", "Call-out hours", "Call-out refusals", _
                                               "Posting hours", "Posting refusals", "Other hours", "Total refused")
    With ws.Range("A4").Resize(1, NC)
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(30, 56, 100)
        .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    ws.Rows(4).RowHeight = 30

    Dim r As Long: r = 5
    Dim a As Variant, sub1(0 To 4) As Double, tot(0 To 4) As Double, q As Long, curUnit As String
    For i = 1 To n
        If i > 1 And units(i) <> curUnit Then
            ReportTotalRow ws, r, curUnit & " total", sub1
            r = r + 1
            For q = 0 To 4
                sub1(q) = 0
            Next q
        End If
        curUnit = units(i)
        a = splitMap(names(i))
        ws.Cells(r, 1).Value = units(i)
        ws.Cells(r, 2).NumberFormat = "@"
        ws.Cells(r, 2).Value = names(i)
        ws.Cells(r, 3).Value = a(SP_CALLOUT)
        ws.Cells(r, 4).Value = a(SP_NCALLOUT)
        ws.Cells(r, 5).Value = a(SP_POSTING)
        ws.Cells(r, 6).Value = a(SP_NPOSTING)
        ws.Cells(r, 7).Value = a(SP_OTHER)
        ws.Cells(r, 8).Value = a(SP_CALLOUT) + a(SP_POSTING) + a(SP_OTHER)
        For q = 0 To 4
            sub1(q) = sub1(q) + a(q): tot(q) = tot(q) + a(q)
        Next q
        r = r + 1
    Next i
    If n > 0 Then
        ReportTotalRow ws, r, curUnit & " total", sub1
        r = r + 1
        If allUnits Then ReportTotalRow ws, r, "ALL UNITS", tot: r = r + 1
    Else
        ws.Cells(r, 1).Value = "Nobody has charged hours in these pay periods" & IIf(allUnits, ".", " in " & unit & ".")
        r = r + 1
    End If
    Dim rTotEnd As Long: rTotEnd = r - 1

    ' ---- v26: the charges behind the figures - one line per charged log row, by person in the same order ----
    If n > 0 Then
        r = r + 1
        ws.Cells(r, 1).Value = "THE CHARGES BEHIND THE FIGURES  -  every charged OT Log row in " & rangeTxt
        ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Size = 12
        r = r + 1
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("Shift date", "Ref", "Unit / shift", "Charged", "Why")
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, NC))
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(30, 56, 100)
            .VerticalAlignment = xlCenter
        End With
        r = r + 1
        Dim lg As Variant: lg = LogRowsNow()
        Dim rr As Long, per As Long, nmRows As Long
        For i = 1 To n
            ws.Cells(r, 1).Value = units(i) & "  -  " & names(i)
            ws.Cells(r, 1).Font.Bold = True
            ws.Range(ws.Cells(r, 1), ws.Cells(r, NC)).Interior.Color = RGB(236, 239, 243)
            r = r + 1
            nmRows = 0
            If IsArray(lg) Then
                For rr = 1 To UBound(lg, 1)                    ' newest first, as the log is
                    If StrComp(Trim$(SafeStr(lg(rr, L_EMP))), names(i), vbTextCompare) = 0 Then
                        If UCase$(Trim$(SafeStr(lg(rr, L_STATUS)))) <> "VOID" And NumOr0(lg(rr, L_CHARGED)) <> 0 Then
                            per = PeriodNumber(lg(rr, L_PERIOD))
                            If per >= firstStart And per <= lastStart Then
                                ws.Cells(r, 1).NumberFormat = "ddd mmm d"
                                If IsDate(lg(rr, L_DATE)) Then ws.Cells(r, 1).Value = CDate(lg(rr, L_DATE))
                                ws.Cells(r, 2).Value = SafeStr(lg(rr, L_REF))
                                ws.Cells(r, 3).Value = SafeStr(lg(rr, L_DEPT)) & " / " & SafeStr(lg(rr, L_SHIFTLBL))
                                ws.Cells(r, 4).Value = NumOr0(lg(rr, L_CHARGED))
                                ws.Cells(r, 5).Value = WhyFromRow(lg, rr, True)
                                With ws.Range(ws.Cells(r, 5), ws.Cells(r, NC))       ' output-only sheet: the sentence across E:H
                                    .Merge: .WrapText = True: .VerticalAlignment = xlTop
                                End With
                                ws.Rows(r).RowHeight = 15 * (1 + Len(CStr(ws.Cells(r, 5).Value)) \ 42)   ' E:H = 48 units wide (review)
                                nmRows = nmRows + 1
                                r = r + 1
                            End If
                        End If
                    End If
                Next rr
            End If
            If nmRows = 0 Then ws.Cells(r, 2).Value = "(no charged rows in these pay periods)": r = r + 1
        Next i
    End If

    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 28
    ws.Range(ws.Columns(3), ws.Columns(NC)).ColumnWidth = 12
    ws.Range(ws.Cells(5, 3), ws.Cells(rTotEnd, NC)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(4, 1), ws.Cells(rTotEnd, NC)).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
    If r > rTotEnd + 1 Then
        ws.Range(ws.Cells(rTotEnd + 4, 1), ws.Cells(r - 1, 1)).HorizontalAlignment = xlLeft
        ws.Range(ws.Cells(rTotEnd + 4, 4), ws.Cells(r - 1, 4)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(rTotEnd + 4, 1), ws.Cells(r - 1, NC)).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
    End If
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, NC)).Address
        .PrintTitleRows = "$4:$4"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        ' v27 (item 5, 5.3): where the figures came from
        .CenterFooter = "&8Produced by the Call Out Roster from the OT Log's charged rows (S- call-outs, P- postings, ADJ- adjustments)  -  page &P of &N"
    End With
    On Error GoTo Failed
    ProtectSheet ws
    ws.Activate
    ReleaseStatusBar                        ' the tab changed with events off
    ActiveWindow.DisplayGridlines = False
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas

    If n > 0 Then
        If MsgBox("The report is on the '" & ws.Name & "' tab (" & n & " people)." & vbNewLine & vbNewLine & _
                  "Save a PDF as well?", vbYesNo + vbQuestion, "Refusals report") = vbYes Then ReportPdf ws, "Refused hours " & IIf(allUnits, "all units", unit)
    End If
    Exit Sub
Failed:
    em = Err.Description
    LogError "RefusalsReport", Err.Number, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    MsgBox "The refusals report couldn't be drawn: " & em, vbExclamation, "Refusals report"
End Sub

Private Sub ReportTotalRow(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, ByRef t() As Double)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 3).Value = t(SP_CALLOUT)
    ws.Cells(r, 4).Value = t(SP_NCALLOUT)
    ws.Cells(r, 5).Value = t(SP_POSTING)
    ws.Cells(r, 6).Value = t(SP_NPOSTING)
    ws.Cells(r, 7).Value = t(SP_OTHER)
    ws.Cells(r, 8).Value = t(SP_CALLOUT) + t(SP_POSTING) + t(SP_OTHER)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 8))
        .Font.Bold = True
        .Interior.Color = RGB(236, 239, 243)
    End With
End Sub

Private Function ReportSheet() As Worksheet
    Set ReportSheet = ReportSheetNamed(REFUSAL_SHEET)
End Function

' v26: an output-only report sheet by name, made on first use (the Refusals Report's pattern)
Public Function ReportSheetNamed(ByVal sheetName As String) As Worksheet     ' v27 (item 23): modEvidence draws its page with it
    If Not SheetExists(sheetName) Then
        Dim was As Object: Set was = ActiveSheet
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)                ' v25.1
        On Error Resume Next
        If Not was Is Nothing Then was.Activate
        On Error GoTo 0
    End If
    Set ReportSheetNamed = ThisWorkbook.Worksheets(sheetName)
End Function


' ====================================================================
'  EMPLOYEE HISTORY REPORT (v26, kit item 1)
' ====================================================================
' One person, a date range -> the History Report sheet (+ PDF): every OT Log
' row for them in the range, newest first, with the Why line, then the
' charged hours per pay period and the total, as the members' My history
' shows them. Two flavours: FULL (admin: who logged it and when, void rows
' struck through) and MEMBER-SAFE (exactly the columns dLog carries, so it
' matches the members' file row for row). Supervisors get member-safe only.
' Reads the log; writes the output sheet. No rule.
' v27 (item 20, B16): one code-built form (frmHistory) replaces the three input boxes: the name picker, From / To
' defaulted to the roster year, Full / Member-safe for an admin in the admin view, a Save-a-PDF tick. The form
' calls RunHistoryReport; ResolvePersonName is the name match both paths share.
Public Sub EmployeeHistory()
    On Error GoTo Oops
    If Not IsSupervisor() Then Exit Sub                 ' the button is theirs (admins included)
    If LogTable() Is Nothing Then Exit Sub
    frmHistory.Show
    Exit Sub
Oops:
    ReportError "EmployeeHistory"
End Sub

' The report for the form: draws it (FULL only for an admin in the admin view - the sheet is admin-only) and,
' asked, saves it as a PDF beside the roster. Raises to its caller (the form's handler).
Public Sub RunHistoryReport(ByVal nm As String, ByVal dFrom As Date, ByVal dTo As Date, ByVal full As Boolean, ByVal pdf As Boolean)
    If full And Not (IsAdmin() And gAdminUnlocked) Then full = False
    DrawHistoryReport nm, dFrom, dTo, full
    If pdf Then
        Dim hs As Worksheet: Set hs = ThisWorkbook.Worksheets(IIf(full, HISTORY_ADMIN_SHEET, HISTORY_SHEET))
        ' (review L9) a draw that failed cleared the sheet and said so: no PDF of it
        If Left$(SafeStr(hs.Range("A1").Value), 16) = "EMPLOYEE HISTORY" Then ReportPdf hs, "History " & CalName(nm), "Employee history - " & nm, True
    End If
End Sub
' v27 (item 20): part of a name -> the roster name (the first match), or someone no longer on the roster who has
' log rows; several matches ask "Did you mean"; "" = nobody, or cancelled. The one match the form and the box share.
Public Function ResolvePersonName(ByVal t As String, ByVal ttl As String) As String
    t = Trim$(t)
    If t = "" Then Exit Function
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim r As Long, nm As String, hit As String, nHits As Long
    For r = R_FIRST To RosterLastRow()                  ' (review M7) a whole name typed or picked wins at once
        nm = Trim$(SafeStr(rs.Cells(r, C_NAME).Value))
        If StrComp(nm, t, vbTextCompare) = 0 Then ResolvePersonName = nm: Exit Function
    Next r
    For r = R_FIRST To RosterLastRow()
        nm = Trim$(SafeStr(rs.Cells(r, C_NAME).Value))
        If nm <> "" Then
            If InStr(1, nm, t, vbTextCompare) > 0 Then
                nHits = nHits + 1
                If hit = "" Then hit = nm
            End If
        End If
    Next r
    If hit = "" Then                                    ' someone who left: the log still holds their rows
        Dim lg As Variant: lg = LogRowsNow()
        Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = 1
        If IsArray(lg) Then
            For r = 1 To UBound(lg, 1)
                nm = Trim$(SafeStr(lg(r, L_EMP)))
                If nm <> "" And Not seen.Exists(nm) Then
                    seen.Add nm, True
                    If InStr(1, nm, t, vbTextCompare) > 0 Then
                        nHits = nHits + 1
                        If hit = "" Then hit = nm
                    End If
                End If
            Next r
        End If
    End If
    If hit = "" Then
        MsgBox "Nobody on the roster or in the OT Log matches """ & t & """.", vbInformation, ttl
        Exit Function
    End If
    If nHits > 1 Then
        If MsgBox(nHits & " names match """ & t & """. Did you mean " & hit & "?" & vbNewLine & vbNewLine & _
                  "(No: type more of the name.)", vbYesNo + vbQuestion, ttl) <> vbYes Then Exit Function
    End If
    ResolvePersonName = hit
End Function
' v26: the Why column as the members' file carries it: the sentence with the Note for a CHARGED row,
' "" for any other (nothing about an uncharged call or an unavailable reason leaves the master)
Public Function PublicWhy(ByRef lg As Variant, ByVal r As Long) As String
    If NumOr0(lg(r, L_CHARGED)) <> 0 Then PublicWhy = WhyFromRow(lg, r, True, True)     ' the code's own Notes only (review)
End Function

Private Sub DrawHistoryReport(ByVal nm As String, ByVal dFrom As Date, ByVal dTo As Date, ByVal full As Boolean)
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim ws As Worksheet, em As String
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- the rows: this person, newest shift date first (ties keep the log's order). FULL: shift date in the
    '      range, void rows too. MEMBER-SAFE: exactly dLog's rows (an Entry ID, not void, Period Start in the
    '      range - a provisional future shift counts in the period it was logged in, as My history shows it) ----
    Dim lg As Variant: lg = LogRowsNow()
    Dim idx() As Long, n As Long, r As Long, i As Long, j As Long, t As Long, sd As Date, keep As Boolean
    If IsArray(lg) Then
        ReDim idx(1 To UBound(lg, 1))
        For r = 1 To UBound(lg, 1)
            If StrComp(Trim$(SafeStr(lg(r, L_EMP))), nm, vbTextCompare) = 0 And IsDate(lg(r, L_DATE)) Then
                keep = False
                If full Then
                    sd = CDate(lg(r, L_DATE))
                    keep = (Int(sd) >= Int(dFrom) And Int(sd) <= Int(dTo))
                ElseIf SafeStr(lg(r, L_ENTRY)) <> "" And UCase$(SafeStr(lg(r, L_STATUS))) <> "VOID" Then   ' BuildLog's own tests
                    If PeriodNumber(lg(r, L_PERIOD)) > 0 Then                 ' a date or a serial, as BuildLog's DateNumber reads it
                        keep = (PeriodNumber(lg(r, L_PERIOD)) >= CLng(Int(dFrom)) And PeriodNumber(lg(r, L_PERIOD)) <= CLng(Int(dTo)))
                    End If
                End If
                If keep Then n = n + 1: idx(n) = r
            End If
        Next r
        For i = 2 To n
            t = idx(i): j = i - 1
            Do While j >= 1
                If Int(CDate(lg(idx(j), L_DATE))) >= Int(CDate(lg(t, L_DATE))) Then Exit Do
                idx(j + 1) = idx(j): j = j - 1
            Loop
            idx(j + 1) = t
        Next i
    End If

    ' ---- the sheet ----
    Set ws = ReportSheetNamed(IIf(full, HISTORY_ADMIN_SHEET, HISTORY_SHEET))   ' the FULL flavour (logins) on its admin-only sheet
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    Dim nc As Long: nc = IIf(full, 15, 10)                  ' v27 (item 23): the FULL report's 15th column = Evidence (scans, disputes)
    ws.Range("A1").Value = "EMPLOYEE HISTORY  -  " & nm & "  -  " & DateText(dFrom) & " to " & DateText(dTo) & IIf(full, "  (FULL, admin)", "  (member-safe)")
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = "Every OT Log entry for " & nm & IIf(full, " with a shift date in the range", " counted in a pay period starting in the range") & _
                           ", newest first" & IIf(full, ", void rows struck through (not counted)", " (void rows left out; the same rows as My history in the members' file)") & ".  " & _
                           "Why = what the row shows; the charge rules are the roster's.  Drawn " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & "."
    ws.Range("A2").Font.Italic = True: ws.Range("A2").Font.Size = 8.5
    Dim heads As Variant
    If full Then
        heads = Array("Shift date", "Ref", "Type", "Unit / shift", "Hours", "Call order", "Outcome", "Charged", "Pay period", "Why", _
                      "Logged by", "When", "Status", "Void note", "Evidence")
    Else
        heads = Array("Shift date", "Ref", "Type", "Unit / shift", "Hours", "Call order", "Outcome", "Charged", "Pay period", "Why")   ' My history's order
    End If
    ws.Range("A4").Resize(1, nc).Value = heads
    With ws.Range("A4").Resize(1, nc)
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(30, 56, 100)
        .WrapText = True: .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
    End With
    ws.Rows(4).RowHeight = 30

    Dim totals As Object: Set totals = CreateObject("Scripting.Dictionary")
    Dim totalsP As Object: Set totalsP = CreateObject("Scripting.Dictionary")     ' v26.5: posting refusals per period
    Dim totalsC As Object: Set totalsC = CreateObject("Scripting.Dictionary")     ' v26.5: call-out refusals per period
    Dim yearTot As Double, kCall As Double, kPost As Double, kOther As Double, ch As Double, ps As String, isVoid As Boolean
    r = 5
    If n = 0 Then
        ws.Cells(r, 1).Value = "Nothing logged for " & nm & " between " & DateText(dFrom) & " and " & DateText(dTo) & "."
        r = r + 1
    End If
    For i = 1 To n
        t = idx(i)
        isVoid = (UCase$(Trim$(SafeStr(lg(t, L_STATUS)))) = "VOID")
        ws.Cells(r, 1).NumberFormat = "ddd mmm d"
        ws.Cells(r, 1).Value = CDate(lg(t, L_DATE))
        ws.Cells(r, 2).Value = SafeStr(lg(t, L_REF))
        ws.Cells(r, 3).Value = SafeStr(lg(t, L_TYPE))
        ws.Cells(r, 4).Value = SafeStr(lg(t, L_DEPT)) & " / " & SafeStr(lg(t, L_SHIFTLBL))
        ws.Cells(r, 5).Value = NumOr0(lg(t, L_BASEHRS))
        If NumOr0(lg(t, L_ORDER)) > 0 Then ws.Cells(r, 6).Value = CLng(NumOr0(lg(t, L_ORDER)))
        ws.Cells(r, 7).Value = SafeStr(lg(t, L_OUTCOME))
        ch = NumOr0(lg(t, L_CHARGED))
        ws.Cells(r, 8).Value = ch
        ws.Cells(r, 9).NumberFormat = "@"
        If IsDate(lg(t, L_PERIOD)) Then ws.Cells(r, 9).Value = Format$(CDate(lg(t, L_PERIOD)), "mmm d")
        ws.Cells(r, 10).Value = IIf(full, WhyFromRow(lg, t, True), PublicWhy(lg, t))
        If full Then
            ws.Cells(r, 11).Value = SafeStr(lg(t, L_USER))
            ws.Cells(r, 12).NumberFormat = "@"                 ' text, or Excel reads it back as a date (convention 9)
            If IsDate(lg(t, L_STAMP)) Then ws.Cells(r, 12).Value = Format$(CDate(lg(t, L_STAMP)), "mmm d, yyyy h:mm AM/PM")
            ws.Cells(r, 13).Value = SafeStr(lg(t, L_STATUS))
            ws.Cells(r, 14).Value = SafeStr(lg(t, L_VOIDNOTE))
            ws.Cells(r, 15).Value = EvidenceFor(SafeStr(lg(t, L_REF)))   ' v27 (item 23)
            If isVoid Then ws.Range(ws.Cells(r, 1), ws.Cells(r, nc)).Font.Strikethrough = True
        End If
        If Not isVoid Then
            If IsDate(lg(t, L_PERIOD)) Then ps = Format$(CDate(lg(t, L_PERIOD)), "yyyy-mm-dd") Else ps = "(no period)"
            If totals.Exists(ps) Then totals(ps) = totals(ps) + ch Else totals.Add ps, ch
            yearTot = yearTot + ch
            Select Case RefusedBucket(SafeStr(lg(t, L_TYPE)))
                Case "callout": kCall = kCall + ch: If totalsC.Exists(ps) Then totalsC(ps) = totalsC(ps) + ch Else totalsC.Add ps, ch
                Case "posting": kPost = kPost + ch: If totalsP.Exists(ps) Then totalsP(ps) = totalsP(ps) + ch Else totalsP.Add ps, ch
                Case Else: kOther = kOther + ch
            End Select
        End If
        r = r + 1
    Next i
    Dim rRowsEnd As Long: rRowsEnd = r - 1

    ' ---- the totals, as the members' My history shows them ----
    r = r + 1
    ' v26.5: per pay period the total charged, the posting refusals, the call-out refusals and the premium hours worked
    ' (the Worked Periods table the Kronos export fills; "-" until one is imported), as My history shows them
    Dim worked As Object: Set worked = WorkedPeriodsFor(nm)
    Dim hasWorked As Boolean: hasWorked = (worked.Count > 0)
    Dim wk As Variant, wkTot As Double
    ' only the periods the report covers (the member-safe report takes the periods starting in the range, the
    ' full one the periods holding its shift dates), so "Total for the range" is one range
    Dim wkFrom As Double: wkFrom = CDbl(Int(dFrom))
    If full Then wkFrom = CDbl(Int(PeriodStartFor(dFrom)))
    Dim inRange As Object: Set inRange = CreateObject("Scripting.Dictionary")
    For Each wk In worked.Keys
        If CDbl(wk) >= wkFrom And CDbl(wk) <= CDbl(Int(dTo)) Then
            inRange(wk) = worked(wk)
            wkTot = wkTot + worked(wk)
            ps = Format$(CDate(CDbl(wk)), "yyyy-mm-dd")
            If Not totals.Exists(ps) Then totals.Add ps, 0#        ' a period with hours worked but nothing charged still gets a row
        End If
    Next wk
    Set worked = inRange
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("Pay period", "Charged hours", "Posting refusals", "Call-out refusals", "Worked (Kronos)")
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, nc))
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(30, 56, 100)
    End With
    r = r + 1
    Dim ks() As String, nk As Long, kk As Variant, tk As String
    If totals.Count > 0 Then
        ReDim ks(1 To totals.Count)
        For Each kk In totals.Keys
            nk = nk + 1: ks(nk) = CStr(kk)
        Next kk
    End If
    For i = 2 To nk                                     ' A-Z, then written newest first
        tk = ks(i): j = i - 1
        Do While j >= 1
            If ks(j) <= tk Then Exit Do
            ks(j + 1) = ks(j): j = j - 1
        Loop
        ks(j + 1) = tk
    Next i
    For i = nk To 1 Step -1
        If ks(i) Like "####-##-##" Then
            sd = DateSerial(CLng(Left$(ks(i), 4)), CLng(Mid$(ks(i), 6, 2)), CLng(Mid$(ks(i), 9, 2)))
            ws.Cells(r, 1).Value = Format$(sd, "mmm d") & " - " & Format$(sd + 13, "mmm d")
        Else
            ws.Cells(r, 1).Value = ks(i)
        End If
        ws.Cells(r, 2).Value = totals(ks(i))
        If totalsP.Exists(ks(i)) Then ws.Cells(r, 3).Value = totalsP(ks(i)) Else ws.Cells(r, 3).Value = 0
        If totalsC.Exists(ks(i)) Then ws.Cells(r, 4).Value = totalsC(ks(i)) Else ws.Cells(r, 4).Value = 0
        If hasWorked Then
            If ks(i) Like "####-##-##" And worked.Exists(CLng(Int(CDbl(sd)))) Then ws.Cells(r, 5).Value = worked(CLng(Int(CDbl(sd)))) Else ws.Cells(r, 5).Value = 0
        Else
            ws.Cells(r, 5).Value = "-"
        End If
        r = r + 1
    Next i
    ws.Cells(r, 1).Value = "Total for the range": ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = yearTot: ws.Cells(r, 3).Value = kPost: ws.Cells(r, 4).Value = kCall
    If hasWorked Then ws.Cells(r, 5).Value = wkTot Else ws.Cells(r, 5).Value = "-"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Font.Bold = True
    r = r + 1
    ws.Cells(r, 1).Value = "  call-out refusals": ws.Cells(r, 2).Value = kCall: r = r + 1
    ws.Cells(r, 1).Value = "  posting refusals": ws.Cells(r, 2).Value = kPost: r = r + 1
    If Abs(kOther) > 0.0001 Then ws.Cells(r, 1).Value = "  other (adjustments)": ws.Cells(r, 2).Value = kOther: r = r + 1

    ' ---- layout ----
    ws.Columns(1).ColumnWidth = 14
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 11
    ws.Columns(4).ColumnWidth = 22
    ws.Columns(5).ColumnWidth = 7
    ws.Columns(6).ColumnWidth = 7
    ws.Columns(7).ColumnWidth = 13
    ws.Columns(8).ColumnWidth = 9
    ws.Columns(9).ColumnWidth = 10
    ws.Columns(10).ColumnWidth = 70
    If full Then
        ws.Columns(11).ColumnWidth = 14
        ws.Columns(12).ColumnWidth = 20
        ws.Columns(13).ColumnWidth = 8
        ws.Columns(14).ColumnWidth = 30
        ws.Columns(15).ColumnWidth = 40              ' (review L15) the Evidence column
    End If
    If rRowsEnd >= 5 Then
        ws.Range(ws.Cells(5, 10), ws.Cells(rRowsEnd, 10)).WrapText = True
        If full Then ws.Range(ws.Cells(5, 15), ws.Cells(rRowsEnd, 15)).WrapText = True
        ws.Range(ws.Cells(5, 1), ws.Cells(rRowsEnd, nc)).VerticalAlignment = xlTop
        ws.Range(ws.Cells(5, 5), ws.Cells(rRowsEnd, 8)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(4, 1), ws.Cells(rRowsEnd, nc)).Borders(xlInsideHorizontal).Color = RGB(220, 220, 220)
        ws.Range(ws.Rows(5), ws.Rows(rRowsEnd)).AutoFit
    End If
    ws.Range(ws.Cells(rRowsEnd + 3, 2), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, nc)).Address
        .PrintTitleRows = "$4:$4"
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    On Error GoTo Failed
    ProtectSheet ws
    ws.Activate
    ReleaseStatusBar                        ' the tab changed with events off
    ActiveWindow.DisplayGridlines = False
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas

    If MsgBox("The report is on the '" & ws.Name & "' tab (" & n & " entries)." & vbNewLine & vbNewLine & _
              "Save a PDF as well?", vbYesNo + vbQuestion, "Employee history") = vbYes Then
        ReportPdf ws, "History " & nm & " " & Format$(dFrom, "yyyy-mm-dd") & " to " & Format$(dTo, "yyyy-mm-dd") & _
                      IIf(full, " FULL", ""), "Employee history", full     ' the FULL one (logins) stays beside the master (review)
    End If
    Exit Sub
Failed:
    em = Err.Description
    LogError "DrawHistoryReport", Err.Number, em
    On Error Resume Next
    If Not ws Is Nothing Then ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    MsgBox "The history report couldn't be drawn: " & em, vbExclamation, "Employee history"
End Sub

' Numbered pick: 0 = all units, 1..8 = one unit. False = cancelled.
Private Function AskReportUnit(ByRef unit As String, ByRef allUnits As Boolean) As Boolean
    Dim depts As Variant: depts = DeptList()
    Dim i As Long, menu As String, s As String, k As Long
    menu = "  0 = All units" & vbNewLine
    For i = LBound(depts) To UBound(depts)
        menu = menu & "  " & (i - LBound(depts) + 1) & " = " & depts(i) & vbNewLine
    Next i
    Do
        s = Trim$(VBA.InputBox("Refused hours by kind (call-out / posting), this roster year." & vbNewLine & vbNewLine & _
                               "Which unit? Enter the number:" & vbNewLine & vbNewLine & menu, "Refusals report", "0"))
        If s = "" Then Exit Function
        k = -1
        If s Like "#" Or s Like "##" Then k = CLng(s)        ' digits only: "$5" or "1,5" must not pick a unit
        If k >= 0 Then
            If k = 0 Then allUnits = True: AskReportUnit = True: Exit Function
            If k >= 1 And k <= UBound(depts) - LBound(depts) + 1 Then
                unit = CStr(depts(LBound(depts) + k - 1))
                AskReportUnit = True
                Exit Function
            End If
        End If
        MsgBox "Enter a number from 0 to " & (UBound(depts) - LBound(depts) + 1) & ".", vbExclamation, "Refusals report"
    Loop
End Function

' v26: fileBase is the PDF's name before the date (was fixed to "Refused hours <what>"); ttl the box title
Public Sub ReportPdf(ByVal ws As Worksheet, ByVal fileBase As String, Optional ByVal ttl As String = "Refusals report", _
                     Optional ByVal besideRoster As Boolean = False, Optional ByVal targetFolder As String = "")     ' v27: Public - the Reconcile sheet saves its page with it; item 23: a folder of the caller's
    Dim folder As String: folder = EnsureSlash(ReportFolder())
    If besideRoster Or Not FolderExists(folder) Then folder = EnsureSlash(ThisWorkbook.Path)
    If targetFolder <> "" Then folder = EnsureSlash(targetFolder)
    Dim pdfPath As String: pdfPath = folder & fileBase & " " & Format$(Date, "yyyy-mm-dd") & ".pdf"
    Dim em As String
    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then em = Err.Description
    On Error GoTo 0
    If em = "" Then
        MsgBox "PDF saved:" & vbNewLine & pdfPath, vbInformation, ttl
    Else
        MsgBox "The PDF couldn't be saved (" & em & "). Is it open?", vbExclamation, ttl
    End If
End Sub

' ====================================================================
'  THE SHIFT HANDOVER (v27 item 25, 1.4) AND "SINCE YOU WERE LAST HERE" (item 3)
' ====================================================================
' One scan, two pages on the Handover sheet: everything logged in the last HANDOVER_HOURS hours (the handover), or
' everything OTHER logins logged since this login's previous open (the Desk's tile, item 3). Three parts: the OT
' Log rows (call-outs, postings, adjustments; a void shows in Status), the time off / sick calls / leave booked,
' the shift switches entered. Supervisors and admins. No logins, no notes, leave shown as "Leave" (TypeLong);
' never published. Changes no data.
Public Sub ShiftHandover()
    On Error GoTo Oops
    If Not IsSupervisor() Then MsgBox "The shift handover is for supervisors and admins.", vbInformation, "Shift handover": Exit Sub
    Dim since As Date: since = Now - HANDOVER_HOURS / 24#
    DrawActivity since, False, "SHIFT HANDOVER  -  logged since " & Format$(since, "ddd mmm d h:mm AM/PM") & " (the last " & HANDOVER_HOURS & " h)", "Shift handover"
    Exit Sub
Oops:
    ReportError "ShiftHandover"
End Sub

' v27 (item 3): the Desk's "Since you were last here" tile - what OTHER logins logged since this login's previous open
Public Sub SinceLastHere()
    On Error GoTo Oops
    If Not IsSupervisor() Then MsgBox "This list is for supervisors and admins.", vbInformation, "Since you were last here": Exit Sub
    Dim since As Date: since = PrevOpenAt()
    If since = 0 Then
        Tell "This is the first open on record for your login: nothing to compare with yet. From the next open on, the tile counts what others logged in between.", vbInformation, "Since you were last here"
        Exit Sub
    End If
    DrawActivity since, True, "SINCE YOU WERE LAST HERE  -  logged by others since " & Format$(since, "ddd mmm d h:mm AM/PM"), "Since you were last here"
    Exit Sub
Oops:
    ReportError "SinceLastHere"
End Sub

' v27 (item 3): the tile's figure - the OT Log entries (one per Ref ID: a call-out session is one), the bookings
' and the switches logged since `since` (by other logins when othersOnly). 0 when since is 0. Never raises.
Public Function ActivitySinceCount(ByVal since As Date, ByVal othersOnly As Boolean) As Long
    On Error GoTo Quit
    If since = 0 Then Exit Function
    Dim r As Long
    ActivitySinceCount = ActivityScan(Nothing, r, since, othersOnly)
Quit:
End Function

' the page: drawn, protected, shown, a print offered
Private Sub DrawActivity(ByVal since As Date, ByVal othersOnly As Boolean, ByVal title As String, ByVal boxTitle As String)
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating     ' captured before the handler (convention 6)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Failed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim ws As Worksheet: Set ws = ReportSheetNamed(HANDOVER_SHEET)
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "'Drawn " & Format$(Now, "ddd mmm d h:mm AM/PM") & ".  " & Replace(TrustStrip(), vbLf, "   ")
    Dim r As Long: r = 4
    Dim n As Long: n = ActivityScan(ws, r, since, othersOnly)
    ws.Columns(1).ColumnWidth = 16
    ws.Columns(2).ColumnWidth = 10
    Dim i As Long
    For i = 3 To 10
        ws.Columns(i).ColumnWidth = 13
    Next i
    ws.Columns(7).ColumnWidth = 22
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = boxTitle & " - " & Format$(Now, "ddd mmm d h:mm AM/PM") & " - page &P of &N"
    End With
    ProtectSheet ws
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    ws.Activate
    If Ask(IIf(n = 0, "Nothing was logged in that time.", n & " entr" & IIf(n = 1, "y", "ies") & " on the Handover tab.") & vbNewLine & vbNewLine & _
           "Print it now?", vbYesNo + vbQuestion, boxTitle, vbNo) = vbYes Then ws.PrintOut
    Exit Sub
Failed:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    Application.EnableEvents = evWas
    Application.ScreenUpdating = scrWas
    Err.Raise en, "DrawActivity", em
End Sub

' The three parts, drawn from row r on ws (ws Nothing = count only). Returns the entries counted: one per OT Log
' Ref ID, one per booking, one per switch.
Private Function ActivityScan(ByVal ws As Worksheet, ByRef r As Long, ByVal since As Date, ByVal othersOnly As Boolean) As Long
    Dim draw As Boolean: draw = Not ws Is Nothing
    Dim me_ As String: me_ = CurrentUser()
    Dim i As Long, n As Long, st As Variant, mine As Boolean
    Dim refs As Object: Set refs = CreateObject("Scripting.Dictionary")
    refs.CompareMode = 1

    ' 1. the OT Log
    If draw Then HandoverHead ws, r, "OT LOG", Array("Logged at", "Ref", "Type", "Shift date", "Unit", "Shift", "Employee", "Outcome", "Charged", "Status")
    Dim lg As Variant: lg = LogRowsNow()
    n = 0
    If IsArray(lg) Then
        For i = 1 To UBound(lg, 1)                  ' newest first, as the log
            st = lg(i, L_STAMP)
            If IsDate(st) Then
                mine = (StrComp(Trim$(SafeStr(lg(i, L_USER))), me_, vbTextCompare) = 0)
                If CDate(st) >= since And Not (othersOnly And mine) Then
                    If draw Then HandoverRow ws, r, Array(Format$(CDate(st), "ddd h:mm AM/PM"), SafeStr(lg(i, L_REF)), SafeStr(lg(i, L_TYPE)), _
                        DayText(lg(i, L_DATE)), SafeStr(lg(i, L_DEPT)), SafeStr(lg(i, L_SHIFTLBL)), SafeStr(lg(i, L_EMP)), _
                        SafeStr(lg(i, L_OUTCOME)), HoursText(NumOr0(lg(i, L_CHARGED))), SafeStr(lg(i, L_STATUS)))
                    n = n + 1
                    refs(SafeStr(lg(i, L_REF)) & "|" & IIf(SafeStr(lg(i, L_REF)) = "", CStr(i), "")) = True
                End If
            End If
        Next i
    End If
    If draw And n = 0 Then HandoverRow ws, r, Array("Nothing logged.")
    If draw Then r = r + 1
    ActivityScan = refs.Count

    ' 2. time off, sick calls, leave
    If draw Then HandoverHead ws, r, "TIME OFF, SICK CALLS AND LEAVE BOOKED", Array("Entered at", "ID", "Employee", "Unit", "Type", "From", "To", "Status")
    n = 0
    Dim alo As ListObject: Set alo = AbsTable(False)
    Dim av As Variant
    If Not alo Is Nothing Then
        If Not alo.DataBodyRange Is Nothing Then
            av = RangeTo2D(alo.DataBodyRange)
            For i = UBound(av, 1) To 1 Step -1      ' newest last in the table: shown newest first
                st = av(i, AB_STAMP)
                If IsDate(st) Then
                    mine = (StrComp(Trim$(SafeStr(av(i, AB_USER))), me_, vbTextCompare) = 0)
                    If CDate(st) >= since And Not (othersOnly And mine) Then
                        If draw Then HandoverRow ws, r, Array(Format$(CDate(st), "ddd h:mm AM/PM"), SafeStr(av(i, AB_ID)), SafeStr(av(i, AB_EMP)), _
                            SafeStr(av(i, AB_UNIT)), TypeLong(SafeStr(av(i, AB_TYPE))), DayText(av(i, AB_START)), _
                            IIf(IsDate(av(i, AB_END)), DayText(av(i, AB_END)), "open-ended"), SafeStr(av(i, AB_STATUS)))
                        n = n + 1
                    End If
                End If
            Next i
        End If
    End If
    If draw And n = 0 Then HandoverRow ws, r, Array("Nothing booked.")
    If draw Then r = r + 1
    ActivityScan = ActivityScan + n

    ' 3. shift switches
    If draw Then HandoverHead ws, r, "SHIFT SWITCHES ENTERED", Array("Entered at", "ID", "Date", "Shift", "Unit", "Gives up", "Works", "Status")
    n = 0
    Dim slo As ListObject: Set slo = SwitchTable(False)
    Dim sv As Variant
    If Not slo Is Nothing Then
        If Not slo.DataBodyRange Is Nothing Then
            sv = RangeTo2D(slo.DataBodyRange)
            For i = UBound(sv, 1) To 1 Step -1
                st = sv(i, SW_STAMP)
                If IsDate(st) Then
                    mine = (StrComp(Trim$(SafeStr(sv(i, SW_USER))), me_, vbTextCompare) = 0)
                    If CDate(st) >= since And Not (othersOnly And mine) Then
                        If draw Then HandoverRow ws, r, Array(Format$(CDate(st), "ddd h:mm AM/PM"), SafeStr(sv(i, SW_ID)), DayText(sv(i, SW_DATE)), _
                            SafeStr(sv(i, SW_SHIFT)), SafeStr(sv(i, SW_UNIT)), SafeStr(sv(i, SW_OFF)), SafeStr(sv(i, SW_WORKS)), _
                            SafeStr(sv(i, SW_STATUS)))
                        n = n + 1
                    End If
                End If
            Next i
        End If
    End If
    If draw And n = 0 Then HandoverRow ws, r, Array("No switches entered.")
    ActivityScan = ActivityScan + n
End Function

Private Sub HandoverHead(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String, ByVal heads As Variant)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = RGB(30, 56, 100)
    r = r + 1
    Dim j As Long
    For j = LBound(heads) To UBound(heads)
        ws.Cells(r, 1 + j - LBound(heads)).Value = heads(j)
    Next j
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 1 + UBound(heads) - LBound(heads)))
        .Font.Bold = True: .Font.Size = 9
        .Interior.Color = RGB(222, 229, 240)
    End With
    r = r + 1
End Sub

Private Sub HandoverRow(ByVal ws As Worksheet, ByRef r As Long, ByVal vals As Variant)
    Dim j As Long
    For j = LBound(vals) To UBound(vals)
        ws.Cells(r, 1 + j - LBound(vals)).Value = "'" & CStr(vals(j))
    Next j
    r = r + 1
End Sub

' a date cell as "Mon Oct 12" ("" when it is not a date)
Private Function DayText(ByVal x As Variant) As String
    If IsDate(x) Then DayText = Format$(CDate(x), "ddd mmm d")
End Function
