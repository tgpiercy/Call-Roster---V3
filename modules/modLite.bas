Attribute VB_Name = "modLite"
Option Explicit

' ====================================================================
'  v29 (briefs/v29.md items 3a / 3b / 4; briefs/admin-lite.md): THE ADMIN LITE DESK'S DATA - the sources behind Ray's cards -
'  and FIX AN ENTRY's work.
'  Graeme, Oct 27, 2026: Ray (an admin on LITE_DESK_USERS) opens on a simpler Desk drawn on the FULL admin view, with
'  one action per card. This module holds what the cards show; the drawing itself (DrawLiteDesk) comes in item 3b and
'  reads only these functions. Nothing here writes to the roster or the OT Log: the three daily-check MARKS are hidden
'  names (DailyKronos / DailyNeeds / DailyPublish = "yyyy-mm-dd hh:nn|login", set by the import, Needs Attention and
'  the publish when they finish), and the counts are read from the tables the roster already keeps. Pure cores
'  (suites postexc, litemark, litedaily, litebar) are mirrored in tests/rules.py.
' ====================================================================
Public Const LITE_VERSION As String = "Lite v29.0.2 (Oct 28)"

Public Const LITE_CLOSING_HOURS As Double = 24      ' a posting with no member sign-up that closes within this many hours is an exception
Public Const LITE_STAFF_DAYS As Long = 14           ' Band 4: below minimum in the next N days
Private Const DAILY_KRONOS As String = "DailyKronos"
Private Const DAILY_NEEDS As String = "DailyNeeds"
Private Const DAILY_PUBLISH As String = "DailyPublish"

' Fix an entry (item 4): what the card shows and the re-entry needs, read by FixEntryFacts
Public Type FixFacts
    Ref As String
    Kind As String          ' the prefix: S, P, ADJ, SS, AB, SC, LV, UM, NH, MAN, OB ...
    Found As Boolean
    Rows As Long            ' live rows (1 for a switch or an absence)
    Text As String          ' the read-only view
    Unit As String
    ShiftDate As Date
    ShiftLbl As String
    Hours As Double
    Emp As String
    Partner As String       ' a switch: who works it
    DFrom As Date           ' an absence
    DTo As Date             ' 0 = open-ended
    AbsType As String       ' an absence's type (Vacation, Personal ...)
    IsLeave As Boolean      ' an absence with a leave group (LV-): re-entered through Leave, not Book Time Off
    ReqId As String         ' the Booked members' request this booking came from ("" = none): the re-entry re-books it
End Type

' ====================================================================
'  PURE
' ====================================================================
' A posting's exception for the Lite Desk's "Posting exceptions" card: "no takers" = Open, past its closing moment,
' no member sign-up (the admin decides: No takers, Cancel or Award Out of Dept - the brief's "awaiting an out-of-department
' award" IS this state seen from the admin's chair, so two words cover its three); "closing soon" = Open, not closed yet,
' closing within LITE_CLOSING_HOURS, no member sign-up; "" otherwise (any other status, a sign-up, or time to spare).
' Paper sign-ups are known only at the award, so the card says "no member sign-ups". Suite postexc.
Public Function PostingExceptionCore(ByVal status As String, ByVal closed As Boolean, ByVal hoursToClose As Double, _
                                     ByVal nSignups As Long) As String
    If Trim$(status) <> "Open" Then Exit Function
    If nSignups > 0 Then Exit Function
    If closed Then
        PostingExceptionCore = "no takers"
    ElseIf hoursToClose <= LITE_CLOSING_HOURS Then
        PostingExceptionCore = "closing soon"
    End If
End Function

' A daily-check mark ("yyyy-mm-dd hh:nn|login") read against today (ISO): "hh:nn (login)" when it was made today,
' "" when not today, blank or malformed. Suite litemark.
Public Function DailyMarkText(ByVal mark As String, ByVal todayIso As String) As String
    Dim p As Long: p = InStr(mark, "|")
    If p = 0 Then Exit Function
    Dim stamp As String: stamp = Left$(mark, p - 1)
    If Len(stamp) < 16 Then Exit Function
    If Left$(stamp, 10) <> todayIso Then Exit Function
    DailyMarkText = Mid$(stamp, 12, 5) & " (" & Mid$(mark, p + 1) & ")"
End Function

' The daily check's collapsed line: all three marks made today -> "Done today: Kronos hh:nn (who); Needs Attention
' hh:nn (who); members' file hh:nn (who)"; "" while any step is not done today (the band shows its three steps).
' Suite litedaily.
Public Function DailyCheckLineCore(ByVal todayIso As String, ByVal kronosMark As String, ByVal needsMark As String, _
                                   ByVal publishMark As String) As String
    Dim k As String, n As String, p As String
    k = DailyMarkText(kronosMark, todayIso)
    n = DailyMarkText(needsMark, todayIso)
    p = DailyMarkText(publishMark, todayIso)
    If k = "" Or n = "" Or p = "" Then Exit Function
    DailyCheckLineCore = "Done today: Kronos " & k & "; Needs Attention " & n & "; members' file " & p
End Function

' How many cells of a bar nCells wide a figure fills against the largest figure: 0 for nothing, at least 1 for
' anything, rounded half up, never past nCells. Suite litebar.
Public Function BarCellsCore(ByVal hours As Double, ByVal maxHours As Double, ByVal nCells As Long) As Long
    If hours <= 0 Or maxHours <= 0 Or nCells <= 0 Then Exit Function
    Dim x As Double: x = hours / maxHours * nCells
    BarCellsCore = Int(x + 0.5)
    If BarCellsCore < 1 Then BarCellsCore = 1
    If BarCellsCore > nCells Then BarCellsCore = nCells
End Function

' ====================================================================
'  THE DAILY CHECK MARKS (hidden names; "anyone has done it today")
' ====================================================================
Private Function LiteNameGet(ByVal key As String) As String
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then Exit Function
    LiteNameGet = Replace(Replace(Mid$(nmObj.RefersTo, 2), """", ""), "'", "")
End Function

Private Sub LiteNameSet(ByVal key As String, ByVal v As String)
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:="=""" & v & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & v & """"
    End If
End Sub

' Called when a daily-check step finishes: "kronos" (SyncKronosData, H written), "needs" (ShowNeedsAttention),
' "publish" (PublishMembers, the file saved). Writes "yyyy-mm-dd hh:nn|login"; a practice copy marks too (its Desk
' shows them; nothing else reads them). Never raises; the mark dirties the workbook so the close-time save keeps it.
Public Sub LiteMark(ByVal step As String)
    On Error Resume Next                            ' never raises
    Dim key As String
    Select Case LCase$(step)
        Case "kronos": key = DAILY_KRONOS
        Case "needs": key = DAILY_NEEDS
        Case "publish": key = DAILY_PUBLISH
        Case Else: Exit Sub
    End Select
    LiteNameSet key, Format$(Now, "yyyy-mm-dd hh:nn") & "|" & LCase$(CurrentUser())   ' (review LOW-9) the workbook is left dirty so the close saves the mark
End Sub

' One step's text for the band: "hh:nn (login)" when done today, else "". step as LiteMark's.
Public Function DailyStepText(ByVal step As String) As String
    Dim key As String
    Select Case LCase$(step)
        Case "kronos": key = DAILY_KRONOS
        Case "needs": key = DAILY_NEEDS
        Case "publish": key = DAILY_PUBLISH
        Case Else: Exit Function
    End Select
    DailyStepText = DailyMarkText(LiteNameGet(key), Format$(Date, "yyyy-mm-dd"))
End Function

' The collapsed line when all three are done today, else ""
Public Function DailyCheckLine() As String
    DailyCheckLine = DailyCheckLineCore(Format$(Date, "yyyy-mm-dd"), LiteNameGet(DAILY_KRONOS), LiteNameGet(DAILY_NEEDS), LiteNameGet(DAILY_PUBLISH))
End Function

' ====================================================================
'  POSTING EXCEPTIONS
' ====================================================================
' The open postings with no member sign-up that are past closing or close within LITE_CLOSING_HOURS, oldest closing
' first; oldestId / oldestWhat name the first (the card's button opens the Award pop-up on it). The sign-ups come from
' the Signups folder (modSignups.SignupStatusMap, as the award pop-up pre-ticks them - review HIGH-1: the list's Signed Up
' column is written only at the award). Reads tblPostings once; the folder is tested once and read once per Open posting
' that would otherwise count (an unreachable share is one wait, not one per posting). Never raises (0 on any trouble).
Public Function PostingExceptionCount(Optional ByRef oldestId As String, Optional ByRef oldestWhat As String) As Long
    On Error GoTo Quit
    oldestId = "": oldestWhat = ""
    Dim pv As Variant: pv = LiteTable(PT_TABLE)
    If Not IsArray(pv) Then Exit Function
    Dim folderOk As Boolean: folderOk = False
    If SIGNUPS_ON Then folderOk = FolderExists(SignupFolderPath())
    Dim i As Long, id As String
    Dim cl As Date, what As String, oldest As Date, n As Long, nSigned As Long
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        cl = AsDate(pv(i, PS_CLOSES))
        If id <> "" And cl > 0 And SafeStr(pv(i, PS_STATUS)) = "Open" Then
            nSigned = 0
            what = PostingExceptionCore("Open", PostingClosed(cl), (PostingCloseAt(cl) - Now) * 24, 0)
            If what <> "" Then                        ' a candidate: now the sign-ups, from the folder (the costly read, only here)
                If folderOk Then nSigned = SignupStatusMap(id, Val(SafeStr(pv(i, PS_HOURS))), cl).Count
                what = PostingExceptionCore("Open", PostingClosed(cl), (PostingCloseAt(cl) - Now) * 24, nSigned)
            End If
            If what <> "" Then
                n = n + 1
                If oldest = 0 Or cl < oldest Then oldest = cl: oldestId = id: oldestWhat = what
            End If
        End If
    Next i
    PostingExceptionCount = n
Quit:
End Function

' ====================================================================
'  THE OTHER CARDS' FIGURES
' ====================================================================
' The members' requests still waiting for their partner's Agree (the card's second number; RequestsWaitingCount is the
' first). Never raises.
Public Function RequestsPartnerCount() As Long
    On Error GoTo Quit
    Dim v As Variant: v = LiteTable(RQ_TABLE)
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, RQ_STATUS)) = RQ_ST_PARTNER Then RequestsPartnerCount = RequestsPartnerCount + 1
    Next i
Quit:
End Function

' People returning from leave inside the next nDays (the day after a leave booking's end), plus open-ended leave that
' has started - the same reading as Needs Attention's LEAVE section (modNeeds.NAGatherWork D), every unit. firstTxt =
' the soonest: "Joe Smith returns Mon Nov 9" / "Joe Smith - open-ended". Never raises.
Public Function LeaveReturningCount(ByVal nDays As Long, Optional ByRef firstTxt As String) As Long
    On Error GoTo Quit
    firstTxt = ""
    Dim v As Variant: v = AbsTableData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, sd As Date, eD As Date, soonest As Date, txt As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" And IsDate(v(i, AB_START)) Then
            sd = CDate(v(i, AB_START)): txt = ""
            If Not IsDate(v(i, AB_END)) Then
                If Int(sd) <= Date + nDays - 1 Then txt = SafeStr(v(i, AB_EMP)) & " - open-ended": eD = sd
            ElseIf IsLeaveCode(TypeCode(SafeStr(v(i, AB_TYPE)))) Then
                eD = CDate(v(i, AB_END)) + 1
                If Int(eD) >= Date And Int(eD) <= Date + nDays - 1 Then txt = SafeStr(v(i, AB_EMP)) & " returns " & Format$(eD, "ddd mmm d")
            End If
            If txt <> "" Then
                LeaveReturningCount = LeaveReturningCount + 1
                If soonest = 0 Or eD < soonest Then soonest = eD: firstTxt = txt
            End If
        End If
    Next i
Quit:
End Function

' ====================================================================
'  OT HOURS THIS PAY PERIOD BY UNIT (Band 4's bars)
' ====================================================================
' The non-void Hours Charged in the OT Log whose Period Start is the current pay period, summed per unit in DeptList's
' order (+ "(other)" for a unit name the roster does not know, only when it has hours). A 2-D array (1 To n, 1 To 2):
' unit, hours; Empty when the log is empty. Display only; never a rule. Never raises.
Public Function UnitChargedThisPeriod() As Variant
    On Error GoTo Quit
    Dim v As Variant: v = LogRowsNow()
    If Not IsArray(v) Then Exit Function
    Dim depts As Variant: depts = DeptList()
    Dim nD As Long: nD = UBound(depts) - LBound(depts) + 1
    Dim out() As Variant: ReDim out(1 To nD + 1, 1 To 2)
    Dim u As Long
    For u = 1 To nD
        out(u, 1) = CStr(depts(LBound(depts) + u - 1)): out(u, 2) = 0#
    Next u
    out(nD + 1, 1) = "(other)": out(nD + 1, 2) = 0#
    Dim p0 As Date: p0 = PeriodStartFor(Date)
    Dim i As Long, k As Long, hit As Boolean, unit As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" And Val(SafeStr(v(i, L_CHARGED))) <> 0 Then     ' (review MED-4) a negative adjustment counts too
            If Int(AsDate(v(i, L_PERIOD))) = Int(p0) Then
                unit = Trim$(SafeStr(v(i, L_DEPT))): hit = False
                For k = 1 To nD
                    If StrComp(out(k, 1), unit, vbTextCompare) = 0 Then out(k, 2) = out(k, 2) + Val(SafeStr(v(i, L_CHARGED))): hit = True: Exit For
                Next k
                If Not hit Then out(nD + 1, 2) = out(nD + 1, 2) + Val(SafeStr(v(i, L_CHARGED)))
            End If
        End If
    Next i
    If out(nD + 1, 2) = 0 Then                       ' drop "(other)" when nothing landed there
        Dim trimmed() As Variant: ReDim trimmed(1 To nD, 1 To 2)
        For u = 1 To nD
            trimmed(u, 1) = out(u, 1)
            trimmed(u, 2) = out(u, 2)
        Next u
        UnitChargedThisPeriod = trimmed
    Else
        UnitChargedThisPeriod = out
    End If
Quit:
End Function

' ====================================================================
'  FIX AN ENTRY (item 4; briefs/admin-lite.md): the guided correction - a Ref's rows read-only, Void through the
'  existing void path (VoidRef for OT Log refs - the rows marked VOID with the reason, the calendar lines released, NoSave
'  on a part-way stop -, VoidSwitch for an SS-, VoidAbsence for an AB- / SC-; an award is withdrawn through
'  WithdrawAward, never voided raw), "Re-enter it now?" through the pop-ups' Prefill, then the
'  Feedback rows for the old Ref marked resolved with the new one ("void" when not re-entered). frmFixEntry is the card;
'  this is its work. Nothing new writes to the roster or the log: the voids and the pop-ups' own saves do.
' ====================================================================
' (the FixFacts Type sits at the top of the module: a Type must come before the first procedure)

Public Function FixEntryKind(ByVal ref As String) As String
    Dim p As Long: p = InStr(ref, "-")
    If p > 1 Then FixEntryKind = UCase$(Left$(ref, p - 1))
End Function

' Everything the card shows and the re-entry needs, read-only. Found = live rows exist for the Ref. Never raises.
Public Function FixEntryFacts(ByVal ref As String) As FixFacts
    Dim f As FixFacts
    On Error GoTo Quit
    f.Ref = UCase$(Trim$(ref)): f.Kind = FixEntryKind(f.Ref)
    Dim v As Variant, i As Long, t As String, who As Object, k As Variant
    Select Case f.Kind
        Case "SS"
            v = TableData(SwitchTable(False))
            If IsArray(v) Then
                For i = 1 To UBound(v, 1)
                    If UCase$(Trim$(SafeStr(v(i, SW_ID)))) = f.Ref Then
                        f.Found = (UCase$(SafeStr(v(i, SW_STATUS))) <> "VOID"): f.Rows = 1
                        f.Emp = SafeStr(v(i, SW_OFF)): f.Partner = SafeStr(v(i, SW_WORKS)): f.Unit = SafeStr(v(i, SW_UNIT))
                        f.ShiftDate = AsDate(v(i, SW_DATE)): f.ShiftLbl = SafeStr(v(i, SW_SHIFT))
                        f.ReqId = RequestIdForRef(f.Ref)
                        t = "Shift switch " & f.Ref & IIf(f.Found, "", " (already VOID)") & vbNewLine & _
                            "  " & f.Emp & " gives up " & Format$(f.ShiftDate, "ddd mmm d") & " " & f.ShiftLbl & " (" & f.Unit & "); " & f.Partner & " works it." & vbNewLine & _
                            IIf(SafeStr(v(i, SW_PAYFOR)) <> "", "  Payback for " & SafeStr(v(i, SW_PAYFOR)) & " (a re-entry comes back as a NEW switch: pick the payback link in the pop-up)" & vbNewLine, "") & _
                            IIf(SafeStr(v(i, SW_PAIDBY)) <> "", "  Paid back by " & SafeStr(v(i, SW_PAIDBY)) & vbNewLine, "") & _
                            IIf(f.ReqId <> "", "  Booked from members' request " & f.ReqId & " (a re-entry re-books it)." & vbNewLine, "") & _
                            "  Changes nobody's Total Hours; it changes who is free and the 16 h rule's reading."
                        Exit For
                    End If
                Next i
            End If
        Case "AB", "SC"
            v = AbsTableData()
            If IsArray(v) Then
                For i = 1 To UBound(v, 1)
                    If UCase$(Trim$(SafeStr(v(i, 1)))) = f.Ref Then
                        f.Found = (UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID"): f.Rows = 1
                        f.Emp = SafeStr(v(i, AB_EMP)): f.Unit = SafeStr(v(i, AB_UNIT)): f.AbsType = SafeStr(v(i, AB_TYPE))
                        f.DFrom = AsDate(v(i, AB_START)): f.DTo = AsDate(v(i, AB_END))
                        If UBound(v, 2) >= AB_GROUP Then f.IsLeave = (Trim$(SafeStr(v(i, AB_GROUP))) <> "")
                        f.ReqId = RequestIdForRef(f.Ref)
                        t = IIf(f.Kind = "SC", "Sick day ", IIf(f.IsLeave, "Leave ", "Time off ")) & f.Ref & IIf(f.Found, "", " (already VOID)") & vbNewLine & _
                            "  " & f.Emp & "  " & f.AbsType & "  " & Format$(f.DFrom, "ddd mmm d") & _
                            IIf(f.DTo > 0, " - " & Format$(f.DTo, "ddd mmm d"), " - open-ended") & " (" & f.Unit & ")" & vbNewLine & _
                            IIf(f.ReqId <> "", "  Booked from members' request " & f.ReqId & IIf(f.Kind = "SC", " (re-book it by hand: the request reads Booking voided once no sick day of it is left).", " (a re-entry re-books it).") & vbNewLine, "") & _
                            IIf(f.Kind = "SC", "  A sick day is re-entered through Someone called in sick (today or later) or Book Time Off > Sick (a past day)." & vbNewLine, "") & _
                            IIf(f.IsLeave, "  A leave is re-entered through Leave: going (its group and return top-up live there)." & vbNewLine, "") & _
                            "  Changes nobody's Total Hours; its calendar lines and covers stay (release them on the Calendar if the cover was wrong)."
                        Exit For
                    End If
                Next i
            End If
        Case Else
            v = LogRowsNow()
            If IsArray(v) Then
                Set who = CreateObject("Scripting.Dictionary"): who.CompareMode = vbTextCompare
                For i = 1 To UBound(v, 1)
                    If UCase$(Trim$(SafeStr(v(i, L_REF)))) = f.Ref And UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then
                        f.Rows = f.Rows + 1
                        If f.Rows = 1 Then
                            f.Unit = SafeStr(v(i, L_DEPT)): f.ShiftDate = AsDate(v(i, L_DATE)): f.ShiftLbl = SafeStr(v(i, L_SHIFTLBL))
                            f.Emp = SafeStr(v(i, L_EMP))
                            t = SafeStr(v(i, L_TYPE)) & " " & f.Ref & "  -  " & Format$(f.ShiftDate, "ddd mmm d") & "  " & f.Unit & "  " & f.ShiftLbl & _
                                "{HOURS}" & IIf(SafeStr(v(i, L_TASK)) <> "", "  (" & SafeStr(v(i, L_TASK)) & ")", "") & vbNewLine
                        End If
                        ' v29.0.3 (review HIGH 1.3): the SESSION's hours on offer = the most any live row shows, reading the "of M h" a
                        ' partial or capped row's Note names (modStats.OfferedOnRow) - the first row's Hours is ONE person's figure (the
                        ' hours a partial acceptor took, a 16 h cap, 0 below a partial), and a re-entry pre-filled with it charged wrong
                        If OfferedOnRow(v(i, L_BASEHRS), SafeStr(v(i, L_TASK))) > f.Hours Then f.Hours = OfferedOnRow(v(i, L_BASEHRS), SafeStr(v(i, L_TASK)))
                        If f.Rows <= 25 Then t = t & "  " & SafeStr(v(i, L_EMP)) & "  " & SafeStr(v(i, L_OUTCOME)) & _
                                                IIf(Val(SafeStr(v(i, L_CHARGED))) <> 0, "  " & SignedHours(Val(SafeStr(v(i, L_CHARGED)))) & " charged", "") & vbNewLine
                        who(SafeStr(v(i, L_EMP))) = who(SafeStr(v(i, L_EMP))) + Val(SafeStr(v(i, L_CHARGED)))
                    End If
                Next i
                If f.Rows > 25 Then t = t & "  ...and " & (f.Rows - 25) & " more" & vbNewLine
                t = Replace(t, "{HOURS}", IIf(f.Hours > 0, "  " & HoursText(f.Hours) & " h on offer", ""))     ' v29.0.3: the session's offer, known once every row is read
                f.Found = (f.Rows > 0)
                If f.Found Then
                    t = t & "What it changed (Total Hours; the void takes it back):" & vbNewLine
                    Dim anyHit As Boolean
                    For Each k In who.Keys
                        If who(k) <> 0 Then t = t & "  " & k & "  " & SignedHours(CDbl(who(k))) & vbNewLine: anyHit = True
                    Next k
                    If Not anyHit Then t = t & "  nobody's (no hours charged)" & vbNewLine
                    ' (review MED-8) what the void does NOT undo, by kind - the warnings Void Entry and Undo last give
                    Select Case f.Kind
                        Case "NH": t = t & "  NOTE: voiding a new hire's opening rows leaves the person on the roster with no opening hours; Remove them through Change a person if they never started." & vbNewLine
                        Case "UM": t = t & "  NOTE: voiding a unit move's rows leaves the person in the new unit; move them back through Change a person if the move was wrong." & vbNewLine
                        Case "LV": t = t & "  NOTE: voiding a leave return's top-up leaves the leave closed; book the return again from Leave: returning only if the leave itself is re-opened." & vbNewLine
                        Case "OB", "MAN": t = t & "  NOTE: an opening balance / typed-cell row - ask " & PRIMARY_ADMIN_NAME & " before voiding it." & vbNewLine
                    End Select
                End If
            End If
    End Select
    If Not f.Found And t = "" Then t = "No live entry with the Ref " & f.Ref & " (a typo, or already voided)."
    f.Text = t
Quit:
    FixEntryFacts = f
End Function

' The void, by kind, through the existing paths; msg says what happened. False when nothing was voided. The P- case is
' refused here: an award is withdrawn (the refusals stay charged) - the card points at Withdraw an award.
Public Function FixEntryVoid(ByRef f As FixFacts, ByVal why As String, ByRef msg As String) As Boolean
    msg = ""
    If Trim$(why) = "" Then msg = "A reason is required.": Exit Function
    If Not f.Found Then msg = "Nothing live to void for " & f.Ref & ".": Exit Function
    Select Case f.Kind
        Case "P"
            msg = "An award is not voided here: use Withdraw an award / edit a posting (the refusals stay charged, the line is released, then a call-out is offered)."
        Case "SS"
            If VoidSwitch(f.Ref, why) Then FixEntryVoid = True: msg = "Switch " & f.Ref & " is VOID."
        Case "AB", "SC"
            If VoidAbsence(f.Ref, why) Then FixEntryVoid = True: msg = "Absence " & f.Ref & " is VOID (its calendar lines stay)."
        Case Else
            Dim calN As Long, n As Long
            n = VoidRef(f.Ref, why, calN)
            If n > 0 Then FixEntryVoid = True: msg = n & " row" & IIf(n = 1, "", "s") & " of " & f.Ref & " VOID" & IIf(calN > 0, "; " & calN & " calendar line" & IIf(calN = 1, "", "s") & " released", "") & "."
        End Select
    If Not FixEntryVoid And msg = "" Then msg = "The void did not go through (the box said why)."
End Function

' "Re-enter it now?": the right pop-up, pre-filled from the voided facts. Returns the Ref ID the re-entry got ("" when
' nothing was saved, or the kind has no pop-up to pre-fill: the message names the button to use). A booking that came from
' a members' request (f.ReqId) rides gRequestId into the pop-up, so its save re-books the request with the new Ref
' (RequestBooked) instead of leaving it "Booking voided" (review MED-7). The pop-ups take "Fix:<ref>" as the request id
' and word their note and caption for a re-entry (review MED-4).
Public Function FixEntryReenter(ByRef f As FixFacts, ByRef msg As String) As String
    Dim rec0 As String: rec0 = gReceipt.Recorded
    msg = ""
    gRequestId = ""
    Select Case f.Kind
        Case "S"
            If LogTable() Is Nothing Then msg = "The OT Log could not be read.": Exit Function
            Load frmOvertimeEntry
            frmOvertimeEntry.Prefill f.Unit, f.ShiftDate, f.ShiftLbl, "Re-entered after " & f.Ref, IIf(f.Hours > 0, f.Hours, 0), ""
            frmOvertimeEntry.Show
            RefreshCalendar
        Case "ADJ"
            Dim r As Long: r = RosterRowOf(f.Emp)
            If r = 0 Then msg = f.Emp & " is not on the roster: Add Adjustment from the roster by hand.": Exit Function
            RosterSheet().Activate
            RosterSheet().Cells(r, C_NAME).Select
            AddAdjustment
        Case "SS"
            If SwitchTable(True) Is Nothing Then msg = "The Shift Switches table could not be read.": Exit Function
            gRequestId = f.ReqId
            frmShiftSwitch.Prefill f.Emp, f.Partner, f.ShiftDate, f.ShiftLbl, f.Unit, "Fix:" & f.Ref
            frmShiftSwitch.Show
            gRequestId = ""
        Case "AB"
            If f.IsLeave Then msg = "A leave is re-entered through Leave: going (the pop-up carries its group and the return top-up).": Exit Function
            If AbsTable(True) Is Nothing Then msg = "The Absences table could not be read.": Exit Function
            If CovTable(True) Is Nothing Then msg = "The Coverage table could not be read.": Exit Function
            gRequestId = f.ReqId
            frmTimeOff.Prefill f.Emp, f.DFrom, f.DTo, "Fix:" & f.Ref, f.AbsType
            frmTimeOff.Show
            gRequestId = ""
            RefreshCalendar
        Case "SC"
            msg = "A sick day is re-entered through Someone called in sick (today or a later shift) or Book Time Off > Sick (a past day)."
            Exit Function
        Case "LV"
            msg = "Re-enter the leave through Leave: going or returning (no pre-fill for a leave return)."
            Exit Function
        Case "UM", "NH"
            msg = "Re-enter it through Change a person (the Employee pop-up: a move or a new hire)."
            Exit Function
        Case Else
            msg = "No pop-up re-enters a " & f.Kind & "- entry; " & PRIMARY_ADMIN_NAME & " enters it by hand."
            Exit Function
    End Select
    If gReceipt.Recorded <> rec0 Then FixEntryReenter = FirstRefIn(gReceipt.Recorded)
    If FixEntryReenter = "" Then msg = "Nothing was saved from the pop-up."
End Function

' the Feedback rows for the old Ref: resolved with the new Ref, or "void" when it was not re-entered ("checked" /
' "already void" when the card was closed without a void: the entry was looked at). Saved (review LOW-14).
Public Sub FixEntryDone(ByVal oldRef As String, ByVal newRef As String)
    On Error Resume Next                            ' never raises
    If ResolveFeedback(oldRef, IIf(Trim$(newRef) = "", "void", Trim$(newRef))) > 0 Then SaveIfWritable
End Sub

' "+12 h" / "-5 h", through the one hours formatter (review LOW-12)
Private Function SignedHours(ByVal h As Double) As String
    SignedHours = IIf(h < 0, "-", "+") & HoursText(Abs(h)) & " h"
End Function

' Alt+F8 / the Lite Desk's Corrections card: the card
Public Sub ShowFixEntry()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    frmFixEntry.Show
    Exit Sub
Oops:
    ReportError "ShowFixEntry"
End Sub

' ====================================================================
'  HELPERS
' ====================================================================
' a table anywhere in the workbook as a 2-D array; Empty when missing or empty. Never raises.
Private Function LiteTable(ByVal tableName As String) As Variant
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(tableName)
        If Not lo Is Nothing Then LiteTable = TableData(lo): Exit Function
    Next ws
End Function

Private Function SafeBool(ByVal v As Variant) As Boolean
    On Error Resume Next                            ' never raises
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then SafeBool = v: Exit Function
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    SafeBool = (s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Or s = "X")
End Function
