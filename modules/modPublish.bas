Attribute VB_Name = "modPublish"
Option Explicit

' ====================================================================
'  PUBLISH THE MEMBERS' WORKBOOK (M1)
'  The Member View button (PublishMemberView in modOTLog) runs
'  PublishMembers:
'   1. works every table out here in the master, as VALUES: who's on
'      the roster and their call position per unit, who's free for each
'      unit / date / shift for the next MEMBER_NEXT_DAYS days, calendar
'      lines with leave masked, open postings and their lists, this
'      year's non-void log rows (no notes unless PUBLISH_REASONS);
'   2. checks nothing private is in them (no "@", no phone-like numbers,
'      no raw leave codes) and stops if anything is;
'   3. opens the template next to this workbook, fills its data sheets,
'      runs the members' own redraw macro, saves it as MEMBER_FILE_NAME in
'      PUBLISH_FOLDER and closes it.
'  The master is never changed, apart from the hidden name that records
'  when it last published. The members' file never reads the master.
' ====================================================================

Public Const PUB_VERSION As String = "Publish v22 (Oct 13)"

Public Const DS_META As String = "dMeta"
Public Const DS_ROSTER As String = "dRoster"
Public Const DS_NEXT As String = "dNext"
Public Const DS_CAL As String = "dCal"
Public Const DS_POST As String = "dPost"
Public Const DS_PLIST As String = "dPostList"
Public Const DS_LOG As String = "dLog"

Private Const ROSTER_FIXED As Long = 7        ' Name, Crew, Home Unit, Total, Worked, Refused, Provisional
Private Const CAL_MONTHS_BACK As Long = 1
Private Const CAL_MONTHS_AHEAD As Long = 3
Private Const CAP_ROWS As Long = 60000      ' tables are sized once (ReDim Preserve cannot grow the first dimension)


' ---------------- when it was last published ----------------
Public Function MembersPublishedAt() As Date
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(PUB_NAME)
    On Error GoTo 0
    If nmObj Is Nothing Then Exit Function
    Dim x As Double: x = Val(Mid$(nmObj.RefersTo, 2))
    If x > 0 Then MembersPublishedAt = CDate(x)
End Function

Private Sub SetPublishedAt(ByVal d As Date)
    Dim ref As String: ref = "=" & Replace(Format$(CDbl(d), "0.00000"), ",", ".")
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(PUB_NAME)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=PUB_NAME, RefersTo:=ref, Visible:=False
    Else
        nmObj.RefersTo = ref
    End If
    On Error GoTo 0
End Sub

Public Function MemberTemplatePath() As String
    MemberTemplatePath = EnsureSlash(ThisWorkbook.Path) & MEMBER_TEMPLATE_NAME
End Function

Public Function MemberTemplateExists() As Boolean
    On Error Resume Next
    MemberTemplateExists = (Dir$(MemberTemplatePath()) <> "")
End Function

' SystemCheck line
Public Function MembersStatusText() As String
    Dim d As Date: d = MembersPublishedAt()
    If d = 0 Then
        MembersStatusText = "Members file: never published"
    Else
        MembersStatusText = "Members file: published " & Format$(d, "mmm d, h:mm AM/PM")
    End If
End Function


' ====================================================================
'  PUBLISH
' ====================================================================
Public Sub PublishMembers()
    If Not RequireAdmin() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not MemberTemplateExists() Then
        MsgBox "The members' template wasn't found:" & vbNewLine & MemberTemplatePath() & vbNewLine & vbNewLine & _
               "Build it once (see the M1 install notes: a blank workbook, import modMember.bas, paste its ThisWorkbook " & _
               "code, run BuildMemberTemplate, save as that name next to this workbook).", vbExclamation, "Member View"
        Exit Sub
    End If
    Dim folder As String: folder = EnsureSlash(PUBLISH_FOLDER)
    If Not FolderExists(folder) Then
        MsgBox "The members' folder can't be reached:" & vbNewLine & folder & vbNewLine & vbNewLine & _
               "Check the drive, or PUBLISH_FOLDER in modOvertimeEntry.", vbExclamation, "Member View"
        Exit Sub
    End If

    Dim t0 As Single: t0 = Timer
    Application.Calculate
    Application.ScreenUpdating = False
    Application.StatusBar = "Publishing the members' file: working out call order and who's free ..."

    ' ---- 1. the tables, as values ----
    Dim meta As Variant, roster As Variant, nxt As Variant, cal As Variant
    Dim post As Variant, plist As Variant, lg As Variant
    Dim nR As Long, nN As Long, nC As Long, nP As Long, nPL As Long, nL As Long
    On Error GoTo Failed
    meta = BuildMeta()
    roster = BuildRoster(nR)
    nxt = BuildNext(nN)
    cal = BuildCal(nC)
    BuildPostings post, nP, plist, nPL
    lg = BuildLog(nL)
    On Error GoTo 0

    ' ---- 2. privacy check (nothing has been opened or written yet) ----
    Dim why As String
    why = PrivateDataIn(roster, DS_ROSTER)
    If why = "" Then why = PrivateDataIn(nxt, DS_NEXT)
    If why = "" Then why = PrivateDataIn(cal, DS_CAL)
    If why = "" Then why = PrivateDataIn(post, DS_POST)
    If why = "" Then why = PrivateDataIn(plist, DS_PLIST)
    If why = "" Then why = PrivateDataIn(lg, DS_LOG)
    If why = "" Then why = LeaveCodesIn(cal)
    If why <> "" Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "NOT published. Something private would have gone into the members' file:" & vbNewLine & vbNewLine & why & _
               vbNewLine & vbNewLine & "Fix that entry (or send me this message) and publish again.", vbCritical, "Member View"
        Exit Sub
    End If
    why = OrderSelfCheck(roster)
    If why <> "" Then
        Application.StatusBar = False
        Application.ScreenUpdating = True
        MsgBox "NOT published. The call positions worked out for the members' file don't match the roster's own call order:" & _
               vbNewLine & vbNewLine & why & vbNewLine & vbNewLine & _
               "Nothing was written. Run SystemCheck and SelfTest, then send me this message.", vbCritical, "Member View"
        Exit Sub
    End If

    ' ---- 3. fill the template and save it as the members' file ----
    Application.StatusBar = "Publishing the members' file: writing " & MEMBER_FILE_NAME & " ..."
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim secWas As Long: secWas = Application.AutomationSecurity
    Dim wb As Workbook, em As String, savedAs As String, redrawn As Boolean
    On Error GoTo OpenFailed
    Application.EnableEvents = False               ' the template's own open code must not run here
    Application.AutomationSecurity = 1             ' msoAutomationSecurityLow: its macros load, so MemberRefreshAll can run
    Set wb = Workbooks.Open(Filename:=MemberTemplatePath(), ReadOnly:=True, UpdateLinks:=0)
    Application.AutomationSecurity = secWas
    On Error GoTo WriteFailed
    If Not HasDataSheets(wb) Then
        Err.Raise vbObjectError + 601, , "the template has no data sheets - open it and run BuildMemberTemplate"
    End If
    WriteTable wb, DS_META, meta
    WriteTable wb, DS_ROSTER, roster
    WriteTable wb, DS_NEXT, nxt
    WriteTable wb, DS_CAL, cal
    WriteTable wb, DS_POST, post
    WriteTable wb, DS_PLIST, plist
    WriteTable wb, DS_LOG, lg

    ' the members' own code draws Home and the Roster (so the file reads right even with macros off);
    ' it is saved on the Roster, A-Z, so a macros-off open lands on something useful
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!MemberRefreshAll"
    redrawn = (Err.Number = 0)
    Err.Clear
    wb.Worksheets("Roster").Activate
    Err.Clear
    On Error GoTo WriteFailed

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.SaveAs Filename:=folder & MEMBER_FILE_NAME, FileFormat:=52    ' 52 = .xlsm
    If Err.Number = 0 Then
        savedAs = folder & MEMBER_FILE_NAME
    Else
        em = Err.Description
        Err.Clear
        savedAs = folder & Replace(MEMBER_FILE_NAME, ".xlsm", "") & " " & Format$(Now, "yyyy-mm-dd hhmm") & ".xlsm"
        wb.SaveAs Filename:=savedAs, FileFormat:=52
        If Err.Number <> 0 Then em = em & " / " & Err.Description: savedAs = ""
        Err.Clear
    End If
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = evWas
    On Error GoTo 0
    ThisWorkbook.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True

    If savedAs = "" Then
        MsgBox "The members' file couldn't be saved:" & vbNewLine & em & vbNewLine & vbNewLine & _
               "Is " & MEMBER_FILE_NAME & " open on a plant PC? Close it there and publish again.", vbCritical, "Member View"
        Exit Sub
    End If
    SetPublishedAt Now
    On Error Resume Next
    MarkCalDirty                                   ' Needs Attention's "members file" line
    SaveIfWritable                                 ' keeps the published-at stamp
    On Error GoTo 0
    Dim msg As String
    msg = "Members' file published (" & Format$(Timer - t0, "0") & " s):" & vbNewLine & savedAs & vbNewLine & vbNewLine & _
          nR & " people, " & nN & " who's-next rows, " & nC & " calendar lines, " & nP & " open posting(s), " & nL & " log rows." & _
          IIf(redrawn, "", vbNewLine & vbNewLine & "Its views couldn't be redrawn (macros didn't run in the template); " & _
          "members will see them refresh when they open it with macros on.")
    If savedAs <> folder & MEMBER_FILE_NAME Then
        msg = msg & vbNewLine & vbNewLine & "NOTE: " & MEMBER_FILE_NAME & " is open somewhere (" & em & "), so this copy was " & _
              "saved with the date in its name. Members still see the old one until it's closed and you publish again."
    End If
    MsgBox msg, vbInformation, "Member View"
    Exit Sub

Failed:
    em = Err.Description
    On Error Resume Next
    EndSwitchBatch
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "NOT published - couldn't work out the members' data: " & em & vbNewLine & vbNewLine & _
           "Nothing was opened or written. Run SystemCheck, then try again; if it happens twice, send me this message.", _
           vbCritical, "Member View"
    Exit Sub
OpenFailed:
    em = Err.Description
    Application.AutomationSecurity = secWas
    Application.EnableEvents = evWas
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "NOT published - the template couldn't be opened: " & em & vbNewLine & MemberTemplatePath() & vbNewLine & vbNewLine & _
           "If the template is open in Excel, close it and publish again. If it's missing, rebuild it (M1 install notes, part B).", _
           vbCritical, "Member View"
    Exit Sub
WriteFailed:
    em = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.EnableEvents = evWas
    Application.AutomationSecurity = secWas
    ThisWorkbook.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "NOT published - stopped while filling the template: " & em & vbNewLine & vbNewLine & _
           "The template was opened read-only and closed without saving, so the members' file is unchanged. " & _
           "If the message mentions data sheets, open the template and run BuildMemberTemplate again; otherwise send me this message.", _
           vbCritical, "Member View"
End Sub

' The call positions in dRoster must be the master's own call order. Re-run
' it for the first unit with people and compare (cheap; catches a wiring slip).
Private Function OrderSelfCheck(ByRef roster As Variant) As String
    Dim depts As Variant: depts = DeptList()
    Dim u As Long, recs() As EmpRec, n As Long, i As Long, r As Long, posCol As Long, found As Boolean
    For u = 0 To UBound(depts)
        n = LoadEligible(CStr(depts(u)), "ALL", recs)
        If n > 0 Then
            posCol = ROSTER_FIXED + 2 * u + 2
            For i = 1 To n
                found = False
                For r = 2 To UBound(roster, 1)
                    If StrComp(SafeStr(roster(r, 1)), recs(i).EmpName, vbTextCompare) = 0 Then
                        found = True
                        If NumOr0(roster(r, posCol)) <> i Then
                            OrderSelfCheck = DS_ROSTER & ": " & recs(i).EmpName & " is #" & i & " in " & depts(u) & _
                                             " call order but the table says #" & SafeStr(roster(r, posCol))
                            Exit Function
                        End If
                        Exit For
                    End If
                Next r
                If Not found Then OrderSelfCheck = DS_ROSTER & ": " & recs(i).EmpName & " missing": Exit Function
            Next i
            Exit Function                              ' one unit is enough
        End If
    Next u
End Function


' ====================================================================
'  THE TABLES (first row = headers)
' ====================================================================
Private Function BuildMeta() As Variant
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim ps As Date: ps = PeriodStartFor(Date)
    Dim yFrom As Variant, yTo As Variant
    yFrom = ws.Cells(PStartRow(), C_PERIOD1).Value
    yTo = ws.Cells(PEndRow(), C_PERIODN).Value
    Dim a(1 To 16, 1 To 2) As Variant
    a(1, 1) = "Key": a(1, 2) = "Value"
    a(14, 1) = "Posting close time": a(14, 2) = POSTING_CLOSE_TIME
    a(15, 1) = "Crews": a(15, 2) = Join(ShiftList(), "|")
    a(16, 1) = "Pay period days": a(16, 2) = 14
    a(2, 1) = "Published at": a(2, 2) = Now
    a(3, 1) = "Published by": a(3, 2) = CurrentUser()
    a(4, 1) = "Master version": a(4, 2) = EXPECTED_RELEASE & " M1  (" & CORE_VERSION & ")"
    a(5, 1) = "Kronos data date": a(5, 2) = KronosDataDate()
    a(6, 1) = "Kronos data through": a(6, 2) = KronosDataThrough()
    a(7, 1) = "Pay period start": a(7, 2) = ps
    a(8, 1) = "Pay period end": a(8, 2) = ps + 13
    a(9, 1) = "Roster year from": a(9, 2) = yFrom
    a(10, 1) = "Roster year to": a(10, 2) = yTo
    a(11, 1) = "Who's next days": a(11, 2) = MEMBER_NEXT_DAYS
    a(12, 1) = "Units": a(12, 2) = Join(DeptList(), "|")
    a(13, 1) = "Charge rule": a(13, 2) = "grade " & CHARGE_MAX_GRADE & " and below, x" & CHARGE_MULTIPLIER
    BuildMeta = a
End Function

' Name, Crew, Home Unit, Total, Worked, Refused, Provisional, then per unit: Code, Position
Private Function BuildRoster(ByRef nR As Long) As Variant
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim depts As Variant: depts = DeptList()
    Dim nU As Long: nU = UBound(depts) + 1
    Dim nCols As Long: nCols = ROSTER_FIXED + 2 * nU
    Dim provLive As Boolean: provLive = ProvisionalLive()

    ' call position per unit: name -> rank (the master's own order)
    Dim pos() As Object, u As Long, recs() As EmpRec, n As Long, i As Long
    ReDim pos(0 To nU - 1)
    For u = 0 To nU - 1
        Set pos(u) = CreateObject("Scripting.Dictionary")
        pos(u).CompareMode = 1
        n = LoadEligible(CStr(depts(u)), "ALL", recs)
        For i = 1 To n
            pos(u)(recs(i).EmpName) = i
        Next i
    Next u

    Dim v As Variant, r As Long, nm As String, k As Long
    If lastRow >= R_FIRST Then v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, UnitColLast())).Value
    Dim out() As Variant: ReDim out(1 To lastRow - R_FIRST + 2, 1 To nCols)
    out(1, 1) = "Name": out(1, 2) = "Crew": out(1, 3) = "Home Unit": out(1, 4) = "Total"
    out(1, 5) = "Worked": out(1, 6) = "Refused": out(1, 7) = "Provisional"
    For u = 0 To nU - 1
        out(1, ROSTER_FIXED + 2 * u + 1) = depts(u) & " code"
        out(1, ROSTER_FIXED + 2 * u + 2) = depts(u) & " #"
    Next u
    k = 1
    If IsArray(v) Then
        For r = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(r, C_NAME)))
            If nm <> "" Then
                k = k + 1                       ' out() is sized to the roster above, so no cap test is needed here
                out(k, 1) = nm
                out(k, 2) = Trim$(SafeStr(v(r, C_SHIFT)))
                out(k, 3) = Trim$(SafeStr(v(r, C_HOME)))
                out(k, 4) = NumOr0(v(r, C_TOTAL))
                out(k, 5) = NumOr0(v(r, C_WORKED))
                out(k, 6) = NumOr0(v(r, C_REFUSED))
                out(k, 7) = IIf(provLive, NumOr0(v(r, C_PROV)), 0)
                For u = 0 To nU - 1
                    out(k, ROSTER_FIXED + 2 * u + 1) = Trim$(SafeStr(v(r, DeptCol(CStr(depts(u))))))
                    If pos(u).Exists(nm) Then out(k, ROSTER_FIXED + 2 * u + 2) = pos(u)(nm) Else out(k, ROSTER_FIXED + 2 * u + 2) = ""
                Next u
            End If
        Next r
    End If
    nR = k - 1
    BuildRoster = out
End Function

' Unit, Date, Shift, Rank, Name, Total, Provisional, Code - everyone free, in call order
Private Function BuildNext(ByRef nN As Long) As Variant
    Dim depts As Variant: depts = DeptList()
    Dim cap As Long: cap = CAP_ROWS
    Dim out() As Variant: ReDim out(1 To cap, 1 To 8)
    out(1, 1) = "Unit": out(1, 2) = "Date": out(1, 3) = "Shift": out(1, 4) = "Rank"
    out(1, 5) = "Name": out(1, 6) = "Total": out(1, 7) = "Provisional": out(1, 8) = "Code"
    Dim k As Long: k = 1
    RefreshCalCache
    BeginSwitchBatch
    On Error GoTo Done
    Dim u As Long, base() As EmpRec, recs() As EmpRec, n As Long, m As Long, i As Long, dd As Long
    Dim d As Date, sh As Variant, removed As String
    For u = 0 To UBound(depts)
        n = LoadEligible(CStr(depts(u)), "ALL", base)
        If n > 0 Then
            For dd = 0 To MEMBER_NEXT_DAYS - 1
                d = Date + dd
                For Each sh In Array("Days", "Nights")
                    recs = base
                    m = FilterNotFree(recs, n, d, CStr(sh), removed, True)
                    For i = 1 To m
                        k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
                        out(k, 1) = depts(u): out(k, 2) = d: out(k, 3) = CStr(sh): out(k, 4) = i
                        out(k, 5) = recs(i).EmpName: out(k, 6) = recs(i).TotalHrs
                        out(k, 7) = recs(i).ProvHrs: out(k, 8) = recs(i).DeptCode
                    Next i
                Next sh
            Next dd
        End If
    Next u
Done:
    EndSwitchBatch
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
    nN = k - 1
    BuildNext = Trimmed(out, k, 8)
End Function

' Date, Shift, Unit, Left text, Right text, Type, Open - per unit; plus one holiday row per date (unit "*")
Private Function BuildCal(ByRef nC As Long) As Variant
    Dim depts As Variant: depts = DeptList()
    Dim d0 As Date, d1 As Date
    d0 = DateSerial(Year(Date), Month(Date) - CAL_MONTHS_BACK, 1)
    d1 = DateSerial(Year(Date), Month(Date) + CAL_MONTHS_AHEAD + 1, 1) - 1
    Dim cap As Long: cap = CAP_ROWS
    Dim out() As Variant: ReDim out(1 To cap, 1 To 7)
    out(1, 1) = "Date": out(1, 2) = "Shift": out(1, 3) = "Unit": out(1, 4) = "Left"
    out(1, 5) = "Right": out(1, 6) = "Type": out(1, 7) = "Open"
    Dim k As Long: k = 1
    Dim d As Date, u As Long, sh As Variant, slots() As CalSlot, n As Long, i As Long, hl As String
    RefreshCalCache
    For d = d0 To d1
        hl = HolidayLabel(d)
        If hl <> "" Then
            k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
            out(k, 1) = d: out(k, 2) = "": out(k, 3) = "*": out(k, 4) = hl: out(k, 5) = "": out(k, 6) = "HOL": out(k, 7) = False
        End If
        For u = 0 To UBound(depts)
            For Each sh In Array("Days", "Nights")
                n = SlotsFor(d, CStr(sh), CStr(depts(u)), slots, True)
                For i = 1 To n
                    k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
                    out(k, 1) = d: out(k, 2) = CStr(sh): out(k, 3) = depts(u)
                    out(k, 4) = slots(i).LeftText: out(k, 5) = slots(i).RightText
                    out(k, 6) = PublicType(slots(i).TypeCode): out(k, 7) = slots(i).IsOpen
                Next i
            Next sh
        Next u
    Next d
    nC = k - 1
    BuildCal = Trimmed(out, k, 7)
End Function

' leave of any kind is one public code; the members' calendar colours it "Leave"
Private Function PublicType(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": PublicType = "LV"
        Case Else: PublicType = UCase$(Trim$(code))            ' SIC stays: "Sick" is shown by decision (Oct 13)
    End Select
End Function

' Open postings: ID, Unit, Date, Shift, Hours, Positions, Reason, Closes;
' their lists: ID, Rank, Name, Hours posted
Private Sub BuildPostings(ByRef post As Variant, ByRef nP As Long, ByRef plist As Variant, ByRef nPL As Long)
    Dim ids As Variant, choices As Variant
    choices = OpenPostingChoices(ids)
    Dim capP As Long: capP = CAP_ROWS
    Dim capL As Long: capL = CAP_ROWS
    Dim p() As Variant: ReDim p(1 To capP, 1 To 8)
    Dim l() As Variant: ReDim l(1 To capL, 1 To 4)
    p(1, 1) = "Posting ID": p(1, 2) = "Unit": p(1, 3) = "Date": p(1, 4) = "Shift"
    p(1, 5) = "Hours": p(1, 6) = "Positions": p(1, 7) = "Reason": p(1, 8) = "Closes"
    l(1, 1) = "Posting ID": l(1, 2) = "Rank": l(1, 3) = "Name": l(1, 4) = "Hours posted"
    Dim kP As Long: kP = 1
    Dim kL As Long: kL = 1
    If IsArray(ids) Then
        Dim i As Long, j As Long, pr As PostRec, blank As PostRec, lst() As PListRec, n As Long
        For i = LBound(ids) To UBound(ids)
            pr = blank                              ' LoadPosting leaves pr untouched when the id isn't found
            n = LoadPosting(CStr(ids(i)), pr, lst)
            If pr.ID <> "" Then                     ' a posting with an empty list still shows (as "list not published")
                kP = kP + 1: If kP > capP Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " postings - raise CAP_ROWS in modPublish"
                p(kP, 1) = pr.ID: p(kP, 2) = pr.Unit: p(kP, 3) = pr.ShiftDate: p(kP, 4) = pr.ShiftLbl
                p(kP, 5) = pr.Hours: p(kP, 6) = pr.Positions: p(kP, 7) = pr.Reason: p(kP, 8) = pr.ClosesOn
                For j = 1 To n
                    kL = kL + 1: If kL > capL Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " posting list rows - raise CAP_ROWS in modPublish"
                    l(kL, 1) = pr.ID: l(kL, 2) = lst(j).Rank: l(kL, 3) = lst(j).EmpName: l(kL, 4) = lst(j).TotalHrs
                Next j
            End If
        Next i
    End If
    nP = kP - 1: nPL = kL - 1
    post = Trimmed(p, kP, 8)
    plist = Trimmed(l, kL, 4)
End Sub

' This year's non-void log rows: Shift Date, Ref ID, Type, Unit, Shift, Hours, Employee,
' Outcome, Hours Charged, Period Start (+ Note when PUBLISH_REASONS)
Private Function BuildLog(ByRef nL As Long) As Variant
    Dim lo As ListObject: Set lo = LogTable()
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim yFrom As Double, yTo As Double
    yFrom = NumOr0(ws.Cells(PStartRow(), C_PERIOD1).Value2)
    yTo = NumOr0(ws.Cells(PEndRow(), C_PERIODN).Value2)
    Dim nCols As Long: nCols = IIf(PUBLISH_REASONS, 11, 10)
    Dim cap As Long: cap = CAP_ROWS
    Dim out() As Variant: ReDim out(1 To cap, 1 To nCols)
    out(1, 1) = "Shift Date": out(1, 2) = "Ref ID": out(1, 3) = "Type": out(1, 4) = "Unit": out(1, 5) = "Shift"
    out(1, 6) = "Hours": out(1, 7) = "Employee": out(1, 8) = "Outcome": out(1, 9) = "Hours Charged": out(1, 10) = "Period Start"
    If PUBLISH_REASONS Then out(1, 11) = "Note"
    Dim k As Long: k = 1
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
            Dim r As Long, sd As Double
            For r = 1 To UBound(v, 1)
                If SafeStr(v(r, L_ENTRY)) <> "" And UCase$(SafeStr(v(r, L_STATUS))) <> "VOID" Then
                    sd = NumOr0(v(r, L_DATE))
                    If (yFrom = 0 Or sd >= yFrom) And (yTo = 0 Or sd <= yTo) Then
                        k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
                        out(k, 1) = v(r, L_DATE): out(k, 2) = SafeStr(v(r, L_REF)): out(k, 3) = SafeStr(v(r, L_TYPE))
                        out(k, 4) = SafeStr(v(r, L_DEPT)): out(k, 5) = SafeStr(v(r, L_SHIFTLBL)): out(k, 6) = NumOr0(v(r, L_BASEHRS))
                        out(k, 7) = SafeStr(v(r, L_EMP)): out(k, 8) = SafeStr(v(r, L_OUTCOME))
                        out(k, 9) = NumOr0(v(r, L_CHARGED)): out(k, 10) = v(r, L_PERIOD)
                        If PUBLISH_REASONS Then out(k, 11) = SafeStr(v(r, L_REASON))
                    End If
                End If
            Next r
        End If
    End If
    nL = k - 1
    BuildLog = Trimmed(out, k, nCols)
End Function

' first n rows of a (bigger) 2-D array
Private Function Trimmed(ByRef a As Variant, ByVal n As Long, ByVal nCols As Long) As Variant
    Dim out() As Variant: ReDim out(1 To n, 1 To nCols)
    Dim i As Long, j As Long
    For i = 1 To n
        For j = 1 To nCols
            out(i, j) = a(i, j)
        Next j
    Next i
    Trimmed = out
End Function


' ====================================================================
'  PRIVACY CHECK - the tables above never copy phones, emails, notes,
'  leave types, void notes or who logged. This catches a slip anyway.
' ====================================================================
Private Function PrivateDataIn(ByRef a As Variant, ByVal tableName As String) As String
    If Not IsArray(a) Then Exit Function
    Dim i As Long, j As Long, s As String
    For i = 2 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            If VarType(a(i, j)) = vbString Then
                s = a(i, j)
                If InStr(s, "@") > 0 Then
                    PrivateDataIn = tableName & " row " & i & ": looks like an email (" & Left$(s, 40) & ")"
                    Exit Function
                End If
                If LongestDigitRun(s) >= 7 Then
                    PrivateDataIn = tableName & " row " & i & ": looks like a phone number (" & Left$(s, 40) & ")"
                    Exit Function
                End If
            End If
        Next j
    Next i
End Function

' Phone-like = a single run of 7+ digits, or 10+ digits across groups joined
' by phone separators (space - ( ) .): "905-555-1234" and "(905) 555 1234"
' are caught; "0700-1900", "07:00-15:00" and "P-0012" are not. Dates written
' yyyy-mm-dd are blanked first so they never count.
Private Function LongestDigitRun(ByVal s As String) As Long
    Dim joined As Long: joined = JoinedDigits(s)
    Dim single As Long: single = SingleRun(s)
    If joined >= 10 Then
        LongestDigitRun = joined
    Else
        LongestDigitRun = single
    End If
End Function

' longest run of digits with no separator at all
Private Function SingleRun(ByVal s As String) As Long
    Dim p As Long, ch As String, run As Long, best As Long
    For p = 1 To Len(s)
        ch = Mid$(s, p, 1)
        If ch Like "#" Then
            run = run + 1
            If run > best Then best = run
        Else
            run = 0
        End If
    Next p
    SingleRun = best
End Function

' longest run of digits counting across phone separators
Private Function JoinedDigits(ByVal s As String) As Long
    Dim p As Long
    p = 1
    Do While p <= Len(s) - 9
        If Mid$(s, p, 10) Like "####-##-##" Then
            s = Left$(s, p - 1) & "          " & Mid$(s, p + 10)
            p = p + 10
        Else
            p = p + 1
        End If
    Loop
    Dim t As String, ch As String, run As Long, best As Long
    For p = 1 To Len(s)
        ch = Mid$(s, p, 1)
        If ch Like "#" Then
            run = run + 1
            If run > best Then best = run
        ElseIf ch = " " Or ch = "-" Or ch = "(" Or ch = ")" Or ch = "." Then
            ' separators inside a phone number: keep counting
        Else
            run = 0
        End If
    Next p
    JoinedDigits = best
End Function

Private Function LeaveCodesIn(ByRef cal As Variant) As String
    If Not IsArray(cal) Then Exit Function
    Dim i As Long, t As String
    For i = 2 To UBound(cal, 1)
        t = UCase$(SafeStr(cal(i, 6)))
        If IsLeaveCode(t) Then
            LeaveCodesIn = DS_CAL & " row " & i & ": leave type " & t & " would be visible"
            Exit Function
        End If
        t = UCase$(SafeStr(cal(i, 4)))
        If InStr(t, "(STD") > 0 Or InStr(t, "(PAR") > 0 Or InStr(t, "(LOA") > 0 Or InStr(t, "(WSIB") > 0 Or InStr(t, "(LTD") > 0 Or InStr(t, "(OLV") > 0 Then
            LeaveCodesIn = DS_CAL & " row " & i & ": leave type in the text"
            Exit Function
        End If
    Next i
End Function


' ====================================================================
'  WRITING INTO THE TEMPLATE
' ====================================================================
Private Function HasDataSheets(ByVal wb As Workbook) As Boolean
    Dim nm As Variant, sh As Object
    For Each nm In Array(DS_META, DS_ROSTER, DS_NEXT, DS_CAL, DS_POST, DS_PLIST, DS_LOG)
        Set sh = Nothing
        On Error Resume Next
        Set sh = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If sh Is Nothing Then Exit Function
    Next nm
    HasDataSheets = True
End Function

Private Sub WriteTable(ByVal wb As Workbook, ByVal sheetName As String, ByRef a As Variant)
    Dim ws As Worksheet: Set ws = wb.Worksheets(sheetName)
    ws.Unprotect Password:=MEMBER_PASSWORD
    ws.Cells.ClearContents
    If IsArray(a) Then ws.Range("A1").Resize(UBound(a, 1), UBound(a, 2)).Value = a
    ws.Protect Password:=MEMBER_PASSWORD, UserInterfaceOnly:=True
End Sub
