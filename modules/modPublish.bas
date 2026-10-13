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

Public Const PUB_VERSION As String = "Publish v24.3.1 (Oct 14)"

Public Const DS_META As String = "dMeta"
Public Const DS_ROSTER As String = "dRoster"
Public Const DS_NEXT As String = "dNext"
Public Const DS_CAL As String = "dCal"
Public Const DS_POST As String = "dPost"
Public Const DS_PLIST As String = "dPostList"
Public Const DS_LOG As String = "dLog"
Public Const DS_ATT As String = "dAttend"     ' v24.1 / M2: sick time COUNTS only (no dates)
Private Const DIRTY_NAME As String = "MembersDirty"   ' hidden name: postings or people changed since the last publish
Private Const MEMBER_NEEDS As String = "Members M2.1" ' v24.2: the members' code this roster's tables are laid out for

' Name, Crew, Home Unit, Total, Worked, Refused, Provisional, Refused call-out, Refused posting, Refused other
Private Const ROSTER_FIXED As Long = 10
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

' ---------------- behind since the last publish? (v24.1) ----------------
' Set when a posting is created / closed / reopened or a person is added,
' moved, renamed or removed; cleared by a successful publish.
Public Sub MarkMembersDirty(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim nmObj As Name
    Set nmObj = ThisWorkbook.Names(DIRTY_NAME)
    If nmObj Is Nothing Then ThisWorkbook.Names.Add Name:=DIRTY_NAME, RefersTo:="=TRUE", Visible:=False
End Sub

Private Sub ClearMembersDirty()
    On Error Resume Next
    ThisWorkbook.Names(DIRTY_NAME).Delete
End Sub

Public Function MembersDirty() As Boolean
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(DIRTY_NAME)
    On Error GoTo 0
    MembersDirty = Not nmObj Is Nothing
End Function

' From ThisWorkbook.Workbook_BeforeClose: an admin closing the live roster
' with postings or people changed since the last publish is asked ONCE
' whether to publish. Never blocks the close; a failed publish reports itself.
Public Sub PublishReminderOnClose()
    On Error GoTo Quit
    If Not IsLiveRoster() Then Exit Sub         ' v24.3.1: read-only, non-admin, a practice copy, a backup or frozen copy
    If NoSave() Then Exit Sub
    If Not MembersDirty() Then Exit Sub
    If Not MemberTemplateExists() Then Exit Sub
    If MsgBox("Postings or people changed since the members' file was last published" & _
              IIf(MembersPublishedAt() > 0, " (" & Format$(MembersPublishedAt(), "ddd mmm d, h:mm AM/PM") & ")", "") & "." & vbNewLine & vbNewLine & _
              "Publish it now, before closing?" & IIf(SIGNUPS_ON, vbNewLine & "(Members can only sign up for postings that are in their file.)", ""), _
              vbYesNo + vbQuestion, "Member View") = vbYes Then PublishMembers
Quit:
End Sub

' PUBLISH_FOLDER with "." (the roster's own folder) and ".\sub" resolved
Public Function PublishFolderPath() As String
    ' v24.2: a practice copy (COPY / DRY RUN in its name) only ever uses its own folder, wherever
    ' PUBLISH_FOLDER points - so its publish and its sign-up inbox can't reach the real ones
    If IsDryRunName(ThisWorkbook.Name) Then
        PublishFolderPath = EnsureSlash(ThisWorkbook.Path)
        Exit Function
    End If
    Dim pf As String: pf = Trim$(PUBLISH_FOLDER)
    If pf = "" Or pf = "." Then
        pf = ThisWorkbook.Path
    ElseIf Left$(pf, 2) = ".\" Then
        pf = EnsureSlash(ThisWorkbook.Path) & Mid$(pf, 3)
    End If
    PublishFolderPath = EnsureSlash(pf)
End Function

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
    ' v24.3.1 (review M2): yesterday's backup sits beside the template; opened to look something up,
    ' Member View from it would overwrite the live members' file with old hours stamped today.
    ' A practice copy is NOT refused, even one made from a FINAL or "Backup before" file (its name then
    ' carries both words): it publishes "PRACTICE - ..." beside itself (PublishFolderPath).
    If IsBackupCopyName(ThisWorkbook.Name) And Not IsDryRunName(ThisWorkbook.Name) Then
        MsgBox "This file is a backup or frozen copy of the roster, not the live roster:" & vbNewLine & _
               ThisWorkbook.Name & vbNewLine & vbNewLine & _
               "Member View won't publish from it. Open the live roster and publish there.", vbExclamation, "Member View"
        Exit Sub
    End If
    If LogTable() Is Nothing Then Exit Sub
    If Not MemberTemplateExists() Then
        MsgBox "The members' template wasn't found:" & vbNewLine & MemberTemplatePath() & vbNewLine & vbNewLine & _
               "Build it once (see the M1 install notes: a blank workbook, import modMember.bas, paste its ThisWorkbook " & _
               "code, run BuildMemberTemplate, save as that name next to this workbook).", vbExclamation, "Member View"
        Exit Sub
    End If
    Dim folder As String: folder = PublishFolderPath()
    ' v24.2: a practice copy writes "PRACTICE - ..." beside itself (see PublishFolderPath), never the members' file
    Dim outName As String: outName = IIf(IsDryRunName(ThisWorkbook.Name), "PRACTICE - ", "") & MEMBER_FILE_NAME
    If gSmokeRun Then                              ' v24.2: SmokeTest publishes a scratch copy, never the members' file
        gSmokeSaved = ""
        If Trim$(gSmokeFolder) = "" Then Exit Sub
        folder = EnsureSlash(gSmokeFolder)
    End If
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
    Dim post As Variant, plist As Variant, lg As Variant, att As Variant
    Dim nR As Long, nN As Long, nC As Long, nP As Long, nPL As Long, nL As Long, nA As Long
    On Error GoTo Failed
    meta = BuildMeta()
    roster = BuildRoster(nR)
    nxt = BuildNext(nN)
    cal = BuildCal(nC)
    BuildPostings post, nP, plist, nPL
    lg = BuildLog(nL)
    att = BuildAttend(nA)
    If SIGNUPS_ON And Not gSmokeRun Then EnsureSignupFolder
    On Error GoTo 0

    ' ---- 2. privacy check (nothing has been opened or written yet) ----
    Dim why As String
    why = PrivateDataIn(roster, DS_ROSTER)
    If why = "" Then why = PrivateDataIn(nxt, DS_NEXT)
    If why = "" Then why = PrivateDataIn(cal, DS_CAL)
    If why = "" Then why = PrivateDataIn(post, DS_POST)
    If why = "" Then why = PrivateDataIn(plist, DS_PLIST)
    If why = "" Then why = PrivateDataIn(lg, DS_LOG)
    If why = "" Then why = LeaveCodesIn(cal, post, lg)
    If why = "" Then why = PrivateDataIn(att, DS_ATT)
    If why = "" Then why = AttendPrivacy(att)
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
    Application.StatusBar = "Publishing the members' file: writing " & outName & " ..."
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim secWas As Long: secWas = Application.AutomationSecurity
    Dim wb As Workbook, em As String, savedAs As String, redrawn As Boolean
    Dim tplVer As String, tplLayout As String
    On Error GoTo OpenFailed
    Application.EnableEvents = False               ' the template's own open code must not run here
    Application.AutomationSecurity = 1             ' msoAutomationSecurityLow: its macros load, so MemberRefreshAll can run
    Set wb = Workbooks.Open(Filename:=MemberTemplatePath(), ReadOnly:=True, UpdateLinks:=0)
    Application.AutomationSecurity = secWas
    On Error GoTo WriteFailed
    If Not HasDataSheets(wb) Then
        Err.Raise vbObjectError + 601, , "the template is older than this roster (it has no " & DS_ATT & _
                                         " sheet): update the members' template (INSTALL note, members' part)"
    End If
    ' v24.2: the template's code and the version that laid its sheets out must both be what this roster fills
    On Error Resume Next
    tplVer = CStr(Application.Run("'" & wb.Name & "'!MemberVersion"))
    tplLayout = CStr(Application.Run("'" & wb.Name & "'!MemberLayoutVersion"))
    On Error GoTo WriteFailed
    If InStr(1, tplVer, MEMBER_NEEDS, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 604, , "the template's code is " & _
                                         IIf(tplVer = "", "older than " & MEMBER_NEEDS & " (or its macros couldn't be run)", tplVer) & _
                                         " and this roster needs " & MEMBER_NEEDS & ": update the members' template (INSTALL note, members' part)"
    End If
    If StrComp(tplLayout, tplVer, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 605, , "the template's code is " & tplVer & " but its sheets were built by " & _
                                         IIf(tplLayout = "", "an older version", tplLayout) & _
                                         ": open the template, run BuildMemberTemplate, save it, close it and publish again"
    End If
    WriteTable wb, DS_META, meta
    WriteTable wb, DS_ROSTER, roster
    WriteTable wb, DS_NEXT, nxt
    WriteTable wb, DS_CAL, cal
    WriteTable wb, DS_POST, post
    WriteTable wb, DS_PLIST, plist
    WriteTable wb, DS_LOG, lg
    WriteTable wb, DS_ATT, att

    ' the members' own code draws Home and the Roster (so the file reads right even with macros off);
    ' it is saved on the Roster, A-Z, so a macros-off open lands on something useful
    On Error Resume Next
    Application.Run "'" & wb.Name & "'!MemberRefreshAll"
    redrawn = (Err.Number = 0)
    Err.Clear
    Dim tplNote As String                          ' v24.2: the members' code writes a line on Home when its redraw fails
    tplNote = Trim$(CStr(wb.Worksheets("Home").Range("B5").Value))
    Err.Clear
    wb.Worksheets("Roster").Activate
    Err.Clear
    On Error GoTo WriteFailed
    If tplNote <> "" Then Err.Raise vbObjectError + 606, , "the members' file reported a problem while it was redrawn: " & tplNote

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.SaveAs Filename:=folder & outName, FileFormat:=52    ' 52 = .xlsm
    If Err.Number = 0 Then
        savedAs = folder & outName
    Else
        em = Err.Description
        Err.Clear
        savedAs = folder & Replace(outName, ".xlsm", "") & " " & Format$(Now, "yyyy-mm-dd hhmm") & ".xlsm"
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
               "Is " & outName & " open on a plant PC? Close it there and publish again.", vbCritical, "Member View"
        Exit Sub
    End If
    If gSmokeRun Then                              ' the scratch copy is written: nothing about the real publish changes
        gSmokeSaved = savedAs
        Exit Sub
    End If
    SetPublishedAt Now
    If savedAs = folder & outName Then ClearMembersDirty     ' a dated copy (file open on a plant PC) leaves members behind
    On Error Resume Next
    MarkCalDirty                                   ' Needs Attention's "members file" line
    SaveIfWritable                                 ' keeps the published-at stamp
    On Error GoTo 0
    Dim msg As String
    msg = "Members' file published (" & Format$(Timer - t0, "0") & " s), " & tplVer & ":" & vbNewLine & savedAs & vbNewLine & vbNewLine & _
          nR & " people, " & nN & " who's-next rows, " & nC & " calendar lines, " & nP & " open posting(s), " & nL & " log rows." & _
          IIf(redrawn, "", vbNewLine & vbNewLine & "Its views couldn't be redrawn (macros didn't run in the template); " & _
          "members will see them refresh when they open it with macros on.")
    If savedAs <> folder & outName Then
        msg = msg & vbNewLine & vbNewLine & "NOTE: " & outName & " is open somewhere (" & em & "), so this copy was " & _
              "saved with the date in its name. Members still see the old one until it's closed and you publish again."
    End If
    MsgBox msg, vbInformation, "Member View"
    Exit Sub

Failed:
    em = Err.Description
    LogError "PublishMembers", Err.Number, em
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
    LogError "PublishMembers", Err.Number, em
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
    ' 601 - 605 are refusals that say what to do (an old template): not errors for the log
    If Err.Number < vbObjectError + 601 Or Err.Number > vbObjectError + 605 Then LogError "PublishMembers", Err.Number, em
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
    yTo = ws.Cells(PEndRow(), PeriodLastCol()).Value
    Dim a(1 To 20, 1 To 2) As Variant
    a(1, 1) = "Key": a(1, 2) = "Value"
    a(14, 1) = "Posting close time": a(14, 2) = POSTING_CLOSE_TIME
    a(15, 1) = "Crews": a(15, 2) = Join(ShiftList(), "|")
    a(16, 1) = "Pay period days": a(16, 2) = 14
    a(2, 1) = "Published at": a(2, 2) = Now
    a(3, 1) = "Published by": a(3, 2) = ""                 ' v24.1: a Windows login is never published (nothing reads this row)
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
    ' v24.1 / M2
    a(17, 1) = "Sign-ups": a(17, 2) = IIf(SIGNUPS_ON, "on", "off")
    a(18, 1) = "Signup folder": a(18, 2) = IIf(SIGNUPS_ON, SignupFolderPath(), "")
    a(19, 1) = "Calendar groups": a(19, 2) = Join(CalGroups(), ";")
    a(20, 1) = "Template needs": a(20, 2) = "M2"
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
    out(1, 8) = "Refused call-out": out(1, 9) = "Refused posting": out(1, 10) = "Refused other"
    Dim splitMap As Object: Set splitMap = SplitNow()      ' v24.1: what Refused is made of
    Dim rc As Double, rp As Double, ro As Double
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
                out(k, 2) = CrewOn(nm, Date, Trim$(SafeStr(v(r, C_SHIFT))))     ' v24: today's crew
                out(k, 3) = Trim$(SafeStr(v(r, C_HOME)))
                out(k, 4) = NumOr0(v(r, C_TOTAL))
                out(k, 5) = NumOr0(v(r, C_WORKED))
                out(k, 6) = NumOr0(v(r, C_REFUSED))
                out(k, 7) = IIf(provLive, NumOr0(v(r, C_PROV)), 0)
                RefusedSplitOf splitMap, nm, rc, rp, ro
                out(k, 8) = rc: out(k, 9) = rp: out(k, 10) = ro
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
    ' v24.3: one list per unit serves all the days, unless a planned unit move takes
    ' effect inside the window - then each date gets its own list, as the master's
    ' call-out and posting lists do
    Dim perDate As Boolean: perDate = PlannedWithin(Date, Date + MEMBER_NEXT_DAYS - 1)
    For u = 0 To UBound(depts)
        n = LoadEligible(CStr(depts(u)), "ALL", base)
        If n > 0 Or perDate Then
            For dd = 0 To MEMBER_NEXT_DAYS - 1
                d = Date + dd
                If perDate Then n = LoadEligible(CStr(depts(u)), "ALL", base, 999, True, d)
                If n > 0 Then
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
                End If
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
        ' v24.1: the crews on days / nights (unit "#"), for the day header of the members' month grid
        k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
        out(k, 1) = d: out(k, 2) = "": out(k, 3) = "#": out(k, 4) = DaysCrew(d): out(k, 5) = NightsCrew(d): out(k, 6) = "CREW": out(k, 7) = False
        For u = 0 To UBound(depts)
            For Each sh In Array("Days", "Nights")
                n = SlotsFor(d, CStr(sh), CStr(depts(u)), slots, True)
                For i = 1 To n
                    k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
                    out(k, 1) = d: out(k, 2) = CStr(sh): out(k, 3) = depts(u)
                    out(k, 4) = PublicLeft(slots(i).LeftText, slots(i).TypeCode): out(k, 5) = slots(i).RightText
                    out(k, 6) = PublicType(slots(i).TypeCode): out(k, 7) = slots(i).IsOpen
                Next i
            Next sh
        Next u
    Next d
    nC = k - 1
    BuildCal = Trimmed(out, k, 7)
End Function

' leave of any kind is one public code; the members' calendar colours it "Leave".
' v24.3.1 (decided D-M1, Oct 13, 2026): a sick call or a personal day goes out as OFF - the kind
' of absence never leaves the master; the five sick-time COUNTS on the card stay (BuildAttend).
' (This reverses the same-day decision to show "Sick".) The members' file has no OFF colour
' and draws it as any other line.
Private Function PublicType(ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": PublicType = "LV"
        Case "SIC", "PER": PublicType = "OFF"
        Case Else: PublicType = UCase$(Trim$(code))
    End Select
End Function

' v24.3.1: the calendar line's text as members may see it - "Smith P (Sick 8h)" becomes
' "Smith P (Off 8h)", a personal day the same; everything else as drawn
Private Function PublicLeft(ByVal lt As String, ByVal code As String) As String
    Select Case UCase$(Trim$(code))
        Case "SIC": PublicLeft = Replace(lt, "(Sick", "(Off", 1, -1, vbTextCompare)
        Case "PER": PublicLeft = Replace(lt, "(Personal", "(Off", 1, -1, vbTextCompare)
        Case Else: PublicLeft = lt
    End Select
End Function

' Open postings: ID, Unit, Date, Shift, Hours, Positions, Reason, Closes, Listed, Signed up;
' their lists: ID, Rank, Name, Hours posted, Signed ("" / "Yes" / "8 h").
' v24.1: Signed comes from the electronic sign-up inbox (names, as on the paper
' sheet - decided Oct 12, 2026). Someone who signed up without being on the
' list is added with Rank 0 (the paper sheet's out-of-department lines).
Private Sub BuildPostings(ByRef post As Variant, ByRef nP As Long, ByRef plist As Variant, ByRef nPL As Long)
    Dim ids As Variant, choices As Variant
    choices = OpenPostingChoices(ids)
    Dim capP As Long: capP = CAP_ROWS
    Dim capL As Long: capL = CAP_ROWS
    Dim p() As Variant: ReDim p(1 To capP, 1 To 10)
    Dim l() As Variant: ReDim l(1 To capL, 1 To 5)
    p(1, 1) = "Posting ID": p(1, 2) = "Unit": p(1, 3) = "Date": p(1, 4) = "Shift"
    p(1, 5) = "Hours": p(1, 6) = "Positions": p(1, 7) = "Reason": p(1, 8) = "Closes"
    p(1, 9) = "Listed": p(1, 10) = "Signed up"
    l(1, 1) = "Posting ID": l(1, 2) = "Rank": l(1, 3) = "Name": l(1, 4) = "Hours posted": l(1, 5) = "Signed"
    Dim kP As Long: kP = 1
    Dim kL As Long: kL = 1
    If IsArray(ids) Then
        Dim i As Long, j As Long, pr As PostRec, blank As PostRec, lst() As PListRec, n As Long
        Dim sg As Object, nSigned As Long, kk As Variant, onList As Object
        For i = LBound(ids) To UBound(ids)
            pr = blank                              ' LoadPosting leaves pr untouched when the id isn't found
            n = LoadPosting(CStr(ids(i)), pr, lst)
            If pr.ID <> "" Then                     ' a posting with an empty list still shows (as "list not published")
                kP = kP + 1: If kP > capP Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " postings - raise CAP_ROWS in modPublish"
                p(kP, 1) = pr.ID: p(kP, 2) = pr.Unit: p(kP, 3) = pr.ShiftDate: p(kP, 4) = pr.ShiftLbl
                p(kP, 5) = pr.Hours: p(kP, 6) = pr.Positions: p(kP, 7) = pr.Reason: p(kP, 8) = pr.ClosesOn
                Set sg = SignupStatusMap(pr.ID, pr.Hours, pr.ClosesOn)
                Set onList = CreateObject("Scripting.Dictionary")
                onList.CompareMode = 1
                nSigned = 0
                For j = 1 To n
                    kL = kL + 1: If kL > capL Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " posting list rows - raise CAP_ROWS in modPublish"
                    l(kL, 1) = pr.ID: l(kL, 2) = lst(j).Rank: l(kL, 3) = lst(j).EmpName: l(kL, 4) = lst(j).TotalHrs
                    l(kL, 5) = ""
                    onList(Trim$(lst(j).EmpName)) = True
                    If sg.Exists(Trim$(lst(j).EmpName)) Then l(kL, 5) = sg(Trim$(lst(j).EmpName)): nSigned = nSigned + 1
                Next j
                For Each kk In sg.Keys                  ' signed up, not on the list: out of department
                    If Not onList.Exists(CStr(kk)) Then
                        kL = kL + 1: If kL > capL Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " posting list rows - raise CAP_ROWS in modPublish"
                        l(kL, 1) = pr.ID: l(kL, 2) = 0: l(kL, 3) = CStr(kk): l(kL, 4) = "": l(kL, 5) = sg(kk)
                        nSigned = nSigned + 1
                    End If
                Next kk
                p(kP, 9) = n: p(kP, 10) = nSigned
            End If
        Next i
    End If
    nP = kP - 1: nPL = kL - 1
    post = Trimmed(p, kP, 10)
    plist = Trimmed(l, kL, 5)
End Sub

' v24.1 / M2 (decided Oct 12, 2026): sick time COUNTS for each person, so the
' members' file can show them to whoever picks that name. Name, Sick days used,
' Paid sick days left, Occurrences used, Personal days used, Status.
' Never a date of a sick or personal day, an occurrence date, a reason or a
' note: the only date allowed is the "until" date in the status.
Private Function BuildAttend(ByRef nA As Long) As Variant
    Dim names As Variant: names = RosterNames()
    Dim n As Long, i As Long, st As AttendanceStatus, nm As String
    If IsArray(names) Then n = UBound(names) - LBound(names) + 1
    Dim out() As Variant: ReDim out(1 To n + 1, 1 To 6)
    out(1, 1) = "Name": out(1, 2) = "Sick days used": out(1, 3) = "Paid sick days left"
    out(1, 4) = "Occurrences used": out(1, 5) = "Personal days used": out(1, 6) = "Status"
    If n > 0 Then
        BeginAttendanceBatch
        On Error GoTo Failed
        For i = 1 To n
            nm = CStr(names(LBound(names) + i - 1))
            AttendanceFor nm, Date, st
            out(i + 1, 1) = nm
            out(i + 1, 2) = st.SickUsed
            out(i + 1, 3) = st.PaidLeft
            out(i + 1, 4) = st.Occurrences
            out(i + 1, 5) = st.PersonalUsed
            out(i + 1, 6) = AttendPublicStatus(st)
        Next i
        EndAttendanceBatch
    End If
    nA = n
    BuildAttend = out
    Exit Function
Failed:
    Dim em As String: em = Err.Description          ' (PublishMembers logs it: this is re-raised below)
    On Error Resume Next
    EndAttendanceBatch
    On Error GoTo 0
    Err.Raise vbObjectError + 603, , "sick time counts: " & em
End Function

' The four things the members' file may say about paid sick time
Public Function AttendPublicStatus(ByRef st As AttendanceStatus) As String
    If Not st.Entitled Then
        AttendPublicStatus = "No paid sick until " & Format$(st.WaitUntil, "mmm d, yyyy")
    ElseIf st.PaidLeft <= 0 Then
        AttendPublicStatus = "Paid sick days used up"
    ElseIf st.Occurrences >= PAID_SICK_OCC Then
        AttendPublicStatus = "Paid sick only while the current occurrence continues"
    Else
        AttendPublicStatus = "Paid sick available"
    End If
End Function

' dAttend may hold a name, four counts and one of the status phrases - nothing else
Private Function AttendPrivacy(ByRef att As Variant) As String
    If Not IsArray(att) Then Exit Function
    If UBound(att, 2) <> 6 Then AttendPrivacy = DS_ATT & ": unexpected columns": Exit Function
    Dim i As Long, j As Long, s As String
    For i = 2 To UBound(att, 1)
        For j = 2 To 5
            If Not AttendCountOk(att(i, j)) Then
                AttendPrivacy = DS_ATT & " row " & i & ": column " & j & " must be a count"
                Exit Function
            End If
        Next j
        s = SafeStr(att(i, 6))
        If Not AttendStatusOk(s) Then
            AttendPrivacy = DS_ATT & " row " & i & ": status isn't one of the four allowed phrases"
            Exit Function
        End If
    Next i
End Function

' A sick-time count may be a number or blank - never text and never a date
Public Function AttendCountOk(ByVal x As Variant) As Boolean
    AttendCountOk = Not (VarType(x) = vbString Or VarType(x) = vbDate)
End Function

' The four things the members' file may say about paid sick time
Public Function AttendStatusOk(ByVal s As String) As Boolean
    AttendStatusOk = (s = "Paid sick available" Or s = "Paid sick days used up" Or _
                      s = "Paid sick only while the current occurrence continues" Or s Like "No paid sick until [A-Z][a-z][a-z] *, ####")
End Function

' This year's (by Period Start) non-void log rows: Shift Date, Ref ID, Type, Unit, Shift, Hours, Employee,
' Outcome, Hours Charged, Period Start (+ Note when PUBLISH_REASONS)
Private Function BuildLog(ByRef nL As Long) As Variant
    Dim lo As ListObject: Set lo = LogTable()
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim yFrom As Double, yTo As Double
    yFrom = NumOr0(ws.Cells(PStartRow(), C_PERIOD1).Value2)
    yTo = NumOr0(ws.Cells(PEndRow(), PeriodLastCol()).Value2)
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
                    sd = DateNumber(v(r, L_PERIOD))         ' v24.1: Period Start, the test the roster's formulas use
                    If (yFrom = 0 Or sd >= yFrom) And (yTo = 0 Or sd <= yTo) Then
                        k = k + 1: If k > cap Then Err.Raise vbObjectError + 602, , "more than " & CAP_ROWS & " rows for the members' file - raise CAP_ROWS in modPublish"
                        out(k, 1) = v(r, L_DATE): out(k, 2) = SafeStr(v(r, L_REF)): out(k, 3) = SafeStr(v(r, L_TYPE))
                        out(k, 4) = SafeStr(v(r, L_DEPT)): out(k, 5) = SafeStr(v(r, L_SHIFTLBL)): out(k, 6) = NumOr0(v(r, L_BASEHRS))
                        out(k, 7) = SafeStr(v(r, L_EMP)): out(k, 8) = SafeStr(v(r, L_OUTCOME))
                        out(k, 9) = NumOr0(v(r, L_CHARGED)): out(k, 10) = v(r, L_PERIOD)
                        If PUBLISH_REASONS Then out(k, 11) = PublicNote(SafeStr(v(r, L_REASON)))
                    End If
                End If
            Next r
        End If
    End If
    nL = k - 1
    BuildLog = Trimmed(out, k, nCols)
End Function

' A date cell as a number (0 = not a date). NumOr0 can't be used: a cell read with
' .Value arrives as a Date, which IsNumeric says is not a number - so every log row
' failed the year test and the members' history was published empty (M1 - v24).
Private Function DateNumber(ByVal x As Variant) As Double
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If IsDate(x) Then
        DateNumber = CDbl(CDate(x))
    ElseIf IsNumeric(x) Then
        DateNumber = CDbl(x)
    End If
End Function

' A log Note as members may see it: the "[electronic sign-up <login> <time>]"
' marker the Award pop-up adds (v24.1) is cut off - a login is never published.
Private Function PublicNote(ByVal note As String) As String
    Dim p As Long: p = InStr(1, note, "[electronic sign-up", vbTextCompare)
    If p > 0 Then note = Trim$(Left$(note, p - 1))
    PublicNote = note
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
    Dim i As Long, j As Long, kind As String
    For i = 2 To UBound(a, 1)
        For j = 1 To UBound(a, 2)
            kind = PrivateValueKind(a(i, j))
            Select Case kind
                Case "email"
                    PrivateDataIn = tableName & " row " & i & ": looks like an email (" & Left$(CStr(a(i, j)), 40) & ")"
                    Exit Function
                Case "phone"
                    PrivateDataIn = tableName & " row " & i & ", column " & j & ": looks like a phone number or an ID (" & _
                                    MaskDigits(Left$(CStr(a(i, j)), 40)) & ")"
                    Exit Function
                Case "number"                      ' the value itself is not shown: it may be a Payroll #
                    PrivateDataIn = tableName & " row " & i & ", column " & j & ": a number of 7 or more digits (a Payroll # or a phone number?)"
                    Exit Function
            End Select
        Next j
    Next i
End Function

' "905-555-0142" -> "###-###-####": the message shows where the digits are, never what they are
Private Function MaskDigits(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then ch = "#"
        MaskDigits = MaskDigits & ch
    Next i
End Function

' What one published value looks like: "" (fine), "email", "phone" (text with a
' phone-like run of digits) or "number" (v24.2: a NUMBER of 1,000,000 or more -
' no hours figure is that large, a Payroll # is; dates are a different type).
Public Function PrivateValueKind(ByVal x As Variant) As String
    Select Case VarType(x)
        Case vbString
            If InStr(x, "@") > 0 Then
                PrivateValueKind = "email"
            ElseIf LongestDigitRun(CStr(x)) >= 7 Then
                PrivateValueKind = "phone"
            End If
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20       ' 20 = LongLong
            If Abs(CDbl(x)) >= 1000000 Then PrivateValueKind = "number"
    End Select
End Function

' Phone-like = a single run of 7+ digits, or 10+ digits across groups joined
' by phone separators (space - ( ) .): "905-555-1234" and "(905) 555 1234"
' are caught; "0700-1900", "07:00-15:00" and "P-0012" are not. Dates written
' yyyy-mm-dd are blanked first so they never count.
Private Function LongestDigitRun(ByVal s As String) As Long
    Dim joined As Long: joined = JoinedDigits(s)
    Dim oneRun As Long: oneRun = SingleRun(s)     ' not "single": that is a reserved word (a data type)
    If joined >= 10 Then
        LongestDigitRun = joined
    Else
        LongestDigitRun = oneRun
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

' v24.3.1 (review M1): every piece of FREE TEXT the members' file carries is screened for a leave
' word - the calendar's Left and Right text (imported old-calendar lines are kept as written, a
' Need line's note and a "set by hand" covered-by are typed), a posting's Reason (printed,
' emailed, drawn), and the log Note when PUBLISH_REASONS is on. Before this only the calendar's
' Left text was screened, and only for a bracketed code.
Private Function LeaveCodesIn(ByRef cal As Variant, ByRef post As Variant, ByRef lg As Variant) As String
    Dim i As Long, t As String, w As String
    If IsArray(cal) Then
        For i = 2 To UBound(cal, 1)
            t = UCase$(SafeStr(cal(i, 6)))
            If IsLeaveCode(t) Then
                LeaveCodesIn = DS_CAL & " row " & i & ": leave type " & t & " would be visible"
                Exit Function
            End If
            w = LeaveWordsIn(SafeStr(cal(i, 4)))
            If w = "" Then w = LeaveWordsIn(SafeStr(cal(i, 5)))
            If w <> "" Then
                LeaveCodesIn = DS_CAL & " row " & i & " (" & Format$(cal(i, 1), "mmm d") & " " & SafeStr(cal(i, 2)) & " " & _
                               SafeStr(cal(i, 3)) & "): the text says '" & w & "'"
                Exit Function
            End If
        Next i
    End If
    If IsArray(post) Then
        For i = 2 To UBound(post, 1)
            w = LeaveWordsIn(SafeStr(post(i, 7)))
            If w <> "" Then
                LeaveCodesIn = DS_POST & " row " & i & " (" & SafeStr(post(i, 1)) & ", Reason): the text says '" & w & "'"
                Exit Function
            End If
        Next i
    End If
    If IsArray(lg) Then
        If UBound(lg, 2) >= 11 Then                  ' the Note column is published only with PUBLISH_REASONS
            For i = 2 To UBound(lg, 1)
                w = LeaveWordsIn(SafeStr(lg(i, 11)))
                If w <> "" Then
                    LeaveCodesIn = DS_LOG & " row " & i & " (" & SafeStr(lg(i, 2)) & ", Note): the text says '" & w & "'"
                    Exit Function
                End If
            Next i
        End If
    End If
End Function

' v24.3.1: the first word or phrase in free text that would tell members the kind of leave, or that
' someone is sick or on a personal day ("" = clean). Whole words only, letters only: "Standard" is
' clean, "Acme Ltd" is not (the admin rewords it); "Leave", "Vac", "Trn", "Off" alone are clean.
' Pure; mirrored (leavewords cases).
Public Function LeaveWordsIn(ByVal txt As String) As String
    Dim toks() As String, n As Long, i As Long
    n = WordTokens(txt, toks)
    For i = 1 To n
        Select Case toks(i)
            Case "STD", "LOA", "WSIB", "LTD", "OLV", "PAR", "PARENTAL", "DISABILITY", "SICK", "PERSONAL"
                LeaveWordsIn = toks(i)
                Exit Function
            Case "LEAVE"
                If i + 2 <= n Then
                    If toks(i + 1) = "OF" And toks(i + 2) = "ABSENCE" Then LeaveWordsIn = "LEAVE OF ABSENCE": Exit Function
                End If
        End Select
    Next i
End Function

' the maximal runs of letters in s, upper-cased, in toks(1..n); returns n
Private Function WordTokens(ByVal s As String, ByRef toks() As String) As Long
    Dim i As Long, ch As String, cur As String, n As Long
    ReDim toks(1 To Len(s) + 1)
    For i = 1 To Len(s) + 1
        If i <= Len(s) Then ch = UCase$(Mid$(s, i, 1)) Else ch = " "
        If ch Like "[A-Z]" Then
            cur = cur & ch
        ElseIf cur <> "" Then
            n = n + 1
            toks(n) = cur
            cur = ""
        End If
    Next i
    WordTokens = n
End Function


' ====================================================================
'  WRITING INTO THE TEMPLATE
' ====================================================================
Private Function HasDataSheets(ByVal wb As Workbook) As Boolean
    Dim nm As Variant, sh As Object
    For Each nm In Array(DS_META, DS_ROSTER, DS_NEXT, DS_CAL, DS_POST, DS_PLIST, DS_LOG, DS_ATT)
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
