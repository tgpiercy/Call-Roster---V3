Attribute VB_Name = "modSignups"
Option Explicit

' ====================================================================
'  ELECTRONIC SIGN-UPS (v24.1)
'
'  The members' workbook can't write to the master, and holds no rules.
'  Its "Sign me up" / "Withdraw" buttons each drop ONE small text file
'  into the inbox folder (SIGNUP_FOLDER, beside the members' file):
'       P-0012__<login>__<PC name>__<yyyymmdd-hhmmss>.txt   (the PC name since M3; older files had none)
'       posting=P-0012
'       employee=Smith, Joe
'       hours=8            (blank = the whole shift)
'       action=signup      (or withdraw)
'       login=jsmith
'       machine=PLANT-PC-07
'       stamp=2027-01-12 14:03:22
'  Nothing in a file is trusted as a decision. When Award Posting loads
'  a posting, the master reads that posting's files, keeps each person's
'  NEWEST action, and pre-ticks Signed up (with partial hours) for names
'  on the posting's list. The supervisor still confirms every tick. The
'  paper sheet stays valid.
'   - a sign-up stamped after the posting closed is listed, not ticked
'   - a name that isn't on the list is reported for the Out of department box
'   - the Windows login is shown to the supervisor and goes in the OT Log
'     note; it is never published (plant PCs are shared, so it shows which
'     login was used, not who clicked)
'  When the posting is closed its files move to Signups\Done\<posting>\.
'  SIGNUPS_ON (modOvertimeEntry) switches the whole thing on.
'
'  Pure (mirrored in tests/rules.py, checked by SelfTest): ReduceSignups.
' ====================================================================

Public Const SG_VERSION As String = "Sign-ups v29.0.1 (Oct 28)"

' rows read from the inbox
Public Const SG_EMP As Long = 1
Public Const SG_HOURS As Long = 2       ' 0 = the whole shift
Public Const SG_ACTION As Long = 3      ' signup / withdraw
Public Const SG_LOGIN As Long = 4
Public Const SG_STAMP As Long = 5
Public Const SG_FILE As Long = 6
Public Const SG_MTIME As Long = 7       ' v27 (item 14): the file's modified time on disk (0 = unknown)
Public Const SG_N As Long = 7

' rows ReduceSignups gives back (people whose newest action is a sign-up)
Public Const SR_EMP As Long = 1
Public Const SR_HOURS As Long = 2
Public Const SR_LATE As Long = 3        ' stamped after the posting closed
Public Const SR_LOGIN As Long = 4
Public Const SR_STAMP As Long = 5
Public Const SR_SUSPECT As Long = 6     ' v27 (item 14): True = the file was written more than SIGNUP_STAMP_SLACK_MIN after its stamp
Public Const SR_LOGINMIX As Long = 7    ' v27 (item 14): True = another file for this name came from a different login
Public Const SR_N As Long = 7
Public Const SIGNUP_STAMP_SLACK_MIN As Long = 10   ' v27 (item 14, A12): a stamp this much older than the file's own time is suspect

' v26 (M3, D-M4): rows WithdrawnSignups gives back (people whose newest action is a withdraw)
Public Const WD_EMP As Long = 1
Public Const WD_LOGIN As Long = 2       ' the login that withdrew
Public Const WD_STAMP As Long = 3
Public Const WD_MISMATCH As Long = 4    ' True = an earlier sign-up came from a different login
Public Const WD_N As Long = 4


' ====================================================================
'  PURE RULE
' ====================================================================
' rows(1..n, 1..SG_N) in any order. For each employee the action with the
' latest stamp wins (of two with the same stamp, the later row). People
' whose newest action is "signup" come back in out(1..m, 1..SR_N), in the
' order each name first appeared. closesAt = 0 means "never late".
' v27 (item 14): asOf = the reader's now (for the clock-change allowance on the suspect flag); 0 = no allowance - the
' function stays PURE: it reads nothing but its arguments
Public Function ReduceSignups(ByRef rows As Variant, ByVal closesAt As Date, ByRef out() As Variant, Optional ByVal asOf As Date = 0) As Long
    If Not IsArray(rows) Then Exit Function
    Dim newest As Object: Set newest = CreateObject("Scripting.Dictionary")
    newest.CompareMode = 1
    Dim i As Long, nm As String, k As Variant, m As Long, r As Long
    For i = 1 To UBound(rows, 1)
        nm = Trim$(SafeStr(rows(i, SG_EMP)))
        If nm <> "" And IsDate(rows(i, SG_STAMP)) Then
            If Not newest.Exists(nm) Then
                newest.Add nm, i
            ElseIf CDate(rows(i, SG_STAMP)) >= CDate(rows(CLng(newest(nm)), SG_STAMP)) Then
                newest(nm) = i
            End If
        End If
    Next i
    If newest.Count = 0 Then Exit Function
    ReDim out(1 To newest.Count, 1 To SR_N)
    Dim lg As String, mt As Date
    For Each k In newest.Keys
        r = CLng(newest(k))
        If LCase$(Trim$(SafeStr(rows(r, SG_ACTION)))) = "signup" Then
            m = m + 1
            out(m, SR_EMP) = Trim$(SafeStr(rows(r, SG_EMP)))
            out(m, SR_HOURS) = NumOr0(rows(r, SG_HOURS))
            out(m, SR_LATE) = (closesAt > 0 And CDate(rows(r, SG_STAMP)) > closesAt)
            out(m, SR_LOGIN) = SafeStr(rows(r, SG_LOGIN))
            out(m, SR_STAMP) = CDate(rows(r, SG_STAMP))
            ' v27 (item 14, A12): the two flags - a back-dated stamp (the file is newer than it says by more than the
            ' slack) and a name whose files came from more than one login (a shared PC, or someone else's click)
            mt = 0
            If UBound(rows, 2) >= SG_MTIME Then
                If IsDate(rows(r, SG_MTIME)) Then mt = CDate(rows(r, SG_MTIME))
            End If
            Dim across As Boolean: across = False
            If mt > 0 And asOf > 0 Then across = ClockChangeBetween(CDate(rows(r, SG_STAMP)), asOf)
            out(m, SR_SUSPECT) = SignupSuspect(CDate(rows(r, SG_STAMP)), mt, across)
            lg = Trim$(SafeStr(rows(r, SG_LOGIN)))
            out(m, SR_LOGINMIX) = False
            If lg <> "" Then
                For i = 1 To UBound(rows, 1)
                    If i <> r Then
                        If StrComp(Trim$(SafeStr(rows(i, SG_EMP))), CStr(k), vbTextCompare) = 0 Then
                            If Trim$(SafeStr(rows(i, SG_LOGIN))) <> "" And StrComp(Trim$(SafeStr(rows(i, SG_LOGIN))), lg, vbTextCompare) <> 0 Then out(m, SR_LOGINMIX) = True: Exit For
                        End If
                    End If
                Next i
            End If
        End If
    Next k
    ReduceSignups = m
End Function

' v27 (item 14, A12): a sign-up file whose own stamp is more than SIGNUP_STAMP_SLACK_MIN older than the time the
' file was written is suspect (a stamp typed or reset by hand; the members' file writes the two within a second).
' PURE (SelfTest "sgsuspect"; mirrored in tests/rules.py). An unknown time on either side = not suspect.
' Compared in WHOLE SECONDS (review H1: a date difference x 1440 is 10.000000001 for an exact ten minutes, so
' the Excel SelfTest failed where the Python mirror passed). acrossClockChange (review M2): a clock change
' between the stamp and now shifts FileDateTime's reading by an hour, so a difference within the slack of 60 min
' is then not suspect. A PC clock more than 10 min off the file server's flags every sign-up: a known limit.
Public Function SignupSuspect(ByVal stamp As Date, ByVal fileTime As Date, Optional ByVal acrossClockChange As Boolean = False) As Boolean
    If stamp = 0 Or fileTime = 0 Then Exit Function
    Dim gapSecs As Double: gapSecs = Round((fileTime - stamp) * 86400#)     ' v27.1: gapSecs (secs would hide modCalendar.Secs)
    If acrossClockChange Then
        If Abs(gapSecs - 3600) <= SIGNUP_STAMP_SLACK_MIN * 60 Then Exit Function
    End If
    SignupSuspect = (gapSecs > SIGNUP_STAMP_SLACK_MIN * 60)
End Function

' True when a clock change lies between the two times: the change happens at 2 AM the morning after a change night
' (modShiftCycle names the nights), so the moment tested is 2 AM on the day after each one (re-review B5); at most a
' year is looked at
Public Function ClockChangeBetween(ByVal d1 As Date, ByVal d2 As Date) As Boolean
    If d1 = 0 Or d2 = 0 Then Exit Function
    Dim a As Date, b As Date, k As Date, m As Date
    a = d1: b = d2
    If a > b Then k = a: a = b: b = k
    If b - a > 366 Then b = a + 366
    For k = Int(a) - 1 To Int(b)
        If DstNightHours(k) <> 12 Then
            m = Int(k) + 1 + TimeSerial(2, 0, 0)
            If a < m And b >= m Then ClockChangeBetween = True: Exit Function
        End If
    Next k
End Function

' v26 (M3; D-M4, Oct 13: every withdrawal listed with login and time at LOAD). PURE (SelfTest
' "withdrawn"): the people whose NEWEST action is a withdraw, in the order each name first appeared:
' out(1..m, 1..WD_N) = name, the withdrawing login, its stamp, and whether a sign-up before it came
' from a different login (a shared plant PC, or someone else's click). Nothing here changes a tick.
Public Function WithdrawnSignups(ByRef rows As Variant, ByRef out() As Variant) As Long
    If Not IsArray(rows) Then Exit Function
    Dim newest As Object: Set newest = CreateObject("Scripting.Dictionary")
    newest.CompareMode = 1
    Dim i As Long, nm As String, k As Variant, m As Long, r As Long, mism As Boolean, wdLogin As String
    For i = 1 To UBound(rows, 1)
        nm = Trim$(SafeStr(rows(i, SG_EMP)))
        If nm <> "" And IsDate(rows(i, SG_STAMP)) Then
            If Not newest.Exists(nm) Then
                newest.Add nm, i
            ElseIf CDate(rows(i, SG_STAMP)) >= CDate(rows(CLng(newest(nm)), SG_STAMP)) Then
                newest(nm) = i
            End If
        End If
    Next i
    If newest.Count = 0 Then Exit Function
    ReDim out(1 To newest.Count, 1 To WD_N)
    For Each k In newest.Keys
        r = CLng(newest(k))
        If LCase$(Trim$(SafeStr(rows(r, SG_ACTION)))) = "withdraw" Then
            wdLogin = Trim$(SafeStr(rows(r, SG_LOGIN)))
            mism = False
            For i = 1 To UBound(rows, 1)                   ' any earlier sign-up by this name from another login
                If StrComp(Trim$(SafeStr(rows(i, SG_EMP))), CStr(k), vbTextCompare) = 0 And i <> r Then
                    If LCase$(Trim$(SafeStr(rows(i, SG_ACTION)))) = "signup" And IsDate(rows(i, SG_STAMP)) Then
                        If CDate(rows(i, SG_STAMP)) <= CDate(rows(r, SG_STAMP)) Then
                            If StrComp(Trim$(SafeStr(rows(i, SG_LOGIN))), wdLogin, vbTextCompare) <> 0 Then mism = True
                        End If
                    End If
                End If
            Next i
            m = m + 1
            out(m, WD_EMP) = Trim$(SafeStr(rows(r, SG_EMP)))
            out(m, WD_LOGIN) = wdLogin
            out(m, WD_STAMP) = CDate(rows(r, SG_STAMP))
            out(m, WD_MISMATCH) = mism
        End If
    Next k
    WithdrawnSignups = m
End Function


' ====================================================================
'  THE INBOX FOLDER
' ====================================================================
' SIGNUP_FOLDER resolved: ".\Signups" = a folder beside the members' file
Public Function SignupFolderPath() As String
    Dim p As String: p = Trim$(SIGNUP_FOLDER)
    If p = "" Then Exit Function
    If p = "." Then
        p = PublishFolderPath()
    ElseIf Left$(p, 2) = ".\" Then
        p = PublishFolderPath() & Mid$(p, 3)
    End If
    SignupFolderPath = EnsureSlash(p)
End Function

' Makes the inbox folder if it isn't there. True when it exists afterwards.
Public Function EnsureSignupFolder() As Boolean
    Dim p As String: p = SignupFolderPath()
    If p = "" Then Exit Function
    On Error Resume Next
    If Not FolderExists(p) Then MkDir Left$(p, Len(p) - 1)
    On Error GoTo 0
    EnsureSignupFolder = FolderExists(p)
End Function

' When a posting stops taking sign-ups: its closing date at POSTING_CLOSE_TIME
Public Function PostingCloseAt(ByVal closesOn As Date) As Date
    If closesOn = 0 Then Exit Function
    PostingCloseAt = Int(closesOn) + TimeValue(POSTING_CLOSE_TIME)
End Function

' "2027-01-12 14:03:22" -> date + time (0 if it can't be read). Built from the
' numbers, never CDate: the files are written on other PCs.
Private Function StampOf(ByVal s As String) As Date
    s = Trim$(s)
    If Len(s) < 19 Then Exit Function
    If Not (s Like "####-##-## ##:##:##*") Then Exit Function
    On Error Resume Next                            ' never raises
    StampOf = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2))) + _
              TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
End Function

' The inbox files for one posting: rows(1..n, 1..SG_N), or Empty. A file that
' can't be opened or read (locked by the PC still writing it, antivirus) is
' skipped on its own, so it never hides the others; skipped says how many.
Public Function ReadSignups(ByVal postingId As String, Optional ByRef skipped As Long) As Variant
    skipped = 0
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Function
    If Not FolderExists(folder) Then Exit Function
    postingId = UCase$(Trim$(postingId))
    Dim files() As String, n As Long, f As String
    On Error Resume Next
    f = Dir$(folder & postingId & "__*.txt")
    Do While f <> "" And Err.Number = 0
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    On Error GoTo 0
    If n = 0 Then Exit Function
    Dim rows() As Variant: ReDim rows(1 To n, 1 To SG_N)
    Dim i As Long, k As Long, p As Long, mt As Date
    Dim post As String, emp As String, hrs As Double, act As String, login As String, stamp As Date
    For i = 1 To n
        If ReadOneSignup(folder & files(i), post, emp, hrs, act, login, stamp) Then
            ' the file must say the same posting as its name, and be complete
            If post = postingId And emp <> "" And stamp > 0 And (act = "signup" Or act = "withdraw") Then
                k = k + 1
                rows(k, SG_EMP) = Left$(emp, 80)
                rows(k, SG_HOURS) = IIf(hrs > 0, hrs, 0)
                rows(k, SG_ACTION) = act
                rows(k, SG_LOGIN) = login
                rows(k, SG_STAMP) = stamp
                rows(k, SG_FILE) = files(i)
                mt = 0                                   ' v27 (item 14): the file's own time, for the suspect flag
                On Error Resume Next
                mt = FileDateTime(folder & files(i))
                On Error GoTo 0
                rows(k, SG_MTIME) = mt
            End If
        Else
            skipped = skipped + 1
        End If
    Next i
    If k = 0 Then Exit Function
    Dim out() As Variant: ReDim out(1 To k, 1 To SG_N)
    For i = 1 To k
        For p = 1 To SG_N
            out(i, p) = rows(i, p)
        Next p
    Next i
    SortSignupRows out, k                            ' v27 (item 14): by stamp, then file name - the same order on every PC
    ReadSignups = out
End Function

' v27 (item 14, A12): the rows in (stamp, file name) order, so a tie on the stamp is broken the same way on every
' PC (Dir$ returns files in no promised order; ReduceSignups keeps the LATER row on a tie). The members' file
' names a file POST__login__pc__yyyymmdd-hhnnss[-n].txt, so a same-second tie falls to the LOGIN (a word sort,
' vbTextCompare: the "-n" suffix barely weighs, which matters only past ten clicks in one second)
Private Sub SortSignupRows(ByRef rows() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, c As Long, tmp(1 To SG_N) As Variant
    For i = 2 To n
        For c = 1 To SG_N
            tmp(c) = rows(i, c)
        Next c
        j = i - 1
        Do While j >= 1
            If Not SignupRowAfter(rows, j, tmp) Then Exit Do
            For c = 1 To SG_N
                rows(j + 1, c) = rows(j, c)
            Next c
            j = j - 1
        Loop
        For c = 1 To SG_N
            rows(j + 1, c) = tmp(c)
        Next c
    Next i
End Sub

' True when row j sorts after the row held in tmp (stamp, then file name)
Private Function SignupRowAfter(ByRef rows() As Variant, ByVal j As Long, ByRef tmp() As Variant) As Boolean
    Dim a As Double, b As Double
    If IsDate(rows(j, SG_STAMP)) Then a = CDbl(CDate(rows(j, SG_STAMP)))
    If IsDate(tmp(SG_STAMP)) Then b = CDbl(CDate(tmp(SG_STAMP)))
    If a > b Then SignupRowAfter = True: Exit Function
    If a < b Then Exit Function
    SignupRowAfter = (StrComp(SafeStr(rows(j, SG_FILE)), SafeStr(tmp(SG_FILE)), vbTextCompare) > 0)
End Function

' v26.5.5 (A8): the inbox files (any posting, not the Done folder) whose employee line is this name - a
' rename leaves them as they are, so the Employee pop-up's Rename preview can say how many will be ignored.
Public Function SignupFilesNaming(ByVal nameTxt As String) As Long
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Function
    If Not FolderExists(folder) Then Exit Function
    nameTxt = Trim$(nameTxt)
    If nameTxt = "" Then Exit Function
    Dim files() As String, n As Long, f As String, i As Long
    On Error Resume Next
    f = Dir$(folder & "*__*.txt")
    Do While f <> "" And Err.Number = 0
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    On Error GoTo 0
    Dim post As String, emp As String, hrs As Double, act As String, login As String, stamp As Date
    For i = 1 To n
        If ReadOneSignup(folder & files(i), post, emp, hrs, act, login, stamp) Then
            If StrComp(Trim$(emp), nameTxt, vbTextCompare) = 0 Then SignupFilesNaming = SignupFilesNaming + 1
        End If
    Next i
End Function

' Reads one inbox file. False = it couldn't be opened or read.
Private Function ReadOneSignup(ByVal fullPath As String, ByRef post As String, ByRef emp As String, ByRef hrs As Double, _
                               ByRef act As String, ByRef login As String, ByRef stamp As Date) As Boolean
    post = "": emp = "": hrs = 0: act = "": login = "": stamp = 0
    Dim ff As Integer, ln As String, p As Long, valTxt As String, opened As Boolean, nLines As Long
    On Error GoTo Quit
    ff = FreeFile
    Open fullPath For Input Access Read Shared As #ff
    opened = True
    Do While Not EOF(ff) And nLines < 50
        Line Input #ff, ln
        nLines = nLines + 1
        p = InStr(ln, "=")
        If p > 1 Then
            valTxt = Trim$(Mid$(ln, p + 1))
            Select Case LCase$(Trim$(Left$(ln, p - 1)))
                Case "posting": post = UCase$(valTxt)
                Case "employee": emp = valTxt
                Case "hours": hrs = Val(valTxt)          ' Val reads "8.5" the same on every PC, whatever its decimal sign
                Case "action": act = LCase$(valTxt)
                Case "login": login = Left$(valTxt, 40)
                Case "stamp": stamp = StampOf(valTxt)
            End Select
        End If
    Loop
    Close #ff
    ReadOneSignup = True
    Exit Function
Quit:
    On Error Resume Next
    If opened Then Close #ff
End Function

' Roster names, lower case -> as spelled on the roster. A sign-up counts only
' for a name that is on the roster: the inbox folder is writable by everyone,
' so nothing typed into a file is ever shown or published as it stands.
Private Function RosterNameMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set RosterNameMap = d
    Dim names As Variant: names = RosterNames()
    If Not IsArray(names) Then Exit Function
    Dim i As Long
    For i = LBound(names) To UBound(names)
        d(Trim$(CStr(names(i)))) = Trim$(CStr(names(i)))
    Next i
End Function


' ====================================================================
'  AWARD POSTING > LOAD
' ====================================================================
' Pre-ticks the electronic sign-ups on a freshly loaded posting list.
' Returns how many names were ticked; summary is the line for the pop-up.
Public Function ApplySignups(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByRef summary As String) As Long
    summary = ""
    If Not SIGNUPS_ON Then Exit Function
    Dim skipped As Long
    Dim rows As Variant: rows = ReadSignups(p.ID, skipped)
    If Not IsArray(rows) Then
        summary = IIf(skipped > 0, skipped & " sign-up file(s) couldn't be read just now - LOAD again in a minute.", "No electronic sign-ups in the inbox.")
        Exit Function
    End If
    Dim red() As Variant, m As Long, k As Long, i As Long, hit As Long
    Dim nTick As Long, nPart As Long, lateTxt As String, outTxt As String, nm As String, nUnknown As Long
    Dim onRoster As Object: Set onRoster = RosterNameMap()
    m = ReduceSignups(rows, PostingCloseAt(p.ClosesOn), red, Now)
    For k = 1 To m
        nm = CStr(red(k, SR_EMP))
        hit = 0
        For i = 1 To n
            If StrComp(Trim$(lst(i).EmpName), nm, vbTextCompare) = 0 Then hit = i: Exit For
        Next i
        If hit = 0 And Not onRoster.Exists(nm) Then
            nUnknown = nUnknown + 1                     ' not a roster name: never shown as typed
        ElseIf hit = 0 Then
            outTxt = outTxt & IIf(outTxt = "", "", ", ") & onRoster(nm) & IIf(NumOr0(red(k, SR_HOURS)) > 0, " (" & HoursText(NumOr0(red(k, SR_HOURS))) & " h)", "")
        ElseIf CBool(red(k, SR_LATE)) Then
            lateTxt = lateTxt & IIf(lateTxt = "", "", ", ") & nm
        Else
            lst(hit).SignedUp = True
            lst(hit).SignupHrs = 0
            If IsPartialSignup(p.Hours, NumOr0(red(k, SR_HOURS))) Then lst(hit).SignupHrs = NumOr0(red(k, SR_HOURS)): nPart = nPart + 1
            lst(hit).ESign = CStr(red(k, SR_LOGIN)) & " " & Format$(CDate(red(k, SR_STAMP)), "mmm d h:nn AM/PM")
            nTick = nTick + 1
        End If
    Next k
    ' v27 (item 14, A12): the two flags named (the tick stands: the supervisor's tick is what counts, the paper sheet too)
    Dim suspTxt As String, mixTxt As String
    For k = 1 To m
        nm = CStr(red(k, SR_EMP))
        If onRoster.Exists(nm) Then
            If CBool(red(k, SR_SUSPECT)) Then suspTxt = suspTxt & IIf(suspTxt = "", "", ", ") & onRoster(nm)
            If CBool(red(k, SR_LOGINMIX)) Then mixTxt = mixTxt & IIf(mixTxt = "", "", ", ") & onRoster(nm)
        End If
    Next k
    ' v26 (M3, D-M4): every withdrawal named, with its login and time; a login other than the sign-up's is said
    Dim wd() As Variant, nWd As Long, wdTxt As String
    nWd = WithdrawnSignups(rows, wd)
    For k = 1 To nWd
        nm = CStr(wd(k, WD_EMP))
        If onRoster.Exists(nm) Then nm = onRoster(nm)           ' the roster's spelling; an unknown name as typed is never shown
        If Not onRoster.Exists(CStr(wd(k, WD_EMP))) Then nm = "(a name not on the roster)"
        wdTxt = wdTxt & IIf(wdTxt = "", "", "; ") & nm & " (" & CStr(wd(k, WD_LOGIN)) & ", " & _
                Format$(CDate(wd(k, WD_STAMP)), "mmm d h:nn AM/PM") & IIf(CBool(wd(k, WD_MISMATCH)), "; a DIFFERENT login than the sign-up", "") & ")"
    Next k
    summary = nTick & " electronic sign-up" & IIf(nTick = 1, "", "s") & " ticked" & IIf(nPart > 0, " (" & nPart & " partial)", "") & "." & _
              IIf(lateTxt <> "", "  AFTER CLOSING, not ticked: " & lateTxt & ".", "") & _
              IIf(outTxt <> "", "  Not on the list (out of department box): " & outTxt & ".", "") & _
              IIf(suspTxt <> "", "  CHECK the stamp (the file was written more than " & SIGNUP_STAMP_SLACK_MIN & " min after the time it claims - a PC clock, or a back-dated sign-up): " & suspTxt & ".", "") & _
              IIf(mixTxt <> "", "  More than one login for the same name: " & mixTxt & ".", "") & _
              IIf(wdTxt <> "", "  Withdrawn (newest file is a withdrawal), not ticked: " & wdTxt & ".", "") & _
              IIf(nUnknown > 0, "  " & nUnknown & " file(s) name someone who isn't on the roster: ignored.", "") & _
              IIf(skipped > 0, "  " & skipped & " file(s) couldn't be read: LOAD again in a minute.", "") & _
              "  Check them against the paper sheet."
    ApplySignups = nTick
End Function

' The sign-up state of every name for one OPEN posting, for the members' file:
' name -> "Yes" / "8 h". Late sign-ups are left out (the supervisor decides).
Public Function SignupStatusMap(ByVal postingId As String, ByVal postHours As Double, ByVal closesOn As Date) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set SignupStatusMap = d
    If Not SIGNUPS_ON Then Exit Function
    Dim rows As Variant: rows = ReadSignups(postingId)
    If Not IsArray(rows) Then Exit Function
    Dim red() As Variant, m As Long, k As Long, nm As String
    Dim onRoster As Object: Set onRoster = RosterNameMap()
    m = ReduceSignups(rows, PostingCloseAt(closesOn), red, Now)
    For k = 1 To m
        ' only a name that is on the roster is ever published, and in the roster's own spelling:
        ' whatever else a file holds (a formula, a phone number, anything typed) goes nowhere
        If Not CBool(red(k, SR_LATE)) And onRoster.Exists(CStr(red(k, SR_EMP))) Then
            nm = CStr(onRoster(CStr(red(k, SR_EMP))))
            If IsPartialSignup(postHours, NumOr0(red(k, SR_HOURS))) Then
                d(nm) = HoursText(NumOr0(red(k, SR_HOURS))) & " h"
            Else
                d(nm) = "Yes"
            End If
        End If
    Next k
End Function

' A posting was closed (award / no takers / cancel): its inbox files move to
' Done\<posting>\. Quiet: a file that can't be moved stays and is reported by
' IntegrityCheck later.
Public Sub ArchiveSignups(ByVal postingId As String)
    On Error Resume Next                            ' never raises
    If Not SIGNUPS_ON Then Exit Sub
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Sub
    If Not FolderExists(folder) Then Exit Sub
    postingId = UCase$(Trim$(postingId))
    Dim files() As String, n As Long, f As String, i As Long
    f = Dir$(folder & postingId & "__*.txt")
    Do While f <> "" And Err.Number = 0             ' (an error from Dir must end the loop, not repeat it)
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    If n = 0 Then Exit Sub
    If Not FolderExists(folder & "Done\") Then MkDir folder & "Done"
    If Not FolderExists(folder & "Done\" & postingId & "\") Then MkDir folder & "Done\" & postingId
    For i = 1 To n
        Name folder & files(i) As folder & "Done\" & postingId & "\" & files(i)
    Next i
End Sub

' A posting was reopened: its sign-up files come back from Done\<posting>\ into
' the inbox, so the people who signed up electronically are ticked again.
Public Sub RestoreSignups(ByVal postingId As String)
    On Error Resume Next                            ' never raises
    If Not SIGNUPS_ON Then Exit Sub
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Sub
    postingId = UCase$(Trim$(postingId))
    Dim doneDir As String: doneDir = folder & "Done\" & postingId & "\"
    If Not FolderExists(doneDir) Then Exit Sub
    Dim files() As String, n As Long, f As String, i As Long
    f = Dir$(doneDir & postingId & "__*.txt")
    Do While f <> "" And Err.Number = 0             ' (an error from Dir must end the loop, not repeat it)
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    For i = 1 To n
        Name doneDir & files(i) As folder & files(i)
    Next i
End Sub

' Inbox files more than minDays old whose posting is no longer open
' (IntegrityCheck). Returns how many; txt names the first few postings.
Public Function StaleSignups(ByVal minDays As Long, ByRef txt As String) As Long
    txt = ""
    On Error GoTo Quit
    If Not SIGNUPS_ON Then Exit Function
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Function
    If Not FolderExists(folder) Then Exit Function
    Dim files() As String, n As Long, f As String, i As Long, id As String, p As Long, cnt As Long
    f = Dir$(folder & "*__*.txt")
    Do While f <> "" And Err.Number = 0             ' (an error from Dir must end the loop, not repeat it)
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For i = 1 To n
        p = InStr(files(i), "__")
        id = UCase$(Left$(files(i), p - 1))
        If Int(Date) - Int(FileDateTime(folder & files(i))) > minDays Then
            If PostingStatus(id) <> "Open" Then
                cnt = cnt + 1
                If Not seen.Exists(id) Then
                    seen.Add id, True
                    If seen.Count <= 5 Then txt = txt & IIf(txt = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    StaleSignups = cnt
Quit:
End Function
