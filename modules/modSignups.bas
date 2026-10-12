Attribute VB_Name = "modSignups"
Option Explicit

' ====================================================================
'  ELECTRONIC SIGN-UPS (v24.1)
'
'  The members' workbook can't write to the master, and holds no rules.
'  Its "Sign me up" / "Withdraw" buttons each drop ONE small text file
'  into the inbox folder (SIGNUP_FOLDER, beside the members' file):
'       P-0012__<login>__<yyyymmdd-hhmmss>.txt
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

Public Const SG_VERSION As String = "Sign-ups v24.1 (Jan 5)"

' rows read from the inbox
Public Const SG_EMP As Long = 1
Public Const SG_HOURS As Long = 2       ' 0 = the whole shift
Public Const SG_ACTION As Long = 3      ' signup / withdraw
Public Const SG_LOGIN As Long = 4
Public Const SG_STAMP As Long = 5
Public Const SG_FILE As Long = 6
Public Const SG_N As Long = 6

' rows ReduceSignups gives back (people whose newest action is a sign-up)
Public Const SR_EMP As Long = 1
Public Const SR_HOURS As Long = 2
Public Const SR_LATE As Long = 3        ' stamped after the posting closed
Public Const SR_LOGIN As Long = 4
Public Const SR_STAMP As Long = 5
Public Const SR_N As Long = 5


' ====================================================================
'  PURE RULE
' ====================================================================
' rows(1..n, 1..SG_N) in any order. For each employee the action with the
' latest stamp wins (of two with the same stamp, the later row). People
' whose newest action is "signup" come back in out(1..m, 1..SR_N), in the
' order their newest action was found. closesAt = 0 means "never late".
Public Function ReduceSignups(ByRef rows As Variant, ByVal closesAt As Date, ByRef out() As Variant) As Long
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
    For Each k In newest.Keys
        r = CLng(newest(k))
        If LCase$(Trim$(SafeStr(rows(r, SG_ACTION)))) = "signup" Then
            m = m + 1
            out(m, SR_EMP) = Trim$(SafeStr(rows(r, SG_EMP)))
            out(m, SR_HOURS) = NumOr0(rows(r, SG_HOURS))
            out(m, SR_LATE) = (closesAt > 0 And CDate(rows(r, SG_STAMP)) > closesAt)
            out(m, SR_LOGIN) = SafeStr(rows(r, SG_LOGIN))
            out(m, SR_STAMP) = CDate(rows(r, SG_STAMP))
        End If
    Next k
    ReduceSignups = m
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
    On Error Resume Next
    StampOf = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2))) + _
              TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
End Function

' The inbox files for one posting: rows(1..n, 1..SG_N), or Empty.
Public Function ReadSignups(ByVal postingId As String) As Variant
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Function
    If Not FolderExists(folder) Then Exit Function
    postingId = UCase$(Trim$(postingId))
    Dim files() As String, n As Long, f As String
    On Error GoTo Quit
    f = Dir$(folder & postingId & "__*.txt")
    Do While f <> ""
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    If n = 0 Then Exit Function
    Dim rows() As Variant: ReDim rows(1 To n, 1 To SG_N)
    Dim i As Long, k As Long, ff As Integer, ln As String, p As Long, key As String, valTxt As String
    Dim post As String, emp As String, hrs As Double, act As String, login As String, stamp As Date
    For i = 1 To n
        post = "": emp = "": hrs = 0: act = "": login = "": stamp = 0
        ff = FreeFile
        Open folder & files(i) For Input Access Read Shared As #ff
        Do While Not EOF(ff)
            Line Input #ff, ln
            p = InStr(ln, "=")
            If p > 1 Then
                key = LCase$(Trim$(Left$(ln, p - 1)))
                valTxt = Trim$(Mid$(ln, p + 1))
                Select Case key
                    Case "posting": post = UCase$(valTxt)
                    Case "employee": emp = valTxt
                    Case "hours": If IsNumeric(valTxt) Then hrs = CDbl(valTxt)
                    Case "action": act = LCase$(valTxt)
                    Case "login": login = Left$(valTxt, 40)
                    Case "stamp": stamp = StampOf(valTxt)
                End Select
            End If
        Loop
        Close #ff
        ' the file must say the same posting as its name, and be complete
        If post = postingId And emp <> "" And stamp > 0 And (act = "signup" Or act = "withdraw") Then
            k = k + 1
            rows(k, SG_EMP) = Left$(emp, 80)
            rows(k, SG_HOURS) = IIf(hrs > 0, hrs, 0)
            rows(k, SG_ACTION) = act
            rows(k, SG_LOGIN) = login
            rows(k, SG_STAMP) = stamp
            rows(k, SG_FILE) = files(i)
        End If
    Next i
    If k = 0 Then Exit Function
    Dim out() As Variant: ReDim out(1 To k, 1 To SG_N)
    For i = 1 To k
        For p = 1 To SG_N
            out(i, p) = rows(i, p)
        Next p
    Next i
    ReadSignups = out
    Exit Function
Quit:
    On Error Resume Next
    Close #ff
End Function


' ====================================================================
'  AWARD POSTING > LOAD
' ====================================================================
' Pre-ticks the electronic sign-ups on a freshly loaded posting list.
' Returns how many names were ticked; summary is the line for the pop-up.
Public Function ApplySignups(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByRef summary As String) As Long
    summary = ""
    If Not SIGNUPS_ON Then Exit Function
    Dim rows As Variant: rows = ReadSignups(p.ID)
    If Not IsArray(rows) Then summary = "No electronic sign-ups in the inbox.": Exit Function
    Dim red() As Variant, m As Long, k As Long, i As Long, hit As Long
    Dim nTick As Long, nPart As Long, lateTxt As String, outTxt As String, nm As String
    m = ReduceSignups(rows, PostingCloseAt(p.ClosesOn), red)
    For k = 1 To m
        nm = CStr(red(k, SR_EMP))
        hit = 0
        For i = 1 To n
            If StrComp(Trim$(lst(i).EmpName), nm, vbTextCompare) = 0 Then hit = i: Exit For
        Next i
        If hit = 0 Then
            outTxt = outTxt & IIf(outTxt = "", "", ", ") & nm & IIf(NumOr0(red(k, SR_HOURS)) > 0, " (" & HoursText(NumOr0(red(k, SR_HOURS))) & " h)", "")
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
    summary = nTick & " electronic sign-up" & IIf(nTick = 1, "", "s") & " ticked" & IIf(nPart > 0, " (" & nPart & " partial)", "") & "." & _
              IIf(lateTxt <> "", "  AFTER CLOSING, not ticked: " & lateTxt & ".", "") & _
              IIf(outTxt <> "", "  Not on the list (out of department box): " & outTxt & ".", "") & _
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
    Dim red() As Variant, m As Long, k As Long
    m = ReduceSignups(rows, PostingCloseAt(closesOn), red)
    For k = 1 To m
        If Not CBool(red(k, SR_LATE)) Then
            If IsPartialSignup(postHours, NumOr0(red(k, SR_HOURS))) Then
                d(CStr(red(k, SR_EMP))) = HoursText(NumOr0(red(k, SR_HOURS))) & " h"
            Else
                d(CStr(red(k, SR_EMP))) = "Yes"
            End If
        End If
    Next k
End Function

' A posting was closed (award / no takers / cancel): its inbox files move to
' Done\<posting>\. Quiet: a file that can't be moved stays and is reported by
' IntegrityCheck later.
Public Sub ArchiveSignups(ByVal postingId As String)
    On Error Resume Next
    If Not SIGNUPS_ON Then Exit Sub
    Dim folder As String: folder = SignupFolderPath()
    If folder = "" Then Exit Sub
    If Not FolderExists(folder) Then Exit Sub
    postingId = UCase$(Trim$(postingId))
    Dim files() As String, n As Long, f As String, i As Long
    f = Dir$(folder & postingId & "__*.txt")
    Do While f <> ""
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
    Do While f <> ""
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
