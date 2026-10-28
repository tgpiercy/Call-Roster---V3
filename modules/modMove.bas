Attribute VB_Name = "modMove"
Option Explicit

' ====================================================================
'  v29 (briefs/v29.md item 1; Graeme, Oct 27, 2026): THE FOLDER SPLIT AND THE GUIDED MOVE.
'  At go-live the master moves to a supervisor-only folder; the members' file, Signups\ and Requests\ go to the
'  members' folder. The folder settings stay constants: the ONE value the move stores in the workbook is the members'
'  folder (the hidden name MembersFolder, read by modPublish.PublishFolderPath when set), so the January cutover needs
'  no code install. Excel cannot open two files of one name, so the move cannot copy-and-reopen itself: MoveRoster
'  (admin, Alt+F8) checks the file, copies itself and the members' template to the new master folder with the members'
'  path already inside, moves the inbox folders (Signups\ with Done\, Requests\ with done\) to the new members' folder,
'  freezes THIS file with the hidden name MovedTo and saves, then says "close this, open the new one". The new file's
'  first admin open (MoveFirstOpen) marks the members' file dirty and asks for SystemCheck + IntegrityCheck + a publish.
'  A COPY / DRY RUN file runs the same steps in REHEARSAL mode into test folders: it writes the members' path flagged
'  "R|", which only a practice copy honours ("L|" for the live file), so a practice copy made from a moved live roster
'  never publishes into the real members' folder. Pure: StoredPathCore (suite mvstored), MoveTargetProblem (suite mvtarget).
' ====================================================================
Public Const MV_VERSION As String = "Move v28 (Oct 27)"        ' new in the v29 build; stamped on the current release until the v29 release prep (test_drift)
Private Const NM_MEMBERS As String = "MembersFolder"            ' "L|<path>" (the live roster) or "R|<path>" (a rehearsal on a copy)
Private Const NM_MOVEDTO As String = "MovedTo"                  ' "<new full path>|<yyyy-mm-dd hh:nn>" in the OLD (frozen) file
Private Const NM_MOVEDFROM As String = "MovedFrom"              ' "<old full path>|<yyyy-mm-dd hh:nn>" in the NEW file
Private Const CHECK_FILE As String = "~roster-check.tmp"

' ====================================================================
'  PURE
' ====================================================================
' The stored members' folder value against the file reading it: "L|<path>" is honoured by the live roster only, "R|<path>"
' (a rehearsal) by a COPY / DRY RUN only; anything else gives "" (the constant applies). Suite mvstored.
Public Function StoredPathCore(ByVal value As String, ByVal isCopy As Boolean) As String
    If InStr(value, "|") <> 2 Then Exit Function
    Dim flag As String: flag = UCase$(Left$(value, 1))
    Dim p As String: p = Trim$(Mid$(value, 3))
    If p = "" Then Exit Function
    If (flag = "L" And Not isCopy) Or (flag = "R" And isCopy) Then StoredPathCore = p
End Function

' The two folders typed for a move: "" when they will do, else why not (one line). masterNow = the roster's folder today.
' Suite mvtarget.
Public Function MoveTargetProblem(ByVal masterNow As String, ByVal masterNew As String, ByVal membersNew As String, _
                                  ByVal masterReachable As Boolean, ByVal membersReachable As Boolean) As String
    masterNow = LCase$(EnsureSlash(Trim$(masterNow))): masterNew = LCase$(EnsureSlash(Trim$(masterNew))): membersNew = LCase$(EnsureSlash(Trim$(membersNew)))
    If Trim$(masterNew) = "" Or masterNew = "\" Then MoveTargetProblem = "no master folder was given": Exit Function
    If Trim$(membersNew) = "" Or membersNew = "\" Then MoveTargetProblem = "no members' folder was given": Exit Function
    If masterNew = masterNow Then MoveTargetProblem = "the new master folder is the folder the roster is in now": Exit Function
    If Not masterReachable Then MoveTargetProblem = "the new master folder can't be reached or written: " & masterNew: Exit Function
    If Not membersReachable Then MoveTargetProblem = "the new members' folder can't be reached or written: " & membersNew: Exit Function
End Function

' ====================================================================
'  THE STORED VALUES (workbook-level hidden names)
' ====================================================================
Private Function NameGet(ByVal key As String) As String
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then Exit Function
    NameGet = Replace(Replace(Mid$(nmObj.RefersTo, 2), """", ""), "'", "")
End Function

Private Sub NameSet(ByVal key As String, ByVal v As String)
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:="=""" & v & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & v & """"
    End If
End Sub

' The members' folder the move stored for THIS kind of file ("" = none: PUBLISH_FOLDER applies), with its slash
Public Function StoredMembersFolder() As String
    Dim p As String: p = StoredPathCore(NameGet(NM_MEMBERS), IsDryRunName(ThisWorkbook.Name))
    If p <> "" Then StoredMembersFolder = EnsureSlash(p)
End Function

' Where this (frozen) copy moved to; "" = it never moved, it is live
Public Function MovedToPath() As String
    Dim s As String: s = NameGet(NM_MOVEDTO)
    If InStr(s, "|") > 0 Then MovedToPath = Left$(s, InStr(s, "|") - 1) Else MovedToPath = s
End Function

Public Function MovedToWhen() As String
    Dim s As String: s = NameGet(NM_MOVEDTO)
    If InStr(s, "|") > 0 Then MovedToWhen = Mid$(s, InStr(s, "|") + 1)
End Function

' "<old path> on <when>" for the file that came out of a move; "" = never moved
Public Function MovedFromText() As String
    Dim s As String: s = NameGet(NM_MOVEDFROM)
    If s = "" Then Exit Function
    If InStr(s, "|") > 0 Then MovedFromText = Left$(s, InStr(s, "|") - 1) & " on " & Mid$(s, InStr(s, "|") + 1) Else MovedFromText = s
End Function

' ====================================================================
'  REACHABLE / WRITABLE
' ====================================================================
' True when a small file can be made and removed in the folder. Never raises.
Public Function FolderWritable(ByVal p As String) As Boolean
    On Error Resume Next                            ' never raises
    p = EnsureSlash(Trim$(p))
    If p = "" Or Not FolderExists(p) Then Exit Function
    Dim f As String: f = p & CHECK_FILE
    Dim ff As Integer: ff = FreeFile
    Err.Clear
    Open f For Output As #ff
    If Err.Number <> 0 Then Exit Function
    Print #ff, "roster write check"
    Close #ff
    FolderWritable = (Dir$(f) <> "")
    Kill f
End Function

' One SystemCheck line per path - the resolved folder, reachable, writable - plus the members' file's presence and age,
' and what the move left behind. Facts into info, troubles into warn (SystemCheckOK stays versions-only). Never raises.
Public Sub MovePathLines(ByRef warn As String, ByRef info As String)
    On Error Resume Next                            ' never raises
    Dim master As String: master = EnsureSlash(ThisWorkbook.Path)
    Dim members As String: members = PublishFolderPath()
    info = info & "Paths:" & vbNewLine
    PathLine warn, info, "master folder", master, True
    PathLine warn, info, "members' folder", members, True
    PathLine warn, info, "sign-ups", SignupFolderPath(), True
    PathLine warn, info, "requests", RequestFolderPath(), True
    PathLine warn, info, "backups", BackupFolderPath(), True
    PathLine warn, info, "archive (FINAL)", ResolveFolder(ARCHIVE_FOLDER), True
    PathLine warn, info, "closed periods", ResolveFolder(CLOSED_FOLDER), False
    PathLine warn, info, "scans", ResolveFolder(SCAN_FOLDER), False
    PathLine warn, info, "reports", EnsureSlash(ReportFolder()), False
    Dim kp As String: kp = Trim$(KRONOS_PATH)
    If kp <> "" Then
        If Right$(kp, 1) <> "\" Then kp = Left$(kp, InStrRev(kp, "\"))
        PathLine warn, info, "Kronos reports", kp, False
    End If
    If Trim$(GUIDE_PATH) <> "" Then
        If Dir$(GUIDE_PATH) <> "" Then info = info & "  guide PDF: " & GUIDE_PATH & vbNewLine Else warn = warn & "- the guide PDF can't be found: " & GUIDE_PATH & vbNewLine
    End If
    info = info & "  error log: " & ErrorLogFilePath() & IIf(Dir$(ErrorLogFilePath()) <> "", "", " (no file yet)") & vbNewLine
    ' the members' file itself
    Dim mf As String: mf = members & IIf(IsDryRunName(ThisWorkbook.Name), "PRACTICE - ", "") & MEMBER_FILE_NAME
    If Dir$(mf) <> "" Then
        info = info & "  members' file present, saved " & Format$(FileDateTime(mf), "ddd mmm d h:mm AM/PM") & vbNewLine
    Else
        warn = warn & "- no members' file in the members' folder yet: publish it (" & mf & ")" & vbNewLine
    End If
    If StoredMembersFolder() <> "" Then info = info & "  the members' folder was set by the guided move (PUBLISH_FOLDER is the default only)" & vbNewLine
    If MovedFromText() <> "" Then
        info = info & "  this roster moved from " & MovedFromText() & vbNewLine
        Dim oldPath As String: oldPath = Left$(MovedFromText(), InStr(MovedFromText(), " on ") - 1)
        If Dir$(oldPath) <> "" Then warn = warn & "- the old copy still sits at " & oldPath & " under its live name: rename it MOVED - ... (the runbook)" & vbNewLine
    End If
End Sub

Private Sub PathLine(ByRef warn As String, ByRef info As String, ByVal what As String, ByVal p As String, ByVal needWrite As Boolean)
    On Error Resume Next                            ' never raises
    If Trim$(p) = "" Then info = info & "  " & what & ": (not set)" & vbNewLine: Exit Sub
    If Not FolderExists(p) Then
        warn = warn & "- " & what & " can't be reached: " & p & vbNewLine
    ElseIf needWrite And Not FolderWritable(p) Then
        warn = warn & "- " & what & " is not writable for you: " & p & vbNewLine
    Else
        info = info & "  " & what & ": " & p & IIf(needWrite, " (reachable, writable)", " (reachable)") & vbNewLine
    End If
End Sub

' ====================================================================
'  THE GUIDED MOVE (admin, Alt+F8; a cutover step, never an install step)
' ====================================================================
Public Sub MoveRoster()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim rehearsal As Boolean: rehearsal = IsDryRunName(ThisWorkbook.Name)
    Dim ttl As String: ttl = IIf(rehearsal, "Move the roster - REHEARSAL", "Move the roster")
    If IsBackupCopyName(ThisWorkbook.Name) Then MsgBox "This is a backup or frozen copy: the move runs from the live roster (or a practice COPY to rehearse).", vbExclamation, ttl: Exit Sub
    If NoSave() Then MsgBox "Saving is blocked in this file: close WITHOUT saving, reopen, and run the move again.", vbExclamation, ttl: Exit Sub
    If ThisWorkbook.Path = "" Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then MsgBox "Open the roster from an ordinary folder first (not SharePoint / OneDrive).", vbExclamation, ttl: Exit Sub
    Dim savedAt As Date
    If StaleOnDisk(savedAt) Then MsgBox "Someone else saved this file at " & Format$(savedAt, "h:mm AM/PM") & " since you opened it. Close WITHOUT saving, reopen, and run the move again.", vbExclamation, ttl: Exit Sub
    If Not ThisWorkbook.Saved Then
        If MsgBox("The roster has unsaved changes. Save it now and go on?", vbOKCancel + vbQuestion, ttl) <> vbOK Then Exit Sub
        If Not SaveIfWritable() Then MsgBox "The save did not go through: nothing moved.", vbExclamation, ttl: Exit Sub
    End If
    If OtherRosterOpen() Then MsgBox "Another workbook with the roster's name is open in this Excel: close it first.", vbExclamation, ttl: Exit Sub

    ' the two folders
    Dim masterNow As String: masterNow = EnsureSlash(ThisWorkbook.Path)
    Dim oldSignups As String: oldSignups = SignupFolderPath()
    Dim oldRequests As String: oldRequests = RequestFolderPath()
    Dim masterNew As String, membersNew As String, prob As String
    masterNew = Trim$(InputBox(IIf(rehearsal, "REHEARSAL: a TEST folder for the master (nothing on P:)." & vbNewLine & vbNewLine, "") & _
                               "The NEW folder for the master roster (the supervisor-only folder). It must exist already.", ttl, masterNow))
    If masterNew = "" Then Exit Sub
    membersNew = Trim$(InputBox(IIf(rehearsal, "REHEARSAL: a TEST folder for the members' side." & vbNewLine & vbNewLine, "") & _
                                "The NEW folder for the members' file, Signups\ and Requests\ (the members' folder). It must exist already.", ttl, PublishFolderPath()))
    If membersNew = "" Then Exit Sub
    masterNew = EnsureSlash(masterNew): membersNew = EnsureSlash(membersNew)
    prob = MoveTargetProblem(masterNow, masterNew, membersNew, FolderWritable(masterNew), FolderWritable(membersNew))
    If prob <> "" Then MsgBox "Nothing moved: " & prob & ".", vbExclamation, ttl: Exit Sub
    Dim newFull As String: newFull = masterNew & ThisWorkbook.Name
    If Dir$(newFull) <> "" Then MsgBox "Nothing moved: " & newFull & " exists already. Move or rename it first.", vbExclamation, ttl: Exit Sub
    Dim nIn As Long: nIn = CountTree(oldSignups) + CountTree(oldRequests)
    If MsgBox(IIf(rehearsal, "REHEARSAL on this practice copy - the real roster and P: are not touched." & vbNewLine & vbNewLine, "") & _
              "The move:" & vbNewLine & _
              "  1. copy this roster and the members' template to " & masterNew & vbNewLine & _
              "  2. move the sign-up and request files (" & nIn & ") to " & membersNew & " (Signups\, Requests\)" & vbNewLine & _
              "  3. write the members' folder into the new copy, so it publishes there" & vbNewLine & _
              "  4. freeze THIS file (nothing can be logged, saved or published in it again)" & vbNewLine & _
              "  5. you close this file and open the new one; its first open asks for SystemCheck, IntegrityCheck and a publish" & vbNewLine & vbNewLine & _
              "Backups, Closed, Scans, Practice and the error log start again beside the new master; the old ones stay here." & vbNewLine & vbNewLine & _
              "Go on?", vbOKCancel + vbQuestion + vbDefaultButton2, ttl) <> vbOK Then Exit Sub

    Application.StatusBar = "Moving the roster ..."
    ' 3 first: the values ride inside the copy
    Dim stamp As String: stamp = Format$(Now, "yyyy-mm-dd hh:nn")
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    NameSet NM_MEMBERS, IIf(rehearsal, "R|", "L|") & membersNew
    NameSet NM_MOVEDFROM, ThisWorkbook.FullName & "|" & stamp
    ' 1. the copy (the restricted view, as every copy) and the template
    SaveRestrictedCopy newFull
    If Dir$(newFull) = "" Then GoTo CopyFailed
    If FileLen(newFull) < 1000 Then GoTo CopyFailed
    If MemberTemplateExists() Then
        On Error Resume Next
        FileCopy MemberTemplatePath(), masterNew & MEMBER_TEMPLATE_NAME
        On Error GoTo Oops
    End If
    ' 2. the inbox folders: copy, count, then clear the old
    Dim nCopied As Long, nExpect As Long, note As String
    nExpect = CountTree(oldSignups)
    If nExpect > 0 Or FolderExists(oldSignups) Then
        If Not CopyTree(oldSignups, membersNew & "Signups\", nCopied) Or CountTree(membersNew & "Signups\") < nExpect Then
            note = "the sign-up files did not all copy (" & nCopied & " of " & nExpect & ")": GoTo InboxFailed
        End If
    End If
    nCopied = 0: nExpect = CountTree(oldRequests)
    If nExpect > 0 Or FolderExists(oldRequests) Then
        If Not CopyTree(oldRequests, membersNew & "Requests\", nCopied) Or CountTree(membersNew & "Requests\") < nExpect Then
            note = "the request files did not all copy (" & nCopied & " of " & nExpect & ")": GoTo InboxFailed
        End If
    End If
    ClearTree oldSignups: ClearTree oldRequests
    ' 4. freeze this file
    NameSet NM_MOVEDTO, newFull & "|" & stamp
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(NM_MOVEDFROM): nmObj.Delete       ' the old file did not come out of a move
    On Error GoTo Oops
    ThisWorkbook.Saved = False
    If Not SaveIfWritable() Then
        MsgBox "The new roster is in place at " & newFull & ", but THIS file could not be saved with its MOVED marker. " & _
               "Close it WITHOUT saving, then rename it MOVED - ... by hand so nobody opens it as the live roster.", vbExclamation, ttl
    End If
    Application.StatusBar = False
    MsgBox IIf(rehearsal, "REHEARSAL done. ", "Done. ") & "Now:" & vbNewLine & vbNewLine & _
           "  1. Close this file (it is frozen: every open says it moved)." & vbNewLine & _
           "  2. Open " & newFull & vbNewLine & _
           "  3. Its first open asks for SystemCheck (every path), IntegrityCheck and Publish the members' file." & vbNewLine & _
           "  4. Rename this file  MOVED " & Format$(Date, "yyyy-mm-dd") & " - " & ThisWorkbook.Name & "  in " & masterNow & vbNewLine & vbNewLine & _
           "Members' folder: " & membersNew & vbNewLine & "Sign-up and request files moved: " & nIn, vbInformation, ttl
    Exit Sub
CopyFailed:
    NameSet NM_MEMBERS, "": NameSet NM_MOVEDFROM, ""
    ThisWorkbook.Saved = wasSaved
    Application.StatusBar = False
    MsgBox "Nothing moved: the copy at " & newFull & " could not be made (the folder, or the file is open there). This file is unchanged.", vbExclamation, ttl
    Exit Sub
InboxFailed:
    NameSet NM_MEMBERS, "": NameSet NM_MOVEDFROM, ""
    ThisWorkbook.Saved = wasSaved
    On Error Resume Next
    Kill newFull
    Kill masterNew & MEMBER_TEMPLATE_NAME
    Application.StatusBar = False
    MsgBox "Nothing moved: " & note & ". The copy was taken back; the old files are untouched. Check the members' folder's rights and run the move again.", vbExclamation, ttl
    Exit Sub
Oops:
    Application.StatusBar = False
    ReportError "MoveRoster"
End Sub

' the first ADMIN open of a roster that came out of a move: the members' file is marked out of date (it must be published
' into the new folder) and one box, once per login, says what to run. Called by OpenChecks. Never raises.
Public Sub MoveFirstOpen()
    On Error Resume Next                            ' never raises
    If MovedFromText() = "" Or MovedToPath() <> "" Then Exit Sub
    If Not IsAdmin() Or ThisWorkbook.ReadOnly Then Exit Sub
    Dim key As String: key = LoginNameKey("MoveSeen")
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If Not nmObj Is Nothing Then Exit Sub
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    ThisWorkbook.Names.Add Name:=key, RefersTo:="=""seen""", Visible:=False
    MarkMembersDirty
    If wasSaved Then ThisWorkbook.Saved = True
    MsgBox "First open since the roster moved from " & MovedFromText() & "." & vbNewLine & vbNewLine & _
           "Run, in this order: SystemCheck (every path reachable and writable), IntegrityCheck, then Publish the members' file into " & _
           PublishFolderPath() & ". Then rename the old file MOVED - ... if that is not done.", vbInformation, "Moved roster"
End Sub

' ====================================================================
'  FILE HELPERS (Dir cannot nest: names are collected first, then walked)
' ====================================================================
Private Function OtherRosterOpen() As Boolean
    Dim w As Workbook
    For Each w In Application.Workbooks
        If Not w Is ThisWorkbook Then
            If StrComp(w.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then OtherRosterOpen = True: Exit Function
        End If
    Next w
End Function

' the files in a folder (one level) and its subfolders, as two arrays; counts returned
Private Sub ListFolder(ByVal p As String, ByRef files() As String, ByRef nF As Long, ByRef subs() As String, ByRef nS As Long)
    On Error Resume Next                            ' never raises
    nF = 0: nS = 0
    ReDim files(1 To 1): ReDim subs(1 To 1)
    p = EnsureSlash(p)
    If Not FolderExists(p) Then Exit Sub
    Dim f As String: f = Dir$(p & "*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            If (GetAttr(p & f) And vbDirectory) = vbDirectory Then
                nS = nS + 1: ReDim Preserve subs(1 To nS): subs(nS) = f
            Else
                nF = nF + 1: ReDim Preserve files(1 To nF): files(nF) = f
            End If
        End If
        f = Dir$()
    Loop
End Sub

' every file under p (two levels: the inbox folders hold Done\<posting>\ and done\)
Private Function CountTree(ByVal p As String) As Long
    Dim files() As String, subs() As String, nF As Long, nS As Long, i As Long
    ListFolder p, files, nF, subs, nS
    CountTree = nF
    For i = 1 To nS
        Dim f2() As String, s2() As String, nF2 As Long, nS2 As Long, j As Long
        ListFolder EnsureSlash(p) & subs(i), f2, nF2, s2, nS2
        CountTree = CountTree + nF2
        For j = 1 To nS2
            Dim f3() As String, s3() As String, nF3 As Long, nS3 As Long
            ListFolder EnsureSlash(p) & subs(i) & "\" & s2(j), f3, nF3, s3, nS3
            CountTree = CountTree + nF3
        Next j
    Next i
End Function

' copies p's files and its subfolders (two levels) under dst, making the folders; nCopied counts files; False on any failure
Private Function CopyTree(ByVal p As String, ByVal dst As String, ByRef nCopied As Long) As Boolean
    On Error GoTo Failed
    p = EnsureSlash(p): dst = EnsureSlash(dst)
    If Not FolderExists(dst) Then MkDir Left$(dst, Len(dst) - 1)
    If Not FolderExists(dst) Then Exit Function
    If Not FolderExists(p) Then CopyTree = True: Exit Function
    Dim files() As String, subs() As String, nF As Long, nS As Long, i As Long, j As Long, k As Long
    ListFolder p, files, nF, subs, nS
    For i = 1 To nF
        FileCopy p & files(i), dst & files(i): nCopied = nCopied + 1
    Next i
    For i = 1 To nS
        Dim f2() As String, s2() As String, nF2 As Long, nS2 As Long
        If Not FolderExists(dst & subs(i)) Then MkDir dst & subs(i)
        ListFolder p & subs(i), f2, nF2, s2, nS2
        For j = 1 To nF2
            FileCopy p & subs(i) & "\" & f2(j), dst & subs(i) & "\" & f2(j): nCopied = nCopied + 1
        Next j
        For j = 1 To nS2
            Dim f3() As String, s3() As String, nF3 As Long, nS3 As Long
            If Not FolderExists(dst & subs(i) & "\" & s2(j)) Then MkDir dst & subs(i) & "\" & s2(j)
            ListFolder p & subs(i) & "\" & s2(j), f3, nF3, s3, nS3
            For k = 1 To nF3
                FileCopy p & subs(i) & "\" & s2(j) & "\" & f3(k), dst & subs(i) & "\" & s2(j) & "\" & f3(k): nCopied = nCopied + 1
            Next k
        Next j
    Next i
    CopyTree = True
    Exit Function
Failed:
    LogError "MoveRoster (CopyTree " & p & ")", Err.Number, Err.Description
    CopyTree = False
End Function

' removes p's files and subfolders (two levels), leaving p itself: the old members' file then sees an empty inbox. Quiet.
Private Sub ClearTree(ByVal p As String)
    On Error Resume Next                            ' never raises (a file in use stays; SystemCheck's old-location line names what is left)
    p = EnsureSlash(p)
    If Not FolderExists(p) Then Exit Sub
    Dim files() As String, subs() As String, nF As Long, nS As Long, i As Long, j As Long, k As Long
    ListFolder p, files, nF, subs, nS
    For i = 1 To nF
        Kill p & files(i)
    Next i
    For i = 1 To nS
        Dim f2() As String, s2() As String, nF2 As Long, nS2 As Long
        ListFolder p & subs(i), f2, nF2, s2, nS2
        For j = 1 To nS2
            Dim f3() As String, s3() As String, nF3 As Long, nS3 As Long
            ListFolder p & subs(i) & "\" & s2(j), f3, nF3, s3, nS3
            For k = 1 To nF3
                Kill p & subs(i) & "\" & s2(j) & "\" & f3(k)
            Next k
            RmDir p & subs(i) & "\" & s2(j)
        Next j
        For j = 1 To nF2
            Kill p & subs(i) & "\" & f2(j)
        Next j
        RmDir p & subs(i)
    Next i
End Sub
