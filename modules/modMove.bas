Attribute VB_Name = "modMove"
Option Explicit

' ====================================================================
'  v29 (briefs/v29.md item 1; Graeme, Oct 27, 2026): THE FOLDER SPLIT AND THE GUIDED MOVE.
'  At go-live the master moves to a supervisor-only folder; the members' file, Signups\ and Requests\ go to the
'  members' folder. The folder settings stay constants: the ONE value the move stores in the workbook is the members'
'  folder (the hidden name MembersFolder, read by modPublish.PublishFolderPath when set), so the January cutover needs
'  no code install. Excel cannot open two files of one name, so the move cannot copy-and-reopen itself: MoveRoster
'  (admin, Alt+F8) checks the file, writes the members' path INTO this workbook and copies itself and the members'
'  template to the new master folder (so the copy carries the path), moves the inbox folders file by file (each file
'  deleted only once its copy is verified), renames the OLD members' file MOVED so nothing keeps writing beside it,
'  freezes THIS file with the hidden name MovedTo and saves, then says "close this, open the new one". Any failure
'  before the freeze rolls the stored values back and takes the copy away; this file stays live. The new file's first
'  admin open (MoveFirstOpen) marks the members' file dirty and asks for SystemCheck + IntegrityCheck + a publish.
'  A COPY / DRY RUN file runs the same steps in REHEARSAL mode into test folders: it writes the members' path flagged
'  "R|", which only a practice copy honours ("L|" for the live file), so a practice copy made from a moved live roster
'  never publishes into the real members' folder. Pure: StoredPathCore (suite mvstored), MoveTargetProblem (suite mvtarget).
'  Review round 1 (Oct 28): the inbox step runs only for ".\<sub>" settings and never onto itself; one rollback for every
'  failure incl. a raised SaveRestrictedCopy; the names are stored quote-safe and read back verified; the old members'
'  file is renamed; the frozen copy never saves on close (gNoSaveOnClose); SystemCheck's path block is its own box.
' ====================================================================
Public Const MV_VERSION As String = "Move v28 (Oct 28)"        ' new in the v29 build; stamped on the current release until the v29 release prep (test_drift)
Private Const NM_MEMBERS As String = "MembersFolder"            ' "L|<path>" (the live roster) or "R|<path>" (a rehearsal on a copy)
Private Const NM_MOVEDTO As String = "MovedTo"                  ' "<new full path>|<yyyy-mm-dd hh:nn>" in the OLD (frozen) file
Private Const NM_MOVEDFROM As String = "MovedFrom"              ' "<old full path>|<yyyy-mm-dd hh:nn>" in the NEW file
Private mPathBlock As String                                   ' SystemCheck's one line per path (its own box: the main box is capped at 1,024 characters)

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

' The two folders typed for a move: "" when they will do, else why not (one line). masterNow = the roster's folder today,
' membersNow = the members' folder in use today (the inbox must never be moved onto itself: review H1). Suite mvtarget.
Public Function MoveTargetProblem(ByVal masterNow As String, ByVal masterNew As String, ByVal membersNew As String, _
                                  ByVal masterReachable As Boolean, ByVal membersReachable As Boolean, _
                                  Optional ByVal membersNow As String = "") As String
    masterNow = LCase$(EnsureSlash(Trim$(masterNow))): masterNew = LCase$(EnsureSlash(Trim$(masterNew)))
    membersNew = LCase$(EnsureSlash(Trim$(membersNew))): membersNow = LCase$(EnsureSlash(Trim$(membersNow)))
    If Trim$(masterNew) = "" Or masterNew = "\" Then MoveTargetProblem = "no master folder was given": Exit Function
    If Trim$(membersNew) = "" Or membersNew = "\" Then MoveTargetProblem = "no members' folder was given": Exit Function
    If masterNew = masterNow Then MoveTargetProblem = "the new master folder is the folder the roster is in now": Exit Function
    If membersNew = masterNow Then MoveTargetProblem = "the new members' folder is the folder the roster is in now (the inbox would be moved onto itself)": Exit Function
    If membersNow <> "\" And membersNew = membersNow Then MoveTargetProblem = "the new members' folder is the members' folder in use now": Exit Function
    If Not masterReachable Then MoveTargetProblem = "the new master folder can't be reached or written: " & masterNew: Exit Function
    If Not membersReachable Then MoveTargetProblem = "the new members' folder can't be reached or written: " & membersNew: Exit Function
End Function

' ====================================================================
'  THE STORED VALUES (workbook-level hidden names; stored as ="text" with inner quotes doubled - review M3)
' ====================================================================
Private Function NameGet(ByVal key As String) As String
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then Exit Function
    Dim s As String: s = nmObj.RefersTo
    If Len(s) >= 3 And Left$(s, 2) = "=""" And Right$(s, 1) = """" Then NameGet = Replace(Mid$(s, 3, Len(s) - 3), """""", """")
End Function

' "" removes the name
Private Sub NameSet(ByVal key As String, ByVal v As String)
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If v = "" Then
        If Not nmObj Is Nothing Then nmObj.Delete
        Exit Sub
    End If
    Dim ref As String: ref = "=""" & Replace(v, """", """""") & """"
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:=ref, Visible:=False
    Else
        nmObj.RefersTo = ref
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

' the old full path of a roster that came out of a move ("" = never moved) and when
Public Function MovedFromPath() As String
    Dim s As String: s = NameGet(NM_MOVEDFROM)
    If InStr(s, "|") > 0 Then MovedFromPath = Left$(s, InStr(s, "|") - 1) Else MovedFromPath = s
End Function

Public Function MovedFromText() As String
    Dim s As String: s = NameGet(NM_MOVEDFROM)
    If s = "" Then Exit Function
    If InStr(s, "|") > 0 Then MovedFromText = MovedFromPath() & " on " & Mid$(s, InStr(s, "|") + 1) Else MovedFromText = s
End Function

' ====================================================================
'  REACHABLE / WRITABLE
' ====================================================================
' True when a small file can be made and removed in the folder (named by login: two admins at once never collide). Never raises.
Public Function FolderWritable(ByVal p As String) As Boolean
    On Error Resume Next                            ' never raises
    p = EnsureSlash(Trim$(p))
    If p = "" Or p = "\" Or Not FolderExists(p) Then Exit Function
    Dim f As String: f = p & "~" & LoginNameKey("rostercheck") & ".tmp"
    Dim ff As Integer: ff = FreeFile
    Err.Clear
    Open f For Output As #ff
    If Err.Number <> 0 Then Exit Function
    Print #ff, "roster write check"
    Close #ff
    FolderWritable = (Dir$(f) <> "")
    Kill f
End Function

' SystemCheck: one line per path (the resolved folder, reachable, writable) into the Paths block (MovePathBlock - its own
' box), a one-line count into info, every trouble into warn (SystemCheckOK stays versions-only); the members' file's
' presence and age; what a move left behind. Never raises.
Public Sub MovePathLines(ByRef warn As String, ByRef info As String)
    On Error Resume Next                            ' never raises
    mPathBlock = ""
    Dim nOk As Long, nAll As Long
    Dim master As String: master = EnsureSlash(ThisWorkbook.Path)
    Dim members As String: members = PublishFolderPath()
    PathLine warn, "master folder", master, True, nOk, nAll
    PathLine warn, "members' folder", members, True, nOk, nAll
    PathLine warn, "sign-ups", SignupFolderPath(), True, nOk, nAll
    PathLine warn, "requests", RequestFolderPath(), True, nOk, nAll
    PathLine warn, "backups", BackupFolderPath(), True, nOk, nAll
    PathLine warn, "archive (FINAL)", ResolveFolder(ARCHIVE_FOLDER), True, nOk, nAll
    PathLine warn, "closed periods", ResolveFolder(CLOSED_FOLDER), False, nOk, nAll
    PathLine warn, "scans", ResolveFolder(SCAN_FOLDER), False, nOk, nAll
    PathLine warn, "reports", EnsureSlash(ReportFolder()), False, nOk, nAll
    Dim kp As String: kp = Trim$(KRONOS_PATH)
    If kp <> "" Then
        If Not FolderExists(kp) Then
            If InStrRev(kp, "\") > 0 Then kp = Left$(kp, InStrRev(kp, "\"))      ' a file or a wildcard: its folder
        End If
        PathLine warn, "Kronos reports", kp, False, nOk, nAll
    End If
    If Trim$(GUIDE_PATH) <> "" Then
        Dim gFound As Boolean: gFound = False
        Err.Clear
        gFound = (Dir$(GUIDE_PATH) <> "")
        If Err.Number <> 0 Then gFound = False
        Err.Clear
        If gFound Then mPathBlock = mPathBlock & "guide PDF: " & GUIDE_PATH & vbNewLine Else warn = warn & "- the guide PDF can't be found: " & GUIDE_PATH & vbNewLine
    End If
    Dim el As String: el = ErrorLogFilePath()
    If el <> "" Then mPathBlock = mPathBlock & "error log: " & el & IIf(Dir$(el) <> "", "", " (no file yet)") & vbNewLine
    ' the members' file itself
    Dim mf As String: mf = members & IIf(IsDryRunName(ThisWorkbook.Name), "PRACTICE - ", "") & MEMBER_FILE_NAME
    If Dir$(mf) <> "" Then
        mPathBlock = mPathBlock & "members' file present, saved " & Format$(FileDateTime(mf), "ddd mmm d h:mm AM/PM") & vbNewLine
    Else
        warn = warn & "- no members' file in the members' folder yet: publish it (" & mf & ")" & vbNewLine
    End If
    If StoredMembersFolder() <> "" Then mPathBlock = mPathBlock & "the members' folder was set by the guided move (PUBLISH_FOLDER is the default only)" & vbNewLine
    If MovedFromText() <> "" And Not IsDryRunName(ThisWorkbook.Name) Then
        mPathBlock = mPathBlock & "this roster moved from " & MovedFromText() & vbNewLine
        Dim oldPath As String: oldPath = MovedFromPath()
        If oldPath <> "" Then
            If Dir$(oldPath) <> "" Then warn = warn & "- the old copy still sits at " & oldPath & " under its live name: rename it MOVED " & Format$(Date, "yyyy-mm-dd") & " - ... (the runbook)" & vbNewLine
        End If
    End If
    info = info & "Paths: " & nOk & " of " & nAll & " folders reachable" & IIf(nOk = nAll, "", " - see the CHECK lines") & " (SystemCheck's second box lists them)" & vbNewLine
End Sub

Public Function MovePathBlock() As String
    MovePathBlock = mPathBlock
End Function

Private Sub PathLine(ByRef warn As String, ByVal what As String, ByVal p As String, ByVal needWrite As Boolean, ByRef nOk As Long, ByRef nAll As Long)
    On Error Resume Next                            ' never raises
    If Trim$(p) = "" Or p = "\" Then mPathBlock = mPathBlock & what & ": (not set)" & vbNewLine: Exit Sub
    nAll = nAll + 1
    If Not FolderExists(p) Then
        warn = warn & "- " & what & " can't be reached: " & p & vbNewLine
        mPathBlock = mPathBlock & what & ": " & p & "  - NOT REACHABLE" & vbNewLine
    ElseIf needWrite And Not FolderWritable(p) Then
        warn = warn & "- " & what & " is not writable for you: " & p & vbNewLine
        mPathBlock = mPathBlock & what & ": " & p & "  - reachable, NOT writable" & vbNewLine
    Else
        nOk = nOk + 1
        mPathBlock = mPathBlock & what & ": " & p & IIf(needWrite, "  - reachable, writable", "  - reachable") & vbNewLine
    End If
End Sub

' ====================================================================
'  THE GUIDED MOVE (admin, Alt+F8; a cutover step, never an install step)
' ====================================================================
' the sub-folder an inbox setting names beside the members' file ("Signups" for ".\Signups"); "" = the setting is off,
' "." (the whole members' folder: never moved as an inbox) or absolute (not ours to move) - review H2
Private Function InboxSub(ByVal setting As String) As String
    setting = Trim$(setting)
    If Left$(setting, 2) <> ".\" Then Exit Function
    Dim s As String: s = Mid$(setting, 3)
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    If s = "" Or InStr(s, "\") > 0 Or InStr(s, "..") > 0 Then Exit Function
    InboxSub = s
End Function

Public Sub MoveRoster()
    Dim prevMembers As String, prevFrom As String, newFull As String, madeCopy As Boolean, wasSaved As Boolean
    Dim masterNew As String, membersNew As String
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
    Dim membersNow As String: membersNow = PublishFolderPath()
    Dim subS As String: subS = InboxSub(SIGNUP_FOLDER)
    Dim subR As String: subR = InboxSub(REQUEST_FOLDER)
    Dim oldSignups As String, oldRequests As String
    If subS <> "" Then oldSignups = membersNow & subS & "\"
    If subR <> "" Then oldRequests = membersNow & subR & "\"
    masterNew = Trim$(InputBox(IIf(rehearsal, "REHEARSAL: a TEST folder for the master (nothing on P:)." & vbNewLine & vbNewLine, "") & _
                               "The NEW folder for the master roster (the supervisor-only folder). It must exist already.", ttl, masterNow))
    If masterNew = "" Then Exit Sub
    membersNew = Trim$(InputBox(IIf(rehearsal, "REHEARSAL: a TEST folder for the members' side." & vbNewLine & vbNewLine, "") & _
                                "The NEW folder for the members' file, " & IIf(subS <> "", subS & "\ ", "") & IIf(subR <> "", "and " & subR & "\ ", "") & "(the members' folder). It must exist already; IT sets the members' rights on it.", ttl, ""))
    If membersNew = "" Then Exit Sub
    masterNew = EnsureSlash(masterNew): membersNew = EnsureSlash(membersNew)
    Dim prob As String
    prob = MoveTargetProblem(masterNow, masterNew, membersNew, FolderWritable(masterNew), FolderWritable(membersNew), membersNow)
    If prob <> "" Then MsgBox "Nothing moved: " & prob & ".", vbExclamation, ttl: Exit Sub
    newFull = masterNew & ThisWorkbook.Name
    If Dir$(newFull) <> "" Then MsgBox "Nothing moved: " & newFull & " exists already. Move or rename it first.", vbExclamation, ttl: Exit Sub
    Dim nIn As Long: nIn = CountTree(oldSignups) + CountTree(oldRequests)
    Dim inboxNote As String
    If subS = "" Then inboxNote = inboxNote & vbNewLine & "  (SIGNUP_FOLDER is not a .\<folder> setting: the sign-up files are NOT moved - check it)"
    If subR = "" Then inboxNote = inboxNote & vbNewLine & "  (REQUEST_FOLDER is not a .\<folder> setting: the request files are NOT moved - check it)"
    If subS <> "" And Not FolderExists(membersNew & subS & "\") Then inboxNote = inboxNote & vbNewLine & "  (" & subS & "\ does not exist yet in the members' folder: it will be made and inherits the folder's rights - members need write rights on it)"
    If subR <> "" And Not FolderExists(membersNew & subR & "\") Then inboxNote = inboxNote & vbNewLine & "  (" & subR & "\ does not exist yet in the members' folder: it will be made and inherits the folder's rights - members need write rights on it)"
    If MsgBox(IIf(rehearsal, "REHEARSAL on this practice copy - the real roster and P: are not touched." & vbNewLine & vbNewLine, "") & _
              "The move:" & vbNewLine & _
              "  1. copy this roster and the members' template to " & masterNew & vbNewLine & _
              "  2. move the sign-up and request files (" & nIn & ") to " & membersNew & inboxNote & vbNewLine & _
              "  3. write the members' folder into the new copy, so it publishes there" & vbNewLine & _
              "  4. rename the old members' file MOVED, so nothing keeps writing beside it" & vbNewLine & _
              "  5. freeze THIS file (nothing can be logged, saved or published in it again)" & vbNewLine & _
              "  6. you close this file and open the new one; its first open asks for SystemCheck, IntegrityCheck and a publish" & vbNewLine & vbNewLine & _
              "Backups, Closed, Scans, Practice and the error log start again beside the new master; the old ones stay here." & vbNewLine & vbNewLine & _
              "Go on?", vbOKCancel + vbQuestion + vbDefaultButton2, ttl) <> vbOK Then Exit Sub

    Application.StatusBar = "Moving the roster ..."
    wasSaved = ThisWorkbook.Saved
    prevMembers = NameGet(NM_MEMBERS): prevFrom = NameGet(NM_MOVEDFROM)
    ' 3 first: the values ride inside the copy; read back and checked before anything is copied (review M3)
    Dim stamp As String: stamp = Format$(Now, "yyyy-mm-dd hh:nn")
    NameSet NM_MEMBERS, IIf(rehearsal, "R|", "L|") & membersNew
    NameSet NM_MOVEDFROM, ThisWorkbook.FullName & "|" & stamp
    If LCase$(EnsureSlash(StoredPathCore(NameGet(NM_MEMBERS), rehearsal))) <> LCase$(membersNew) Then
        MoveRollback prevMembers, prevFrom, "", False, wasSaved
        Application.StatusBar = False
        MsgBox "Nothing moved: the members' folder could not be stored in the workbook (a very long path, or an odd character). This file is unchanged.", vbExclamation, ttl
        Exit Sub
    End If
    ' 1. the copy (the restricted view, as every copy) and the template
    SaveRestrictedCopy newFull                       ' raises on a failed SaveCopyAs: Oops rolls back
    madeCopy = True
    If Dir$(newFull) = "" Then GoTo CopyFailed
    If FileLen(newFull) < 1000 Then GoTo CopyFailed
    Dim tplNote As String
    If MemberTemplateExists() Then
        On Error Resume Next
        Err.Clear
        FileCopy MemberTemplatePath(), masterNew & MEMBER_TEMPLATE_NAME
        If Err.Number <> 0 Then tplNote = vbNewLine & "The members' template could NOT be copied (" & Err.Description & "): copy " & MEMBER_TEMPLATE_NAME & " to " & masterNew & " by hand before the first publish."
        On Error GoTo Oops
    End If
    ' 2. the inbox folders, file by file (a file is deleted only once its copy is verified - review H4)
    Dim nMoved As Long, note As String
    If oldSignups <> "" Then
        If Not MoveTree(oldSignups, membersNew & subS & "\", nMoved) Then note = "the sign-up files did not all move (" & nMoved & " moved so far)": GoTo InboxFailed
    End If
    If oldRequests <> "" Then
        If Not MoveTree(oldRequests, membersNew & subR & "\", nMoved) Then note = "the request files did not all move (" & nMoved & " moved so far)": GoTo InboxFailed
    End If
    ' 4. the old members' file and its emptied inbox folders (review H5): nothing keeps writing beside the old master
    Dim oldMf As String, mfNote As String
    oldMf = membersNow & IIf(rehearsal, "PRACTICE - ", "") & MEMBER_FILE_NAME
    If LCase$(membersNow) <> LCase$(membersNew) Then
        On Error Resume Next
        If Dir$(oldMf) <> "" Then
            Err.Clear
            Name oldMf As membersNow & "MOVED " & Format$(Date, "yyyy-mm-dd") & " - " & IIf(rehearsal, "PRACTICE - ", "") & MEMBER_FILE_NAME
            If Err.Number <> 0 Then mfNote = vbNewLine & "The old members' file could NOT be renamed (" & Err.Description & " - open on a PC?): rename it MOVED by hand, or members keep opening it."
        End If
        Err.Clear
        If oldSignups <> "" Then RmDir Left$(oldSignups, Len(oldSignups) - 1)
        If oldRequests <> "" Then RmDir Left$(oldRequests, Len(oldRequests) - 1)
        Err.Clear
        On Error GoTo Oops
    End If
    ' 5. freeze this file: the old stored value back (this file never publishes anywhere again), MovedTo set, MovedFrom off
    NameSet NM_MEMBERS, prevMembers
    NameSet NM_MOVEDFROM, prevFrom
    NameSet NM_MOVEDTO, newFull & "|" & stamp
    ThisWorkbook.Saved = False
    If Not SaveIfWritable() Then
        MsgBox "The new roster is in place at " & newFull & ", but THIS file could not be saved with its MOVED marker. " & _
               "Close it WITHOUT saving, then rename it MOVED " & Format$(Date, "yyyy-mm-dd") & " - ... by hand so nobody opens it as the live roster.", vbExclamation, ttl
    End If
    Application.StatusBar = False
    MsgBox IIf(rehearsal, "REHEARSAL done. ", "Done. ") & "Now:" & vbNewLine & vbNewLine & _
           "  1. Close this file (it is frozen: every open says it moved)." & vbNewLine & _
           "  2. Open " & newFull & vbNewLine & _
           "  3. Its first open asks for SystemCheck (every path), IntegrityCheck and Publish the members' file." & vbNewLine & _
           "  4. Rename this file  MOVED " & Format$(Date, "yyyy-mm-dd") & " - " & ThisWorkbook.Name & "  in " & masterNow & vbNewLine & vbNewLine & _
           "Members' folder: " & membersNew & vbNewLine & "Sign-up and request files moved: " & nMoved & IIf(nMoved <> nIn, " (of " & nIn & " seen)", "") & tplNote & mfNote, vbInformation, ttl
    Exit Sub
CopyFailed:
    MoveRollback prevMembers, prevFrom, newFull, madeCopy, wasSaved
    Application.StatusBar = False
    MsgBox "Nothing moved: the copy at " & newFull & " could not be made (the folder, or the file is open there). This file is unchanged.", vbExclamation, ttl
    Exit Sub
InboxFailed:
    MoveRollback prevMembers, prevFrom, newFull, madeCopy, wasSaved
    Application.StatusBar = False
    MsgBox "Nothing moved: " & note & ". The copy of the roster was taken back and this file is unchanged. " & _
           IIf(nMoved > 0, nMoved & " sign-up / request file(s) had already moved to " & membersNew & ": move them back by hand (Signups\ / Requests\), then ", "") & _
           "check the members' folder's rights and run the move again.", vbExclamation, ttl
    Exit Sub
Oops:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    On Error Resume Next
    MoveRollback prevMembers, prevFrom, newFull, madeCopy, wasSaved    ' (review H3) a raised SaveRestrictedCopy too: the live file keeps its old values
    Application.StatusBar = False
    LogError "MoveRoster", en, em
    MsgBox "The move stopped: " & em & vbNewLine & vbNewLine & "This file keeps its old settings and stays live; " & _
           IIf(madeCopy, "the copy at " & newFull & " was removed. ", "no copy was made. ") & "Run SystemCheck, then the move again.", vbExclamation, ttl
End Sub

' every failure before the freeze: the stored values as they were, the copy and the template taken away when this run
' made them, the Saved flag back
Private Sub MoveRollback(ByVal prevMembers As String, ByVal prevFrom As String, ByVal newFull As String, ByVal madeCopy As Boolean, ByVal wasSaved As Boolean)
    On Error Resume Next                            ' never raises
    NameSet NM_MEMBERS, prevMembers
    NameSet NM_MOVEDFROM, prevFrom
    If madeCopy And newFull <> "" Then
        If Dir$(newFull) <> "" Then Kill newFull
        Dim tpl As String: tpl = Left$(newFull, InStrRev(newFull, "\")) & MEMBER_TEMPLATE_NAME
        If Dir$(tpl) <> "" Then Kill tpl
    End If
    ThisWorkbook.Saved = wasSaved
End Sub

' the first ADMIN open of a roster that came out of a move: the members' file is marked out of date (it must be published
' into the new folder) and one box, once per login, says what to run. Called by OpenChecks. A practice copy of a moved
' roster carries MovedFrom too and is left alone. Never raises.
Public Sub MoveFirstOpen()
    On Error Resume Next                            ' never raises
    If MovedFromText() = "" Or MovedToPath() <> "" Then Exit Sub
    If IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name) Then Exit Sub
    If Not IsAdmin() Or ThisWorkbook.ReadOnly Then Exit Sub
    Dim key As String: key = LoginNameKey("MoveSeen")
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If Not nmObj Is Nothing Then Exit Sub
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    ThisWorkbook.Names.Add Name:=key, RefersTo:="=""seen""", Visible:=False
    MarkMembersDirty
    If wasSaved Then ThisWorkbook.Saved = True
    MsgBox "First open since the roster moved from " & MovedFromText() & "." & vbNewLine & vbNewLine & _
           "Run, in this order: SystemCheck (every path reachable and writable - its second box), IntegrityCheck, then Publish the members' file into " & _
           PublishFolderPath() & ". Then rename the old roster file MOVED " & Format$(Date, "yyyy-mm-dd") & " - ... if that is not done.", vbInformation, "Moved roster"
End Sub

' ====================================================================
'  FILE HELPERS (Dir cannot nest: names are collected first, then walked; two levels - the inbox folders hold Done\<posting>\ and done\)
' ====================================================================
Private Function OtherRosterOpen() As Boolean
    Dim w As Workbook
    For Each w In Application.Workbooks
        If Not w Is ThisWorkbook Then
            If StrComp(w.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then OtherRosterOpen = True: Exit Function
        End If
    Next w
End Function

' the files in a folder (one level) and its subfolders, as two arrays; counts returned. A name GetAttr cannot read counts as a file.
Private Sub ListFolder(ByVal p As String, ByRef files() As String, ByRef nF As Long, ByRef subs() As String, ByRef nS As Long)
    On Error Resume Next                            ' never raises
    nF = 0: nS = 0
    ReDim files(1 To 1): ReDim subs(1 To 1)
    p = EnsureSlash(p)
    If p = "" Or p = "\" Or Not FolderExists(p) Then Exit Sub
    Dim f As String, isDir As Boolean, att As Long
    f = Dir$(p & "*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            Err.Clear
            att = GetAttr(p & f)
            isDir = (Err.Number = 0) And ((att And vbDirectory) = vbDirectory)
            Err.Clear
            If isDir Then
                nS = nS + 1: ReDim Preserve subs(1 To nS): subs(nS) = f
            Else
                nF = nF + 1: ReDim Preserve files(1 To nF): files(nF) = f
            End If
        End If
        f = Dir$()
    Loop
End Sub

' every file under p (two levels)
Private Function CountTree(ByVal p As String) As Long
    If Trim$(p) = "" Then Exit Function
    Dim files() As String, subs() As String, nF As Long, nS As Long, i As Long, j As Long
    Dim f2() As String, s2() As String, nF2 As Long, nS2 As Long
    Dim f3() As String, s3() As String, nF3 As Long, nS3 As Long
    ListFolder p, files, nF, subs, nS
    CountTree = nF
    For i = 1 To nS
        ListFolder EnsureSlash(p) & subs(i), f2, nF2, s2, nS2
        CountTree = CountTree + nF2
        For j = 1 To nS2
            ListFolder EnsureSlash(p) & subs(i) & "\" & s2(j), f3, nF3, s3, nS3
            CountTree = CountTree + nF3
        Next j
    Next i
End Function

' one file: copied, the copy checked (present, same size), then the original removed; False when the copy is not there
Private Function MoveOne(ByVal src As String, ByVal dst As String) As Boolean
    On Error Resume Next                            ' never raises
    Err.Clear
    FileCopy src, dst
    If Err.Number <> 0 Then Exit Function
    If Dir$(dst) = "" Then Exit Function
    If FileLen(dst) <> FileLen(src) Then Exit Function
    Kill src
    MoveOne = True
End Function

' p's files and subfolders (two levels) moved under dst, making the folders, the emptied subfolders removed; nMoved counts
' files; False on the first file that would not move (the ones before it are in dst, the rest still in p)
Private Function MoveTree(ByVal p As String, ByVal dst As String, ByRef nMoved As Long) As Boolean
    On Error GoTo Failed
    p = EnsureSlash(p): dst = EnsureSlash(dst)
    If p = "" Or p = "\" Or dst = "" Or dst = "\" Then Exit Function
    If LCase$(p) = LCase$(dst) Then Exit Function    ' never onto itself
    If Not FolderExists(dst) Then MkDir Left$(dst, Len(dst) - 1)
    If Not FolderExists(dst) Then Exit Function
    If Not FolderExists(p) Then MoveTree = True: Exit Function
    Dim files() As String, subs() As String, nF As Long, nS As Long, i As Long, j As Long, k As Long
    Dim f2() As String, s2() As String, nF2 As Long, nS2 As Long
    Dim f3() As String, s3() As String, nF3 As Long, nS3 As Long
    ListFolder p, files, nF, subs, nS
    For i = 1 To nF
        If Not MoveOne(p & files(i), dst & files(i)) Then Exit Function
        nMoved = nMoved + 1
    Next i
    For i = 1 To nS
        If Not FolderExists(dst & subs(i)) Then MkDir dst & subs(i)
        ListFolder p & subs(i), f2, nF2, s2, nS2
        For j = 1 To nF2
            If Not MoveOne(p & subs(i) & "\" & f2(j), dst & subs(i) & "\" & f2(j)) Then Exit Function
            nMoved = nMoved + 1
        Next j
        For j = 1 To nS2
            If Not FolderExists(dst & subs(i) & "\" & s2(j)) Then MkDir dst & subs(i) & "\" & s2(j)
            ListFolder p & subs(i) & "\" & s2(j), f3, nF3, s3, nS3
            For k = 1 To nF3
                If Not MoveOne(p & subs(i) & "\" & s2(j) & "\" & f3(k), dst & subs(i) & "\" & s2(j) & "\" & f3(k)) Then Exit Function
                nMoved = nMoved + 1
            Next k
            On Error Resume Next
            RmDir p & subs(i) & "\" & s2(j)
            On Error GoTo Failed
        Next j
        On Error Resume Next
        RmDir p & subs(i)
        On Error GoTo Failed
    Next i
    MoveTree = True
    Exit Function
Failed:
    LogError "MoveRoster (MoveTree " & p & ")", Err.Number, Err.Description
    MoveTree = False
End Function
