Attribute VB_Name = "modSystem"
Option Explicit

' ====================================================================
'  SYSTEM CHECK + OPENING CHECKS (v20, v21)
'
'  Alt+F8 > SystemCheck  - run after every update. Shows each module's
'  version, flags leftovers ("modCalendar1", modFixShift ...) and checks
'  the sheets, the OT Log layout, holidays, the Kronos import, blank Home
'  Units and the daily backup.
'  Listing modules and checking the pasted form code needs Excel's
'  "Trust access to the VBA project object model" (File > Options >
'  Trust Center > Trust Center Settings > Macro Settings). Without it,
'  the module versions are still shown.
'
'  OpenChecks (from Workbook_Open):
'   - read-only copy: says who has the file open for editing
'   - file opened from SharePoint / OneDrive: AutoSave off (each submit
'     already saves), so two people's edits can't be merged silently
'
'  BackupOnOpen (v21, from Workbook_Open): the first admin open of the
'  day saves a dated copy into BACKUP_FOLDER and prunes old ones. It
'  never blocks opening: no folder = skip quietly, unreachable = warn.
' ====================================================================

Public Const SYS_VERSION As String = "System v24 (Jan 5)"
Public Const EXPECTED_RELEASE As String = "v24"
Private Const N_MODULES As Long = 18
' ThisWorkbook and each form are checked against the release that last changed them
Private Const THISWB_RELEASE As String = "v24"
Private Const FORM_RELEASES As String = "frmOvertimeEntry|v22,frmNewPosting|v21,frmPostingEntry|v24,frmShiftSwitch|v21," & _
                                        "frmTimeOff|v22.1,frmLeave|v22,frmSickCall|v22.1,frmConvert|v22,frmEmployee|v24"

Private mBackupWarned As Boolean


' ---------------- SystemCheck ----------------
' Problems come first; the message stays short enough for one MsgBox.
Public Sub SystemCheck()
    Dim warn As String, info As String, nOk As Long

    ' ---- modules: each one's version constant (a missing module won't compile) ----
    VerCheck "modOvertimeEntry", CORE_VERSION, warn, nOk
    VerCheck "modCalendar", CAL_VERSION, warn, nOk
    VerCheck "modHolidays", HOL_VERSION, warn, nOk
    VerCheck "modShiftCycle", ROT_VERSION, warn, nOk
    VerCheck "modShiftSwitch", SS_VERSION, warn, nOk
    VerCheck "modPostings", POST_VERSION, warn, nOk
    VerCheck "ImportKronosData", KR_VERSION, warn, nOk
    VerCheck "modUI", UI_VERSION, warn, nOk
    VerCheck "modSystem", SYS_VERSION, warn, nOk
    VerCheck "modGradeLog", GL_VERSION, warn, nOk
    VerCheck "modSelfTest", ST_VERSION, warn, nOk
    VerCheck "Call_Out_Roster", CO_VERSION, warn, nOk
    VerCheck "modPublish", PUB_VERSION, warn, nOk
    VerCheck "modLeave", LV_VERSION, warn, nOk
    VerCheck "modAttendance", AT_VERSION, warn, nOk
    VerCheck "modIntegrity", IC_VERSION, warn, nOk
    VerCheck "modRollover", RO_VERSION, warn, nOk
    VerCheck "modPeople", PPL_VERSION, warn, nOk
    info = "Modules: " & nOk & " of " & N_MODULES & " on " & EXPECTED_RELEASE & vbNewLine
    If InStr(ThisWorkbookVersion(), THISWB_RELEASE) = 0 Then
        warn = warn & "- ThisWorkbook code is older than " & THISWB_RELEASE & ": re-paste ThisWorkbook.txt" & vbNewLine
    End If

    ' ---- VBA project (needs trusted access) ----
    Dim proj As Object, comp As Object, nm As String, trusted As Boolean, cnt As Long
    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    trusted = Not proj Is Nothing
    If trusted Then
        cnt = proj.VBComponents.Count
        If Err.Number <> 0 Then trusted = False
    End If
    On Error GoTo 0
    If Not trusted Then
        info = info & "Forms / leftovers: not checked (VBA project access is off)" & vbNewLine
    Else
        Dim compNames As String, f As Variant
        For Each comp In proj.VBComponents
            nm = comp.Name
            compNames = compNames & "|" & LCase$(nm) & "|"
            If nm Like "*[0-9]" And Not nm Like "Sheet*" And Not nm Like "Module*" And Not nm Like "Class*" And Not nm Like "Chart*" Then
                warn = warn & "- '" & nm & "' looks like a duplicate from an import: remove it" & vbNewLine
            End If
            Select Case LCase$(nm)
                Case "modfixshift", "refusal_report", "module1", "class1"
                    warn = warn & "- '" & nm & "' is left over from an earlier version: remove it" & vbNewLine
            End Select
        Next comp
        For Each f In Array("frmShiftSwitch", "frmOvertimeEntry", "frmNewPosting", "frmPostingEntry", "frmTimeOff", _
                            "frmLeave", "frmSickCall", "frmConvert", "frmEmployee")
            If InStr(compNames, "|" & LCase$(CStr(f)) & "|") = 0 Then
                warn = warn & "- form " & f & " is missing (Insert > UserForm, name it, paste " & f & ".txt)" & vbNewLine
            End If
        Next f
        Dim fr As Variant, fNm As String, fRel As String
        For Each fr In Split(FORM_RELEASES, ",")
            fNm = Split(CStr(fr), "|")(0): fRel = Split(CStr(fr), "|")(1)
            If InStr(compNames, "|" & LCase$(fNm) & "|") > 0 Then
                If InStr(FormMarker(proj, fNm), fRel) = 0 Then
                    warn = warn & "- " & fNm & " code is older than " & fRel & ": re-paste " & fNm & ".txt" & vbNewLine
                End If
            End If
        Next fr
        info = info & "Forms and leftovers: checked" & vbNewLine
    End If

    ' ---- data ----
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then
        warn = warn & "- OT Log table not found" & vbNewLine
    ElseIf Not LogLayoutCurrent(lo) Then
        warn = warn & "- OT Log is in the old layout: run UpgradeOTLogLayout" & vbNewLine
    Else
        info = info & "OT Log: " & Format$(lo.ListRows.Count, "#,##0") & " rows" & vbNewLine
    End If
    ' ---- v23: the roster year, year-end folders, integrity ----
    On Error Resume Next
    info = info & "Roster year: " & RosterYear() & " (" & (PeriodLastCol() - C_PERIOD1 + 1) & " pay periods from " & _
           Format$(RosterYearFrom(), "mmm d, yyyy") & ")" & vbNewLine
    On Error GoTo 0
    If Not FolderExists(ResolveFolder(ARCHIVE_FOLDER)) Then
        warn = warn & "- archive folder for the year-end frozen copy can't be reached: " & ResolveFolder(ARCHIVE_FOLDER) & vbNewLine
    End If
    Dim rf As String, rfParent As String
    rf = EnsureSlash(ReportFolder())
    If Not FolderExists(rf) And Len(rf) > 3 Then
        rfParent = Left$(rf, InStrRev(rf, "\", Len(rf) - 1))
        If FolderExists(rfParent) Then
            warn = warn & "- this year's reports folder doesn't exist yet: " & rf & " (create it; PDFs go next to the workbook until then)" & vbNewLine
        End If
    End If
    If IntegrityStatusText() <> "" Then info = info & IntegrityStatusText() & vbNewLine
    If IntegrityHasProblems() Then warn = warn & "- the integrity check found problems: run IntegrityCheck" & vbNewLine
    info = info & KronosStatusText() & vbNewLine
    If KronosIsStale() Then warn = warn & "- Kronos data is old: run Import Kronos" & vbNewLine
    info = info & DataThroughText() & vbNewLine
    If RosterLastRow() >= R_FIRST And Not ProvisionalLive() Then
        warn = warn & "- Total Hours doesn't include provisional accepted OT yet: run RefreshRosterFormulas" & vbNewLine
    End If
    Dim hv As Variant, k As Long
    hv = HolidaysToConfirm(Date, Date + HOL_CONFIRM_DAYS - 1)
    If IsArray(hv) Then
        For k = LBound(hv) To UBound(hv)
            warn = warn & "- " & Mid$(hv(k), 12) & " " & Left$(hv(k), 10) & _
                   " is on a weekend: enter the weekday off on the Holidays sheet" & vbNewLine
        Next k
    End If
    info = info & "Holidays sheet: " & IIf(SheetExists(HOL_SHEET), "present", "created on next open (admin)") & vbNewLine

    ' ---- v21: Home Unit, grade log, backup ----
    Dim hu As Variant, nHU As Long, huTxt As String
    On Error Resume Next
    hu = BlankHomeUnits()
    On Error GoTo 0
    If UCase$(Trim$(SafeStr(RosterSheet().Cells(2, C_HOME).Value))) <> "HOME UNIT" Then
        warn = warn & "- Home Unit column not set up: run SetupHomeUnit" & vbNewLine
    ElseIf IsArray(hu) Then
        nHU = UBound(hu) - LBound(hu) + 1
        For k = LBound(hu) To UBound(hu)
            If k - LBound(hu) < 6 Then huTxt = huTxt & IIf(huTxt = "", "", ", ") & hu(k)
        Next k
        If nHU > 6 Then huTxt = huTxt & " ... and " & (nHU - 6) & " more"
        warn = warn & "- " & nHU & " with no Home Unit (column B): " & huTxt & vbNewLine
    End If
    info = info & "Grade change log: " & IIf(SheetExists(GL_SHEET), "present", "created on the first change") & vbNewLine
    ' v22: absences table columns, imported lines not booked
    If Not AbsenceColumnsCurrent() Then warn = warn & "- Absences table needs its v22 columns: run UpgradeAbsenceColumns" & vbNewLine
    Dim nImp As Long
    On Error Resume Next
    nImp = ImportedNotBookedCount()
    On Error GoTo 0
    If nImp > 0 Then warn = warn & "- " & nImp & " imported vacation line(s) in the next " & IMPORT_WATCH_DAYS & " days not booked yet: Convert imports" & vbNewLine
    ' v24: crew history and the Former sheet
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        warn = warn & "- Crew History / Former sheets not set up: run SetupPeople" & vbNewLine
    Else
        info = info & "Crew History: " & Format$(CrewTable(False).ListRows.Count, "#,##0") & " rows" & vbNewLine
    End If
    If Not PostingListCurrent() Then warn = warn & "- posting lists need their Signup Hours column: run SetupPeople" & vbNewLine
    ' M1: members' file
    If Not MemberTemplateExists() Then
        warn = warn & "- members' template not found next to this workbook: " & MEMBER_TEMPLATE_NAME & " (see the M1 install notes)" & vbNewLine
    End If
    info = info & MembersStatusText() & vbNewLine
    Dim bkWarn As String, bkInfo As String
    BackupStatus bkWarn, bkInfo
    warn = warn & bkWarn
    info = info & bkInfo
    If ThisWorkbook.ReadOnly Then warn = warn & "- This copy is read-only" & vbNewLine

    If Len(warn) > 700 Then warn = Left$(warn, InStrRev(Left$(warn, 700), vbNewLine)) & "- ...and more" & vbNewLine
    Dim msg As String
    If warn = "" Then
        msg = "Everything looks right." & vbNewLine & vbNewLine & info
    Else
        msg = "CHECK:" & vbNewLine & warn & vbNewLine & info
    End If
    MsgBox msg, IIf(warn = "", vbInformation, vbExclamation), "System check  (" & EXPECTED_RELEASE & ")"
End Sub

Private Sub VerCheck(ByVal modName As String, ByVal ver As String, ByRef warn As String, ByRef nOk As Long)
    If InStr(ver, EXPECTED_RELEASE) > 0 Then
        nOk = nOk + 1
    Else
        warn = warn & "- " & modName & " is " & ver & ": remove it and import the " & EXPECTED_RELEASE & " file" & vbNewLine
    End If
End Sub

' ThisWorkbook exposes CodeVersion from v20 on; older code doesn't have it
Private Function ThisWorkbookVersion() As String
    Dim o As Object: Set o = ThisWorkbook
    On Error Resume Next
    ThisWorkbookVersion = o.CodeVersion
    On Error GoTo 0
    If ThisWorkbookVersion = "" Then ThisWorkbookVersion = "before v20"
End Function

' the "' @version ..." line in a form's code, or "no version line"
Private Function FormMarker(ByVal proj As Object, ByVal formName As String) As String
    On Error Resume Next
    Dim cm As Object: Set cm = proj.VBComponents(formName).CodeModule
    Dim n As Long: n = cm.CountOfLines
    If n > 60 Then n = 60
    Dim i As Long, ln As String
    For i = 1 To n
        ln = Trim$(cm.Lines(i, 1))
        If Left$(ln, 10) = "' @version" Then FormMarker = Trim$(Mid$(ln, 11)): Exit Function
    Next i
    FormMarker = "no version line (before v20)"
End Function

' ---------------- opening checks ----------------
Public Sub OpenChecks()
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then
        Dim who As String: who = LockOwner()
        MsgBox IIf(who <> "", who & " has", "Someone else has") & " the roster open for editing, so this copy is read-only." & _
               vbNewLine & vbNewLine & "You can look at everything, but nothing can be logged or saved." & vbNewLine & _
               "Close it and open it again once they're out.", vbInformation, "Read-only"
        Exit Sub
    End If
    ' SharePoint / OneDrive: AutoSave off - every submit saves anyway, and
    ' two people's copies must never be merged in the background.
    If LCase$(Left$(ThisWorkbook.FullName, 4)) = "http" Then
        Dim wbo As Object: Set wbo = ThisWorkbook      ' late-bound: older Excel has no AutoSaveOn
        If wbo.AutoSaveOn Then wbo.AutoSaveOn = False
    End If
End Sub

' Name of whoever holds the edit lock on a network share (from Excel's
' "~$" owner file next to the workbook), or "" if it can't be read.
Public Function LockOwner() As String
    On Error GoTo Done
    If LCase$(Left$(ThisWorkbook.FullName, 4)) = "http" Then Exit Function
    Dim p As String: p = ThisWorkbook.Path & Application.PathSeparator & "~$" & ThisWorkbook.Name
    If Dir$(p, vbHidden) = "" Then Exit Function
    Dim f As Integer: f = FreeFile
    Dim b() As Byte, n As Long
    Open p For Binary Access Read Shared As #f
    n = LOF(f)
    If n < 2 Then Close #f: Exit Function
    If n > 200 Then n = 200
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    Dim L As Long: L = b(0)
    If L < 1 Or L > n - 1 Then Exit Function
    Dim i As Long, s As String
    For i = 1 To L
        s = s & Chr$(b(i))
    Next i
    LockOwner = Trim$(s)
    Exit Function
Done:
    On Error Resume Next
    Close #f
End Function


' ====================================================================
'  DAILY BACKUP (v21)
' ====================================================================
' Today's backup file name, e.g. "Call Out Roster 2026-10-31.xlsm"
Public Function BackupFileName(ByVal d As Date) As String
    BackupFileName = BACKUP_BASENAME & " " & Format$(d, "yyyy-mm-dd") & ".xlsm"
End Function

' The date in a backup file name that matches the exact pattern, or 0
Public Function BackupFileDate(ByVal fileName As String) As Date
    Dim pat As String: pat = BACKUP_BASENAME & " ####-##-##.xlsm"
    If Len(fileName) <> Len(pat) Then Exit Function
    If Not fileName Like pat Then Exit Function
    Dim s As String: s = Mid$(fileName, Len(BACKUP_BASENAME) + 2, 10)
    On Error Resume Next
    BackupFileDate = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
End Function

' True when the backup should be skipped without a word
Private Function BackupOff() As Boolean
    ' a practice copy (COPY / DRY RUN in its name) never takes the day's backup slot from the real roster
    BackupOff = (Trim$(BACKUP_FOLDER) = "" Or ThisWorkbook.ReadOnly Or Not IsAdmin() Or IsDryRunName(ThisWorkbook.Name))
End Function

' BACKUP_FOLDER with "." / ".\sub" resolved against the roster's own folder
Public Function BackupFolderPath() As String
    BackupFolderPath = ResolveFolder(BACKUP_FOLDER)
End Function

Public Sub BackupOnOpen()
    On Error Resume Next
    If BackupOff() Then Exit Sub
    Dim folder As String: folder = BackupFolderPath()
    If Not FolderExists(folder) Then
        If Not mBackupWarned Then
            mBackupWarned = True
            MsgBox "Today's backup copy wasn't saved: the backup folder can't be reached." & vbNewLine & _
                   folder & vbNewLine & vbNewLine & "The roster opens as usual. Check the folder (BACKUP_FOLDER in modOvertimeEntry).", _
                   vbExclamation, "Backup"
        End If
        Exit Sub
    End If
    Dim target As String: target = folder & BackupFileName(Date)
    If Dir$(target) <> "" Then Exit Sub                 ' already taken today
    Err.Clear
    ThisWorkbook.SaveCopyAs target
    If Err.Number <> 0 Then
        If Not mBackupWarned Then
            mBackupWarned = True
            MsgBox "Today's backup copy couldn't be saved to:" & vbNewLine & target & vbNewLine & Err.Description & _
                   vbNewLine & vbNewLine & "The roster opens as usual.", vbExclamation, "Backup"
        End If
        Exit Sub
    End If
    PruneBackups folder
End Sub

' Deletes copies older than BACKUP_KEEP_DAYS in the backup folder only,
' matched by the exact name pattern. Returns how many were removed.
Public Function PruneBackups(ByVal folder As String) As Long
    On Error Resume Next
    Dim names() As String, n As Long, f As String, d As Date, i As Long, removed As Long
    f = Dir$(folder & BACKUP_BASENAME & " ????-??-??.xlsm")
    Do While f <> ""
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$()
    Loop
    For i = 1 To n
        d = BackupFileDate(names(i))
        If d > 0 Then
            If d < Date - BACKUP_KEEP_DAYS Then
                Err.Clear
                Kill folder & names(i)
                If Err.Number = 0 Then removed = removed + 1
                Err.Clear
            End If
        End If
    Next i
    PruneBackups = removed
End Function

' SystemCheck lines for the backup
Private Sub BackupStatus(ByRef warn As String, ByRef info As String)
    On Error Resume Next
    If Trim$(BACKUP_FOLDER) = "" Then
        warn = warn & "- BACKUP_FOLDER isn't set (modOvertimeEntry), so daily backups are off" & vbNewLine
        Exit Sub
    End If
    If IsDryRunName(ThisWorkbook.Name) Then
        info = info & "Backup: off for a practice copy (COPY / DRY RUN in the file name)" & vbNewLine
        Exit Sub
    End If
    Dim folder As String: folder = BackupFolderPath()
    If Not FolderExists(folder) Then
        warn = warn & "- backup folder can't be reached: " & folder & vbNewLine
    ElseIf Dir$(folder & BackupFileName(Date)) = "" Then
        If IsAdmin() And Not ThisWorkbook.ReadOnly Then
            warn = warn & "- no backup copy taken today yet (it's saved when an admin opens the file)" & vbNewLine
        Else
            info = info & "Backup: none today (read-only / non-admin open)" & vbNewLine
        End If
    Else
        info = info & "Backup: today's copy is in " & folder & vbNewLine
    End If
End Sub
