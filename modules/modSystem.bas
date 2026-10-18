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

Public Const SYS_VERSION As String = "System v25.3 (Oct 17)"
Public Const EXPECTED_RELEASE As String = "v25.3"
Private Const N_MODULES As Long = 25           ' v24.4: + modSmoke; v25.3: + modDesk
' ThisWorkbook and each form are checked against the release that last changed them
Private Const THISWB_RELEASE As String = "v25.5"
Public Const IL_SHEET As String = "Install Log"    ' v24.4: written by the Roster Updater's Install panel (tblInstallLog); admin-only, never protected
Private Const FORM_RELEASES As String = "frmOvertimeEntry|v25.5,frmNewPosting|v25.5,frmPostingEntry|v25.5,frmShiftSwitch|v25.5," & _
                                        "frmTimeOff|v25.5,frmLeave|v25.5,frmSickCall|v25.5,frmConvert|v25.5,frmEmployee|v25.5"   ' v25.5: every form re-pasted (UI_SCALE)

Private mBackupWarned As Boolean


' ---------------- SystemCheck ----------------
' Problems come first; the message stays short enough for one MsgBox.
' v24.4: the check itself; SystemCheck shows it, SystemCheckOK shows it and says whether it was clean
' (the Install panel's "System check" step). An error inside raises to the caller's handler.
Private Sub SystemCheckBuild(ByRef warn As String, ByRef info As String, ByRef verOk As Boolean)
    On Error GoTo Failed
    Dim nOk As Long
    verOk = True

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
    VerCheck "modReports", REP_VERSION, warn, nOk
    VerCheck "modSignups", SG_VERSION, warn, nOk
    VerCheck "modErrorLog", EL_VERSION, warn, nOk
    VerCheck "modSmokeTest", SMK_VERSION, warn, nOk
    VerCheck "modOTLog", OTL_VERSION, warn, nOk               ' v24.3.1: live code (Void Entry, Add Adjustment, Rename), stamped at last
    VerCheck "modSmoke", SMO_VERSION, warn, nOk               ' v24.4: the Install panel's pop-up openers
    VerCheck "modDesk", DESK_VERSION, warn, nOk               ' v25.3: the second admin's Desk
    info = "Modules: " & nOk & " of " & N_MODULES & " on " & EXPECTED_RELEASE & vbNewLine
    If nOk <> N_MODULES Then verOk = False
    If InStr(ThisWorkbookVersion(), THISWB_RELEASE) = 0 Then
        warn = warn & "- ThisWorkbook code is older than " & THISWB_RELEASE & ": re-paste ThisWorkbook.txt" & vbNewLine
        verOk = False
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
    On Error GoTo Failed
    If Not trusted Then
        info = info & "Forms / leftovers: not checked (VBA project access is off)" & vbNewLine
        ' (v25: this line no longer fails SystemCheckOK - on the v24.4 first run the roster's own VBProject call was
        '  refused while the Install panel had just swapped the code through that model; the module count decides)
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
                    verOk = False
                End If
            End If
        Next fr
        info = info & "Forms and leftovers: checked" & vbNewLine
    End If

    ' ---- data ----
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo Failed
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
    On Error GoTo Failed
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
    On Error GoTo Failed
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
    On Error GoTo Failed
    If nImp > 0 Then warn = warn & "- " & nImp & " imported vacation line(s) in the next " & IMPORT_WATCH_DAYS & " days not booked yet: Convert imports" & vbNewLine
    ' v24: crew history and the Former sheet
    If CrewTable(False) Is Nothing Or FormerTable(False) Is Nothing Then
        warn = warn & "- Crew History / Former sheets not set up: run SetupPeople" & vbNewLine
    Else
        info = info & "Crew History: " & Format$(CrewTable(False).ListRows.Count, "#,##0") & " rows" & vbNewLine
    End If
    ' v24.3: unit moves entered ahead of time
    Dim pmRows As Variant, nPm As Long
    On Error Resume Next
    pmRows = PlannedRows()
    On Error GoTo Failed
    If IsArray(pmRows) Then nPm = UBound(pmRows, 1)
    info = info & "Planned unit moves waiting: " & nPm & vbNewLine
    If Not PostingListCurrent() Then warn = warn & "- posting lists need their Signup Hours column: run SetupPeople" & vbNewLine
    ' M1: members' file
    If Not MemberTemplateExists() Then
        warn = warn & "- members' template not found next to this workbook: " & MEMBER_TEMPLATE_NAME & " (see the M1 install notes)" & vbNewLine
    End If
    info = info & MembersStatusText() & vbNewLine
    ' v24.1: electronic sign-ups
    If SIGNUPS_ON Then
        If FolderExists(SignupFolderPath()) Then
            info = info & "Sign-ups: on (" & SignupFolderPath() & ")" & vbNewLine
        Else
            warn = warn & "- sign-up inbox folder not found: " & SignupFolderPath() & " (it is made at the next publish; members need write access to it)" & vbNewLine
        End If
    Else
        info = info & "Sign-ups: off (SIGNUPS_ON in modOvertimeEntry)" & vbNewLine
    End If
    Dim prInfo As String, prWarn As String          ' v25: every sheet protected, totals locked
    prWarn = ProtectionStatus(prInfo)
    info = info & prInfo & vbNewLine
    If prWarn <> "" Then warn = warn & prWarn & vbNewLine
    Dim bkWarn As String, bkInfo As String
    BackupStatus bkWarn, bkInfo
    warn = warn & bkWarn
    info = info & bkInfo
    Dim nUnseen As Long: nUnseen = UnseenErrorCount()
    If nUnseen > 0 Then warn = warn & "- " & nUnseen & " new line(s) in the Error Log: run ShowErrorLog" & vbNewLine
    If ThisWorkbook.ReadOnly Then warn = warn & "- This copy is read-only" & vbNewLine

    Exit Sub
Failed:
    Err.Raise Err.Number, , Err.Description        ' up to SystemCheck / SystemCheckOK (their Oops)
End Sub

Public Sub SystemCheck()
    On Error GoTo Oops                         ' v24.3.1: the first macro run after every install had no handler at all
    Dim warn As String, info As String, verOk As Boolean
    SystemCheckBuild warn, info, verOk
    If Len(warn) > 700 Then warn = Left$(warn, InStrRev(Left$(warn, 700), vbNewLine)) & "- ...and more" & vbNewLine
    Dim msg As String
    If warn = "" Then
        msg = "Everything looks right." & vbNewLine & vbNewLine & info
    Else
        msg = "CHECK:" & vbNewLine & warn & vbNewLine & info
    End If
    MsgBox msg, IIf(warn = "", vbInformation, vbExclamation), "System check  (" & EXPECTED_RELEASE & ")"
    Exit Sub
Oops:
    ReportError "SystemCheck"
End Sub

' v24.4: SystemCheck as a result, for the Install panel: True when every module, form and ThisWorkbook
' is on this release (the VERSION lines). The environment lines (Kronos age, P: paths, unseen Error
' Log lines ...) are shown as always but don't fail an install: they are expected until go-live.
' An error inside logs and returns False.
Public Function SystemCheckOK() As Boolean
    On Error GoTo Oops
    Dim warn As String, info As String, verOk As Boolean, msg As String
    SystemCheckBuild warn, info, verOk
    If Len(warn) > 700 Then warn = Left$(warn, InStrRev(Left$(warn, 700), vbNewLine)) & "- ...and more" & vbNewLine
    If warn = "" Then
        msg = "Everything looks right." & vbNewLine & vbNewLine & info
    Else
        msg = "CHECK:" & vbNewLine & warn & vbNewLine & info
    End If
    MsgBox msg, IIf(warn = "", vbInformation, vbExclamation), "System check  (" & EXPECTED_RELEASE & ")"
    SystemCheckOK = verOk
    Exit Function
Oops:
    ReportError "SystemCheckOK"
End Function

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

' ---------------- restricted copies (v25, deep review I2) ----------------
' Every copy of the roster that code saves while an admin is working (a "Backup before ..."
' copy, the practice copy, the updater's backup) is saved in the RESTRICTED view, as the daily
' backup and the rollover's frozen copy already were: admin sheets very hidden, phone / email /
' Payroll # columns hidden, every sheet protected. Opened with macros off, such a copy shows
' nothing an ordinary supervisor can't see. The admin view is put back after the save, whatever
' happened. An error from SaveCopyAs itself raises to the caller (which says so).
Public Sub SaveRestrictedCopy(ByVal fullPath As String)
    ' v25 review: the admin view is read from the sheet (the phone column shown), not only from the flag a VBA
    ' reset clears; the live file's Saved state and active sheet are put back (the Install panel's check step
    ' refuses an unsaved roster); an error in the view switch itself raises after the admin view is restored.
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim wasAdmin As Boolean: wasAdmin = gAdminUnlocked
    If Not rs Is Nothing Then wasAdmin = wasAdmin Or Not rs.Columns(C_PHONE).Hidden
    Dim wasDesk As Boolean: wasDesk = gDeskMode               ' v25.3: a Desk user gets the Desk back, not the full layout
    Dim savedWas As Boolean: savedWas = ThisWorkbook.Saved
    Dim actWas As Object: Set actWas = ActiveSheet
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim en As Long, em As String, copied As Boolean
    Application.ScreenUpdating = False
    On Error GoTo Failed
    ApplyViewCore False
    On Error Resume Next
    ThisWorkbook.SaveCopyAs fullPath
    en = Err.Number: em = Err.Description
    copied = (en = 0)
    On Error GoTo Failed
    If wasAdmin Then ApplyViewCore True, wasDesk
    On Error Resume Next
    actWas.Activate
    Application.ScreenUpdating = scrWas
    ThisWorkbook.Saved = savedWas
    On Error GoTo 0
    If en <> 0 Then Err.Raise en, , em
    Exit Sub
Failed:
    en = Err.Number: em = Err.Description
    On Error Resume Next
    If wasAdmin Then ApplyViewCore True, wasDesk
    actWas.Activate
    Application.ScreenUpdating = scrWas
    ThisWorkbook.Saved = savedWas
    On Error GoTo 0
    Err.Raise en, , IIf(copied, "the copy was saved, but the admin view didn't come back (close and reopen the roster): ", "the copy wasn't saved: ") & em
End Sub

' ---------------- the Install Log sheet (v25) ----------------
' The Install panel (Roster Updater.xlsm, v24.4) writes one row per install step to this sheet.
' Under structure protection (v25) the updater can't add a sheet, so the roster makes it on an
' admin open (ApplyView) with the same layout the updater used: tblInstallLog, very hidden,
' never protected (ProtectSheet skips IL_SHEET).
Public Sub EnsureInstallLog()
    If SheetExists(IL_SHEET) Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    On Error GoTo Oops
    UnlockStructure
    Dim actWas As Object: Set actWas = ActiveSheet
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = IL_SHEET
    ws.Range("A1:G1").Value = Array("Version", "Step", "Kind", "Status", "When", "User", "Note")
    ws.Range("A1:G1").Font.Bold = True
    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:G1"), , xlYes)
    lo.Name = "tblInstallLog"
    ws.Columns("A:G").ColumnWidth = 18: ws.Columns("B").ColumnWidth = 60: ws.Columns("G").ColumnWidth = 70
    SetSheetVisible ws, xlSheetVeryHidden
    RelockStructure
    On Error Resume Next
    actWas.Activate
    Exit Sub
Oops:
    LogError "EnsureInstallLog", Err.Number, Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then                       ' v25 review: no half-made sheet left behind (the next admin open tries again)
        UnlockStructure
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    RelockStructure
    actWas.Activate
End Sub

' v25: the Install panel's "Show the log" calls this (it can't show a sheet under structure protection)
Public Sub ShowInstallLog()
    On Error GoTo Oops
    If Not SheetExists(IL_SHEET) Then EnsureInstallLog
    If Not SheetExists(IL_SHEET) Then Exit Sub
    SetSheetVisible ThisWorkbook.Worksheets(IL_SHEET), xlSheetVisible
    ThisWorkbook.Worksheets(IL_SHEET).Activate
    Exit Sub
Oops:
    ReportError "ShowInstallLog"
End Sub

' ---------------- a practice copy (v24.2) ----------------
' Alt+F8 > MakePracticeCopy (admin): saves "<name> COPY <date time>.xlsm" in a
' Practice folder beside the roster, puts a copy of the members' template next
' to it, and opens it. Daily backups, the publish reminder and SmokeTest's
' guard all key on COPY in the name, so nothing done in it touches the roster.
Public Sub MakePracticeCopy()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    ' v24.3.2 (Graeme's Oct 13 test): a copy made from inside a copy carries everything done in
    ' it, saved or not - after SmokeTest, its test person and the rows it left. Copies come from the roster.
    If NoSave() Then
        MsgBox "Saving is blocked in this file (SmokeTest ran here, or a change did not verify), so a copy of it would carry " & _
               "that state: after SmokeTest, its test person and the rows it left." & vbNewLine & vbNewLine & _
               "Close this file without saving and run MakePracticeCopy from the roster you are testing.", vbExclamation, "Practice copy"
        Exit Sub
    End If
    If IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name) Then
        MsgBox "This file is itself a copy:" & vbNewLine & ThisWorkbook.Name & vbNewLine & vbNewLine & _
               "A practice copy made from here would carry everything done in this one, saved or not. " & _
               "Close it without saving and run MakePracticeCopy from the roster you are testing." & vbNewLine & vbNewLine & _
               "(If this IS the roster you are testing and Windows put 'Copy' in its name, rename the file first: " & _
               "COPY, DRY RUN, FINAL or Backup in the name marks a practice or backup file to every check.)", _
               vbExclamation, "Practice copy"
        Exit Sub
    End If
    If ThisWorkbook.Path = "" Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Open the roster from an ordinary folder first (not SharePoint / OneDrive).", vbExclamation, "Practice copy"
        Exit Sub
    End If
    Dim folder As String: folder = EnsureSlash(ThisWorkbook.Path) & "Practice\"
    If Not FolderExists(folder) Then
        On Error Resume Next
        MkDir folder
        On Error GoTo Oops
    End If
    If Not FolderExists(folder) Then
        MsgBox "The Practice folder couldn't be made:" & vbNewLine & folder, vbExclamation, "Practice copy"
        Exit Sub
    End If
    Dim baseName As String: baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 1 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    Dim target As String: target = folder & baseName & " COPY " & Format$(Now, "yyyy-mm-dd hhmm") & ".xlsm"
    If MsgBox("Make a practice copy and open it?" & vbNewLine & vbNewLine & target & vbNewLine & vbNewLine & _
              "It is a separate file with the roster as it is right now (saved or not). Nothing you do in it reaches the real roster " & _
              "or the members' file. Its title bar says COPY: check that before you test anything.", _
              vbOKCancel + vbQuestion, "Practice copy") <> vbOK Then Exit Sub
    SaveRestrictedCopy target                      ' v25 (I2): the copy never carries the admin view
    Dim tplNote As String
    If MemberTemplateExists() Then
        On Error Resume Next
        FileCopy MemberTemplatePath(), folder & MEMBER_TEMPLATE_NAME
        If Err.Number <> 0 Then tplNote = vbNewLine & "The members' template couldn't be copied (" & Err.Description & "): close it and copy it by hand if you want to test Member View."
        On Error GoTo Oops
    Else
        tplNote = vbNewLine & "There is no members' template beside the roster, so Member View can't be tested in the copy."
    End If
    Workbooks.Open Filename:=target
    MsgBox "You are now in the practice copy:" & vbNewLine & target & tplNote & vbNewLine & vbNewLine & _
           "Run SmokeTest here (Alt+F8), or try things by hand. Close it when you are done; delete the Practice folder whenever you like.", _
           vbInformation, "Practice copy"
    Exit Sub
Oops:
    ReportError "MakePracticeCopy"
End Sub

' ---------------- opening checks ----------------
Public Sub OpenChecks()
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then
        Dim who As String: who = LockOwner()
        Dim onDisk As Boolean: onDisk = False
        onDisk = ((GetAttr(ThisWorkbook.FullName) And vbReadOnly) = vbReadOnly)
        If onDisk And IsBackupCopyName(ThisWorkbook.Name) Then       ' v25 (I6): the year-end FINAL, or a backup set read-only
            MsgBox "This is a frozen copy of the roster (" & ThisWorkbook.Name & "): read-only on disk by design." & vbNewLine & vbNewLine & _
                   "You can look at everything; nothing can be logged or saved in it. The live roster is the file without FINAL / a date in its name.", _
                   vbInformation, "Frozen copy"
            Exit Sub
        End If
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

' v24.3.1: a file that is itself a copy of the roster - a daily backup, the year-end
' FINAL, a "Backup before ..." pre-change copy. Opened "to check the backup" it must
' not take today's backup slot or publish the members' file. Pure; mirrored and cased.
Public Function IsBackupCopyName(ByVal fileName As String) As Boolean
    If BackupFileDate(fileName) > 0 Then IsBackupCopyName = True: Exit Function
    IsBackupCopyName = (InStr(1, fileName, "FINAL", vbTextCompare) > 0 Or InStr(1, fileName, "Backup", vbTextCompare) > 0)
End Function

' v24.3.1: the live roster, open for editing by an admin - not a practice copy, not a
' backup. The one test for "may this file change anything beside itself".
Public Function IsLiveRoster() As Boolean
    If ThisWorkbook.ReadOnly Then Exit Function
    If Not IsAdmin() Then Exit Function
    If IsDryRunName(ThisWorkbook.Name) Then Exit Function
    If IsBackupCopyName(ThisWorkbook.Name) Then Exit Function
    IsLiveRoster = True
End Function

' True when the backup should be skipped without a word
Private Function BackupOff() As Boolean
    ' a practice copy (COPY / DRY RUN in its name) never takes the day's backup slot from the real roster;
    ' nor does a backup, FINAL or pre-change copy opened to look at (v24.3.1, review I1)
    BackupOff = (Trim$(BACKUP_FOLDER) = "" Or Not IsLiveRoster())
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
    If IsBackupCopyName(ThisWorkbook.Name) Then
        info = info & "Backup: off - this file is itself a backup or frozen copy (" & ThisWorkbook.Name & ")" & vbNewLine
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
