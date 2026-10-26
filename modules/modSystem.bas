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

Public Const SYS_VERSION As String = "System v28 (Oct 25)"
Public Const EXPECTED_RELEASE As String = "v28"
' v27.1: the release whose news WHATS_NEW carries - a login that has seen that box is not shown it again by a
' release with nothing new for them (v27.1 was the behaviour-free refactor; v27.5 changes three rules, so its box shows)
Public Const WHATS_NEW_RELEASE As String = "v28"
' v25.5 (A6): three lines shown once per login on the first open after an install (WhatsNewOnOpen; | = a new line)
Public Const WHATS_NEW As String = "v28: members can ASK from their file. The members' file gets a Requests page: time off between two dates, or a shift switch with the partner who would work it (the partner clicks Agree on their own PC first). The request books nothing by itself: it lands on the admin Desk and Needs Attention (MEMBERS' REQUESTS), where a double-click opens the time-off or switch pop-up prefilled and a right-click declines with a reason; the answer reaches the member at the next update.|" & _
                                   "The Payroll # is now the key: every people table carries it (filled from the roster at the install and before every save), a rename follows the number, and the Employee pop-up takes it at Add and Edit (digits only). Nothing is published: members never see the number.|" & _
                                   "Also: IntegrityCheck 16 'Keys' names a number on two rows or a row without one; the new Requests and Keys macros are listed in the admin guide. v27.5's three rule fixes (the 16 h stop on switches, the Kronos gap rows, the staffing count's partial covers) are in this install if v27.5 was skipped."
Private Const N_MODULES As Long = 43           ' v24.4: + modSmoke; v25.3: + modDesk; v27: + modReconcile, modStats, modReceipt, modRules, modEvidence; v27.1: + the 11 split out of modCalendar / modPeople / modPostings; v28: + modRequests, modKeys
' ThisWorkbook and each form are checked against the release that last changed them
Private Const THISWB_RELEASE As String = "v28"   ' v28: open step 10 (RequestsOnOpen; the later steps moved down one)
Public Const OP_SHEET As String = "Opens"         ' v27 (item 32): one row per open (tblOpens); admin-only, never published
Private Const OP_TABLE As String = "tblOpens"
Private Const OP_N As Long = 7
Private Const OPENS_KEEP_DAYS As Long = 90
Public Const IL_SHEET As String = "Install Log"    ' v24.4: written by the Roster Updater's Install panel (tblInstallLog); admin-only, never protected
Private Const FORM_RELEASES As String = "frmOvertimeEntry|v27,frmNewPosting|v27.1,frmPostingEntry|v27,frmShiftSwitch|v28," & _
                                        "frmTimeOff|v28,frmLeave|v27.1,frmSickCall|v27.1,frmConvert|v27.1,frmEmployee|v28,frmReceipt|v27.1,frmHistory|v27.1,frmTableEdit|v27.1"   ' v27.1: the ten code-built forms build through modUI's shared helpers; the two designer forms stay v27   ' v27: every form but frmConvert changed (the receipts, the draft, the history pop-up, the restriction button, the Holidays / Min Staff pop-up); three are new (frmReceipt, frmHistory, frmTableEdit); ThisWorkbook too (open steps 12-13)

Private mBackupWarned As Boolean
Private mNewestBackupAt As Date                 ' v27 (item 12): the newest backup's time, so a save does not list the folder every time
Private mPrevOpenAt As Date                     ' v27 (item 3): this login's previous open, read once a session
Private mPrevOpenRead As Boolean


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
    VerCheck "modReconcile", RC_VERSION, warn, nOk           ' v27: the Reconcile sheet
    VerCheck "modStats", STATS_VERSION, warn, nOk            ' v27: the quarterly Stats report
    VerCheck "modReceipt", RCPT_VERSION, warn, nOk           ' v27: the receipt card after every save
    VerCheck "modRules", RULES_VERSION, warn, nOk            ' v27 (item 22): the Rules and Settings sheets
    VerCheck "modEvidence", EVID_VERSION, warn, nOk          ' v27 (item 23): the period close, scans, disputes
    VerCheck "modRequests", RQ_VERSION, warn, nOk            ' v28 (item 1): members' requests
    VerCheck "modKeys", KEY_VERSION, warn, nOk               ' v28 (item 2): the Payroll # key
    VerCheck "modSignups", SG_VERSION, warn, nOk
    VerCheck "modErrorLog", EL_VERSION, warn, nOk
    VerCheck "modSmokeTest", SMK_VERSION, warn, nOk
    VerCheck "modOTLog", OTL_VERSION, warn, nOk               ' v24.3.1: live code (Void Entry, Add Adjustment, Rename), stamped at last
    VerCheck "modSmoke", SMO_VERSION, warn, nOk               ' v24.4: the Install panel's pop-up openers
    VerCheck "modDesk", DESK_VERSION, warn, nOk               ' v25.3: the second admin's Desk
    VerCheck "modBoard", BOARD_VERSION, warn, nOk   ' v27.1: split out
    VerCheck "modMinStaff", MINSTAFF_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modCalImport", CALIMP_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modNeeds", NEEDS_VERSION, warn, nOk   ' v27.1: split out
    VerCheck "modCalMenus", CALMENU_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modDates", DATES_VERSION, warn, nOk   ' v27.1: split out
    VerCheck "modFormer", FORMER_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modPlannedMoves", PLANMOVE_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modPostingSheet", PSHEET_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modPostingSync", PSYNC_VERSION, warn, nOk ' v27.1: split out
    VerCheck "modPostingMail", PMAIL_VERSION, warn, nOk ' v27.1: split out
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
    ' v27: what Worked (H) is read from - the cumulative report's snapshot and the daily snapshots' ledger
    If KronosCumThrough() > 0 Or KronosLedgerFrom() > 0 Then
        info = info & "Kronos reports: Totals through " & IIf(KronosCumThrough() > 0, Format$(KronosCumThrough(), "mmm d, yyyy"), "none yet") & _
               "; Worked Days ledger from " & IIf(KronosLedgerFrom() > 0, Format$(KronosLedgerFrom(), "mmm d, yyyy"), "none yet") & vbNewLine
    Else
        info = info & "Kronos reports: none imported yet (Worked comes from the pivot sheet)" & vbNewLine
    End If
    info = info & DiskStampText() & vbNewLine      ' v26.5.5 (A2)
    info = info & SettingsOverrideText() & vbNewLine   ' v27 (item 22): the Settings sheet's overrides in force
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
    ' v26: minimum staffing (the Min Staff sheet; made with the plant's numbers on the first calendar draw by an admin)
    Dim nMin As Long
    On Error Resume Next
    nMin = MinStaffMap().Count
    On Error GoTo Failed
    info = info & "Minimum staffing: " & IIf(SheetExists(MINSTAFF_SHEET), nMin & " unit / shift minimum(s) on the " & MINSTAFF_SHEET & " sheet", _
                                           "the " & MINSTAFF_SHEET & " sheet is made on an admin's next open (once the Calendar sheet exists)") & vbNewLine
    If Not PostingListCurrent() Then warn = warn & "- posting lists need their Signup Hours column: run SetupPeople" & vbNewLine
    ' M1: members' file
    If Not MemberTemplateExists() Then
        warn = warn & "- members' template not found next to this workbook: " & MEMBER_TEMPLATE_NAME & " (see the M1 install notes)" & vbNewLine
    End If
    info = info & MembersStatusText() & vbNewLine
    info = info & OpensSummaryText() & vbNewLine     ' v27 (item 32): errors this release, opens this week
    ' v26.5.5 (A13): where the members' file goes (PUBLISH_FOLDER resolved; "." = beside the master until the folder is named)
    If FolderExists(PublishFolderPath()) Then
        info = info & "Members' folder: " & PublishFolderPath() & IIf(Trim$(PUBLISH_FOLDER) = ".", " (beside the master: PUBLISH_FOLDER not named yet)", "") & vbNewLine
    Else
        warn = warn & "- members' folder can't be reached: " & PublishFolderPath() & " (PUBLISH_FOLDER in modOvertimeEntry; Member View will refuse)" & vbNewLine
    End If
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
' v27 (item 31, C7): the SystemCheck text with no box - the Install panel keeps it in the Install Log row's Note
Public Function SystemCheckText() As String
    On Error GoTo Oops
    Dim warn As String, info As String, verOk As Boolean
    SystemCheckBuild warn, info, verOk
    SystemCheckText = IIf(warn = "", "Everything looks right.", "CHECK: " & warn) & vbNewLine & info
    Exit Function
Oops:
    SystemCheckText = "SystemCheck could not run: " & Err.Description
End Function

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
    On Error Resume Next                            ' never raises
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
    If Not RequireSupervisor() Then Exit Sub        ' v26.5.5 (5.4): supervisors practise too (the Desk's Practice button)
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
           "Every sheet carries a red PRACTICE COPY banner: nothing done here counts, reaches the roster or the members' file. " & _
           "Try things by hand, or run SmokeTest (Alt+F8). Close it when you are done: copies older than a day are deleted " & _
           "with the roster's daily backup (the first open of the day).", vbInformation, "Practice copy"
    Exit Sub
Oops:
    ReportError "MakePracticeCopy"
End Sub

' v26.5.5 (5.4): a practice copy says so on every visible sheet (a red banner shape, top left) and in the
' window title, so nobody mistakes it for the roster. Workbook_Open runs it; the live roster is untouched.
Public Sub PracticeBannerOnOpen()
    On Error Resume Next                            ' never raises
    If Not IsDryRunName(ThisWorkbook.Name) Then Exit Sub
    Dim ws As Worksheet, s As Shape, i As Long
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            UnlockSheet ws
            For i = ws.Shapes.Count To 1 Step -1
                If ws.Shapes(i).Name = "prac_banner" Then ws.Shapes(i).Delete
            Next i
            Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 330, 16)
            s.Name = "prac_banner"
            s.Fill.ForeColor.RGB = RGB(190, 30, 30)
            s.Line.Visible = msoFalse
            s.Placement = xlFreeFloating
            s.Locked = True
            With s.TextFrame
                .Characters.Text = "PRACTICE COPY: nothing here counts"
                .Characters.Font.Bold = True
                .Characters.Font.Size = 9
                .Characters.Font.Color = vbWhite
                .MarginTop = 1: .MarginBottom = 1: .MarginLeft = 4
            End With
            ProtectSheet ws
        End If
    Next ws
    ThisWorkbook.Windows(1).Caption = "PRACTICE COPY  -  " & ThisWorkbook.Name      ' (review) this copy's own window, whichever is in front
End Sub

' v26.5.5 (5.4). PURE (SelfTest "practice"): the date and time in a practice copy's name,
' "<roster name> COPY yyyy-mm-dd hhmm.xlsm" (MakePracticeCopy writes it), or 0 when the name isn't that
Public Function PracticeCopyDate(ByVal fileName As String) As Date
    Dim p As Long: p = InStr(1, fileName, " COPY ", vbTextCompare)
    If p = 0 Then Exit Function
    Dim s As String: s = Mid$(fileName, p + 6)
    If Not s Like "####-##-## ####.xlsm" Then Exit Function
    Dim y As Long, mo As Long, dy As Long, hh As Long, mi As Long
    y = CLng(Left$(s, 4)): mo = CLng(Mid$(s, 6, 2)): dy = CLng(Mid$(s, 9, 2)): hh = CLng(Mid$(s, 12, 2)): mi = CLng(Mid$(s, 14, 2))
    If mo < 1 Or mo > 12 Or dy < 1 Or dy > 31 Or hh > 23 Or mi > 59 Then Exit Function      ' DateSerial would roll these over
    If Day(DateSerial(y, mo, dy)) <> dy Then Exit Function                                  ' Feb 30 and the like
    PracticeCopyDate = DateSerial(y, mo, dy) + TimeSerial(hh, mi, 0)
End Function

' Deletes the practice copies (and nothing else) in the Practice folder beside the roster that are more
' than a day old - the copy can't delete itself on close. Returns how many went.
Public Function PrunePracticeCopies(ByVal folder As String) As Long
    On Error Resume Next                            ' never raises
    If Not FolderExists(folder) Then Exit Function
    Dim names() As String, n As Long, f As String, i As Long, removed As Long, d As Date
    f = Dir$(folder & "* COPY ????-??-?? ????.xlsm")
    Do While f <> "" And Err.Number = 0
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$()
    Loop
    Err.Clear
    For i = 1 To n
        d = PracticeCopyDate(names(i))
        If d > 0 Then
            If Now - d > 1 Then
                Err.Clear
                Kill folder & names(i)
                If Err.Number = 0 Then removed = removed + 1
                Err.Clear
            End If
        End If
    Next i
    PrunePracticeCopies = removed
End Function

' ---------------- opening checks ----------------
Public Sub OpenChecks()
    On Error Resume Next                            ' never raises
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

' v27 (item 20): a hidden-name key for this login - <prefix>_<login>, the login's odd characters made safe
' v27 (item 3): this login's previous open (0 = none on record), from the hidden name LastOpen_<login>; the first
' call of a session reads it and moves the name on to now (kept with the next real save - a workbook that was
' saved stays marked saved). Never raises.
Public Function PrevOpenAt() As Date
    If mPrevOpenRead Then PrevOpenAt = mPrevOpenAt: Exit Function
    mPrevOpenRead = True
    On Error Resume Next
    Dim key As String: key = LoginNameKey("LastOpen")
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    Dim v As Double
    If Not nmObj Is Nothing Then v = Val(Replace(nmObj.RefersTo, "=", ""))
    If v > 40000 And v < CDbl(Now) Then mPrevOpenAt = CDate(v)       ' a real date before now, else nothing on record
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    Dim nowTxt As String: nowTxt = "=" & Trim$(Str$(CDbl(Now)))      ' Str$ writes a point whatever the locale
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:=nowTxt, Visible:=False
    Else
        nmObj.RefersTo = nowTxt
    End If
    If wasSaved Then ThisWorkbook.Saved = True
    PrevOpenAt = mPrevOpenAt
End Function

Public Function LoginNameKey(ByVal prefix As String) As String
    Dim u As String, i As Long, ch As String, s As String
    u = LCase$(CurrentUser())
    For i = 1 To Len(u)
        ch = Mid$(u, i, 1)
        If ch Like "[a-z0-9]" Then s = s & ch Else s = s & "_"
    Next i
    If s = "" Then s = "nobody"
    LoginNameKey = prefix & "_" & s
End Function

' v27 (item 20, B15): the first time a login ever opens the roster, one box points at the three things that matter
' (a hidden name FirstOpen_<login> remembers it). Run by Workbook_Open BEFORE the what's-new box, which it reads. Quiet.
Public Sub FirstOpenOnOpen()
    On Error Resume Next                            ' never raises
    If CurrentUser() = "" Or gSmokeRun Then Exit Sub
    Dim nmKey As String: nmKey = LoginNameKey("FirstOpen")
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(nmKey)
    If Not nmObj Is Nothing Then Exit Sub
    ' (review MED-9) a login that opened an earlier release has a what's-new name: not a first open - remembered quietly
    Dim seen As Name: Set seen = ThisWorkbook.Names(WhatsNewName())
    If Not seen Is Nothing Then
        ThisWorkbook.Names.Add Name:=nmKey, RefersTo:="=""" & EXPECTED_RELEASE & """", Visible:=False
        Exit Sub
    End If
    Application.ScreenUpdating = True
    MsgBox "First time here, " & CurrentUser() & "?" & vbNewLine & vbNewLine & _
           "- To fill a shift: Log Call-Out on the bar (or Fill a shift now on the Desk; Ctrl+Shift+C). Load the list, work down it, SUBMIT." & vbNewLine & _
           "- The Calendar shows who is off and what needs cover." & vbNewLine & _
           "- Every job is explained under How do I ... on the bar (the Help tab)." & vbNewLine & vbNewLine & _
           SignedInText() & IIf(IsSupervisor(), ".", ".  To book sick calls, post and award shifts too, ask " & PRIMARY_ADMIN_NAME & " to be added as a supervisor."), vbInformation, "Call Out Roster"
    ThisWorkbook.Names.Add Name:=nmKey, RefersTo:="=""" & EXPECTED_RELEASE & """", Visible:=False
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
    On Error Resume Next                            ' never raises
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

' ---------------- what's new (v25.5, A6) ----------------
' Run by Workbook_Open after OpenChecks (re-review: on a painted screen): the WHATS_NEW box once per login per release. A hidden name per login
' ("WhatsNewSeen_<login>") holds the last release seen; it is written here and saved with the next save, so a
' read-only open sees the box again. Never raises (the open step logs an error, the box is not worth one).
Private Function WhatsNewName() As String
    WhatsNewName = LoginNameKey("WhatsNewSeen")   ' v27: the one per-login key builder
End Function

Public Sub WhatsNewOnOpen()
    On Error Resume Next                            ' never raises
    If CurrentUser() = "" Then Exit Sub
    Dim nmObj As Name
    Set nmObj = ThisWorkbook.Names(WhatsNewName())
    If Not nmObj Is Nothing Then
        If InStr(nmObj.RefersTo, WHATS_NEW_RELEASE) > 0 Then Exit Sub     ' v27.1: the news's release, not the install's
    End If
    ShowWhatsNew
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=WhatsNewName(), RefersTo:="=""" & WHATS_NEW_RELEASE & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & WHATS_NEW_RELEASE & """"
    End If
End Sub

' Alt+F8 (the Help tab names it): the box itself
Public Sub ShowWhatsNew()
    On Error GoTo Oops
    Application.ScreenUpdating = True                ' Workbook_Open has it back on by then; belt and braces for an Alt+F8 run with it off
    MsgBox "What's new in " & WHATS_NEW_RELEASE & ":" & vbNewLine & vbNewLine & _
           "- " & Replace(WHATS_NEW, "|", vbNewLine & vbNewLine & "- ") & vbNewLine & vbNewLine & _
           "The Help tab (How do I ... on the Desk) has every job. This box shows once; Alt+F8 > ShowWhatsNew brings it back.", _
           vbInformation, "Call Out Roster " & WHATS_NEW_RELEASE
    Exit Sub
Oops:
    ReportError "ShowWhatsNew"
End Sub

' True when the backup should be skipped without a word
Private Function BackupOff() As Boolean
    ' a practice copy (COPY / DRY RUN in its name) never takes the day's backup slot from the real roster;
    ' nor does a backup, FINAL or pre-change copy opened to look at (v24.3.1, review I1).
    ' v27 (item 12, Graeme Oct 19-20): ANY writable open of the live file backs up, not only an admin's
    BackupOff = (Trim$(BACKUP_FOLDER) = "" Or ThisWorkbook.ReadOnly Or IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name))
End Function

' v27 (item 2): today's daily backup - 1 taken, 0 not in the folder, -1 this copy takes none (a practice, backup or
' read-only copy). Never raises.
Public Function BackupTodayState() As Long
    On Error Resume Next
    If BackupOff() Then BackupTodayState = -1: Exit Function
    Dim folder As String: folder = EnsureSlash(BackupFolderPath())
    If FolderExists(folder) Then
        If Dir$(folder & BackupFileName(Date)) <> "" Then BackupTodayState = 1
    End If
End Function

' BACKUP_FOLDER with "." / ".\sub" resolved against the roster's own folder
Public Function BackupFolderPath() As String
    BackupFolderPath = ResolveFolder(BACKUP_FOLDER)
End Function

Public Sub BackupOnOpen()
    On Error Resume Next                            ' never raises
    If BackupOff() Then Exit Sub
    Dim folder As String: folder = BackupFolderPath()
    If Not FolderExists(folder) Then
        If Not mBackupWarned Then
            mBackupWarned = True
            MsgBox "Today's backup copy wasn't saved: the backup folder can't be reached." & vbNewLine & _
                   folder & vbNewLine & vbNewLine & "The roster opens as usual. Tell " & PRIMARY_ADMIN_NAME & " (the backup folder setting).", _
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
    PrunePreChangeCopies EnsureSlash(ThisWorkbook.Path)      ' v26: the "Backup before ..." copies live beside the roster
    PrunePracticeCopies EnsureSlash(ThisWorkbook.Path) & "Practice\"      ' v26.5.5 (5.4): practice copies older than a day
End Sub

' v26 (backup housekeeping). PURE (SelfTest "prechange"): the date in a pre-change copy's
' name, "Backup before <what> yyyy-mm-dd hhmm[ss] - <roster file name>" (modPeople, modCalendar,
' modUI, modMigration write them), or 0 when the name isn't exactly that for THIS roster's name.
' A practice copy's own name never matches the real roster's copies, and a copy whose stamp
' can't be read is never touched.
Public Function PreChangeCopyDate(ByVal fileName As String, ByVal rosterName As String) As Date
    Const HEAD As String = "Backup before "
    Dim tail As String: tail = " - " & rosterName
    If rosterName = "" Then Exit Function
    If Len(fileName) <= Len(HEAD) + Len(tail) Then Exit Function
    If StrComp(Left$(fileName, Len(HEAD)), HEAD, vbTextCompare) <> 0 Then Exit Function
    If StrComp(Right$(fileName, Len(tail)), tail, vbTextCompare) <> 0 Then Exit Function
    Dim core As String: core = Left$(fileName, Len(fileName) - Len(tail))    ' "Backup before planned moves 2026-10-13 091500"
    Dim p As Long: p = InStrRev(core, " ")
    If p = 0 Then Exit Function
    Dim tm As String: tm = Mid$(core, p + 1)
    If Not (tm Like "####" Or tm Like "######") Then Exit Function
    core = Left$(core, p - 1)
    p = InStrRev(core, " ")
    If p <= Len(HEAD) Then Exit Function                  ' no "what" between the head and the date
    Dim ds As String: ds = Mid$(core, p + 1)
    If Not ds Like "####-##-##" Then Exit Function
    PreChangeCopyDate = DateSerial(CLng(Left$(ds, 4)), CLng(Mid$(ds, 6, 2)), CLng(Mid$(ds, 9, 2)))
End Function

' v26: deletes this roster's "Backup before ..." copies older than BACKUP_KEEP_DAYS from the
' folder they are written to (beside the roster), matched by PreChangeCopyDate. Returns how many.
Public Function PrunePreChangeCopies(ByVal folder As String) As Long
    On Error Resume Next                            ' never raises
    Dim names() As String, n As Long, f As String, d As Date, i As Long, removed As Long
    f = Dir$(folder & "Backup before * - " & ThisWorkbook.Name)
    Do While f <> ""
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$()
    Loop
    For i = 1 To n
        d = PreChangeCopyDate(names(i), ThisWorkbook.Name)
        If d > 0 Then
            If d < Date - BACKUP_KEEP_DAYS Then
                Err.Clear
                Kill folder & names(i)
                If Err.Number = 0 Then removed = removed + 1
                Err.Clear
            End If
        End If
    Next i
    PrunePreChangeCopies = removed
End Function

' v26: how many daily backups (in the backup folder) and pre-change copies (beside the roster)
' there are, and their size in bytes
Private Sub BackupInventory(ByVal folder As String, ByRef nDaily As Long, ByRef bDaily As Double, _
                            ByRef nPre As Long, ByRef bPre As Double)
    On Error Resume Next                            ' never raises
    Dim f As String, here As String
    f = Dir$(folder & BACKUP_BASENAME & " ????-??-??.xlsm")
    Do While f <> ""
        If BackupFileDate(f) > 0 Then nDaily = nDaily + 1: bDaily = bDaily + FileLen(folder & f)
        f = Dir$()
    Loop
    here = EnsureSlash(ThisWorkbook.Path)
    f = Dir$(here & "Backup before * - " & ThisWorkbook.Name)
    Do While f <> ""
        If PreChangeCopyDate(f, ThisWorkbook.Name) > 0 Then nPre = nPre + 1: bPre = bPre + FileLen(here & f)
        f = Dir$()
    Loop
End Sub

Private Function MbText(ByVal bytes As Double) As String
    MbText = Format$(bytes / 1048576, "0") & " MB"
End Function

' Deletes copies older than BACKUP_KEEP_DAYS in the backup folder only,
' matched by the exact name pattern. Returns how many were removed.
Public Function PruneBackups(ByVal folder As String) As Long
    On Error Resume Next                            ' never raises
    Dim names() As String, n As Long, f As String, d As Date, i As Long, removed As Long
    f = Dir$(folder & BACKUP_BASENAME & " ????-??-??.xlsm")
    Do While f <> ""
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$()
    Loop
    f = Dir$(folder & BACKUP_BASENAME & " backup ????-??-?? ??.xlsm")      ' v27: the hourly backups
    Do While f <> ""
        n = n + 1
        ReDim Preserve names(1 To n)
        names(n) = f
        f = Dir$()
    Loop
    ' v27 (item 12; review): the FIRST daily backup of each pay period is the long-kept one - found per period before the
    ' loop. The period anchor is read once; if the roster's header cannot be read, nothing is pruned (the backups are
    ' exactly what a broken file needs).
    Dim firstOf As Object: Set firstOf = CreateObject("Scripting.Dictionary")
    Dim ps As Date, key As Long, anchorOk As Boolean
    Err.Clear
    ps = PeriodStartFor(Date)
    anchorOk = (Err.Number = 0 And ps > 0)
    Err.Clear
    If Not anchorOk Then Exit Function
    For i = 1 To n
        d = BackupFileDate(names(i))
        If d > 0 Then
            key = CLng(Int(CDbl(PeriodStartFor(d))))
            If Err.Number <> 0 Then Err.Clear: Exit Function
            If Not firstOf.Exists(key) Then firstOf.Add key, d Else If d < firstOf(key) Then firstOf(key) = d
        End If
    Next i
    Dim keep As Boolean
    For i = 1 To n
        keep = True                                    ' (review HIGH): a read that fails leaves the file alone
        d = BackupFileDate(names(i))
        If d > 0 Then
            key = CLng(Int(CDbl(PeriodStartFor(d))))
            keep = BackupKeepCore("daily", d, Date, (firstOf.Exists(key) And firstOf(key) = d))
        Else
            d = HourlyBackupDate(names(i))
            If d > 0 Then keep = BackupKeepCore("hourly", d, Date, False)
        End If
        If Err.Number <> 0 Then keep = True: Err.Clear
        If Not keep Then
            Err.Clear
            Kill folder & names(i)
            If Err.Number = 0 Then removed = removed + 1
            Err.Clear
        End If
    Next i
    PruneBackups = removed
End Function

' v27 (item 12). PURE (SelfTest "bkeep"; mirrored): is a backup kept today? A daily backup for BACKUP_KEEP_DAYS, or
' BACKUP_PERIOD_KEEP_DAYS when it is the FIRST daily backup of its pay period (the caller works that out); an hourly
' backup for BACKUP_HOURLY_KEEP_DAYS. Dates only (the hour is dropped).
Public Function BackupKeepCore(ByVal kind As String, ByVal fileDate As Date, ByVal today As Date, ByVal periodStart As Boolean) As Boolean
    Dim age As Long: age = CLng(Int(CDbl(today)) - Int(CDbl(fileDate)))
    If LCase$(kind) = "hourly" Then
        BackupKeepCore = (age <= BACKUP_HOURLY_KEEP_DAYS)
    ElseIf periodStart Then
        BackupKeepCore = (age <= BACKUP_PERIOD_KEEP_DAYS)
    Else
        BackupKeepCore = (age <= BACKUP_KEEP_DAYS)
    End If
End Function

' v27: the hourly backup's name, "<basename> backup yyyy-mm-dd hh.xlsm" (one per hour at most)
Public Function HourlyBackupName(ByVal t As Date) As String
    HourlyBackupName = BACKUP_BASENAME & " backup " & Format$(t, "yyyy-mm-dd hh") & ".xlsm"
End Function

' PURE (SelfTest "bhourly"; mirrored): the day and hour in an hourly backup's name, or 0
Public Function HourlyBackupDate(ByVal fileName As String) As Date
    Dim pat As String: pat = BACKUP_BASENAME & " backup ####-##-## ##.xlsm"
    If Len(fileName) <> Len(pat) Then Exit Function
    If Not fileName Like pat Then Exit Function
    Dim s As String: s = Mid$(fileName, Len(BACKUP_BASENAME) + 9, 13)
    Dim y As Long, mo As Long, dy As Long, hh As Long
    y = CLng(Left$(s, 4)): mo = CLng(Mid$(s, 6, 2)): dy = CLng(Mid$(s, 9, 2)): hh = CLng(Mid$(s, 12, 2))
    If mo < 1 Or mo > 12 Or dy < 1 Or dy > 31 Or hh > 23 Then Exit Function
    If Day(DateSerial(y, mo, dy)) <> dy Then Exit Function
    HourlyBackupDate = DateSerial(y, mo, dy) + TimeSerial(hh, 0, 0)
End Function

' v27 (item 12): called by Workbook_BeforeSave - when the newest backup (daily or hourly) is older than BACKUP_SAVE_HOURS,
' a copy of the file AS IT IS NOW goes beside the daily backups before the save overwrites it. Quiet: a backup never stops a save.
Public Sub BackupBeforeSave()
    On Error Resume Next                            ' never raises
    If BackupOff() Then Exit Sub
    If mNewestBackupAt > 0 Then
        If (Now - mNewestBackupAt) * 24 < BACKUP_SAVE_HOURS Then Exit Sub
    End If
    Dim folder As String: folder = BackupFolderPath()
    If Not FolderExists(folder) Then Exit Sub
    Dim newest As Date, f As String, at As Date
    f = Dir$(folder & BACKUP_BASENAME & " ????-??-??.xlsm")
    Do While f <> ""
        at = 0: at = FileDateTime(folder & f)
        If at > newest Then newest = at
        f = "": f = Dir$()
    Loop
    f = Dir$(folder & BACKUP_BASENAME & " backup ????-??-?? ??.xlsm")
    Do While f <> ""
        at = 0: at = FileDateTime(folder & f)
        If at > newest Then newest = at
        f = "": f = Dir$()
    Loop
    If newest > 0 Then
        mNewestBackupAt = newest
        If (Now - newest) * 24 < BACKUP_SAVE_HOURS Then Exit Sub
    End If
    Dim target As String: target = folder & HourlyBackupName(Now)
    If Dir$(target) <> "" Then mNewestBackupAt = Now: Exit Sub
    Err.Clear
    ThisWorkbook.SaveCopyAs target
    If Err.Number <> 0 Then LogError "BackupBeforeSave", Err.Number, Err.Description Else mNewestBackupAt = Now
    Err.Clear
    PruneBackups folder
End Sub

' v27 (item 12): the restore drill (admin; every release's group 4 from v27; the December dry run). The newest daily
' backup is copied to the Practice folder as a COPY, opened, and asked: SystemCheckOK, IntegrityOnOpen (ran at its open)
' -> IntegrityHasProblems, SelfTestOK; its OT Log row count and everyone's Total Hours are compared with the live file's,
' so the box says what a restore from that backup would lose. One Install Log row (kind drill). The copy is closed.
Public Sub RestoreDrill()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name) Then
        MsgBox "Run the restore drill from the live roster, not from a copy.", vbExclamation, "Restore drill": Exit Sub
    End If
    Dim folder As String: folder = BackupFolderPath()
    If Not FolderExists(folder) Then MsgBox "The backup folder can't be reached:" & vbNewLine & folder, vbExclamation, "Restore drill": Exit Sub
    Dim f As String, best As String, bestD As Date, d As Date
    f = Dir$(folder & BACKUP_BASENAME & " ????-??-??.xlsm")
    Do While f <> ""
        d = BackupFileDate(f)
        If d > bestD Then bestD = d: best = f
        f = Dir$()
    Loop
    If best = "" Then MsgBox "No daily backup was found in " & folder, vbExclamation, "Restore drill": Exit Sub
    Dim prac As String: prac = EnsureSlash(ThisWorkbook.Path) & "Practice\"
    If Not FolderExists(prac) Then MkDir prac
    Dim baseName As String: baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 1 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    Dim target As String: target = prac & baseName & " DRILL COPY " & Format$(Now, "yyyy-mm-dd hhmm") & ".xlsm"
    If MsgBox("Restore drill: copy the newest daily backup (" & best & ") to" & vbNewLine & target & vbNewLine & vbNewLine & _
              "open it as a practice copy, run SystemCheck, IntegrityCheck and SelfTest in it, and compare it with this file? " & _
              "Nothing here changes; the copy is closed at the end (and deleted with tomorrow's backup).", vbOKCancel + vbQuestion, "Restore drill") <> vbOK Then Exit Sub
    FileCopy folder & best, target
    ' this file's figures first
    Dim liveRows As Long: liveRows = 0
    If Not LogTable() Is Nothing Then liveRows = LogTable().ListRows.Count
    Dim liveTot As Object: Set liveTot = RosterTotalsMap()
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=target)                ' its own Workbook_Open runs: the banner, IntegrityOnOpen (admin)
    Dim nm As String: nm = "'" & wb.Name & "'!"
    Dim sysOk As Boolean, icBad As Boolean, stOk As Boolean, em As String
    On Error Resume Next
    sysOk = Application.Run(nm & "SystemCheckOK")
    If Err.Number <> 0 Then em = em & "SystemCheck could not run (" & Err.Description & "). ": Err.Clear
    icBad = (Application.Run(nm & "IntegrityRunCount") > 0)            ' (review) run for real in the copy, not the open's quiet result
    If Err.Number <> 0 Then em = em & "IntegrityCheck could not run (" & Err.Description & "). ": icBad = True: Err.Clear
    stOk = Application.Run(nm & "SelfTestOK")
    If Err.Number <> 0 Then em = em & "SelfTest could not run (" & Err.Description & "). ": Err.Clear
    ' the copy's log rows and totals
    Dim copyRows As Long, lost As String, nLost As Long
    copyRows = wb.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE).ListRows.Count
    If Err.Number <> 0 Then em = em & "the copy's OT Log could not be read (" & Err.Description & "). ": Err.Clear
    Dim ws As Worksheet: Set ws = wb.Worksheets(ROSTER_SHEET)
    If Err.Number <> 0 Then em = em & "the copy's roster could not be read (" & Err.Description & "). ": Err.Clear
    Dim r As Long, k As String, tLive As String, tCopy As String
    Dim copyTot As Object: Set copyTot = CreateObject("Scripting.Dictionary"): copyTot.CompareMode = 1
    For r = R_FIRST To ws.Cells(ws.Rows.Count, C_NAME).End(xlUp).Row
        k = Trim$(SafeStr(ws.Cells(r, C_NAME).Value))
        If k <> "" Then copyTot(k) = Format$(NumOr0(ws.Cells(r, C_TOTAL).Value), "0.0000")
    Next r
    Dim kk As Variant
    For Each kk In liveTot.Keys
        tLive = Split(CStr(liveTot(kk)), "|")(0)
        If copyTot.Exists(kk) Then tCopy = CStr(copyTot(kk)) Else tCopy = ""
        If tCopy = "" Then
            nLost = nLost + 1: If nLost <= 10 Then lost = lost & "  " & kk & ": not in the backup" & vbNewLine
        ElseIf Abs(Val(tLive) - Val(tCopy)) > 0.0001 Then
            nLost = nLost + 1: If nLost <= 10 Then lost = lost & "  " & kk & ": " & HoursText(Val(tCopy)) & " in the backup, " & HoursText(Val(tLive)) & " now" & vbNewLine
        End If
    Next kk
    If Err.Number <> 0 Then em = em & "the comparison stopped part-way (" & Err.Description & "). ": Err.Clear
    wb.Close SaveChanges:=False
    On Error GoTo Oops
    Dim verdict As String
    verdict = "Backup " & best & " (" & Format$(bestD, "ddd mmm d") & ")" & vbNewLine & _
              "  SystemCheck: " & IIf(sysOk, "OK", "NOT clean") & "   IntegrityCheck: " & IIf(icBad, "PROBLEMS", "no PROBLEM") & "   SelfTest: " & IIf(stOk, "passed", "FAILED") & vbNewLine & _
              "  A restore from it would lose " & (liveRows - copyRows) & " OT Log row(s) logged since, and these people's totals would move:" & vbNewLine & _
              IIf(lost = "", "  (nobody's)" & vbNewLine, lost & IIf(nLost > 10, "  ... and " & (nLost - 10) & " more" & vbNewLine, "")) & _
              IIf(em <> "", vbNewLine & em, "")
    Dim rowOk As Boolean: rowOk = LogInstallRow("drill", "Restore drill: " & best, IIf(sysOk And Not icBad And stOk And em = "", "PASS", "CHECK"), Replace(verdict, vbNewLine, " | "))
    MsgBox "Restore drill done." & vbNewLine & vbNewLine & verdict & vbNewLine & IIf(rowOk, "One Install Log row was written (saved with the file).", "The Install Log row could NOT be written."), _
           IIf(sysOk And Not icBad And stOk And em = "", vbInformation, vbExclamation), "Restore drill"
    Exit Sub
Oops:
    ReportError "RestoreDrill"
End Sub

' v27: one row into the Install Log from inside the roster (the updater writes the install's rows; the drill its own).
' The sheet is never protected by design.
Public Function LogInstallRow(ByVal kind As String, ByVal stepTxt As String, ByVal status As String, ByVal note As String) As Boolean
    On Error Resume Next                            ' never raises
    EnsureInstallLog
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(IL_SHEET)
    Dim lo As ListObject: Set lo = ws.ListObjects("tblInstallLog")
    If lo Is Nothing Then Exit Function
    Dim r As Long: r = lo.HeaderRowRange.Row + lo.ListRows.Count + 1
    If lo.ListRows.Count = 1 Then
        If Not lo.DataBodyRange Is Nothing Then
            If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then r = lo.HeaderRowRange.Row + 1
        End If
    End If
    Err.Clear
    ws.Cells(r, 1).Value = EXPECTED_RELEASE
    ws.Cells(r, 2).Value = stepTxt
    ws.Cells(r, 3).Value = kind
    ws.Cells(r, 4).Value = status
    ws.Cells(r, 5).Value = Now
    ws.Cells(r, 6).Value = CurrentUser()
    ws.Cells(r, 7).Value = "'" & Left$(note, 250)
    lo.Resize ws.Range(lo.HeaderRowRange, ws.Cells(r, 7))
    LogInstallRow = (Err.Number = 0 And Trim$(SafeStr(ws.Cells(r, 2).Value)) = stepTxt)
End Function

' SystemCheck lines for the backup
Private Sub BackupStatus(ByRef warn As String, ByRef info As String)
    On Error Resume Next                            ' never raises
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
    ' v26: what the folders hold (the prune runs after the day's backup is saved)
    Dim nDaily As Long, bDaily As Double, nPre As Long, bPre As Double
    BackupInventory folder, nDaily, bDaily, nPre, bPre
    info = info & "Backups kept: " & nDaily & " daily (" & MbText(bDaily) & ") in the backup folder, " & nPre & _
           " pre-change (" & MbText(bPre) & ") beside the roster; copies older than " & BACKUP_KEEP_DAYS & _
           " days are removed on the first admin open of the day" & vbNewLine
End Sub

' ====================================================================
'  THE OPENS TABLE (v27 item 32, C12)
' ====================================================================
' One row per open of the live roster (When, Login, View, Release, Read-only, Integrity PROBLEMs, Calendar draw s) on
' the admin-only Opens sheet (made on first use), pruned at OPENS_KEEP_DAYS; never published. Workbook_Open step 13,
' after the integrity check. A practice / backup copy writes nothing; a read-only open could never keep its row, so
' it writes nothing either (the column says "no" on every kept row). The row goes with the next real save. Never raises.
Public Sub OpensOnOpen()
    Dim evWas As Boolean: evWas = Application.EnableEvents
    On Error GoTo Quit
    If ThisWorkbook.ReadOnly Then Exit Sub
    If IsDryRunName(ThisWorkbook.Name) Or IsBackupCopyName(ThisWorkbook.Name) Then Exit Sub
    Application.EnableEvents = False                ' (review) before the sheet is made: no SheetActivate during the open
    Dim lo As ListObject: Set lo = OpensTable(True)
    If lo Is Nothing Then Application.EnableEvents = evWas: Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    Dim i As Long
    If Not lo.DataBodyRange Is Nothing Then          ' the rows older than the keep go first
        For i = lo.ListRows.Count To 1 Step -1
            If IsDate(lo.DataBodyRange.Cells(i, 1).Value) Then
                If CDate(lo.DataBodyRange.Cells(i, 1).Value) < Now - OPENS_KEEP_DAYS Then lo.ListRows(i).Delete
            End If
        Next i
    End If
    Dim rw As ListRow
    If lo.DataBodyRange Is Nothing Then
        Set rw = lo.ListRows.Add
    ElseIf lo.ListRows.Count = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
        Set rw = lo.ListRows(1)                       ' the blank row a new table is made with
    Else
        Set rw = lo.ListRows.Add
    End If
    Dim v(1 To 1, 1 To OP_N) As Variant
    v(1, 1) = Now
    v(1, 2) = CurrentUser()
    v(1, 3) = ViewModeFor(CurrentUser(), ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS)
    v(1, 4) = EXPECTED_RELEASE
    v(1, 5) = "no"
    Dim nProb As Long: nProb = IntegrityProblemCount()
    If nProb < 0 Then v(1, 6) = "not run" Else v(1, 6) = nProb
    v(1, 7) = CalDrawSeconds()
    rw.Range.Value = v
    ProtectSheet ws
    Application.EnableEvents = evWas
    Exit Sub
Quit:
    Dim en As Long: en = Err.Number
    Dim em As String: em = Err.Description
    Application.EnableEvents = evWas
    If SheetExists(OP_SHEET) Then ProtectSheetQuiet ThisWorkbook.Worksheets(OP_SHEET)     ' (review) never left unprotected
    LogError "OpensOnOpen", en, em
End Sub

Public Function OpensTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(OP_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(OP_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(OP_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet     ' adding a sheet activates it - put the user back
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = OP_SHEET
            RelockStructure
            SetTabColor ws, RGB(30, 56, 100)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, OP_N).Value = Array("When", "Login", "View", "Release", "Read-only", "Integrity PROBLEMs", "Calendar draw s")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, OP_N), , xlYes)
        lo.Name = OP_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 20
        ws.Columns(2).ColumnWidth = 14
        ws.Columns(3).ColumnWidth = 11
        ws.Columns(4).ColumnWidth = 10
        ws.Columns(5).ColumnWidth = 9
        ws.Columns(6).ColumnWidth = 17
        ws.Columns(7).ColumnWidth = 15
        ws.Columns(1).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked And Not gDeskMode, xlSheetVisible, xlSheetVeryHidden)     ' (review) the Desk view hides admin sheets
    End If
    Set OpensTable = lo
End Function


' SystemCheck's line: the errors logged under this release (and how many nobody has read), the opens in 7 days
Public Function OpensSummaryText() As String
    Dim nOpen As Long, i As Long
    On Error Resume Next                            ' never raises
    Dim lo As ListObject: Set lo = OpensTable(False)
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
            For i = 1 To UBound(v, 1)
                If IsDate(v(i, 1)) Then
                    If CDate(v(i, 1)) >= Now - 7 Then nOpen = nOpen + 1
                End If
            Next i
        End If
    End If
    OpensSummaryText = "Errors this release: " & ErrorsThisRelease() & " (" & UnseenErrorCount() & " unseen); opens in the last 7 days: " & nOpen
End Function
