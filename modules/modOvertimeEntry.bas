Attribute VB_Name = "modOvertimeEntry"
Option Explicit

' ====================================================================
'  CALL OUT ROSTER - CORE MODULE
'  Configuration, shared helpers, eligibility/sorting, OT Log access.
'  Every other module and both forms depend on this one.
' ====================================================================

' --------------------------------------------------------------------
'  CONFIGURATION - change things here, not scattered through the code
' --------------------------------------------------------------------
Public Const CORE_VERSION As String = "Core v26.5 (Oct 19)"
Public Const ROSTER_SHEET As String = "OT Roster"     ' tab name; the code also finds it by its headers
Public Const LOG_SHEET As String = "OT Log"
Public Const LOG_TABLE As String = "tblOTLog"
Public Const CALLOUT_SHEET As String = "Blank Call Out Roster"
Public Const CALLOUT_PHONE_COL As Long = 5          ' phone column on the printed call out roster (hidden for non-admins) - v21
Public Const POSTING_SHEET As String = "Overtime Posting"
Public Const REFUSAL_SHEET As String = "Refusal Report"
Public Const HISTORY_SHEET As String = "History Report"      ' v26: the employee history report, member-safe (output only; supervisors and admins)
Public Const HISTORY_ADMIN_SHEET As String = "History Report (admin)"   ' v26: the FULL flavour (who logged, when): admin only, like the Refusal Report
Public Const MINSTAFF_SHEET As String = "Min Staff"           ' v26: minimum staffing per unit and shift (admin-typed settings table)
Public Const MINSTAFF_TABLE As String = "tblMinStaff"

Public Const SHEET_PASSWORD As String = "mytest"
' Windows usernames allowed to run admin macros (comma separated, any case)
Public Const ADMIN_USERS As String = "piercyg,lanthierr"
' v25.3: the logins that land on the Desk (the second admin's screen, modDesk). A login on BOTH lists has
' every admin right and lands on the Desk; on ADMIN_USERS alone it gets today's layout; on DESK_USERS alone
' it is an ordinary user. Set at install (Graeme, Oct 17: a placeholder ships).
Public Const DESK_USERS As String = "lanthierr"
' v25.5 (Graeme, Oct 18, 2026): the SUPERVISORS by login. A login on this list (or on ADMIN_USERS) may do the
' supervisor's jobs (sick call, new posting, award to the list, print, publish) and lands on the supervisor Desk
' (SUPERVISOR_DESK); every other login is a plain user (log a call-out, the calendar, the sorts) on the roster.
Public Const SUPERVISOR_USERS As String = "mihaychukc,sheltond,campeauc"
Public Const PRIMARY_ADMIN_NAME As String = "Graeme"                       ' named by the Desk's help buttons
Public Const PRIMARY_ADMIN_EMAIL As String = "graeme.piercy@syensqo.com"   ' "Something went wrong" writes to this address
Public Const GUIDE_PATH As String = ""                                    ' the quick guide as a PDF on the plant PC; "" until go-live
Public Const SUPERVISOR_DESK As Boolean = True                            ' a login on SUPERVISOR_USERS lands on the supervisor Desk (Graeme, Oct 17 / v25.5); False = the roster
' v25.5 (briefs/v25.5.md item 2, D1 taken by default Oct 18): every pop-up's positions, sizes and point sizes are
' multiplied by this, for everyone. 1 = the sizes as written (the call-out pop-up is 500 x 615 pt = 820 px tall at 96 dpi:
' a 1920 x 1080 screen; a 1366 x 768 screen cuts its bottom - Graeme's call, Oct 18); 1.2 on a 1080p screen only.
Public Const UI_SCALE As Double = 1                                       ' (ScaleForm / UiPx in modUI apply it)

Public Const CHARGE_MULTIPLIER As Double = 1.75   ' refused / passed OT is charged at this rate
Public Const CHARGE_MAX_GRADE As Double = 2       ' grades above this are called but never charged

' Posting rules (paper posting sheet AND the posting form use the same list)
Public Const POSTING_USE_GRADE_CUTOFFS As Boolean = False  ' False = anyone with a code in that department is listed
' (POSTING_INCLUDE_DAYS removed in v20 - it was never used: DAYS people are on postings
'  whenever they're free that date, the same as everyone else)
Public Const CREW_ALL As String = "All not scheduled"   ' crew picker default: everyone with the code who isn't scheduled that day

' Callout form rule
Public Const CALLOUT_INCLUDE_DAYS As Boolean = True        ' DAYS employees appear for every shift

' Clock-change nights (v21): True = New Posting / Log Call-Out default Hours to the
' actual 13 h (fall back) or 11 h (spring forward); False = keep 12 and only warn
Public Const DST_USE_ACTUAL_HOURS As Boolean = True
Public Const USUAL_SHIFT_HOURS As Double = 12
' v26.5.5 (item 9, Graeme Oct 19-21): the plant's fatigue rule. At most FATIGUE_MAX_HOURS of consecutive
' work (the shift already worked that touches the OT, plus the OT), and at least FATIGUE_MIN_REST before the
' next shift. Accepted call-outs, awards and switches are WORKED shifts for it (modCalendar.FatigueMaxFor).
Public Const FATIGUE_MAX_HOURS As Double = 16
Public Const FATIGUE_MIN_REST As Double = 11

' {YEAR} is replaced by the roster year (v23), so reports follow the rollover: use ReportFolder()
Public Const REPORT_FOLDER As String = "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\{YEAR}\"
' Where Member View saves the members' file. "." = the folder the roster is in (for now: Oct 12, 2026).
' v26.5.5 (A13): at the install Graeme names the MEMBERS' folder and this constant takes its absolute
' path (the sign-up inbox below follows it: ".\Signups" = beside the members' file); the master stays in
' the supervisor-only folder with its backups, pre-change copies, the FINAL and the error log. Until the
' folder is named, "." keeps everything beside the master. SystemCheck shows both folders. The P: path:
'   "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\"
Public Const PUBLISH_FOLDER As String = "."
Public Const PUBLISH_FILE As String = "Call Out Roster - Member View.xlsx"   ' the old values-only file (no longer written from M1)

' Members' workbook (M1): the template sits next to this workbook; the published copy goes to PUBLISH_FOLDER
Public Const MEMBER_TEMPLATE_NAME As String = "Call Out Roster - Members TEMPLATE.xlsm"
Public Const MEMBER_FILE_NAME As String = "Call Out Roster - Members.xlsm"
Public Const MEMBER_PASSWORD As String = "members"  ' sheet protection in the members' file (modMember uses the same)
Public Const MEMBERS_STALE_DAYS As Long = 1         ' Needs Attention reminds after this many days
Public Const MEMBER_NEXT_DAYS As Long = 14          ' Who's next horizon
Public Const PUB_NAME As String = "MembersPublishedAt"
' Electronic sign-ups (v24.1): the members' workbook drops one small file per sign-up into this
' folder and Award Posting reads them. ".\Signups" = a folder beside the members' file; every
' member needs WRITE access to it (and can list it: the file names carry the login that signed up).
' On since v26.5 for the beta; the supervisor's tick is what counts, the paper sheet stays valid.
Public Const SIGNUPS_ON As Boolean = True       ' v26.5 (Graeme, Oct 19): on for the beta - members' Sign me up / Withdraw write into the Signups folder beside the members' file
Public Const SIGNUP_FOLDER As String = ".\Signups"
Public Const SIGNUP_STALE_DAYS As Long = 7          ' IntegrityCheck: inbox files this old for a closed posting
Public Const KRONOS_PATH As String = "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\Kronos - Raw Data Processing.xlsm"   ' v26.5: the Kronos EXPORT file ("Employee Transactions and Totals"), or the FOLDER the exports are saved in (ending in \: the newest file is read); the pivot workbook still works, without the per-period hours
Public Const KRONOS_SHEET As String = "Sorted Data Here"
Public Const KRONOS_STALE_DAYS As Long = 7            ' warn on call-outs / postings when Kronos data is older than this
Public Const PROV_NAME As String = "KronosDataThrough" ' hidden name: last date the Kronos data covers (v21)
Public Const REFUSED_NAME As String = "RefusedThrough"  ' v26.5 hidden name: the current pay period's start - Refused (I) counts the periods up to it

' Daily backup (v21): the first admin open each day saves "Call Out Roster yyyy-mm-dd.xlsm"
' here and deletes copies older than BACKUP_KEEP_DAYS. Blank = backups off (SystemCheck says so).
' "." = the folder the roster is in (Graeme, Oct 2026); ".\Backups" = a subfolder of it (must exist).
Public Const BACKUP_FOLDER As String = "."
Public Const BACKUP_KEEP_DAYS As Long = 30
Public Const BACKUP_BASENAME As String = "Call Out Roster"

' Year end (v23). The frozen copy "Call Out Roster 2026 FINAL.xlsm" goes here at the rollover.
' "." = the folder the roster is in; ".\Archive" = a subfolder of it (must exist).
Public Const ARCHIVE_FOLDER As String = "."
Public Const PAYBACK_WARN_DAYS As Long = 42            ' IntegrityCheck: a shift-switch payback owed longer than this
Public Const AWAIT_NAME As String = "KronosAwaitFrom"  ' hidden name: after a rollover, Worked (H) stays 0 until Kronos data reaches this date

' --------------------------------------------------------------------
'  ROSTER LAYOUT (CALL IN_JAN)
' --------------------------------------------------------------------
Public Const R_HDR As Long = 4          ' sort / filter header row (row just above the employees)
Public Const PERIOD_FMT As String = "General;-General;"   ' hides zeros; NO text (@) section
' Pay-period START and END dates live in rows 2 and 4 - see PStartRow / PEndRow
Public Const R_FIRST As Long = 5        ' first employee row
Public Const C_PAYROLL As Long = 1      ' A  Payroll # from Kronos (v24.2) - admin view only, never published, not a key yet
Public Const C_HOME As Long = 2         ' B  Home Unit (v21) - one of DeptList; blank = Graeme fills it in
Public Const C_PROV As Long = 3         ' C  (hidden) provisional accepted OT not yet in Kronos (v21)
Public Const C_NAME As Long = 4         ' D
Public Const C_SHIFT As Long = 5        ' E
Public Const C_PHONE As Long = 6        ' F
Public Const C_TOTAL As Long = 7        ' G  = Worked + Refused + Provisional (v21)
Public Const C_WORKED As Long = 8       ' H  (from Kronos)
Public Const C_REFUSED As Long = 9      ' I  = the period cells up to the current period (v26.5 RefusedFormula)
Public Const C_SENIORITY As Long = 10   ' J
Public Const C_EMAIL As Long = 11       ' K  operator email (award emails) - admin view only
Public Const C_PERIOD1 As Long = 20     ' T  first pay period
Public Const C_PERIODN As Long = 45     ' AS 26th pay period (the last one in a 26-period year)
Public Const C_PERIODX As Long = 46     ' AT 27th pay period: in use only in a 27-period year (first: 2033) - see PeriodLastCol
Public Const C_LAST As Long = 52        ' AZ - right edge used by sorts (same as the old macros)

' --------------------------------------------------------------------
'  OT LOG LAYOUT (tblOTLog) - column numbers inside the table
'  Event block (same words as Postings) | person block | audit block
'  (UpgradeOTLogLayout converts the old 20-column layout to this one)
' --------------------------------------------------------------------
Public Const L_REF As Long = 1          ' Ref ID     (S-#### / P-#### / ADJ-####)
Public Const L_TYPE As Long = 2         ' Type       (Callout / Posting / Adjustment)
Public Const L_DATE As Long = 3         ' Shift Date
Public Const L_DEPT As Long = 4         ' Unit
Public Const L_SHIFTLBL As Long = 5     ' Shift      (Days / Nights)
Public Const L_CREW As Long = 6         ' Crew       (A/B, C, ...)
Public Const L_TASK As Long = 7         ' Reason
Public Const L_BASEHRS As Long = 8      ' Hours
Public Const L_ORDER As Long = 9        ' Call Order
Public Const L_EMP As Long = 10         ' Employee
Public Const L_CODE As Long = 11        ' Dept Code
Public Const L_EMPHRS As Long = 12      ' Total Hrs Then
Public Const L_OUTCOME As Long = 13     ' Outcome
Public Const L_REASON As Long = 14      ' Note
Public Const L_CHARGED As Long = 15     ' Hours Charged
Public Const L_PERIOD As Long = 16      ' Period Start
Public Const L_ENTRY As Long = 17       ' Entry ID   (audit block from here)
Public Const L_USER As Long = 18        ' Logged By
Public Const L_STAMP As Long = 19       ' Timestamp
Public Const L_STATUS As Long = 20      ' Status
Public Const L_VOIDNOTE As Long = 21    ' Void Note
Public Const L_NCOLS As Long = 21

' --------------------------------------------------------------------
'  EMPLOYEE RECORD used by both forms and the printed outputs
' --------------------------------------------------------------------
Public Type EmpRec
    EmpName As String
    EmpRow As Long
    TotalHrs As Double
    ProvHrs As Double           ' part of TotalHrs that is provisional accepted OT (v21)
    Seniority As Double
    DeptCode As String
    Grade As Double
    Phone As String
    EmpShift As String
    Outcome As String
    Reason As String
    RosterHrs As Double         ' v24.3: Total Hours on the roster (set by LoadEligible)
    EstHrs As Boolean           ' v24.3: TotalHrs is a planned mover's hours AFTER the top-up, an estimate for call order
    MaxHours As Double          ' v26.5.5 (item 9): the most OT they may take on the listed shift under the 16 h rule (FilterNotFree sets it)
End Type

' Session flags
Public gAdminUnlocked As Boolean    ' admin view: sheets left unprotected (except the OT Log)
Public gNoSaveOnClose As Boolean    ' set when a migration/verification fails - never save this state
' v24.2: SmokeTest (modSmokeTest) sets gSmokeRun on a practice copy only. While it is True
' nothing is saved to disk, confirms answer themselves (Ask) and success messages stay quiet (Tell).
Public gSmokeRun As Boolean
Public gSmokeFolder As String       ' where the publish step writes its scratch members' file
Public gSmokeSaved As String        ' ...and the file it wrote ("" = it didn't)
Public gSmokeRefusal As String      ' v25: why the publish refused (the box it would have shown), for SmokeTest's planted-word step
Public gSaveFailedAt As Date        ' v26.5.5 (A1): when the last save to disk FAILED (0 = the last save went through); the Desk header shows it
Public gOpenedMtime As Date         ' v26.5.5 (A2): the file's on-disk time at open and after every own save (0 = not readable: the guard is off)


' ====================================================================
'  LAUNCH FORMS (button names unchanged)
' ====================================================================
Sub ShowOvertimeForm()
    On Error GoTo Oops
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    frmOvertimeEntry.Show
    Exit Sub
Oops:
    ReportError "ShowOvertimeForm"
End Sub

' Award Posting pop-up (v25.3: every supervisor; Award Out of Dept, No Takers and Cancel inside it stay admin)
Sub ShowPostingForm()
    On Error GoTo Oops
    If Not RequireSupervisor() Then Exit Sub     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    frmPostingEntry.Show
    Exit Sub
Oops:
    ReportError "ShowPostingForm"
End Sub


' ====================================================================
'  LISTS
' ====================================================================
Public Function LogHeaders() As Variant
    LogHeaders = Array("Ref ID", "Type", "Shift Date", "Unit", "Shift", "Crew", _
                       "Reason", "Hours", "Call Order", "Employee", "Dept Code", _
                       "Total Hrs Then", "Outcome", "Note", "Hours Charged", "Period Start", _
                       "Entry ID", "Logged By", "Timestamp", "Status", "Void Note")
End Function

Public Function DeptList() As Variant
    DeptList = Array("Phosphine", "Derivatives", "Cypure", "Material Handler", _
                     "QC Gases", "QC Derivatives", "Utilities", "SBP")
End Function

Public Function ShiftList() As Variant
    ShiftList = Array("A", "B", "A/B", "C", "D", "C/D", "DAYS")
End Function

' Crew picker on the posting and call-out pop-ups (CREW_ALL is in the settings at the top)
Public Function CrewPickList() As Variant
    CrewPickList = Array(CREW_ALL, "A", "B", "A/B", "C", "D", "C/D", "DAYS")
End Function

' crew picker text -> LoadEligible's shift filter
Public Function CrewFilter(ByVal sel As String) As String
    If Trim$(sel) = "" Or StrComp(Trim$(sel), CREW_ALL, vbTextCompare) = 0 Then CrewFilter = "ALL" Else CrewFilter = sel
End Function

Public Function TaskList() As Variant
    ' coverage of someone off ... | additional support (nobody off)
    TaskList = Array("Vacation Coverage", "Leave Coverage", "Training Coverage", _
                     "PHA", "Unit Operator Support", "Other")
End Function

' Add Adjustment must pick one of these (v21). Stored in the OT Log's
' Reason column (7) so reports can split adjustments by type; the Type
' column stays "Adjustment".
Public Function AdjustmentTypes() As Variant
    AdjustmentTypes = Array("Leave return", "Unit move", "New hire", "Correction", "Other")
End Function

Public Function UnavailableReasons() As Variant
    UnavailableReasons = Array("Vacation", "Already Working OT", "Training", _
                               "Scheduled Off Site", "Modified Duties", _
                               "Not Qualified for Task", "Other", "Working days elsewhere")   ' v26.5.5 (Graeme, Oct 21): was "Days"
End Function

' Department -> roster column holding that department's grade code
Public Function DeptCol(ByVal d As String) As Long
    Select Case d
        Case "QC Derivatives": DeptCol = 12     ' L - QC LAB
        Case "Derivatives": DeptCol = 13        ' M - DERIVS
        Case "Cypure": DeptCol = 14             ' N - GASES
        Case "Phosphine": DeptCol = 15          ' O - PH3
        Case "Material Handler": DeptCol = 16   ' P - MH
        Case "QC Gases": DeptCol = 17           ' Q - QC GAS
        Case "Utilities": DeptCol = 18          ' R - UTL
        Case "SBP": DeptCol = 19                ' S - SBP
        Case Else: DeptCol = 0
    End Select
End Function

' Unit column number -> department name (reverse of DeptCol); "" if not a unit column
Public Function DeptOfCol(ByVal col As Long) As String
    Dim d As Variant
    For Each d In DeptList()
        If DeptCol(CStr(d)) = col Then DeptOfCol = CStr(d): Exit Function
    Next d
End Function

' First and last unit-code columns on the roster (L:S)
Public Function UnitColFirst() As Long
    UnitColFirst = 12
End Function

Public Function UnitColLast() As Long
    UnitColLast = 19
End Function

' Home Unit for one roster row read as a 2-D array slice: the unit whose
' column holds that person's ONLY code; "" when they have codes in more
' than one unit (or none). v21 - SetupHomeUnit uses it to pre-fill B.
Public Function SingleUnitFor(ByVal rowVals As Variant, ByVal rowIdx As Long) As String
    Dim d As Variant, n As Long, hit As String
    For Each d In DeptList()
        If Trim$(SafeStr(rowVals(rowIdx, DeptCol(CStr(d))))) <> "" Then
            n = n + 1
            hit = CStr(d)
        End If
    Next d
    If n = 1 Then SingleUnitFor = hit
End Function

' Names with a unit code but no Home Unit (column B), A-Z. Empty if none.
Public Function BlankHomeUnits() As Variant
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, UnitColLast())).Value
    Dim i As Long, n As Long, nm As String, d As Variant, hasCode As Boolean
    Dim a() As String: ReDim a(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        If nm <> "" And Trim$(SafeStr(v(i, C_HOME))) = "" Then
            hasCode = False
            For Each d In DeptList()
                If Trim$(SafeStr(v(i, DeptCol(CStr(d))))) <> "" Then hasCode = True: Exit For
            Next d
            If hasCode Then n = n + 1: a(n) = nm
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve a(1 To n)
    Dim j As Long, t As String
    For i = 2 To n
        t = a(i): j = i - 1
        Do While j >= 1
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
    BlankHomeUnits = a
End Function

' Department -> section title on the Blank Call Out Roster sheet
Public Function DeptRosterTitle(ByVal d As String) As String
    Select Case d
        Case "Derivatives": DeptRosterTitle = "Derivatives Call Out Roster"
        Case "Phosphine": DeptRosterTitle = "Phosphine Call Out Roster"
        Case "Cypure": DeptRosterTitle = "Cypure Call Out Roster"
        Case "Utilities": DeptRosterTitle = "Utilities Call Out Roster"
        Case "Material Handler": DeptRosterTitle = "Fumigation Call Out Roster"
        Case "QC Derivatives": DeptRosterTitle = "QC Derivatives Lab Call Out Roster"
        Case "QC Gases": DeptRosterTitle = "QC Gases Lab Call Out Roster"
        Case "SBP": DeptRosterTitle = "SBP Call Out Roster"
    End Select
End Function

' Highest grade listed on an overtime posting (carried over from the old
' OT_Posting filters: P1.5, D2.5, G2, M1; QC lists everyone)
Public Function DeptPostingMaxGrade(ByVal d As String) As Double
    If Not POSTING_USE_GRADE_CUTOFFS Then DeptPostingMaxGrade = 999: Exit Function
    Select Case d
        Case "Phosphine": DeptPostingMaxGrade = 1.5
        Case "Derivatives": DeptPostingMaxGrade = 2.5
        Case "Cypure": DeptPostingMaxGrade = 2
        Case "Material Handler": DeptPostingMaxGrade = 1
        Case Else: DeptPostingMaxGrade = 999
    End Select
End Function


' ====================================================================
'  SHIFT MATCHING
'  - A/B and C/D (or A&B, C&D) are treated as "either"
'  - an A/B employee matches A, B or A/B
'  - DAYS employees match DAYS, and every other shift if includeDays
'  - "ALL" matches everyone
' ====================================================================
Public Function ShiftMatches(ByVal empShift As String, ByVal selectedShift As String, _
                             Optional ByVal includeDays As Boolean = True) As Boolean
    Dim es As String: es = NormShift(empShift)
    Dim ss As String: ss = NormShift(selectedShift)

    If ss = "ALL" Then ShiftMatches = True: Exit Function
    If es = "" Or ss = "" Then Exit Function

    If es = "DAYS" Then
        ShiftMatches = (ss = "DAYS") Or includeDays
        Exit Function
    End If
    If ss = "DAYS" Then Exit Function

    Dim e As Variant, s As Variant
    For Each e In Split(es, "/")
        For Each s In Split(ss, "/")
            If e <> "" And e = s Then ShiftMatches = True: Exit Function
        Next s
    Next e
End Function

Private Function NormShift(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, "&", "/")
    s = Replace(s, " ", "")
    If s = "DAY" Then s = "DAYS"
    NormShift = s
End Function


' Numeric part of a grade code:  D1 -> 1, P1.5 -> 1.5, G3 -> 3, blank -> 999
Public Function DeptCodeValue(ByVal code As String) As Double
    Dim numStr As String, ch As String, i As Long
    Dim foundDigit As Boolean
    For i = 1 To Len(code)
        ch = Mid$(code, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            numStr = numStr & ch
            foundDigit = True
        ElseIf foundDigit Then
            Exit For
        End If
    Next i
    If numStr <> "" And IsNumeric(numStr) Then
        DeptCodeValue = Val(numStr)
    Else
        DeptCodeValue = 999
    End If
End Function


' ====================================================================
'  ELIGIBLE LIST - the ONE place call order is decided
'  Order: grade (low first) > total hours (low first) > seniority
'  (earliest date first) > name.
'  Returns the count; recs(1..count) is filled and sorted.
' ====================================================================
Public Function LoadEligible(ByVal dept As String, ByVal shiftSel As String, _
                             ByRef recs() As EmpRec, _
                             Optional ByVal maxGrade As Double = 999, _
                             Optional ByVal includeDays As Boolean = True, _
                             Optional ByVal asOf As Variant) As Long
    EnsureRefusedThrough                                     ' v26.5: a session open across a period start keeps I, G and the call order current
    ' asOf (v24): a shift date. Each person's crew is then the crew they are on
    ' that day (crew history); left out, it is the roster's crew column.
    ' v24.3: a unit move that is waiting counts from its date (UnitCodeOn): the
    ' person is off the cleared units' lists and on the new unit's, at the hours
    ' they would have after the top-up (PlannedListHours). With no asOf the date
    ' is today, so a move that is due but not applied yet already counts.
    Dim dc As Long: dc = DeptCol(dept)
    If dc = 0 Then Exit Function

    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function

    Dim v As Variant
    v = ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, C_LAST)).Value
    ' column C only means "provisional" once RefreshRosterFormulas has written it (v21)
    Dim provLive As Boolean: provLive = ProvisionalLive()

    Dim n As Long, i As Long, nm As String, code As String, g As Double, crewTxt As String
    Dim hasAsOf As Boolean, asOfDate As Date
    If Not IsMissing(asOf) Then
        If IsDate(asOf) Then hasAsOf = True: asOfDate = CDate(asOf)
    End If
    ReDim recs(1 To UBound(v, 1))

    ' planned unit moves (v24.3): everyone holding a code in this unit on the date,
    ' whatever their crew or grade, is a peer for the mover's listed hours
    Dim planned As Boolean: planned = HasPlannedMoves()
    Dim planDate As Date: planDate = Date
    If hasAsOf Then planDate = asOfDate
    Dim movedIn As Boolean, anyMoved As Boolean, nAll As Long
    Dim allGrade() As Double, allTotal() As Double, allMover() As Boolean, recMover() As Boolean
    If planned Then
        ReDim allGrade(1 To UBound(v, 1)): ReDim allTotal(1 To UBound(v, 1))
        ReDim allMover(1 To UBound(v, 1)): ReDim recMover(1 To UBound(v, 1))
    End If

    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        code = Trim$(SafeStr(v(i, dc)))
        movedIn = False
        If planned And nm <> "" Then
            code = UnitCodeOn(nm, dept, planDate, code, movedIn)
            If code <> "" Then
                nAll = nAll + 1
                allGrade(nAll) = DeptCodeValue(code)
                allTotal(nAll) = NumOr0(v(i, C_TOTAL))
                allMover(nAll) = movedIn
            End If
        End If
        If nm <> "" And code <> "" Then
            ' v26.5.5 (A4): an error in G (a #REF! after a row delete, #NAME? when the RefusedThrough name is
            ' missing on a copy) read as 0 and put the person at the TOP of the call order. It is refused instead.
            If IsError(v(i, C_TOTAL)) Then RaiseTotalError nm
            If provLive Then
                If IsError(v(i, C_REFUSED)) Or IsError(v(i, C_PROV)) Then RaiseTotalError nm
            End If
            crewTxt = Trim$(SafeStr(v(i, C_SHIFT)))
            If hasAsOf Then crewTxt = CrewOn(nm, asOfDate, crewTxt)
            If ShiftMatches(crewTxt, shiftSel, includeDays) Then
                g = DeptCodeValue(code)
                If g <= maxGrade Then
                    n = n + 1
                    With recs(n)
                        .EmpName = nm
                        .EmpRow = R_FIRST + i - 1
                        .DeptCode = code
                        .Grade = g
                        .TotalHrs = NumOr0(v(i, C_TOTAL))
                        .RosterHrs = .TotalHrs
                        .EstHrs = False
                        If provLive Then .ProvHrs = NumOr0(v(i, C_PROV)) Else .ProvHrs = 0
                        .Seniority = SeniorityKey(v(i, C_SENIORITY))
                        .Phone = Trim$(SafeStr(v(i, C_PHONE)))
                        .EmpShift = crewTxt
                        .Outcome = ""
                        .Reason = ""
                        .MaxHours = FATIGUE_MAX_HOURS       ' v26.5.5: no limit until FilterNotFree looks at the date and shift
                    End With
                    If movedIn Then
                        recMover(n) = True
                        anyMoved = True
                    End If
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    If anyMoved Then                                ' before the sort: these hours decide the mover's place
        Dim listHrs As Double
        For i = 1 To n
            If recMover(i) Then
                listHrs = PlannedListHours(recs(i).TotalHrs, allGrade, allMover, allTotal, nAll)
                If listHrs > recs(i).TotalHrs Then
                    recs(i).TotalHrs = listHrs
                    recs(i).EstHrs = True
                End If
            End If
        Next i
    End If
    ReDim Preserve recs(1 To n)
    SortEmpRecs recs, n
    LoadEligible = n
End Function

' ---- v26.5.5 (5.2): the trust strip and the two hour texts ----
' One line that says what the figures on a pop-up, a printed roster, the board or the month PDF rest on:
' how current the worked hours are, how many totals carry provisional OT, that Refused is capped at the
' current pay period, and when the members' file was last updated.
Public Function TrustStrip() As String
    Dim dt As Date: dt = KronosDataThrough()
    Dim s As String
    If dt = 0 Then s = "Worked hours: none imported from Kronos yet" Else s = "Worked hours through " & Format$(dt, "ddd mmm d") & " (Kronos)"
    Dim nProv As Long: nProv = ProvisionalCount()
    s = s & "  |  " & IIf(nProv = 0, "no", CStr(nProv)) & " total" & IIf(nProv = 1, "", "s") & " with provisional OT"
    s = s & "  |  Refused counted to the current pay period (from " & Format$(RefusedThrough(), "mmm d") & ")"
    Dim p As Date: p = MembersPublishedAt()
    s = s & "  |  Members' file " & IIf(p = 0, "never published", "updated " & Format$(p, "ddd h:mm AM/PM"))
    TrustStrip = s
End Function

' How many roster totals include provisional accepted OT (column C > 0) right now
Public Function ProvisionalCount() As Long
    If Not ProvisionalLive() Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Function
    Dim v As Variant: v = ws.Range(ws.Cells(R_FIRST, C_PROV), ws.Cells(lastRow, C_PROV)).Value2
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If NumOr0(v(i, 1)) > 0 Then ProvisionalCount = ProvisionalCount + 1
    Next i
End Function

' "21 charged (12 h x 1.75)" - the hours on offer and what they cost
Public Function ChargedText(ByVal hours As Double, ByVal charged As Double) As String
    ChargedText = HoursText(charged) & " charged (" & HoursText(hours) & " h x " & CHARGE_MULTIPLIER & ")"
End Function

' "141 h (to Oct 18)" - a total and the date its worked hours run to
Public Function TotalText(ByVal total As Double, ByVal through As Date) As String
    TotalText = HoursText(total) & " h" & IIf(through = 0, "", " (to " & Format$(through, "mmm d") & ")")
End Function

' v26.5.5 (A4): the one message for a roster figure that is an error (LoadEligible, the publish's BuildRoster).
' The pop-ups show it through their handlers (ReportError); the publish stops with it.
Public Sub RaiseTotalError(ByVal nm As String)
    Err.Raise vbObjectError + 776, "LoadEligible", "Total Hours shows an error for " & nm & _
              ". An admin runs Alt+F8 > RefreshRosterFormulas, then try again."
End Sub

' Call order on an in-memory list (public from v21 so SelfTest can use it)
Public Sub SortEmpRecs(recs() As EmpRec, ByVal n As Long)
    Dim i As Long, j As Long, tmp As EmpRec
    For i = 2 To n
        tmp = recs(i)
        j = i - 1
        Do While j >= 1
            If RecBefore(tmp, recs(j)) Then
                recs(j + 1) = recs(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        recs(j + 1) = tmp
    Next i
End Sub

' v24.3.1. Where a newcomer goes in a list that is already in order and must
' not be re-sorted (a posting's frozen list): before the first row they come
' before in call order. Returns 1 .. n + 1. Pure; mirrored in tests/rules.py.
Public Function ListInsertIndex(ByRef recs() As EmpRec, ByVal n As Long, ByRef newRec As EmpRec) As Long
    Dim i As Long
    For i = 1 To n
        If RecBefore(newRec, recs(i)) Then
            ListInsertIndex = i
            Exit Function
        End If
    Next i
    ListInsertIndex = n + 1
End Function

Private Function RecBefore(a As EmpRec, b As EmpRec) As Boolean
    If a.Grade <> b.Grade Then RecBefore = (a.Grade < b.Grade): Exit Function
    ' v25 (E6): hours are compared to the hundredth - two totals that differ only in floating-point dust tie
    If Round(a.TotalHrs, 2) <> Round(b.TotalHrs, 2) Then RecBefore = (a.TotalHrs < b.TotalHrs): Exit Function
    If a.Seniority <> b.Seniority Then RecBefore = (a.Seniority < b.Seniority): Exit Function
    RecBefore = (StrComp(a.EmpName, b.EmpName, vbTextCompare) < 0)
End Function

' Hours charged for one call-out outcome (v21: the one place for this rule).
' Refused and Left Message are charged at CHARGE_MULTIPLIER, grade
' CHARGE_MAX_GRADE and below only. Accepted / Unavailable: nothing.
Public Function CalloutCharge(ByVal outcome As String, ByVal grade As Double, ByVal hours As Double) As Double
    If grade > CHARGE_MAX_GRADE Then Exit Function
    Select Case outcome
        Case "Refused", "Left Message": CalloutCharge = hours * CHARGE_MULTIPLIER
    End Select
End Function

' ---- part of a call-out accepted (v24.1) ----
' Someone can take part of a call-out ("8 of 12"). Calling then continues
' down the list for what is left, so the hours ON OFFER fall by what was
' taken for everyone BELOW that person in call order. A whole-shift
' acceptance changes nothing for the people below (as before v24.1: a
' second opening on the same call-out is still the full hours).
'   outcomes(i)  the outcome recorded for person i ("" = not called)
'   taken(i)     hours an Accepted person took; 0 = everything on offer
'   offers(i)    out: the hours person i was offered
Public Sub CalloutOffers(ByVal hours As Double, ByRef outcomes() As String, ByRef taken() As Double, _
                         ByVal n As Long, ByRef offers() As Double)
    Dim i As Long, remain As Double
    remain = hours
    ReDim offers(1 To IIf(n > 0, n, 1))
    For i = 1 To n
        offers(i) = remain
        If outcomes(i) = "Accepted" Then
            ' v25 (D-E4, Graeme Oct 14): a PARTIAL acceptance closes the shift - the rest is dropped,
            ' uncharged; nobody below is offered it (v24.1 offered the rest down the list)
            If taken(i) > 0 And taken(i) < remain Then remain = 0
        End If
    Next i
End Sub

' v25 (D-E4): has a partial acceptance closed this call-out's shift? (the pop-up refuses a further Accepted)
Public Function CalloutClosed(ByRef outcomes() As String, ByRef taken() As Double, ByVal n As Long, ByVal hours As Double) As Boolean
    Dim i As Long
    For i = 1 To n
        If outcomes(i) = "Accepted" Then
            If taken(i) > 0 And taken(i) < hours Then CalloutClosed = True: Exit Function
        End If
    Next i
End Function

' Hours an Accepted person works: the part they took, or everything on offer
Public Function AcceptedHours(ByVal offer As Double, ByVal taken As Double) As Double
    If taken > 0 And taken < offer Then AcceptedHours = taken Else AcceptedHours = offer
End Function

' Does one OT Log row count as provisional accepted OT right now? Mirrors
' the column-C SUMIFS exactly (SelfTest / tests check this against it).
Public Function ProvisionalCounts(ByVal outcome As String, ByVal status As String, ByVal shiftDate As Date, _
                                  ByVal isNight As Boolean, ByVal dataThrough As Date, _
                                  ByVal yearFrom As Date, ByVal yearTo As Date) As Boolean
    If UCase$(status) = "VOID" Then Exit Function
    If outcome <> "Accepted" And outcome <> "Awarded" Then Exit Function
    If Int(shiftDate) < Int(yearFrom) Or Int(shiftDate) > Int(yearTo) Then Exit Function
    ProvisionalCounts = Not ProvisionalCovered(shiftDate, isNight, dataThrough)
End Function

Public Function SeniorityKey(ByVal x As Variant) As Double
    If IsError(x) Or IsEmpty(x) Then
        SeniorityKey = 1E+15
    ElseIf IsDate(x) Then
        SeniorityKey = CDbl(CDate(x))
    ElseIf IsNumeric(x) Then
        SeniorityKey = CDbl(x)
    Else
        SeniorityKey = 1E+15
    End If
End Function


' ====================================================================
'  ROSTER HELPERS
' ====================================================================
' The roster tab is found by name, or (if it has been renamed) by its
' headers: EMPLOYEE in D2 and PLANT SENIORITY in J2.
Public Function RosterSheet() As Worksheet
    Dim ws As Worksheet, cand As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ROSTER_SHEET)
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets("CALL IN_JAN")     ' original tab name
    On Error GoTo 0
    If ws Is Nothing Then
        ' renamed tab: headers must match; prefer the one whose periods read the OT Log
        For Each cand In ThisWorkbook.Worksheets
            If UCase$(Trim$(SafeStr(cand.Cells(2, C_NAME).Value))) = "EMPLOYEE" And _
               InStr(1, SafeStr(cand.Cells(2, C_SENIORITY).Value), "SENIORITY", vbTextCompare) > 0 Then
                If ws Is Nothing Then Set ws = cand
                If InStr(1, cand.Cells(R_FIRST, C_PERIOD1).Formula, LOG_TABLE, vbTextCompare) > 0 Then Set ws = cand: Exit For
            End If
        Next cand
    End If
    If ws Is Nothing Then Err.Raise vbObjectError + 513, , "Roster sheet not found (looked for '" & ROSTER_SHEET & "' or a sheet with EMPLOYEE in D2)."
    Set RosterSheet = ws
End Function

' Row holding pay-period START dates (2 after the Oct 2026 layout, 4 before)
Public Function PStartRow() As Long
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim a As Variant, b As Variant
    a = ws.Cells(2, C_PERIOD1).Value2: b = ws.Cells(4, C_PERIOD1).Value2
    PStartRow = 4
    If Not IsError(a) And Not IsError(b) Then
        If IsNumeric(a) And IsNumeric(b) And Not IsEmpty(a) And Not IsEmpty(b) Then
            If CDbl(a) < CDbl(b) Then PStartRow = 2
        End If
    End If
End Function

Public Function PEndRow() As Long
    PEndRow = 6 - PStartRow()          ' the other of rows 2 / 4
End Function

' Last row with a name in column D. Reads the cells directly, so rows hidden
' by a filter/sort view are always included (Range.Find skips them).
Public Function RosterLastRow() As Long
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim bottom As Long: bottom = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    RosterLastRow = R_FIRST - 1
    If bottom < R_FIRST Then Exit Function
    Dim v As Variant: v = RangeTo2D(ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(bottom, C_NAME)))
    Dim i As Long
    For i = UBound(v, 1) To 1 Step -1
        If Trim$(SafeStr(v(i, 1))) <> "" Then RosterLastRow = R_FIRST + i - 1: Exit Function
    Next i
End Function

Public Sub ClearRosterFilter(ByVal ws As Worksheet)
    If ws.AutoFilterMode Then
        If ws.FilterMode Then ws.ShowAllData
        ws.AutoFilterMode = False
    End If
    ws.Rows.Hidden = False
End Sub

' The filter covers whole rows (A:AS) so a sort from a dropdown can never
' move one block of columns on its own; dropdowns are only SHOWN on L:S.
' Excel may trim a filter range, so everything below works from the
' ACTUAL filter range (ws.AutoFilter.Range), never an assumed one.
' Leaves an existing filter alone unless the roster grew/shrank.
' Sheet must be unprotected.
Public Sub EnsureRosterFilter(ByVal ws As Worksheet)
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub
    If ws.AutoFilterMode Then
        With ws.AutoFilter.Range
            ' a 27-period year must carry column AT in the filter, or a dropdown sort would leave it behind
            If .Row = R_HDR And .Row + .Rows.Count - 1 = lastRow And _
               .Column <= 12 And .Column + .Columns.Count - 1 >= IIf(PeriodLastCol() = C_PERIODX, C_PERIODX, 19) Then Exit Sub
        End With
        If ws.FilterMode Then ws.ShowAllData
        ws.AutoFilterMode = False
    End If
    ws.Range(ws.Cells(R_HDR, 1), ws.Cells(lastRow, PeriodLastCol())).AutoFilter Field:=1    ' creates it (field 1 always exists)

    Dim rng As Range: Set rng = ws.AutoFilter.Range
    Dim k As Long, col As Long
    For k = 1 To rng.Columns.Count
        col = rng.Column + k - 1
        If col < 12 Or col > 19 Then
            On Error Resume Next              ' a hidden dropdown is cosmetic - never stop for it
            rng.AutoFilter Field:=k, VisibleDropDown:=False
            On Error GoTo 0
        End If
    Next k
End Sub

' Filter field number for a sheet column, based on the actual filter range
Public Function RosterFilterField(ByVal ws As Worksheet, ByVal col As Long) As Long
    RosterFilterField = col - ws.AutoFilter.Range.Column + 1
End Function

' SUMIFS formula for one roster row's pay-period cells (written across T:AS)
Public Function PeriodFormula(ByVal r As Long) As String
    PeriodFormula = "=SUMIFS(" & _
        LOG_TABLE & "[[Hours Charged]:[Hours Charged]]," & _
        LOG_TABLE & "[[Employee]:[Employee]],$" & ColLetter(C_NAME) & r & "," & _
        LOG_TABLE & "[[Period Start]:[Period Start]]," & ColLetter(C_PERIOD1) & "$" & PStartRow() & "," & _
        LOG_TABLE & "[[Status]:[Status]],""<>VOID"")"
End Function

' --------------------------------------------------------------------
'  PROVISIONAL ACCEPTED OT (v21)
'  An Accepted call-out or Awarded posting counts in Total Hours from the
'  moment it is logged, until Kronos covers the shift. Kronos data runs
'  through the day before the file was saved (KronosDataThrough, a hidden
'  workbook name the import keeps up to date). A shift is covered once it
'  has fully ended by the end of that day:
'     day shift on D    covered when D     <= data-through
'     night shift on D  covered when D + 1 <= data-through
'  Only rows whose shift date falls inside this roster's year count, the
'  row's shift Hours (not x 1.75), non-void only.
' --------------------------------------------------------------------
Public Function ProvisionalCovered(ByVal shiftDate As Date, ByVal isNight As Boolean, ByVal dataThrough As Date) As Boolean
    Dim endsOn As Date: endsOn = Int(shiftDate)
    If isNight Then endsOn = endsOn + 1
    ProvisionalCovered = (endsOn <= Int(dataThrough))
End Function

' One roster row's provisional formula (column C). Four SUMIFS over the
' log: Accepted / Awarded x day rows (Shift <> Nights) / night rows.
Public Function ProvFormula(ByVal r As Long) As String
    Dim yearFrom As String, yearTo As String
    yearFrom = "$" & ColLetter(C_PERIOD1) & "$" & PStartRow()
    yearTo = "$" & ColLetter(PeriodLastCol()) & "$" & PEndRow()
    ProvFormula = "=" & ProvSumIfs(r, "Accepted", False, yearFrom, yearTo) & _
                  "+" & ProvSumIfs(r, "Accepted", True, yearFrom, yearTo) & _
                  "+" & ProvSumIfs(r, "Awarded", False, yearFrom, yearTo) & _
                  "+" & ProvSumIfs(r, "Awarded", True, yearFrom, yearTo)
End Function

Private Function ProvSumIfs(ByVal r As Long, ByVal outcome As String, ByVal nights As Boolean, _
                            ByVal yearFrom As String, ByVal yearTo As String) As String
    Dim dateCol As String: dateCol = LOG_TABLE & "[[Shift Date]:[Shift Date]]"
    ProvSumIfs = "SUMIFS(" & LOG_TABLE & "[[Hours]:[Hours]]," & _
        LOG_TABLE & "[[Employee]:[Employee]],$" & ColLetter(C_NAME) & r & "," & _
        LOG_TABLE & "[[Status]:[Status]],""<>VOID""," & _
        LOG_TABLE & "[[Outcome]:[Outcome]],""" & outcome & """," & _
        dateCol & ","">=""&" & yearFrom & "," & _
        dateCol & ",""<=""&" & yearTo & "," & _
        LOG_TABLE & "[[Shift]:[Shift]],""" & IIf(nights, "Nights", "<>Nights") & """," & _
        dateCol & ","">""&" & IIf(nights, "(" & PROV_NAME & "-1)", PROV_NAME) & ")"
End Function

' True once the roster's Total (G) formulas add column C, i.e. after
' RefreshRosterFormulas has run on v21 (checks the first named row)
Public Function ProvisionalLive() As Boolean
    On Error Resume Next
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If Trim$(SafeStr(ws.Cells(r, C_NAME).Value)) <> "" Then
            ProvisionalLive = (InStr(1, ws.Cells(r, C_TOTAL).Formula, "+" & ColLetter(C_PROV) & r, vbTextCompare) > 0)
            Exit Function
        End If
    Next r
End Function

' Total Hours (column G) = Worked + Refused + Provisional
Public Function TotalFormula(ByVal r As Long) As String
    TotalFormula = "=" & ColLetter(C_WORKED) & r & "+" & ColLetter(C_REFUSED) & r & "+" & ColLetter(C_PROV) & r
End Function

' Writes one row's provisional (C) and total (G) formulas. Sheet must be unprotected.
Public Sub WriteTotalRow(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Cells(r, C_PROV)
        .NumberFormat = "General"
        .Formula = ProvFormula(r)
        .NumberFormat = PERIOD_FMT
        .Locked = True
    End With
    With ws.Cells(r, C_TOTAL)
        .NumberFormat = "General"
        .Formula = TotalFormula(r)
        .NumberFormat = "#,##0"
        .Locked = True
    End With
End Sub

Public Function IsOurFormula(ByVal cel As Range) As Boolean
    If cel.HasFormula Then
        IsOurFormula = (InStr(1, cel.Formula, LOG_TABLE, vbTextCompare) > 0)
    End If
End Function

' A period formula that Excel stored as plain text (the cell had a text
' format when it was written). Shows as "=SUMIFS(tblOTLog..." on screen.
Public Function IsBrokenFormula(ByVal cel As Range) As Boolean
    If cel.HasFormula Then Exit Function
    Dim v As Variant: v = cel.Value2
    If VarType(v) = vbString Then
        If Left$(v, 1) = "=" And InStr(1, v, LOG_TABLE, vbTextCompare) > 0 Then IsBrokenFormula = True
    End If
End Function

' Writes one roster row's pay-period formulas. The format is set to General
' FIRST so Excel stores real formulas, then the zero-hiding format goes on.
' Sheet must be unprotected.
Public Sub WritePeriodRow(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Range(ws.Cells(r, C_PERIOD1), ws.Cells(r, PeriodLastCol()))
        .NumberFormat = "General"
        .Formula = PeriodFormula(r)
        .NumberFormat = PERIOD_FMT
        .Locked = True
    End With
End Sub

' --------------------------------------------------------------------
'  REFUSED (I) COUNTS THE PERIODS UP TO THE CURRENT ONE (v26.5, Graeme Oct 19)
'  A posting's refusal charges carry the SHIFT's pay period (Period
'  Start). Until v26.5 Refused summed every period column, so a charge
'  for a shift two periods out counted in Total Hours and the call order
'  today. Now Refused = the charges whose Period Start is on or before
'  the current period's start (the hidden name RefusedThrough, kept by
'  EnsureRefusedThrough on every open and formula refresh); a future
'  period's charges show in their column and start counting when that
'  period begins. Call-outs carry the call-out date's period, so nothing
'  changes for them. Provisional accepted OT (C) is untouched (v21).
' --------------------------------------------------------------------
' PURE (SelfTest "refthrough"): does a charge with this Period Start count today?
Public Function RefusedCounts(ByVal periodStart As Date, ByVal through As Date) As Boolean
    RefusedCounts = (Int(periodStart) <= Int(through))
End Function

Public Function RefusedThrough() As Date
    RefusedThrough = GetDateName(REFUSED_NAME)
End Function

' The current period's start into the hidden name (when it moved, or is missing). Safe any time;
' a read-only copy keeps what it has (the formulas then read the last saved value).
Public Function EnsureRefusedThrough() As Date
    Dim want As Date: want = PeriodStartFor(Date)
    Dim have As Date: have = RefusedThrough()
    ' v26.5.5 (A14): the cap stays inside this roster's year. A PC clock in the past would drop every charge
    ' out of I (everyone's Total = Worked, the call order wrong); a date past the year (the rollover not run
    ' yet) counts every period, which is right. Moving the cap BACK by more than one period is a clock
    ' problem too: an admin is asked, anyone else keeps the stored cap.
    If want < RosterYearFrom() Then want = RosterYearFrom()
    If want > RosterYearTo() Then want = RosterYearTo()
    If have > 0 And want < have - 14 And Not gSmokeRun Then
        If IsAdmin() Then
            If MsgBox("This PC says today is " & Format$(Date, "ddd mmm d, yyyy") & ", more than a pay period before the roster's " & _
                      "last current period (" & Format$(have, "mmm d") & "). Is the PC's date wrong?" & vbNewLine & vbNewLine & _
                      "Yes = keep the roster as it is (fix the clock). No = move the current period back to " & Format$(want, "mmm d") & ".", _
                      vbYesNo + vbExclamation + vbDefaultButton1, "Pay period") = vbYes Then want = have
        Else
            want = have
        End If
    End If
    If have <> want And Not ThisWorkbook.ReadOnly Then
        SetDateName REFUSED_NAME, want: have = want
        Application.Calculate                                ' whoever reads I or G next sees the new cap, manual calculation included
        MarkMembersDirty                                     ' every I, G and the call order moved: the members' file is out of date
    End If
    If have = 0 Then have = want
    EnsureRefusedThrough = have
End Function

' One roster row's Refused formula: the row's own pay-period cells whose header date is on or before
' RefusedThrough (same test as RefusedCounts). Summing the row's cells, not the log, keeps I inside this
' roster's year by construction (the v26.5 reviews: from Dec 27 the current period is next year's first,
' and a log row off the period grid counts nowhere, as before) - exactly what IntegrityCheck and
' RefreshRosterFormulas add up.
Public Function RefusedFormula(ByVal r As Long) As String
    Dim c1 As String: c1 = ColLetter(C_PERIOD1)
    Dim cN As String: cN = ColLetter(PeriodLastCol())
    Dim hdr As Long: hdr = PStartRow()
    RefusedFormula = "=SUMIFS($" & c1 & r & ":$" & cN & r & ",$" & c1 & "$" & hdr & ":$" & cN & "$" & hdr & ",""<=""&" & REFUSED_NAME & ")"
End Function

' True when a Refused cell still holds the pre-v26.5 formula (=SUM of the period cells) - the first open after
' the v26.5 install, before RefreshRosterFormulas rewrites it (IntegrityCheck then expects the old rule)
Public Function OldRefusedFormula(ByVal cel As Range) As Boolean
    If Not cel.HasFormula Then Exit Function
    OldRefusedFormula = (InStr(1, cel.Formula, "SUM(", vbTextCompare) > 0 And InStr(1, cel.Formula, LOG_TABLE, vbTextCompare) = 0)
End Function

' Refused (column I): v26.5 the charges up to the current period (RefusedFormula); was =SUM of the
' row's pay-period cells. Sheet must be unprotected.
Public Sub WriteRefusedRow(ByVal ws As Worksheet, ByVal r As Long)
    EnsureRefusedThrough
    With ws.Cells(r, C_REFUSED)
        .NumberFormat = "General"
        .Formula = RefusedFormula(r)
        .NumberFormat = "#,##0"
        .Locked = True
    End With
End Sub

' Non-void hours charged in the OT Log for one employee, across the pay
' periods shown on this roster. Used to verify the roster after rewrites.
Public Function LogTotalOnRoster(ByVal empName As String, Optional ByVal throughCurrent As Boolean = False) As Double
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim sRow As Long: sRow = PStartRow()
    Dim f As Long: f = CLng(ws.Cells(sRow, C_PERIOD1).Value2)
    Dim l As Long: l = CLng(ws.Cells(sRow, PeriodLastCol()).Value2)
    If throughCurrent Then                                   ' v26.5: as Refused (I) counts - the periods up to the current one
        Dim thr As Long: thr = CLng(Int(EnsureRefusedThrough()))
        If thr < l Then l = thr
    End If
    LogTotalOnRoster = Application.WorksheetFunction.SumIfs( _
        lo.ListColumns(L_CHARGED).DataBodyRange, _
        lo.ListColumns(L_EMP).DataBodyRange, empName, _
        lo.ListColumns(L_PERIOD).DataBodyRange, ">=" & f, _
        lo.ListColumns(L_PERIOD).DataBodyRange, "<=" & l, _
        lo.ListColumns(L_STATUS).DataBodyRange, "<>VOID")
End Function


' ====================================================================
'  PAY PERIODS
'  Periods are 14 days, anchored on the first period START date (T4).
'  Works for any date, including Dec 27-31 (next year's first period).
' ====================================================================
Public Function PeriodStartFor(ByVal d As Date) As Date
    Dim a As Double: a = CDbl(CDate(RosterSheet().Cells(PStartRow(), C_PERIOD1).Value))
    Dim x As Double: x = Int(CDbl(d))
    PeriodStartFor = CDate(a + 14 * Int((x - a) / 14))
End Function

Public Function PeriodColumn(ByVal periodStart As Date) As Long
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim j As Long, v As Variant, sRow As Long: sRow = PStartRow()
    For j = C_PERIOD1 To PeriodLastCol()
        v = ws.Cells(sRow, j).Value
        If IsDate(v) Then
            If CLng(CDbl(CDate(v))) = CLng(CDbl(periodStart)) Then
                PeriodColumn = j
                Exit Function
            End If
        End If
    Next j
End Function

' ---- the roster year (v23) ----
' The year is whatever the roster's date row says: period 1 starts in T2 and
' is the 14-day period that contains Jan 1. A year has 26 periods, or 27
' when one more period fits before the next Jan 1 (first: 2033).

' Last pay-period column in use: AT only when its start-date header is the
' date 14 days after AS's (a 27-period year); otherwise AS.
Public Function PeriodLastCol() As Long
    PeriodLastCol = C_PERIODN
    On Error Resume Next
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim sRow As Long: sRow = PStartRow()
    Dim a As Variant, b As Variant
    a = ws.Cells(sRow, C_PERIODN).Value2: b = ws.Cells(sRow, C_PERIODX).Value2
    If IsError(a) Or IsError(b) Then Exit Function
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If Not IsNumeric(a) Or Not IsNumeric(b) Then Exit Function
    If CDbl(b) = CDbl(a) + 14 Then PeriodLastCol = C_PERIODX
End Function

' First day of period 1 / last day of the last period on this roster
Public Function RosterYearFrom() As Date
    RosterYearFrom = Int(CDate(RosterSheet().Cells(PStartRow(), C_PERIOD1).Value))
End Function

Public Function RosterYearTo() As Date
    RosterYearTo = Int(CDate(RosterSheet().Cells(PEndRow(), PeriodLastCol()).Value))
End Function

' The year this roster is for (2026 for the periods starting Dec 28, 2025)
Public Function RosterYear() As Long
    RosterYear = RosterYearOf(RosterYearFrom())
End Function

' PURE (SelfTest "year"): the year a period-1 start date belongs to
Public Function RosterYearOf(ByVal firstStart As Date) As Long
    RosterYearOf = Year(Int(firstStart) + 13)
End Function

' PURE: first day of the period containing Jan 1 of year y, on the 14-day grid through anchor
Public Function FirstPeriodStart(ByVal anchor As Date, ByVal y As Long) As Date
    Dim a As Double: a = Int(CDbl(anchor))
    Dim x As Double: x = CDbl(DateSerial(y, 1, 1))
    FirstPeriodStart = CDate(a + 14 * Int((x - a) / 14))      ' Int rounds down for dates before the anchor too
End Function

' PURE: 26, or 27 when one more period fits before the next year's period 1
Public Function PeriodsInYear(ByVal anchor As Date, ByVal y As Long) As Long
    PeriodsInYear = CLng((CDbl(FirstPeriodStart(anchor, y + 1)) - CDbl(FirstPeriodStart(anchor, y))) / 14)
End Function

' REPORT_FOLDER with {YEAR} filled in from the roster
Public Function ReportFolder() As String
    Dim y As Long
    On Error Resume Next
    y = RosterYear()
    On Error GoTo 0
    If y < 2000 Then y = Year(Date)
    ReportFolder = Replace(REPORT_FOLDER, "{YEAR}", CStr(y))
End Function

' A folder setting with "." / ".\sub" resolved against the roster's own folder ("" stays "")
Public Function ResolveFolder(ByVal p As String) As String
    p = Trim$(p)
    If p = "" Then Exit Function
    If p = "." Then
        p = ThisWorkbook.Path
    ElseIf Left$(p, 2) = ".\" Then
        p = EnsureSlash(ThisWorkbook.Path) & Mid$(p, 3)
    End If
    ResolveFolder = EnsureSlash(p)
End Function

' True if the period is on this roster, or the user agrees to log it anyway
Public Function ConfirmPeriod(ByVal ps As Date) As Boolean
    If PeriodColumn(ps) > 0 Then ConfirmPeriod = True: Exit Function
    If gSmokeRun Then ConfirmPeriod = True: Exit Function
    ConfirmPeriod = (MsgBox("These hours fall in the pay period starting " & _
        Format$(ps, "mmm d, yyyy") & ", which is not one of this roster's columns." & _
        vbNewLine & vbNewLine & _
        "They will be saved in the OT Log but won't appear in this roster's totals " & _
        "until a roster with that period exists." & vbNewLine & vbNewLine & _
        "Log them anyway?", vbYesNo + vbExclamation, "Pay Period") = vbYes)
End Function


' ====================================================================
'  PAYROLL #  (v24.2) - pure rules for the Kronos import
' ====================================================================
' A Payroll # as text: digits kept as they are (a number loses its decimals,
' text keeps its leading zeros), anything else trimmed. "" = none.
Public Function PayrollText(ByVal x As Variant) As String
    If IsError(x) Or IsEmpty(x) Or IsNull(x) Then Exit Function
    If VarType(x) = vbString Then
        PayrollText = Trim$(x)
    ElseIf IsNumeric(x) Then
        If CDbl(x) > 0 Then PayrollText = Format$(CDbl(x), "0")
    End If
End Function

' Is this header cell the Payroll # column? ("Payroll #", "PAYROLL NO.", "Payroll Number" ...)
Public Function IsPayrollHeader(ByVal s As String) As Boolean
    Dim t As String, i As Long, ch As String
    For i = 1 To Len(s)
        ch = LCase$(Mid$(s, i, 1))
        If ch Like "[a-z0-9#]" Then t = t & ch
    Next i
    IsPayrollHeader = (t = "payroll#" Or t = "payroll" Or t = "payrollno" Or t = "payrollnumber" Or t = "payrollnum" Or t = "payrollid")
End Function

' v24.3.1: does a heading cell say `want`? Letters and digits only, case and spacing ignored:
' "Employee  Name" and "EMPLOYEE NAME" match "Employee Name"; "Name" alone does not; an empty cell
' never matches. Pure; mirrored (krhdr cases). The Kronos import finds its columns with it.
Public Function IsHeaderMatch(ByVal cellText As String, ByVal want As String) As Boolean
    Dim a As String: a = AlnumKey(cellText)
    IsHeaderMatch = (a <> "" And a = AlnumKey(want))
End Function

Private Function AlnumKey(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = LCase$(Mid$(s, i, 1))
        If ch Like "[a-z0-9]" Then AlnumKey = AlnumKey & ch
    Next i
End Function

' How a roster row is matched to Kronos: "name" (no number on the row, or Kronos
' doesn't have that number), "number" (the number is in Kronos under the same
' name), "mismatch" (the number is in Kronos under a DIFFERENT name: nothing is
' changed for that row and the import says so).
Public Function PayrollDecision(ByVal rosterNum As String, ByVal numInKronos As Boolean, ByVal namesAgree As Boolean) As String
    If rosterNum = "" Or Not numInKronos Then
        PayrollDecision = "name"
    ElseIf namesAgree Then
        PayrollDecision = "number"
    Else
        PayrollDecision = "mismatch"
    End If
End Function


' ====================================================================
'  ONE CALL-OUT SESSION -> OT LOG ROWS + CALENDAR LINES  (v24.2)
'  Moved here from the call-out pop-up so there is one copy: the pop-up
'  and SmokeTest both call it. entries(1..n) is the call list in call
'  order with each Outcome ("" = not reached); taken(i) = hours an
'  Accepted person took when it was only part of what was on offer
'  (0 = all of it). False = nothing was written.
' ====================================================================
Public Function SubmitCalloutSession(ByVal dept As String, ByVal shiftLbl As String, ByVal crewSel As String, _
                                     ByVal entryDate As Date, ByVal hours As Double, ByRef entries() As EmpRec, _
                                     ByRef taken() As Double, ByVal n As Long, ByRef sessionID As String, _
                                     ByRef summary As String, Optional ByRef shortSummary As String = "") As Boolean
    ' v25.5: shortSummary = the four-line box the call-out pop-up shows; summary = the full detail (status bar, SmokeTest)
    Dim i As Long, k As Long, logged As Long
    For i = 1 To n
        If entries(i).Outcome <> "" Then logged = logged + 1
    Next i
    If logged = 0 Then Exit Function

    Dim ps As Date: ps = PeriodStartFor(entryDate)
    If Not ConfirmPeriod(ps) Then Exit Function

    sessionID = NextRefID("S")
    Dim user As String: user = CurrentUser()
    Dim submTime As Date: submTime = Now

    Dim data() As Variant: ReDim data(1 To logged, 1 To L_NCOLS)
    Dim charged As Double, totalCharged As Double
    Dim nAcc As Long, nRef As Long, nMsg As Long, nUna As Long, accTxt As String
    ' v24.1: the hours each person was offered (less than the call-out's below a partial acceptance)
    Dim oc() As String, offers() As Double, worked As Double
    ReDim oc(1 To n)
    For i = 1 To n
        oc(i) = entries(i).Outcome
    Next i
    CalloutOffers hours, oc, taken, n, offers

    For i = 1 To n
        If entries(i).Outcome <> "" Then
            k = k + 1
            charged = CalloutCharge(entries(i).Outcome, entries(i).Grade, offers(i))   ' v21: one rule; v24.1: on the hours offered
            totalCharged = totalCharged + charged

            data(k, L_REF) = sessionID
            data(k, L_TYPE) = "Callout"
            data(k, L_DATE) = entryDate
            data(k, L_DEPT) = dept
            data(k, L_SHIFTLBL) = shiftLbl
            data(k, L_CREW) = crewSel
            data(k, L_BASEHRS) = offers(i)                 ' the hours this person was offered
            If entries(i).Outcome = "Accepted" Then data(k, L_BASEHRS) = AcceptedHours(offers(i), taken(i))   ' provisional OT sums this
            data(k, L_ORDER) = i          ' position in the call list
            data(k, L_EMP) = entries(i).EmpName
            data(k, L_CODE) = entries(i).DeptCode
            data(k, L_EMPHRS) = entries(i).TotalHrs
            If entries(i).EstHrs Then data(k, L_EMPHRS) = entries(i).RosterHrs     ' v24.3: the log keeps their real total, never the estimate
            data(k, L_OUTCOME) = entries(i).Outcome
            data(k, L_REASON) = entries(i).Reason
            If entries(i).Outcome = "Accepted" And taken(i) > 0 And taken(i) < offers(i) Then
                data(k, L_REASON) = "Took " & HoursText(taken(i)) & " of " & HoursText(offers(i)) & " h (the rest dropped, uncharged)"
            ElseIf offers(i) = 0 And entries(i).Reason = "" Then
                data(k, L_REASON) = "Not offered: a partial acceptance closed the shift"      ' v25 (D-E4): never charged
            ElseIf offers(i) < hours And entries(i).Reason = "" Then
                data(k, L_REASON) = "Offered " & HoursText(offers(i)) & " of " & HoursText(hours) & " h (part already taken)"
            End If
            data(k, L_CHARGED) = charged
            data(k, L_PERIOD) = ps
            data(k, L_USER) = user
            data(k, L_STAMP) = submTime

            Select Case entries(i).Outcome
                Case "Accepted": nAcc = nAcc + 1
                Case "Refused": nRef = nRef + 1
                Case "Left Message": nMsg = nMsg + 1
                Case "Unavailable": nUna = nUna + 1
            End Select
        End If
    Next i

    If Not AppendLogRows(data) Then Exit Function     ' the caller keeps its entries so nothing is lost
    Dim savedOk As Boolean: savedOk = SaveIfWritable()     ' v26.5.5 (A1): a failed save is in the summary; the rows ARE in the log

    ' accepted -> onto the calendar
    Dim calTxt As String, accShort As String
    For i = 1 To n
        If entries(i).Outcome = "Accepted" Then
            worked = AcceptedHours(offers(i), taken(i))
            accTxt = accTxt & vbNewLine & "      " & entries(i).EmpName & ": " & HoursText(worked) & " h"
            accShort = accShort & IIf(accShort = "", "", ", ") & entries(i).EmpName & " " & HoursText(worked) & " h"
            calTxt = calTxt & vbNewLine & "  " & CalendarAddAccepted(dept, entryDate, shiftLbl, entries(i).EmpName, worked, sessionID)
        End If
    Next i

    ' v25.5: the four lines the pop-up shows (the detail below stays on the status bar and in the OT Log)
    shortSummary = "Session " & sessionID & " submitted: " & nAcc & " accepted" & IIf(accShort <> "", " (" & accShort & ")", "") & _
                   ", " & nRef & " refused, " & nMsg & " left message, " & nUna & " unavailable." & vbNewLine & _
                   "Charged " & HoursText(totalCharged) & " h (Refused and Left Message, grade " & CHARGE_MAX_GRADE & " and below, x " & CHARGE_MULTIPLIER & ")." & vbNewLine & _
                   dept & "  |  " & shiftLbl & "  |  crew " & crewSel & "  |  " & Format$(entryDate, "ddd mmm d") & "  |  " & HoursText(hours) & " h" & _
                   IIf(calTxt <> "", vbNewLine & "Calendar:" & Replace(calTxt, vbNewLine & "  ", " "), "")

    summary = "Session " & sessionID & " submitted!" & vbNewLine & vbNewLine & _
              dept & " | " & shiftLbl & " | Crew: " & crewSel & " | " & Format$(entryDate, "ddd mmm d") & vbNewLine & _
              "Base hours: " & hours & "   Rate: " & CHARGE_MULTIPLIER & "x" & vbNewLine & vbNewLine & _
              "  Accepted: " & nAcc & accTxt & vbNewLine & _
              "  Refused: " & nRef & vbNewLine & _
              "  Left Message: " & nMsg & vbNewLine & _
              "  Unavailable: " & nUna & vbNewLine & vbNewLine & _
              "  Total charged: " & HoursText(totalCharged) & " hrs" & _
              IIf(calTxt <> "", vbNewLine & vbNewLine & "Calendar:" & calTxt, "")
    ' v26.5.5 (B13): what to do next, from the outcome (a partial acceptance closed the shift; nobody = the line stays open)
    Dim partHrs As Double
    For i = 1 To n
        If entries(i).Outcome = "Accepted" And taken(i) > 0 Then partHrs = taken(i)
    Next i
    Dim nextTxt As String: nextTxt = NextStepText(nAcc, partHrs, nMsg, IsAdmin(), PRIMARY_ADMIN_NAME)
    shortSummary = shortSummary & vbNewLine & nextTxt
    summary = summary & vbNewLine & vbNewLine & nextTxt
    If Not savedOk Then                                ' v26.5.5 (A1): logged in the workbook, NOT on disk
        shortSummary = "NOT SAVED TO DISK - try File > Save now. These entries are logged in the workbook; do not log them again." & _
                       vbNewLine & shortSummary
        summary = "NOT SAVED TO DISK." & vbNewLine & summary
    End If
    SubmitCalloutSession = True
End Function

' v26.5.5 (B13): the last line of a submitted session - what happens to the line and what to do next.
' PURE (SelfTest "nextstep"; mirrored in tests/rules.py). partialHours > 0 = someone took part of the hours
' (the shift closed, v25 D-E4); nAccepted = 0 = the line stays open.
Public Function NextStepText(ByVal nAccepted As Long, ByVal partialHours As Double, ByVal nLeftMsg As Long, _
                             ByVal isAdminUser As Boolean, ByVal adminName As String) As String
    If nAccepted = 0 Then
        NextStepText = "The line stays OPEN on the calendar. Next: " & _
                       IIf(nLeftMsg > 0, "call back later (Fill a shift now shows ? beside them), ", "") & _
                       "post the shift, or " & IIf(isAdminUser, "award out of department.", "ask " & adminName & " to award out of department.")
    ElseIf partialHours > 0 Then
        NextStepText = "The shift is closed at " & HoursText(partialHours) & " h; the rest is dropped, nobody charged."
    Else
        NextStepText = "The line is covered on the calendar."
    End If
End Function


' ====================================================================
'  ACCESS / PROTECTION
' ====================================================================
Public Function CurrentUser() As String
    CurrentUser = Environ$("USERNAME")
End Function

Public Function IsAdmin() As Boolean
    Dim u As String: u = LCase$(Trim$(CurrentUser()))
    If u = "" Then Exit Function
    IsAdmin = (InStr(1, "," & LCase$(Replace(ADMIN_USERS, " ", "")) & ",", _
                     "," & u & ",", vbBinaryCompare) > 0)
End Function

Public Function RequireAdmin() As Boolean
    If IsAdmin() Then
        RequireAdmin = True
    Else
        MsgBox "Access denied. Your Windows login (" & CurrentUser() & _
               ") is not on the admin list for this workbook.", vbCritical, "Authorization"
    End If
End Function

' v25.3 (Graeme, Oct 17, 2026): SUPERVISOR rights - the jobs every supervisor with the file open may do: log a
' call-out, sick call (books the sick shift, can start the call-out), new posting, award posting, print the
' call-out roster, and (v25.5) publish the members' file. Everything else (time off, switches, leave, void, adjustments, people, Kronos, the
' reopen / no-takers / out-of-department / cancel of a posting - the three Submit* carry RequireAdmin before
' CanFinalize) stays behind RequireAdmin. v25.5 (Graeme, Oct 18): a supervisor is a login on SUPERVISOR_USERS or
' an admin (ViewModeFor, modDesk), with edit access; anyone else is refused by name, a read-only copy as before.
Public Function IsSupervisor() As Boolean
    IsSupervisor = (ViewModeFor(CurrentUser(), ADMIN_USERS, DESK_USERS, SUPERVISOR_USERS) <> "user")
End Function

Public Function RequireSupervisor() As Boolean
    If Not IsSupervisor() Then
        MsgBox "Access denied. Your Windows login (" & CurrentUser() & ") is not on the supervisor list for this workbook." & _
               vbNewLine & vbNewLine & "Ask " & PRIMARY_ADMIN_NAME & " to add your login to the supervisors' list.", vbCritical, "Authorization"
        Exit Function
    End If
    RequireSupervisor = RequireWritable()
End Function

Public Function RequireWritable() As Boolean
    If ThisWorkbook.ReadOnly Then
        MsgBox "This copy is open read-only, so nothing can be logged or saved." & vbNewLine & _
               "Close it and reopen with edit access.", vbExclamation, "Read-only"
    Else
        RequireWritable = True
    End If
End Function

' v24.2: a confirm. SmokeTest answers it itself (smokeAnswer); everyone else sees the box.
Public Function Ask(ByVal prompt As String, ByVal buttons As Long, ByVal title As String, ByVal smokeAnswer As Long) As Long
    If gSmokeRun Then
        Ask = smokeAnswer
    Else
        Ask = MsgBox(prompt, buttons, title)
    End If
End Function

' v24.2: a message that only reports success. Quiet under SmokeTest.
Public Sub Tell(ByVal prompt As String, ByVal buttons As Long, ByVal title As String)
    If Not gSmokeRun Then MsgBox prompt, buttons, title
End Sub

' A failed conversion must never be saved. The flag is also kept as a
' hidden workbook name so it survives a VBA reset ("End" on an error).
Public Sub MarkNoSave()
    gNoSaveOnClose = True
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="NoSaveFlag", RefersTo:="=TRUE", Visible:=False
    On Error GoTo 0
End Sub

Public Function NoSave() As Boolean
    If gNoSaveOnClose Then NoSave = True: Exit Function
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names("NoSaveFlag")
    On Error GoTo 0
    NoSave = Not nm Is Nothing
End Function

' Lifts the save block - only called after a full check against the OT Log passes
Public Sub ClearNoSave()
    gNoSaveOnClose = False
    On Error Resume Next
    ThisWorkbook.Names("NoSaveFlag").Delete
    On Error GoTo 0
End Sub

' Saves the workbook after a write. True = saved (or nothing to save: a read-only copy, SmokeTest).
' v26.5.5 (A1): a save that FAILS is reported, never swallowed. Until now the Save sat under
' On Error Resume Next: a locked file or a dropped share left the pop-up saying "submitted!"
' with nothing on disk. Excel's Save raises on most failures; when Workbook_BeforeSave refuses
' it (a stale copy, item A2) it comes back quietly with Saved still False, so both are tested.
' The callers that log something keep what they logged on screen when this returns False.
Public Function SaveIfWritable() As Boolean
    If gSmokeRun Then SaveIfWritable = True: Exit Function          ' v24.2: SmokeTest never saves
    If NoSave() Then
        MsgBox "Not saved - saving is blocked because an earlier check failed in this session." & vbNewLine & _
               "Close WITHOUT saving and reopen.", vbExclamation, "Not saved"
        Exit Function
    End If
    If ThisWorkbook.ReadOnly Then SaveIfWritable = True: Exit Function
    Dim en As Long, em As String, before As Date, after As Date
    On Error Resume Next
    before = FileDateTime(ThisWorkbook.FullName)
    Err.Clear
    ThisWorkbook.Save
    en = Err.Number: em = Err.Description
    after = FileDateTime(ThisWorkbook.FullName)
    On Error GoTo 0
    If en = 0 And ThisWorkbook.Saved Then NoteSaveResult True: SaveIfWritable = True: Exit Function
    If NoSave() Then Exit Function                                  ' BeforeSave refused it and said why (a stale copy)
    If en = 0 And before > 0 And after > before Then                ' written; something dirtied the workbook again after the save
        NoteSaveResult True: SaveIfWritable = True: Exit Function
    End If
    NoteSaveResult False
    If em = "" Then em = "the file on disk did not change"
    LogError "SaveIfWritable", en, em
    MsgBox "NOT SAVED to disk (" & em & ")." & vbNewLine & vbNewLine & _
           "Don't close Excel. Try File > Save now; if that fails, tell " & PRIMARY_ADMIN_NAME & "." & vbNewLine & _
           "What you logged is in this workbook until it is saved.", vbCritical, "Not saved"
End Function

' ---- v26.5.5 (A2): a stale copy blocks the save ----
' Two supervisors can open the roster minutes apart (the second gets it writable when the first has
' closed, or through a dropped lock). Whoever saves second would overwrite the first one's entries.
' Workbook_Open records the file's on-disk time; every own save records it again (Workbook_AfterSave);
' Workbook_BeforeSave compares. A changed time means someone else saved meanwhile: the save is refused,
' saving is blocked for the session (MarkNoSave) and the box says to close WITHOUT saving and reopen.
' A read-only copy, SmokeTest, and a file whose time can't be read (an http path) are not guarded.
Public Sub RecordDiskStamp()
    gOpenedMtime = 0
    On Error Resume Next
    If ThisWorkbook.Path <> "" And LCase$(Left$(ThisWorkbook.Path, 4)) <> "http" Then gOpenedMtime = FileDateTime(ThisWorkbook.FullName)
    On Error GoTo 0
End Sub

' True when the file on disk was saved by someone else since this session opened or last saved it
Public Function StaleOnDisk(ByRef savedAt As Date) As Boolean
    savedAt = 0
    If gOpenedMtime = 0 Or gSmokeRun Or ThisWorkbook.ReadOnly Then Exit Function
    On Error Resume Next
    savedAt = FileDateTime(ThisWorkbook.FullName)
    If Err.Number <> 0 Then savedAt = 0
    On Error GoTo 0
    If savedAt = 0 Then Exit Function                   ' the file is unreachable right now: the save itself will say so (A1)
    StaleOnDisk = (Abs(CDbl(savedAt) - CDbl(gOpenedMtime)) > 2 / 86400#)      ' two seconds: a share's time rounding
End Function

' SystemCheck's line
Public Function DiskStampText() As String
    If gOpenedMtime = 0 Then
        DiskStampText = "Stale-copy guard: off (the file's time can't be read here)"
    Else
        DiskStampText = "Stale-copy guard: on (file last saved " & Format$(gOpenedMtime, "ddd mmm d h:mm:ss AM/PM") & ")"
    End If
End Function

' v26.5.5 (A1): every save's result lands here (SaveIfWritable, and Workbook_AfterSave for a File > Save by hand),
' so the Desk header can say "Not saved since 3:14 AM" until a save goes through
Public Sub NoteSaveResult(ByVal ok As Boolean)
    If ok Then
        gSaveFailedAt = 0
    ElseIf gSaveFailedAt = 0 Then
        gSaveFailedAt = Now
    End If
End Sub

' The warning the pop-ups add to their success line after a failed save ("" when the save went through)
Public Function SaveWarnText() As String
    If gSaveFailedAt > 0 Then SaveWarnText = "   NOT SAVED to disk - File > Save now."
End Function

' UserInterfaceOnly lets macros edit while users can't. It is not saved
' with the file, so Workbook_Open re-applies it every session.
' v25 (R4, decided Oct 13, 2026): every sheet stays protected in EVERY view. What an admin
' may type into is the set of cells ApplyView leaves UNLOCKED (roster B, E, F, J, K, L:S and
' the Holidays "Observed on" column); everything else, formulas and tables included, goes
' through the buttons. Before this the admin view unprotected every sheet but the OT Log.
' Code writes through the UnlockSheet / ProtectSheet pair (both idempotent).
Public Sub ProtectSheet(ByVal ws As Worksheet)
    ' v24.4: the Install Log is written by the Roster Updater from outside this workbook, which has no
    ' password: it is never protected (admin-only through ApplyView, like the Error Log)
    If ws.Name = IL_SHEET Then UnlockSheet ws: Exit Sub
    ' v25 review: AllowSorting stays OFF - Excel lets a user sort a block made only of unlocked cells, so an
    ' admin could reorder seniority or grade codes against the names; the sort buttons sort in code
    ws.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, UserInterfaceOnly:=True, _
               AllowSorting:=False, AllowFiltering:=True
    If ws.Name = LOG_SHEET Then ws.EnableOutlining = True     ' +/- on the audit columns
End Sub

' v25: the one way code opens a sheet to write to it (ProtectSheet closes it again). The only
' Unprotect in the codebase: tests/test_drift.py checks that.
Public Sub UnlockSheet(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

' v25 (R4): the cells an admin types into, unlocked; every other roster cell locked. Non-admins
' get nothing unlocked. Idempotent, so ApplyView does it every time (no one-time migration).
' Worked (H) is unlocked too (v25 review): today the hours are typed from the Kronos PDF by hand and
' only the Excel import writes H in code; lock it again when the daily import is running.
Public Sub LockRosterCells(ByVal rs As Worksheet, ByVal adminView As Boolean)
    ' v26.5 (Graeme, Oct 19): nothing on the roster is typed into any more, in any view. A person's details
    ' change through the Employee pop-up's Edit (a double-click on the row); Worked (H) comes from the
    ' Kronos import. (v25 R4 had left B, E, F, H, J, K and L:S open for an admin.) adminView is kept for
    ' the callers; every cell is locked either way.
    rs.Cells.Locked = True
End Sub

' v25: how many sheets are protected, for SystemCheck ("" = as expected; otherwise the odd ones out)
Public Function ProtectionStatus(ByRef info As String) As String
    Dim ws As Worksheet, nProt As Long, nAll As Long, odd As String
    For Each ws In ThisWorkbook.Worksheets
        nAll = nAll + 1
        If ws.ProtectContents Then
            nProt = nProt + 1
        ElseIf ws.Name <> IL_SHEET Then
            odd = odd & IIf(odd = "", "", ", ") & ws.Name
        End If
    Next ws
    Dim rs As Worksheet: Set rs = RosterSheet()
    Dim totalsLocked As Boolean
    If rs Is Nothing Then
        info = "Protection: " & nProt & " of " & nAll & " sheets protected; no roster sheet found"
        ProtectionStatus = "- the roster sheet wasn't found (ROSTER_SHEET)"
        Exit Function
    End If
    totalsLocked = rs.Cells(R_FIRST, C_TOTAL).Locked And rs.Cells(R_FIRST, C_PERIOD1).Locked
    info = "Protection: " & nProt & " of " & nAll & " sheets protected" & IIf(SheetExists(IL_SHEET), " (the Install Log open by design)", "") & _
           "; roster totals " & IIf(totalsLocked, "locked", "NOT locked") & "; tabs " & IIf(ThisWorkbook.ProtectStructure, "locked", "NOT locked")
    If odd <> "" Then ProtectionStatus = "- unprotected sheet(s): " & odd & " (reopen the roster; ApplyView protects every sheet)"
    If Not totalsLocked Then ProtectionStatus = ProtectionStatus & IIf(ProtectionStatus = "", "", vbNewLine) & "- the roster's totals are not locked: reopen the roster"
    If Not ThisWorkbook.ProtectStructure Then ProtectionStatus = ProtectionStatus & IIf(ProtectionStatus = "", "", vbNewLine) & "- the workbook's tabs are not protected: reopen the roster"
End Function

' v25 review: every sheet is protected, so a message that asks for a hand edit says how
' v26.5.5 (B10): the hint names the BUTTON that changes the sheet, never the VBA editor; a repair by hand (a
' stray row, a cell that should be empty) goes through Alt+F8 > UnlockForRepair, then ProtectAll.
Public Function ByHandHint(ByVal sheetName As String) As String
    Dim button As String
    Select Case sheetName
        Case ROSTER_SHEET: button = "Employee > Edit (double-click the person's row)"
        Case "Crew History": button = "Employee > Edit (a crew change is dated into the history)"
        Case "Absences": button = "Book Time Off / Sick Call"
        Case "Postings": button = "the posting's own buttons (Award, No Takers, Cancel Request, Reopen)"
        Case "Shift Switches": button = "Shift Switch"
        Case "Coverage", "Calendar": button = "the calendar line's menu"
        Case "Holidays": button = "the Observed on cells (the only typed-into ones there)"
    End Select
    ByHandHint = " (The " & sheetName & " sheet is protected: nothing on it is typed into. " & _
                 IIf(button <> "", "An admin changes it through " & button & ". ", "") & _
                 "For a repair by hand an admin runs Alt+F8 > UnlockForRepair, makes the change, then Alt+F8 > ProtectAll.)"
End Function

' v26.5.5 (B10): the admin's way to a hand repair without the VBA editor - unlocks ONE sheet (the active one
' by default) until ProtectAll, a save, or the next open protects everything again
Public Sub UnlockForRepair()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim nm As String
    nm = Trim$(InputBox("Which sheet do you need to change by hand? (Nothing on the roster is typed into normally: use the buttons " & _
                        "where one exists. This is for a repair IntegrityCheck or a message asked for.)", "Unlock for repair", ActiveSheet.Name))
    If nm = "" Then Exit Sub
    If Not SheetExists(nm) Then MsgBox "There is no sheet called " & nm & ".", vbExclamation, "Unlock for repair": Exit Sub
    If IsYearEndSheet(nm) Or nm = LOG_SHEET Then
        MsgBox "The " & nm & " sheet is never changed by hand: corrections are voids and new entries (Void Entry, Add Adjustment).", _
               vbExclamation, "Unlock for repair"
        Exit Sub
    End If
    UnlockSheet ThisWorkbook.Worksheets(nm)
    ThisWorkbook.Worksheets(nm).Activate
    MsgBox "The " & nm & " sheet is unlocked. Make the change, then run Alt+F8 > ProtectAll (a save or the next open also protects everything)." & _
           vbNewLine & vbNewLine & "IntegrityCheck afterwards tells you whether the roster still agrees with the OT Log.", vbInformation, "Unlock for repair"
    Exit Sub
Oops:
    ReportError "UnlockForRepair"
End Sub

' The archived logs ("OT Log 2026") and the Year End record: never unprotected, even for admins (v23)
Public Function IsYearEndSheet(ByVal nm As String) As Boolean
    IsYearEndSheet = (nm Like LOG_SHEET & " ####" Or nm = YE_SHEET)
End Function

Public Sub ProtectAll()
    Dim ws As Worksheet
    On Error Resume Next                            ' v25 re-review: one sheet that won't protect doesn't stop the rest (ReportError calls this)
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws
    Next ws
    On Error GoTo 0
    RelockStructure                                 ' v25: and the tabs themselves
End Sub

' v25 (R4 / the 2.7 design): the workbook's STRUCTURE is protected too - a tab can't be renamed,
' moved, deleted or hidden by hand (a renamed OT Log would break every SUMIFS). Code that adds,
' removes, renames or shows / hides a sheet goes through this pair (both idempotent, no password
' outside the roster: the updater asks the roster's own macros).
Public Sub UnlockStructure()
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=SHEET_PASSWORD
    On Error GoTo 0
End Sub

Public Sub RelockStructure()
    On Error Resume Next
    If Not ThisWorkbook.ProtectStructure Then ThisWorkbook.Protect Password:=SHEET_PASSWORD, Structure:=True
    On Error GoTo 0
End Sub

' v25: show / hide a sheet under structure protection
Public Sub SetSheetVisible(ByVal ws As Worksheet, ByVal state As Long)
    If ws.Visible = state Then Exit Sub
    Dim en As Long, em As String
    UnlockStructure
    On Error Resume Next
    ws.Visible = state
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    RelockStructure                                 ' v25 review: relocked whatever happened
    If en <> 0 Then Err.Raise en, , em
End Sub

' v25.1: a tab colour is a change to the workbook's structure, which v25 keeps protected, so it is set
' between the UnlockStructure / RelockStructure pair like a sheet add. Cosmetic: it never raises, it
' writes only when the colour differs, and the tabs are relocked whatever happened. Every Tab.Color
' write in the roster goes through here (the Calendar and Needs Attention accessors set theirs on
' every call).
Public Sub SetTabColor(ByVal ws As Worksheet, ByVal colour As Long)
    On Error Resume Next
    If CStr(ws.Tab.Color) = CStr(colour) Then Exit Sub
    UnlockStructure
    ws.Tab.Color = colour
    RelockStructure
End Sub

' Admin view shows the admin sheets and phone numbers; everyone else doesn't see them.
' v25: and LockRosterCells (v26.5: every roster cell locked, nothing is typed into); error frame (E4).
Public Sub ApplyView(ByVal adminView As Boolean, Optional ByVal deskMode As Boolean = False)
    On Error GoTo Oops
    ApplyViewCore adminView, deskMode
    Exit Sub
Oops:
    ReportError "ApplyView"
End Sub

' v25 review: the work itself, raising to its caller (SaveRestrictedCopy refuses the copy on an error;
' the updater's "did the new code run" probe sees a real error). Buttons and the open use ApplyView.
Public Sub ApplyViewCore(ByVal adminView As Boolean, Optional ByVal deskMode As Boolean = False)
    gAdminUnlocked = adminView
    ' v25.3: the Desk view = an admin (every right, the unlocked cells) who sees a supervisor's tabs plus the Desk
    ' and Needs Attention, no admin groups on the bar, Phone shown, Email / Payroll # hidden (modDesk)
    gDeskMode = adminView And deskMode
    If adminView Then EnsureInstallLog               ' v25: the Install panel writes it; made BEFORE the visibility loop so it shows this session
    Dim nm As Variant
    For Each nm In Array(POSTING_SHEET, AB_SHEET, CV_SHEET, NA_SHEET, HOL_SHEET, HH_SHEET, GL_SHEET, AT_SHEET, IC_SHEET, YE_SHEET, CH_SHEET, FM_SHEET, PM_SHEET, REFUSAL_SHEET, EL_SHEET, IL_SHEET, _
                         HISTORY_ADMIN_SHEET, MINSTAFF_SHEET, WP_SHEET)   ' absences hold leave types - admin only; IL_SHEET v24.4; v26: the FULL history (logins) and the staffing minimums; v26.5: the per-period Kronos hours
        If SheetExists(CStr(nm)) Then
            Dim showIt As Boolean
            showIt = adminView And Not gDeskMode
            If gDeskMode Then showIt = (CStr(nm) = NA_SHEET)          ' v25.3: a Desk user sees a supervisor's tabs + the Desk + Needs Attention
            SetSheetVisible ThisWorkbook.Worksheets(CStr(nm)), IIf(showIt, xlSheetVisible, xlSheetVeryHidden)
        End If
    Next nm
    ' v25.5: the supervisor's buttons (ui_sup_) and the Desk sheet go to admins and the logins on SUPERVISOR_USERS;
    ' a plain user sees neither. Decided by the login, whatever view is being applied (the close-time save by a
    ' supervisor keeps them; the next open applies the opener's own view)
    Dim supView As Boolean: supView = adminView Or IsSupervisor()
    If SheetExists(DESK_SHEET) Then SetSheetVisible ThisWorkbook.Worksheets(DESK_SHEET), IIf(adminView Or (SUPERVISOR_DESK And supView), xlSheetVisible, xlSheetVeryHidden)   ' v25.3: admins' and supervisors' Desk
    ' v26: the member-safe history report and the board are a supervisor's output sheets (hidden from a plain user)
    For Each nm In Array(HISTORY_SHEET, BOARD_SHEET)
        If SheetExists(CStr(nm)) Then SetSheetVisible ThisWorkbook.Worksheets(CStr(nm)), IIf(supView, xlSheetVisible, xlSheetVeryHidden)
    Next nm
    Dim rs As Worksheet: Set rs = RosterSheet()
    UnlockSheet rs
    LockRosterCells rs, adminView                   ' v25 (R4)
    If SheetExists(HOL_SHEET) Then LockHolidayCells adminView
    If SheetExists(MINSTAFF_SHEET) Then LockMinStaffCells adminView   ' v26: the Minimum / Note cells open for an admin
    Dim leftWas As Double: leftWas = rs.Cells(1, C_SHIFT).Left
    rs.Columns(C_PHONE).Hidden = Not adminView
    rs.Columns(C_EMAIL).Hidden = Not adminView Or gDeskMode          ' v25.3: a Desk user keeps Phone, not these two
    rs.Columns(C_PAYROLL).Hidden = Not adminView Or gDeskMode        ' v24.2: a personal identifier, like phone and email
    ' column A is LEFT of the command bar, and the bar's shapes don't move with cells: carry them along
    Dim dx As Double: dx = rs.Cells(1, C_SHIFT).Left - leftWas
    Dim shp As Shape
    For Each shp In rs.Shapes
        If Left$(shp.Name, 3) = "ui_" And Abs(dx) > 0.01 Then shp.Left = shp.Left + dx
        If Left$(shp.Name, 7) = "ui_adm_" Then shp.Visible = IIf(adminView And (Not gDeskMode Or shp.Name = "ui_adm_Desk"), msoTrue, msoFalse)   ' v25.3: the Desk hides the admin groups, keeps "< Desk"
        If Left$(shp.Name, 7) = "ui_sup_" Then shp.Visible = IIf(supView, msoTrue, msoFalse)   ' v25.5: the supervisor's jobs
    Next shp
    ' v25.5 (review): the FIND group sits after ADMIN; when the ADMIN group is hidden (a supervisor, the Desk view) it
    ' takes ADMIN's place, so it is on screen on a 1366-wide monitor instead of past an empty gap
    On Error Resume Next
    Dim fx As Double: fx = rs.Shapes("ui_adm_lbl_admin").Left
    If adminView And Not gDeskMode Then fx = fx + 187          ' 2 * wA + G + GG in BuildCommandBar
    rs.Shapes("ui_sup_lbl_find").Left = fx
    rs.Shapes("ui_sup_Find").Left = fx
    rs.Shapes("ui_sup_History").Left = fx              ' v26: the group's second button moves with it
    On Error GoTo 0
    If SheetExists(CALLOUT_SHEET) Then
        ThisWorkbook.Worksheets(CALLOUT_SHEET).Columns(CALLOUT_PHONE_COL).Hidden = Not adminView
    End If
    If SheetExists(CAL_SHEET) Then
        For Each shp In ThisWorkbook.Worksheets(CAL_SHEET).Shapes
            If Left$(shp.Name, 8) = "cal_adm_" Then shp.Visible = IIf(adminView, msoTrue, msoFalse)
            If Left$(shp.Name, 8) = "cal_sup_" Then shp.Visible = IIf(supView, msoTrue, msoFalse)   ' v26: the Board button
        Next shp
    End If
    ProtectAll      ' v25: protects every sheet in every view
    On Error Resume Next
    SetRosterStatus                 ' refresh pay period / log count (keeps the current view text)
    On Error GoTo 0
End Sub


' ====================================================================
'  OT LOG ACCESS
' ====================================================================
' True when the table has the current column layout (see L_ constants)
Public Function LogLayoutCurrent(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count <> L_NCOLS Then Exit Function
    LogLayoutCurrent = (lo.ListColumns(L_REF).Name = "Ref ID" And lo.ListColumns(L_ENTRY).Name = "Entry ID" _
                        And lo.ListColumns(L_CHARGED).Name = "Hours Charged")
End Function

Public Function LogTable() As ListObject
    Dim lo As ListObject
    If SheetExists(LOG_SHEET) Then
        On Error Resume Next
        Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing Then
        MsgBox "The OT Log table (" & LOG_TABLE & " on sheet '" & LOG_SHEET & _
               "') was not found." & vbNewLine & "If this workbook hasn't been migrated yet, " & _
               "run RunMigration first.", vbCritical, "OT Log missing"
    ElseIf Not LogLayoutCurrent(lo) Then
        ' never write new-layout rows into the old column order
        MsgBox "The OT Log is still in the old column layout." & vbNewLine & vbNewLine & _
               "Run Alt+F8 > UpgradeOTLogLayout once, then try again.", vbExclamation, "OT Log upgrade needed"
        Set lo = Nothing
    End If
    Set LogTable = lo
End Function

' Current (non-void) logged total for one employee + pay period
Public Function LogSumFor(ByVal empName As String, ByVal periodStart As Date) As Double
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    LogSumFor = Application.WorksheetFunction.SumIfs( _
        lo.ListColumns(L_CHARGED).DataBodyRange, _
        lo.ListColumns(L_EMP).DataBodyRange, empName, _
        lo.ListColumns(L_PERIOD).DataBodyRange, CDbl(periodStart), _
        lo.ListColumns(L_STATUS).DataBodyRange, "<>VOID")
End Function

' Most recent entries at the top (sheet must be unprotected)
Public Sub SortLogNewestFirst(ByVal lo As ListObject)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add Key:=lo.ListColumns(L_ENTRY).Range, SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Public Function NextEntryID() As Long
    Dim lo As ListObject: Set lo = LogTable()
    Dim mx As Double
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            mx = Application.WorksheetFunction.Max(lo.ListColumns(L_ENTRY).DataBodyRange)
        End If
    End If
    NextEntryID = NextNumber(CLng(mx), IdMark("ENTRY"))
End Function

' ---- ID high-water marks (v23) ----
' The year-end rollover moves last year's rows out of the live log, so the
' highest number ever used per prefix is kept in a hidden workbook name
' (IdMark_S, IdMark_P, IdMark_ADJ, IdMark_ENTRY ...). Marks only go up.
Public Function IdMark(ByVal prefix As String) As Long
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names("IdMark_" & UCase$(prefix))
    On Error GoTo 0
    If nmObj Is Nothing Then Exit Function
    IdMark = CLng(Val(Mid$(nmObj.RefersTo, 2)))
End Function

Public Sub SetIdMark(ByVal prefix As String, ByVal n As Long)
    If n <= IdMark(prefix) Then Exit Sub
    Dim nmObj As Name, ref As String: ref = "=" & CStr(n)
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names("IdMark_" & UCase$(prefix))
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:="IdMark_" & UCase$(prefix), RefersTo:=ref, Visible:=False
    Else
        nmObj.RefersTo = ref
        nmObj.Visible = False
    End If
    On Error GoTo 0
End Sub

' v26 (I8): a Ref ID prefix as the marks know it - letters only (S, P, ADJ, LV ...), at most 8
' (was Private in modRollover; the integrity line and the repair need it too)
Public Function IsPlainPrefix(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 8 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "[A-Z]" Then Exit Function
    Next i
    IsPlainPrefix = True
End Function

' v26 (I8): the highest number per plain prefix among these Ref IDs, added into d
' (prefix -> number; a higher number already there wins). refs = a 2-D column array
' (RangeTo2D) or a 1-D array of strings. The ONE parse of "S-0042": the rollover's
' marks, the integrity line and RepairIdMarks all use it.
Public Sub AddRefMaxima(ByRef refs As Variant, ByVal d As Object)
    If Not IsArray(refs) Then Exit Sub
    Dim i As Long, s As String, p As Long, pre As String, k As Long, twoD As Boolean
    On Error Resume Next
    twoD = (UBound(refs, 2) >= 1)
    On Error GoTo 0
    For i = LBound(refs, 1) To UBound(refs, 1)
        If twoD Then s = UCase$(Trim$(SafeStr(refs(i, 1)))) Else s = UCase$(Trim$(SafeStr(refs(i))))
        p = InStrRev(s, "-")
        If p > 1 And p < Len(s) Then
            Dim tail As String: tail = Mid$(s, p + 1)
            If Len(tail) <= 9 And tail Like String$(Len(tail), "#") Then      ' digits only, no overflow (review: IsNumeric takes "1E3")
                pre = Left$(s, p - 1)
                If IsPlainPrefix(pre) Then
                    k = CLng(Val(tail))
                    If Not d.Exists(pre) Then
                        d.Add pre, k
                    ElseIf k > d(pre) Then
                        d(pre) = k
                    End If
                End If
            End If
        End If
    Next i
End Sub

' PURE (SelfTest "idmark"): Ref IDs joined by | -> "LV=3|P=17|S=431" (prefixes A-Z sorted),
' "" when none parses. Exercises AddRefMaxima on plain values.
Public Function RefPrefixMaxText(ByVal refList As String) As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim refs As Variant: refs = Split(refList, "|")
    AddRefMaxima refs, d
    If d.Count = 0 Then Exit Function
    Dim keys() As String, n As Long, kk As Variant, i As Long, j As Long, t As String
    ReDim keys(1 To d.Count)
    For Each kk In d.Keys
        n = n + 1: keys(n) = CStr(kk)
    Next kk
    For i = 2 To n                                   ' insertion sort: a handful of prefixes
        t = keys(i): j = i - 1
        Do While j >= 1
            If keys(j) <= t Then Exit Do
            keys(j + 1) = keys(j): j = j - 1
        Loop
        keys(j + 1) = t
    Next i
    For i = 1 To n
        RefPrefixMaxText = RefPrefixMaxText & IIf(i = 1, "", "|") & keys(i) & "=" & d(keys(i))
    Next i
End Function

' PURE (SelfTest "nextid"): the next number after the highest in the log and the mark
Public Function NextNumber(ByVal maxInLog As Long, ByVal mark As Long) As Long
    If mark > maxInLog Then NextNumber = mark + 1 Else NextNumber = maxInLog + 1
End Function

' Next ID for a prefix: S-0001 (callout session), P-0001 (posting),
' ADJ-0001 (adjustment), OB-0001 (opening balance), MAN-0001 (manual cells)
Public Function NextRefID(ByVal prefix As String) As String
    Dim lo As ListObject: Set lo = LogTable()
    Dim mx As Long, i As Long, s As String, v As Variant, k As Long
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            v = RangeTo2D(lo.ListColumns(L_REF).DataBodyRange)
            For i = 1 To UBound(v, 1)
                s = UCase$(SafeStr(v(i, 1)))
                If Left$(s, Len(prefix) + 1) = UCase$(prefix) & "-" Then
                    k = CLng(Val(Mid$(s, Len(prefix) + 2)))
                    If k > mx Then mx = k
                End If
            Next i
        End If
    End If
    NextRefID = UCase$(prefix) & "-" & Format$(NextNumber(mx, IdMark(prefix)), "0000")
End Function

' Appends rows to the log. data = 2-D array (1..n, 1..L_NCOLS).
' Entry IDs are filled in here. Returns False (and says why) on failure.
Public Function AppendLogRows(data As Variant) As Boolean
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function

    Dim n As Long: n = UBound(data, 1)
    If n < 1 Then AppendLogRows = True: Exit Function

    Dim firstId As Long: firstId = NextEntryID()
    Dim r As Long
    For r = 1 To n
        data(r, L_ENTRY) = firstId + r - 1
    Next r

    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    ' a brand-new table has one empty placeholder row - write over it
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If

    Dim errMsg As String
    Application.EnableCancelKey = xlErrorHandler    ' v26.5.5 (A3): Esc lands in Fail, not between the resize and the values
    On Error GoTo Fail
    UnlockSheet ws        ' tables can't grow on a protected sheet
    If ws.FilterMode Then ws.ShowAllData              ' v26.5.5 (A14): a filtered log can't be resized onto hidden rows
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, L_NCOLS)
        .Value = data
        .Columns(L_DATE).NumberFormat = "mm/dd/yyyy"
        .Columns(L_PERIOD).NumberFormat = "mm/dd/yyyy"
        .Columns(L_STAMP).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    End With
    On Error Resume Next            ' rows are already written - a sort hiccup must not report failure
    SortLogNewestFirst lo
    On Error GoTo Fail
    ProtectSheet ws
    Application.Calculate
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    SetRosterStatus                 ' log count on the status line
    On Error GoTo 0
    AppendLogRows = True
    Exit Function

Fail:
    errMsg = Err.Description
    On Error Resume Next
    Application.EnableCancelKey = xlInterrupt
    ProtectSheet ws
    ' v26.5.5 (A14): the resize went through but the values or the sort did not: the rows ARE in the table
    ' (blank or half-filled). Submitting again would log the session twice, so the save is blocked instead.
    If lo.ListRows.Count > cnt Then
        MarkNoSave
        MsgBox "The OT Log could not be finished:" & vbNewLine & errMsg & vbNewLine & vbNewLine & _
               "Its rows were written but not completed, so saving is blocked. Close WITHOUT saving, reopen, " & _
               "and log this again.", vbCritical, "OT Log"
    Else
        MsgBox "Could not write to the OT Log:" & vbNewLine & errMsg & vbNewLine & vbNewLine & _
               "Nothing was saved for this entry.", vbCritical, "OT Log"
    End If
End Function


' ====================================================================
'  SMALL UTILITIES
' ====================================================================
Public Function SheetExists(ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = ThisWorkbook.Sheets(nm)
    On Error GoTo 0
    SheetExists = Not sh Is Nothing
End Function

Public Function SafeStr(ByVal x As Variant) As String
    If IsError(x) Or IsNull(x) Then
        SafeStr = ""
    Else
        SafeStr = CStr(x)
    End If
End Function

Public Function NumOr0(ByVal x As Variant) As Double
    If IsError(x) Then
        NumOr0 = 0
    ElseIf IsNumeric(x) Then
        NumOr0 = CDbl(x)
    Else
        NumOr0 = 0
    End If
End Function

' .Value of a single cell is a scalar - this always returns a 2-D array
Public Function RangeTo2D(ByVal rng As Range) As Variant
    Dim a() As Variant
    If rng.Cells.Count = 1 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = rng.Value
        RangeTo2D = a
    Else
        RangeTo2D = rng.Value
    End If
End Function

Public Function ColLetter(ByVal colNum As Long) As String
    ColLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, colNum).Address(True, False), "$")(0)
End Function

Public Function EnsureSlash(ByVal p As String) As String
    If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then
        EnsureSlash = p
    Else
        EnsureSlash = p & Application.PathSeparator
    End If
End Function

Public Function FolderExists(ByVal p As String) As Boolean
    If Trim$(p) = "" Then Exit Function
    On Error Resume Next
    FolderExists = (Len(Dir(p, vbDirectory)) > 0)
    On Error GoTo 0
End Function
