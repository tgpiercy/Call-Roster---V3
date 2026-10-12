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
Public Const CORE_VERSION As String = "Core v24.1 (Jan 5)"
Public Const ROSTER_SHEET As String = "OT Roster"     ' tab name; the code also finds it by its headers
Public Const LOG_SHEET As String = "OT Log"
Public Const LOG_TABLE As String = "tblOTLog"
Public Const CALLOUT_SHEET As String = "Blank Call Out Roster"
Public Const CALLOUT_PHONE_COL As Long = 5          ' phone column on the printed call out roster (hidden for non-admins) - v21
Public Const POSTING_SHEET As String = "Overtime Posting"
Public Const REFUSAL_SHEET As String = "Refusal Report"

Public Const SHEET_PASSWORD As String = "mytest"
' Windows usernames allowed to run admin macros (comma separated, any case)
Public Const ADMIN_USERS As String = "piercyg,lanthierr"

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

' {YEAR} is replaced by the roster year (v23), so reports follow the rollover: use ReportFolder()
Public Const REPORT_FOLDER As String = "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\{YEAR}\"
' Where Member View saves the members' file. "." = the folder the roster is in (for now: Oct 12, 2026).
' At the network roll-out this becomes the members' folder, e.g.
'   "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\"
Public Const PUBLISH_FOLDER As String = "."
Public Const PUBLISH_FILE As String = "Call Out Roster - Member View.xlsx"   ' the old values-only file (no longer written from M1)
Public Const PUBLISH_REASONS As Boolean = False   ' False = leave Reason/Note column out of the member view

' Members' workbook (M1): the template sits next to this workbook; the published copy goes to PUBLISH_FOLDER
Public Const MEMBER_TEMPLATE_NAME As String = "Call Out Roster - Members TEMPLATE.xlsm"
Public Const MEMBER_FILE_NAME As String = "Call Out Roster - Members.xlsm"
Public Const MEMBER_PASSWORD As String = "members"  ' sheet protection in the members' file (modMember uses the same)
Public Const MEMBERS_STALE_DAYS As Long = 1         ' Needs Attention reminds after this many days
Public Const MEMBER_NEXT_DAYS As Long = 14          ' Who's next horizon
Public Const PUB_NAME As String = "MembersPublishedAt"
' Electronic sign-ups (v24.1): the members' workbook drops one small file per sign-up into this
' folder and Award Posting reads them. ".\Signups" = a folder beside the members' file; every
' member needs WRITE access to it. Off until the union and management agree it counts like
' initialling the posted sheet.
Public Const SIGNUPS_ON As Boolean = False
Public Const SIGNUP_FOLDER As String = ".\Signups"
Public Const SIGNUP_STALE_DAYS As Long = 7          ' IntegrityCheck: inbox files this old for a closed posting
Public Const KRONOS_PATH As String = "P:\INDIVIDUAL FOLDERS\Union\Local 21-O-02 Production\Production Overtime Roster\Kronos - Raw Data Processing.xlsm"
Public Const KRONOS_SHEET As String = "Sorted Data Here"
Public Const KRONOS_STALE_DAYS As Long = 7            ' warn on call-outs / postings when Kronos data is older than this
Public Const PROV_NAME As String = "KronosDataThrough" ' hidden name: last date the Kronos data covers (v21)

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
Public Const C_REFUSED As Long = 9      ' I  = SUM(period columns)
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
End Type

' Session flags
Public gAdminUnlocked As Boolean    ' admin view: sheets left unprotected (except the OT Log)
Public gNoSaveOnClose As Boolean    ' set when a migration/verification fails - never save this state
' v24.2: SmokeTest (modSmokeTest) sets gSmokeRun on a practice copy only. While it is True
' nothing is saved to disk, confirms answer themselves (Ask) and success messages stay quiet (Tell).
Public gSmokeRun As Boolean
Public gSmokeFolder As String       ' where the publish step writes its scratch members' file
Public gSmokeSaved As String        ' ...and the file it wrote ("" = it didn't)


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

' Award Posting pop-up (admin)
Sub ShowPostingForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
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
                               "Not Qualified for Task", "Other", "Days")
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
    ' asOf (v24): a shift date. Each person's crew is then the crew they are on
    ' that day (crew history); left out, it is the roster's crew column.
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

    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, C_NAME)))
        code = Trim$(SafeStr(v(i, dc)))
        If nm <> "" And code <> "" Then
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
                        If provLive Then .ProvHrs = NumOr0(v(i, C_PROV)) Else .ProvHrs = 0
                        .Seniority = SeniorityKey(v(i, C_SENIORITY))
                        .Phone = Trim$(SafeStr(v(i, C_PHONE)))
                        .EmpShift = crewTxt
                        .Outcome = ""
                        .Reason = ""
                    End With
                End If
            End If
        End If
    Next i

    If n = 0 Then Exit Function
    ReDim Preserve recs(1 To n)
    SortEmpRecs recs, n
    LoadEligible = n
End Function

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

Private Function RecBefore(a As EmpRec, b As EmpRec) As Boolean
    If a.Grade <> b.Grade Then RecBefore = (a.Grade < b.Grade): Exit Function
    If a.TotalHrs <> b.TotalHrs Then RecBefore = (a.TotalHrs < b.TotalHrs): Exit Function
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
            If taken(i) > 0 And taken(i) < remain Then remain = remain - taken(i)
        End If
    Next i
End Sub

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

' Refused (column I) = the sum of the row's pay-period cells. Sheet must be unprotected.
Public Sub WriteRefusedRow(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, C_REFUSED).Formula = "=SUM(" & ColLetter(C_PERIOD1) & r & ":" & ColLetter(PeriodLastCol()) & r & ")"
    ws.Cells(r, C_REFUSED).Locked = True
End Sub

' Non-void hours charged in the OT Log for one employee, across the pay
' periods shown on this roster. Used to verify the roster after rewrites.
Public Function LogTotalOnRoster(ByVal empName As String) As Double
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
                                     ByRef summary As String) As Boolean
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
            data(k, L_OUTCOME) = entries(i).Outcome
            data(k, L_REASON) = entries(i).Reason
            If entries(i).Outcome = "Accepted" And taken(i) > 0 And taken(i) < offers(i) Then
                data(k, L_REASON) = "Took " & HoursText(taken(i)) & " of " & HoursText(offers(i)) & " h"
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
    SaveIfWritable

    ' accepted -> onto the calendar
    Dim calTxt As String
    For i = 1 To n
        If entries(i).Outcome = "Accepted" Then
            worked = AcceptedHours(offers(i), taken(i))
            accTxt = accTxt & vbNewLine & "      " & entries(i).EmpName & ": " & HoursText(worked) & " h"
            calTxt = calTxt & vbNewLine & "  " & CalendarAddAccepted(dept, entryDate, shiftLbl, entries(i).EmpName, worked, sessionID)
        End If
    Next i

    summary = "Session " & sessionID & " submitted!" & vbNewLine & vbNewLine & _
              dept & " | " & shiftLbl & " | Crew: " & crewSel & " | " & Format$(entryDate, "MM/DD/YYYY") & vbNewLine & _
              "Base hours: " & hours & "   Rate: " & CHARGE_MULTIPLIER & "x" & vbNewLine & vbNewLine & _
              "  Accepted: " & nAcc & accTxt & vbNewLine & _
              "  Refused: " & nRef & vbNewLine & _
              "  Left Message: " & nMsg & vbNewLine & _
              "  Unavailable: " & nUna & vbNewLine & vbNewLine & _
              "  Total charged: " & HoursText(totalCharged) & " hrs" & _
              IIf(calTxt <> "", vbNewLine & vbNewLine & "Calendar:" & calTxt, "")
    SubmitCalloutSession = True
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

Public Sub SaveIfWritable()
    If gSmokeRun Then Exit Sub                     ' v24.2: SmokeTest never saves
    If NoSave() Then
        MsgBox "Not saved - saving is blocked because an earlier check failed in this session." & vbNewLine & _
               "Close WITHOUT saving and reopen.", vbExclamation
        Exit Sub
    End If
    If Not ThisWorkbook.ReadOnly Then
        On Error Resume Next
        ThisWorkbook.Save
        On Error GoTo 0
    End If
End Sub

' UserInterfaceOnly lets macros edit while users can't. It is not saved
' with the file, so Workbook_Open re-applies it every session.
' In admin view every sheet except the OT Log stays unprotected (as the old
' password login did), so admins can edit grades, shifts, postings etc.
Public Sub ProtectSheet(ByVal ws As Worksheet)
    If gAdminUnlocked And ws.Name <> LOG_SHEET And Not IsYearEndSheet(ws.Name) Then
        ws.Unprotect Password:=SHEET_PASSWORD
        Exit Sub
    End If
    ws.Protect Password:=SHEET_PASSWORD, DrawingObjects:=True, Contents:=True, _
               Scenarios:=True, UserInterfaceOnly:=True, _
               AllowSorting:=True, AllowFiltering:=True
    If ws.Name = LOG_SHEET Then ws.EnableOutlining = True     ' +/- on the audit columns
End Sub

' The archived logs ("OT Log 2026") and the Year End record: never unprotected, even for admins (v23)
Public Function IsYearEndSheet(ByVal nm As String) As Boolean
    IsYearEndSheet = (nm Like LOG_SHEET & " ####" Or nm = YE_SHEET)
End Function

Public Sub ProtectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProtectSheet ws
    Next ws
End Sub

' Admin view shows the admin sheets and phone numbers; everyone else doesn't see them
Public Sub ApplyView(ByVal adminView As Boolean)
    gAdminUnlocked = adminView
    Dim nm As Variant
    For Each nm In Array(POSTING_SHEET, AB_SHEET, CV_SHEET, NA_SHEET, HOL_SHEET, HH_SHEET, GL_SHEET, AT_SHEET, IC_SHEET, YE_SHEET, CH_SHEET, FM_SHEET, REFUSAL_SHEET, EL_SHEET)   ' absences hold leave types - admin only
        If SheetExists(CStr(nm)) Then
            If adminView Then
                ThisWorkbook.Worksheets(CStr(nm)).Visible = xlSheetVisible
            Else
                ThisWorkbook.Worksheets(CStr(nm)).Visible = xlSheetVeryHidden
            End If
        End If
    Next nm
    Dim rs As Worksheet: Set rs = RosterSheet()
    rs.Unprotect Password:=SHEET_PASSWORD
    rs.Columns(C_PHONE).Hidden = Not adminView
    rs.Columns(C_EMAIL).Hidden = Not adminView
    rs.Columns(C_PAYROLL).Hidden = Not adminView       ' v24.2: a personal identifier, like phone and email
    Dim shp As Shape
    For Each shp In rs.Shapes
        If Left$(shp.Name, 7) = "ui_adm_" Then shp.Visible = IIf(adminView, msoTrue, msoFalse)
    Next shp
    If SheetExists(CALLOUT_SHEET) Then
        ThisWorkbook.Worksheets(CALLOUT_SHEET).Columns(CALLOUT_PHONE_COL).Hidden = Not adminView
    End If
    If SheetExists(CAL_SHEET) Then
        For Each shp In ThisWorkbook.Worksheets(CAL_SHEET).Shapes
            If Left$(shp.Name, 8) = "cal_adm_" Then shp.Visible = IIf(adminView, msoTrue, msoFalse)
        Next shp
    End If
    ProtectAll      ' protects for everyone else; unlocks (except OT Log) for admins
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
    On Error GoTo Fail
    ws.Unprotect Password:=SHEET_PASSWORD        ' tables can't grow on a protected sheet
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
    SetRosterStatus                 ' log count on the status line
    On Error GoTo 0
    AppendLogRows = True
    Exit Function

Fail:
    errMsg = Err.Description
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Could not write to the OT Log:" & vbNewLine & errMsg & vbNewLine & vbNewLine & _
           "Nothing was saved for this entry.", vbCritical, "OT Log"
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
