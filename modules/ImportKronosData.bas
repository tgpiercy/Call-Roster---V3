Attribute VB_Name = "ImportKronosData"
Option Explicit

' ====================================================================
'  KRONOS -> Worked Hours (column H)
'  v27: reads the Kronos report's Totals block (see the v27 note below); the
'  processing workbook's "Sorted Data Here" sheet is the fall-back.
'  Names match ignoring spaces/case; a Kronos name with a trailing
'  middle initial (e.g. "Dobbin,Stephen A") also matches.
'  Matched names turn green; unmatched names are listed at the end.
'
'  v20 (Oct 9):
'   - Remembers when the import ran and the Kronos file's saved date, for
'     the status line and the stale-data warning (KronosStatusText).
'   - Every import adds a dated snapshot of each person's worked hours to
'     the hidden Hours History sheet (needed later for the leave return
'     top-up and provisional accepted OT).
'   - Also lists Kronos names that aren't on the roster.
'  v21 (Oct 11):
'   - Each snapshot row records the "Data Through" date: the Kronos data
'     runs through the day BEFORE the file was saved. The latest value is
'     kept as the hidden name KronosDataThrough, which the roster's
'     provisional-OT formulas (column C) read: accepted / awarded shifts
'     that have ended by that date are in Kronos, so they stop counting.
'  v26.5 (Oct 19):
'   - Read the Kronos EXPORT itself when KRONOS_PATH named it, or the folder
'     the exports are saved in; per roster pay period into the admin-only
'     Worked Periods sheet (WriteWorkedPeriods; published as dWorked).
'     The pivot sheet "Sorted Data Here" stays as the fall-back.
'  v27 (Oct 21, briefs/v27-kronos.md) - THE REPORT READ RIGHT:
'   - The export "Employee Transactions and Totals" has THREE blocks. The
'     dated "Transactions:" block is timecard EDITS (adds and reversals; it
'     nets to a few hundred hours a year) - v26.5 summed it and wrote wrong
'     Worked figures. The undated "Totals:" block (Global ID, Employee, Pay
'     Code, Hours; the heading SHIFTED two columns left) holds the real hours
'     per person per pay code for the report's Time Period. ReadKronosReport
'     reads the title's Time Period and Query and the two blocks by heading.
'   - Two Time Periods, one import: a ONE-DAY snapshot (every morning for
'     yesterday) adds that day's rows to the Worked Days LEDGER; the CUMULATIVE
'     report (Jan 11 .. yesterday, at each period end) replaces the Kronos
'     Totals SNAPSHOT and logs every difference from what the ledger expected
'     into Worked Changes (a timecard edit, dated by the Transactions block
'     when it names the day). Worked (H) = the snapshot through KronosCumThrough
'     + the ledger days after it, each code's hours x its multiplier
'     (PREMIUM_CODES / PREMIUM_MULTIPLIERS: straight-time equivalent, as the
'     processing workbook's Grand Total was), rounded whole per person
'     (WeightedTotal). Other codes on a roster person are reported, never
'     counted. PivotFromLedger gives that in the pivot's shape, so the
'     name / Payroll # matching below runs unchanged.
'   - The H write is verified: a backup copy first, the sanity confirm (anyone
'     whose H falls, or rises by more than KRONOS_JUMP_HOURS), H from one array,
'     Refused and the period cells compared before / after, MarkNoSave on a
'     mismatch. Data Through = the report's Time Period end, never the file's
'     saved date. Every report file newer than the last import is read, oldest
'     first (KronosFilesToImport). The pivot sheet is still the fall-back.
' ====================================================================

Public Const KR_VERSION As String = "Kronos import v29.0.3 (Oct 28)"
Public Const HH_SHEET As String = "Hours History"
Public Const HH_TABLE As String = "tblHoursHistory"
Private Const HH_NCOLS As Long = 8          ' v21: + Data Through; v23: + Year (the roster year the hours belong to)
' v26.5: premium hours per person per pay period, from the Kronos EXPORT's dated transactions
' (Graeme, Oct 19: the roster reads the export itself; the pivot workbook's "Sorted Data Here" stays
' as the fall-back). Rebuilt in full at every import of an export (it spans the year).
Public Const WP_SHEET As String = "Worked Periods"
Public Const WP_TABLE As String = "tblWorkedPeriods"
Private Const WP_NCOLS As Long = 6          ' Payroll #, Employee, Period Start, Hours, Codes, Imported At
Private mTxPeriods As Object                ' the per-period sums for WriteWorkedPeriods: "namekey|serial" -> weighted hours
Private mTxCodes As Object                  ' "namekey|serial" -> "code h; code h"
Private mTxRaw As Object                    ' namekey -> the name as the roster spells it
Private mTxPay As Object                    ' namekey -> Global ID (the Payroll #)
Private mTxSkipped As Long                  ' (kept for the message; 0 since v27)
Private mClashes As String                  ' v27 (review M3): report rows whose Payroll # the roster holds under another name
Private mClashNames As Object               ' v27 (review N1): the roster names a clash touches - H is left alone for them
Private mIngested As Long                   ' v27 (review N2): reports written into the tables in this run (set by the FIRST write)
Private mGapNote As String                  ' v27.5 (item 2): the last cumulative's gap Note ("" = no missed one-day report between the Throughs)
Private mGapMsg As String                   '   the same days as words for the message ("Oct 14, Oct 15")
Private Const GAP_NOTE_PREFIX As String = "no one-day report for "   ' v27.5 (item 2): a gap row's Note = this + "yyyy-mm-dd; ..."
Public Const KF_SHEET As String = "Kronos Files"
Public Const KF_TABLE As String = "tblKronosFiles"
Private Const KF_NCOLS As Long = 4          ' File, Result, Time Period, Imported At  (review N7: every file read, taken or skipped)
' v27: the three admin-only tables behind Worked (H) - see the header
Public Const KT_SHEET As String = "Kronos Totals"
Public Const KT_TABLE As String = "tblKronosTotals"
Private Const KT_NCOLS As Long = 6          ' Payroll #, Employee, Pay Code, Hours, Through, Imported At
Public Const WD_SHEET As String = "Worked Days"
Public Const WD_TABLE As String = "tblWorkedDays"
Private Const WD_NCOLS As Long = 8          ' Payroll #, Employee, Date, Pay Code, Hours, Weighted, Source, Imported At
Public Const WC_SHEET As String = "Worked Changes"
Public Const WC_TABLE As String = "tblWorkedChanges"
Private Const WC_NCOLS As Long = 13         ' When, Payroll #, Employee, Date, Pay Code, Hours Before, Hours After, Change, Weighted Change, Kind, Note, Source, Imported By
Public Const KRONOS_JUMP_HOURS As Double = 60   ' the sanity confirm names anyone whose H rises by more than this in one import (or falls at all)
Private Const KRONOS_YEAR_SLACK As Long = 14    ' a cumulative report must start within this many days of the roster year's first day (Kronos's year starts a period later)

' v27: one report, read into arrays (the Totals block is what counts; the Transactions block only dates an edit)
Private Type KronosReport
    HasTotals As Boolean
    FromD As Date
    ToD As Date
    Query As String
    nTot As Long
    TotId() As String
    TotName() As String
    TotCode() As String
    TotHrs() As Double
    nTx As Long
    TxId() As String
    TxDate() As Date
    TxCode() As String
    TxBy() As String
End Type

' ====================================================================
'  v27: THE PURE RULES (mirrored in tests/rules.py)
' ====================================================================
' The multiplier for a Kronos pay code (-1 = not a roster code): the code without its "CA " prefix, any case, trimmed,
' against the parallel lists (suite kmult).
Public Function PremiumMultiplier(ByVal code As String, ByVal codes As String, ByVal mults As String) As Double
    PremiumMultiplier = -1
    Dim c As String: c = Trim$(code)
    If UCase$(Left$(c, 3)) = "CA " Then c = Trim$(Mid$(c, 4))
    If c = "" Then Exit Function
    Dim cs As Variant: cs = Split(codes, ",")
    Dim ms As Variant: ms = Split(mults, ",")
    Dim i As Long
    For i = 0 To UBound(cs)
        If StrComp(Trim$(CStr(cs(i))), c, vbTextCompare) = 0 Then
            If i <= UBound(ms) Then PremiumMultiplier = Val(Trim$(CStr(ms(i))))
            Exit Function
        End If
    Next i
End Function

' Excel's ROUND(x, 0): half away from zero (VBA's Round is half-to-even; the processing workbook used Excel's)
Public Function RoundHalfUp(ByVal x As Double) As Double
    RoundHalfUp = Fix(x + 0.5000001 * Sgn(x))
End Function

' One person's Worked (H) from "code=hours;code=hours": the roster codes' hours x their multiplier, summed and rounded
' to a whole number; the other codes are left out and named after "|" so the caller reports them (suite kweight).
Public Function WeightedTotal(ByVal spec As String, ByVal codes As String, ByVal mults As String) As String
    Dim parts As Variant: parts = Split(spec, ";")
    Dim i As Long, p As Long, code As String, m As Double, tot As Double, other As String
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            p = InStr(CStr(parts(i)), "=")
            If p > 0 Then
                code = Trim$(Left$(CStr(parts(i)), p - 1))
                m = PremiumMultiplier(code, codes, mults)
                If m < 0 Then
                    other = other & IIf(other = "", "", ";") & code
                Else
                    tot = tot + Val(Mid$(CStr(parts(i)), p + 1)) * m
                End If
            End If
        End If
    Next i
    WeightedTotal = CStr(RoundHalfUp(tot)) & IIf(other <> "", "|" & other, "")
End Function

' The report's "Time Period : mm/dd/yyyy - mm/dd/yyyy" read by hand, never CDate: "" with fromD / toD set, or what is
' wrong ("no dates", "end before start"). Month first, as Kronos prints it (suite kperiod).
Public Function TimePeriodCore(ByVal text As String, ByRef fromD As Date, ByRef toD As Date) As String
    fromD = 0: toD = 0
    Dim s As String, i As Long, ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "/" Then s = s & ch Else s = s & " "
    Next i
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Dim toks As Variant: toks = Split(Trim$(s), " ")
    Dim n As Long, parts As Variant, m As Long, d As Long, y As Long, dt As Date
    For i = 0 To UBound(toks)
        parts = Split(CStr(toks(i)), "/")
        If UBound(parts) = 2 Then
            m = Val(parts(0)): d = Val(parts(1)): y = Val(parts(2))
            If y >= 1900 And y <= 2999 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                If d <= Day(DateSerial(y, m + 1, 0)) Then
                    dt = DateSerial(y, m, d)
                    n = n + 1
                    If n = 1 Then fromD = dt
                    If n = 2 Then toD = dt
                End If
            End If
        End If
    Next i
    If n < 2 Then fromD = 0: toD = 0: TimePeriodCore = "no dates": Exit Function
    If toD < fromD Then TimePeriodCore = "end before start"
End Function

' What a report is, from its Time Period against the roster year: "day" (one day: a snapshot for the ledger),
' "cumulative" (starts within KRONOS_YEAR_SLACK days of the year's first day: the year's totals), "range" (several
' days that are not the year to date: refused - the Totals block cannot be split into days), "outside" (not this
' roster year) (suite kkind).
Public Function ReportKindCore(ByVal fromD As Date, ByVal toD As Date, ByVal yearFrom As Date, ByVal yearTo As Date) As String
    If fromD = 0 Or toD = 0 Then ReportKindCore = "outside": Exit Function
    If fromD < yearFrom Or toD < yearFrom Or fromD > yearTo Then ReportKindCore = "outside": Exit Function   ' (review) a start before the year would bring last year's hours
    If Int(fromD) = Int(toD) Then ReportKindCore = "day": Exit Function
    If fromD <= yearFrom + KRONOS_YEAR_SLACK Then ReportKindCore = "cumulative" Else ReportKindCore = "range"
End Function

' v27.1: THE LEDGER RULE (PivotFromLedger, RebuildWorkedPeriods, WriteLedgerDay, RosterNameFor; mirrored in
' tests/rules_kronos.py, cases in tests/cases-kronos.csv)
' (review H3) a Kronos Totals snapshot from before this roster year is not this year's: 0 = none
Private Function LedgerCumOf(ByVal cumThrough As Date, ByVal yearFrom As Date) As Date
    If cumThrough < yearFrom Then LedgerCumOf = 0 Else LedgerCumOf = cumThrough
End Function

' The last day Worked (H) covers (suite kledger): the Totals' Through, then each next day the ledger holds, along the
' UNBROKEN run (review M7 / N6: a missed morning stops both the hours and the data-through date there). daySerials =
' the ledger's days as whole-day serials, ";"-joined (any order; only those after the Totals and in this year count).
' 0 = no snapshot this year.
Public Function LedgerRunCore(ByVal cumThrough As Date, ByVal yearFrom As Date, ByVal daySerials As String) As Date
    cumThrough = LedgerCumOf(cumThrough, yearFrom)
    LedgerRunCore = cumThrough
    Dim days As Object: Set days = CreateObject("Scripting.Dictionary")
    Dim parts As Variant: parts = Split(daySerials, ";")
    Dim i As Long, k As Long
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            k = CLng(Val(CStr(parts(i))))
            If CDbl(k) > CDbl(cumThrough) And CDbl(k) >= CDbl(yearFrom) Then days(k) = True
        End If
    Next i
    Dim d As Date
    If cumThrough > 0 Then
        d = cumThrough + 1
        Do While days.Exists(CLng(Int(CDbl(d))))
            LedgerRunCore = d
            d = d + 1
        Loop
    End If
End Function

' A ledger day counts toward Worked (H) when it is after the Totals' Through and inside the run (suite kledger)
Public Function LedgerDayIn(ByVal d As Date, ByVal cumThrough As Date, ByVal runEnd As Date) As Boolean
    LedgerDayIn = (d > cumThrough And d <= runEnd)
End Function

' One person's Worked (H) from the ledger: the snapshot's "code=h;code=h", then the run's days' "code=h;code=h", weighted
' and rounded ONCE (WeightedTotal; the codes not counted are dropped) (suite kledger)
Public Function LedgerPersonCore(ByVal snapSpec As String, ByVal daySpec As String) As Double
    Dim s As String: s = snapSpec
    If daySpec <> "" Then s = s & IIf(s = "", "", ";") & daySpec
    Dim w As String: w = WeightedTotal(s, PREMIUM_CODES, PREMIUM_MULTIPLIERS)
    Dim p As Long: p = InStr(w, "|")
    If p > 0 Then w = Left$(w, p - 1)
    LedgerPersonCore = Val(w)
End Function

' (review N4 / N5) a timecard edit (a Worked Changes "edit" row dated d) counts in Worked Periods when it is this year,
' from the first one-day report (ledgerFrom, 0 = none yet) on, and day d was not RE-imported at or after the edit was
' logged (reimportedAt = that day's newest "day" change row's When, Empty = never; editWhen = the edit row's When) -
' that re-import already holds the edited figure (suite kedit)
Public Function EditCountsCore(ByVal d As Date, ByVal yearFrom As Date, ByVal ledgerFrom As Date, _
                               ByVal reimportedAt As Variant, ByVal editWhen As Variant) As Boolean
    EditCountsCore = (d >= yearFrom And ledgerFrom > 0 And d >= ledgerFrom And Not DayReimportedAfter(reimportedAt, editWhen))
End Function

' (review N5 / R2) True when the day was RE-imported (reimportedAt, Empty = never) at or after the moment whenV (an edit
' row's When): that re-import already holds the edited figure. A day first imported after the edit is not a re-import.
Private Function DayReimportedAfter(ByVal reimportedAt As Variant, ByVal whenV As Variant) As Boolean
    If IsEmpty(reimportedAt) Then Exit Function
    If Not IsDate(whenV) Then Exit Function
    DayReimportedAfter = (CDbl(reimportedAt) >= CDbl(CDate(whenV)))
End Function

' (review M2 / M7) a one-day report for day d replaces a ledger row (its Date rowDate, its Employee emp) when the row is
' that day and is a roster person's (onRoster) or the quiet-day marker (a blank Employee); a Former person's row stays
' (suite kday)
Public Function LedgerDayDrop(ByVal rowDate As Variant, ByVal d As Date, ByVal emp As String, ByVal onRoster As Boolean) As Boolean
    If Not IsDate(rowDate) Then Exit Function
    If Int(CDbl(CDate(rowDate))) = Int(CDbl(d)) Then LedgerDayDrop = (emp = "" Or onRoster)
End Function

' v27.5 (item 2, Graeme Oct 25): the days in (oldThrough, toD] the ledger does not hold - a missed one-day report -
' as "yyyy-mm-dd; yyyy-mm-dd" ("" = none; oldThrough 0 = no snapshot yet, none). heldSerials = the ledger's days as
' whole-day serials, ";"-joined (the quiet-day marker's day counts as held) (suite kgapdays)
Public Function GapDaysText(ByVal oldThrough As Date, ByVal toD As Date, ByVal heldSerials As String) As String
    If oldThrough <= 0 Then Exit Function
    Dim held As Object: Set held = SerialSet(heldSerials)
    Dim d As Date: d = oldThrough + 1
    Do While Int(CDbl(d)) <= Int(CDbl(toD))
        If Not held.Exists(CLng(Int(CDbl(d)))) Then GapDaysText = GapDaysText & IIf(GapDaysText = "", "", "; ") & Format$(d, "yyyy-mm-dd")
        d = d + 1
    Loop
End Function

' v27.5 (item 2): the days a gap row's Note names (every "yyyy-mm-dd" in it), as ";"-joined whole-day serials ("" = none)
Public Function GapNoteDays(ByVal note As String) As String
    Dim i As Long, s As String
    i = 1
    Do While i <= Len(note) - 9
        s = Mid$(note, i, 10)
        If IsIsoDay(s) Then
            GapNoteDays = GapNoteDays & IIf(GapNoteDays = "", "", ";") & _
                          CStr(CLng(DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Right$(s, 2)))))
            i = i + 10
        Else
            i = i + 1
        End If
    Loop
End Function

' "yyyy-mm-dd" with digits in every place and a month 1-12, a day 1-31 (no CDate: the text is ours)
Private Function IsIsoDay(ByVal s As String) As Boolean
    If Len(s) <> 10 Then Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then Exit Function
    Dim i As Long
    For i = 1 To 10
        If i <> 5 And i <> 8 Then
            If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
        End If
    Next i
    Dim m As Long, dd As Long: m = CLng(Mid$(s, 6, 2)): dd = CLng(Right$(s, 2))
    IsIsoDay = (m >= 1 And m <= 12 And dd >= 1 And dd <= 31)
End Function

' v27.5 (item 2): a gap row (a year-to-date difference across missed one-day reports, dated d = that report's Through,
' its Note naming the days) counts in Worked Periods when it is this year, from the first one-day report on
' (as EditCountsCore), and the person has NO ledger row on any day it names (gapSerials; heldSerials = the days this
' person has rows on): once one of those days is imported, the day's own rows count and the gap row is set aside
' (suite kgap). A Note naming no day counts like an edit.
Public Function GapCountsCore(ByVal d As Date, ByVal yearFrom As Date, ByVal ledgerFrom As Date, _
                              ByVal gapSerials As String, ByVal heldSerials As String) As Boolean
    If Not (d >= yearFrom And ledgerFrom > 0 And d >= ledgerFrom) Then Exit Function
    Dim held As Object: Set held = SerialSet(heldSerials)
    Dim parts As Variant, i As Long
    parts = Split(gapSerials, ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            If held.Exists(CLng(Val(CStr(parts(i))))) Then Exit Function
        End If
    Next i
    GapCountsCore = True
End Function

' ";"-joined whole-day serials -> a set (a blank part skipped)
Private Function SerialSet(ByVal serials As String) As Object
    ' (review v27.5, HIGH) built in a local: inside a Function its own name followed by (...) is a CALL to itself, so
    ' "SerialSet(key) = True" recursed until the stack ran out - callcheck SELFCALL now flags that shape
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim parts As Variant, i As Long
    parts = Split(serials, ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then dict(CLng(Val(CStr(parts(i))))) = True
    Next i
    Set SerialSet = dict
End Function

' v27.5 (item 2): a gap Note's ISO days as words for a message - "Oct 14, Oct 15", at most five, then "and N more"
Private Function GapDaysWords(ByVal isoList As String) As String
    Dim parts As Variant, i As Long, n As Long
    parts = Split(GapNoteDays(isoList), ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            n = n + 1
            If n <= 5 Then GapDaysWords = GapDaysWords & IIf(GapDaysWords = "", "", ", ") & Format$(CDate(CLng(Val(CStr(parts(i))))), "mmm d")
        End If
    Next i
    If n > 5 Then GapDaysWords = GapDaysWords & " and " & (n - 5) & " more"
End Function

' a dictionary's keys ";"-joined
Private Function KeysJoined(ByVal dict As Object) As String
    Dim kk As Variant
    For Each kk In dict.Keys
        KeysJoined = KeysJoined & IIf(KeysJoined = "", "", ";") & CStr(kk)
    Next kk
End Function

' (review M3 / N1) a report row's roster person: byNum = the roster name its Payroll # gives, byName = the one its name
' gives ("" = none). "clash" = both found and they differ (nothing is taken for either), "number" = byNum, "name" =
' byName, "none" = not a roster person (suite kname)
Public Function RosterNameCore(ByVal byNum As String, ByVal byName As String) As String
    If byNum <> "" And byName <> "" And StrComp(byNum, byName, vbTextCompare) <> 0 Then
        RosterNameCore = "clash"
    ElseIf byNum <> "" Then
        RosterNameCore = "number"
    ElseIf byName <> "" Then
        RosterNameCore = "name"
    Else
        RosterNameCore = "none"
    End If
End Function

' v25: the by-heading rule on its own (D-I4): "" when 'Employee Name' and 'Grand Total' are both found
' on one heading row within the first 10 rows of src (a 2-D array of the sheet), else what is missing.
Public Function KronosHeadingProblem(ByRef src As Variant, ByRef nameCol As Long, ByRef totCol As Long, _
                                     ByRef hdrName As Long, ByRef hdrTot As Long) As String
    Dim missingHdr As String
    nameCol = HeaderColumn(src, hdrName, "Employee Name")
    totCol = HeaderColumn(src, hdrTot, "Grand Total")
    If nameCol = 0 Then missingHdr = "'Employee Name'"
    If totCol = 0 Then missingHdr = missingHdr & IIf(missingHdr = "", "", " and ") & "'Grand Total'"
    If missingHdr = "" And hdrName <> hdrTot Then missingHdr = "'Employee Name' and 'Grand Total' on the same heading row"
    KronosHeadingProblem = missingHdr
End Function

Sub SyncKronosData()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub

    Dim wsDest As Worksheet: Set wsDest = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then Exit Sub

    ' ---- v27: every report file newer than the last import, oldest first (one file: that file) ----
    Dim files As Variant: files = KronosFilesToImport()
    If Not IsArray(files) Then
        MsgBox "No Kronos file was found at" & vbNewLine & KRONOS_PATH & vbNewLine & vbNewLine & _
               "KRONOS_PATH (modOvertimeEntry) names the folder the Kronos reports are saved in (every report newer than the last import is read), or one report file.", _
               vbCritical, "Kronos import"
        Exit Sub
    End If
    Dim stamp As Date: stamp = Now
    ' v27 (A7): a backup copy before anything is written, and the hours that must NOT move captured now
    Dim bk As String
    If Not PeopleBackup("kronos", bk) Then Exit Sub
    Dim refBefore As Object: Set refBefore = RosterRefusedMap()

    ' ---- the roster's names and numbers, for matching the report's people (the roster's spelling wins) ----
    Dim nameMap As Object: Set nameMap = CreateObject("Scripting.Dictionary"): nameMap.CompareMode = 1   ' NameKey -> roster name
    Dim numMap As Object: Set numMap = CreateObject("Scripting.Dictionary"): numMap.CompareMode = 1    ' Payroll # -> roster name
    Dim r As Long, nm As String
    Dim pn As String
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(wsDest.Cells(r, C_NAME).Value))
        If nm <> "" Then
            nameMap(NameKey(nm)) = nm
            pn = PayrollText(wsDest.Cells(r, C_PAYROLL).Value)
            If pn <> "" Then numMap(pn) = nm
        End If
    Next r

    ' ---- each file: a report (its Totals block into the snapshot or the ledger) or the pivot sheet (the fall-back) ----
    Dim fileName As String, wbSource As Workbook, openedHere As Boolean
    Dim repTxt As String, nReports As Long, pivotSrc As Variant, pivotName As String
    Dim fileDate As Date, pivotDate As Date, awaitFrom As Date: awaitFrom = KronosAwaitFrom()
    Dim tablesTouched As Boolean                      ' (review H2 / M5): once a report is in the tables, an exit without the H write blocks the save
    mIngested = 0: mClashes = ""
    Set mClashNames = HeldClashNames()                   ' (review R1): the clashes earlier runs found, until a year-to-date report clears them
    Application.EnableCancelKey = xlErrorHandler      ' v26.5.5 (A3): Esc lands in Oops, never mid-write (review M5: before the table writes too)
    If Not KronosReadFiles(files, stamp, bk, awaitFrom, nameMap, numMap, wbSource, openedHere, _
                           repTxt, nReports, pivotSrc, pivotName, fileDate, pivotDate, tablesTouched) Then Exit Sub
    ThisWorkbook.Activate
    If mIngested > 0 Then SaveClashNames                  ' (review R1): remembered for the next run (a one-day report cannot see them)

    ' ---- what Worked (H) is read from: the snapshot + ledger once a year-to-date report has been taken (review H5:
    '      one-day reports alone are kept in the ledger but cannot give the year's hours), else the pivot sheet ----
    Dim src As Variant, fromReport As Boolean, through As Date
    If KronosCumThrough() >= RosterYearFrom() And KronosCumThrough() > 0 Then        ' (review L6) a snapshot from before this year is none
        src = PivotFromLedger(through)
        fromReport = True
        ' v29.0.3 (review 1.10): on the report path the "Kronos file date" (KronosDataDate: the Desk pill, the stale box,
        ' IntegrityCheck's Check line, the status line) is the data's own age - the Time Period end + 1 - not the report
        ' FILE's saved time: a report copied in keeps the time it was made, so from a few days on every open said stale
        If through > 0 Then fileDate = Int(through) + 1
        If fileDate > Int(stamp) Then fileDate = Int(stamp)      ' (re-review R2) a report ending TODAY read as "never imported" (a negative age)
        If Not IsArray(pivotSrc) Then pivotName = ""
    ElseIf IsArray(pivotSrc) Then
        src = pivotSrc
        fileDate = pivotDate
        If nReports > 0 Then repTxt = repTxt & vbNewLine & "The one-day reports are kept in the Worked Days ledger; Worked (H) comes from the pivot sheet until the year-to-date report (Jan 11 .. yesterday) has been imported."
    Else
        Application.EnableCancelKey = xlInterrupt
        If nReports > 0 Then
            ' the one-day rows in the ledger are right and harmless to save: no save block here
            MsgBox "The one-day report(s) were kept in the Worked Days ledger, but Worked (H) needs the YEAR-TO-DATE report first:" & repTxt & vbNewLine & vbNewLine & _
                   "Run the Kronos report for Jan 11 .. yesterday (Welland Hourly) into the Kronos folder and import again. Worked hours were not changed.", vbExclamation, "Kronos import"
        Else
            MsgBox "Nothing to import:" & repTxt & vbNewLine & vbNewLine & "Nothing was changed.", vbExclamation, "Kronos import"
        End If
        Exit Sub
    End If
    If fromReport And pivotName <> "" Then repTxt = repTxt & vbNewLine & pivotName & ": the pivot sheet is not used while reports are (Worked comes from the reports)"
    fileName = IIf(fromReport, "the Kronos reports", pivotName)

    ' ---- Kronos names -> grand total ----
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim rawOf As Object: Set rawOf = CreateObject("Scripting.Dictionary")   ' key -> name as Kronos spells it
    Dim used As Object: Set used = CreateObject("Scripting.Dictionary")    ' keys matched to a roster name
    Dim alt As Object: Set alt = CreateObject("Scripting.Dictionary")   ' "Last,First X" -> "last,first"
    Dim payOf As Object, keyOfNum As Object, pCol As Long        ' name key -> Payroll #; Payroll # -> name key; the Payroll # column
    If Not KronosNameIndex(src, fileName, bk, tablesTouched, dict, rawOf, alt, payOf, keyOfNum, pCol) Then Exit Sub
    ' ---- v23: after a rollover, Worked (H) stays 0 until the data covers the new year (the pivot path: by file date) ----
    If Not fromReport And awaitFrom > 0 Then
        Dim fdChk As Date: fdChk = fileDate
        If fdChk = 0 Then fdChk = Now
        If DataThroughFor(fdChk) < awaitFrom Then
            Application.EnableCancelKey = xlInterrupt
            MsgBox "This Kronos file was saved " & Format$(fdChk, "ddd mmm d, yyyy") & ", so it still holds last year's hours." & vbNewLine & vbNewLine & _
                   "Worked hours stay at 0 until you import a report that reaches " & Format$(awaitFrom, "mmm d, yyyy") & _
                   " (the first one of the new year). Worked hours were not changed.", vbExclamation, "Kronos import"
            Exit Sub
        End If
    End If

    ' ---- work out column H (nothing is written yet: the new figures go into an array) ----
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlErrorHandler      ' v26.5.5 (A3): Esc lands in Oops (ReportError restores the key), never mid-column
    Dim hit As String, found As Long, missing As String, kVal As Variant
    Dim rosterNum As String, numKey As String, nameClash As String, nFilled As Long, noNum As String, nNoNum As Long, payDiff As String, payShared As String
    Dim nRows As Long: nRows = lastRow - R_FIRST + 1
    Dim hNew As Variant: hNew = wsDest.Range(wsDest.Cells(R_FIRST, C_WORKED), wsDest.Cells(lastRow, C_WORKED)).Value2   ' the old H, row by row; matched rows are replaced
    Dim payNew As Variant: payNew = wsDest.Range(wsDest.Cells(R_FIRST, C_PAYROLL), wsDest.Cells(lastRow, C_PAYROLL)).Value2
    Dim greenRow() As Boolean: ReDim greenRow(1 To nRows)
    Dim fell As String, nFell As Long, jumped As String, nJumped As Long, oldH As Double, zeroed As String, nZeroed As Long
    Dim histN As Long, histName() As String, histHrs() As Double
    ReDim histName(1 To nRows): ReDim histHrs(1 To nRows)
    For r = R_FIRST To lastRow
        nm = Trim$(SafeStr(wsDest.Cells(r, C_NAME).Value))
        If nm <> "" Then
            hit = FindKey(dict, alt, NameKey(nm))
            ' v24.2: a row that has its Payroll # is matched by the number first
            rosterNum = "": numKey = ""
            If pCol > 0 Then
                rosterNum = PayrollText(wsDest.Cells(r, C_PAYROLL).Value)
                If rosterNum <> "" Then
                    If keyOfNum.Exists(rosterNum) Then numKey = CStr(keyOfNum(rosterNum))
                End If
            End If
            If fromReport And mClashNames.Exists(nm) Then
                hit = "~clash~"                    ' (review N1): a Payroll # clash in the report - H left alone, named below
            Else
                Select Case PayrollDecision(rosterNum, numKey <> "", (numKey <> "" And numKey = hit))
                    Case "number"
                        hit = numKey
                    Case "mismatch"                ' same number, another name: a rename in Kronos, or a typing slip
                        used(numKey) = True
                        nameClash = nameClash & "  " & nm & "   (Kronos has that Payroll # as " & rawOf(numKey) & ")" & vbNewLine
                        hit = "~leave~"
                End Select
            End If
            If hit = "~clash~" Then
                nameClash = nameClash & "  " & nm & "   (a Kronos row's Payroll # and name point at two roster people)" & vbNewLine
                If FindKey(dict, alt, NameKey(nm)) <> "" Then used(FindKey(dict, alt, NameKey(nm))) = True
            ElseIf hit = "~leave~" Then
                ' nothing is changed for this row
            ElseIf hit <> "" Then
                If pCol > 0 And payOf.Exists(hit) Then
                    If rosterNum = "" Then
                        If CStr(keyOfNum(CStr(payOf(hit)))) = "" Then
                            payShared = payShared & "  " & nm & vbNewLine          ' Kronos has that number under two names: don't guess
                        Else
                            payNew(r - R_FIRST + 1, 1) = CStr(payOf(hit))
                            nFilled = nFilled + 1
                        End If
                    ElseIf rosterNum <> CStr(payOf(hit)) Then
                        payDiff = payDiff & "  " & nm & vbNewLine                  ' matched by name; the roster's number isn't Kronos's
                    End If
                End If
                used(hit) = True
                kVal = dict(hit)
                If Not IsError(kVal) And IsNumeric(kVal) And Not IsEmpty(kVal) Then
                    oldH = NumOr0(hNew(r - R_FIRST + 1, 1))
                    hNew(r - R_FIRST + 1, 1) = CDbl(kVal)
                    greenRow(r - R_FIRST + 1) = True
                    found = found + 1
                    histN = histN + 1: histName(histN) = nm: histHrs(histN) = CDbl(kVal)
                    ' v27 (the sanity flags, 1.4): a fall, or a rise over KRONOS_JUMP_HOURS, is named before anything is written
                    If CDbl(kVal) < oldH - 0.001 Then
                        nFell = nFell + 1
                        If nFell <= 15 Then fell = fell & "  " & nm & ": " & HoursText(oldH) & " -> " & HoursText(CDbl(kVal)) & vbNewLine
                    ElseIf CDbl(kVal) > oldH + KRONOS_JUMP_HOURS Then
                        nJumped = nJumped + 1
                        If nJumped <= 15 Then jumped = jumped & "  " & nm & ": " & HoursText(oldH) & " -> " & HoursText(CDbl(kVal)) & vbNewLine
                    End If
                Else
                    missing = missing & "  " & nm & " (no number in Kronos)" & vbNewLine
                End If
            ElseIf fromReport And rosterNum <> "" Then
                ' (review M4 / N1) a roster row WITH a Payroll # that the year-to-date report has no row for: no premium hours
                ' this year, so H is 0 - never a figure left over from an earlier import. A row with no number is only "not found".
                oldH = NumOr0(hNew(r - R_FIRST + 1, 1))
                hNew(r - R_FIRST + 1, 1) = 0#
                histN = histN + 1: histName(histN) = nm: histHrs(histN) = 0#
                nZeroed = nZeroed + 1
                If nZeroed <= 10 Then zeroed = zeroed & "  " & nm & IIf(oldH > 0.001, " (was " & HoursText(oldH) & ")", "") & vbNewLine
                If oldH > 0.001 Then
                    nFell = nFell + 1
                    If nFell <= 15 Then fell = fell & "  " & nm & ": " & HoursText(oldH) & " -> 0 (not in the Kronos report)" & vbNewLine
                End If
            Else
                missing = missing & "  " & nm & IIf(fromReport, " (no Payroll # in column A: fill it, or match the Kronos spelling)", "") & vbNewLine
            End If
        End If
    Next r
    If pCol > 0 Then                               ' who still has no Payroll #
        For r = 1 To nRows
            If Trim$(SafeStr(wsDest.Cells(r + R_FIRST - 1, C_NAME).Value)) <> "" And PayrollText(payNew(r, 1)) = "" Then
                nNoNum = nNoNum + 1
                If nNoNum <= 10 Then noNum = noNum & "  " & Trim$(SafeStr(wsDest.Cells(r + R_FIRST - 1, C_NAME).Value)) & vbNewLine
            End If
        Next r
        If nNoNum > 10 Then noNum = noNum & "  ... and " & (nNoNum - 10) & " more" & vbNewLine
    End If

    ' ---- v27: the sanity confirm, then ONE write of column H, then the verify ----
    If nFell + nJumped > 0 Then
        Application.EnableCancelKey = xlInterrupt
        Application.ScreenUpdating = True
        Dim sane As String
        sane = "Before the Worked hours are written:" & vbNewLine
        If nFell > 0 Then sane = sane & vbNewLine & nFell & " people's Worked hours FALL (hours already imported should only grow):" & vbNewLine & fell & IIf(nFell > 15, "  ... and " & (nFell - 15) & " more" & vbNewLine, "")
        If nJumped > 0 Then sane = sane & vbNewLine & nJumped & " people's Worked hours rise by more than " & KRONOS_JUMP_HOURS & " h:" & vbNewLine & jumped & IIf(nJumped > 15, "  ... and " & (nJumped - 15) & " more" & vbNewLine, "")
        sane = sane & vbNewLine & "Expected after a wrong import was put right, after the LIEU weights were corrected (v27), or the first report of the year. " & _
               "Otherwise the report may be the wrong one (its Query and Time Period are in the box after this)." & vbNewLine & vbNewLine & _
               "Write these hours?" & IIf(tablesTouched, " (No = the Kronos tables this run already updated are undone by closing WITHOUT saving; the backup copy " & bk & " stays.)", _
                                           " (No = nothing changes; the backup copy " & bk & " stays.)")
        If Ask(sane, vbYesNo + vbExclamation + vbDefaultButton2, "Kronos import - check first", vbYes) <> vbYes Then
            If tablesTouched Then
                MarkNoSave                                   ' (review H2): the ledger / snapshot rows are in the workbook; only an unsaved close takes them back
                Tell "Worked hours were not written. The Kronos tables were updated by this run, so SAVING IS BLOCKED: close WITHOUT saving and reopen (the backup copy is " & bk & ").", vbExclamation, "Kronos import"
            Else
                Tell "Nothing was changed.", vbInformation, "Kronos import"
            End If
            Exit Sub
        End If
        Application.ScreenUpdating = False
        Application.EnableCancelKey = xlErrorHandler
    End If
    UnlockSheet wsDest
    If pCol > 0 Then                               ' the column keeps text, so a leading zero survives
        wsDest.Range(wsDest.Cells(R_FIRST, C_PAYROLL), wsDest.Cells(lastRow, C_PAYROLL)).NumberFormat = "@"
        If Trim$(SafeStr(wsDest.Cells(2, C_PAYROLL).Value)) = "" Then
            With wsDest.Cells(2, C_PAYROLL)
                .Value = "PAYROLL #"
                .Interior.Color = RGB(31, 56, 100): .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 8
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            End With
        End If
        wsDest.Range(wsDest.Cells(R_FIRST, C_PAYROLL), wsDest.Cells(lastRow, C_PAYROLL)).Value = payNew
    End If
    wsDest.Range(wsDest.Cells(R_FIRST, C_WORKED), wsDest.Cells(lastRow, C_WORKED)).Value = hNew      ' one assignment (A7)
    wsDest.Range(wsDest.Cells(R_FIRST, C_NAME), wsDest.Cells(lastRow, C_NAME)).Interior.ColorIndex = xlNone
    For r = 1 To nRows
        If greenRow(r) Then wsDest.Cells(r + R_FIRST - 1, C_NAME).Interior.Color = RGB(226, 239, 218)
    Next r
    wsDest.Rows(R_FIRST & ":" & lastRow).AutoFit
    ProtectSheet wsDest
    FillPayrollKeys True                            ' v28 (item 2): a number the import just filled reaches that person's rows
    Application.Calculate
    Dim moved As String: moved = TotalsDifferenceRef(refBefore, RosterRefusedMap())
    If moved <> "" Then
        MarkNoSave
        Application.EnableCancelKey = xlInterrupt
        Application.ScreenUpdating = True
        MsgBox "The Kronos import changed something it must not (Refused or a pay-period cell):" & vbNewLine & moved & vbNewLine & _
               "SAVING IS BLOCKED. Close WITHOUT saving and reopen; the backup copy is " & bk & ". Tell " & PRIMARY_ADMIN_NAME & ".", vbCritical, "Kronos import"
        Exit Sub
    End If

    ' ---- v20: remember when, and keep a dated snapshot ----
    If fileDate = 0 Then fileDate = stamp
    ' v26.5.5 (A14): a copied file can carry a saved time from years ago or from the future; the pivot path's data-through
    ' date reads it, so it is bounded to this roster's year (today otherwise, said so)
    If fileDate < RosterYearFrom() - 14 Or fileDate > Now + 1 Then
        If Not fromReport Then MsgBox "The Kronos file's saved time (" & Format$(fileDate, "ddd mmm d, yyyy h:mm AM/PM") & ") is outside this roster's year, " & _
               "so today is used as its date. If the export is older than that, its hours are old too.", vbExclamation, "Kronos import"
        fileDate = stamp
    End If
    SetDateName "KronosImportedAt", stamp
    SetDateName "KronosFileDate", fileDate
    ' v27: a report's data runs through its Time Period's end (PivotFromLedger gave it); the pivot sheet's through
    ' the day before its workbook was saved (v21)
    If Not fromReport Then through = DataThroughFor(fileDate)
    If through = 0 Then through = DataThroughFor(stamp)
    SetDateName PROV_NAME, through
    If awaitFrom > 0 Then ClearKronosAwait      ' v23: the new year's hours are in
    Dim histId As String
    If histN > 0 Then histId = AppendHoursHistory(stamp, fileDate, through, histName, histHrs, histN)
    Dim wpTxt As String                        ' v27: the per-period premium hours from the ledger (snapshot days only)
    If fromReport Then wpTxt = RebuildWorkedPeriods(nameMap, stamp)
    Application.CalculateFull                 ' provisional (C) and Total (G) pick up the new date
    wsDest.Activate                           ' before the screen comes back: the Worked Periods sheet must not flash
    Application.EnableCancelKey = xlInterrupt
    Application.ScreenUpdating = True
    On Error Resume Next
    SetRosterStatus
    On Error GoTo Oops
    LiteMark "kronos"                              ' v29: the Lite Desk's daily check - H is written

    KronosImportMessages dict, alt, used, rawOf, found, nZeroed, fromReport, fileDate, through, histId, repTxt, wpTxt, _
                         awaitFrom, pCol, nFilled, nameClash, missing, zeroed, payDiff, payShared, noNum, nNoNum
    Exit Sub
Oops:
    If tablesTouched Or mIngested > 0 Then MarkNoSave   ' (review M5 / N2): a report is in the tables but H is not written - an unsaved close is the way back
    ReportError "SyncKronosData"
    KronosCleanUp wsDest, wbSource, openedHere
End Sub

' SyncKronosData's file loop (v27.1: moved out of it unchanged): each report file read and taken into the tables, or
' its pivot sheet kept as the fall-back. False = a file would not open (the box is shown; SyncKronosData stops). No
' handler of its own: an error goes to SyncKronosData's Oops with wbSource / openedHere (ByRef) as they stand, so the
' file is closed there as before.
Private Function KronosReadFiles(ByRef files As Variant, ByVal stamp As Date, ByVal bk As String, ByVal awaitFrom As Date, _
                                 ByVal nameMap As Object, ByVal numMap As Object, ByRef wbSource As Workbook, ByRef openedHere As Boolean, _
                                 ByRef repTxt As String, ByRef nReports As Long, ByRef pivotSrc As Variant, ByRef pivotName As String, _
                                 ByRef fileDate As Date, ByRef pivotDate As Date, ByRef tablesTouched As Boolean) As Boolean
    Dim srcPath As String, fileName As String, wsSource As Worksheet
    Dim rep As KronosReport, fi As Long, ingest As String
    Dim lastFileDate As Date
    ' v27.5 (item 2, review): the one-day reports of a run are taken BEFORE its year-to-date, whatever the files' saved
    ' order, so a year-to-date saved a minute before that morning's one-day report sees the day and names no gap
    Dim deferred() As Boolean, rdPass As Long
    ReDim deferred(LBound(files) To UBound(files))
    For rdPass = 1 To 2
    For fi = LBound(files) To UBound(files)
        If rdPass = 2 And Not deferred(fi) Then GoTo NextFile
        srcPath = CStr(files(fi))
        fileName = Mid$(srcPath, InStrRev(srcPath, "\") + 1)
        Set wbSource = Nothing: openedHere = False
        On Error Resume Next
        Set wbSource = Workbooks(fileName)
        On Error GoTo 0                                ' an error from here raises to SyncKronosData's Oops
        If wbSource Is Nothing Then
            Application.EnableEvents = False          ' don't run the Kronos file's own macros
            On Error Resume Next
            Set wbSource = Workbooks.Open(Filename:=srcPath, ReadOnly:=True, UpdateLinks:=0)
            On Error GoTo 0
            Application.EnableEvents = True
            If wbSource Is Nothing Then
                Application.EnableCancelKey = xlInterrupt
                If tablesTouched Then MarkNoSave
                MsgBox "Could not open the Kronos file:" & vbNewLine & srcPath & vbNewLine & vbNewLine & _
                       IIf(tablesTouched, "An earlier file of this run was already taken into the Kronos tables, so SAVING IS BLOCKED: close WITHOUT saving, reopen, fix the file and import again. The backup copy is " & bk & ".", _
                           "Nothing was changed."), vbCritical, "Kronos import"
                Exit Function
            End If
            openedHere = True
        End If
        fileDate = 0
        On Error Resume Next
        fileDate = FileDateTime(srcPath)
        On Error GoTo 0
        If fileDate > lastFileDate Then lastFileDate = fileDate
        ingest = ReadKronosReport(wbSource, rep)
        If rdPass = 1 And rep.HasTotals And ingest = "" Then
            If ReportKindCore(rep.FromD, rep.ToD, RosterYearFrom(), RosterYearTo()) = "cumulative" Then
                deferred(fi) = True                       ' read again on pass 2, after every one-day report of the run
                If openedHere Then wbSource.Close SaveChanges:=False
                Set wbSource = Nothing: openedHere = False
                GoTo NextFile
            End If
        End If
        If rep.HasTotals And ingest = "" Then
            ' v23: after a rollover nothing is taken until a report reaches the new year
            If awaitFrom > 0 And rep.ToD < awaitFrom Then
                ingest = "its Time Period ends " & Format$(rep.ToD, "ddd mmm d, yyyy") & ": last year's hours (Worked stays 0 until a report reaches " & Format$(awaitFrom, "mmm d, yyyy") & ") - skipped"
            Else
                ingest = IngestReport(rep, fileName, stamp, nameMap, numMap)
                nReports = mIngested: tablesTouched = (mIngested > 0)       ' (review N2 / L2): set by the first table write, not by the text
            End If
            If Right$(ingest, 9) = "- skipped" Or InStr(ingest, "- skipped" & vbNewLine) > 0 Then RecordKronosFile fileName, "skipped: " & ingest, rep   ' (review N7): a skipped file is not reopened next time
            repTxt = repTxt & vbNewLine & fileName & ": " & ingest
        Else
            If rep.HasTotals Then repTxt = repTxt & vbNewLine & fileName & ": " & ingest   ' a 'Totals:' block the reader could not use (review M8): the pivot sheet is still tried
            ' the fall-back: the processing workbook's pivot sheet (Payroll #, Employee Name, Grand Total)
            Set wsSource = Nothing
            On Error Resume Next
            Set wsSource = wbSource.Worksheets(KRONOS_SHEET)
            On Error GoTo 0
            If wsSource Is Nothing Then
                If Not rep.HasTotals Then repTxt = repTxt & vbNewLine & fileName & ": not a Kronos report (no 'Totals:' block) and no sheet '" & KRONOS_SHEET & "' - skipped"
            Else
                pivotSrc = SheetValues(wsSource): pivotName = fileName: pivotDate = fileDate
            End If
        End If
        If openedHere Then wbSource.Close SaveChanges:=False
        Set wbSource = Nothing: openedHere = False
NextFile:
    Next fi
    Next rdPass
    If lastFileDate > 0 Then fileDate = lastFileDate           ' (re-review) the NEWEST file's time, not the deferred year-to-date's
    KronosReadFiles = True
End Function

' SyncKronosData's index of the Kronos side (v27.1: moved out of it unchanged): the heading rule, then each Kronos name
' -> its Grand Total (dict), its spelling (rawOf), the name without a middle initial (alt), its Payroll # (payOf,
' keyOfNum, made here; pCol = that column, 0 = none). False = a heading is missing (the box is shown; SyncKronosData
' stops). No handler of its own: an error goes to SyncKronosData's Oops.
Private Function KronosNameIndex(ByRef src As Variant, ByVal fileName As String, ByVal bk As String, ByVal tablesTouched As Boolean, _
                                 ByVal dict As Object, ByVal rawOf As Object, ByVal alt As Object, _
                                 ByRef payOf As Object, ByRef keyOfNum As Object, ByRef pCol As Long) As Boolean
    ' v24.3.1 (review I3, decided D-I4): the name and total columns are found by their headings, never by position.
    Dim i As Long, k As String, raw As String, a As String, pn As String
    Dim nameCol As Long, totCol As Long, hdrName As Long, hdrTot As Long, missingHdr As String
    missingHdr = KronosHeadingProblem(src, nameCol, totCol, hdrName, hdrTot)     ' v25: the rule on its own (SmokeTest runs it)
    If missingHdr <> "" Then
        Application.EnableCancelKey = xlInterrupt
        If tablesTouched Then MarkNoSave
        MsgBox "The Kronos sheet doesn't have the heading(s) the import needs: " & missingHdr & "." & vbNewLine & _
               IIf(tablesTouched, "The Kronos tables were updated by this run, so SAVING IS BLOCKED: close WITHOUT saving (the backup copy is " & bk & ").", "Nothing was changed.") & vbNewLine & vbNewLine & _
               "Open " & fileName & ", sheet '" & KRONOS_SHEET & "', and check its heading row (it must be within the first 10 rows). " & _
               "If the report's layout has changed, send " & PRIMARY_ADMIN_NAME & " a screenshot of its first rows.", vbExclamation, "Kronos import"
        Exit Function
    End If
    ' v24.2: the Payroll # column, found by its header (not there = the import works as before)
    Set payOf = CreateObject("Scripting.Dictionary")        ' name key -> Payroll #
    Set keyOfNum = CreateObject("Scripting.Dictionary")  ' Payroll # -> name key ("" = two people share it)
    Dim hdrRow As Long
    hdrRow = hdrName
    pCol = PayrollColumn(src, hdrRow)                 ' v24.3.1: on the heading row only (a title cell above it isn't the column)
    For i = hdrRow + 1 To UBound(src, 1)
        raw = Trim$(SafeStr(src(i, nameCol)))
        k = NameKey(raw)
        If k <> "" Then
            If pCol > 0 And Not dict.Exists(k) Then
                pn = PayrollText(src(i, pCol))
                If pn <> "" Then
                    payOf(k) = pn
                    If keyOfNum.Exists(pn) Then keyOfNum(pn) = "" Else keyOfNum.Add pn, k
                End If
            End If
            If Not dict.Exists(k) Then dict.Add k, src(i, totCol): rawOf.Add k, raw
            If raw Like "* [A-Za-z]" Or raw Like "* [A-Za-z]." Then
                a = NameKey(Left$(raw, InStrRev(raw, " ") - 1))
                If alt.Exists(a) Then
                    alt(a) = ""                 ' two people share it - don't guess
                Else
                    alt.Add a, k
                End If
            End If
        End If
    Next i
    KronosNameIndex = True
End Function

' SyncKronosData's closing boxes (v27.1: moved out of it unchanged): the Kronos names that matched nobody, then the
' result and the Payroll # lists. Reads only. No handler of its own: an error goes to SyncKronosData's Oops.
Private Sub KronosImportMessages(ByVal dict As Object, ByVal alt As Object, ByVal used As Object, ByVal rawOf As Object, _
                                 ByVal found As Long, ByVal nZeroed As Long, ByVal fromReport As Boolean, ByVal fileDate As Date, _
                                 ByVal through As Date, ByVal histId As String, ByVal repTxt As String, ByVal wpTxt As String, _
                                 ByVal awaitFrom As Date, ByVal pCol As Long, ByVal nFilled As Long, ByVal nameClash As String, _
                                 ByVal missing As String, ByVal zeroed As String, ByVal payDiff As String, ByVal payShared As String, _
                                 ByVal noNum As String, ByVal nNoNum As Long)
    ' Kronos names that matched nobody on the roster
    Dim extra As String, nExtra As Long, kk As Variant
    Dim formerKeys As Object: Set formerKeys = CreateObject("Scripting.Dictionary")     ' v24: people on the Former sheet are known
    On Error Resume Next
    Dim fk As String
    For Each kk In FormerNameSet().Keys
        fk = FindKey(dict, alt, NameKey(CStr(kk)))         ' the same match a roster name gets (middle initial in Kronos)
        If fk <> "" Then formerKeys(fk) = True
    Next kk
    On Error GoTo 0                                    ' an error from here raises to SyncKronosData's Oops
    For Each kk In dict.Keys
        If Not used.Exists(kk) And Not formerKeys.Exists(kk) Then
            nExtra = nExtra + 1
            If nExtra <= 15 Then extra = extra & "  " & rawOf(kk) & vbNewLine
        End If
    Next kk
    If nExtra > 15 Then extra = extra & "  ... and " & (nExtra - 15) & " more" & vbNewLine

    Dim msg As String
    msg = "Kronos import complete. " & found & " names matched (green)." & IIf(nZeroed > 0, " " & nZeroed & " roster people have no hours in the report: Worked set to 0.", "") & vbNewLine & _
          IIf(fromReport, "Worked = the Kronos Totals through " & Format$(KronosCumThrough(), "ddd mmm d") & " + the one-day reports after it", _
                          "Kronos file saved " & Format$(fileDate, "ddd mmm d, h:mm AM/PM")) & ", data through " & _
          Format$(through, "ddd mmm d") & "." & vbNewLine & _
          "Accepted / awarded shifts that ended by then now count from Kronos instead of as provisional." & _
          IIf(histId <> "", vbNewLine & "Snapshot " & histId & " added to Hours History.", "") & _
          IIf(fromReport, repTxt & vbNewLine & wpTxt, vbNewLine & "Read from the pivot sheet '" & KRONOS_SHEET & "' (no per-period hours: the daily Kronos reports give those)." & repTxt) & _
          IIf(awaitFrom > 0, vbNewLine & vbNewLine & "FIRST IMPORT OF THE NEW YEAR: check two or three people's Worked hours against Kronos. " & _
                             "They should be this year's hours only, not last year's total.", "")
    If pCol = 0 Then
        msg = msg & vbNewLine & "No 'Payroll #' column was found in the Kronos sheet, so names were matched as before."
    ElseIf nFilled > 0 Then
        msg = msg & vbNewLine & "Payroll # filled in for " & nFilled & " people (roster column A, admin view)."
    End If
    If nameClash <> "" Then msg = msg & vbNewLine & vbNewLine & _
        "SAME Payroll #, DIFFERENT NAME - nothing changed for these (rename them with Employee > Rename, or fix column A):" & vbNewLine & nameClash
    If missing <> "" Then msg = msg & vbNewLine & vbNewLine & _
        "On the roster but not found in Kronos - Worked Hours NOT updated:" & vbNewLine & missing
    If zeroed <> "" Then msg = msg & vbNewLine & "No premium hours in the Kronos report this year - Worked set to 0:" & vbNewLine & zeroed & _
        IIf(nZeroed > 10, "  ... and " & (nZeroed - 10) & " more" & vbNewLine, "")
    If extra <> "" Then msg = msg & vbNewLine & _
        "In Kronos but not on the roster (spelling different, or not on this roster):" & vbNewLine & extra
    If Len(msg) > 1000 Then msg = Left$(msg, 980) & vbNewLine & "... (more)"
    MsgBox msg, IIf(missing = "" And extra = "" And nameClash = "", vbInformation, vbExclamation), "Kronos import"
    ' v24.2: the Payroll # lists get a box of their own (names only, never a number)
    Dim payMsg As String
    If payDiff <> "" Then payMsg = payMsg & "Matched by name, but column A doesn't hold the Payroll # Kronos has - check column A:" & vbNewLine & payDiff & vbNewLine
    If payShared <> "" Then payMsg = payMsg & "Kronos has one Payroll # under two names, so column A was left blank for:" & vbNewLine & payShared & vbNewLine
    If noNum <> "" Then payMsg = payMsg & "Roster rows with no Payroll # yet (" & nNoNum & "):" & vbNewLine & noNum
    If payMsg <> "" Then
        If Len(payMsg) > 1000 Then payMsg = Left$(payMsg, 980) & vbNewLine & "... (more)"
        MsgBox payMsg, IIf(payDiff = "" And payShared = "", vbInformation, vbExclamation), "Kronos import - Payroll #"
    End If
End Sub

' v27 (review H3): the rollover empties the Kronos Totals snapshot and the Worked Days ledger (last year's) and drops their
' two dates; the Worked Changes log stays (append-only history). Quiet when the tables were never made.
Public Sub ClearKronosTables()
    Dim lo As ListObject, outA() As Variant
    ReDim outA(1 To 1, 1 To 1)
    Set lo = KronosTotalsTable(False)
    If Not lo Is Nothing Then ReplaceRows lo, outA, 0, KT_NCOLS, "", ""
    Set lo = WorkedDaysTable(False)
    If Not lo Is Nothing Then ReplaceRows lo, outA, 0, WD_NCOLS, "", ""
    ' (review R3) Kronos Files is KEPT: it holds file names only, and emptying it would make January reopen all of last year's
    On Error Resume Next
    ThisWorkbook.Names(KCUM_NAME).Delete
    ThisWorkbook.Names(KLEDGER_NAME).Delete
    On Error GoTo 0
End Sub

' ====================================================================
'  v26.5: THE KRONOS EXPORT READ DIRECTLY, AND THE PREMIUM HOURS PER PAY PERIOD
' ====================================================================
' KRONOS_PATH resolved: a file as it is; a folder (ending in \) or a wildcard
' ("...\*.xlsx") = the newest matching file by its saved time ("" = nothing)
Public Function KronosFilePath() As String
    Dim p As String: p = Trim$(KRONOS_PATH)
    If p = "" Then Exit Function
    On Error Resume Next                                     ' an unreachable drive: "" (the caller says no file was found) (never raises)
    Dim hit As String
    Dim pattern As String, folder As String
    If Right$(p, 1) = "\" Then
        folder = p: pattern = "*.xls*"
    ElseIf InStr(p, "*") > 0 Or InStr(p, "?") > 0 Then
        folder = Left$(p, InStrRev(p, "\")): pattern = Mid$(p, InStrRev(p, "\") + 1)
    Else
        hit = Dir$(p)                                        ' on its own line: under Resume Next an error inside an If's test would run the Then part
        If hit <> "" Then KronosFilePath = p
        Exit Function
    End If
    Dim f As String, best As String, bestAt As Date, at As Date
    f = Dir$(folder & pattern)
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then                      ' not Excel's lock files
            at = FileDateTime(folder & f)
            If at > bestAt Then bestAt = at: best = f
        End If
        f = ""                                           ' cleared first: under Resume Next a share dropping mid-listing must end the loop
        f = Dir$()
    Loop
    If best <> "" Then KronosFilePath = folder & best
End Function

' ====================================================================
'  v27: THE REPORT, THE SNAPSHOT, THE LEDGER, THE CHANGE LOG
' ====================================================================
' The files to import: KRONOS_PATH a folder (or wildcard) = every report file saved after the last import, oldest
' first (none newer = the newest one, so a re-run re-reads it); a file = that file. Empty (not an array) = nothing.
Public Function KronosFilesToImport(Optional ByRef nNew As Long = 0) As Variant     ' v27 (item 2): nNew = how many the roster has not read
    nNew = 0
    Dim p As String: p = Trim$(KRONOS_PATH)
    If p = "" Then Exit Function
    On Error Resume Next                                     ' an unreachable drive: nothing (the caller says no file was found) (never raises)
    Dim folder As String, pattern As String, hit As String
    If Right$(p, 1) = "\" Then
        folder = p: pattern = "*.xls*"
    ElseIf InStr(p, "*") > 0 Or InStr(p, "?") > 0 Then
        folder = Left$(p, InStrRev(p, "\")): pattern = Mid$(p, InStrRev(p, "\") + 1)
    Else
        hit = Dir$(p)
        If hit <> "" Then
            KronosFilesToImport = Array(p)
            Dim fAt As Date: fAt = FileDateTime(p)       ' read first: under Resume Next an error in an If test runs its Then
            ' (review M1) one file (the processing workbook, read through the pivot sheet, never recorded in Kronos Files):
            ' new = saved after the last import, nothing else
            If fAt > KronosImportedAt() Then nNew = 1
        End If
        Exit Function
    End If
    Dim since As Date: since = KronosImportedAt()
    Dim known As Object: Set known = KnownSources()          ' (review M10): a file copied in with an old saved time is still new to the roster
    Dim names() As String, ats() As Date, n As Long, f As String, at As Date, best As String, bestAt As Date
    f = Dir$(folder & pattern)
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then                      ' not Excel's lock files
            at = 0
            at = FileDateTime(folder & f)
            If at > bestAt Then bestAt = at: best = f
            If at > since Or Not known.Exists(f) Then
                n = n + 1
                ReDim Preserve names(1 To n): ReDim Preserve ats(1 To n)
                names(n) = f: ats(n) = at
            End If
        End If
        f = ""                                           ' cleared first: under Resume Next a share dropping mid-listing must end the loop
        f = Dir$()
    Loop
    nNew = n
    If n = 0 Then
        If best <> "" Then KronosFilesToImport = Array(folder & best)
        Exit Function
    End If
    ' oldest first (a few files: a plain exchange sort)
    Dim i As Long, j As Long, tn As String, ta As Date
    For i = 1 To n - 1
        For j = i + 1 To n
            If ats(j) < ats(i) Then
                tn = names(i): names(i) = names(j): names(j) = tn
                ta = ats(i): ats(i) = ats(j): ats(j) = ta
            End If
        Next j
    Next i
    Dim out() As Variant: ReDim out(0 To n - 1)
    For i = 1 To n
        out(i - 1) = folder & names(i)
    Next i
    KronosFilesToImport = out
End Function

' v27 (item 2): how many report files in KRONOS_PATH the roster has not read yet (the admin Desk's Kronos pill)
Public Function KronosNewFileCount() As Long
    On Error Resume Next                                     ' never raises (the Desk draws on it)
    Dim n As Long, v As Variant
    v = KronosFilesToImport(n)
    KronosNewFileCount = n
End Function

' The report file names the tables already hold (the Source columns of Worked Days and Worked Changes)
Private Function KnownSources() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Set KnownSources = d
    Dim v As Variant, i As Long
    v = TableRows(KronosFilesTable(False))                   ' (review N7): every file read before, taken or skipped
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Trim$(SafeStr(v(i, 1))) <> "" And LCase$(Left$(Trim$(SafeStr(v(i, 1))), 7)) <> "clash: " Then d(Trim$(SafeStr(v(i, 1)))) = True
        Next i
    End If
    v = TableRows(WorkedDaysTable(False))
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Trim$(SafeStr(v(i, 7))) <> "" Then d(Trim$(SafeStr(v(i, 7)))) = True
        Next i
    End If
End Function

' The record of every report file read (admin-only, never published): File, Result, Time Period, Imported At. A
' file read again gets its row replaced, so the table holds one row per file name.
Public Function KronosFilesTable(Optional ByVal create As Boolean = False) As ListObject
    Set KronosFilesTable = KronosTable(KF_SHEET, KF_TABLE, create, _
        Array("File", "Result", "Time Period", "Imported At"), Array(44, 60, 24, 17))
End Function

Private Sub RecordKronosFile(ByVal fileName As String, ByVal result As String, ByRef rep As KronosReport)
    Dim lo As ListObject: Set lo = KronosFilesTable(True)
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableRows(lo)
    Dim n As Long, i As Long, r As Long
    If IsArray(v) Then n = UBound(v, 1)
    Dim outA() As Variant: ReDim outA(1 To n + 1, 1 To KF_NCOLS)
    If IsArray(v) Then
        For i = 1 To n
            If StrComp(Trim$(SafeStr(v(i, 1))), fileName, vbTextCompare) <> 0 Then
                r = r + 1
                outA(r, 1) = v(i, 1): outA(r, 2) = v(i, 2): outA(r, 3) = v(i, 3): outA(r, 4) = v(i, 4)
            End If
        Next i
    End If
    r = r + 1
    outA(r, 1) = fileName: outA(r, 2) = Left$(result, 250)
    outA(r, 3) = IIf(rep.FromD > 0, Format$(rep.FromD, "yyyy-mm-dd") & " - " & Format$(rep.ToD, "yyyy-mm-dd"), "")
    outA(r, 4) = Now
    ReplaceRows lo, outA, r, KF_NCOLS, "", "4"
End Sub

' (review R1) the clash names held in Kronos Files as "clash: <name>" rows (KnownSources ignores them: not file names)
Private Function HeldClashNames() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Set HeldClashNames = d
    Dim v As Variant: v = TableRows(KronosFilesTable(False))
    Dim i As Long, f As String
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        f = Trim$(SafeStr(v(i, 1)))
        If LCase$(Left$(f, 7)) = "clash: " Then d(Trim$(Mid$(f, 8))) = True
    Next i
End Function

' The current clash set written back: the old "clash: " rows replaced by this run's
Private Sub SaveClashNames()
    Dim lo As ListObject: Set lo = KronosFilesTable(True)
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableRows(lo)
    Dim n As Long, i As Long, r As Long, kk As Variant
    If IsArray(v) Then n = UBound(v, 1)
    Dim outA() As Variant: ReDim outA(1 To n + mClashNames.Count + 1, 1 To KF_NCOLS)
    If IsArray(v) Then
        For i = 1 To n
            If LCase$(Left$(Trim$(SafeStr(v(i, 1))), 7)) <> "clash: " Then
                r = r + 1
                outA(r, 1) = v(i, 1): outA(r, 2) = v(i, 2): outA(r, 3) = v(i, 3): outA(r, 4) = v(i, 4)
            End If
        Next i
    End If
    For Each kk In mClashNames.Keys
        r = r + 1
        outA(r, 1) = "clash: " & kk: outA(r, 2) = "a Kronos row's Payroll # and name point at two roster people: Worked left alone (Employee > Rename, or fix column A); cleared by the next clean year-to-date import"
        outA(r, 3) = "": outA(r, 4) = Now
    Next kk
    ReplaceRows lo, outA, r, KF_NCOLS, "", "4"
End Sub

' A report record that carries only a Time Period (for the file record of a one-day report)
Private Function KronosReportStub(ByVal fromD As Date, ByVal toD As Date) As KronosReport
    Dim rep As KronosReport
    rep.FromD = fromD: rep.ToD = toD
    KronosReportStub = rep
End Function

' A sheet read whole as a 2-D array (at least 2 x 2); the last cell by rows and by columns, LookIn:=xlFormulas so a
' hidden or filtered-out row still counts (v24.3.1)
Private Function SheetValues(ByVal ws As Worksheet) As Variant
    Dim lastSrc As Long, lastCol As Long, lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastSrc = 2 Else lastSrc = lastCell.Row
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then lastCol = 2 Else lastCol = lastCell.Column
    If lastSrc < 2 Then lastSrc = 2
    If lastCol < 2 Then lastCol = 2
    SheetValues = ws.Range(ws.Cells(1, 1), ws.Cells(lastSrc, lastCol)).Value
End Function

' The first cell to the right of column c on row r that holds anything ("" when none)
Private Function RowText(ByRef v As Variant, ByVal r As Long, ByVal c As Long) As String
    Dim j As Long
    For j = c + 1 To UBound(v, 2)
        If Trim$(SafeStr(v(r, j))) <> "" Then RowText = Trim$(SafeStr(v(r, j))): Exit Function
    Next j
End Function

' Reads the report "Employee Transactions and Totals" from whichever sheet holds its "Totals:" block: the title's
' Time Period and Query, the Totals rows (Global ID, Employee, Pay Code, Hours) and the Transactions rows (+ Date,
' Edited By). rep.HasTotals False = not a report (the caller tries the pivot sheet). Returns "" or the problem.
Private Function ReadKronosReport(ByVal wb As Workbook, ByRef rep As KronosReport) As String
    Dim ws As Worksheet, v As Variant, prob As String
    rep.HasTotals = False: rep.nTot = 0: rep.nTx = 0: rep.Query = "": rep.FromD = 0: rep.ToD = 0
    For Each ws In wb.Worksheets
        v = SheetValues(ws)
        prob = ReadKronosArray(v, rep)
        If rep.HasTotals Then ReadKronosReport = prob: Exit Function
    Next ws
End Function

' The report as a 2-D array (one sheet): the title's Time Period and Query, the Totals rows (Global ID, Employee,
' Pay Code, Hours) and the Transactions rows (+ Date, Edited By). rep.HasTotals False = this sheet is not the report.
' Returns "" or the problem. (review M9: the array form is what SmokeTest step 15 feeds)
Private Function ReadKronosArray(ByRef v As Variant, ByRef rep As KronosReport) As String
    Dim r As Long, c As Long, txt As String
    Dim totHdr As Long, txHdr As Long, sumRow As Long, periodTxt As String
    rep.HasTotals = False: rep.nTot = 0: rep.nTx = 0: rep.Query = "": rep.FromD = 0: rep.ToD = 0
    For r = 1 To UBound(v, 1)
        txt = Trim$(SafeStr(v(r, 1)))
        If LCase$(Left$(txt, 11)) = "time period" Then
            periodTxt = RowText(v, r, 1)
            If periodTxt = "" Then periodTxt = txt
        ElseIf LCase$(Left$(txt, 5)) = "query" Then
            rep.Query = RowText(v, r, 1)
        ElseIf LCase$(txt) = "transactions:" Then
            txHdr = r + 1
        ElseIf LCase$(txt) = "totals:" Then
            totHdr = r + 1
        ElseIf LCase$(txt) = "summary totals:" Then
            sumRow = r
            Exit For
        End If
    Next r
    If totHdr = 0 Or totHdr > UBound(v, 1) Then Exit Function
    rep.HasTotals = True
    If sumRow = 0 Then sumRow = UBound(v, 1) + 1
    ' the Time Period: the report's two dates, read by hand
    Dim prob As String: prob = TimePeriodCore(periodTxt, rep.FromD, rep.ToD)
    If prob <> "" Then ReadKronosArray = "the Time Period line could not be read (" & prob & "): " & periodTxt: Exit Function
    ' the Totals block's columns, by heading (the heading sits two columns left of the Transactions block's)
    Dim cId As Long, cEmp As Long, cCode As Long, cHrs As Long
    For c = 1 To UBound(v, 2)
        If IsHeaderMatch(SafeStr(v(totHdr, c)), "Global ID") Then cId = c
        If IsHeaderMatch(SafeStr(v(totHdr, c)), "Employee") Then cEmp = c
        If IsHeaderMatch(SafeStr(v(totHdr, c)), "Pay Code") Then cCode = c
        If IsHeaderMatch(SafeStr(v(totHdr, c)), "Hours") Then cHrs = c
    Next c
    If cId = 0 Or cEmp = 0 Or cCode = 0 Or cHrs = 0 Then
        ReadKronosArray = "the Totals block's heading lacks " & IIf(cId = 0, "Global ID ", "") & IIf(cEmp = 0, "Employee ", "") & _
                          IIf(cCode = 0, "Pay Code ", "") & IIf(cHrs = 0, "Hours", "")
        Exit Function
    End If
    Dim n As Long: n = sumRow - totHdr
    If n < 1 Then n = 1
    ReDim rep.TotId(1 To n): ReDim rep.TotName(1 To n): ReDim rep.TotCode(1 To n): ReDim rep.TotHrs(1 To n)
    For r = totHdr + 1 To sumRow - 1
        If Trim$(SafeStr(v(r, cEmp))) <> "" And Trim$(SafeStr(v(r, cCode))) <> "" And IsNumeric(v(r, cHrs)) And Not IsEmpty(v(r, cHrs)) Then
            rep.nTot = rep.nTot + 1
            rep.TotId(rep.nTot) = PayrollText(v(r, cId))
            rep.TotName(rep.nTot) = Trim$(SafeStr(v(r, cEmp)))
            rep.TotCode(rep.nTot) = Trim$(SafeStr(v(r, cCode)))
            rep.TotHrs(rep.nTot) = CDbl(v(r, cHrs))
        End If
    Next r
    ' the Transactions block (edits): only to date a change; missing = fine
    If txHdr > 0 And txHdr < totHdr - 1 Then
        Dim tId As Long, tDt As Long, tCode As Long, tBy As Long
        For c = 1 To UBound(v, 2)
            If IsHeaderMatch(SafeStr(v(txHdr, c)), "Global ID") Then tId = c
            If IsHeaderMatch(SafeStr(v(txHdr, c)), "Date") Then tDt = c
            If IsHeaderMatch(SafeStr(v(txHdr, c)), "Pay Code") Then tCode = c
            If IsHeaderMatch(SafeStr(v(txHdr, c)), "Edited By") Then tBy = c
        Next c
        If tId > 0 And tDt > 0 And tCode > 0 Then
            n = totHdr - 1 - txHdr
            If n < 1 Then n = 1
            ReDim rep.TxId(1 To n): ReDim rep.TxDate(1 To n): ReDim rep.TxCode(1 To n): ReDim rep.TxBy(1 To n)
            For r = txHdr + 1 To totHdr - 2
                If VarType(v(r, tDt)) = vbDate And Trim$(SafeStr(v(r, tCode))) <> "" Then     ' a real date cell only (convention 9: never CDate a string)
                    rep.nTx = rep.nTx + 1
                    rep.TxId(rep.nTx) = PayrollText(v(r, tId))
                    rep.TxDate(rep.nTx) = CDate(v(r, tDt))
                    rep.TxCode(rep.nTx) = Trim$(SafeStr(v(r, tCode)))
                    If tBy > 0 Then rep.TxBy(rep.nTx) = Trim$(SafeStr(v(r, tBy)))
                End If
            Next r
        End If
    End If
End Function

' SmokeTest step 15's hook: a made-up report array -> "from|to|totals rows|transactions rows|sum of the Totals hours
' for <id>" so the step proves the dated edit rows are NOT in the sum (review M9). Public, holds no rule of its own.
Public Function KronosArrayCheck(ByRef v As Variant, ByVal payrollId As String) As String
    Dim rep As KronosReport, prob As String, i As Long, s As Double
    prob = ReadKronosArray(v, rep)
    If Not rep.HasTotals Then KronosArrayCheck = "no Totals block": Exit Function
    If prob <> "" Then KronosArrayCheck = prob: Exit Function
    For i = 1 To rep.nTot
        If rep.TotId(i) = payrollId Then s = s + rep.TotHrs(i)
    Next i
    KronosArrayCheck = Format$(rep.FromD, "yyyy-mm-dd") & "|" & Format$(rep.ToD, "yyyy-mm-dd") & "|" & rep.nTot & "|" & rep.nTx & "|" & Trim$(Str$(s))
End Function

' The roster's spelling for a report row: by Payroll # (column A) first, then the name's key, then the name without
' a trailing middle initial ("" = not a roster person)
Private Function RosterNameFor(ByVal id As String, ByVal raw As String, ByVal nameMap As Object, ByVal numMap As Object) As String
    Dim byNum As String, byName As String
    If id <> "" Then
        If numMap.Exists(id) Then byNum = CStr(numMap(id))
    End If
    Dim k As String: k = NameKey(raw)
    If nameMap.Exists(k) Then
        byName = CStr(nameMap(k))
    ElseIf raw Like "* [A-Za-z]" Or raw Like "* [A-Za-z]." Then
        k = NameKey(Left$(raw, InStrRev(raw, " ") - 1))
        If nameMap.Exists(k) Then byName = CStr(nameMap(k))
    End If
    ' (review M3 / N1) the same rule as PayrollDecision: a number that the roster holds under ANOTHER name is a clash -
    ' nothing is taken for either, BOTH roster people keep the Worked they have, and the import names them
    Select Case RosterNameCore(byNum, byName)
        Case "clash"
            If Not mClashNames.Exists(byNum) Then mClashNames.Add byNum, True
            If Not mClashNames.Exists(byName) Then mClashNames.Add byName, True
            If InStr(mClashes, vbNewLine & "    " & raw & " ") = 0 Then mClashes = mClashes & vbNewLine & "    " & raw & " (its Payroll # is on the roster as " & byNum & ")"
            Exit Function
        Case "number"
            RosterNameFor = byNum
        Case Else
            RosterNameFor = byName
    End Select
End Function

' A report into the tables. A ONE-DAY report: that day's rows into the Worked Days ledger (a day already held is
' replaced, every difference logged, kind "day"). A CUMULATIVE report: the Kronos Totals snapshot replaced; what the
' ledger expected (the old snapshot + the snapshot days after its Through) compared code by code, every difference
' logged, kind "edit", dated by the Transactions block when it names that person and code on an earlier day. Roster
' people only (the roster's spelling). Returns the line for the message.
Private Function IngestReport(ByRef rep As KronosReport, ByVal srcName As String, ByVal stamp As Date, _
                              ByVal nameMap As Object, ByVal numMap As Object) As String
    Dim kind As String: kind = ReportKindCore(rep.FromD, rep.ToD, RosterYearFrom(), RosterYearTo())
    Dim span As String: span = Format$(rep.FromD, "mmm d") & " - " & Format$(rep.ToD, "mmm d, yyyy")
    Select Case kind
        Case "outside"
            IngestReport = "Time Period " & span & " is outside this roster's year - skipped": Exit Function
        Case "range"
            IngestReport = "Time Period " & span & " covers several days but not the year to date: its totals cannot be split into days - run one day at a time, or Jan .. yesterday - skipped"
            Exit Function
    End Select
    ' (review H4) an older year-to-date than the one held would move the snapshot back and log every difference as an edit
    If kind = "cumulative" And rep.ToD < KronosCumThrough() Then
        IngestReport = "the year to date " & span & " is OLDER than the Kronos Totals already held (through " & Format$(KronosCumThrough(), "mmm d, yyyy") & ") - skipped"
        Exit Function
    End If
    Dim qTxt As String
    If InStr(1, rep.Query, KRONOS_QUERY, vbTextCompare) = 0 Then qTxt = " (Query '" & rep.Query & "', not '" & KRONOS_QUERY & "': other codes are reported, not counted)"
    ' the report's rows for roster people, summed per person and code (a person's two rows for one code add up)
    Dim tot As Object: Set tot = CreateObject("Scripting.Dictionary"): tot.CompareMode = 1     ' "name|code" -> hours
    Dim payOf As Object: Set payOf = CreateObject("Scripting.Dictionary"): payOf.CompareMode = 1   ' name -> Payroll #
    Dim i As Long, nm As String, key As String, nPeople As Long, other As String, nOther As Long, m As Double
    mClashes = ""
    ' (review R1 / M1): the year to date is the truth about clashes - the held set is swapped out while it is read, and put
    ' back if the file is refused below (a wrong-query file must not drop what earlier runs found)
    Dim heldClashes As Object
    If kind = "cumulative" Then
        Set heldClashes = mClashNames
        Set mClashNames = CreateObject("Scripting.Dictionary"): mClashNames.CompareMode = 1
    End If
    For i = 1 To rep.nTot
        nm = RosterNameFor(rep.TotId(i), rep.TotName(i), nameMap, numMap)
        If nm <> "" Then
            If Not payOf.Exists(nm) Then payOf.Add nm, rep.TotId(i): nPeople = nPeople + 1
            key = nm & "|" & rep.TotCode(i)
            If tot.Exists(key) Then tot(key) = tot(key) + rep.TotHrs(i) Else tot.Add key, rep.TotHrs(i)
            m = PremiumMultiplier(rep.TotCode(i), PREMIUM_CODES, PREMIUM_MULTIPLIERS)
            If m < 0 And Abs(rep.TotHrs(i)) > 0.001 Then
                nOther = nOther + 1
                If nOther <= 10 Then other = other & vbNewLine & "    " & nm & ": " & rep.TotCode(i) & " " & HoursText(rep.TotHrs(i)) & " h"
            End If
        End If
    Next i
    If nOther > 10 Then other = other & vbNewLine & "    ... and " & (nOther - 10) & " more"
    ' (review H4) a year to date with nobody from the roster in it would empty the snapshot: refused before any write
    If kind = "cumulative" And nPeople = 0 Then
        Set mClashNames = heldClashes                        ' (review M1): nothing taken, so the held clashes stand
        IngestReport = "the year to date " & span & " names nobody on the roster (the wrong query or site?) - skipped" & IIf(mClashes <> "", vbNewLine & "  Payroll # clashes:" & mClashes, "")
        Exit Function
    End If
    ' (review M2) every roster person: a day's rows are replaced for all of them, not only those in this report
    Dim rosterSet As Object: Set rosterSet = CreateObject("Scripting.Dictionary"): rosterSet.CompareMode = 1
    Dim kk As Variant
    For Each kk In nameMap.Keys
        rosterSet(CStr(nameMap(kk))) = True
    Next kk
    Dim nChanges As Long, line As String
    If kind = "day" Then
        nChanges = WriteLedgerDay(rep.ToD, tot, payOf, rosterSet, srcName, stamp)
        Dim lf As Date: lf = KronosLedgerFrom()
        If lf = 0 Or rep.ToD < lf Then SetDateName KLEDGER_NAME, rep.ToD
        line = "one-day report for " & Format$(rep.ToD, "ddd mmm d") & ": " & nPeople & " roster people into the Worked Days ledger" & _
               IIf(rep.ToD <= KronosCumThrough(), " (inside the Totals through " & Format$(KronosCumThrough(), "mmm d") & ": kept, not added)", "") & _
               IIf(nChanges > 0, "; " & nChanges & " figure(s) differed from the day as first imported (Worked Changes)", "")
    Else
        nChanges = WriteCumulative(rep, tot, payOf, rosterSet, srcName, stamp)     ' v29.0.3 (review 1.11): removed people stay out of the expected map
        line = "the year to date " & span & ": " & nPeople & " roster people into the Kronos Totals"
        If mGapNote <> "" Then                        ' v27.5 (item 2): a missed morning is named, its differences are gap rows
            line = line & "; NO one-day report for " & mGapMsg & " - " & IIf(nChanges > 0, nChanges & " figure(s) differed from what the ledger expected: " & _
                   "in Worked Changes as gap rows, counted in Worked Periods until those days are imported", "nothing differed")
        ElseIf nChanges > 0 Then
            line = line & "; " & nChanges & " figure(s) differed from what the ledger expected - timecard edits, in Worked Changes"
        End If
    End If
    IngestReport = line & qTxt & IIf(other <> "", vbNewLine & "  codes NOT counted on roster people (" & nOther & "):" & other, "") & _
                   IIf(mClashes <> "", vbNewLine & "  SAME Payroll #, DIFFERENT NAME - nothing taken for these (Employee > Rename, or fix column A):" & mClashes, "")
End Function

' The ledger table, made on first use (admin-only, never published)
Public Function WorkedDaysTable(Optional ByVal create As Boolean = False) As ListObject
    Set WorkedDaysTable = KronosTable(WD_SHEET, WD_TABLE, create, _
        Array("Payroll #", "Employee", "Date", "Pay Code", "Hours", "Weighted", "Source", "Imported At"), _
        Array(11, 24, 12, 20, 8, 9, 40, 17))
End Function

' The snapshot of the last cumulative report, made on first use (admin-only, never published)
Public Function KronosTotalsTable(Optional ByVal create As Boolean = False) As ListObject
    Set KronosTotalsTable = KronosTable(KT_SHEET, KT_TABLE, create, _
        Array("Payroll #", "Employee", "Pay Code", "Hours", "Through", "Imported At"), _
        Array(11, 24, 20, 8, 12, 17))
End Function

' The change log, made on first use (admin-only, never published; rows are only ever appended)
Public Function WorkedChangesTable(Optional ByVal create As Boolean = False) As ListObject
    Set WorkedChangesTable = KronosTable(WC_SHEET, WC_TABLE, create, _
        Array("When", "Payroll #", "Employee", "Date", "Pay Code", "Hours Before", "Hours After", "Change", "Weighted Change", "Kind", "Note", "Source", "Imported By"), _
        Array(17, 11, 24, 12, 20, 9, 9, 8, 9, 8, 36, 40, 12))
End Function

' One pattern for the three tables above (and WorkedPeriodsTable's): the sheet at the end, the table on A1
Private Function KronosTable(ByVal sheetName As String, ByVal tableName As String, ByVal create As Boolean, _
                             ByVal heads As Variant, ByVal widths As Variant) As ListObject
    Dim ws As Worksheet, lo As ListObject, nCols As Long, c As Long
    nCols = UBound(heads) - LBound(heads) + 1
    If SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets(sheetName)
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = sheetName
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, nCols).Value = heads
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, nCols), , xlYes)
        lo.Name = tableName
        lo.TableStyle = "TableStyleLight9"
        For c = 1 To nCols
            ws.Columns(c).ColumnWidth = widths(LBound(widths) + c - 1)
        Next c
        ProtectSheet ws
    End If
    Set KronosTable = lo
End Function

' A table's rows as a 2-D array (Empty when none; a blank first row counts as none)
Private Function TableRows(ByVal lo As ListObject) As Variant
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then Exit Function
    TableRows = RangeTo2D(lo.DataBodyRange)
End Function

' Replaces a table's rows with outA (n rows), or empties it; the sheet unlocked and protected around the write
Private Sub ReplaceRows(ByVal lo As ListObject, ByRef outA As Variant, ByVal n As Long, ByVal nCols As Long, ByVal dateCols As String, ByVal stampCols As String)
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If n > 0 Then
        lo.Resize lo.HeaderRowRange.Resize(n + 1)
        With lo.HeaderRowRange.Offset(1, 0).Resize(n, nCols)
            .Columns(1).NumberFormat = "@"
            .Value = outA
            FormatCols .Cells, dateCols, "mmm d, yyyy"
            FormatCols .Cells, stampCols, "mmm d, yyyy h:mm"
        End With
    End If
    ProtectSheet ws
End Sub

' Appends outA (n rows) to a table
Private Sub AppendRows(ByVal lo As ListObject, ByRef outA As Variant, ByVal n As Long, ByVal nCols As Long, ByVal dateCols As String, ByVal stampCols As String)
    If n = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If Not lo.DataBodyRange Is Nothing Then
        If cnt = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, nCols)
        .Columns(2).NumberFormat = "@"
        .Value = outA
        FormatCols .Cells, dateCols, "mmm d, yyyy"
        FormatCols .Cells, stampCols, "mmm d, yyyy h:mm"
    End With
    ProtectSheet ws
End Sub

Private Sub FormatCols(ByVal rng As Range, ByVal cols As String, ByVal fmt As String)
    Dim p As Variant, c As Variant
    If cols = "" Then Exit Sub
    p = Split(cols, ",")
    For Each c In p
        rng.Columns(CLng(c)).NumberFormat = fmt
    Next c
End Sub

Public Function KronosCumThrough() As Date
    KronosCumThrough = GetDateName(KCUM_NAME)
End Function

Public Function KronosLedgerFrom() As Date
    KronosLedgerFrom = GetDateName(KLEDGER_NAME)
End Function

' One snapshot day into the ledger: the day's rows for roster people replaced (a difference from the day as
' first imported logged, kind "day"). Returns the number of change rows written.
Private Function WriteLedgerDay(ByVal d As Date, ByVal tot As Object, ByVal payOf As Object, ByVal rosterSet As Object, ByVal srcName As String, ByVal stamp As Date) As Long
    Dim lo As ListObject: Set lo = WorkedDaysTable(True)
    If lo Is Nothing Then Err.Raise vbObjectError + 790, , "the Worked Days sheet couldn't be made"
    Dim v As Variant: v = TableRows(lo)
    Dim keep As Long, i As Long, n As Long, key As String, c As Long, emp As String, dayRow As Boolean, hadDay As Boolean
    Dim oldDay As Object: Set oldDay = CreateObject("Scripting.Dictionary"): oldDay.CompareMode = 1
    Dim dropRow() As Boolean
    If IsArray(v) Then
        ReDim dropRow(1 To UBound(v, 1))
        For i = 1 To UBound(v, 1)
            emp = Trim$(SafeStr(v(i, 2)))
            dayRow = LedgerDayDrop(v(i, 3), d, emp, rosterSet.Exists(emp))   ' (review M2 / M7): every roster row of the day, and the quiet-day marker
            dropRow(i) = dayRow
            If dayRow Then
                hadDay = True
                If emp <> "" Then
                    key = emp & "|" & Trim$(SafeStr(v(i, 4)))
                    If oldDay.Exists(key) Then oldDay(key) = oldDay(key) + NumOr0(v(i, 5)) Else oldDay.Add key, NumOr0(v(i, 5))
                End If
            Else
                keep = keep + 1
            End If
        Next i
    End If
    n = keep + tot.Count + 1
    Dim outA() As Variant: ReDim outA(1 To n, 1 To WD_NCOLS)
    Dim r As Long
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Not dropRow(i) Then
                r = r + 1
                For c = 1 To WD_NCOLS
                    outA(r, c) = v(i, c)
                Next c
            End If
        Next i
    End If
    Dim kk As Variant, nm As String, code As String, m As Double
    For Each kk In tot.Keys
        r = r + 1
        nm = Left$(CStr(kk), InStr(CStr(kk), "|") - 1): code = Mid$(CStr(kk), InStr(CStr(kk), "|") + 1)
        m = PremiumMultiplier(code, PREMIUM_CODES, PREMIUM_MULTIPLIERS)
        outA(r, 1) = SafeStr(payOf(nm)): outA(r, 2) = nm: outA(r, 3) = CDate(Int(CDbl(d))): outA(r, 4) = code
        outA(r, 5) = tot(kk): outA(r, 6) = IIf(m < 0, 0, tot(kk) * m): outA(r, 7) = srcName: outA(r, 8) = stamp
    Next kk
    If tot.Count = 0 Then
        ' (review M7) a day with no roster hours still leaves its mark, so the day counts as imported (a blank Employee: every reader skips it)
        r = r + 1
        outA(r, 1) = "": outA(r, 2) = "": outA(r, 3) = CDate(Int(CDbl(d))): outA(r, 4) = "(no roster hours)"
        outA(r, 5) = 0: outA(r, 6) = 0: outA(r, 7) = srcName: outA(r, 8) = stamp
    End If
    mIngested = mIngested + 1                     ' (review N2): from here the tables differ from the saved file
    ReplaceRows lo, outA, r, WD_NCOLS, "3", "8"
    RecordKronosFile srcName, "one-day report taken", KronosReportStub(d, d)
    ' the differences from the day as first imported (the first import of a day logs nothing: the ledger IS the record)
    If hadDay Then
        Dim noRep As KronosReport                     ' a "day" change has no edit to date: an empty report
        WriteLedgerDay = LogChanges(oldDay, tot, payOf, d, "day", "the day re-imported", srcName, stamp, noRep, 0)
    End If
End Function

' The cumulative report into the Kronos Totals snapshot; the differences from what the ledger expected logged, kind
' "edit" (dated by the report's Transactions block when it names that person and code on an earlier day). Returns the
' number of change rows written. The first cumulative ever (no snapshot) logs nothing: it is the baseline.
' v27.5 (item 2): when a day between the old Through and this one has no one-day report, every difference is a "gap"
' row instead - dated at this Through, its Note naming the missed days (mGapNote for the message).
' v29.0.3 (review 1.11): rosterSet = the roster's names - the old snapshot and the ledger keep people removed since the last
' year-to-date, while the report's rows are taken for roster people only, so each removed person logged false "lost hours"
' rows (Lane Wexford: three gap rows at the close-out); names not on the roster stay out of what the ledger expected
Private Function WriteCumulative(ByRef rep As KronosReport, ByVal tot As Object, ByVal payOf As Object, ByVal rosterSet As Object, _
                                 ByVal srcName As String, ByVal stamp As Date) As Long
    Dim lo As ListObject: Set lo = KronosTotalsTable(True)
    If lo Is Nothing Then Err.Raise vbObjectError + 791, , "the Kronos Totals sheet couldn't be made"
    Dim oldThrough As Date: oldThrough = KronosCumThrough()
    Dim v As Variant: v = TableRows(lo)
    Dim i As Long, key As String
    ' what the ledger expected: the old snapshot + the snapshot days after its Through, up to the new Through
    Dim expected As Object: Set expected = CreateObject("Scripting.Dictionary"): expected.CompareMode = 1
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If Trim$(SafeStr(v(i, 2))) <> "" And rosterSet.Exists(Trim$(SafeStr(v(i, 2)))) Then
                key = Trim$(SafeStr(v(i, 2))) & "|" & Trim$(SafeStr(v(i, 3)))
                If expected.Exists(key) Then expected(key) = expected(key) + NumOr0(v(i, 4)) Else expected.Add key, NumOr0(v(i, 4))
                If Not payOf.Exists(Trim$(SafeStr(v(i, 2)))) Then payOf.Add Trim$(SafeStr(v(i, 2))), SafeStr(v(i, 1))   ' (review) the Payroll # for a person only in the old snapshot
            End If
        Next i
    End If
    Dim lv As Variant: lv = TableRows(WorkedDaysTable(False))
    Dim held As Object: Set held = CreateObject("Scripting.Dictionary")      ' v27.5 (item 2): every day the ledger holds (the quiet-day marker too)
    If IsArray(lv) Then
        For i = 1 To UBound(lv, 1)
            If IsDate(lv(i, 3)) Then
                held(CLng(Int(CDbl(CDate(lv(i, 3)))))) = True
                If CDate(lv(i, 3)) > oldThrough And CDate(lv(i, 3)) <= rep.ToD And Trim$(SafeStr(lv(i, 2))) <> "" And rosterSet.Exists(Trim$(SafeStr(lv(i, 2)))) Then
                    key = Trim$(SafeStr(lv(i, 2))) & "|" & Trim$(SafeStr(lv(i, 4)))
                    If expected.Exists(key) Then expected(key) = expected(key) + NumOr0(lv(i, 5)) Else expected.Add key, NumOr0(lv(i, 5))
                End If
            End If
        Next i
    End If
    ' the new snapshot
    Dim n As Long: n = tot.Count
    Dim outA() As Variant: ReDim outA(1 To IIf(n > 0, n, 1), 1 To KT_NCOLS)
    Dim kk As Variant, r As Long, nm As String, code As String
    For Each kk In tot.Keys
        r = r + 1
        nm = Left$(CStr(kk), InStr(CStr(kk), "|") - 1): code = Mid$(CStr(kk), InStr(CStr(kk), "|") + 1)
        outA(r, 1) = SafeStr(payOf(nm)): outA(r, 2) = nm: outA(r, 3) = code: outA(r, 4) = tot(kk)
        outA(r, 5) = CDate(Int(CDbl(rep.ToD))): outA(r, 6) = stamp
    Next kk
    mIngested = mIngested + 1                     ' (review N2)
    ReplaceRows lo, outA, r, KT_NCOLS, "5", "6"
    RecordKronosFile srcName, "year to date taken", rep
    ' (review M1) the edits are dated against the OLD Through, so the log is written before the date moves
    ' v27.5 (item 2, Graeme Oct 25): a day between the Throughs the ledger does not hold is a GAP - a missed one-day
    ' report - so each difference is a "gap" row naming those days (counted in Worked Periods only until the person's
    ' rows for one of them arrive: GapCountsCore), not a timecard "edit" dated by the Transactions block
    mGapNote = "": mGapMsg = ""
    If oldThrough > 0 Then
        Dim gapTxt As String: gapTxt = GapDaysText(oldThrough, rep.ToD, KeysJoined(held))
        If gapTxt <> "" Then
            mGapNote = GAP_NOTE_PREFIX & gapTxt
            mGapMsg = GapDaysWords(gapTxt)
            WriteCumulative = LogChanges(expected, tot, payOf, rep.ToD, "gap", mGapNote, srcName, stamp, rep, oldThrough)
        Else
            WriteCumulative = LogChanges(expected, tot, payOf, rep.ToD, "edit", "differs from the ledger", srcName, stamp, rep, oldThrough)
        End If
    End If
    SetDateName KCUM_NAME, rep.ToD
End Function

' Every "name|code" whose hours differ between before and after into Worked Changes; the date is d, or for an "edit"
' the latest day the report's Transactions block names for that person and code on or before the old Through
' (with who edited it in the Note). Returns the rows written.
Private Function LogChanges(ByVal before As Object, ByVal after As Object, ByVal payOf As Object, ByVal d As Date, ByVal kind As String, _
                            ByVal note As String, ByVal srcName As String, ByVal stamp As Date, ByRef rep As KronosReport, ByVal oldThrough As Date) As Long
    Dim keys As Object: Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = 1
    Dim kk As Variant
    For Each kk In before.Keys
        keys(kk) = True
    Next kk
    For Each kk In after.Keys
        keys(kk) = True
    Next kk
    Dim outA() As Variant: ReDim outA(1 To IIf(keys.Count > 0, keys.Count, 1), 1 To WC_NCOLS)
    Dim n As Long, b As Double, a As Double, nm As String, code As String, m As Double, whenD As Date, who As String, i As Long, best As Date
    For Each kk In keys.Keys
        b = 0: a = 0
        If before.Exists(kk) Then b = before(kk)
        If after.Exists(kk) Then a = after(kk)
        If Abs(a - b) > 0.005 Then
            nm = Left$(CStr(kk), InStr(CStr(kk), "|") - 1): code = Mid$(CStr(kk), InStr(CStr(kk), "|") + 1)
            m = PremiumMultiplier(code, PREMIUM_CODES, PREMIUM_MULTIPLIERS)
            If m < 0 Then m = 0
            whenD = CDate(Int(CDbl(d))): who = note
            If kind = "edit" Then
                ' the latest day the Transactions block names for this person and code on or before the old Through
                best = 0
                For i = 1 To rep.nTx
                    If rep.TxId(i) <> "" And rep.TxId(i) = SafeStr(payOf(nm)) And StrComp(rep.TxCode(i), code, vbTextCompare) = 0 Then
                        If (rep.TxDate(i) <= oldThrough Or oldThrough = 0) And rep.TxDate(i) > best Then
                            best = rep.TxDate(i)
                            who = IIf(rep.TxBy(i) <> "", "timecard edited by " & rep.TxBy(i), "timecard edited")
                        End If
                    End If
                Next i
                If best > 0 Then whenD = CDate(Int(CDbl(best)))
            End If
            n = n + 1
            outA(n, 1) = stamp: outA(n, 2) = SafeStr(payOf(nm)): outA(n, 3) = nm: outA(n, 4) = whenD: outA(n, 5) = code
            outA(n, 6) = b: outA(n, 7) = a: outA(n, 8) = a - b: outA(n, 9) = (a - b) * m: outA(n, 10) = kind
            outA(n, 11) = who: outA(n, 12) = srcName: outA(n, 13) = CurrentUser()
        End If
    Next kk
    If n > 0 Then AppendRows WorkedChangesTable(True), outA, n, WC_NCOLS, "4", "1"
    LogChanges = n
End Function

' Worked (H) for every roster person the snapshot or the ledger knows, in the pivot's shape (Payroll #, Employee
' Name, Grand Total) so the matching in SyncKronosData runs unchanged: the snapshot's codes + the ledger days after
' its Through, each code's hours x its multiplier, rounded whole (WeightedTotal). through = the last day covered.
Public Function PivotFromLedger(ByRef through As Date) As Variant
    Dim cumThrough As Date: cumThrough = KronosCumThrough()
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    cumThrough = LedgerCumOf(cumThrough, yearFrom)            ' (review H3) a snapshot from before this roster year is not this year's
    through = cumThrough
    Dim spec As Object: Set spec = CreateObject("Scripting.Dictionary"): spec.CompareMode = 1   ' name -> the snapshot's "code=h;code=h" (its keys = the pivot's row order)
    Dim dayOf As Object: Set dayOf = CreateObject("Scripting.Dictionary"): dayOf.CompareMode = 1   ' name -> the run's days' "code=h;code=h"
    Dim payOf As Object: Set payOf = CreateObject("Scripting.Dictionary"): payOf.CompareMode = 1
    Dim days As Object: Set days = CreateObject("Scripting.Dictionary")
    Dim v As Variant, i As Long, nm As String, d As Date
    If cumThrough > 0 Then
        v = TableRows(KronosTotalsTable(False))
        If IsArray(v) Then
            For i = 1 To UBound(v, 1)
                nm = Trim$(SafeStr(v(i, 2)))
                If nm <> "" Then
                    spec(nm) = SafeStr(spec(nm)) & IIf(SafeStr(spec(nm)) = "", "", ";") & Trim$(SafeStr(v(i, 3))) & "=" & Trim$(Str$(NumOr0(v(i, 4))))   ' Str$: "." whatever the PC's separator
                    If SafeStr(payOf(nm)) = "" Then payOf(nm) = SafeStr(v(i, 1))
                End If
            Next i
        End If
    End If
    ' (review M7 / N6) the one-day reports count only along the UNBROKEN run after the Totals: a missed morning stops both
    ' the hours and the data-through date there (provisional OT keeps counting the shifts beyond it; IntegrityCheck names the gap)
    Dim lv As Variant: lv = TableRows(WorkedDaysTable(False))
    If IsArray(lv) Then
        For i = 1 To UBound(lv, 1)
            If IsDate(lv(i, 3)) Then days(CLng(Int(CDbl(CDate(lv(i, 3)))))) = True      ' every day the ledger holds: LedgerRunCore keeps those after the Totals
        Next i
    End If
    Dim kk As Variant, serials As String
    For Each kk In days.Keys
        serials = serials & IIf(serials = "", "", ";") & CStr(kk)
    Next kk
    through = LedgerRunCore(cumThrough, yearFrom, serials)
    If IsArray(lv) And cumThrough > 0 Then
        For i = 1 To UBound(lv, 1)
            If IsDate(lv(i, 3)) Then
                d = CDate(lv(i, 3))
                If LedgerDayIn(d, cumThrough, through) Then
                    nm = Trim$(SafeStr(lv(i, 2)))
                    If nm <> "" Then
                        If Not spec.Exists(nm) Then spec.Add nm, ""      ' a ledger-only person: a row of the pivot only with a day inside the run
                        dayOf(nm) = SafeStr(dayOf(nm)) & IIf(SafeStr(dayOf(nm)) = "", "", ";") & Trim$(SafeStr(lv(i, 4))) & "=" & Trim$(Str$(NumOr0(lv(i, 5))))
                        If SafeStr(payOf(nm)) = "" Then payOf(nm) = SafeStr(lv(i, 1))
                    End If
                End If
            End If
        Next i
    End If
    Dim out() As Variant: ReDim out(1 To spec.Count + 1, 1 To 3)
    out(1, 1) = "Payroll #": out(1, 2) = "Employee Name": out(1, 3) = "Grand Total"
    Dim n As Long, ds As String
    n = 1
    For Each kk In spec.Keys
        n = n + 1
        ds = ""
        If dayOf.Exists(kk) Then ds = CStr(dayOf(kk))
        out(n, 1) = SafeStr(payOf(kk)): out(n, 2) = CStr(kk): out(n, 3) = LedgerPersonCore(CStr(spec(kk)), ds)
    Next kk
    PivotFromLedger = out
End Function

' The Worked Periods table from the ledger: every snapshot day's weighted hours summed per roster person per pay
' period (exact from KronosLedgerFrom on; the days inside the cumulative's Through included - they are history, the
' snapshot only replaces them in H). Returns the line for the message.
Private Function RebuildWorkedPeriods(ByVal nameMap As Object, ByVal stamp As Date) As String
    Set mTxPeriods = CreateObject("Scripting.Dictionary"): mTxPeriods.CompareMode = 1
    Set mTxCodes = CreateObject("Scripting.Dictionary"): mTxCodes.CompareMode = 1
    Set mTxRaw = CreateObject("Scripting.Dictionary"): mTxRaw.CompareMode = 1
    Set mTxPay = CreateObject("Scripting.Dictionary"): mTxPay.CompareMode = 1
    mTxSkipped = 0
    Dim rosterOf As Object: Set rosterOf = CreateObject("Scripting.Dictionary"): rosterOf.CompareMode = 1
    Dim rosterSet As Object: Set rosterSet = CreateObject("Scripting.Dictionary"): rosterSet.CompareMode = 1
    Dim kk As Variant
    For Each kk In nameMap.Keys
        rosterSet(CStr(nameMap(kk))) = True
    Next kk
    Dim v As Variant: v = TableRows(WorkedDaysTable(False))
    Dim i As Long, nm As String, k As String, pk As String, w As Double, d As Date, hk As String
    Dim yearFrom As Date: yearFrom = RosterYearFrom()
    Dim lf As Date: lf = KronosLedgerFrom()
    Dim heldOf As Object: Set heldOf = CreateObject("Scripting.Dictionary"): heldOf.CompareMode = 1   ' v27.5 (item 2): "namekey|code" -> the days this person has ledger rows on for that pay code (";" serials; review: per code, so a backfill sets aside only its own code's gap row)
    Dim dayStamp As Object: Set dayStamp = CreateObject("Scripting.Dictionary")   ' (review N5 / R2): day -> when it was last RE-imported (a "day" change row)
    Dim cv As Variant: cv = TableRows(WorkedChangesTable(False))
    Dim ck As Long
    If IsArray(cv) Then
        For i = 1 To UBound(cv, 1)
            If LCase$(Trim$(SafeStr(cv(i, 10)))) = "day" And IsDate(cv(i, 4)) And IsDate(cv(i, 1)) Then
                ck = CLng(Int(CDbl(CDate(cv(i, 4)))))
                If Not dayStamp.Exists(ck) Then
                    dayStamp.Add ck, CDbl(CDate(cv(i, 1)))
                ElseIf CDbl(CDate(cv(i, 1))) > dayStamp(ck) Then
                    dayStamp(ck) = CDbl(CDate(cv(i, 1)))
                End If
            End If
        Next i
    End If
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, 2)))
            If nm <> "" And IsDate(v(i, 3)) Then
                hk = NameKey(nm) & "|" & Trim$(SafeStr(v(i, 4)))                                                 ' v27.5 (item 2)
                heldOf(hk) = SafeStr(heldOf(hk)) & ";" & CStr(CLng(Int(CDbl(CDate(v(i, 3))))))
                If rosterSet.Exists(nm) And CDate(v(i, 3)) >= yearFrom Then                 ' roster people only, this year only (a Former person's rows stay in the ledger)
                    k = NameKey(nm)
                    If Not rosterOf.Exists(k) Then rosterOf.Add k, nm: mTxRaw(k) = nm: mTxPay(k) = SafeStr(v(i, 1))
                    d = CDate(v(i, 3))
                    pk = k & "|" & CLng(Int(CDbl(PeriodStartFor(d))))
                    w = NumOr0(v(i, 6))
                    If mTxPeriods.Exists(pk) Then mTxPeriods(pk) = mTxPeriods(pk) + w Else mTxPeriods.Add pk, w
                    If w <> 0 Then mTxCodes(pk) = SafeStr(mTxCodes(pk)) & IIf(SafeStr(mTxCodes(pk)) = "", "", "; ") & Format$(d, "mmm d") & " " & Trim$(SafeStr(v(i, 4))) & " " & HoursText(NumOr0(v(i, 5)))
                End If
            End If
        Next i
    End If
    ' (review M9) the edits a year-to-date report revealed belong to the period of the day they name. v27.5 (item 2): a
    ' "gap" row (the difference across missed one-day reports, dated at that report's Through) counts while the person
    ' has no ledger row on any day its Note names (GapCountsCore) - the day's own rows take over once it is imported
    v = TableRows(WorkedChangesTable(False))
    Dim reAt As Variant, kindTxt As String, counts As Boolean
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            nm = Trim$(SafeStr(v(i, 3)))
            kindTxt = LCase$(Trim$(SafeStr(v(i, 10))))
            If (kindTxt = "edit" Or kindTxt = "gap") And nm <> "" And IsDate(v(i, 4)) Then
                d = CDate(v(i, 4))
                k = NameKey(nm)
                If kindTxt = "gap" Then
                    counts = rosterSet.Exists(nm) And GapCountsCore(d, yearFrom, lf, GapNoteDays(SafeStr(v(i, 11))), SafeStr(heldOf(k & "|" & Trim$(SafeStr(v(i, 5))))))
                Else
                    ' (review N4 / N5) this year, from the first one-day report on, and not a day re-imported AFTER the edit was logged
                    ' (that re-import already holds the edited figure)
                    reAt = Empty
                    ck = CLng(Int(CDbl(d)))
                    If dayStamp.Exists(ck) Then reAt = dayStamp(ck)
                    counts = rosterSet.Exists(nm) And EditCountsCore(d, yearFrom, lf, reAt, v(i, 1))
                End If
                If counts Then
                    If Not rosterOf.Exists(k) Then rosterOf.Add k, nm: mTxRaw(k) = nm: mTxPay(k) = SafeStr(v(i, 2))
                    pk = k & "|" & CLng(Int(CDbl(PeriodStartFor(d))))
                    w = NumOr0(v(i, 9))
                    If mTxPeriods.Exists(pk) Then mTxPeriods(pk) = mTxPeriods(pk) + w Else mTxPeriods.Add pk, w
                    If w <> 0 Then mTxCodes(pk) = SafeStr(mTxCodes(pk)) & IIf(SafeStr(mTxCodes(pk)) = "", "", "; ") & Format$(d, "mmm d") & " " & kindTxt & " " & Trim$(SafeStr(v(i, 5))) & " " & HoursText(NumOr0(v(i, 8)))
                End If
            End If
        Next i
    End If
    RebuildWorkedPeriods = WriteWorkedPeriods(rosterOf, stamp)
    If lf > 0 Then RebuildWorkedPeriods = RebuildWorkedPeriods & " Exact from " & Format$(lf, "mmm d, yyyy") & " (the first one-day report); earlier periods are not available."
End Function

' name -> "Refused|sum of the pay-period cells": what a Kronos import must leave untouched (RosterTotalsMap without Total,
' which H moves by design)
Private Function RosterRefusedMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    Dim m As Object: Set m = RosterTotalsMap()
    Dim k As Variant, s As String
    For Each k In m.Keys
        s = CStr(m(k))
        d(k) = Mid$(s, InStr(s, "|") + 1)
    Next k
    Set RosterRefusedMap = d
End Function

' TotalsDifference for the Refused map: names whose Refused or period sum moved, or that came or went ("" = nothing)
Private Function TotalsDifferenceRef(ByVal before As Object, ByVal after As Object) As String
    Dim k As Variant, s As String, n As Long
    For Each k In before.Keys
        If Not after.Exists(k) Then
            n = n + 1: If n <= 10 Then s = s & "  " & k & ": no longer on the roster" & vbNewLine
        ElseIf before(k) <> after(k) Then
            n = n + 1: If n <= 10 Then s = s & "  " & k & ": Refused / periods were " & Replace(before(k), "|", " / ") & ", now " & Replace(after(k), "|", " / ") & vbNewLine
        End If
    Next k
    For Each k In after.Keys
        If Not before.Exists(k) Then n = n + 1: If n <= 10 Then s = s & "  " & k & ": appeared on the roster" & vbNewLine
    Next k
    If n > 10 Then s = s & "  ...and " & (n - 10) & " more" & vbNewLine
    TotalsDifferenceRef = s
End Function

' The Worked Periods sheet's table, made on first use (admin-only, never published as a sheet)
Public Function WorkedPeriodsTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(WP_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(WP_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(WP_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = WP_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, WP_NCOLS).Value = Array("Payroll #", "Employee", "Period Start", "Hours", "Codes", "Imported At")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, WP_NCOLS), , xlYes)
        lo.Name = WP_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns("A").ColumnWidth = 11: ws.Columns("B").ColumnWidth = 24: ws.Columns("C").ColumnWidth = 13
        ws.Columns("D").ColumnWidth = 9: ws.Columns("E").ColumnWidth = 60: ws.Columns("F").ColumnWidth = 17
        ProtectSheet ws
    End If
    Set WorkedPeriodsTable = lo
End Function

' Rebuilds the table from this import's per-period sums: one row per person per pay period, the
' roster's spelling where the name matched (rosterOf), the export's otherwise. Returns the message line.
Private Function WriteWorkedPeriods(ByVal rosterOf As Object, ByVal stamp As Date) As String
    If mTxPeriods Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = WorkedPeriodsTable(True)
    If lo Is Nothing Then WriteWorkedPeriods = "The Worked Periods sheet couldn't be made.": Exit Function
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim n As Long: n = mTxPeriods.Count
    Dim outA() As Variant, kk As Variant, p As Long, k As String, nm As String, i As Long, people As Object
    Set people = CreateObject("Scripting.Dictionary"): people.CompareMode = 1
    If n > 0 Then
        ReDim outA(1 To n, 1 To WP_NCOLS)
        For Each kk In mTxPeriods.Keys
            p = InStr(CStr(kk), "|")
            k = Left$(CStr(kk), p - 1)
            If rosterOf.Exists(k) Then                       ' roster people only: a name not on the roster (or Former) is not kept
                nm = CStr(rosterOf(k))
                i = i + 1
                people(nm) = True
                outA(i, 1) = SafeStr(mTxPay(k))
                outA(i, 2) = nm
                outA(i, 3) = CDate(CDbl(Mid$(CStr(kk), p + 1)))
                outA(i, 4) = mTxPeriods(kk)
                outA(i, 5) = SafeStr(mTxCodes(kk))
                outA(i, 6) = stamp
            End If
        Next kk
        n = i
    End If
    Dim em As String
    On Error GoTo Fail
    UnlockSheet ws
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If n > 0 Then
        lo.Resize lo.HeaderRowRange.Resize(n + 1)
        With lo.HeaderRowRange.Offset(1, 0).Resize(n, WP_NCOLS)
            .Columns(1).NumberFormat = "@"
            .Value = outA
            .Columns(3).NumberFormat = "mmm d, yyyy"
            .Columns(6).NumberFormat = "mmm d, yyyy h:mm"
        End With
    End If
    ProtectSheet ws
    WriteWorkedPeriods = "Premium hours per pay period: " & people.Count & " people, " & n & " person-periods (Worked Periods sheet)." & _
                         IIf(mTxSkipped > 0, " " & mTxSkipped & " export row(s) with no pay code or no date were not counted.", "")
    Exit Function
Fail:
    em = Err.Description
    LogError "WriteWorkedPeriods", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    WriteWorkedPeriods = "The per-period hours could NOT be written: " & em
End Function

' One person's premium hours per pay period, period serial -> hours (empty when none / no table)
Public Function WorkedPeriodsFor(ByVal emp As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set WorkedPeriodsFor = d
    Dim lo As ListObject: Set lo = WorkedPeriodsTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, ps As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 2))), Trim$(emp), vbTextCompare) = 0 And IsDate(v(i, 3)) Then
            ps = CLng(Int(CDbl(CDate(v(i, 3)))))
            If d.Exists(ps) Then d(ps) = d(ps) + NumOr0(v(i, 4)) Else d.Add ps, NumOr0(v(i, 4))
        End If
    Next i
End Function

' Every row as Employee, Period Start, Hours (the members' dWorked); Empty when there is none
Public Function WorkedPeriodRows() As Variant
    Dim lo As ListObject: Set lo = WorkedPeriodsTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim out() As Variant: ReDim out(1 To UBound(v, 1) + 1, 1 To 3)
    out(1, 1) = "Employee": out(1, 2) = "Period Start": out(1, 3) = "Hours"
    Dim i As Long
    For i = 1 To UBound(v, 1)
        out(i + 1, 1) = SafeStr(v(i, 2)): out(i + 1, 2) = v(i, 3): out(i + 1, 3) = NumOr0(v(i, 4))
    Next i
    WorkedPeriodRows = out
End Function

' After an error part-way through the import: the roster is protected again and
' the Kronos file is closed if the import opened it. Quiet.
Private Sub KronosCleanUp(ByVal wsDest As Worksheet, ByVal wbSource As Workbook, ByVal openedHere As Boolean)
    On Error Resume Next                            ' never raises
    If Not wsDest Is Nothing Then ProtectSheet wsDest
    If openedHere And Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    ThisWorkbook.Activate
End Sub


' ====================================================================
'  IMPORT DATE / STALE CHECK (v20)
' ====================================================================
Public Sub SetDateName(ByVal nmName As String, ByVal d As Date)     ' v26.5: Public (RefusedThrough uses it too)
    Dim ref As String: ref = "=" & Replace(Format$(CDbl(d), "0.00000"), ",", ".")
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(nmName)
    On Error GoTo 0
    On Error Resume Next
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=nmName, RefersTo:=ref, Visible:=False
    Else
        nmObj.RefersTo = ref                  ' in place: a formula that reads the name never sees it missing
        nmObj.Visible = False
    End If
    On Error GoTo 0
End Sub

Public Function GetDateName(ByVal nmName As String) As Date
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(nmName)
    On Error GoTo 0
    If nmObj Is Nothing Then Exit Function
    Dim x As Double: x = Val(Mid$(nmObj.RefersTo, 2))
    If x > 0 Then GetDateName = CDate(x)
End Function

' ---- waiting for the new year's Kronos data (v23) ----
' Set by RolloverYear to Jan 1 of the new year: Worked (H) is 0 and an
' import is refused until the file's data runs through that date.
Public Function KronosAwaitFrom() As Date
    KronosAwaitFrom = GetDateName(AWAIT_NAME)
End Function

Public Sub SetKronosAwait(ByVal d As Date)
    SetDateName AWAIT_NAME, d
End Sub

Public Sub ClearKronosAwait(Optional ByVal fromCode As Boolean = True)     ' the argument keeps it out of the Alt+F8 list
    On Error Resume Next
    ThisWorkbook.Names(AWAIT_NAME).Delete
    On Error GoTo 0
End Sub

' When the Kronos data was saved (falls back to when it was imported); 0 = never
Public Function KronosDataDate() As Date
    KronosDataDate = GetDateName("KronosFileDate")
    If KronosDataDate = 0 Then KronosDataDate = GetDateName("KronosImportedAt")
End Function

Public Function KronosImportedAt() As Date
    KronosImportedAt = GetDateName("KronosImportedAt")
End Function

' ---- data-through date (v21) ----
' The Kronos data runs through the day before the file was saved.
Public Function DataThroughFor(ByVal fileDate As Date) As Date
    DataThroughFor = Int(fileDate) - 1
End Function

' Last date the imported Kronos data covers (0 = not known)
Public Function KronosDataThrough() As Date
    KronosDataThrough = GetDateName(PROV_NAME)
End Function

' Makes sure the hidden name the provisional formulas read exists, and
' returns its value. Before the first v21 import it is worked out from the
' v20 Kronos file date; with no import at all, yesterday (so nothing that
' has already ended is counted twice). Safe to call any time.
Public Function EnsureDataThrough() As Date
    Dim d As Date: d = KronosDataThrough()
    If d = 0 Then
        If KronosDataDate() > 0 Then
            d = DataThroughFor(KronosDataDate())
        Else
            d = Date - 1
        End If
        If Not ThisWorkbook.ReadOnly Then SetDateName PROV_NAME, d
    End If
    EnsureDataThrough = d
End Function

' Status-line / SystemCheck text, e.g. "Provisional OT: Kronos data through Oct 9"
Public Function DataThroughText() As String
    Dim d As Date: d = KronosDataThrough()
    If d = 0 Then
        DataThroughText = "Provisional OT: data-through date not set (run Import Kronos)"
    Else
        DataThroughText = "Provisional OT: Kronos data through " & Format$(d, "mmm d")
    End If
End Function

' Days since the Kronos data was saved; -1 = never imported (v20)
Public Function KronosAgeDays() As Long
    Dim d As Date: d = KronosDataDate()
    If d = 0 Then KronosAgeDays = -1 Else KronosAgeDays = Int(Date) - Int(d)
End Function

Public Function KronosIsStale() As Boolean
    Dim aw As Date: aw = KronosAwaitFrom()       ' v23: after a rollover the clock starts on Jan 1
    If aw > 0 Then KronosIsStale = (Int(Date) > Int(aw) + KronosStaleDays()): Exit Function
    Dim a As Long: a = KronosAgeDays()
    KronosIsStale = (a < 0 Or a > KronosStaleDays())
End Function

' Status line text, e.g. "Kronos: Oct 7 (2 days ago)"
Public Function KronosStatusText() As String
    Dim aw As Date: aw = KronosAwaitFrom()
    If aw > 0 Then KronosStatusText = "Kronos: awaiting " & Year(aw) & " data (Worked = 0)": Exit Function
    Dim a As Long: a = KronosAgeDays()
    If a < 0 Then KronosStatusText = "Kronos: not imported since v20": Exit Function
    Dim t As String
    Select Case a
        Case 0: t = "today"
        Case 1: t = "yesterday"
        Case Else: t = a & " days ago"
    End Select
    KronosStatusText = "Kronos: " & Format$(KronosDataDate(), "mmm d") & " (" & t & ")" & _
                       IIf(a > KronosStaleDays(), " - import", "")
End Function

' v26.5.5 (B5): the one red line a supervisor's pop-up shows when the worked hours are old ("" when fresh)
Public Function KronosStaleLine() As String
    If Not KronosIsStale() Then Exit Function
    Dim a As Long: a = KronosAgeDays()
    If a < 0 Then
        KronosStaleLine = "Worked hours have not been imported from Kronos since the install: the call order may be off. Tell " & _
                          PRIMARY_ADMIN_NAME & " to import Kronos."
    Else
        KronosStaleLine = "Worked hours are " & a & " days old: the call order may be off. Tell " & PRIMARY_ADMIN_NAME & " to import Kronos."
    End If
End Function

' Call-out / posting warning. True = go ahead. v26.5.5 (B5): a supervisor can't import, so they get no
' box here - the calling pop-up shows KronosStaleLine in red instead; an admin keeps the Yes / No, Yes the default.
Public Function KronosOkToContinue(ByVal what As String) As Boolean
    KronosOkToContinue = True
    If Not KronosIsStale() Then Exit Function
    If Not IsAdmin() Then Exit Function
    Dim a As Long: a = KronosAgeDays()
    Dim msg As String
    If a < 0 Then
        msg = "Worked hours haven't been imported from Kronos since this update was installed, " & _
              "so there's no record of how current they are."
    Else
        msg = "Worked hours come from Kronos data through " & Format$(KronosDataThrough(), "ddd mmm d") & _
              " (" & a & " days old: the age counts from the data's own end)."
    End If
    msg = msg & vbNewLine & vbNewLine & "Call order is only as current as that import. Continue with the " & what & " anyway?" & _
          vbNewLine & vbNewLine & "(No = stop here and run Import Kronos first.)"
    KronosOkToContinue = (MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton1, "Kronos data is old") = vbYes)
End Function


' ====================================================================
'  HOURS HISTORY (v20) - one dated row per person per import
' ====================================================================
Public Function HoursHistoryTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(HH_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(HH_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(HH_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            UnlockStructure                                 ' v25: structure protection
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = HH_SHEET
            RelockStructure
            SetTabColor ws, RGB(96, 125, 139)           ' v25.1
        End If
        UnlockSheet ws
        ws.Range("A1:H1").Value = Array("Import ID", "Imported At", "Kronos File Date", "Employee", "Worked Hours", "Imported By", "Data Through", "Year")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H2"), , xlYes)
        lo.Name = HH_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns("A").ColumnWidth = 10
        ws.Columns("B:C").ColumnWidth = 17
        ws.Columns("D").ColumnWidth = 24
        ws.Columns("E").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 12
        ws.Columns("G").ColumnWidth = 13
        ws.Columns("H").ColumnWidth = 7
        ProtectSheet ws
    End If
    If Not lo Is Nothing Then EnsureHoursHistoryColumns lo
    Set HoursHistoryTable = lo
End Function

' A v20 table has 6 columns: "Data Through" (v21) and "Year" (v23) go on the end, once
Private Sub EnsureHoursHistoryColumns(ByVal lo As ListObject)
    If lo.ListColumns.Count >= HH_NCOLS Then Exit Sub
    If ThisWorkbook.ReadOnly Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    On Error Resume Next
    UnlockSheet ws
    If lo.ListColumns.Count < 7 Then
        lo.ListColumns.Add.Name = "Data Through"
        ws.Columns(7).ColumnWidth = 13
    End If
    If lo.ListColumns.Count < 8 Then
        lo.ListColumns.Add.Name = "Year"
        ws.Columns(8).ColumnWidth = 7
    End If
    ProtectSheet ws
    On Error GoTo 0
End Sub

' Gives every snapshot that has no Year yet the year y (the rollover stamps
' the year it is closing, so a late-December snapshot is never read as the
' new year's hours). Returns how many rows were stamped.
Public Function StampHoursHistoryYear(ByVal y As Long) As Long
    Dim lo As ListObject: Set lo = HoursHistoryTable(False)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < HH_NCOLS Then Err.Raise vbObjectError + 750, , "the Hours History sheet couldn't be given its Year column"
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim yv As Variant: yv = RangeTo2D(lo.ListColumns(HH_NCOLS).DataBodyRange)
    Dim idv As Variant: idv = RangeTo2D(lo.ListColumns(1).DataBodyRange)
    Dim i As Long, n As Long
    For i = 1 To UBound(yv, 1)
        If Trim$(SafeStr(yv(i, 1))) = "" And Trim$(SafeStr(idv(i, 1))) <> "" Then
            yv(i, 1) = y
            n = n + 1
        End If
    Next i
    If n > 0 Then
        UnlockSheet lo.Parent
        lo.ListColumns(HH_NCOLS).DataBodyRange.Value = yv
        ProtectSheet lo.Parent
    End If
    StampHoursHistoryYear = n
End Function

' Appends the snapshot; returns its Import ID ("KR-0007") or "" on failure
Private Function AppendHoursHistory(ByVal stamp As Date, ByVal fileDate As Date, ByVal through As Date, _
                                    histName() As String, histHrs() As Double, ByVal n As Long) As String
    Dim lo As ListObject: Set lo = HoursHistoryTable(True)
    If lo Is Nothing Then Exit Function
    If lo.ListColumns.Count < HH_NCOLS Then
        MsgBox "Worked hours were imported, but the Hours History sheet couldn't be given its Data Through / Year columns." & _
               vbNewLine & "Is the file read-only?", vbExclamation, "Call Out Roster"
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = lo.Parent

    ' next Import ID
    Dim mx As Long, v As Variant, i As Long, cnt As Long, blankFirst As Boolean
    cnt = lo.ListRows.Count
    If Not lo.DataBodyRange Is Nothing Then
        If cnt = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
            blankFirst = True: cnt = 0
        Else
            v = RangeTo2D(lo.ListColumns(1).DataBodyRange)
            For i = 1 To UBound(v, 1)
                If Left$(SafeStr(v(i, 1)), 3) = "KR-" Then
                    If Val(Mid$(SafeStr(v(i, 1)), 4)) > mx Then mx = Val(Mid$(SafeStr(v(i, 1)), 4))
                End If
            Next i
        End If
    End If
    Dim id As String: id = "KR-" & Format$(mx + 1, "0000")

    Dim outA() As Variant: ReDim outA(1 To n, 1 To HH_NCOLS)
    Dim yrNow As Long
    On Error Resume Next
    yrNow = RosterYear()
    On Error GoTo 0
    For i = 1 To n
        outA(i, 1) = id
        outA(i, 2) = stamp
        outA(i, 3) = fileDate
        outA(i, 4) = histName(i)
        outA(i, 5) = histHrs(i)
        outA(i, 6) = CurrentUser()
        outA(i, 7) = through
        If yrNow > 0 Then outA(i, 8) = yrNow
    Next i

    Dim em As String
    On Error GoTo Fail
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    With lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, HH_NCOLS)
        .Value = outA
        .Columns(2).NumberFormat = "mmm d, yyyy h:mm"
        .Columns(3).NumberFormat = "mmm d, yyyy h:mm"
        .Columns(7).NumberFormat = "mmm d, yyyy"
    End With
    ProtectSheet ws
    AppendHoursHistory = id
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendHoursHistory", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Worked hours were imported, but the Hours History snapshot couldn't be saved:" & vbNewLine & em, vbExclamation, "Call Out Roster"
End Function

' v24.3.1: the column whose heading (first 10 rows) says `want` by its letters and digits
' (IsHeaderMatch in modOvertimeEntry: "Employee  Name" and "EMPLOYEE NAME" match, "Name" doesn't).
' 0 = none; hdrRow = the row it was found on.
Private Function HeaderColumn(ByRef src As Variant, ByRef hdrRow As Long, ByVal want As String) As Long
    Dim r As Long, c As Long, rMax As Long
    hdrRow = 0
    rMax = UBound(src, 1)
    If rMax > 10 Then rMax = 10
    For r = 1 To rMax
        For c = 1 To UBound(src, 2)
            If VarType(src(r, c)) = vbString Then
                If IsHeaderMatch(CStr(src(r, c)), want) Then
                    HeaderColumn = c
                    hdrRow = r
                    Exit Function
                End If
            End If
        Next c
    Next r
End Function

' The column of the Kronos sheet whose heading says Payroll # (0 = none), on the heading row hdrRow
' (v24.3.1: that row only - before, the first match in the first 10 rows).
Private Function PayrollColumn(ByRef src As Variant, ByVal hdrRow As Long) As Long
    Dim c As Long
    If hdrRow < 1 Or hdrRow > UBound(src, 1) Then Exit Function
    For c = 1 To UBound(src, 2)
        If VarType(src(hdrRow, c)) = vbString Then
            If IsPayrollHeader(CStr(src(hdrRow, c))) Then
                PayrollColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

Public Function NameKey(ByVal s As String) As String
    NameKey = LCase$(Replace(Replace(Trim$(s), " ", ""), ".", ""))
End Function

' exact match, or the roster name + a middle initial in Kronos ("Dobbin,Stephen A")
Private Function FindKey(ByVal dict As Object, ByVal alt As Object, ByVal k As String) As String
    If k = "" Then Exit Function
    If dict.Exists(k) Then FindKey = k: Exit Function
    If alt.Exists(k) Then FindKey = CStr(alt(k))
End Function
