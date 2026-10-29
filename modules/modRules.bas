Attribute VB_Name = "modRules"
Option Explicit

' ====================================================================
'  RULES AND SETTINGS (v27, item 22 - briefs/v27.md 1.3)
'
'  Two admin sheets, nothing published.
'  Rules: one row per rule the roster applies - ID, the rule in one
'  sentence, the SelfTest suite that proves it, the release it came in
'  (all from RuleList below, rewritten by DrawRules), and a Basis column
'  an admin TYPES (the CBA article, or "agreed <date> by <who>"); the
'  only typed cells. RuleText(id) gives the sentence to anything that
'  wants to name a rule (the receipts' "on paper" line, item 6).
'  Settings: one row per OVERRIDE of a tunable setting - Name, Value,
'  Effective, Changed By, Previous, Note - written only by SetSetting
'  (admin, Alt+F8); the code reads a setting through its typed wrapper
'  (KronosStaleDays, WeekCap, MembersStaleDays, SignupStaleDays,
'  PaybackWarnDays, PostingCloseDays), each = SettingNum(name, default)
'  with modOvertimeEntry's constant as the default. A rule's own numbers
'  (the 16 h rule, the charge rate) are NOT tunable here: a rule change
'  is a release. SystemCheck lists the overrides; IntegrityCheck flags a
'  row whose name is not tunable or whose value is not a number.
' ====================================================================

Public Const RULES_VERSION As String = "Rules v29.1 (Oct 28)"
Public Const RULES_SHEET As String = "Rules"
Public Const RULES_TABLE As String = "tblRules"
Public Const SETTINGS_SHEET As String = "Settings"
Public Const SETTINGS_TABLE As String = "tblSettings"
Private Const RL_N As Long = 5                   ' ID, Rule, Basis (typed), Suite, Since
Private Const ST_N As Long = 6                   ' Name, Value, Effective, Changed By, Previous, Note
' the settings an admin may override on the sheet (name=default), in the order SetSetting lists them
Private Const TUNABLE As String = "KRONOS_STALE_DAYS,WEEK_MAX_HOURS,MEMBERS_STALE_DAYS,SIGNUP_STALE_DAYS,PAYBACK_WARN_DAYS,POSTING_CLOSE_DAYS"

' ---- the rules, one line each: ID|the sentence|suite|since ----
Private Function RuleList() As Variant
    RuleList = Array( _
        "CO-1|Call order: the grade code in the unit's column, then the lowest Total Hours, then the earliest plant seniority, then the name.|rotation|v20", _
        "CO-2|Everyone with a code in the unit, any grade, is listed when free that shift; grades above " & CHARGE_MAX_GRADE & " are called but never charged.|chargewhy|v21", _
        "EL-1|Not free: scheduled on that shift (the crew of that date), coming off nights that morning, or starting days the next morning.|rotation|v20", _
        "EL-2|Not free: booked off, sick, or inside the 24 h after a sick shift ends.|attendocc|v22", _
        "EL-3|Not free: working a shift switch that day, or gave the shift up in one (a shift given up relieves the next day's rest rule).|fatigue|v25", _
        "EL-4|The 16 h rule: the work that ends when the offer starts plus the offer may not pass " & FATIGUE_MAX_HOURS & " h, and " & FATIGUE_MIN_REST & " h must remain before the next work; a capped person is offered their maximum. Every switch must fit on the shift worked - a scheduled shift beside it counts - and the shift given up is not worked; the same-day sister swap of both shifts (the only swap an A/B pair can make) is exempt from the 16 h arithmetic, stopped only by OT already held on the shift (v27.5, Graeme Oct 25).|fatigue|v26.5.5", _
        "EL-5|OT restricted by an admin: not listed between the dates, never charged.|-|v27", _
        "WK-1|The weekly cap is OFF: the 48 h is a limit on scheduled work and overtime is requested (Graeme, Oct 22). An admin may turn a warning on with a number on the Settings sheet: then an acceptance or award that takes a person's overtime past it in the 7 days before the shift asks for a reason; it never blocks.|weekcap|v27", _
        "CH-1|A call-out refusal or a message left is charged the hours offered x " & CHARGE_MULTIPLIER & " (grade " & CHARGE_MAX_GRADE & " and below).|calloffer|v20", _
        "CH-2|A partial acceptance closes the call-out shift: nobody logged after it is offered the rest, nobody is charged for it.|calloffer|v29.1", _
        "CH-3|A posting award charges everyone ahead of the last awarded person in call order who did not sign up, at the posting's full hours x " & CHARGE_MULTIPLIER & ".|awardorder|v20", _
        "CH-4|No takers charges everyone eligible on the list; an award out of department charges everyone eligible in the unit who did not sign up; a cancel charges nobody.|awardcover|v20", _
        "CH-5|A full-shift sign-up beats a partial one; a partial is awarded only when no full-shift sign-up is left; uncovered hours are dropped, uncharged.|awardorder|v24", _
        "CH-6|Someone who moves unit or leaves is never charged on an open posting for a shift on or after that date.|plancode|v24", _
        "TH-1|Total Hours = Worked (Kronos) + Refused (charged hours in the pay periods up to the current one) + provisional accepted OT not yet in Kronos.|refsplit|v26.5", _
        "TH-2|Worked (Kronos) = each pay code's hours x its multiplier, rounded whole per person (OVERTIME 1.75, OT - DT Split 2, LIEU OT 1.5, LIEU DT 2, LIEU OT - DT Split 1.5, LIEU ST 1, CALLIN OT 1.75, MIN 4 1.5, STATWK 2.0 2).|kweight|v27", _
        "TH-3|DST nights charge actual hours (13 h in the fall, 11 h in the spring).|shiftdate|v21", _
        "PP-1|A leave return adds the home unit's average Total Hours gained while away (grade " & CHARGE_MAX_GRADE & " and below, not on leave); a leave over the year reset tops up from the reset.|newhire|v22", _
        "PP-2|A new hire starts at the home unit's average; a unit move raises to the highest in the new unit among grade " & CHARGE_MAX_GRADE & " and below, never lowers.|unitmove|v24", _
        "PP-3|A planned unit move lists the mover in the new unit from its date at the hours they would have after the top-up (an estimate for order only).|plancode|v24.3", _
        "AT-1|Sick time: 5 paid sick days a year in up to 3 occurrences, 5 unpaid personal days; over 5 combined = no paid sick until a year after the 2nd occurrence.|attendocc|v22.1", _
        "YR-1|Totals reset each year; pay periods are biweekly from the year's first Sunday start; the year is the roster's one cell.|snapyear|v23", _
        "HL-1|Plant holidays: DAYS workers are off and free; the rotating crews stay scheduled; a weekend holiday uses the plant's observed weekday.|rotation|v20")
End Function

' The sentence for a rule id ("" = unknown id)
Public Function RuleText(ByVal id As String) As String
    Dim v As Variant, k As Variant, p As Variant
    v = RuleList()
    For Each k In v
        p = Split(CStr(k), "|")
        If StrComp(CStr(p(0)), Trim$(id), vbTextCompare) = 0 Then RuleText = CStr(p(1)): Exit Function
    Next k
End Function

' ---- the Rules sheet ----
Public Function RulesTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(RULES_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(RULES_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(RULES_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = RULES_SHEET
            RelockStructure
            SetTabColor ws, RGB(30, 56, 100)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, RL_N).Value = Array("ID", "Rule", "Basis (CBA article, or agreed <date> by <who>) - TYPE HERE", "Suite", "Since")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, RL_N), , xlYes)
        lo.Name = RULES_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 7: ws.Columns(2).ColumnWidth = 110: ws.Columns(3).ColumnWidth = 44
        ws.Columns(4).ColumnWidth = 11: ws.Columns(5).ColumnWidth = 8
        ws.Columns(2).WrapText = True: ws.Columns(3).WrapText = True
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set RulesTable = lo
End Function

' Admin (Alt+F8, the Desk's FINISH band): the sheet made or refreshed - ID / Rule / Suite / Since from the code,
' a row appended for a new id, the typed Basis kept; then shown
Public Sub DrawRules()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim lo As ListObject: Set lo = RulesTable(True)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1
    Dim v As Variant, i As Long, cnt As Long
    If Not lo.DataBodyRange Is Nothing Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then
            v = RangeTo2D(lo.DataBodyRange)
            cnt = UBound(v, 1)
            For i = 1 To cnt
                If Trim$(SafeStr(v(i, 1))) <> "" Then have(Trim$(SafeStr(v(i, 1)))) = i
            Next i
        End If
    End If
    Dim rules As Variant: rules = RuleList()
    Dim k As Variant, p As Variant, r As Long, nNew As Long, nUpd As Long
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    UnlockSheet ws
    For Each k In rules
        p = Split(CStr(k), "|")
        If have.Exists(CStr(p(0))) Then
            r = have(CStr(p(0)))
            lo.DataBodyRange.Cells(r, 2).Value = CStr(p(1))
            lo.DataBodyRange.Cells(r, 4).Value = CStr(p(2))
            lo.DataBodyRange.Cells(r, 5).Value = CStr(p(3))
            nUpd = nUpd + 1
        Else
            cnt = cnt + 1
            lo.Resize lo.HeaderRowRange.Resize(cnt + 1)
            lo.HeaderRowRange.Offset(cnt, 0).Resize(1, RL_N).Value = Array(CStr(p(0)), CStr(p(1)), "", CStr(p(2)), CStr(p(3)))
            nNew = nNew + 1
        End If
    Next k
    ProtectSheet ws
    LockRulesCells True
    Application.EnableEvents = evWas
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
    MsgBox "Rules sheet: " & nUpd & " rule(s) refreshed, " & nNew & " added. Type the Basis of each (the CBA article, or ""agreed <date> by <who>""): that column is yours, the rest is rewritten from the code.", vbInformation, "Rules"
    Exit Sub
Oops:
    ReportError "DrawRules"
End Sub

' v25 (R4): the Basis column is the only typed-into cell on the sheet (ApplyViewCore calls this)
Public Sub LockRulesCells(ByVal adminView As Boolean)
    Dim lo As ListObject: Set lo = RulesTable(False)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    UnlockSheet ws
    ws.Cells.Locked = True
    If adminView Then
        On Error Resume Next
        If Not lo.DataBodyRange Is Nothing Then lo.ListColumns(3).DataBodyRange.Locked = False
        On Error GoTo 0
    End If
    ProtectSheet ws
End Sub

' ---- the Settings sheet ----
Public Function SettingsTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(SETTINGS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(SETTINGS_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(SETTINGS_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = SETTINGS_SHEET
            RelockStructure
            SetTabColor ws, RGB(30, 56, 100)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, ST_N).Value = Array("Name", "Value", "Effective", "Changed By", "Previous", "Note")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, ST_N), , xlYes)
        lo.Name = SETTINGS_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(1).ColumnWidth = 24: ws.Columns(3).ColumnWidth = 16: ws.Columns(6).ColumnWidth = 50
        ws.Columns(3).NumberFormat = "ddd mmm d, yyyy"
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set SettingsTable = lo
End Function

' True when the name is one an admin may override here
Public Function IsTunable(ByVal nm As String) As Boolean
    IsTunable = (InStr(1, "," & TUNABLE & ",", "," & UCase$(Trim$(nm)) & ",", vbTextCompare) > 0)
End Function

' The numeric override for a setting, or the default. Never raises; a non-numeric cell = the default.
Public Function SettingNum(ByVal nm As String, ByVal dflt As Double) As Double
    SettingNum = dflt
    On Error Resume Next
    Dim lo As ListObject: Set lo = SettingsTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), Trim$(nm), vbTextCompare) = 0 Then
            If IsNumeric(v(i, 2)) And Not IsEmpty(v(i, 2)) Then SettingNum = CDbl(v(i, 2))
            Exit Function
        End If
    Next i
End Function

' the typed wrappers the code reads (each: the sheet's override, else the constant)
Public Function KronosStaleDays() As Long
    KronosStaleDays = CLng(SettingNum("KRONOS_STALE_DAYS", KRONOS_STALE_DAYS))
End Function
Public Function WeekCap() As Double
    WeekCap = SettingNum("WEEK_MAX_HOURS", WEEK_MAX_HOURS)
End Function
Public Function MembersStaleDays() As Long
    MembersStaleDays = CLng(SettingNum("MEMBERS_STALE_DAYS", MEMBERS_STALE_DAYS))
End Function
Public Function SignupStaleDays() As Long
    SignupStaleDays = CLng(SettingNum("SIGNUP_STALE_DAYS", SIGNUP_STALE_DAYS))
End Function
Public Function PaybackWarnDays() As Long
    PaybackWarnDays = CLng(SettingNum("PAYBACK_WARN_DAYS", PAYBACK_WARN_DAYS))
End Function
Public Function PostingCloseDays() As Long
    PostingCloseDays = CLng(SettingNum("POSTING_CLOSE_DAYS", POSTING_CLOSE_DAYS))
End Function

' The default of a tunable name (for SetSetting's prompt and SystemCheck's line)
Private Function TunableDefault(ByVal nm As String) As Double
    Select Case UCase$(Trim$(nm))
        Case "KRONOS_STALE_DAYS": TunableDefault = KRONOS_STALE_DAYS
        Case "WEEK_MAX_HOURS": TunableDefault = WEEK_MAX_HOURS
        Case "MEMBERS_STALE_DAYS": TunableDefault = MEMBERS_STALE_DAYS
        Case "SIGNUP_STALE_DAYS": TunableDefault = SIGNUP_STALE_DAYS
        Case "PAYBACK_WARN_DAYS": TunableDefault = PAYBACK_WARN_DAYS
        Case "POSTING_CLOSE_DAYS": TunableDefault = POSTING_CLOSE_DAYS
    End Select
End Function

' Admin (Alt+F8): one override written - the name picked from the tunable list, the value a number, the previous kept
Public Sub SetSetting()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim names As Variant: names = Split(TUNABLE, ",")
    Dim msg As String, i As Long
    For i = LBound(names) To UBound(names)
        msg = msg & (i + 1) & ")  " & names(i) & "   (default " & HoursText(TunableDefault(CStr(names(i)))) & ", now " & HoursText(SettingNum(CStr(names(i)), TunableDefault(CStr(names(i))))) & ")" & vbNewLine
    Next i
    Dim pk As Variant
    pk = Application.InputBox(Prompt:="Which setting?" & vbNewLine & vbNewLine & msg & vbNewLine & "Type the number (a rule's own numbers are not settings: a rule change is a release):", Title:="Settings", Default:=1, Type:=1)
    If VarType(pk) = vbBoolean Then Exit Sub
    Dim pick As Long: pick = CLng(Val(CStr(pk)))
    If pick < 1 Or pick > UBound(names) + 1 Then MsgBox "Type one of the numbers listed.", vbExclamation, "Settings": Exit Sub
    Dim nm As String: nm = CStr(names(pick - 1))
    Dim a As Variant
    a = Application.InputBox(Prompt:="New value for " & nm & " (a number; the default is " & HoursText(TunableDefault(nm)) & "):", Title:="Settings", Default:=HoursText(SettingNum(nm, TunableDefault(nm))), Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim t As String: t = Replace(Trim$(CStr(a)), ",", ".")
    If Not IsNumeric(t) Then MsgBox "That isn't a number.", vbExclamation, "Settings": Exit Sub
    Dim newVal As Double: newVal = Val(t)
    If newVal < 0 Then MsgBox "A setting is never negative.", vbExclamation, "Settings": Exit Sub
    Dim note As String: note = Trim$(SafeStr(Application.InputBox(Prompt:="Why (optional; kept on the row):", Title:="Settings", Type:=2)))
    If note = "False" Then note = ""
    Dim lo As ListObject: Set lo = SettingsTable(True)
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim v As Variant, r As Long, cnt As Long, prevTxt As String
    If Not lo.DataBodyRange Is Nothing Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) > 0 Then
            v = RangeTo2D(lo.DataBodyRange)
            cnt = UBound(v, 1)
            For i = 1 To cnt
                If StrComp(Trim$(SafeStr(v(i, 1))), nm, vbTextCompare) = 0 Then r = i: prevTxt = SafeStr(v(i, 2)): Exit For
            Next i
        End If
    End If
    If prevTxt = "" Then prevTxt = HoursText(TunableDefault(nm)) & " (default)"
    Application.EnableEvents = False
    UnlockSheet ws
    If r = 0 Then
        cnt = cnt + 1
        lo.Resize lo.HeaderRowRange.Resize(cnt + 1)
        r = cnt
        lo.HeaderRowRange.Offset(r, 0).Resize(1, ST_N).Value = Array(nm, newVal, Int(Date), CurrentUser(), prevTxt, "'" & note)
    Else
        lo.DataBodyRange.Rows(r).Resize(1, ST_N).Value = Array(nm, newVal, Int(Date), CurrentUser(), prevTxt, "'" & note)
    End If
    ProtectSheet ws
    Application.EnableEvents = True
    MarkCalDirty
    SaveIfWritable
    MsgBox nm & " = " & HoursText(newVal) & " from today (was " & prevTxt & "). SystemCheck lists every override; SetSetting again puts the default back by typing it.", vbInformation, "Settings"
    Exit Sub
Oops:
    Application.EnableEvents = True
    ReportError "SetSetting"
End Sub

' SystemCheck's line: the overrides in force ("none" when the sheet is empty or absent); IntegrityCheck's problems in bad
Public Function SettingsOverrideText(Optional ByRef bad As String) As String
    bad = ""
    SettingsOverrideText = "Settings overrides: none"
    On Error Resume Next                            ' never raises
    Dim lo As ListObject: Set lo = SettingsTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, s As String, nm As String
    For i = 1 To UBound(v, 1)
        nm = Trim$(SafeStr(v(i, 1)))
        If nm <> "" Then
            If Not IsTunable(nm) Then
                bad = bad & IIf(bad = "", "", "; ") & nm & " is not a setting the code reads"
            ElseIf Not IsNumeric(v(i, 2)) Or IsEmpty(v(i, 2)) Then
                bad = bad & IIf(bad = "", "", "; ") & nm & "'s value is not a number (the default applies)"
            Else
                s = s & IIf(s = "", "", ", ") & nm & "=" & HoursText(CDbl(v(i, 2))) & IIf(IsDate(v(i, 3)), " (since " & Format$(CDate(v(i, 3)), "mmm d") & ")", "")
            End If
        End If
    Next i
    If s <> "" Then SettingsOverrideText = "Settings overrides: " & s
End Function

' Admin: the Settings sheet shown (made on first use)
Public Sub ShowSettings()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    Dim lo As ListObject: Set lo = SettingsTable(True)
    If lo Is Nothing Then Exit Sub
    SetSheetVisible lo.Parent, xlSheetVisible
    lo.Parent.Activate
    MsgBox "Nothing is typed on this sheet: Alt+F8 > SetSetting writes a row. " & SettingsOverrideText(), vbInformation, "Settings"
    Exit Sub
Oops:
    ReportError "ShowSettings"
End Sub

' v27 (item 6, 5.3): what the Rules sheet records as a rule's basis (the Basis an admin typed: the CBA article, or
' "agreed <date> by <who>"); "" when none is typed or the sheet is not made yet. Never raises.
Public Function RuleBasis(ByVal id As String) As String
    On Error Resume Next
    Dim lo As ListObject: Set lo = RulesTable(False)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(SafeStr(v(i, 1))), Trim$(id), vbTextCompare) = 0 Then RuleBasis = Trim$(SafeStr(v(i, 3))): Exit Function
    Next i
End Function

' The release a rule came in with ("" = unknown id)
Public Function RuleSince(ByVal id As String) As String
    Dim v As Variant, k As Variant, p As Variant
    v = RuleList()
    For Each k In v
        p = Split(CStr(k), "|")
        If StrComp(CStr(p(0)), Trim$(id), vbTextCompare) = 0 Then
            If UBound(p) >= 3 Then RuleSince = CStr(p(3))
            Exit Function
        End If
    Next k
End Function

' v27 (item 6): the receipt's "On paper this would have been different" line for one rule the paper roster never had
Public Function PaperRuleLine(ByVal id As String) As String
    Dim b As String: b = RuleBasis(id)
    PaperRuleLine = "  On paper this would have been different - " & id & ": " & RuleText(id) & "  (" & _
                    IIf(b <> "", b, "in the roster since " & RuleSince(id) & "; no basis typed on the Rules sheet yet") & ")" & vbNewLine
End Function
