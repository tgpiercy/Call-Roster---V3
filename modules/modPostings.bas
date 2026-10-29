Attribute VB_Name = "modPostings"
Option Explicit

' ====================================================================
'  OVERTIME POSTINGS - create, award, charge, email
'
'  1. Admin clicks New Posting on the OT Roster and fills the pop-up.
'     A Posting ID is assigned and the eligible list - with each
'     operator's Total Hours AT THAT MOMENT - is saved as a snapshot.
'  2. The Overtime Posting tab is drawn, printed and saved as PDF;
'     operators sign up on paper.
'  3. After it closes, Award Posting (pop-up): pick the ID, tick who
'     signed up, then SUBMIT AWARD / NO TAKERS / CANCEL REQUEST
'     (charging rules are with ComputeAward below).
'  4. An Outlook email opens to the awarded operator(s), cc supervisors.
'
'  v27.1 (the refactor, moved unchanged): the printed posting sheet is in modPostingSheet, a person joining or
'  leaving an open posting's list (and the posting records, reopen) in modPostingSync, the award email in
'  modPostingMail. The settings, the Types and the column constants stay here.
' ====================================================================

' ---------------- configuration ----------------
Public Const POST_VERSION As String = "Postings v29.1 (Oct 28)"
Public Const POSTINGS_SHEET As String = "Postings"
Public Const PT_TABLE As String = "tblPostings"
Public Const PL_TABLE As String = "tblPostingList"
' ---- award email CC lists (semicolon separated, e.g. "a@syensqo.com; b@syensqo.com") ----
'  Phosphine, Derivatives, SBP          -> shift supervisors + ops superintendent
'  Cypure, Material Handler (Fumigation)-> gases supervisors + ops superintendent
'  QC Derivatives (Lab), QC Gases       -> QC supervisor + ops superintendent
'  Utilities                            -> utilities supervisor + ops superintendent
'  (the unit -> group mapping is in AwardCC below)
Public Const CC_OPS_SUPT As String = "Raymond.Lanthier@syensqo.com"  ' Operations Superintendent
Public Const CC_SHIFT_SUPS As String = "graeme.piercy@syensqo.com; Chris.Mihaychuk@syensqo.com; " & _
                                       "David.Shelton@syensqo.com; Claude.Campeau@syensqo.com"  ' Shift Supervisors
Public Const CC_GASES_SUPS As String = "Ian.Honculada@syensqo.com"     ' Gases Supervisors
Public Const CC_QC_SUP As String = "Nicole.Matthews@syensqo.com"     ' Quality Control Supervisor
Public Const CC_UTIL_SUP As String = "mark.mcdougall@syensqo.com"     ' Utilities Supervisor
Public Const AWARD_CC As String = ""                                 ' extra CC on EVERY award email

' Overtime Posting sheet (output only - the macro redraws it for each posting)
Public Const POSTING_CLOSE_TIME As String = "8:00 AM"    ' printed after the closing date
Public Const POSTING_PDF_SUBFOLDER As String = "Postings" ' inside REPORT_FOLDER
Public Const POSTING_CLOSE_DAYS As Long = 14              ' default: closes this many days before the shift
Public Const POSTING_OUTSIDE_LINES As Long = 3            ' blank out-of-department sign-up lines on the sheet

' tblPostings columns
Public Const PS_ID As Long = 1
Public Const PS_STATUS As Long = 2
Public Const PS_POSTED As Long = 3
Public Const PS_CLOSES As Long = 4
Public Const PS_DATE As Long = 5
Public Const PS_UNIT As Long = 6
Public Const PS_SHIFT As Long = 7
Public Const PS_CREW As Long = 8
Public Const PS_HOURS As Long = 9
Public Const PS_POS As Long = 10
Public Const PS_REASON As Long = 11
Public Const PS_AWARDTO As Long = 12
Public Const PS_AWARDON As Long = 13
Public Const PS_AWARDBY As Long = 14
Public Const PS_POSTEDBY As Long = 15
Public Const PS_N As Long = 15

' tblPostingList columns (snapshot of the posted list)
Public Const PL_ID As Long = 1
Public Const PL_RANK As Long = 2
Public Const PL_EMP As Long = 3
Public Const PL_SHIFT As Long = 4
Public Const PL_CODE As Long = 5
Public Const PL_HRS As Long = 6
Public Const PL_SIGNED As Long = 7
Public Const PL_RESULT As Long = 8
Public Const PL_CHARGED As Long = 9
Public Const PL_SIGNHRS As Long = 10        ' v24: hours signed up for when it is only PART of the shift (blank = the whole shift)
Public Const PL_N As Long = 10
Public Const EXEMPT_TAG As String = "Exempt - "    ' v24: starts the "never charged on this posting" note in a list row's Result

Public Type PostRec
    ID As String
    Status As String
    PostedOn As Date
    ClosesOn As Date
    ShiftDate As Date
    Unit As String
    ShiftLbl As String
    Crew As String
    Hours As Double
    Positions As Long
    Reason As String
End Type

Public Type PListRec
    Rank As Long
    EmpName As String
    EmpShift As String
    DeptCode As String
    Grade As Double
    TotalHrs As Double
    SignedUp As Boolean
    Awarded As Boolean
    Charged As Double
    Result As String
    SignupHrs As Double         ' v24: 0 = signed up for the whole shift; otherwise the part they'll take
    ESign As String             ' v24.1: "<login> <time>" when the tick came from an electronic sign-up
    ExemptNote As String        ' v24: "Exempt - moved Jan 12" / "Exempt - left Jan 12": never charged on this posting
    MaxHours As Double          ' v26.5.5 (item 9): the most of this shift they may take under the 16 h rule, as of now
    MaxSet As Boolean           ' v26.5.5: MaxHours was worked out (the award pop-up fills it; a bare LoadPosting has no limit)
    OverrideNote As String      ' v27 (item 21): "[over 48 h/week: ...]" - the weekly-cap reason the award pop-up took; goes into the row's Reason
End Type


' ====================================================================
'  ONE-TIME SETUP (safe to run again)
' ====================================================================
Public Sub SetupPostings()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    CreatePostingTables
    SetupEmailColumn
    ResetPostingSheet
    Application.ScreenUpdating = True
    SaveIfWritable
    MsgBox "Postings are set up." & vbNewLine & vbNewLine & _
           "- 'Postings' sheet holds every posting and its posted list" & vbNewLine & _
           "- Roster column K is now EMAIL - fill in operator addresses" & vbNewLine & _
           "- Overtime Posting tab is now output only (New Posting fills it in)" & vbNewLine & vbNewLine & _
           "Run ApplyRosterLayout next to refresh the buttons.", vbInformation, "Call Out Roster"
End Sub

Private Sub CreatePostingTables()
    Dim ws As Worksheet
    If SheetExists(POSTINGS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(POSTINGS_SHEET)
    Else
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(LOG_SHEET))
        ws.Name = POSTINGS_SHEET
        RelockStructure
        SetTabColor ws, RGB(0, 128, 70)                 ' v25.1
    End If
    UnlockSheet ws

    Dim lo As ListObject
    If GetTable(PT_TABLE) Is Nothing Then
        ws.Range("A1").Value = "Overtime postings"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 13
        ws.Range("A3").Resize(1, PS_N).Value = Array("Posting ID", "Status", "Posted On", "Closes On", _
            "Shift Date", "Unit", "Shift", "Crew", "Hours", "Positions", "Reason", _
            "Awarded To", "Awarded On", "Awarded By", "Posted By")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").Resize(2, PS_N), , xlYes)
        lo.Name = PT_TABLE
        lo.TableStyle = "TableStyleMedium7"
        ws.Range("C:E").NumberFormat = "mmm d, yyyy"
        ws.Range("M:M").NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ws.Range("A:O").ColumnWidth = 12
        ws.Columns("F").ColumnWidth = 16
        ws.Columns("K").ColumnWidth = 18
        ws.Columns("L").ColumnWidth = 28
    End If

    If GetTable(PL_TABLE) Is Nothing Then
        ws.Range("Q1").Value = "Posted lists (snapshot of hours when each posting went up)"
        ws.Range("Q1").Font.Bold = True
        ws.Range("Q3").Resize(1, PL_N).Value = Array("Posting ID", "Rank", "Employee", "Shift", _
            "Dept Code", "Total Hrs (posted)", "Signed Up", "Result", "Hours Charged", "Signup Hours")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("Q3").Resize(2, PL_N), , xlYes)
        lo.Name = PL_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Range("Q:Z").ColumnWidth = 11
        ws.Columns("S").ColumnWidth = 24
    End If
    ws.Rows(3).WrapText = True
    ProtectSheet ws
End Sub

Private Sub SetupEmailColumn()
    Dim ws As Worksheet: Set ws = RosterSheet()
    UnlockSheet ws
    If UCase$(Trim$(SafeStr(ws.Cells(2, C_EMAIL).Value))) <> "EMAIL" Then
        Dim lastRow As Long: lastRow = RosterLastRow()
        ' column K only held copies of the Shift column - free it up for email addresses
        If lastRow >= R_FIRST Then ws.Range(ws.Cells(R_FIRST, C_EMAIL), ws.Cells(lastRow, C_EMAIL)).ClearContents
        ws.Cells(2, C_EMAIL).Value = "EMAIL"
    End If
    ws.Columns(C_EMAIL).Hidden = False
    ws.Columns(C_EMAIL).ColumnWidth = 26
    ProtectSheet ws
End Sub

Public Function PostingTablesReady() As Boolean
    If GetTable(PT_TABLE) Is Nothing Or GetTable(PL_TABLE) Is Nothing Then
        MsgBox "Postings aren't set up yet: an admin runs Alt+F8 > SetupPostings once.", vbExclamation, "Postings"
    ElseIf Not PostingListCurrent() Then
        MsgBox "The posting lists need their Signup Hours column (v24). Run Alt+F8 > SetupPeople once, then try again.", vbExclamation, "Call Out Roster"
    Else
        PostingTablesReady = True
    End If
End Function

' True when tblPostingList has the v24 Signup Hours column (or doesn't exist yet)
Public Function PostingListCurrent() As Boolean
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lr Is Nothing Then PostingListCurrent = True: Exit Function
    PostingListCurrent = (lr.ListColumns.Count >= PL_N)
End Function

' Adds the Signup Hours column on the END of tblPostingList (v24; called by
' SetupPeople, which takes the backup and checks the roster totals). Nothing
' moves. False = not done, why says what is in the way.
Public Function UpgradePostingList(ByRef why As String) As Boolean
    why = ""
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lr Is Nothing Then UpgradePostingList = True: Exit Function          ' made with every column by SetupPostings
    If lr.ListColumns.Count >= PL_N Then UpgradePostingList = True: Exit Function
    If lr.ListColumns.Count <> PL_N - 1 Then
        why = "the posting list table has " & lr.ListColumns.Count & " columns, expected " & (PL_N - 1)
        Exit Function
    End If
    Dim ws As Worksheet: Set ws = lr.Parent
    Dim nextCol As Range
    Set nextCol = ws.Range(ws.Cells(lr.Range.Row, lr.Range.Column + lr.Range.Columns.Count), _
                           ws.Cells(lr.Range.Row + lr.Range.Rows.Count, lr.Range.Column + lr.Range.Columns.Count))
    If Application.WorksheetFunction.CountA(nextCol) > 0 Then
        why = "column " & ColLetter(nextCol.Column) & " on the Postings sheet, right beside the posted lists, isn't empty. Clear it and run SetupPeople again" & ByHandHint(ws.Name)
        Exit Function
    End If
    Dim nRows As Long: nRows = lr.ListRows.Count
    UnlockSheet ws
    lr.ListColumns.Add.Name = "Signup Hours"
    ws.Columns(lr.Range.Column + PL_SIGNHRS - 1).ColumnWidth = 11
    ProtectSheet ws
    If lr.ListColumns.Count <> PL_N Or lr.ListRows.Count <> nRows Then
        why = "the posting list table has " & lr.ListColumns.Count & " columns and " & lr.ListRows.Count & " rows after the change (expected " & PL_N & " and " & nRows & ")"
        Exit Function
    End If
    UpgradePostingList = True
End Function

Sub GoToPostingSheet()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not SheetExists(POSTING_SHEET) Then MsgBox "Sheet '" & POSTING_SHEET & "' not found.", , "Call Out Roster": Exit Sub
    SetSheetVisible ThisWorkbook.Worksheets(POSTING_SHEET), xlSheetVisible
    ThisWorkbook.Worksheets(POSTING_SHEET).Activate
    Exit Sub
Oops:
    ReportError "GoToPostingSheet"
End Sub


' ====================================================================
'  CREATE A POSTING  (New Posting pop-up on the OT Roster)
' ====================================================================
Public Sub ShowNewPostingForm()
    On Error GoTo Oops
    If Not RequireSupervisor() Then Exit Sub     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    frmNewPosting.Show
    Exit Sub
Oops:
    ReportError "ShowNewPostingForm"
End Sub

Public Sub ShowLoadPostingForm()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    frmNewPosting.LoadOnly
    frmNewPosting.Show
    Exit Sub
Oops:
    ReportError "ShowLoadPostingForm"
End Sub

' old macro names still work
Public Sub CreatePosting()
    On Error GoTo Oops
    ShowNewPostingForm
    Exit Sub
Oops:
    ReportError "CreatePosting"
End Sub

Public Sub ShowPosting()
    On Error GoTo Oops
    ShowLoadPostingForm
    Exit Sub
Oops:
    ReportError "ShowPosting"
End Sub

' How many people would be on the list right now (pop-up preview)
' How many would be listed. With a date + shift, only people free that day count;
' notFree gets the summary of who's left off.
Public Function EligibleCount(ByVal unit As String, ByVal crew As String, Optional ByVal d As Variant, _
                              Optional ByVal shiftLbl As String = "", Optional ByRef notFree As String, _
                              Optional ByRef nProv As Long) As Long
    nProv = 0
    If DeptCol(unit) = 0 Then Exit Function
    Dim recs() As EmpRec, n As Long, i As Long
    Dim hasDate As Boolean
    If Not IsMissing(d) Then hasDate = IsDate(d)
    If hasDate Then
        n = LoadEligible(unit, CrewFilter(crew), recs, DeptPostingMaxGrade(unit), True, CDate(d))   ' v24: crews as of the shift date
    Else
        n = LoadEligible(unit, CrewFilter(crew), recs, DeptPostingMaxGrade(unit), True)
    End If
    If n > 0 And hasDate And shiftLbl <> "" Then n = FilterNotFree(recs, n, CDate(d), shiftLbl, notFree)
    For i = 1 To n
        If recs(i).ProvHrs > 0 Then nProv = nProv + 1      ' v21: totals that include provisional accepted OT
    Next i
    EligibleCount = n
End Function

' Core: assigns the ID, snapshots the eligible list, saves both, draws
' the Overtime Posting sheet, then prints it and/or saves the PDF.
' p needs Unit, ShiftDate, ShiftLbl, Reason, Crew, Hours, Positions, ClosesOn.
' v27 (item 24): quiet = no "Create posting?" box (PostAllLines makes several in one go after ONE question)
Public Function CreatePostingFrom(ByRef p As PostRec, Optional ByVal doPrint As Boolean = True, _
                                  Optional ByVal doPdf As Boolean = True, Optional ByVal quiet As Boolean = False) As Boolean
    If Not RequireSupervisor() Then Exit Function     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Function
    If Not PostingTablesReady() Then Exit Function

    If p.Positions < 1 Then p.Positions = 1
    p.PostedOn = Date
    p.Status = "Open"

    If p.ClosesOn > p.ShiftDate Then
        If Ask("The posting closes after the shift date. Create it anyway?", vbYesNo + vbExclamation, "New Posting", vbYes) = vbNo Then Exit Function
    End If

    ' eligible list = same rules and order as before (call order, posting grade cutoffs)
    Application.Calculate
    Dim recs() As EmpRec, n As Long
    ' everyone with the department code (any grade) who is free that day:
    ' not scheduled, not under the rest rule, not booked off, not on a shift switch
    n = LoadEligible(p.Unit, CrewFilter(p.Crew), recs, DeptPostingMaxGrade(p.Unit), True, p.ShiftDate)
    Dim offList As String
    If n > 0 Then n = FilterNotFree(recs, n, p.ShiftDate, p.ShiftLbl, offList)
    If n = 0 Then
        If Not quiet Then MsgBox "No one is free to sign up for " & p.Unit & " / " & p.Crew & "." & _
                                 IIf(offList <> "", vbNewLine & "Left off: " & offList, ""), vbInformation, "Call Out Roster"
        Exit Function                                ' (review L5) quiet: the caller's one summary counts it
    End If

    p.ID = NextPostingID()
    Dim outTxt As String
    If doPrint And doPdf Then
        outTxt = "It will be printed and saved as a PDF."
    ElseIf doPrint Then
        outTxt = "It will be printed."
    ElseIf doPdf Then
        outTxt = "It will be saved as a PDF."
    Else
        outTxt = "It will be put on the Overtime Posting tab (not printed)."
    End If
    If Not quiet Then
        If Ask("Create posting " & p.ID & "?" & vbNewLine & vbNewLine & _
                  p.Unit & "  |  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & "  |  " & p.ShiftLbl & vbNewLine & _
                  "Crew: " & p.Crew & "   Hours: " & p.Hours & "   Positions: " & p.Positions & vbNewLine & _
                  "Reason: " & p.Reason & vbNewLine & _
                  "Closes: " & Format$(p.ClosesOn, "ddd mmm d, yyyy") & " @ " & POSTING_CLOSE_TIME & vbNewLine & vbNewLine & _
                  n & " eligible employees will be listed. " & outTxt & _
                  IIf(offList <> "", vbNewLine & vbNewLine & "Not listed: " & offList, ""), _
                  vbOKCancel + vbQuestion, "Create Posting", vbOK) <> vbOK Then Exit Function
    End If

    ' --- save posting + snapshot list ---
    Dim pr() As Variant: ReDim pr(1 To 1, 1 To PS_N)
    pr(1, PS_ID) = p.ID
    pr(1, PS_STATUS) = p.Status
    pr(1, PS_POSTED) = p.PostedOn
    pr(1, PS_CLOSES) = p.ClosesOn
    pr(1, PS_DATE) = p.ShiftDate
    pr(1, PS_UNIT) = p.Unit
    pr(1, PS_SHIFT) = p.ShiftLbl
    pr(1, PS_CREW) = p.Crew
    pr(1, PS_HOURS) = p.Hours
    pr(1, PS_POS) = p.Positions
    pr(1, PS_REASON) = "'" & p.Reason                    ' v24.3.2: as text, whatever it starts with (reads back without the ')
    pr(1, PS_POSTEDBY) = CurrentUser()

    Dim lr() As Variant: ReDim lr(1 To n, 1 To PL_N)
    Dim lst() As PListRec: ReDim lst(1 To n)
    Dim i As Long
    For i = 1 To n
        lr(i, PL_ID) = p.ID
        lr(i, PL_RANK) = i
        lr(i, PL_EMP) = recs(i).EmpName
        lr(i, PL_SHIFT) = recs(i).EmpShift
        lr(i, PL_CODE) = recs(i).DeptCode
        lr(i, PL_HRS) = recs(i).TotalHrs
        lst(i).Rank = i
        lst(i).EmpName = recs(i).EmpName
        lst(i).EmpShift = recs(i).EmpShift
        lst(i).DeptCode = recs(i).DeptCode
        lst(i).Grade = recs(i).Grade
        lst(i).TotalHrs = recs(i).TotalHrs
    Next i

    ' list first: a posting row only appears once its list is safely saved
    If Not AppendRows(GetTable(PL_TABLE), lr, PL_N) Then Exit Function
    If Not AppendRows(GetTable(PT_TABLE), pr, PS_N) Then Exit Function
    Dim calTxt As String: calTxt = CalendarPosted(p)
    MarkMembersDirty                    ' v24.1: a posting members can't see yet
    SaveIfWritable
    CreatePostingFrom = True

    Dim report As String
    OutputPosting p, lst, n, False, doPrint, doPdf, report
    If quiet Then Exit Function                     ' (review L5) the caller says it once for all of them
    Tell "Posting " & p.ID & " created." & vbNewLine & vbNewLine & report & vbNewLine & vbNewLine & _
         "Calendar: " & calTxt, vbInformation, "Create Posting"
End Function


' Load posting <id> and view / print / PDF it. report gets a one-line
' summary of what happened. False if the posting isn't found.
Public Function PostingOutput(ByVal id As String, ByVal doView As Boolean, ByVal doPrint As Boolean, _
                              ByVal doPdf As Boolean, ByRef report As String) As Boolean
    Dim p As PostRec, lst() As PListRec, n As Long
    n = LoadPosting(UCase$(Trim$(id)), p, lst)
    If n = 0 Then Exit Function
    PostingOutput = OutputPosting(p, lst, n, doView, doPrint, doPdf, report)
End Function

Private Function OutputPosting(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                               ByVal doView As Boolean, ByVal doPrint As Boolean, ByVal doPdf As Boolean, _
                               ByRef report As String) As Boolean
    Dim ws As Worksheet: Set ws = PostingSheet()
    Application.ScreenUpdating = False
    RenderPostingSheet ws, p, lst, n
    Application.ScreenUpdating = True
    report = "It's on the Overtime Posting tab."

    ' a hidden sheet can't print or export - show it for a moment if needed
    Dim wasVis As XlSheetVisibility: wasVis = ws.Visible
    If (doPrint Or doPdf) And wasVis <> xlSheetVisible Then SetSheetVisible ws, xlSheetVisible

    Dim em As String
    If doPdf Then
        Dim pdfPath As String
        If PostingPdf(ws, p, pdfPath, em) Then
            report = report & vbNewLine & "PDF saved: " & pdfPath
        Else
            report = report & vbNewLine & "PDF NOT saved (" & em & ")"
        End If
    End If
    If doPrint Then
        em = ""
        On Error Resume Next
        ws.PrintOut Copies:=1
        If Err.Number <> 0 Then em = Err.Description
        On Error GoTo 0
        If em = "" Then
            report = report & vbNewLine & "Sent to printer: " & Application.ActivePrinter
        Else
            report = report & vbNewLine & "NOT printed (" & em & ")"
        End If
    End If

    If ws.Visible <> wasVis And Not doView Then SetSheetVisible ws, wasVis
    If doView Then
        SetSheetVisible ws, xlSheetVisible
        ws.Activate
        On Error Resume Next
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.View = xlNormalView
        ActiveWindow.ScrollRow = 1
        ws.Range("A1").Select
        On Error GoTo 0
    End If
    OutputPosting = True
End Function

' The PDF's file name (no folder, no extension): "P-0012 Phosphine 2026-11-05 A" with the illegal characters replaced
Private Function PostingPdfName(ByRef p As PostRec) As String
    Dim fName As String, ch As Variant
    fName = p.ID & " " & p.Unit & " " & Format$(p.ShiftDate, "yyyy-mm-dd") & " " & p.Crew
    For Each ch In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        fName = Replace(fName, CStr(ch), "-")
    Next ch
    PostingPdfName = fName
End Function

' The folder the posting PDFs go to (REPORT_FOLDER\Postings\ when it exists, else beside the workbook)
Private Function PostingPdfFolder() As String
    Dim folder As String: folder = EnsureSlash(ReportFolder())
    If FolderExists(folder) Then
        Dim subF As String: subF = folder & POSTING_PDF_SUBFOLDER & Application.PathSeparator
        If FolderExists(subF) Then folder = subF
    Else
        folder = EnsureSlash(ThisWorkbook.Path)
    End If
    PostingPdfFolder = folder
End Function

' ====================================================================
'  v27 (item 24, 1.2): THE JOBS THE FORMS COULD NOT RECORD
' ====================================================================
' Withdraw award (admin): a closed posting's AWARDED rows are voided (the refusals stay charged - the people ahead
' still refused), the calendar line opens again (VoidRef releases it), the list row reads "Awarded - withdrawn",
' "still paid" yes / no and the reason go into the Void Note, a receipt; then the offer of a call-out for the shift.
' Not a reopen: a re-award would charge the refusers again, and Reopen needs every row voided (Void Entry > P-n).
Public Sub WithdrawAward()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, listTxt As String, nAw As Long
    Dim live As Object: Set live = PostingsWithAwardedRow()      ' (review L1) a withdrawn award is not offered again
    For i = UBound(pv, 1) To 1 Step -1
        If SafeStr(pv(i, PS_STATUS)) = "Awarded" And nAw < 12 And live.Exists(Trim$(SafeStr(pv(i, PS_ID)))) Then
            nAw = nAw + 1
            listTxt = listTxt & SafeStr(pv(i, PS_ID)) & "  " & SafeStr(pv(i, PS_UNIT)) & "  " & Format$(ToDate(pv(i, PS_DATE)), "ddd mmm d") & "  " & SafeStr(pv(i, PS_SHIFT)) & "  -> " & SafeStr(pv(i, PS_AWARDTO)) & vbNewLine
        End If
    Next i
    If nAw = 0 Then MsgBox "No awarded postings.", vbInformation, "Withdraw award": Exit Sub
    Dim a As Variant
    a = Application.InputBox(Prompt:="Awarded postings (newest first):" & vbNewLine & listTxt & vbNewLine & "Withdraw which award? Type the posting id:", Title:="Withdraw award", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim id As String: id = UCase$(Trim$(CStr(a)))
    Dim p As PostRec, lst() As PListRec, n As Long
    n = LoadPosting(id, p, lst)
    If n = 0 Then MsgBox id & " was not found.", vbExclamation, "Withdraw award": Exit Sub
    If p.Status <> "Awarded" Then MsgBox id & " is " & p.Status & ", not Awarded.", vbExclamation, "Withdraw award": Exit Sub
    Dim paid As Long: paid = MsgBox("Withdraw the award of " & id & " (" & p.Unit & " " & Format$(p.ShiftDate, "ddd mmm d") & " " & p.ShiftLbl & ")?" & vbNewLine & vbNewLine & _
                                    "Is the awarded person STILL PAID for it?  Yes = still paid.  No = not paid.  Cancel = leave the award.", vbYesNoCancel + vbQuestion, "Withdraw award")
    If paid = vbCancel Then Exit Sub
    Dim why As String: why = AskRequired("Why is the award withdrawn? (kept in the Void Note)", "Withdraw award")
    If why = "" Then Exit Sub
    why = "Award withdrawn (still paid: " & IIf(paid = vbYes, "yes", "no") & ") - " & why
    Dim before As Object: Set before = ReceiptBefore()
    Dim calN As Long, cnt As Long
    cnt = VoidRef(id, why, calN, "Awarded")
    If cnt = 0 Then MsgBox "Nothing was voided for " & id & ": no live Awarded row (withdrawn already?), or the void stopped on an error (see the box before this one).", vbExclamation, "Withdraw award": Exit Sub
    ' the list row says so (LoadPosting reads Awarded = Result "Awarded", so this takes the person off the awarded set)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If Not lr Is Nothing Then
        If Not lr.DataBodyRange Is Nothing Then
            Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
            Dim hit As Boolean
            For i = 1 To UBound(lv, 1)
                If StrComp(Trim$(SafeStr(lv(i, PL_ID))), id, vbTextCompare) = 0 And Left$(SafeStr(lv(i, PL_RESULT)), 7) = "Awarded" _
                   And InStr(1, SafeStr(lv(i, PL_RESULT)), "withdrawn", vbTextCompare) = 0 Then      ' (review L1) "Awarded | Exempt - ..." too
                    If Not hit Then UnlockSheet lr.Parent: hit = True
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = "Awarded - withdrawn" & Mid$(SafeStr(lv(i, PL_RESULT)), 8)
                End If
            Next i
            If hit Then ProtectSheet lr.Parent
        End If
    End If
    SaveIfWritable
    ShowReceiptNow "Award of " & id & " withdrawn", _
                   "  " & id & "  " & cnt & " Awarded row(s) voided: " & why & vbNewLine & "  (the refusals stay charged: the people ahead still refused)" & vbNewLine, _
                   ReceiptChanges(before), _
                   "  Calendar: " & p.Unit & " " & Format$(p.ShiftDate, "ddd mmm d") & " " & p.ShiftLbl & IIf(calN > 0, " - the line is OPEN again (" & calN & " line(s) released).", " - no calendar line held the award (nothing to release).") & vbNewLine & _
                   "  Postings sheet: " & id & " stays Awarded with the list row marked withdrawn; Void Entry > " & id & " + Reopen is the way to award it again." & vbNewLine, _
                   WrongText()
    If Ask("Fill the shift now with a call-out?", vbYesNo + vbQuestion, "Withdraw award", vbNo) = vbYes Then
        If LogTable() Is Nothing Then Exit Sub
        Load frmOvertimeEntry
        frmOvertimeEntry.Prefill p.Unit, p.ShiftDate, p.ShiftLbl, "Withdrawn award", p.Hours
        frmOvertimeEntry.Show
    End If
    Exit Sub
Oops:
    ReportError "WithdrawAward"
End Sub

' Edit posting (admin): the hours, date or positions of an OPEN posting. The frozen list is untouched (its hours are
' the people's hours at posting; the sign-ups stay); a changed date moves the calendar line (the old one voided,
' the posting placed again); the old PDF is renamed "... superseded"; the sheet is redrawn and printed.
Public Sub EditPosting()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim ids As Variant, ch As Variant: ch = OpenPostingChoices(ids)
    If Not IsArray(ids) Then MsgBox "No open postings.", vbInformation, "Edit posting": Exit Sub
    Dim i As Long, listTxt As String
    For i = LBound(ids) To UBound(ids)
        If i - LBound(ids) < 14 Then listTxt = listTxt & CStr(ch(i)) & vbNewLine
    Next i
    Dim a As Variant
    a = Application.InputBox(Prompt:="Open postings:" & vbNewLine & listTxt & vbNewLine & "Edit which? Type the posting id:", Title:="Edit posting", Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim id As String: id = UCase$(Trim$(CStr(a)))
    Dim p As PostRec, lst() As PListRec, n As Long
    n = LoadPosting(id, p, lst)
    If n = 0 Then MsgBox id & " was not found.", vbExclamation, "Edit posting": Exit Sub
    If p.Status <> "Open" Then MsgBox id & " is " & p.Status & ": only an open posting is edited.", vbExclamation, "Edit posting": Exit Sub
    Dim pOld As PostRec: pOld = p
    a = Application.InputBox(Prompt:="Hours (now " & HoursText(p.Hours) & "):", Title:="Edit posting", Default:=HoursText(p.Hours), Type:=1)
    If VarType(a) = vbBoolean Then Exit Sub
    If CDbl(a) <= 0 Or CDbl(a) > 24 Then MsgBox "Hours must be between 0 and 24.", vbExclamation, "Edit posting": Exit Sub
    p.Hours = CDbl(a)
    a = Application.InputBox(Prompt:="Shift date (now " & DateText(p.ShiftDate) & "):", Title:="Edit posting", Default:=DateText(p.ShiftDate), Type:=2)
    If VarType(a) = vbBoolean Then Exit Sub
    Dim d2 As Date
    If Not ParseTypedDate(Trim$(CStr(a)), d2) Then MsgBox "That isn't a date I can read.", vbExclamation, "Edit posting": Exit Sub
    If Int(d2) < RosterYearFrom() Then MsgBox "A shift before the roster year.", vbExclamation, "Edit posting": Exit Sub
    p.ShiftDate = Int(d2)
    a = Application.InputBox(Prompt:="Positions (now " & p.Positions & "):", Title:="Edit posting", Default:=p.Positions, Type:=1)
    If VarType(a) = vbBoolean Then Exit Sub
    If CLng(a) < 1 Then MsgBox "At least one position.", vbExclamation, "Edit posting": Exit Sub
    p.Positions = CLng(a)
    If p.Hours = pOld.Hours And Int(p.ShiftDate) = Int(pOld.ShiftDate) And p.Positions = pOld.Positions Then MsgBox "Nothing changed.", vbInformation, "Edit posting": Exit Sub
    If Int(p.ShiftDate) <> Int(pOld.ShiftDate) Then
        ' (review H1) the list is frozen: a date that changes who is free (scheduled, rest rule, booked off, a switch,
        ' restricted) would charge people who could not have worked and leave out people who could - a new posting
        Dim notNow As String, newFree As String
        If Not SameListOn(p, lst, n, notNow, newFree) Then
            MsgBox "On " & Format$(p.ShiftDate, "ddd mmm d") & " " & p.ShiftLbl & " the list would not be the same people, and a posting's list is frozen." & vbNewLine & _
                   IIf(notNow <> "", vbNewLine & "On the list, not free that day: " & notNow, "") & _
                   IIf(newFree <> "", vbNewLine & "Free that day, not on the list: " & newFree, "") & vbNewLine & vbNewLine & _
                   "Cancel Request on the award pop-up and post the shift again for the new date.", vbExclamation, "Edit posting"
            Exit Sub
        End If
        p.ClosesOn = DefaultClosesOn(p.ShiftDate, Now, PostingCloseDays())      ' (review M3) as a new posting's, on any date change; v29.1 (1.2): the one rule
    End If
    Dim nSign As Long, nPart As Long                 ' (review M1) the sign-ups an hours change re-means
    For i = 1 To n
        If lst(i).SignedUp Then
            nSign = nSign + 1
            If lst(i).SignupHrs > 0 Then nPart = nPart + 1
        End If
    Next i
    If Ask(id & ": " & HoursText(pOld.Hours) & " h -> " & HoursText(p.Hours) & " h, " & DateText(pOld.ShiftDate) & " -> " & DateText(p.ShiftDate) & ", " & pOld.Positions & " -> " & p.Positions & " position(s)." & _
           IIf(Int(p.ClosesOn) <> Int(pOld.ClosesOn), vbNewLine & "Closes " & DateText(pOld.ClosesOn) & " -> " & DateText(p.ClosesOn) & " " & POSTING_CLOSE_TIME & ".", "") & _
           IIf(p.Hours <> pOld.Hours And nSign > 0, vbNewLine & nSign & " signed up (" & nPart & " for part of the shift): a whole-shift sign-up now means " & HoursText(p.Hours) & " h; a partial keeps its hours.", "") & vbNewLine & vbNewLine & _
           "The list and its sign-ups stay as they are; the sheet is redrawn and printed, the old PDF renamed superseded. Go ahead?", vbYesNo + vbQuestion, "Edit posting", vbYes) = vbNo Then Exit Sub
    ' the row
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim r As Long
    For i = 1 To UBound(pv, 1)
        If StrComp(Trim$(SafeStr(pv(i, PS_ID))), id, vbTextCompare) = 0 Then r = i: Exit For
    Next i
    If r = 0 Then MsgBox id & " was not found on the Postings sheet.", vbExclamation, "Edit posting": Exit Sub
    UnlockSheet lo.Parent
    lo.DataBodyRange.Cells(r, PS_HOURS).Value = p.Hours
    lo.DataBodyRange.Cells(r, PS_DATE).Value = p.ShiftDate
    lo.DataBodyRange.Cells(r, PS_POS).Value = p.Positions
    lo.DataBodyRange.Cells(r, PS_CLOSES).Value = p.ClosesOn
    ProtectSheet lo.Parent
    ' the calendar line follows a date change; the hours on the line follow the hours
    Dim calTxt As String
    If Int(p.ShiftDate) <> Int(pOld.ShiftDate) Then
        VoidCoverageBySource id
        calTxt = CalendarPosted(p)
    ElseIf p.Hours <> pOld.Hours Then
        UpdateCoverageBySource id, "POSTED " & id, p.Hours, (p.Hours >= DefaultShiftHours(p.ShiftDate, IsNightShift(p.ShiftLbl), USUAL_SHIFT_HOURS) - 0.001)     ' (review M2 / L7) a full shift, DST nights too, clears the part hours
        calTxt = "the line's hours updated"
    End If
    MarkCalDirty
    MarkMembersDirty                                 ' (review M1) members see the new date / hours at the next publish
    ' the old PDF, if it is there, becomes "superseded" (review L4: a second edit gets its own name; said only when done)
    Dim oldPdf As String: oldPdf = PostingPdfFolder() & PostingPdfName(pOld) & ".pdf"
    Dim supTxt As String: supTxt = "No earlier PDF to rename."
    Dim supPath As String: supPath = PostingPdfFolder() & PostingPdfName(pOld) & " superseded"
    On Error Resume Next
    If Len(Dir$(oldPdf)) > 0 Then
        If Len(Dir$(supPath & ".pdf")) > 0 Then supPath = supPath & " " & Format$(Now, "yyyy-mm-dd hhnn")
        Err.Clear
        Name oldPdf As supPath & ".pdf"
        If Err.Number = 0 Then supTxt = "The old PDF is renamed: " & Mid$(supPath, InStrRev(supPath, "\") + 1) & ".pdf" Else supTxt = "The old PDF could not be renamed (open elsewhere?): the new one replaces it."
    End If
    On Error GoTo Oops
    Dim rpt As String
    PostingOutput id, False, True, True, rpt
    SaveIfWritable
    ShowReceiptNow "Posting " & id & " edited", _
                   "  " & id & "  " & HoursText(p.Hours) & " h  " & Format$(p.ShiftDate, "ddd mmm d") & "  " & p.Positions & " position(s)  (the Postings row; no OT Log row)" & vbNewLine, _
                   "  Nobody's Total Hours changed. The list's frozen hours and sign-ups are as they were." & vbNewLine, _
                   "  Calendar: " & IIf(calTxt = "", "unchanged", calTxt) & vbNewLine & "  " & rpt & vbNewLine & "  " & supTxt & vbNewLine, _
                   "  Edit it again, or Cancel Request on the award pop-up." & vbNewLine     ' (Oct 22: a stray ) here was a syntax error)
    Exit Sub
Oops:
    ReportError "EditPosting"
End Sub

' v27 (item 24, review H1): would a list made for p's (new) date and shift hold the same people as the frozen list?
' The list is frozen, so a date that changes who is free must be a new posting. notNow = on the list but not free
' then (an exempt row is never charged, so it does not count); newFree = free then but not on the list.
Private Function SameListOn(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByRef notNow As String, ByRef newFree As String) As Boolean
    notNow = "": newFree = ""
    Dim recs() As EmpRec, m As Long, offList As String, i As Long, k As Variant, nNot As Long, nNew As Long
    m = LoadEligible(p.Unit, CrewFilter(p.Crew), recs, DeptPostingMaxGrade(p.Unit), True, p.ShiftDate)
    If m > 0 Then m = FilterNotFree(recs, m, p.ShiftDate, p.ShiftLbl, offList)
    Dim want As Object: Set want = CreateObject("Scripting.Dictionary")
    want.CompareMode = 1
    For i = 1 To m
        want(Trim$(recs(i).EmpName)) = True
    Next i
    Dim have As Object: Set have = CreateObject("Scripting.Dictionary")
    have.CompareMode = 1
    For i = 1 To n
        have(Trim$(lst(i).EmpName)) = (lst(i).ExemptNote <> "")      ' True = exempt
    Next i
    For Each k In have.Keys
        If Not want.Exists(k) And Not have(k) Then
            nNot = nNot + 1
            If nNot <= 8 Then notNow = notNow & IIf(notNow = "", "", ", ") & k
        End If
    Next k
    For Each k In want.Keys
        If Not have.Exists(k) Then
            nNew = nNew + 1
            If nNew <= 8 Then newFree = newFree & IIf(newFree = "", "", ", ") & k
        End If
    Next k
    If nNot > 8 Then notNow = notNow & " and " & (nNot - 8) & " more"
    If nNew > 8 Then newFree = newFree & " and " & (nNew - 8) & " more"
    SameListOn = (nNot = 0 And nNew = 0)
End Function

' v29.1 (review HIGH 1.2): the closing date a posting gets when nobody types one - closeDays (PostingCloseDays(), 14)
' before the shift; when that closing MOMENT (8 AM, PostingClosed) has passed by nowAt, the day before the shift; when
' that has passed too, today (a shift tomorrow posted after 8 AM is born closed - the one unavoidable case). PURE
' (SelfTest "closeson"; mirrored in tests/rules.py). The ONE rule for New Posting, Post all lines and Edit posting: the
' last two used "d - 14, else today", so a posting made inside 14 days was born CLOSED at 8 AM - nobody could sign up,
' the award pop-up said READY TO AWARD and the Lite card called it "no takers". (re-review R1) the test is the closing
' moment, not the date: a shift exactly closeDays out posted after 8 AM closed at once under "d - 14 is not before today".
Public Function DefaultClosesOn(ByVal d As Date, ByVal nowAt As Date, ByVal closeDays As Long) As Date
    Dim cl As Date: cl = Int(d) - closeDays
    If PostingClosed(cl, nowAt) Then cl = Int(d) - 1
    If PostingClosed(cl, nowAt) Then cl = Int(nowAt)
    DefaultClosesOn = cl
End Function

' v27 (item 24, review L7): an open posting already on this unit, date and shift for the same Reason ("" when none) -
' Post all lines' Reason names the person off, so a second person off the same shift is not taken for a duplicate
Private Function OpenPostingOn(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String, ByVal reasonTxt As String) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, PS_STATUS)) = "Open" And StrComp(SafeStr(v(i, PS_UNIT)), unit, vbTextCompare) = 0 Then
            If IsDate(v(i, PS_DATE)) Then
                If Int(CDate(v(i, PS_DATE))) = Int(d) And IsNightShift(SafeStr(v(i, PS_SHIFT))) = IsNightShift(shiftLbl) _
                   And StrComp(Trim$(SafeStr(v(i, PS_REASON))), Trim$(reasonTxt), vbTextCompare) = 0 Then OpenPostingOn = SafeStr(v(i, PS_ID)): Exit Function
            End If
        End If
    Next i
End Function

' v27 (item 24, review L1): the posting ids whose list still holds an Awarded row (a withdrawn award's do not)
Private Function PostingsWithAwardedRow() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set PostingsWithAwardedRow = d
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lr Is Nothing Then Exit Function
    If lr.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lr.DataBodyRange)
    Dim i As Long, res As String
    For i = 1 To UBound(v, 1)
        res = SafeStr(v(i, PL_RESULT))
        If Left$(res, 7) = "Awarded" And InStr(1, res, "withdrawn", vbTextCompare) = 0 Then d(Trim$(SafeStr(v(i, PL_ID)))) = True
    Next i
End Function

' Post all lines (supervisors; from Book Time Off): one posting per scheduled shift of the booking, each covering its own
' calendar line (gCalPostKey), no question per posting, no print (each posting makes its PDF). Returns the ids.
Public Function PostAllLines(ByVal emp As String, ByVal unit As String, ByVal absId As String, ByVal dFrom As Date, ByVal dTo As Date) As String
    If Not RequireSupervisor() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If Not PostingTablesReady() Then Exit Function
    Dim d As Date, sh As String, p As PostRec, ids As String, nMade As Long, nSkip As Long, nPast As Long, dupTxt As String, cutTxt As String, dupId As String
    If dTo - dFrom > 31 Then cutTxt = "Only the first 32 days were looked at (to " & DateText(dFrom + 31) & "): post the rest from the Calendar.": dTo = dFrom + 31
    For d = Int(dFrom) To Int(dTo)
        sh = EmpShiftOn(emp, d)
        If sh <> "" And d < Date Then nPast = nPast + 1: sh = ""      ' (review L5) a shift already gone is not posted
        If sh <> "" Then
            dupId = OpenPostingOn(unit, d, sh, CalName(emp) & " off")     ' (review L7 / M3) this person's posting for that shift
            If dupId <> "" Then dupTxt = dupTxt & IIf(dupTxt = "", "", ", ") & dupId & " (" & Format$(d, "ddd mmm d") & ")": sh = ""
        End If
        If sh <> "" Then
            p.ID = "": p.Status = "": p.Unit = unit: p.ShiftDate = d: p.ShiftLbl = sh: p.Crew = CREW_ALL
            p.Hours = DefaultShiftHours(d, IsNightShift(sh), USUAL_SHIFT_HOURS): p.Positions = 1
            p.Reason = CalName(emp) & " off"                  ' never the type: a leave word must not reach the posting
            p.ClosesOn = DefaultClosesOn(d, Now, PostingCloseDays())     ' v29.1 (review 1.2): the one rule - inside 14 days it closes the day before the shift
            gCalPostKey = "A:" & absId: gCalPostDate = d: gCalPostShift = sh     ' the posting covers this booking's line
            If CreatePostingFrom(p, False, True, True) Then
                nMade = nMade + 1
                ids = ids & IIf(ids = "", "", ", ") & p.ID
            Else
                nSkip = nSkip + 1
            End If
            gCalPostKey = ""
        End If
    Next d
    ' (review L5) ONE summary for the whole run
    Dim sumTxt As String
    If nMade > 0 Then sumTxt = nMade & " posting(s) made, each on its own calendar line with a PDF: " & ids & "."
    If nSkip > 0 Then sumTxt = sumTxt & IIf(sumTxt = "", "", vbNewLine) & nSkip & " shift(s) not posted: nobody free to sign up - post them one at a time from the Calendar."
    If dupTxt <> "" Then sumTxt = sumTxt & IIf(sumTxt = "", "", vbNewLine) & "Already posted (left as they are): " & dupTxt & "."
    If nPast > 0 Then sumTxt = sumTxt & IIf(sumTxt = "", "", vbNewLine) & nPast & " shift(s) before today not posted."
    If cutTxt <> "" Then sumTxt = sumTxt & IIf(sumTxt = "", "", vbNewLine) & cutTxt
    If sumTxt = "" Then sumTxt = emp & " has no scheduled shift between " & DateText(dFrom) & " and " & DateText(dTo) & ": nothing to post."
    Tell sumTxt, IIf(nSkip > 0, vbExclamation, vbInformation), "Post all lines"
    PostAllLines = ids
End Function

' PDF goes to REPORT_FOLDER\Postings\ (made if missing), else next to the workbook
Private Function PostingPdf(ByVal ws As Worksheet, ByRef p As PostRec, ByRef pdfPath As String, _
                            ByRef em As String) As Boolean
    Dim folder As String, subF As String
    folder = EnsureSlash(ReportFolder())
    If FolderExists(folder) Then
        subF = folder & POSTING_PDF_SUBFOLDER & Application.PathSeparator
        If Not FolderExists(subF) Then
            On Error Resume Next
            MkDir subF
            On Error GoTo 0
        End If
        If FolderExists(subF) Then folder = subF
    Else
        folder = EnsureSlash(ThisWorkbook.Path)
    End If

    pdfPath = folder & PostingPdfName(p) & ".pdf"

    On Error Resume Next
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then
        em = Err.Description & " - is the file open?"
    Else
        PostingPdf = True
    End If
    On Error GoTo 0
End Function


' ====================================================================
'  READ
' ====================================================================
Public Function GetTable(ByVal nm As String) As ListObject
    If Not SheetExists(POSTINGS_SHEET) Then Exit Function
    On Error Resume Next
    Set GetTable = ThisWorkbook.Worksheets(POSTINGS_SHEET).ListObjects(nm)
    On Error GoTo 0
End Function

Public Function ToDate(ByVal x As Variant) As Date
    If IsDate(x) Then
        ToDate = CDate(x)
    ElseIf IsNumeric(x) And Not IsEmpty(x) Then
        ToDate = CDate(CDbl(x))
    End If
End Function

' Returns the list count (0 = not found). p and lst() are filled.
Public Function LoadPosting(ByVal id As String, ByRef p As PostRec, ByRef lst() As PListRec) As Long
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Function

    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, found As Boolean
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(id) Then
            p.ID = SafeStr(v(i, PS_ID))
            p.Status = SafeStr(v(i, PS_STATUS))
            p.PostedOn = ToDate(v(i, PS_POSTED))
            p.ClosesOn = ToDate(v(i, PS_CLOSES))
            p.ShiftDate = ToDate(v(i, PS_DATE))
            p.Unit = SafeStr(v(i, PS_UNIT))
            p.ShiftLbl = SafeStr(v(i, PS_SHIFT))
            p.Crew = SafeStr(v(i, PS_CREW))
            p.Hours = NumOr0(v(i, PS_HOURS))
            p.Positions = CLng(NumOr0(v(i, PS_POS)))
            p.Reason = SafeStr(v(i, PS_REASON))
            found = True
            Exit For
        End If
    Next i
    If Not found Then Exit Function

    v = RangeTo2D(lr.DataBodyRange)
    Dim n As Long
    ReDim lst(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PL_ID)))) = UCase$(id) Then
            n = n + 1
            With lst(n)
                .Rank = CLng(NumOr0(v(i, PL_RANK)))
                .EmpName = SafeStr(v(i, PL_EMP))
                .EmpShift = SafeStr(v(i, PL_SHIFT))
                .DeptCode = SafeStr(v(i, PL_CODE))
                .Grade = DeptCodeValue(.DeptCode)
                .TotalHrs = NumOr0(v(i, PL_HRS))
                .SignedUp = (UCase$(SafeStr(v(i, PL_SIGNED))) = "YES")
                .Result = SafeStr(v(i, PL_RESULT))
                SplitExempt .Result, .ExemptNote          ' "Signed up | Exempt - moved Jan 12" -> the two parts
                If UBound(v, 2) >= PL_SIGNHRS Then .SignupHrs = NumOr0(v(i, PL_SIGNHRS))
                .Awarded = (.Result = "Awarded")
                .Charged = NumOr0(v(i, PL_CHARGED))
            End With
        End If
    Next i
    If n = 0 Then Exit Function
    ReDim Preserve lst(1 To n)

    ' rank order
    Dim j As Long, tmp As PListRec
    For i = 2 To n
        tmp = lst(i): j = i - 1
        Do While j >= 1
            If lst(j).Rank > tmp.Rank Then
                lst(j + 1) = lst(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        lst(j + 1) = tmp
    Next i
    LoadPosting = n
End Function

' The Result cell of a list row can carry "Exempt - ..." on its own or after
' the result ("Signed up | Exempt - moved Jan 12"), so marking an exempt
' person signed up or awarded never loses the exemption.
Public Sub SplitExempt(ByRef result As String, ByRef exemptNote As String)
    Dim p As Long: p = InStr(1, result, EXEMPT_TAG, vbTextCompare)
    exemptNote = ""
    If p = 0 Then Exit Sub
    exemptNote = Trim$(Mid$(result, p))
    result = Trim$(Left$(result, p - 1))
    If Right$(result, 1) = "|" Then result = Trim$(Left$(result, Len(result) - 1))
End Sub


' Hours as text without a trailing point: 8 -> "8", 22.75 -> "22.75" (Format$(8, "0.##") gives "8.")
Public Function HoursText(ByVal x As Double) As String
    HoursText = CStr(Round(x, 2))
End Function

' Open postings for the award pop-up. ids() gets the Posting IDs.
Public Function OpenPostingChoices(ByRef ids As Variant) As Variant
    ids = Empty
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, cnt As Long
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, PS_STATUS)) = "Open" Then cnt = cnt + 1
    Next i
    If cnt = 0 Then Exit Function

    Dim d() As String, idArr() As String, k As Long, closes As Date
    ReDim d(0 To cnt - 1): ReDim idArr(0 To cnt - 1)
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, PS_STATUS)) = "Open" Then
            closes = ToDate(v(i, PS_CLOSES))
            idArr(k) = SafeStr(v(i, PS_ID))
            d(k) = idArr(k) & "   " & Format$(ToDate(v(i, PS_DATE)), "mmm d") & "   " & _
                   SafeStr(v(i, PS_UNIT)) & "   " & SafeStr(v(i, PS_CREW)) & "   - " & _
                   IIf(PostingClosed(closes), "READY TO AWARD", "open until " & Format$(closes, "mmm d"))
            k = k + 1
        End If
    Next i
    ids = idArr
    OpenPostingChoices = d
End Function

' Every posting, newest first, for the Load Posting pop-up.
Public Function AllPostingChoices(ByRef ids As Variant) As Variant
    ids = Empty
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, cnt As Long
    For i = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(i, PS_ID))) <> "" Then cnt = cnt + 1
    Next i
    If cnt = 0 Then Exit Function

    ' order by posting number, newest first
    Dim ord() As Long, num() As Long, j As Long, k As Long, t As Long
    ReDim ord(1 To cnt): ReDim num(1 To cnt)
    For i = 1 To UBound(v, 1)
        If Trim$(SafeStr(v(i, PS_ID))) <> "" Then
            k = k + 1
            ord(k) = i
            num(k) = CLng(Val(Mid$(Trim$(SafeStr(v(i, PS_ID))), 3)))
        End If
    Next i
    For i = 2 To cnt
        t = ord(i): k = num(i): j = i - 1
        Do While j >= 1
            If num(j) < k Then
                ord(j + 1) = ord(j): num(j + 1) = num(j): j = j - 1
            Else
                Exit Do
            End If
        Loop
        ord(j + 1) = t: num(j + 1) = k
    Next i

    Dim d() As String, idArr() As String, st As String, r As Long
    ReDim d(0 To cnt - 1): ReDim idArr(0 To cnt - 1)
    For i = 1 To cnt
        r = ord(i)
        idArr(i - 1) = SafeStr(v(r, PS_ID))
        st = SafeStr(v(r, PS_STATUS))
        If st = "Open" Then
            If PostingClosed(ToDate(v(r, PS_CLOSES))) Then st = "Open - ready to award" Else st = "Open until " & Format$(ToDate(v(r, PS_CLOSES)), "mmm d")
        End If
        d(i - 1) = idArr(i - 1) & "   " & Format$(ToDate(v(r, PS_DATE)), "ddd mmm d") & "   " & _
                   SafeStr(v(r, PS_UNIT)) & "   " & SafeStr(v(r, PS_CREW)) & "   (" & st & ")"
    Next i
    ids = idArr
    AllPostingChoices = d
End Function

' Posting ID -> Shift (Days / Nights), used when upgrading the OT Log layout
Public Function PostingShiftMap() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set PostingShiftMap = d
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long, k As String
    For i = 1 To UBound(v, 1)
        k = Trim$(SafeStr(v(i, PS_ID)))
        If k <> "" Then d(k) = SafeStr(v(i, PS_SHIFT))
    Next i
End Function

Public Function PostingStatus(ByVal id As String) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(id) Then PostingStatus = SafeStr(v(i, PS_STATUS)): Exit Function
    Next i
End Function

' Next P- number across the Postings table, its list table and the OT Log.
' v24.3.1 (review P1): list rows are written before the header, so a header deleted by hand
' or a creation that failed between the two leaves list rows whose ID must never be reused
' (LoadPosting would collect both lists and the award would charge across them).
Private Function NextPostingID() As String
    Dim s As String: s = NextRefID("P")
    Dim mx As Long: mx = CLng(Val(Mid$(s, 3))) - 1
    RaiseToMaxPostingNumber GetTable(PT_TABLE), PS_ID, mx
    RaiseToMaxPostingNumber GetTable(PL_TABLE), PL_ID, mx
    NextPostingID = "P-" & Format$(mx + 1, "0000")
End Function

' mx = the highest P-#### number in that table's ID column, if higher than mx already is
Private Sub RaiseToMaxPostingNumber(ByVal lo As ListObject, ByVal idCol As Long, ByRef mx As Long)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim v As Variant, i As Long, k As Long, s As String
    v = RangeTo2D(lo.ListColumns(idCol).DataBodyRange)
    For i = 1 To UBound(v, 1)
        s = UCase$(Trim$(SafeStr(v(i, 1))))
        If Left$(s, 2) = "P-" Then
            k = CLng(Val(Mid$(s, 3)))
            If k > mx Then mx = k
        End If
    Next i
End Sub


' ====================================================================
'  AWARD / NO TAKERS / CANCEL
'  Charging (refusal = hours x CHARGE_MULTIPLIER, grade limit applies):
'    Award     - anyone AHEAD of the awarded person in the posted call
'                order (grade, then hours, then seniority) who did NOT
'                sign up is charged. People below the awarded person,
'                people who signed up, and the awarded are not charged.
'                With 2+ positions "ahead" means ahead of the LAST
'                awarded person in call order.
'    No takers - nobody took it: everyone eligible on the list is charged.
'    Cancel    - the request was withdrawn: nobody is charged.
'  Charges go in the pay period of the SHIFT DATE.
'  Partial sign-ups (v24): someone can sign up for part of the shift.
'    Full-shift sign-ups are awarded first; a partial only when no
'    full-shift sign-up is left for the position (AwardOrderProblem).
'    Charging doesn't change: people ahead of the last awarded person
'    who didn't sign up AT ALL are charged the posting's FULL hours;
'    a partial sign-up that isn't awarded is not charged; hours nobody
'    covers drop off, uncharged.
' ====================================================================
' Works out Result / Charged for every person on the list (Award mode).
' maxRank / maxHrs = call-order position and hours of the last awarded person.
Public Sub ComputeAward(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                        ByRef maxRank As Long, ByRef maxHrs As Double, ByRef nAwarded As Long, _
                        ByRef nCharged As Long, ByRef totCharged As Double)
    Dim i As Long
    maxRank = 0: maxHrs = 0: nAwarded = 0: nCharged = 0: totCharged = 0
    For i = 1 To n
        If lst(i).Awarded Then
            nAwarded = nAwarded + 1
            If lst(i).Rank > maxRank Then maxRank = lst(i).Rank: maxHrs = lst(i).TotalHrs
        End If
    Next i
    For i = 1 To n
        lst(i).Charged = 0
        If lst(i).Awarded Then
            lst(i).Result = "Awarded"
        ElseIf lst(i).SignedUp Then
            lst(i).Result = "Signed up"
        ElseIf lst(i).ExemptNote <> "" Then
            lst(i).Result = lst(i).ExemptNote             ' moved unit or left before the shift (v24)
        ElseIf nAwarded > 0 And lst(i).Rank < maxRank And lst(i).Grade <= CHARGE_MAX_GRADE Then
            ' v26.5.5 (the 16 h rule, review): charged on the hours they could have taken, never on hours they may not work
            lst(i).Charged = CappedPostingHours(p.Hours, lst(i).MaxHours, lst(i).MaxSet) * CHARGE_MULTIPLIER
            If lst(i).Charged > 0 Then
                lst(i).Result = "Charged"
                nCharged = nCharged + 1
                totCharged = totCharged + lst(i).Charged
            Else
                lst(i).Result = "16 h rule"                ' could not take any of it: nothing to refuse
            End If
        Else
            lst(i).Result = ""
        End If
    Next i
End Sub

' (re-review) the confirms' "N x 21 hrs" line is wrong when the 16 h rule capped someone: this names them
Public Function CapNote(ByRef lst() As PListRec, ByVal n As Long, ByVal postHours As Double) As String
    Dim i As Long, nCap As Long
    For i = 1 To n
        If lst(i).Charged > 0 And lst(i).MaxSet And lst(i).MaxHours < postHours - 0.001 Then nCap = nCap + 1
    Next i
    If nCap > 0 Then CapNote = " (the 16 h rule caps " & nCap & " of them at their own hours)"
End Function

' PURE (SelfTest "postcap"): the hours a listed person is charged on - the posting's, or their 16 h rule
' maximum when the award pop-up worked one out (MaxSet) and it is lower; 0 = nothing (they could take none)
Public Function CappedPostingHours(ByVal postHours As Double, ByVal maxHours As Double, ByVal maxSet As Boolean) As Double
    CappedPostingHours = postHours
    If maxSet And maxHours < postHours Then CappedPostingHours = IIf(maxHours > 0, maxHours, 0)
End Function

' ---- partial sign-ups (v24) ----
' A sign-up for only part of the shift: more than 0 and fewer hours than the posting
Public Function IsPartialSignup(ByVal postHours As Double, ByVal signupHrs As Double) As Boolean
    IsPartialSignup = (signupHrs > 0 And signupHrs < postHours)
End Function

' Hours one awarded person takes: the whole shift, or the part they signed up for
Public Function AwardedHours(ByVal postHours As Double, ByVal signupHrs As Double) As Double
    If IsPartialSignup(postHours, signupHrs) Then AwardedHours = signupHrs Else AwardedHours = postHours
End Function

' The decided order: full-shift sign-ups first; a partial is awarded only when
' no full-shift sign-up is left for that position. Counted per position: with
' 2 positions and 1 full sign-up, partials may take the second.
'   ""           fine (or no partial is awarded)
'   "filled"     every position already has a full-shift award
'   "full-first" a partial is awarded while a full-shift sign-up isn't (who = that person)
'   "over-hours" the awarded partials add up to more than the hours left
Public Function AwardOrderProblem(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                                  Optional ByRef who As String) As String
    Dim i As Long, nFullAw As Long, nPartAw As Long, partHrs As Double, nFullWaiting As Long, pos As Long
    who = ""
    For i = 1 To n
        If lst(i).Awarded Then
            If IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
                nPartAw = nPartAw + 1
                partHrs = partHrs + lst(i).SignupHrs
            Else
                nFullAw = nFullAw + 1
            End If
        ElseIf lst(i).SignedUp And Not IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
            ' v26.5.5 (review): a full-shift sign-up the 16 h rule forbids is not "waiting" - it can never be awarded whole
            If Not (lst(i).MaxSet And lst(i).MaxHours < p.Hours - 0.001) Then
                nFullWaiting = nFullWaiting + 1
                If who = "" Then who = "#" & lst(i).Rank & " " & lst(i).EmpName
            End If
        End If
    Next i
    If nPartAw = 0 Then who = "": Exit Function
    pos = p.Positions
    If pos < 1 Then pos = 1
    If nFullAw >= pos Then                      ' nothing left for a partial: say so before pointing at anyone
        who = "": AwardOrderProblem = "filled"
    ElseIf nFullWaiting > 0 Then
        AwardOrderProblem = "full-first"
    ElseIf partHrs > (pos - nFullAw) * p.Hours + 0.0001 Then
        who = "": AwardOrderProblem = "over-hours"
    Else
        who = ""
    End If
End Function

' Hours the awards cover, and what is left of the posting. The remainder
' drops off: nobody is charged for it.
Public Sub AwardCoverage(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                         ByRef covered As Double, ByRef uncovered As Double)
    Dim i As Long, pos As Long
    covered = 0
    For i = 1 To n
        If lst(i).Awarded Then covered = covered + AwardedHours(p.Hours, lst(i).SignupHrs)
    Next i
    pos = p.Positions
    If pos < 1 Then pos = 1
    uncovered = pos * p.Hours - covered
    If uncovered < 0 Then uncovered = 0
End Sub

' No takers: everyone eligible is charged (nobody should be marked signed up).
Public Sub ComputeNoTakers(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                           ByRef nCharged As Long, ByRef totCharged As Double)
    Dim i As Long
    nCharged = 0: totCharged = 0
    For i = 1 To n
        lst(i).Awarded = False
        lst(i).Charged = 0
        lst(i).Result = lst(i).ExemptNote
        If lst(i).ExemptNote = "" And lst(i).Grade <= CHARGE_MAX_GRADE Then
            lst(i).Charged = CappedPostingHours(p.Hours, lst(i).MaxHours, lst(i).MaxSet) * CHARGE_MULTIPLIER   ' v26.5.5: the 16 h rule caps it
            If lst(i).Charged > 0 Then
                lst(i).Result = "Charged"
                nCharged = nCharged + 1
                totCharged = totCharged + lst(i).Charged
            Else
                lst(i).Result = "16 h rule"
            End If
        End If
    Next i
End Sub

' Awarded out of department: in-department people who didn't sign up are
' charged as for no takers; anyone who signed up isn't (v21: pulled out of
' SubmitAwardOutside so SelfTest can check it).
Public Sub ComputeOutside(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                          ByRef nCharged As Long, ByRef totCharged As Double)
    Dim i As Long
    nCharged = 0: totCharged = 0
    For i = 1 To n
        lst(i).Awarded = False
        lst(i).Charged = 0
        lst(i).Result = lst(i).ExemptNote
        If lst(i).SignedUp Then
            lst(i).Result = "Signed up"
        ElseIf lst(i).ExemptNote = "" And lst(i).Grade <= CHARGE_MAX_GRADE Then
            lst(i).Charged = CappedPostingHours(p.Hours, lst(i).MaxHours, lst(i).MaxSet) * CHARGE_MULTIPLIER   ' v26.5.5: the 16 h rule caps it
            If lst(i).Charged > 0 Then
                lst(i).Result = "Charged"
                nCharged = nCharged + 1
                totCharged = totCharged + lst(i).Charged
            Else
                lst(i).Result = "16 h rule"
            End If
        End If
    Next i
End Sub

' Cancelled: nobody is awarded and nobody is charged.
Public Sub ComputeCancel(ByRef lst() As PListRec, ByVal n As Long)
    Dim i As Long
    For i = 1 To n
        lst(i).Awarded = False
        lst(i).Charged = 0
        lst(i).Result = lst(i).ExemptNote
    Next i
End Sub

' Checks shared by Award / No takers / Cancel
' v25 (P6): THE closing test. A posting closes at POSTING_CLOSE_TIME on its closing date - the moment
' the printed sheet, the members' file and the sign-up inbox already used (v25 review: the master
' said "the day after", so a mover could still be added to a list at 2 PM on the closing day).
' asOf is a date-time (Now); SelfTest passes a date, which is that day's midnight. Pure (SelfTest
' postclosed). Every "ready to award" / "still open" / "can still sign up" asks this.
Public Function PostingClosed(ByVal closesOn As Date, Optional ByVal asOf As Date = 0) As Boolean
    If asOf = 0 Then asOf = Now
    If closesOn = 0 Then Exit Function
    PostingClosed = (asOf >= PostingCloseAt(closesOn))
End Function

Private Function CanFinalize(ByRef p As PostRec) As Boolean
    If Not RequireSupervisor() Then Exit Function     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not RequireWritable() Then Exit Function
    If PostingStatus(p.ID) <> "Open" Then
        MsgBox p.ID & " is no longer open (" & PostingStatus(p.ID) & ").", vbExclamation, "Call Out Roster"
        Exit Function
    End If
    If HasActiveLogRows(p.ID) Then
        MsgBox p.ID & " already has entries in the OT Log, so it can't be closed again." & vbNewLine & _
               "If they're wrong, an admin voids them (Void Entry > " & p.ID & ") and reopens the posting.", vbExclamation, "Call Out Roster"
        Exit Function
    End If
    CanFinalize = True
End Function

Private Function EarlyOK(ByRef p As PostRec, ByVal what As String) As Boolean
    EarlyOK = True
    If Not PostingClosed(p.ClosesOn) Then
        EarlyOK = (Ask(p.ID & " doesn't close until " & Format$(p.ClosesOn, "mmm d") & " @ " & POSTING_CLOSE_TIME & _
                   ". " & what & " it now anyway?", vbYesNo + vbExclamation, what, vbYes) = vbYes)
    End If
End Function

' ---------------- AWARD ----------------
Public Function SubmitAward(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long) As Boolean
    If Not CanFinalize(p) Then Exit Function

    Dim maxRank As Long, maxHrs As Double, nAw As Long, nCh As Long, tot As Double
    ComputeAward p, lst, n, maxRank, maxHrs, nAw, nCh, tot
    If nAw = 0 Then
        ' v26.5.5 (B10): the two closings are an admin's buttons; a supervisor is told whom to ask
        If IsAdmin() Then
            MsgBox "Nobody is awarded." & vbNewLine & vbNewLine & _
                   "If nobody took it, use NO TAKERS (charges everyone eligible)." & vbNewLine & _
                   "If the overtime isn't needed any more, use CANCEL REQUEST (no charges).", vbExclamation, "Award"
        Else
            MsgBox "Nobody is awarded." & vbNewLine & vbNewLine & _
                   "If nobody took it, or the overtime isn't needed any more, ask " & PRIMARY_ADMIN_NAME & " to close the posting " & _
                   "(No Takers charges everyone eligible; Cancel Request charges nobody).", vbExclamation, "Award"
        End If
        Exit Function
    End If

    ' v24: full-shift sign-ups first; a partial only when no full-shift sign-up is left
    Dim who As String, anyPartial As Boolean, covered As Double, uncovered As Double
    Select Case AwardOrderProblem(p, lst, n, who)
        Case "full-first"
            MsgBox who & " signed up for the WHOLE shift and isn't awarded, but a partial sign-up is." & vbNewLine & vbNewLine & _
                   "A full-shift sign-up comes first, even with more hours. Award " & who & " instead." & vbNewLine & _
                   "If they withdrew, take their sign-up off (double-click the name), then submit again.", vbExclamation, "Award"
            Exit Function
        Case "filled"
            MsgBox "Every position is already awarded to someone taking the whole shift, so a partial sign-up can't be awarded as well." & vbNewLine & _
                   "Take the AWARD mark off the partial sign-up.", vbExclamation, "Award"
            Exit Function
        Case "over-hours"
            MsgBox "The partial sign-ups marked AWARDED add up to more hours than the posting has left." & vbNewLine & _
                   "Take one off, or correct the hours with PARTIAL HOURS.", vbExclamation, "Award"
            Exit Function
    End Select
    AwardCoverage p, lst, n, covered, uncovered

    Dim i As Long, awardNames As String, awardShow As String, skipped As String
    For i = 1 To n
        If lst(i).Awarded Then
            awardNames = awardNames & IIf(awardNames = "", "", "; ") & lst(i).EmpName
            awardShow = awardShow & IIf(awardShow = "", "", "; ") & lst(i).EmpName
            If IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
                anyPartial = True
                awardShow = awardShow & " (" & HoursText(lst(i).SignupHrs) & " h)"
            End If
        End If
        If lst(i).SignedUp And Not lst(i).Awarded And lst(i).Rank < maxRank Then
            skipped = skipped & "  #" & lst(i).Rank & "  " & lst(i).EmpName & vbNewLine
        End If
    Next i

    If skipped <> "" Then
        If Ask("These people signed up and are AHEAD of the awarded person in call order, " & _
               "but aren't awarded:" & vbNewLine & vbNewLine & skipped & vbNewLine & _
               "They won't be charged. Continue?", vbYesNo + vbExclamation, "Award", vbYes) = vbNo Then Exit Function
    End If
    If nAw <> p.Positions And Not anyPartial Then
        If Ask(nAw & " awarded but the posting is for " & p.Positions & " position(s). Continue?", _
               vbYesNo + vbExclamation, "Award", vbYes) = vbNo Then Exit Function
    End If
    If Not EarlyOK(p, "Award") Then Exit Function

    Dim ps As Date: ps = PeriodStartFor(p.ShiftDate)
    If nCh > 0 Then
        If Not ConfirmPeriod(ps) Then Exit Function
    End If

    Dim msg As String
    msg = p.ID & "  -  " & p.Unit & "  -  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & vbNewLine & vbNewLine & _
          "Awarded: " & awardShow & vbNewLine & _
          IIf(anyPartial, "Covered: " & HoursText(covered) & " h" & _
              IIf(uncovered > 0, "; the other " & HoursText(uncovered) & " h are not covered and nobody is charged for them", "") & vbNewLine, "") & _
          "Charged (didn't sign up, ahead in call order): " & nCh & " x " & _
          HoursText(p.Hours * CHARGE_MULTIPLIER) & " hrs = " & HoursText(tot) & " hrs" & CapNote(lst, n, p.Hours) & _
          IIf(nCh > 0, vbNewLine & "Pay period starting " & Format$(ps, "mmm d, yyyy"), "")
    If Ask(msg & vbNewLine & vbNewLine & "Submit?", vbOKCancel + vbQuestion, "Confirm award", vbOK) <> vbOK Then Exit Function

    If Not WriteOutcome(p, lst, n, "Awarded", awardNames, ps, maxRank) Then Exit Function
    Dim calNames As String
    For i = 1 To n
        If lst(i).Awarded Then
            calNames = calNames & IIf(calNames = "", "", " / ") & CalName(lst(i).EmpName)
            ' a partial award shows its hours beside the name: "Joe D 8h"
            If IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
                calNames = calNames & " " & CDbl(lst(i).SignupHrs) & "h"
            ElseIf anyPartial And p.Hours < USUAL_SHIFT_HOURS Then
                calNames = calNames & " " & CDbl(p.Hours) & "h"        ' a short posting: the whole-shift person keeps their tag too
            End If
        End If
    Next i
    Dim calTxt As String: calTxt = CalendarOutcome(p, calNames, anyPartial)
    If Not gSmokeRun Then EmailAward p, lst, n, awardShow      ' v24.2: SmokeTest opens no email draft
    ' v27 (item 7): the receipt instead of the box (WriteOutcome filled WHAT WAS RECORDED and WHO IT CHANGED)
    ShowReceiptNow p.ID & " awarded to " & awardShow, gReceipt.Recorded, gReceipt.Changed, _
                   "  Calendar: " & calTxt & vbNewLine & "  Postings sheet: " & p.ID & " is Awarded; the award email " & IIf(gSmokeRun, "was not sent (test)", "went to the crew list") & "." & vbNewLine & _
                   "  OT Log: the rows with Ref ID " & p.ID & " (the charges carry the pay period starting " & Format$(ps, "mmm d") & ")." & vbNewLine & _
                   "  The members' file shows it after the next Update the members' file." & vbNewLine, WrongText() & "  Reopen Posting (admin) opens it again and marks the charged people not chargeable." & vbNewLine
    SubmitAward = True
End Function

' ---------------- AWARD OUT OF DEPARTMENT ----------------
' Nobody in the department took it; someone from another department signed
' a blank line. In-department operators who didn't sign up are charged as
' for NO TAKERS; anyone who signed up isn't charged.
Public Function SubmitAwardOutside(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                                   ByVal outName As String, Optional ByVal outTake As Double = 0) As Boolean
    ' outTake (v24.1): hours the out-of-department person takes when it is only part of the shift (0 = all)
    If Not RequireAdmin() Then Exit Function         ' v25.3: an admin's job (a supervisor awards to the list only)
    If Not CanFinalize(p) Then Exit Function
    outName = Trim$(outName)
    Dim outPartial As Boolean: outPartial = IsPartialSignup(p.Hours, outTake)
    If outName = "" Then MsgBox "Pick or type the out-of-department name first.", vbExclamation, "Call Out Roster": Exit Function
    Dim i As Long, marked As String, signedIn As String
    For i = 1 To n
        If StrComp(Trim$(lst(i).EmpName), outName, vbTextCompare) = 0 Then
            MsgBox outName & " is on this posting's list - select them and use AWARD instead.", vbExclamation, "Call Out Roster"
            Exit Function
        End If
        If lst(i).Awarded Then marked = marked & "  " & lst(i).EmpName & vbNewLine
        If lst(i).SignedUp And Not lst(i).Awarded Then signedIn = signedIn & "  #" & lst(i).Rank & "  " & lst(i).EmpName & vbNewLine
    Next i
    If marked <> "" Then
        MsgBox "These in-department people are marked AWARDED:" & vbNewLine & vbNewLine & marked & vbNewLine & _
               "Use SUBMIT AWARD for them, or clear their award marks first.", vbExclamation, "Out of department"
        Exit Function
    End If
    If signedIn <> "" And outPartial Then
        ' the same order as in-department awards: a partial only when nobody else is left
        MsgBox "These in-department people signed up:" & vbNewLine & vbNewLine & signedIn & vbNewLine & _
               "An out-of-department person can take PART of the shift only when nobody in the department signed up." & vbNewLine & _
               "Award one of them with SUBMIT AWARD, or leave the hours box blank to award the whole shift out of department.", _
               vbExclamation, "Out of department"
        Exit Function
    End If
    If signedIn <> "" Then
        If Ask("These in-department people signed up and would normally get it first:" & vbNewLine & vbNewLine & _
               signedIn & vbNewLine & "Award it out of department anyway? (They won't be charged.)", _
               vbYesNo + vbExclamation, "Out of department", vbYes) = vbNo Then Exit Function
    End If
    If Not EarlyOK(p, "Award") Then Exit Function

    Dim nCh As Long, tot As Double
    ComputeOutside p, lst, n, nCh, tot
    Dim ps As Date: ps = PeriodStartFor(p.ShiftDate)
    If nCh > 0 Then
        If Not ConfirmPeriod(ps) Then Exit Function
    End If

    Dim code As String, hrs As Double, onRoster As Boolean, inUnit As Boolean, crew As String, why As String
    onRoster = RosterLookup(outName, code, hrs, p.Unit, inUnit, crew)
    If onRoster Then
        crew = CrewOn(outName, p.ShiftDate, crew)          ' v24: the crew they are on that date
        code = "OD " & crew
        ' v24.3: a unit move that is waiting counts from its date
        inUnit = (UnitCodeOn(outName, p.Unit, p.ShiftDate, IIf(inUnit, "x", "")) <> "")
    End If
    If onRoster And Not inUnit Then
        If Ask(outName & " isn't signed off in " & p.Unit & " (no code in that column on the OT Roster)." & vbNewLine & _
               "Award anyway?", vbYesNo + vbExclamation, "Out of department", vbYes) = vbNo Then Exit Function
    End If
    If onRoster Then why = NotFreeReason(outName, crew, p.ShiftDate, p.ShiftLbl)
    If onRoster And why = "" Then
        If OtRestrictedOn(outName, p.ShiftDate) Then why = "OT restricted"     ' v27 (item 21, review H1)
    End If
    If why <> "" Then
        If Ask(outName & " isn't free that shift (" & why & ")." & vbNewLine & _
               "Award anyway?", vbYesNo + vbExclamation, "Out of department", vbYes) = vbNo Then Exit Function
    End If
    If Not onRoster Then
        If Ask(outName & " isn't on the OT Roster, so no email address or hours are on file." & vbNewLine & _
               "Award anyway?", vbYesNo + vbExclamation, "Out of department", vbYes) = vbNo Then Exit Function
    End If

    If Ask(p.ID & "  -  " & p.Unit & "  -  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & vbNewLine & vbNewLine & _
              "Award OUT OF DEPARTMENT to: " & outName & IIf(code <> "", "  (" & code & ")", "") & _
              IIf(outPartial, "  for " & HoursText(outTake) & " of " & HoursText(p.Hours) & " h (the rest is not covered; nobody is charged for it)", "") & vbNewLine & _
              "Charged (in department, didn't sign up): " & nCh & " x " & HoursText(p.Hours * CHARGE_MULTIPLIER) & _
              " hrs = " & HoursText(tot) & " hrs" & CapNote(lst, n, p.Hours) & IIf(nCh > 0, vbNewLine & "Pay period starting " & Format$(ps, "mmm d, yyyy"), "") & _
              vbNewLine & vbNewLine & "Submit?", vbOKCancel + vbQuestion, "Confirm out-of-department award", vbOK) <> vbOK Then Exit Function

    Dim awardTxt As String: awardTxt = outName & " (out of dept)"
    If Not WriteOutcome(p, lst, n, "Awarded", awardTxt, ps, 0, outName, IIf(code = "", "OD", code), hrs, outTake) Then Exit Function
    Dim calTxt As String
    If outPartial Then
        calTxt = CalendarOutcome(p, CalName(outName) & " " & CDbl(outTake) & "h", True)
    Else
        calTxt = CalendarOutcome(p, CalName(outName))
    End If
    If Not gSmokeRun Then EmailAward p, lst, n, awardTxt & IIf(outPartial, " (" & HoursText(outTake) & " h)", ""), outName
    ShowReceiptNow p.ID & " awarded out of department to " & outName, gReceipt.Recorded, gReceipt.Changed, _
                   "  Calendar: " & calTxt & vbNewLine & "  Postings sheet: " & p.ID & " is Awarded (out of department)." & vbNewLine & _
                   "  OT Log: the rows with Ref ID " & p.ID & " (everyone eligible in the unit who did not sign up is charged, pay period starting " & Format$(ps, "mmm d") & ")." & vbNewLine & _
                   "  The members' file shows it after the next Update the members' file." & vbNewLine, WrongText() & "  Reopen Posting (admin) opens it again." & vbNewLine
    SubmitAwardOutside = True
End Function

' Out-of-department candidates for a posting: signed off in its unit (a code in
' that unit's column), NOT on the posting's list, and free that shift (not
' scheduled, rest rule, booked off or on a switch). 2-D array (rows x 4):
' name, unit code, crew, total hours (now) - lowest hours first. Empty if nobody.
Public Function OutsideCandidates(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long) As Variant
    Dim recs() As EmpRec, k As Long, i As Long, j As Long, onList As Boolean, why As String
    k = LoadEligible(p.Unit, "ALL", recs, 999, True, p.ShiftDate)
    If k > 0 Then k = FilterNotFree(recs, k, p.ShiftDate, p.ShiftLbl, why)
    If k = 0 Then Exit Function
    Dim keep() As Long, m As Long
    ReDim keep(1 To k)
    For i = 1 To k
        onList = False
        For j = 1 To n
            If StrComp(Trim$(lst(j).EmpName), Trim$(recs(i).EmpName), vbTextCompare) = 0 Then onList = True: Exit For
        Next j
        If Not onList Then m = m + 1: keep(m) = i
    Next i
    If m = 0 Then Exit Function
    ' v25 (D-E3): THE call order (grade, then hours, then seniority: SortEmpRecs), not a hours-only re-sort
    Dim sub_() As EmpRec: ReDim sub_(1 To m)
    For i = 1 To m
        sub_(i) = recs(keep(i))
    Next i
    SortEmpRecs sub_, m
    Dim out() As String: ReDim out(0 To m - 1, 0 To 3)
    For i = 1 To m
        out(i - 1, 0) = sub_(i).EmpName
        out(i - 1, 1) = sub_(i).DeptCode
        out(i - 1, 2) = sub_(i).EmpShift
        out(i - 1, 3) = Format$(sub_(i).TotalHrs, "#,##0")
    Next i
    OutsideCandidates = out
End Function

' Roster row for a name: dept code (Shift column text) and total hours. False if not on the roster.
Private Function RosterLookup(ByVal empName As String, ByRef code As String, ByRef hrs As Double, _
                              Optional ByVal unit As String = "", Optional ByRef inUnit As Boolean, _
                              Optional ByRef crew As String) As Boolean
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), empName, vbTextCompare) = 0 Then
            crew = Trim$(SafeStr(ws.Cells(r, C_SHIFT).Value))
            code = "OD " & crew
            If IsNumeric(ws.Cells(r, C_TOTAL).Value) Then hrs = CDbl(ws.Cells(r, C_TOTAL).Value)
            If unit <> "" Then
                If DeptCol(unit) > 0 Then inUnit = (Trim$(SafeStr(ws.Cells(r, DeptCol(unit)).Value)) <> "")
            End If
            RosterLookup = True
            Exit Function
        End If
    Next r
End Function

' ---------------- NO TAKERS ----------------
Public Function SubmitNoTakers(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long) As Boolean
    If Not RequireAdmin() Then Exit Function         ' v25.3: an admin's job (it charges the whole list)
    If Not CanFinalize(p) Then Exit Function
    Dim i As Long, marked As String
    For i = 1 To n
        If lst(i).SignedUp Or lst(i).Awarded Then marked = marked & "  " & lst(i).EmpName & vbNewLine
    Next i
    If marked <> "" Then
        MsgBox "These people are marked as signed up / awarded:" & vbNewLine & vbNewLine & marked & vbNewLine & _
               "If somebody took it, use SUBMIT AWARD. Otherwise clear their marks first.", vbExclamation, "No takers"
        Exit Function
    End If
    If Not EarlyOK(p, "Close") Then Exit Function

    Dim nCh As Long, tot As Double
    ComputeNoTakers p, lst, n, nCh, tot
    Dim ps As Date: ps = PeriodStartFor(p.ShiftDate)
    If nCh > 0 Then
        If Not ConfirmPeriod(ps) Then Exit Function
    End If
    If Ask(p.ID & "  -  " & p.Unit & "  -  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & vbNewLine & vbNewLine & _
              "NO TAKERS: charge all " & nCh & " eligible operator(s) " & HoursText(p.Hours * CHARGE_MULTIPLIER) & _
              " hrs each (" & HoursText(tot) & " hrs total)" & CapNote(lst, n, p.Hours) & vbNewLine & _
              "Pay period starting " & Format$(ps, "mmm d, yyyy") & vbNewLine & vbNewLine & "Submit?", _
              vbOKCancel + vbExclamation, "Confirm no takers", vbOK) <> vbOK Then Exit Function
    If Not WriteOutcome(p, lst, n, "No takers", "", ps, 0) Then Exit Function
    Dim ntCal As String: ntCal = CalendarOutcome(p, "No takers")
    ShowReceiptNow p.ID & " closed - no takers", gReceipt.Recorded, gReceipt.Changed, _
                   "  Calendar: " & ntCal & vbNewLine & "  Postings sheet: " & p.ID & " is No takers." & vbNewLine & _
                   "  OT Log: one Refused row per eligible person (pay period starting " & Format$(ps, "mmm d") & ")." & vbNewLine & _
                   "  The members' file shows it after the next Update the members' file." & vbNewLine, WrongText() & "  Reopen Posting (admin) opens it again." & vbNewLine
    SubmitNoTakers = True
End Function

' ---------------- v25: the Award pop-up's rules, out of the form ----------------
' Signed up on / off. Off also clears the award, the partial hours and the electronic sign-up.
Public Sub SetSignedUp(ByRef lst() As PListRec, ByVal i As Long, ByVal onNow As Boolean)
    lst(i).SignedUp = onNow
    If Not onNow Then lst(i).Awarded = False: lst(i).SignupHrs = 0: lst(i).ESign = ""
End Sub

' Awarded on / off. The winner signed up by definition.
Public Sub SetAwarded(ByRef lst() As PListRec, ByVal i As Long, ByVal onNow As Boolean)
    lst(i).Awarded = onNow
    If onNow Then lst(i).SignedUp = True
End Sub

' The hours typed for a partial sign-up or an out-of-department award, against the posting's
' hours: "" = fine and hrs holds them (0 = the whole shift: blank, or the posting's hours or
' more); otherwise why it is refused. Pure (SelfTest partialhrs).
Public Function PartialHoursValue(ByVal txt As String, ByVal postHours As Double, ByRef hrs As Double) As String
    hrs = 0
    txt = Trim$(txt)
    If txt = "" Then Exit Function
    If Not IsNumeric(txt) Then PartialHoursValue = "Enter a number of hours, e.g. 8 - or leave it blank for the whole shift.": Exit Function
    hrs = CDbl(txt)
    If hrs <= 0 Then hrs = 0: PartialHoursValue = "The hours must be more than 0 - or blank for the whole shift.": Exit Function
    If hrs >= postHours Then hrs = 0
End Function

' v26.5.5 (item 9): the hours a listed person would take against their 16 h rule maximum: "" = fine,
' otherwise why it is refused. A list whose maximum was never worked out (MaxSet False) has no limit.
Public Function AwardMaxProblem(ByRef lst() As PListRec, ByVal i As Long, ByVal take As Double) As String
    If Not lst(i).MaxSet Then Exit Function
    If take > lst(i).MaxHours + 0.001 Then
        AwardMaxProblem = lst(i).EmpName & " can take at most " & HoursText(lst(i).MaxHours) & " h of this shift (the 16 h rule: " & _
                          "OT accepted, an award or a switch touches it). A partial sign-up at or under that is allowed."
    End If
End Function

' The selected person signed up for part of the shift (0 = the whole of it). False = refused (why).
Public Function SetPartialHours(ByRef lst() As PListRec, ByVal i As Long, ByVal txt As String, ByVal postHours As Double, _
                                ByRef why As String) As Boolean
    Dim h As Double
    why = PartialHoursValue(txt, postHours, h)
    If why <> "" Then Exit Function
    why = AwardMaxProblem(lst, i, IIf(h > 0, h, postHours))      ' v26.5.5 (item 9)
    If why <> "" Then Exit Function
    lst(i).SignedUp = True
    lst(i).SignupHrs = h
    SetPartialHours = True
End Function

' ---------------- CANCEL ----------------
Public Function SubmitCancel(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long) As Boolean
    If Not RequireAdmin() Then Exit Function         ' v25.3: an admin's job (Graeme listed post and award, not cancel)
    If Not CanFinalize(p) Then Exit Function
    ' v24.3.1: Ask / Tell so SmokeTest can cancel a posting (step 9)
    If Ask("Cancel " & p.ID & " (" & p.Unit & ", " & Format$(p.ShiftDate, "ddd mmm d") & ")?" & vbNewLine & vbNewLine & _
           "The overtime is no longer needed. Nobody is awarded and NOBODY is charged.", _
           vbYesNo + vbQuestion, "Cancel request", vbYes) = vbNo Then Exit Function
    ComputeCancel lst, n
    On Error GoTo Failed
    UpdatePostingRecords p, lst, n, "Cancelled", ""
    Dim onCal As Boolean: onCal = (UpdateCoverageBySource(p.ID, "NCN") > 0)
    RefreshCalendar
    MarkMembersDirty
    If Not gSmokeRun Then ArchiveSignups p.ID      ' as WriteOutcome: SmokeTest never touches the inbox
    SaveIfWritable
    SubmitCancel = True
    On Error Resume Next                           ' (review M5) the receipt is outside the Failed frame: the cancel is done and saved
    ' v27 (item 7): the receipt (no OT Log rows: nobody is awarded, nobody is charged)
    ShowReceiptNow p.ID & " cancelled", "  (no OT Log rows: nobody awarded, nobody charged)" & vbNewLine, "  Nobody's Total Hours changed." & vbNewLine, _
                   IIf(onCal, "  Calendar: the line keeps its text with NCN (no coverage needed) as the coverage." & vbNewLine, "  Calendar: it was not on the calendar, so nothing changed there." & vbNewLine) & _
                   "  Postings sheet: " & p.ID & " is Cancelled." & vbNewLine, "  Reopen Posting (admin) opens it again." & vbNewLine
    Exit Function
Failed:
    ' v24.3.1: the same handler as WriteOutcome. A cancel that wrote Status "Cancelled" and then
    ' failed must not be saved as if it finished (P7).
    Dim em As String: em = Err.Description
    LogError "SubmitCancel", Err.Number, em
    On Error Resume Next
    MarkNoSave
    MsgBox "Cancelling " & p.ID & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving, reopen, and cancel " & p.ID & " again.", vbCritical, "Call Out Roster"
End Function

' OT Log rows (awarded, signed up, charged) + posting / list records.
Private Function WriteOutcome(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                              ByVal newStatus As String, ByVal awardNames As String, _
                              ByVal ps As Date, ByVal maxRank As Long, _
                              Optional ByVal outName As String = "", Optional ByVal outCode As String = "", _
                              Optional ByVal outHrs As Double = 0, Optional ByVal outTake As Double = 0) As Boolean
    Dim i As Long, k As Long, nRows As Long
    For i = 1 To n
        If lst(i).Awarded Or lst(i).SignedUp Or lst(i).Charged > 0 Then nRows = nRows + 1
    Next i
    If outName <> "" Then nRows = nRows + 1

    On Error GoTo Failed
    Dim before As Object: Set before = ReceiptBefore()      ' v27 (receipts): Total Hours before the rows land
    gReceipt.Recorded = "": gReceipt.Changed = ""
    If nRows > 0 Then
        Dim data() As Variant: ReDim data(1 To nRows, 1 To L_NCOLS)
        Dim user As String, stamp As Date
        user = CurrentUser(): stamp = Now
        For i = 1 To n
            If lst(i).Awarded Or lst(i).SignedUp Or lst(i).Charged > 0 Then
                k = k + 1
                data(k, L_REF) = p.ID
                data(k, L_TYPE) = "Posting"
                data(k, L_DATE) = p.ShiftDate
                data(k, L_DEPT) = p.Unit
                data(k, L_SHIFTLBL) = p.ShiftLbl
                data(k, L_CREW) = p.Crew
                data(k, L_TASK) = p.Reason
                data(k, L_BASEHRS) = p.Hours
                data(k, L_ORDER) = lst(i).Rank
                data(k, L_EMP) = lst(i).EmpName
                data(k, L_CODE) = lst(i).DeptCode
                data(k, L_EMPHRS) = lst(i).TotalHrs
                If lst(i).Awarded Then
                    data(k, L_OUTCOME) = "Awarded"
                    ' v24: provisional accepted OT (roster column C) sums this row's Hours, so a
                    ' partial award must carry the hours awarded, not the posting's
                    data(k, L_BASEHRS) = AwardedHours(p.Hours, lst(i).SignupHrs)
                    If IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
                        data(k, L_REASON) = "Partial award: " & HoursText(lst(i).SignupHrs) & " of " & HoursText(p.Hours) & " h"
                    End If
                    If lst(i).OverrideNote <> "" Then data(k, L_REASON) = Trim$(SafeStr(data(k, L_REASON)) & "  " & lst(i).OverrideNote)   ' v27 (item 21)
                ElseIf lst(i).Charged > 0 Then
                    data(k, L_OUTCOME) = "Refused"
                    ' v26.5.5 (the 16 h rule): the row's Hours are the hours they could have taken, so the Why sentence adds up
                    data(k, L_BASEHRS) = CappedPostingHours(p.Hours, lst(i).MaxHours, lst(i).MaxSet)
                    If newStatus = "No takers" Then
                        data(k, L_REASON) = "Posting - no takers"
                    ElseIf outName <> "" Then
                        data(k, L_REASON) = "Posting - didn't sign up (awarded out of department)"
                    Else
                        data(k, L_REASON) = "Posting - didn't sign up (#" & lst(i).Rank & ", awarded #" & maxRank & ")"
                    End If
                Else
                    data(k, L_OUTCOME) = "Signed Up"
                    If IsPartialSignup(p.Hours, lst(i).SignupHrs) Then
                        data(k, L_REASON) = "Signed up for " & HoursText(lst(i).SignupHrs) & " of " & HoursText(p.Hours) & " h"
                    End If
                End If
                If lst(i).ESign <> "" And lst(i).SignedUp Then          ' v24.1 (never published)
                    data(k, L_REASON) = Trim$(SafeStr(data(k, L_REASON)) & "  [electronic sign-up " & lst(i).ESign & "]")
                End If
                data(k, L_CHARGED) = lst(i).Charged
                data(k, L_PERIOD) = ps
                data(k, L_USER) = user
                data(k, L_STAMP) = stamp
            End If
        Next i
        If outName <> "" Then
            k = k + 1
            data(k, L_REF) = p.ID
            data(k, L_TYPE) = "Posting"
            data(k, L_DATE) = p.ShiftDate
            data(k, L_DEPT) = p.Unit
            data(k, L_SHIFTLBL) = p.ShiftLbl
            data(k, L_CREW) = p.Crew
            data(k, L_TASK) = p.Reason
            data(k, L_BASEHRS) = AwardedHours(p.Hours, outTake)      ' v24.1: the hours taken (provisional OT sums this)
            data(k, L_EMP) = outName
            data(k, L_CODE) = outCode
            data(k, L_EMPHRS) = outHrs
            data(k, L_OUTCOME) = "Awarded"
            data(k, L_REASON) = "Out of department" & IIf(IsPartialSignup(p.Hours, outTake), _
                                " - partial award: " & HoursText(outTake) & " of " & HoursText(p.Hours) & " h", "")
            data(k, L_CHARGED) = 0
            data(k, L_PERIOD) = ps
            data(k, L_USER) = user
            data(k, L_STAMP) = stamp
        End If
        If Not AppendLogRows(data) Then Exit Function
        gReceipt.Recorded = RowsText(data, k)                   ' v27: the receipt's first part (the finisher adds the rest and shows it)
    End If
    gReceipt.Changed = ReceiptChanges(before)

    UpdatePostingRecords p, lst, n, newStatus, awardNames
    MarkMembersDirty                    ' v24.1: the members' file is behind now
    If Not gSmokeRun Then ArchiveSignups p.ID      ' v24.1: its inbox files are done with (SmokeTest never touches the inbox)
    SaveIfWritable
    WriteOutcome = True
    Exit Function

Failed:
    Dim em As String: em = Err.Description
    LogError "WriteOutcome", Err.Number, em
    On Error Resume Next
    MarkNoSave
    MsgBox "Closing " & p.ID & " stopped part-way: " & em & vbNewLine & vbNewLine & _
           "Saving is blocked. Close WITHOUT saving, reopen, and do " & p.ID & " again.", vbCritical, "Call Out Roster"
End Function

' v24.3.2: the Reason of one posting, changed from the Calendar ("Change the text of this
' line" on the line the posting wrote). Words only: the Reason drives no rule once the posting
' exists, and the list, hours, dates and status stay as they are. True when the row was found.
Public Function SetPostingReason(ByVal id As String, ByVal reason As String) As Boolean
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim i As Long, en As Long, em As String
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(Trim$(id)) Then
            On Error GoTo Failed
            UnlockSheet ws
            lo.DataBodyRange.Cells(i, PS_REASON).Value = "'" & reason      ' text, whatever it starts with (= + - or a date)
            ProtectSheet ws
            On Error GoTo 0
            MarkMembersDirty
            MarkCalDirty
            SetPostingReason = True
            Exit Function
        End If
    Next i
    Exit Function
Failed:
    en = Err.Number: em = Err.Description
    On Error Resume Next
    ProtectSheet ws
    On Error GoTo 0                                     ' (re-review) or the raise below would be swallowed by Resume Next
    Err.Raise en, , em
End Function

' v24.3.2: one posting's Reason as it is on the Postings sheet ("" when the row isn't there)
Public Function PostingReasonOf(ByVal id As String) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant: v = RangeTo2D(lo.DataBodyRange)
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(Trim$(id)) Then PostingReasonOf = Trim$(SafeStr(v(i, PS_REASON))): Exit Function
    Next i
End Function

' ---------------- calendar links ----------------
' New posting: pick the calendar line it covers -> "POSTED P-####"
' Returns a line for the "created" message: where it went, or why it didn't.
Private Function CalendarPosted(ByRef p As PostRec) As String
    On Error GoTo Failed
    Dim k As String
    If gCalPostKey <> "" And Int(p.ShiftDate) = Int(gCalPostDate) And StrComp(p.ShiftLbl, gCalPostShift, vbTextCompare) = 0 Then
        k = gCalPostKey                          ' posted from the Calendar: that line (or "NEW")
    End If
    gCalPostKey = ""
    ' support postings (PHA / unit operator ...) are usually a new line;
    ' coverage postings usually fill someone's time off
    CalendarPosted = CalendarPlace(p.ShiftDate, p.ShiftLbl, p.Unit, "Posting " & p.ID & " - which calendar line does it cover?", _
                                   "POSTED " & p.ID, p.Hours, p.ID, ReasonType(p.Reason), p.Reason, _
                                   IsSupportReason(p.Reason), k)
    If CalendarPosted = "" Then
        CalendarPosted = "not put on the calendar. To add it later: Needs Attention > the posting > Put on the calendar."
    End If
    RefreshCalendar
    Exit Function
Failed:
    CalendarPosted = "NOT updated - " & Err.Description & ". Use Needs Attention > the posting > Put on the calendar."
End Function

' After an award / no takers: update the posting's calendar line, or put it
' on the calendar now if it never got there. Returns a line for the message.
' perNameHours (v24): byText already carries each partial award's hours ("Joe D 8h /
' Sam K 4h"), so the line's own Hours cell is cleared instead of set.
Private Function CalendarOutcome(ByRef p As PostRec, ByVal byText As String, _
                                 Optional ByVal perNameHours As Boolean = False) As String
    On Error GoTo Failed
    Dim lineHrs As Variant
    If perNameHours Then lineHrs = Empty Else lineHrs = p.Hours
    If UpdateCoverageBySource(p.ID, byText, lineHrs, perNameHours) > 0 Then
        CalendarOutcome = "updated to """ & byText & """"
    Else
        CalendarOutcome = CalendarPlace(p.ShiftDate, p.ShiftLbl, p.Unit, _
                                        p.ID & " wasn't on the calendar - which line does it cover?", _
                                        byText, lineHrs, p.ID, ReasonType(p.Reason), p.Reason, IsSupportReason(p.Reason), "", _
                                        Not IsSupportReason(p.Reason))   ' v27 (item 10): a coverage posting takes the unit's one open line, no question; a support posting (its own line) still asks
        If CalendarOutcome = "" Then
            CalendarOutcome = "NOT ON THE CALENDAR (left off) - Needs Attention > the posting > Put on the calendar"   ' v27 (item 10): the receipt shows it in red
        Else
            ' a line that already had its own Hours would draw them after the per-name hours
            If perNameHours Then UpdateCoverageBySource p.ID, byText, lineHrs, True
            CalendarOutcome = "added: " & CalendarOutcome
        End If
    End If
    SaveIfWritable
    RefreshCalendar
    Exit Function
Failed:
    CalendarOutcome = "NOT updated - " & Err.Description
End Function

' Posting reason -> calendar type
Public Function ReasonType(ByVal reason As String) As String
    Dim r As String: r = LCase$(reason)
    If InStr(r, "pha") > 0 Then
        ReasonType = "PHA"
    ElseIf InStr(r, "vac") > 0 Then
        ReasonType = "VAC"
    ElseIf InStr(r, "leave") > 0 Then
        ReasonType = "LOA"
    ElseIf InStr(r, "train") > 0 Then
        ReasonType = "TRN"
    Else
        ReasonType = "SUP"                 ' unit operator support / other
    End If
End Function

' True for "additional support" postings (nobody off - a new calendar line)
Public Function IsSupportReason(ByVal reason As String) As Boolean
    Select Case ReasonType(reason)
        Case "PHA", "SUP": IsSupportReason = True
    End Select
End Function


' v24: someone who moved out of a unit, or left, is still on the frozen list
' of that unit's open postings. For shifts on or after fromDate their list
' row is marked "Exempt - <note>", and ComputeAward / ComputeNoTakers /
' ComputeOutside never charge an exempt row.
'   units    "" = every unit; else "|Unit A|Unit B|"
'   listOnly True = change nothing, just say which postings
' Returns the posting IDs, e.g. "P-0012, P-0014" ("" = none).
Public Function ExemptFromOpenPostings(ByVal emp As String, ByVal fromDate As Date, ByVal units As String, _
                                       ByVal noteTxt As String, Optional ByVal listOnly As Boolean = False) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Function
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    Dim openIds As Object: Set openIds = CreateObject("Scripting.Dictionary")
    openIds.CompareMode = 1
    Dim i As Long, id As String, hit As Boolean, outS As String
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And SafeStr(pv(i, PS_STATUS)) = "Open" Then
            If Int(ToDate(pv(i, PS_DATE))) >= Int(fromDate) Then
                If units = "" Or InStr(1, units, "|" & SafeStr(pv(i, PS_UNIT)) & "|", vbTextCompare) > 0 Then openIds(id) = True
            End If
        End If
    Next i
    If openIds.Count = 0 Then Exit Function
    For i = 1 To UBound(lv, 1)
        id = Trim$(SafeStr(lv(i, PL_ID)))
        If openIds.Exists(id) Then
            If StrComp(Trim$(SafeStr(lv(i, PL_EMP))), Trim$(emp), vbTextCompare) = 0 And _
               InStr(1, SafeStr(lv(i, PL_RESULT)), EXEMPT_TAG, vbTextCompare) = 0 Then
                outS = outS & IIf(outS = "", "", ", ") & id
                If Not listOnly Then
                    If Not hit Then UnlockSheet lr.Parent: hit = True
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = EXEMPT_TAG & noteTxt
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ExemptFromOpenPostings = outS
End Function

' Bring back (v24): the person is on the roster again, so the "Exempt - left ..."
' marks on OPEN postings come off. Returns the posting IDs ("" = none).
Public Function ClearLeftExemptions(ByVal emp As String) As String
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    If lo Is Nothing Or lr Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Or lr.DataBodyRange Is Nothing Then Exit Function
    Dim pv As Variant: pv = RangeTo2D(lo.DataBodyRange)
    Dim lv As Variant: lv = RangeTo2D(lr.DataBodyRange)
    Dim openIds As Object: Set openIds = CreateObject("Scripting.Dictionary")
    openIds.CompareMode = 1
    Dim i As Long, id As String, hit As Boolean, outS As String, res As String, note As String
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        If id <> "" And SafeStr(pv(i, PS_STATUS)) = "Open" Then openIds(id) = True
    Next i
    If openIds.Count = 0 Then Exit Function
    For i = 1 To UBound(lv, 1)
        id = Trim$(SafeStr(lv(i, PL_ID)))
        If openIds.Exists(id) Then
            If StrComp(Trim$(SafeStr(lv(i, PL_EMP))), Trim$(emp), vbTextCompare) = 0 Then
                res = SafeStr(lv(i, PL_RESULT))
                SplitExempt res, note
                If LCase$(Left$(note, Len(EXEMPT_TAG) + 4)) = LCase$(EXEMPT_TAG & "left") Then
                    If Not hit Then UnlockSheet lr.Parent: hit = True
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = res
                    outS = outS & IIf(outS = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ClearLeftExemptions = outS
End Function

' "Exempt - moves Feb 1 (PM-0007)" -> "PM-0007" ("" = the note carries no planned-move ID)
Public Function PlanIdIn(ByVal note As String) As String
    Dim a As Long, b As Long
    a = InStr(1, note, "(PM-", vbTextCompare)
    If a = 0 Then Exit Function
    b = InStr(a, note, ")")
    If b > a Then PlanIdIn = Mid$(note, a + 1, b - a - 1)
End Function

' ====================================================================
'  TABLE APPEND (Postings sheet)
' ====================================================================
Public Function AppendRows(ByVal lo As ListObject, data As Variant, ByVal nCols As Long) As Boolean
    If lo Is Nothing Then Exit Function
    Dim n As Long: n = UBound(data, 1)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim em As String
    On Error GoTo Fail
    UnlockSheet ws
    lo.Resize lo.HeaderRowRange.Resize(cnt + 1 + n)
    lo.HeaderRowRange.Offset(cnt + 1, 0).Resize(n, nCols).Value = data
    ProtectSheet ws
    AppendRows = True
    Exit Function
Fail:
    em = Err.Description
    LogError "AppendRows", Err.Number, em
    On Error Resume Next
    ProtectSheet ws
    MsgBox "Could not write to " & lo.Name & ": " & em, vbCritical, "Call Out Roster"
End Function
