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
' ====================================================================

' ---------------- configuration ----------------
Public Const POST_VERSION As String = "Postings v24.3 (Jan 5)"
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
           "Run ApplyRosterLayout next to refresh the buttons.", vbInformation
End Sub

Private Sub CreatePostingTables()
    Dim ws As Worksheet
    If SheetExists(POSTINGS_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(POSTINGS_SHEET)
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(LOG_SHEET))
        ws.Name = POSTINGS_SHEET
        ws.Tab.Color = RGB(0, 128, 70)
    End If
    ws.Unprotect Password:=SHEET_PASSWORD

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
    ws.Unprotect Password:=SHEET_PASSWORD
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
        MsgBox "Postings aren't set up yet. Run SetupPostings first (Alt+F8).", vbExclamation
    ElseIf Not PostingListCurrent() Then
        MsgBox "The posting lists need their Signup Hours column (v24). Run Alt+F8 > SetupPeople once, then try again.", vbExclamation
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
        why = "column " & ColLetter(nextCol.Column) & " on the Postings sheet, right beside the posted lists, isn't empty. Clear it and run SetupPeople again"
        Exit Function
    End If
    Dim nRows As Long: nRows = lr.ListRows.Count
    ws.Unprotect Password:=SHEET_PASSWORD
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
    If Not SheetExists(POSTING_SHEET) Then MsgBox "Sheet '" & POSTING_SHEET & "' not found.": Exit Sub
    ThisWorkbook.Worksheets(POSTING_SHEET).Visible = xlSheetVisible
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
    If Not RequireAdmin() Then Exit Sub
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
Public Function CreatePostingFrom(ByRef p As PostRec, Optional ByVal doPrint As Boolean = True, _
                                  Optional ByVal doPdf As Boolean = True) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If Not PostingTablesReady() Then Exit Function

    If p.Positions < 1 Then p.Positions = 1
    p.PostedOn = Date
    p.Status = "Open"

    If p.ClosesOn > p.ShiftDate Then
        If Ask("The posting closes after the shift date. Continue anyway?", vbYesNo + vbExclamation, "Microsoft Excel", vbYes) = vbNo Then Exit Function
    End If

    ' eligible list = same rules and order as before (call order, posting grade cutoffs)
    Application.Calculate
    Dim recs() As EmpRec, n As Long
    ' everyone with the department code (any grade) who is free that day:
    ' not scheduled, not under the rest rule, not booked off, not on a shift switch
    n = LoadEligible(p.Unit, CrewFilter(p.Crew), recs, DeptPostingMaxGrade(p.Unit), True, p.ShiftDate)
    Dim offList As String
    If n > 0 Then n = FilterNotFree(recs, n, p.ShiftDate, p.ShiftLbl, offList)
    If n = 0 Then MsgBox "No one is free to sign up for " & p.Unit & " / " & p.Crew & "." & _
                         IIf(offList <> "", vbNewLine & "Left off: " & offList, ""), vbInformation: Exit Function

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
    If Ask("Create posting " & p.ID & "?" & vbNewLine & vbNewLine & _
              p.Unit & "  |  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & "  |  " & p.ShiftLbl & vbNewLine & _
              "Crew: " & p.Crew & "   Hours: " & p.Hours & "   Positions: " & p.Positions & vbNewLine & _
              "Reason: " & p.Reason & vbNewLine & _
              "Closes: " & Format$(p.ClosesOn, "ddd mmm d, yyyy") & " @ " & POSTING_CLOSE_TIME & vbNewLine & vbNewLine & _
              n & " eligible employees will be listed. " & outTxt & _
              IIf(offList <> "", vbNewLine & vbNewLine & "Not listed: " & offList, ""), _
              vbOKCancel + vbQuestion, "Create Posting", vbOK) <> vbOK Then Exit Function

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
    pr(1, PS_REASON) = p.Reason
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
    If (doPrint Or doPdf) And wasVis <> xlSheetVisible Then ws.Visible = xlSheetVisible

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

    If ws.Visible <> wasVis And Not doView Then ws.Visible = wasVis
    If doView Then
        ws.Visible = xlSheetVisible
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

    Dim fName As String, ch As Variant
    fName = p.ID & " " & p.Unit & " " & Format$(p.ShiftDate, "yyyy-mm-dd") & " " & p.Crew
    For Each ch In Array("/", "\", ":", "*", "?", """", "<", ">", "|")
        fName = Replace(fName, CStr(ch), "-")
    Next ch
    pdfPath = folder & fName & ".pdf"

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
'  THE OVERTIME POSTING SHEET  (output only - redrawn every time)
' ====================================================================
Private Function PostingSheet() As Worksheet
    If Not SheetExists(POSTING_SHEET) Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=RosterSheet())
        ws.Name = POSTING_SHEET
    End If
    Set PostingSheet = ThisWorkbook.Worksheets(POSTING_SHEET)
End Function

' Blank page (used by SetupPostings) - removes the old inputs and buttons
Public Sub ResetPostingSheet()
    Dim p As PostRec, lst() As PListRec
    ReDim lst(1 To 1)
    Application.ScreenUpdating = False
    RenderPostingSheet PostingSheet(), p, lst, 0
    Application.ScreenUpdating = True
End Sub

Private Sub RenderPostingSheet(ByVal ws As Worksheet, ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long)
    Const NAVY_C As Long = 6567966        ' RGB(30, 56, 100)
    Dim lightFill As Long: lightFill = RGB(242, 245, 249)
    Dim greyTxt As Long: greyTxt = RGB(90, 90, 90)
    Dim lineCol As Long: lineCol = RGB(200, 206, 215)
    Dim blank As Boolean: blank = (p.ID = "")
    Dim i As Long, r As Long

    ws.Unprotect Password:=SHEET_PASSWORD
    Application.DisplayAlerts = False       ' no "merging cells only keeps..." prompts

    ' ---- wipe everything: old inputs, dropdowns, buttons, formats ----
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Cells.UnMerge
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    ws.Cells.VerticalAlignment = xlCenter
    ws.Rows.RowHeight = 15
    ws.Columns.ColumnWidth = 8.43
    ws.Columns("A").ColumnWidth = 5       ' #
    ws.Columns("B").ColumnWidth = 9       ' hours
    ws.Columns("C").ColumnWidth = 34      ' name
    ws.Columns("D").ColumnWidth = 12      ' shift
    ws.Columns("E").ColumnWidth = 30      ' initials
    ws.Range("A1:E200").NumberFormat = "@"

    ' ---- title + posting ID ----
    ws.Rows(1).RowHeight = 36
    With ws.Range("A1:D1")
        .Merge
        .Value = "OVERTIME POSTING"
        .Font.Size = 22: .Font.Bold = True: .Font.Color = NAVY_C
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range("E1")
        .Value = IIf(blank, "", p.ID)
        .Font.Size = 18: .Font.Bold = True: .Font.Color = NAVY_C
        .HorizontalAlignment = xlCenter
        .Interior.Color = lightFill
        .BorderAround xlContinuous, xlMedium, , NAVY_C
    End With

    ws.Rows(2).RowHeight = 30
    With ws.Range("A2:E2")
        .Merge
        .Value = "The shift below requires coverage. If you're interested, initial beside your name. " & _
                 "The shift will be awarded according to the awarding process defined in the CBA."
        .WrapText = True
        .Font.Size = 9.5: .Font.Italic = True: .Font.Color = greyTxt
        .HorizontalAlignment = xlLeft
    End With

    ' row 3: rule (and a status banner when reprinting a closed posting)
    With ws.Range("A3:E3")
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = NAVY_C
    End With
    If Not blank And p.Status <> "Open" And p.Status <> "" Then
        ws.Rows(3).RowHeight = 18
        With ws.Range("A3:E3")
            .Merge
            .Value = "CLOSED - " & UCase$(p.Status) & "  (reprint for reference only)"
            .Font.Bold = True: .Font.Color = RGB(190, 30, 30)
            .HorizontalAlignment = xlCenter
        End With
    Else
        ws.Rows(3).RowHeight = 6
    End If

    ' ---- details (rows 5-9) ----
    ws.Rows(4).RowHeight = 6
    For r = 5 To 9
        ws.Rows(r).RowHeight = 21
    Next r
    DetailRow ws, 5, "Unit", IIf(blank, "", p.Unit), "Hours", IIf(blank, "", p.Hours)
    DetailRow ws, 6, "Shift date", IIf(blank, "", p.ShiftDate), "Positions", IIf(blank, "", p.Positions)
    DetailRow ws, 7, "Shift", IIf(blank, "", p.ShiftLbl), "Posted", IIf(blank, "", p.PostedOn)
    DetailRow ws, 8, "Crew", IIf(blank, "", p.Crew), "Closes", IIf(blank, "", p.ClosesOn)
    ws.Range("C9:E9").Merge
    DetailRow ws, 9, "Reason", IIf(blank, "", p.Reason), "", ""
    ws.Range("C6").NumberFormat = "dddd, mmmm d, yyyy"
    ws.Range("E7").NumberFormat = "mmm d, yyyy"
    ws.Range("E8").NumberFormat = "ddd mmm d, yyyy"" @ " & POSTING_CLOSE_TIME & """"
    ws.Range("E5:E6").NumberFormat = "General"
    ws.Range("E5:E9").HorizontalAlignment = xlLeft
    With ws.Range("A5:E9")
        .Interior.Color = lightFill
        .BorderAround xlContinuous, xlThin, , lineCol
    End With

    ' ---- list ----
    ws.Rows(10).RowHeight = 10
    Const HDR As Long = 11
    ws.Rows(HDR).RowHeight = 21
    ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, 5)).Value = Array("#", "Hours", "Name", "Shift", "Initials to sign up")
    With ws.Range(ws.Cells(HDR, 1), ws.Cells(HDR, 5))
        .Interior.Color = NAVY_C
        .Font.Color = vbWhite: .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Cells(HDR, 3).HorizontalAlignment = xlLeft

    Dim last As Long
    If n = 0 Then
        last = HDR + 1
        ws.Rows(last).RowHeight = 26
        With ws.Range(ws.Cells(last, 1), ws.Cells(last, 5))
            .Merge
            .Value = "No posting loaded - use New Posting or Load Posting on the OT Roster."
            .Font.Italic = True: .Font.Color = greyTxt
            .HorizontalAlignment = xlCenter
        End With
    Else
        last = HDR + n
        Dim v() As Variant: ReDim v(1 To n, 1 To 4)
        For i = 1 To n
            v(i, 1) = lst(i).Rank
            v(i, 2) = lst(i).TotalHrs
            v(i, 3) = lst(i).EmpName
            v(i, 4) = lst(i).EmpShift
        Next i
        With ws.Range(ws.Cells(HDR + 1, 1), ws.Cells(last, 2))
            .NumberFormat = "#,##0"
            .HorizontalAlignment = xlCenter
        End With
        ws.Range(ws.Cells(HDR + 1, 1), ws.Cells(last, 4)).Value = v
        ws.Range(ws.Cells(HDR + 1, 1), ws.Cells(last, 1)).Font.Color = greyTxt
        ws.Range(ws.Cells(HDR + 1, 3), ws.Cells(last, 3)).Font.Size = 12
        ws.Range(ws.Cells(HDR + 1, 4), ws.Cells(last, 4)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(HDR + 1, 1), ws.Cells(last, 1)).EntireRow.RowHeight = 26
        With ws.Range(ws.Cells(HDR + 1, 1), ws.Cells(last, 4)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = lineCol
        End With
        With ws.Range(ws.Cells(last, 1), ws.Cells(last, 4)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = lineCol
        End With
        ' signature line for each person
        With ws.Range(ws.Cells(HDR + 1, 5), ws.Cells(last, 5))
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
        End With
    End If

    ' ---- out-of-department sign-up (blank lines) ----
    If n > 0 And POSTING_OUTSIDE_LINES > 0 Then
        r = last + 2
        ws.Rows(last + 1).RowHeight = 14
        ws.Rows(r).RowHeight = 21
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("", "Out of department - print your name", "", "", "Initials to sign up")
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
            .Interior.Color = lightFill
            .Font.Bold = True: .Font.Color = NAVY_C
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Color = NAVY_C
        End With
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 4)).Merge
        ws.Cells(r, 2).HorizontalAlignment = xlLeft
        For i = 1 To POSTING_OUTSIDE_LINES
            ws.Rows(r + i).RowHeight = 26
            ws.Range(ws.Cells(r + i, 2), ws.Cells(r + i, 4)).Merge          ' one wide name box
            With ws.Range(ws.Cells(r + i, 2), ws.Cells(r + i, 5)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(150, 150, 150)
            End With
            With ws.Cells(r + i, 5).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = lineCol
            End With
        Next i
        With ws.Cells(r + POSTING_OUTSIDE_LINES + 1, 3)
            .Value = "Out-of-department sign-ups are awarded only if no one in the department signs up."
            .Font.Size = 8.5: .Font.Italic = True: .Font.Color = greyTxt
        End With
        last = r + POSTING_OUTSIDE_LINES + 1
    End If

    ' ---- footer ----
    If Not blank Then
        r = last + 2
        ws.Rows(r).RowHeight = 30
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
            .Merge
            .Value = "Sign-up closes " & Format$(p.ClosesOn, "dddd, mmmm d") & " @ " & POSTING_CLOSE_TIME & _
                     ".  Listed in call order.  Hours are each person's total overtime hours as of " & _
                     Format$(p.PostedOn, "mmmm d, yyyy") & ", including accepted overtime not yet in Kronos."
            .WrapText = True
            .Font.Size = 9: .Font.Color = greyTxt
            .HorizontalAlignment = xlLeft
        End With
        With ws.Cells(r + 1, 5)
            .Value = p.ID & "  |  printed " & Format$(Now, "mmm d, yyyy h:mm AM/PM")
            .Font.Size = 8: .Font.Color = greyTxt
            .HorizontalAlignment = xlRight
        End With
        last = r + 1
    End If

    ' ---- page ----
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(last, 5)).Address
        .PrintTitleRows = ""
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .CenterVertically = False
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintGridlines = False
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .CenterFooter = "": .RightFooter = ""
    End With
    Application.PrintCommunication = True
    ws.DisplayPageBreaks = False
    On Error GoTo 0

    Application.DisplayAlerts = True
    ProtectSheet ws
End Sub

Private Sub DetailRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl1 As String, ByVal val1 As Variant, _
                      ByVal lbl2 As String, ByVal val2 As Variant)
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 2))
        .Merge
        .Value = lbl1
        .HorizontalAlignment = xlRight
        .Font.Bold = True: .Font.Size = 10: .Font.Color = RGB(90, 90, 90)
        .IndentLevel = 1
    End With
    With ws.Cells(r, 3)
        If VarType(val1) <> vbString Then .NumberFormat = "General"
        .Value = val1
        .Font.Bold = True: .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    If lbl2 = "" Then Exit Sub              ' row 9: reason spans C:E
    With ws.Cells(r, 4)
        .Value = lbl2
        .HorizontalAlignment = xlRight
        .Font.Bold = True: .Font.Size = 10: .Font.Color = RGB(90, 90, 90)
    End With
    With ws.Cells(r, 5)
        If VarType(val2) <> vbString Then .NumberFormat = "General"
        .Value = val2
        .Font.Bold = True: .Font.Size = 12
        .IndentLevel = 1
    End With
End Sub


' ====================================================================
'  READ
' ====================================================================
Private Function GetTable(ByVal nm As String) As ListObject
    If Not SheetExists(POSTINGS_SHEET) Then Exit Function
    On Error Resume Next
    Set GetTable = ThisWorkbook.Worksheets(POSTINGS_SHEET).ListObjects(nm)
    On Error GoTo 0
End Function

Private Function ToDate(ByVal x As Variant) As Date
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
Private Sub SplitExempt(ByRef result As String, ByRef exemptNote As String)
    Dim p As Long: p = InStr(1, result, EXEMPT_TAG, vbTextCompare)
    exemptNote = ""
    If p = 0 Then Exit Sub
    exemptNote = Trim$(Mid$(result, p))
    result = Trim$(Left$(result, p - 1))
    If Right$(result, 1) = "|" Then result = Trim$(Left$(result, Len(result) - 1))
End Sub

' What goes in the Result cell for one list person
Private Function ResultCellText(ByRef rec As PListRec) As String
    If rec.ExemptNote = "" Or rec.Result = rec.ExemptNote Then
        ResultCellText = rec.Result
    ElseIf rec.Result = "" Then
        ResultCellText = rec.ExemptNote
    Else
        ResultCellText = rec.Result & " | " & rec.ExemptNote
    End If
End Function

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
                   IIf(Date > closes, "READY TO AWARD", "open until " & Format$(closes, "mmm d"))
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
            If Date > ToDate(v(r, PS_CLOSES)) Then st = "Open - ready to award" Else st = "Open until " & Format$(ToDate(v(r, PS_CLOSES)), "mmm d")
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

' Next P- number across both the Postings table and the OT Log
Private Function NextPostingID() As String
    Dim s As String: s = NextRefID("P")
    Dim mx As Long: mx = CLng(Val(Mid$(s, 3))) - 1
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim v As Variant, i As Long, k As Long
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            v = RangeTo2D(lo.ListColumns(PS_ID).DataBodyRange)
            For i = 1 To UBound(v, 1)
                s = UCase$(Trim$(SafeStr(v(i, 1))))
                If Left$(s, 2) = "P-" Then
                    k = CLng(Val(Mid$(s, 3)))
                    If k > mx Then mx = k
                End If
            Next i
        End If
    End If
    NextPostingID = "P-" & Format$(mx + 1, "0000")
End Function


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
            lst(i).Charged = p.Hours * CHARGE_MULTIPLIER
            lst(i).Result = "Charged"
            nCharged = nCharged + 1
            totCharged = totCharged + lst(i).Charged
        Else
            lst(i).Result = ""
        End If
    Next i
End Sub

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
            nFullWaiting = nFullWaiting + 1
            If who = "" Then who = "#" & lst(i).Rank & " " & lst(i).EmpName
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
            lst(i).Charged = p.Hours * CHARGE_MULTIPLIER
            lst(i).Result = "Charged"
            nCharged = nCharged + 1
            totCharged = totCharged + lst(i).Charged
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
            lst(i).Charged = p.Hours * CHARGE_MULTIPLIER
            lst(i).Result = "Charged"
            nCharged = nCharged + 1
            totCharged = totCharged + lst(i).Charged
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
Private Function CanFinalize(ByRef p As PostRec) As Boolean
    If Not RequireAdmin() Then Exit Function
    If Not RequireWritable() Then Exit Function
    If PostingStatus(p.ID) <> "Open" Then
        MsgBox p.ID & " is no longer open (" & PostingStatus(p.ID) & ").", vbExclamation
        Exit Function
    End If
    If HasActiveLogRows(p.ID) Then
        MsgBox p.ID & " already has entries in the OT Log, so it can't be closed again." & vbNewLine & _
               "If they're wrong: Void Entry > " & p.ID & ", then ReopenPosting.", vbExclamation
        Exit Function
    End If
    CanFinalize = True
End Function

Private Function EarlyOK(ByRef p As PostRec, ByVal what As String) As Boolean
    EarlyOK = True
    If Date <= p.ClosesOn Then
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
        MsgBox "Nobody is awarded." & vbNewLine & vbNewLine & _
               "If nobody took it, use NO TAKERS (charges everyone eligible)." & vbNewLine & _
               "If the overtime isn't needed any more, use CANCEL REQUEST (no charges).", vbExclamation, "Award"
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
          HoursText(p.Hours * CHARGE_MULTIPLIER) & " hrs = " & HoursText(tot) & " hrs" & _
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
    Tell p.ID & " awarded to " & awardShow & "." & vbNewLine & "Calendar: " & calTxt, vbInformation, "Award"
    SubmitAward = True
End Function

' ---------------- AWARD OUT OF DEPARTMENT ----------------
' Nobody in the department took it; someone from another department signed
' a blank line. In-department operators who didn't sign up are charged as
' for NO TAKERS; anyone who signed up isn't charged.
Public Function SubmitAwardOutside(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                                   ByVal outName As String, Optional ByVal outTake As Double = 0) As Boolean
    ' outTake (v24.1): hours the out-of-department person takes when it is only part of the shift (0 = all)
    If Not CanFinalize(p) Then Exit Function
    outName = Trim$(outName)
    Dim outPartial As Boolean: outPartial = IsPartialSignup(p.Hours, outTake)
    If outName = "" Then MsgBox "Pick or type the out-of-department name first.", vbExclamation: Exit Function
    Dim i As Long, marked As String, signedIn As String
    For i = 1 To n
        If StrComp(Trim$(lst(i).EmpName), outName, vbTextCompare) = 0 Then
            MsgBox outName & " is on this posting's list - select them and use AWARD instead.", vbExclamation
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
        If MsgBox("These in-department people signed up and would normally get it first:" & vbNewLine & vbNewLine & _
                  signedIn & vbNewLine & "Award it out of department anyway? (They won't be charged.)", _
                  vbYesNo + vbExclamation, "Out of department") = vbNo Then Exit Function
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
        If MsgBox(outName & " isn't signed off in " & p.Unit & " (no code in that column on the OT Roster)." & vbNewLine & _
                  "Award anyway?", vbYesNo + vbExclamation, "Out of department") = vbNo Then Exit Function
    End If
    If onRoster Then why = NotFreeReason(outName, crew, p.ShiftDate, p.ShiftLbl)
    If why <> "" Then
        If MsgBox(outName & " isn't free that shift (" & why & ")." & vbNewLine & _
                  "Award anyway?", vbYesNo + vbExclamation, "Out of department") = vbNo Then Exit Function
    End If
    If Not onRoster Then
        If MsgBox(outName & " isn't on the OT Roster, so no email address or hours are on file." & vbNewLine & _
                  "Award anyway?", vbYesNo + vbExclamation, "Out of department") = vbNo Then Exit Function
    End If

    If MsgBox(p.ID & "  -  " & p.Unit & "  -  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & vbNewLine & vbNewLine & _
              "Award OUT OF DEPARTMENT to: " & outName & IIf(code <> "", "  (" & code & ")", "") & _
              IIf(outPartial, "  for " & HoursText(outTake) & " of " & HoursText(p.Hours) & " h (the rest is not covered; nobody is charged for it)", "") & vbNewLine & _
              "Charged (in department, didn't sign up): " & nCh & " x " & HoursText(p.Hours * CHARGE_MULTIPLIER) & _
              " hrs = " & HoursText(tot) & " hrs" & IIf(nCh > 0, vbNewLine & "Pay period starting " & Format$(ps, "mmm d, yyyy"), "") & _
              vbNewLine & vbNewLine & "Submit?", vbOKCancel + vbQuestion, "Confirm out-of-department award") <> vbOK Then Exit Function

    Dim awardTxt As String: awardTxt = outName & " (out of dept)"
    If Not WriteOutcome(p, lst, n, "Awarded", awardTxt, ps, 0, outName, IIf(code = "", "OD", code), hrs, outTake) Then Exit Function
    Dim calTxt As String
    If outPartial Then
        calTxt = CalendarOutcome(p, CalName(outName) & " " & CDbl(outTake) & "h", True)
    Else
        calTxt = CalendarOutcome(p, CalName(outName))
    End If
    EmailAward p, lst, n, awardTxt & IIf(outPartial, " (" & HoursText(outTake) & " h)", ""), outName
    MsgBox p.ID & " awarded out of department to " & outName & "." & vbNewLine & "Calendar: " & calTxt, vbInformation, "Award"
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
    ' lowest hours first (ties: seniority, as in call order)
    Dim a As Long, b As Long, t As Long
    For a = 2 To m
        t = keep(a)
        b = a - 1
        Do While b >= 1
            If recs(keep(b)).TotalHrs < recs(t).TotalHrs Then Exit Do
            If recs(keep(b)).TotalHrs = recs(t).TotalHrs And recs(keep(b)).Seniority <= recs(t).Seniority Then Exit Do
            keep(b + 1) = keep(b)
            b = b - 1
        Loop
        keep(b + 1) = t
    Next a
    Dim out() As String: ReDim out(0 To m - 1, 0 To 3)
    For i = 1 To m
        out(i - 1, 0) = recs(keep(i)).EmpName
        out(i - 1, 1) = recs(keep(i)).DeptCode
        out(i - 1, 2) = recs(keep(i)).EmpShift
        out(i - 1, 3) = Format$(recs(keep(i)).TotalHrs, "#,##0")
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
    If MsgBox(p.ID & "  -  " & p.Unit & "  -  " & Format$(p.ShiftDate, "ddd mmm d, yyyy") & vbNewLine & vbNewLine & _
              "NO TAKERS: charge all " & nCh & " eligible operator(s) " & HoursText(p.Hours * CHARGE_MULTIPLIER) & _
              " hrs each (" & HoursText(tot) & " hrs total)" & vbNewLine & _
              "Pay period starting " & Format$(ps, "mmm d, yyyy") & vbNewLine & vbNewLine & "Submit?", _
              vbOKCancel + vbExclamation, "Confirm no takers") <> vbOK Then Exit Function
    If Not WriteOutcome(p, lst, n, "No takers", "", ps, 0) Then Exit Function
    MsgBox p.ID & " closed - no takers." & vbNewLine & "Calendar: " & CalendarOutcome(p, "No takers"), vbInformation, "No takers"
    SubmitNoTakers = True
End Function

' ---------------- CANCEL ----------------
Public Function SubmitCancel(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long) As Boolean
    If Not CanFinalize(p) Then Exit Function
    If MsgBox("Cancel " & p.ID & " (" & p.Unit & ", " & Format$(p.ShiftDate, "ddd mmm d") & ")?" & vbNewLine & vbNewLine & _
              "The overtime is no longer needed. Nobody is awarded and NOBODY is charged.", _
              vbYesNo + vbQuestion, "Cancel request") = vbNo Then Exit Function
    ComputeCancel lst, n
    On Error GoTo Failed
    UpdatePostingRecords p, lst, n, "Cancelled", ""
    If UpdateCoverageBySource(p.ID, "NCN") = 0 Then
        MsgBox p.ID & " cancelled. It wasn't on the calendar, so there was nothing to change there.", vbInformation
    End If
    RefreshCalendar
    MarkMembersDirty
    ArchiveSignups p.ID
    SaveIfWritable
    SubmitCancel = True
    Exit Function
Failed:
    MsgBox "Couldn't cancel " & p.ID & ": " & Err.Description, vbCritical
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
                ElseIf lst(i).Charged > 0 Then
                    data(k, L_OUTCOME) = "Refused"
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
    End If

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
           "Saving is blocked. Close WITHOUT saving, reopen, and do " & p.ID & " again.", vbCritical
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
                                        byText, lineHrs, p.ID, ReasonType(p.Reason), p.Reason, IsSupportReason(p.Reason))
        If CalendarOutcome = "" Then
            CalendarOutcome = "not on the calendar (left off)"
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
                    If Not hit Then lr.Parent.Unprotect Password:=SHEET_PASSWORD: hit = True
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
                    If Not hit Then lr.Parent.Unprotect Password:=SHEET_PASSWORD: hit = True
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = res
                    outS = outS & IIf(outS = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ClearLeftExemptions = outS
End Function

' v24.3: a planned unit move was cancelled, so its "Exempt - moves ... (PM-0007)"
' marks come off OPEN postings. Only marks carrying that move's ID are touched.
' Returns the posting IDs ("" = none).
Public Function ClearMoveExemptions(ByVal emp As String, ByVal moveId As String) As String
    If Trim$(moveId) = "" Then Exit Function
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
                If InStr(1, note, "(" & Trim$(moveId) & ")", vbTextCompare) > 0 Then
                    If Not hit Then
                        lr.Parent.Unprotect Password:=SHEET_PASSWORD
                        hit = True
                    End If
                    lr.DataBodyRange.Cells(i, PL_RESULT).Value = res
                    outS = outS & IIf(outS = "", "", ", ") & id
                End If
            End If
        End If
    Next i
    If hit Then ProtectSheet lr.Parent
    ClearMoveExemptions = outS
End Function

Public Function HasActiveLogRows(ByVal id As String) As Boolean
    Dim lo As ListObject: Set lo = LogTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim v As Variant, i As Long
    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, L_REF)))) = UCase$(id) And UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" Then
            HasActiveLogRows = True
            Exit Function
        End If
    Next i
End Function

Private Sub UpdatePostingRecords(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, _
                                 ByVal newStatus As String, ByVal awardNames As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(POSTINGS_SHEET)
    Dim lo As ListObject: Set lo = GetTable(PT_TABLE)
    Dim lr As ListObject: Set lr = GetTable(PL_TABLE)
    Dim v As Variant, i As Long, j As Long
    ws.Unprotect Password:=SHEET_PASSWORD

    v = RangeTo2D(lo.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PS_ID)))) = UCase$(p.ID) Then
            lo.DataBodyRange.Cells(i, PS_STATUS).Value = newStatus
            If newStatus = "Open" Then
                lo.DataBodyRange.Cells(i, PS_AWARDTO).ClearContents
                lo.DataBodyRange.Cells(i, PS_AWARDON).ClearContents
                lo.DataBodyRange.Cells(i, PS_AWARDBY).ClearContents
            Else
                lo.DataBodyRange.Cells(i, PS_AWARDTO).Value = awardNames
                lo.DataBodyRange.Cells(i, PS_AWARDON).Value = Now
                lo.DataBodyRange.Cells(i, PS_AWARDBY).Value = CurrentUser()
            End If
            Exit For
        End If
    Next i

    v = RangeTo2D(lr.DataBodyRange)
    For i = 1 To UBound(v, 1)
        If UCase$(Trim$(SafeStr(v(i, PL_ID)))) = UCase$(p.ID) Then
            For j = 1 To n
                If CLng(NumOr0(v(i, PL_RANK))) = lst(j).Rank Then
                    If newStatus = "Open" Then
                        lr.DataBodyRange.Cells(i, PL_SIGNED).ClearContents
                        If lst(j).ExemptNote <> "" Then          ' a reopened posting keeps "moved / left" (v24)
                            lr.DataBodyRange.Cells(i, PL_RESULT).Value = lst(j).ExemptNote
                        Else
                            lr.DataBodyRange.Cells(i, PL_RESULT).ClearContents
                        End If
                        lr.DataBodyRange.Cells(i, PL_CHARGED).ClearContents
                        If lr.ListColumns.Count >= PL_SIGNHRS Then lr.DataBodyRange.Cells(i, PL_SIGNHRS).ClearContents
                    Else
                        lr.DataBodyRange.Cells(i, PL_SIGNED).Value = IIf(lst(j).SignedUp, "Yes", "")
                        lr.DataBodyRange.Cells(i, PL_RESULT).Value = ResultCellText(lst(j))
                        lr.DataBodyRange.Cells(i, PL_CHARGED).Value = lst(j).Charged
                        If lr.ListColumns.Count >= PL_SIGNHRS Then
                            If lst(j).SignedUp And IsPartialSignup(p.Hours, lst(j).SignupHrs) Then
                                lr.DataBodyRange.Cells(i, PL_SIGNHRS).Value = lst(j).SignupHrs
                            Else
                                lr.DataBodyRange.Cells(i, PL_SIGNHRS).ClearContents
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next j
        End If
    Next i
    ProtectSheet ws
End Sub

' Reopen an awarded posting (after its log rows have been voided)
Public Sub ReopenPosting()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim id As Variant
    id = Application.InputBox(Prompt:="Posting ID to reopen (e.g. P-0012):", Title:="Reopen Posting", Type:=2)
    If VarType(id) = vbBoolean Then Exit Sub
    id = UCase$(Trim$(CStr(id)))
    Dim p As PostRec, lst() As PListRec, n As Long
    n = LoadPosting(CStr(id), p, lst)
    If n = 0 Then MsgBox "Posting " & id & " not found.", vbExclamation: Exit Sub
    If p.Status = "Open" Then MsgBox id & " is already open.", vbInformation: Exit Sub

    ' its OT Log rows must be voided first so nothing is charged twice
    If HasActiveLogRows(CStr(id)) Then
        MsgBox id & " still has active OT Log rows. Void them first (Void Entry > " & id & "), then reopen.", vbExclamation
        Exit Sub
    End If
    If ArchiveHasActiveRef(CStr(id)) Then          ' v23: its rows left the live log at the year-end rollover
        MsgBox id & " belongs to a closed year, so it can't be reopened. Its OT Log rows are on the archive sheet and that year's file is frozen.", vbExclamation
        Exit Sub
    End If
    If MsgBox("Reopen " & id & " so it can be awarded again?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    ' v24: the list is frozen, the people aren't. Someone who has left by the shift
    ' date, or no longer holds a code in this unit, must not be charged on the re-award.
    ' (These two marks stay until the posting is closed. If the roster was wrong and is
    ' corrected while the posting is open, clear the person's Result cell on the Postings sheet.)
    Dim j As Long, leftOn As Date, dc As Long
    dc = DeptCol(p.Unit)
    For j = 1 To n
        If lst(j).ExemptNote = "" Then
            If RosterRowOf(lst(j).EmpName) = 0 Then
                leftOn = FormerLeftOn(lst(j).EmpName)
                If leftOn = 0 Then
                    lst(j).ExemptNote = EXEMPT_TAG & "not on the roster"
                ElseIf Int(leftOn) <= Int(p.ShiftDate) Then
                    lst(j).ExemptNote = EXEMPT_TAG & "left " & Format$(leftOn, "mmm d")
                End If
            ElseIf dc > 0 Then
                ' v24.3: the code they hold on the SHIFT date (a waiting unit move counts from its date)
                If UnitCodeOn(lst(j).EmpName, p.Unit, p.ShiftDate, _
                              Trim$(SafeStr(RosterSheet().Cells(RosterRowOf(lst(j).EmpName), dc).Value))) = "" Then
                    lst(j).ExemptNote = EXEMPT_TAG & "no code in " & p.Unit & " on the shift date"
                End If
            End If
        End If
    Next j
    UpdatePostingRecords p, lst, n, "Open", ""
    On Error Resume Next
    UpdateCoverageBySource p.ID, "POSTED " & p.ID
    RefreshCalendar
    On Error GoTo Oops
    RestoreSignups CStr(id)             ' v24.1: its electronic sign-ups count again
    MarkMembersDirty
    SaveIfWritable
    MsgBox id & " is open again.", vbInformation
    Exit Sub
Oops:
    ReportError "ReopenPosting"
End Sub


' ====================================================================
'  EMAIL (Outlook draft - you review and click Send)
' ====================================================================
' CC list for an award email, by unit. missingGroups returns the names of
' any group whose address isn't set yet (shown as a warning).
Private Function AwardCC(ByVal unit As String, ByRef missingGroups As String) As String
    Dim cc As String
    missingGroups = ""
    Select Case unit
        Case "Cypure", "Material Handler"              ' Material Handler = Fumigation (gases)
            AddCC cc, missingGroups, CC_GASES_SUPS, "Gases supervisors"
        Case "QC Derivatives", "QC Gases"              ' Quality Control Lab
            AddCC cc, missingGroups, CC_QC_SUP, "Quality Control supervisor"
        Case "Utilities"
            AddCC cc, missingGroups, CC_UTIL_SUP, "Utilities supervisor"
        Case Else                                      ' Phosphine, Derivatives, SBP
            AddCC cc, missingGroups, CC_SHIFT_SUPS, "Shift supervisors"
    End Select
    AddCC cc, missingGroups, CC_OPS_SUPT, "Operations superintendent"
    If Trim$(AWARD_CC) <> "" Then AddCC cc, missingGroups, AWARD_CC, ""
    AwardCC = cc
End Function

Private Sub AddCC(ByRef cc As String, ByRef missingGroups As String, ByVal addr As String, ByVal groupName As String)
    addr = Trim$(addr)
    If addr = "" Then
        If groupName <> "" Then missingGroups = missingGroups & "  " & groupName & vbNewLine
    ElseIf InStr(1, "; " & cc & ";", "; " & addr & ";", vbTextCompare) = 0 Then
        cc = cc & IIf(cc = "", "", "; ") & addr
    End If
End Sub

Public Function EmailFor(ByVal empName As String) As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), empName, vbTextCompare) = 0 Then
            EmailFor = Trim$(SafeStr(ws.Cells(r, C_EMAIL).Value))
            Exit Function
        End If
    Next r
End Function

Private Sub EmailAward(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByVal awardNames As String, _
                       Optional ByVal outName As String = "")
    Dim toList As String, missing As String, em As String, i As Long
    For i = 1 To n
        If lst(i).Awarded Then
            em = EmailFor(lst(i).EmpName)
            If em = "" Or InStr(em, "@") = 0 Then
                missing = missing & "  " & lst(i).EmpName & vbNewLine
            Else
                toList = toList & IIf(toList = "", "", "; ") & em
            End If
        End If
    Next i
    If outName <> "" Then
        em = EmailFor(outName)
        If em = "" Or InStr(em, "@") = 0 Then
            missing = missing & "  " & outName & vbNewLine
        Else
            toList = toList & IIf(toList = "", "", "; ") & em
        End If
    End If

    Dim subj As String
    subj = "Overtime awarded - " & p.ID & " - " & p.Unit & " - " & _
           Format$(p.ShiftDate, "ddd mmm d, yyyy") & " (" & p.Crew & ")"
    ' plain-text version for the fallback (v20)
    Dim plain As String
    plain = "The overtime posting below has been awarded." & vbCrLf & vbCrLf & _
            "Posting: " & p.ID & vbCrLf & "Unit: " & p.Unit & vbCrLf & _
            "Shift date: " & Format$(p.ShiftDate, "dddd, mmmm d, yyyy") & vbCrLf & _
            "Shift / crew: " & p.ShiftLbl & " / " & p.Crew & vbCrLf & _
            "Hours: " & CStr(p.Hours) & vbCrLf & "Reason: " & p.Reason & vbCrLf & _
            "Awarded to: " & awardNames & vbCrLf & vbCrLf & _
            "Awarded by " & CurrentUser() & " on " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & "."

    Dim ccMissing As String, ccList As String
    ccList = AwardCC(p.Unit, ccMissing)
    If missing <> "" Or ccMissing <> "" Then
        MsgBox IIf(missing <> "", "No email address on the roster (column K) for:" & vbNewLine & missing & vbNewLine, "") & _
               IIf(ccMissing <> "", "No CC address set in modPostings for:" & vbNewLine & ccMissing & vbNewLine, "") & _
               "Add them to the email before sending.", vbExclamation
    End If

    Dim ol As Object, mi As Object, shown As Boolean
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        ' new Outlook / no Outlook: open a plain email in the default mail app instead
        MailtoFallback toList, ccList, subj, plain, "Outlook couldn't be opened"
        Exit Sub
    End If

    Dim body As String
    body = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt'>" & _
           "<p>The overtime posting below has been awarded.</p>" & _
           "<table cellpadding='4' style='border-collapse:collapse'>" & _
           HtmlRow("Posting", p.ID) & _
           HtmlRow("Unit", p.Unit) & _
           HtmlRow("Shift date", Format$(p.ShiftDate, "dddd, mmmm d, yyyy")) & _
           HtmlRow("Shift / crew", p.ShiftLbl & " / " & p.Crew) & _
           HtmlRow("Hours", CStr(p.Hours)) & _
           HtmlRow("Reason", p.Reason) & _
           HtmlRow("Awarded to", awardNames) & _
           "</table><p>Awarded by " & CurrentUser() & " on " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & ".</p></div>"

    On Error GoTo MailFail
    Set mi = ol.CreateItem(0)
    mi.To = toList
    mi.CC = ccList
    mi.Subject = subj
    mi.Display                                   ' opens with your signature
    shown = True
    Dim h As String, b1 As Long, b2 As Long
    h = mi.HTMLBody
    b1 = InStr(1, h, "<body", vbTextCompare)
    If b1 > 0 Then b2 = InStr(b1, h, ">")
    If b2 > 0 Then
        mi.HTMLBody = Left$(h, b2) & body & Mid$(h, b2 + 1)
    Else
        mi.HTMLBody = body & h
    End If
    On Error GoTo 0
    Exit Sub
MailFail:
    Dim mailErr As String: mailErr = Err.Description
    On Error GoTo 0
    If shown Then
        MsgBox "Award saved. The email opened, but the award details couldn't be added to it (" & mailErr & ")." & _
               vbNewLine & "Type them in before sending.", vbExclamation
    Else
        MailtoFallback toList, ccList, subj, plain, "Outlook couldn't create the email (" & mailErr & ")"
    End If
End Sub

' v20: when Outlook can't be driven (new Outlook, no Outlook), open a plain
' email in the default mail app via mailto:, or show the text to copy.
Private Sub MailtoFallback(ByVal toList As String, ByVal ccList As String, ByVal subj As String, _
                           ByVal plain As String, ByVal why As String)
    Dim url As String
    url = "mailto:" & UrlEnc(Replace(Replace(toList, " ", ""), ";", ",")) & _
          "?cc=" & UrlEnc(Replace(Replace(ccList, " ", ""), ";", ",")) & _
          "&subject=" & UrlEnc(subj) & "&body=" & UrlEnc(plain)
    Dim ok As Boolean
    If Len(url) < 2000 Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink Address:=url
        ok = (Err.Number = 0)
        On Error GoTo 0
    End If
    If ok Then
        MsgBox "Award saved. " & why & ", so the email opened in your default mail app instead." & vbNewLine & _
               "Check the To and CC lines, then send.", vbInformation, "Award email"
    Else
        MsgBox "Award saved. " & why & " and no mail app could be opened." & vbNewLine & vbNewLine & _
               "Send this yourself:" & vbNewLine & vbNewLine & _
               "To: " & toList & vbNewLine & "CC: " & ccList & vbNewLine & "Subject: " & subj & vbNewLine & vbNewLine & _
               plain, vbExclamation, "Award email"
    End If
End Sub

' percent-encodes text for a mailto: link (UTF-8)
Private Function UrlEnc(ByVal t As String) As String
    Dim i As Long, ch As String, c As Long, out As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        c = AscW(ch) And &HFFFF&
        If ch Like "[A-Za-z0-9]" Or InStr("-_.~@,", ch) > 0 Then
            out = out & ch
        ElseIf c < &H80& Then
            out = out & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < &H800& Then
            out = out & "%" & Hex$(&HC0& Or (c \ &H40&)) & "%" & Hex$(&H80& Or (c And &H3F&))
        Else
            out = out & "%" & Hex$(&HE0& Or (c \ &H1000&)) & "%" & Hex$(&H80& Or ((c \ &H40&) And &H3F&)) & _
                  "%" & Hex$(&H80& Or (c And &H3F&))
        End If
    Next i
    UrlEnc = out
End Function

Private Function HtmlRow(ByVal k As String, ByVal v As String) As String
    HtmlRow = "<tr><td style='color:#555;padding-right:14px'><b>" & k & "</b></td><td>" & _
              Replace(Replace(v, "&", "&amp;"), "<", "&lt;") & "</td></tr>"
End Function


' ====================================================================
'  TABLE APPEND (Postings sheet)
' ====================================================================
Private Function AppendRows(ByVal lo As ListObject, data As Variant, ByVal nCols As Long) As Boolean
    If lo Is Nothing Then Exit Function
    Dim n As Long: n = UBound(data, 1)
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim cnt As Long: cnt = lo.ListRows.Count
    If cnt = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then cnt = 0
    End If
    Dim em As String
    On Error GoTo Fail
    ws.Unprotect Password:=SHEET_PASSWORD
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
    MsgBox "Could not write to " & lo.Name & ": " & em, vbCritical
End Function
