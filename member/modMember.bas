Attribute VB_Name = "modMember"
Option Explicit

' ====================================================================
'  CALL OUT ROSTER - MEMBERS' WORKBOOK (M1, M2)
'  This module lives in the members' file, NOT in the master. It holds
'  no rules: everything it shows was worked out by the master and
'  pushed here as values (the very hidden sheets dMeta, dRoster, dNext,
'  dCal, dPost, dPostList, dLog, dAttend). The code only sorts, filters
'  and looks those values up.
'
'  M2: Find me fills in the person's home unit on every sheet and draws
'  a "Where I stand" card on Home (hours, Refused by kind, call positions,
'  sick time COUNTS, the postings they can sign up for). The calendar is
'  the master's month grid. Postings: mine first, one plain table of all
'  open postings, and Sign me up / Withdraw, which drop ONE small text
'  file into the sign-up inbox folder the master published. That file is
'  a request: the master's Award pop-up decides what it means. This
'  workbook is never written to.
'
'  BuildMemberTemplate (Alt+F8, once, in a blank .xlsm with this module
'  and the ThisWorkbook code): lays out every sheet and button.
'  MemberRefreshAll: redraws Home and the Roster from the data (the
'  master runs it after each publish; it also runs on open).
'  Everything else is a button or a picker cell.
' ====================================================================

Public Const MEM_VERSION As String = "Members M2.1 (Jan 5)"
Private Const PW As String = "members"            ' same as MEMBER_PASSWORD in the master

Private Const SH_HOME As String = "Home"
Private Const SH_ROSTER As String = "Roster"
Private Const SH_NEXT As String = "Who's next"
Private Const SH_CAL As String = "Calendar"
Private Const SH_POST As String = "Postings"
Private Const SH_HIST As String = "My history"
Private Const SH_LISTS As String = "dLists"       ' picker lists, rebuilt by MemberRefreshAll

Private Const NAVY As Long = 6567966              ' RGB(30, 56, 100)
Private Const INK As Long = 3355443               ' RGB(51, 51, 51)
Private Const LIGHT As Long = 15787486            ' RGB(222, 229, 240)
Private Const GREEN As Long = 4620288             ' RGB(0, 128, 70)

' dRoster columns
Private Const RC_NAME As Long = 1
Private Const RC_CREW As Long = 2
Private Const RC_HOME As Long = 3
Private Const RC_TOTAL As Long = 4
Private Const RC_WORKED As Long = 5
Private Const RC_REFUSED As Long = 6
Private Const RC_PROV As Long = 7
Private Const RC_RCALL As Long = 8                ' M2: Refused = call-out + posting + other
Private Const RC_RPOST As Long = 9
Private Const RC_ROTHER As Long = 10
Private Const RC_FIXED As Long = 10

' picker cells (unlocked; a change redraws the view)
Private Const ROSTER_FIND As String = "C2"
Private Const ROSTER_CREW As String = "F2"
Private Const NEXT_UNIT As String = "B2"
Private Const NEXT_DATE As String = "D2"
Private Const NEXT_SHIFT As String = "F2"
Private Const CAL_MONTH As String = "C2"
Private Const CAL_UNIT As String = "G2"            ' a calendar group (the master's groups) or All units
Private Const HIST_NAME As String = "B2"
Private Const HOME_FIND As String = "C14"
Private Const POST_PICK As String = "C2"           ' Postings: "show the list for"
Private Const CARD_ROW As Long = 21                ' Home: the Where I stand card starts here
Private Const CAL_COLW As Double = 17              ' calendar: width of each half of a day
Private Const US As String = "~|~"                 ' separators for the calendar's per-day line lists
Private Const RS As String = "~#~"

Private mDrawing As Boolean


' ====================================================================
'  DATA ACCESS (values only)
' ====================================================================
Private Function Data(ByVal sheetName As String) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim last As Long: last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If last < 1 Or lastCol < 1 Then Exit Function
    If last = 1 And lastCol = 1 And ws.Cells(1, 1).Value = "" Then Exit Function
    Data = To2D(ws.Range(ws.Cells(1, 1), ws.Cells(last, lastCol)).Value)
End Function

Private Function To2D(ByVal v As Variant) As Variant
    If IsArray(v) Then
        To2D = v
    Else
        Dim a(1 To 1, 1 To 1) As Variant
        a(1, 1) = v
        To2D = a
    End If
End Function

Private Function Meta(ByVal key As String) As Variant
    Dim v As Variant: v = Data("dMeta")
    Meta = ""
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If StrComp(Txt(v(i, 1)), key, vbTextCompare) = 0 Then Meta = v(i, 2): Exit Function
    Next i
End Function

Private Function Txt(ByVal x As Variant) As String
    If IsError(x) Or IsNull(x) Or IsEmpty(x) Then Txt = "" Else Txt = CStr(x)
End Function

Private Function Num(ByVal x As Variant) As Double
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If IsNumeric(x) Then Num = CDbl(x)
End Function

Private Function Dt(ByVal x As Variant) As Date
    If IsDate(x) Then
        Dt = CDate(x)
    ElseIf IsNumeric(x) And Not IsEmpty(x) Then
        Dt = CDate(CDbl(x))
    End If
End Function

Private Function UnitList() As Variant
    UnitList = Split(Txt(Meta("Units")), "|")
End Function

Private Function HasData() As Boolean
    HasData = (Txt(Meta("Published at")) <> "")
End Function

' ---- M2.1: is this file what its own code expects? ----
' The supervisors' roster reads these two before it fills the template.
Public Function MemberVersion() As String
    MemberVersion = MEM_VERSION
End Function

' The version whose BuildMemberTemplate laid the sheets out (Home B19: only the build writes it)
Public Function MemberLayoutVersion() As String
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    If Not ws Is Nothing Then MemberLayoutVersion = Trim$(Txt(ws.Range("B19").Value))
End Function

' "" = fine. New code on sheets an older version laid out reads the wrong cells.
Private Function LayoutProblem() As String
    Dim built As String: built = MemberLayoutVersion()
    If StrComp(built, MEM_VERSION, vbTextCompare) <> 0 Then
        LayoutProblem = "This file's code is " & MEM_VERSION & " but its sheets were built by " & IIf(built = "", "an older version", built) & _
                        ". Supervisor: open the template, run BuildMemberTemplate, save it and publish again."
    End If
End Function

' "" = fine. A roster older than this file publishes fewer columns per person, so
' every unit would be read from the wrong column: show nothing rather than wrong figures.
Private Function DataProblem() As String
    If Not HasData() Then Exit Function
    Dim v As Variant: v = Data("dRoster")
    Dim ok As Boolean
    If IsArray(v) Then
        If UBound(v, 2) >= RC_ROTHER Then ok = (StrComp(Txt(v(1, RC_RCALL)), "Refused call-out", vbTextCompare) = 0)
    End If
    If Not ok Then
        DataProblem = "This file was filled by an older supervisors' roster (" & Txt(Meta("Master version")) & _
                      "), so its figures can't be shown. Tell your supervisor: the roster needs its update."
    End If
End Function

' One warning line on Home (row 5). "" clears it.
Private Sub HomeNote(ByVal noteTxt As String)
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    UnprotectWs ws
    With ws.Range("B5")
        .Value = noteTxt
        .Font.Bold = True: .Font.Size = 10: .Font.Color = RGB(176, 0, 32)
    End With
    ProtectWs ws
End Sub

' Every view sheet emptied, with one line pointing at Home
Private Sub BlankViews()
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array(SH_ROSTER, SH_NEXT, SH_CAL, SH_POST, SH_HIST)
        Set ws = WsOf(CStr(nm))
        If Not ws Is Nothing Then
            UnprotectWs ws
            ClearBelow ws, 4
            ws.Cells(5, 1).Value = "Not available - see the note on the Home sheet."
            ProtectWs ws
        End If
    Next nm
    Set ws = WsOf(SH_HOME)
    If Not ws Is Nothing Then
        UnprotectWs ws
        ClearBelow ws, CARD_ROW
        ProtectWs ws
    End If
End Sub


' ====================================================================
'  SHEET HELPERS
' ====================================================================
Private Function WsOf(ByVal nm As String, Optional ByVal create As Boolean = False) As Worksheet
    On Error Resume Next
    Set WsOf = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If WsOf Is Nothing And create Then
        Set WsOf = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        WsOf.Name = nm
    End If
End Function

Private Sub UnprotectWs(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=PW
    On Error GoTo 0
End Sub

Private Sub ProtectWs(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=PW, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True, _
               DrawingObjects:=True, Contents:=True, Scenarios:=True
    On Error GoTo 0
End Sub

' wipe a view sheet below its picker rows (row 4 down), keep buttons
Private Sub ClearBelow(ByVal ws As Worksheet, ByVal fromRow As Long)
    With ws.Range(ws.Rows(fromRow), ws.Rows(ws.Rows.Count))
        .ClearContents
        .ClearFormats
        .RowHeight = 15
        .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Color = INK
    End With
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

Private Sub Title(ByVal ws As Worksheet, ByVal cap As String, ByVal sub_ As String)
    ws.Rows(1).RowHeight = 28
    With ws.Range("A1")
        .Value = cap
        .Font.Size = 16: .Font.Bold = True: .Font.Color = NAVY
    End With
    With ws.Range("A3")
        .Value = sub_
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = RGB(100, 100, 100)
    End With
End Sub

Private Sub Header(ByVal ws As Worksheet, ByVal r As Long, ByVal heads As Variant)
    Dim i As Long
    For i = LBound(heads) To UBound(heads)
        ws.Cells(r, i - LBound(heads) + 1).Value = heads(i)
    Next i
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, UBound(heads) - LBound(heads) + 1))
        .Interior.Color = NAVY
        .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 9
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(r).RowHeight = 20
End Sub

Private Sub Btn(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal mac As String, _
                ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal h As Double, _
                ByVal fillCol As Long, ByVal textCol As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, l, t, w, h)
    shp.Name = nm
    shp.Adjustments.Item(1) = 0.18
    shp.Fill.ForeColor.RGB = fillCol
    shp.Line.Visible = msoFalse
    shp.Shadow.Visible = msoFalse
    shp.Placement = xlFreeFloating
    shp.OnAction = mac
    With shp.TextFrame2
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = cap
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = textCol
        End With
    End With
End Sub

Private Sub Picker(ByVal ws As Worksheet, ByVal addr As String, ByVal refTxt As String, ByVal labelTxt As String)
    Dim c As Range: Set c = ws.Range(addr)
    If labelTxt <> "" Then
        With c.Offset(0, -1)
            .Value = labelTxt
            .Font.Bold = True: .Font.Size = 9
            .HorizontalAlignment = xlRight
        End With
    End If
    With c
        .Locked = False
        .NumberFormat = "@"                        ' "Oct 2026" / "Mon Oct 12" stay text, never auto-dates
        .Interior.Color = RGB(255, 251, 230)
        .Font.Size = 10
        .BorderAround xlContinuous, xlThin, , RGB(160, 175, 200)
        .Validation.Delete
        If refTxt <> "" Then
            .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=refTxt
            .Validation.InCellDropdown = True
            .Validation.ShowError = False
        End If
    End With
End Sub


' ====================================================================
'  BUILD THE TEMPLATE (once)
' ====================================================================
Public Sub BuildMemberTemplate()
    If MsgBox("Build the members' template in THIS workbook?" & vbNewLine & vbNewLine & _
              "It adds the Home, Roster, Who's next, Calendar, Postings and My history sheets, the hidden data sheets " & _
              "and the buttons, then protects them. Run it in the blank workbook you saved as the template.", _
              vbOKCancel + vbQuestion, "Members' template") <> vbOK Then Exit Sub
    Application.ScreenUpdating = False
    Dim nm As Variant, ws As Worksheet
    ' data sheets (very hidden)
    For Each nm In Array("dMeta", "dRoster", "dNext", "dCal", "dPost", "dPostList", "dLog", "dAttend", SH_LISTS)
        Set ws = WsOf(CStr(nm), True)
        UnprotectWs ws
        ws.Visible = xlSheetVeryHidden
        ProtectWs ws
    Next nm
    ' view sheets, in order
    For Each nm In Array(SH_HOME, SH_ROSTER, SH_NEXT, SH_CAL, SH_POST, SH_HIST)
        Set ws = WsOf(CStr(nm), True)
        ws.Visible = xlSheetVisible
    Next nm
    Dim i As Long, order As Variant
    order = Array(SH_HOME, SH_ROSTER, SH_NEXT, SH_CAL, SH_POST, SH_HIST)
    For i = 0 To 5
        ThisWorkbook.Worksheets(CStr(order(i))).Move Before:=ThisWorkbook.Worksheets(i + 1)
    Next i
    ' the blank workbook's own sheet(s) - by index, backwards, so none is skipped
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set sh = ThisWorkbook.Worksheets(i)
        If sh.Visible = xlSheetVisible Then
            If InStr(1, "|" & SH_HOME & "|" & SH_ROSTER & "|" & SH_NEXT & "|" & SH_CAL & "|" & SH_POST & "|" & SH_HIST & "|", _
                     "|" & sh.Name & "|") = 0 Then sh.Delete
        End If
    Next i
    Application.DisplayAlerts = True

    LayoutHome
    LayoutRoster
    LayoutNext
    LayoutCalendar
    LayoutPostings
    LayoutHistory
    MemberRefreshAll
    ThisWorkbook.Worksheets(SH_HOME).Activate
    Application.ScreenUpdating = True
    MsgBox "Template built (" & MEM_VERSION & "). Save it as 'Call Out Roster - Members TEMPLATE.xlsm' next to the master. " & _
           "The master's Member View button fills it and publishes the members' copy.", vbInformation, "Members' template"
End Sub

Private Sub LayoutHome()
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:H").ColumnWidth = 18
    ws.Rows(2).RowHeight = 36
    With ws.Range("B2")
        .Value = "OVERTIME CALL OUT ROSTER"
        .Font.Size = 20: .Font.Bold = True: .Font.Color = NAVY
    End With
    ws.Range("B3").Font.Size = 9: ws.Range("B3").Font.Color = RGB(100, 100, 100)
    ws.Range("B4").Font.Size = 9: ws.Range("B4").Font.Color = RGB(100, 100, 100)
    Dim top As Double: top = ws.Range("B7").Top
    Dim lft As Double: lft = ws.Range("B7").Left
    Btn ws, "mb_roster", "ROSTER", "MemberGoRoster", lft, top, 150, 44, NAVY, vbWhite
    Btn ws, "mb_next", "WHO'S NEXT", "MemberGoNext", lft + 160, top, 150, 44, GREEN, vbWhite
    Btn ws, "mb_cal", "CALENDAR", "MemberGoCalendar", lft + 320, top, 150, 44, RGB(31, 56, 100), vbWhite
    Btn ws, "mb_post", "OPEN POSTINGS", "MemberGoPostings", lft, top + 54, 150, 44, RGB(93, 64, 55), vbWhite
    Btn ws, "mb_hist", "MY HISTORY", "MemberGoHistory", lft + 160, top + 54, 150, 44, RGB(106, 27, 154), vbWhite
    Picker ws, HOME_FIND, "", "Find me:"
    ws.Range(HOME_FIND).ColumnWidth = 26
    ' 26 points clear of the box: Excel draws the drop-down arrow just outside its right edge
    Btn ws, "mb_find", "Show me on the roster", "MemberFindFromHome", ws.Range("D14").Left + 26, ws.Range("D14").Top - 2, 150, 20, LIGHT, INK
    Btn ws, "mb_notme", "Not me (clear)", "MemberNotMe", ws.Range("D14").Left + 182, ws.Range("D14").Top - 2, 110, 20, vbWhite, RGB(100, 100, 100)
    With ws.Range("B17")
        .Value = "Pick your name in Find me: every sheet opens on your home unit and your own figures show below. " & _
                 "Hours and call order are worked out by the supervisors' roster and published here as values. Leave shows only as 'Leave'."
        .Font.Size = 8.5: .Font.Italic = True: .Font.Color = RGB(100, 100, 100)
    End With
    ws.Range("B17:H17").Merge
    ws.Range("B17").WrapText = True
    ws.Rows(17).RowHeight = 30
    ws.Range("B19").Value = MEM_VERSION
    ws.Range("B19").Font.Size = 8: ws.Range("B19").Font.Color = RGB(160, 165, 175)
    ProtectWs ws
End Sub

Private Sub LayoutRoster()
    Dim ws As Worksheet: Set ws = WsOf(SH_ROSTER)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "ROSTER", "Everyone, with total overtime hours and the call position in each unit (1 = called first). " & _
                        "* = total includes accepted overtime not yet in Kronos."
    ws.Columns("A").ColumnWidth = 26
    ws.Columns("B").ColumnWidth = 7
    ws.Columns("C").ColumnWidth = 14
    ws.Range("D:G").ColumnWidth = 8
    ws.Columns("H").ColumnWidth = 22                  ' M2: Refused by kind
    Picker ws, ROSTER_FIND, "", "Find me:"
    Picker ws, ROSTER_CREW, "", "Crew:"
    Dim x As Double: x = ws.Range("I1").Left
    Dim y As Double: y = ws.Range("A1").Top + 2
    Btn ws, "mb_az", "A - Z", "MemberShowAll", x, y, 60, 22, LIGHT, INK
    Btn ws, "mb_find", "Find me", "MemberFindMe", x + 66, y, 66, 22, LIGHT, INK
    Btn ws, "mb_help", "Sorted A - Z. Click a unit to see its call order.", "", x + 140, y, 260, 22, vbWhite, RGB(100, 100, 100)
    ProtectWs ws
End Sub

Private Sub LayoutNext()
    Dim ws As Worksheet: Set ws = WsOf(SH_NEXT)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "WHO'S NEXT", "Pick a unit, date and shift: everyone free for that shift, in call order (top is called first). " & _
                            "Free = not scheduled, not on the rest rule, not booked off, not on a shift switch."
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 8
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 10
    Picker ws, NEXT_UNIT, "", "Unit:"
    Picker ws, NEXT_DATE, "", "Date:"
    Picker ws, NEXT_SHIFT, "", "Shift:"
    ProtectWs ws
End Sub

Private Sub LayoutCalendar()
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 9: ws.Cells.Font.Color = INK
    Title ws, "COVERAGE CALENDAR", "Pick a month and a group. Each day: who is off on the left, who covers on the right; days (yellow) above nights (blue). " & _
                                   "OPEN = still needs cover. Leave shows only as 'Leave'."
    Dim k As Long
    ws.Columns(1).ColumnWidth = 3
    For k = 2 To 15
        ws.Columns(k).ColumnWidth = CAL_COLW
    Next k
    Picker ws, CAL_MONTH, "", "Month:"
    Picker ws, CAL_UNIT, "", "Group:"
    ws.Range(ws.Range(CAL_UNIT), ws.Range(CAL_UNIT).Offset(0, 1)).Merge
    Dim y As Double: y = ws.Rows(2).Top
    Dim x As Double: x = ws.Columns(10).Left
    Btn ws, "mb_cprev", "< Month", "MemberCalPrev", x, y, 70, 18, LIGHT, INK
    Btn ws, "mb_cnext", "Month >", "MemberCalNext", x + 76, y, 70, 18, LIGHT, INK
    Btn ws, "mb_cprint", "Print", "MemberCalPrint", x + 152, y, 60, 18, NAVY, vbWhite
    ProtectWs ws
End Sub

Private Sub LayoutPostings()
    Dim ws As Worksheet: Set ws = WsOf(SH_POST)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "OPEN POSTINGS", "Pick your name on Home to see the postings you can sign up for. Below them: every open posting in one table."
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 16
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    Picker ws, POST_PICK, "", "Show the list for:"
    ProtectWs ws
End Sub

Private Sub LayoutHistory()
    Dim ws As Worksheet: Set ws = WsOf(SH_HIST)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "MY HISTORY", "Pick a name: that person's call-outs and postings this year, newest first, with charged hours by pay period."
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 8
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 12
    ws.Columns("H").ColumnWidth = 14
    Picker ws, HIST_NAME, "", "Name:"
    ProtectWs ws
End Sub


' ====================================================================
'  REFRESH (after a publish, and on open)
' ====================================================================
Public Sub MemberRefreshAll()
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim suWas As Boolean: suWas = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mDrawing = True
    Dim stepTxt As String, note As String, en As Long, em As String
    On Error GoTo Done
    ' the file always starts with nobody picked, even if it was saved with a name showing
    stepTxt = "the pickers"
    SetCell WsOf(SH_HOME), HOME_FIND, ""
    SetCell WsOf(SH_ROSTER), ROSTER_FIND, ""
    SetCell WsOf(SH_HIST), HIST_NAME, ""
    SetCell WsOf(SH_POST), POST_PICK, ""
    stepTxt = "the lists"
    BuildLists
    stepTxt = "Home"
    DrawHome
    note = LayoutProblem()
    If note = "" Then note = DataProblem()
    If note <> "" Then
        stepTxt = "clearing the views"
        BlankViews
    Else
        stepTxt = "the Roster"
        MemberShowAll
        stepTxt = "Who's next"
        DrawNextDefault
        stepTxt = "the Calendar"
        DrawCalendarDefault
        stepTxt = "the Postings"
        MemberShowPostings
        stepTxt = "My history"
        DrawHistory ""
        stepTxt = "the Home card"
        DrawHomeCard
    End If
Done:
    en = Err.Number: em = Err.Description          ' reached normally (0) or from an error anywhere above
    On Error Resume Next
    If en <> 0 Then note = "This file didn't finish loading (" & stepTxt & ": " & em & "). Tell your supervisor."
    HomeNote note
    mDrawing = False
    MarkClean
    Application.ScreenUpdating = suWas
    Application.EnableEvents = evWas
End Sub

' redraws change nothing worth saving - no "save changes?" on close
Private Sub MarkClean()
    On Error Resume Next
    ThisWorkbook.Saved = True
End Sub

' picker lists on dLists: A names, B units (+ All units), C shifts, D next dates, E months,
' F open posting IDs, G calendar groups (M2)
Private Sub BuildLists()
    Dim ws As Worksheet: Set ws = WsOf(SH_LISTS, True)
    UnprotectWs ws
    ws.Cells.ClearContents
    ws.Columns("A:G").NumberFormat = "@"           ' lists are text: "Oct 2026" must not become a date
    Dim v As Variant, i As Long, nNames As Long, nUnits As Long, nMonths As Long, nPosts As Long, nGroups As Long
    v = Data("dRoster")
    If IsArray(v) Then
        Dim names() As String: ReDim names(1 To UBound(v, 1))
        For i = 2 To UBound(v, 1)
            If Txt(v(i, RC_NAME)) <> "" Then nNames = nNames + 1: names(nNames) = Txt(v(i, RC_NAME))
        Next i
        If nNames > 0 Then
            ReDim Preserve names(1 To nNames)
            SortStrings names, nNames
            For i = 1 To nNames
                ws.Cells(i, 1).Value = names(i)
            Next i
        End If
    End If
    Dim u As Variant: u = UnitList()
    If UBound(u) >= LBound(u) Then
        For i = LBound(u) To UBound(u)
            If Txt(u(i)) <> "" Then nUnits = nUnits + 1: ws.Cells(nUnits, 2).Value = u(i)
        Next i
    End If
    If nUnits > 0 Then ws.Cells(nUnits + 1, 2).Value = "All units"
    ws.Cells(1, 3).Value = "Days": ws.Cells(2, 3).Value = "Nights"
    Dim nDays As Long: nDays = CLng(Num(Meta("Who's next days")))
    If nDays <= 0 Then nDays = 14
    Dim d0 As Date: d0 = Dt(Meta("Published at"))
    If d0 = 0 Then d0 = Date
    For i = 0 To nDays - 1
        ws.Cells(i + 1, 4).Value = Format$(Int(d0) + i, "ddd mmm d")
    Next i
    ' months present in dCal
    v = Data("dCal")
    If IsArray(v) Then
        Dim seen As String, m As String
        For i = 2 To UBound(v, 1)
            m = Format$(Dt(v(i, 1)), "mmm yyyy")
            If InStr("|" & seen, "|" & m & "|") = 0 Then
                seen = seen & m & "|"
                nMonths = nMonths + 1
                ws.Cells(nMonths, 5).Value = m
            End If
        Next i
    End If
    ' open postings (M2)
    v = Data("dPost")
    If IsArray(v) Then
        For i = 2 To UBound(v, 1)
            If Txt(v(i, 1)) <> "" Then nPosts = nPosts + 1: ws.Cells(nPosts, 6).Value = Txt(v(i, 1))
        Next i
    End If
    ' calendar groups, as the master's calendar offers them (M2)
    Dim g As Variant
    For Each g In Split(Txt(Meta("Calendar groups")), ";")
        If InStr(CStr(g), "|") > 0 Then nGroups = nGroups + 1: ws.Cells(nGroups, 7).Value = Split(CStr(g), "|")(0)
    Next g
    ProtectWs ws
    ' point the pickers at the lists (an empty list = no drop-down until the first publish)
    Dim crews As String: crews = Replace(Txt(Meta("Crews")), "|", ",")
    If crews = "" Then crews = "A,B,A/B,C,D,C/D,DAYS"
    SetList WsOf(SH_ROSTER), ROSTER_FIND, ListRef("A", nNames)
    SetList WsOf(SH_ROSTER), ROSTER_CREW, crews & ",All"
    SetList WsOf(SH_NEXT), NEXT_UNIT, ListRef("B", nUnits)
    SetList WsOf(SH_NEXT), NEXT_DATE, ListRef("D", nDays)
    SetList WsOf(SH_NEXT), NEXT_SHIFT, ListRef("C", 2)
    SetList WsOf(SH_CAL), CAL_MONTH, ListRef("E", nMonths)
    SetList WsOf(SH_CAL), CAL_UNIT, ListRef("G", nGroups)
    SetList WsOf(SH_HIST), HIST_NAME, ListRef("A", nNames)
    SetList WsOf(SH_HOME), HOME_FIND, ListRef("A", nNames)
    SetList WsOf(SH_POST), POST_PICK, ListRef("F", nPosts)
End Sub

Private Function ListRef(ByVal col As String, ByVal n As Long) As String
    If n > 0 Then ListRef = "='" & SH_LISTS & "'!$" & col & "$1:$" & col & "$" & n
End Function

Private Sub SetList(ByVal ws As Worksheet, ByVal addr As String, ByVal refTxt As String)
    UnprotectWs ws
    With ws.Range(addr).Validation
        .Delete
        If refTxt <> "" Then
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=refTxt
            .InCellDropdown = True
            .ShowError = False
        End If
    End With
    ws.Range(addr).MergeArea.Locked = False        ' the calendar's Group picker is two merged cells
    ProtectWs ws
End Sub

Private Sub SortStrings(ByRef a() As String, ByVal n As Long)
    Dim i As Long, j As Long, t As String
    For i = 2 To n
        t = a(i): j = i - 1
        Do While j >= 1
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

Private Sub DrawHome()
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    UnprotectWs ws
    If HasData() Then
        ws.Range("B3").Value = "Updated " & Format$(Dt(Meta("Published at")), "ddd mmm d, yyyy h:mm AM/PM") & _
                               "   |   Kronos data " & Format$(Dt(Meta("Kronos data date")), "mmm d") & _
                               "   |   pay period " & Format$(Dt(Meta("Pay period start")), "mmm d") & " - " & _
                               Format$(Dt(Meta("Pay period end")), "mmm d")
        ws.Range("B4").Value = "Roster " & Txt(Meta("Master version")) & "   |   charged: " & Txt(Meta("Charge rule")) & _
                               "   |   members file " & MEM_VERSION
    Else
        ws.Range("B3").Value = "Not published yet - the supervisors' roster fills this file with its Member View button."
        ws.Range("B4").Value = ""
    End If
    ProtectWs ws
End Sub


' ====================================================================
'  ROSTER VIEW
' ====================================================================
Public Sub MemberShowAll()
    DrawRoster "", "A - Z"
End Sub

' called from the unit buttons: OnAction "'MemberSortUnit ""Phosphine""'"
Public Sub MemberSortUnit(ByVal unit As String)
    DrawRoster unit, unit & " call order"
End Sub

' Find me on the Roster sheet (button, or its picker cell)
Public Sub MemberFindMe()
    Dim nm As String: nm = Txt(WsOf(SH_ROSTER).Range(ROSTER_FIND).Value)
    If nm = "" Then MsgBox "Pick your name in the Find me box first.", vbInformation: Exit Sub
    ApplyMe nm, SH_ROSTER
End Sub

' Home: the "Show me on the roster" button
Public Sub MemberFindFromHome()
    Dim nm As String: nm = WhoAmI()
    If nm = "" Then MsgBox "Pick your name first.", vbInformation: Exit Sub
    ApplyMe nm, SH_ROSTER
End Sub

' Home: back to the neutral views
Public Sub MemberNotMe()
    ApplyMe "", SH_HOME
End Sub

' The name picked in Find me on Home ("" = nobody). Nothing is remembered
' between opens: the published file always starts with no name.
Private Function WhoAmI() As String
    WhoAmI = Txt(WsOf(SH_HOME).Range(HOME_FIND).Value)
End Function

Private Sub SetCell(ByVal ws As Worksheet, ByVal addr As String, ByVal txtVal As String)
    UnprotectWs ws
    ws.Range(addr).Value = txtVal
    ProtectWs ws
End Sub

' M2: one name drives every sheet. The Roster opens sorted for the person's
' home unit with their row highlighted; Who's next and the Calendar open on
' the same unit; My history, the Postings and the Home card show theirs.
Private Sub ApplyMe(ByVal nm As String, ByVal stayOn As String)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim suWas As Boolean: suWas = Application.ScreenUpdating
    Dim homeU As String, grp As String, en As Long, em As String
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    mDrawing = True
    On Error GoTo Done
    SetCell WsOf(SH_HOME), HOME_FIND, nm
    SetCell WsOf(SH_ROSTER), ROSTER_FIND, nm
    SetCell WsOf(SH_HIST), HIST_NAME, nm
    homeU = HomeUnitOf(nm)
    If homeU <> "" Then
        If UnitCodeCol(homeU) = 0 Then homeU = ""          ' a home unit the published unit list doesn't have
    End If
    If nm <> "" Then SetCell WsOf(SH_ROSTER), ROSTER_CREW, "All"      ' a crew filter left on would hide them
    If homeU <> "" Then
        SetCell WsOf(SH_NEXT), NEXT_UNIT, homeU
        grp = GroupOfUnit(homeU)
        If grp <> "" Then SetCell WsOf(SH_CAL), CAL_UNIT, grp
        DrawRoster homeU, homeU & " call order", nm
    ElseIf nm <> "" Then
        DrawRoster "", "A - Z  (no home unit on file for " & nm & ")", nm
    Else
        DrawRoster "", "A - Z"
    End If
    MemberShowNext
    MemberShowCalendar
    DrawHistory nm
    MemberShowPostings
    DrawHomeCard
Done:
    en = Err.Number: em = Err.Description          ' reached normally (0) or from an error anywhere above
    On Error Resume Next
    HomeNote IIf(en <> 0, "Couldn't show everything for " & nm & " (" & em & "). Tell your supervisor.", "")
    mDrawing = False
    WsOf(stayOn).Activate
    If stayOn = SH_HOME Then WsOf(SH_HOME).Range(HOME_FIND).Select
    MarkClean
    Application.ScreenUpdating = suWas
    Application.EnableEvents = evWas
End Sub

' ---- lookups in the published roster (values only) ----
Private Function RosterRowOf(ByRef v As Variant, ByVal nm As String) As Long
    If Not IsArray(v) Or nm = "" Then Exit Function
    Dim i As Long
    For i = 2 To UBound(v, 1)
        If StrComp(Txt(v(i, RC_NAME)), nm, vbTextCompare) = 0 Then RosterRowOf = i: Exit Function
    Next i
End Function

Private Function HomeUnitOf(ByVal nm As String) As String
    Dim v As Variant: v = Data("dRoster")
    Dim r As Long: r = RosterRowOf(v, nm)
    If r > 0 Then HomeUnitOf = Txt(v(r, RC_HOME))
End Function

' dRoster column of a unit's code (its call position is the next column); 0 = unknown unit
Private Function UnitCodeCol(ByVal unit As String) As Long
    Dim units As Variant: units = UnitList()
    Dim i As Long
    For i = LBound(units) To UBound(units)
        If StrComp(Txt(units(i)), unit, vbTextCompare) = 0 Then UnitCodeCol = RC_FIXED + 2 * (i - LBound(units)) + 1: Exit Function
    Next i
End Function

' "63  (call-out 42 / posting 21)" - "other" only when it isn't zero
Private Function RefusedKinds(ByRef v As Variant, ByVal r As Long) As String
    If UBound(v, 2) < RC_ROTHER Then Exit Function
    If StrComp(Txt(v(1, RC_RCALL)), "Refused call-out", vbTextCompare) <> 0 Then Exit Function    ' data from an older roster
    RefusedKinds = "call-out " & NumTxt(Num(v(r, RC_RCALL))) & " / posting " & NumTxt(Num(v(r, RC_RPOST)))
    If Abs(Num(v(r, RC_ROTHER))) > 0.0001 Then RefusedKinds = RefusedKinds & " / other " & NumTxt(Num(v(r, RC_ROTHER)))
End Function

' 8 -> "8", 22.75 -> "22.75" (Format$(8, "0.##") gives "8.")
Private Function NumTxt(ByVal x As Double) As String
    NumTxt = CStr(Round(x, 2))
End Function

' The master's calendar groups: "Name|Unit,Unit;Name|Unit;All units|*"
Private Function InGroup(ByVal unit As String, ByVal groupName As String) As Boolean
    Dim g As Variant, parts As Variant, u As Variant
    If groupName = "" Or StrComp(groupName, "All units", vbTextCompare) = 0 Then InGroup = True: Exit Function
    For Each g In Split(Txt(Meta("Calendar groups")), ";")
        parts = Split(CStr(g), "|")
        If UBound(parts) >= 1 Then
            If StrComp(CStr(parts(0)), groupName, vbTextCompare) = 0 Then
                If CStr(parts(1)) = "*" Then InGroup = True: Exit Function
                For Each u In Split(CStr(parts(1)), ",")
                    If StrComp(Trim$(CStr(u)), Trim$(unit), vbTextCompare) = 0 Then InGroup = True: Exit Function
                Next u
                Exit Function
            End If
        End If
    Next g
    InGroup = (StrComp(unit, groupName, vbTextCompare) = 0)        ' a plain unit name still works
End Function

' The first group (not "All units") that holds the unit
Private Function GroupOfUnit(ByVal unit As String) As String
    Dim g As Variant, parts As Variant, u As Variant
    For Each g In Split(Txt(Meta("Calendar groups")), ";")
        parts = Split(CStr(g), "|")
        If UBound(parts) >= 1 Then
            If CStr(parts(1)) <> "*" Then
                For Each u In Split(CStr(parts(1)), ",")
                    If StrComp(Trim$(CStr(u)), Trim$(unit), vbTextCompare) = 0 Then GroupOfUnit = CStr(parts(0)): Exit Function
                Next u
            End If
        End If
    Next g
End Function

Private Sub DrawRoster(ByVal unit As String, ByVal viewTxt As String, Optional ByVal highlight As String = "")
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
    Dim ws As Worksheet: Set ws = WsOf(SH_ROSTER)
    Dim v As Variant: v = Data("dRoster")
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    ClearBelow ws, 4
    Dim units As Variant: units = UnitList()
    Dim nU As Long: nU = UBound(units) - LBound(units) + 1
    Dim crew As String: crew = UCase$(Txt(ws.Range(ROSTER_CREW).Value))
    If crew = "ALL" Then crew = ""

    ' buttons: one per unit, inside row 4 (rebuilt each draw so the active one shows)
    Dim i As Long, x As Double, y As Double
    For i = ws.Shapes.Count To 1 Step -1               ' old unit buttons go first (the unit list may change)
        If Left$(ws.Shapes(i).Name, 4) = "mb_u" Then ws.Shapes(i).Delete
    Next i
    ws.Rows(4).RowHeight = 30
    x = ws.Range("A4").Left: y = ws.Range("A4").Top + 4
    If UBound(units) >= LBound(units) Then
        For i = 0 To nU - 1
            Btn ws, "mb_u" & i, Txt(units(i)), "'MemberSortUnit """ & Txt(units(i)) & """'", x + i * 100, y, 96, 22, _
                IIf(Txt(units(i)) = unit, NAVY, LIGHT), IIf(Txt(units(i)) = unit, vbWhite, INK)
        Next i
    End If

    Dim heads As Variant
    If unit = "" Then
        heads = Array("Name", "Crew", "Home unit", "Total", "Worked", "Refused", "Refused by kind")
    Else
        heads = Array("#", "Name", "Crew", "Code", "Total", "Worked", "Refused", "Refused by kind")
    End If
    Header ws, 5, heads
    ws.Range("A3").Value = "View: " & viewTxt & IIf(crew <> "", "  |  crew " & crew, "") & _
                           ".  * = total includes accepted overtime not yet in Kronos."

    If Not IsArray(v) Then ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    ' collect rows
    Dim idx() As Long, keys() As Double, n As Long, uCode As Long, uPos As Long
    ReDim idx(1 To UBound(v, 1)): ReDim keys(1 To UBound(v, 1))
    If unit <> "" Then
        For i = LBound(units) To UBound(units)
            If StrComp(Txt(units(i)), unit, vbTextCompare) = 0 Then
                uCode = RC_FIXED + 2 * (i - LBound(units)) + 1
                uPos = uCode + 1
            End If
        Next i
    End If
    For i = 2 To UBound(v, 1)
        If Txt(v(i, RC_NAME)) <> "" Then
            If crew = "" Or UCase$(Txt(v(i, RC_CREW))) = crew Then
                If unit = "" Then
                    n = n + 1: idx(n) = i: keys(n) = 0
                ElseIf uCode > 0 Then                  ' (nested: VBA would read column 0 in an And)
                    If Txt(v(i, uCode)) <> "" Then n = n + 1: idx(n) = i: keys(n) = Num(v(i, uPos))
                End If
            End If
        End If
    Next i
    ' order: A-Z (already published A-Z? sort anyway) or by position
    Dim j As Long, t As Long, tk As Double
    For i = 2 To n
        t = idx(i): tk = keys(i): j = i - 1
        Do While j >= 1
            If unit = "" Then
                If StrComp(Txt(v(idx(j), RC_NAME)), Txt(v(t, RC_NAME)), vbTextCompare) <= 0 Then Exit Do
            Else
                If keys(j) <= tk Then Exit Do
            End If
            idx(j + 1) = idx(j): keys(j + 1) = keys(j): j = j - 1
        Loop
        idx(j + 1) = t: keys(j + 1) = tk
    Next i

    Dim out() As Variant, nc As Long: nc = UBound(heads) + 1
    If n = 0 Then
        ws.Cells(6, 1).Value = "Nobody to show."
        ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    End If
    ReDim out(1 To n, 1 To nc)
    Dim r As Long, hiRow As Long
    For r = 1 To n
        i = idx(r)
        If unit = "" Then
            out(r, 1) = Txt(v(i, RC_NAME)): out(r, 2) = Txt(v(i, RC_CREW)): out(r, 3) = Txt(v(i, RC_HOME))
            out(r, 4) = HoursText(v(i, RC_TOTAL), v(i, RC_PROV)): out(r, 5) = Num(v(i, RC_WORKED)): out(r, 6) = Num(v(i, RC_REFUSED))
            out(r, 7) = RefusedKinds(v, i)
        Else
            out(r, 1) = keys(r): out(r, 2) = Txt(v(i, RC_NAME)): out(r, 3) = Txt(v(i, RC_CREW)): out(r, 4) = Txt(v(i, uCode))
            out(r, 5) = HoursText(v(i, RC_TOTAL), v(i, RC_PROV)): out(r, 6) = Num(v(i, RC_WORKED)): out(r, 7) = Num(v(i, RC_REFUSED))
            out(r, 8) = RefusedKinds(v, i)
        End If
        If highlight <> "" Then
            If StrComp(Txt(v(i, RC_NAME)), highlight, vbTextCompare) = 0 Then hiRow = 5 + r
        End If
    Next r
    With ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, nc))
        .Value = out
        .Font.Size = 10
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(228, 232, 238)
    End With
    ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, nc)).RowHeight = 17
    ws.Range(ws.Cells(6, nc), ws.Cells(5 + n, nc)).Font.Size = 8.5
    ws.Range(ws.Cells(6, nc), ws.Cells(5 + n, nc)).Font.Color = RGB(100, 100, 100)
    If unit = "" Then
        ws.Range(ws.Cells(6, 4), ws.Cells(5 + n, 6)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(6, 2), ws.Cells(5 + n, 3)).HorizontalAlignment = xlCenter
    Else
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(6, 3), ws.Cells(5 + n, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).Font.Bold = True
    End If
    If hiRow > 0 Then
        With ws.Range(ws.Cells(hiRow, 1), ws.Cells(hiRow, nc))
            .Interior.Color = RGB(255, 236, 179)
            .Font.Bold = True
        End With
        Application.Goto ws.Cells(hiRow, 1), True
    End If
    ' AutoFilter so the A-Z view still filters with macros off
    If unit = "" Then ws.Range(ws.Cells(5, 1), ws.Cells(5 + n, nc)).AutoFilter
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub

Private Function HoursText(ByVal total As Variant, ByVal prov As Variant) As Variant
    If Num(prov) > 0 Then
        HoursText = Format$(Num(total), "#,##0") & "*"
    Else
        HoursText = Num(total)
    End If
End Function


' ====================================================================
'  WHO'S NEXT
' ====================================================================
Private Sub DrawNextDefault()
    If Not HasData() Then Exit Sub                     ' template build: nothing to show yet
    Dim ws As Worksheet: Set ws = WsOf(SH_NEXT)
    Dim u As Variant: u = UnitList()
    UnprotectWs ws
    If Txt(ws.Range(NEXT_UNIT).Value) = "" And UBound(u) >= LBound(u) Then ws.Range(NEXT_UNIT).Value = Txt(u(LBound(u)))
    ws.Range(NEXT_DATE).NumberFormat = "@"
    ws.Range(NEXT_DATE).Value = Format$(Int(Dt(Meta("Published at"))), "ddd mmm d")
    If Txt(ws.Range(NEXT_SHIFT).Value) = "" Then ws.Range(NEXT_SHIFT).Value = "Days"
    ProtectWs ws
    MemberShowNext
End Sub

Public Sub MemberShowNext()
    Dim ws As Worksheet: Set ws = WsOf(SH_NEXT)
    Dim unit As String: unit = Txt(ws.Range(NEXT_UNIT).Value)
    Dim dTxt As String: dTxt = Txt(ws.Range(NEXT_DATE).Value)
    Dim dPick As Date                              ' if Excel turned the picker into a real date, match on that
    If VarType(ws.Range(NEXT_DATE).Value) = vbDate Then dPick = Int(ws.Range(NEXT_DATE).Value): dTxt = Format$(dPick, "ddd mmm d")
    Dim sh As String: sh = Txt(ws.Range(NEXT_SHIFT).Value)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    On Error GoTo Fail
    ClearBelow ws, 4
    Header ws, 5, Array("#", "Name", "Crew", "Code", "Total")
    Dim v As Variant: v = Data("dNext")
    Dim i As Long, n As Long, out() As Variant
    If IsArray(v) And unit <> "" And dTxt <> "" And sh <> "" Then
        ReDim out(1 To UBound(v, 1), 1 To 5)
        For i = 2 To UBound(v, 1)
            If StrComp(Txt(v(i, 1)), unit, vbTextCompare) = 0 And StrComp(Txt(v(i, 3)), sh, vbTextCompare) = 0 And _
               (Format$(Dt(v(i, 2)), "ddd mmm d") = dTxt Or (dPick > 0 And Int(Dt(v(i, 2))) = dPick)) Then
                n = n + 1
                out(n, 1) = Num(v(i, 4)): out(n, 2) = Txt(v(i, 5)): out(n, 3) = "": out(n, 4) = Txt(v(i, 8))
                out(n, 5) = HoursText(v(i, 6), v(i, 7))
            End If
        Next i
    End If
    ws.Range("A3").Value = IIf(n = 0, "Nobody free", n & " free") & " for " & unit & ", " & dTxt & ", " & sh & _
                           ".  Order: grade code, then total hours, then plant seniority.  (as of " & _
                           Format$(Dt(Meta("Published at")), "mmm d h:mm AM/PM") & ")"
    If n > 0 Then
        ' crew from the roster
        Dim rv As Variant: rv = Data("dRoster")
        Dim j As Long
        If IsArray(rv) Then
            For i = 1 To n
                For j = 2 To UBound(rv, 1)
                    If StrComp(Txt(rv(j, RC_NAME)), CStr(out(i, 2)), vbTextCompare) = 0 Then out(i, 3) = Txt(rv(j, RC_CREW)): Exit For
                Next j
            Next i
        End If
        With ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 5))
            .Value = out
            .RowHeight = 17
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(228, 232, 238)
        End With
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).Font.Bold = True
        ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(6, 3), ws.Cells(5 + n, 5)).HorizontalAlignment = xlCenter
        With ws.Range(ws.Cells(6, 1), ws.Cells(6, 5))
            .Interior.Color = RGB(198, 239, 206)
        End With
    Else
        ws.Cells(6, 2).Value = "No one is free for that shift (or the date is outside the published " & Txt(Meta("Who's next days")) & " days)."
        ws.Cells(6, 2).Font.Italic = True
    End If
Fail:                                            ' reached normally, or from an error anywhere above
    On Error Resume Next
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  CALENDAR
' ====================================================================
Private Sub DrawCalendarDefault()
    If Not HasData() Then Exit Sub                     ' template build: nothing to show yet
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    UnprotectWs ws
    ws.Range(CAL_MONTH).NumberFormat = "@"
    ws.Range(CAL_MONTH).Value = Format$(Dt(Meta("Published at")), "mmm yyyy")
    If Txt(ws.Range(CAL_UNIT).Value) = "" Then ws.Range(CAL_UNIT).Value = "All units"
    ProtectWs ws
    MemberShowCalendar
End Sub

' the master's TypeColor values, copied (no rule here, just the same colours)
Private Function LineColour(ByVal code As String) As Long
    Select Case UCase$(code)
        Case "VAC": LineColour = RGB(21, 101, 192)
        Case "LV", "LOA", "STD", "PAR", "WSIB", "LTD", "OLV": LineColour = RGB(106, 27, 154)
        Case "SIC": LineColour = RGB(230, 81, 0)
        Case "PER": LineColour = RGB(69, 90, 100)
        Case "TRN": LineColour = RGB(0, 121, 107)
        Case "PHA": LineColour = RGB(46, 125, 50)
        Case "SUP": LineColour = RGB(93, 64, 55)
        Case "SS": LineColour = RGB(194, 24, 91)
        Case "OPEN": LineColour = RGB(190, 30, 30)
        Case Else: LineColour = RGB(97, 97, 97)
    End Select
End Function

' The month shown, as its first day (0 = can't be read)
Private Function CalMonthStart(ByVal ws As Worksheet) As Date
    Dim mTxt As String: mTxt = Txt(ws.Range(CAL_MONTH).Value)
    Dim m0 As Date
    If VarType(ws.Range(CAL_MONTH).Value) = vbDate Then
        m0 = ws.Range(CAL_MONTH).Value
    ElseIf mTxt <> "" Then
        On Error Resume Next
        m0 = CDate("1 " & mTxt)
        If m0 = 0 Then m0 = CDate(mTxt)
        On Error GoTo 0
    End If
    If m0 <> 0 Then CalMonthStart = DateSerial(Year(m0), Month(m0), 1)
End Function

' One day's lines for the group: each line = left US right US colour-code, joined
' with RS. Fills the holiday name and the crews on days / nights for that date.
Private Sub DayLines(ByRef v As Variant, ByVal rowList As Collection, ByVal grp As String, _
                     ByRef dayTxt As String, ByRef nightTxt As String, ByRef hol As String, _
                     ByRef crewD As String, ByRef crewN As String)
    dayTxt = "": nightTxt = "": hol = "": crewD = "": crewN = ""
    If rowList Is Nothing Then Exit Sub
    Dim k As Variant, i As Long, u As String, line As String
    For Each k In rowList
        i = CLng(k)
        u = Txt(v(i, 3))
        If u = "*" Then
            hol = Txt(v(i, 4))
        ElseIf u = "#" Then
            crewD = Txt(v(i, 4)): crewN = Txt(v(i, 5))
        ElseIf InGroup(u, grp) Then
            line = Txt(v(i, 4)) & US & IIf(Txt(v(i, 5)) <> "", Txt(v(i, 5)), IIf(CBool(v(i, 7)), "OPEN", "")) & US & _
                   IIf(CBool(v(i, 7)), "OPEN", Txt(v(i, 6)))
            If Txt(v(i, 2)) = "Nights" Then
                nightTxt = nightTxt & IIf(nightTxt = "", "", RS) & line
            Else
                dayTxt = dayTxt & IIf(dayTxt = "", "", RS) & line
            End If
        End If
    Next k
End Sub

Private Function LineCount(ByVal lines As String) As Long
    If lines <> "" Then LineCount = UBound(Split(lines, RS)) + 1
End Function

' Writes one block of lines (a day's days or nights) into the pair of columns c, c + 1
Private Sub PaintLines(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal lines As String)
    If lines = "" Then Exit Sub
    Dim arr As Variant: arr = Split(lines, RS)
    Dim i As Long, p As Variant, col As Long
    For i = 0 To UBound(arr)
        p = Split(CStr(arr(i)), US)
        If UBound(p) >= 2 Then
            col = LineColour(CStr(p(2)))
            ws.Cells(r + i, c).Value = CStr(p(0))
            ws.Cells(r + i, c).Font.Color = col
            ws.Cells(r + i, c + 1).Value = CStr(p(1))
            If CStr(p(1)) = "OPEN" Then
                ws.Cells(r + i, c + 1).Font.Color = RGB(190, 30, 30)
                ws.Cells(r + i, c + 1).Font.Bold = True
            ElseIf Left$(CStr(p(1)), 6) = "POSTED" Then
                ws.Cells(r + i, c + 1).Font.Color = RGB(180, 100, 0)
            End If
        End If
    Next i
End Sub

' M2: the master's month grid. Week rows, Sunday to Saturday; each day is two
' columns (who's off | covered by) with the day lines above the night lines.
Public Sub MemberShowCalendar()
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    Dim grp As String: grp = Txt(ws.Range(CAL_UNIT).Value)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    On Error GoTo Finish
    ClearBelow ws, 4
    Dim v As Variant: v = Data("dCal")
    Dim m0 As Date: m0 = CalMonthStart(ws)
    ws.Range("A3").Value = "Month: " & IIf(m0 = 0, "", Format$(m0, "mmmm yyyy")) & "   |   " & IIf(grp = "", "All units", grp) & _
                           "   |   OPEN = still needs cover.  Leave shows only as 'Leave'."
    If Not IsArray(v) Or m0 = 0 Then GoTo Finish

    ' rows of dCal by date, so each day is one lookup
    Dim byDate As Object: Set byDate = CreateObject("Scripting.Dictionary")
    Dim i As Long, dk As Long, bag As Collection
    For i = 2 To UBound(v, 1)
        dk = CLng(Int(CDbl(Dt(v(i, 1)))))
        If byDate.Exists(dk) Then
            Set bag = byDate(dk)
        Else
            Set bag = New Collection
            byDate.Add dk, bag
        End If
        bag.Add i
    Next i

    Dim k As Long, r As Long: r = 5
    Dim dn As Variant: dn = Array("SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY")
    ws.Rows(r).RowHeight = 20
    For k = 0 To 6
        With ws.Range(ws.Cells(r, 2 + 2 * k), ws.Cells(r, 3 + 2 * k))
            .Merge
            .Value = dn(k)
            .Font.Bold = True: .Font.Size = 9: .Font.Color = vbWhite
            .Interior.Color = NAVY
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        End With
    Next k
    r = r + 1

    Dim mEnd As Date: mEnd = DateSerial(Year(m0), Month(m0) + 1, 0)
    Dim wk As Date: wk = m0 - (Weekday(m0, vbSunday) - 1)
    Dim d As Date, inMonth As Boolean, bg As Long, mxD As Long, mxN As Long, c As Long
    Dim dTxt(0 To 6) As String, nTxt(0 To 6) As String, hol(0 To 6) As String, cD(0 To 6) As String, cN(0 To 6) As String
    Do While wk <= mEnd
        mxD = 1: mxN = 1
        For k = 0 To 6
            d = wk + k
            dTxt(k) = "": nTxt(k) = "": hol(k) = "": cD(k) = "": cN(k) = ""
            If Month(d) = Month(m0) Then
                Set bag = Nothing
                If byDate.Exists(CLng(d)) Then Set bag = byDate(CLng(d))
                DayLines v, bag, grp, dTxt(k), nTxt(k), hol(k), cD(k), cN(k)
                If LineCount(dTxt(k)) > mxD Then mxD = LineCount(dTxt(k))
                If LineCount(nTxt(k)) > mxN Then mxN = LineCount(nTxt(k))
            End If
        Next k
        ' the date row: the day number (and holiday) | the crews on days / nights
        ws.Rows(r).RowHeight = 28
        For k = 0 To 6
            d = wk + k
            c = 2 + 2 * k
            inMonth = (Month(d) = Month(m0))
            If Not inMonth Then
                bg = RGB(236, 239, 243)
            ElseIf Int(d) = Int(Date) Then
                bg = RGB(255, 224, 130)
            ElseIf k = 0 Or k = 6 Then
                bg = RGB(245, 247, 250)
            Else
                bg = vbWhite
            End If
            ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1)).Interior.Color = bg
            If inMonth Then
                With ws.Cells(r, c)
                    .NumberFormat = "@"
                    .Value = Format$(d, "d") & IIf(hol(k) <> "", "  " & hol(k), "")
                    .Font.Bold = True: .Font.Size = 11
                    .VerticalAlignment = xlCenter
                    If hol(k) <> "" Then .Font.Color = RGB(176, 0, 32)
                End With
                With ws.Cells(r, c + 1)
                    If cD(k) <> "" Then .Value = "Days " & cD(k) & "  |  Nights " & cN(k)
                    .Font.Size = 7.5: .Font.Color = RGB(100, 100, 100)
                    .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
                End With
            End If
        Next k
        r = r + 1
        ' day lines (yellow), then night lines (blue)
        For k = 0 To 6
            d = wk + k
            c = 2 + 2 * k
            If Month(d) = Month(m0) Then
                ws.Range(ws.Cells(r, c), ws.Cells(r + mxD - 1, c + 1)).Interior.Color = RGB(255, 249, 196)
                ws.Range(ws.Cells(r + mxD, c), ws.Cells(r + mxD + mxN - 1, c + 1)).Interior.Color = RGB(227, 236, 250)
                PaintLines ws, r, c, dTxt(k)
                PaintLines ws, r + mxD, c, nTxt(k)
            Else
                ws.Range(ws.Cells(r, c), ws.Cells(r + mxD + mxN - 1, c + 1)).Interior.Color = RGB(236, 239, 243)
            End If
            With ws.Range(ws.Cells(r - 1, c), ws.Cells(r + mxD + mxN - 1, c + 1))
                .BorderAround xlContinuous, xlThin, , RGB(160, 175, 200)
            End With
        Next k
        With ws.Range(ws.Cells(r, 2), ws.Cells(r + mxD + mxN - 1, 15))
            .Font.Size = 8.5
            .ShrinkToFit = True
            .RowHeight = 14
        End With
        r = r + mxD + mxN
        wk = wk + 7
    Loop
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 15)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    On Error GoTo 0
Finish:
    On Error Resume Next
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub

' < Month / Month >: the previous or next month the roster published
Private Sub StepMonth(ByVal delta As Long)
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    Dim ls As Worksheet: Set ls = WsOf(SH_LISTS)
    Dim m0 As Date: m0 = CalMonthStart(ws)
    Dim cur As String: cur = IIf(m0 = 0, Txt(ws.Range(CAL_MONTH).Value), Format$(m0, "mmm yyyy"))
    Dim n As Long: n = ls.Cells(ls.Rows.Count, 5).End(xlUp).Row
    Dim i As Long, at As Long
    For i = 1 To n
        If StrComp(Txt(ls.Cells(i, 5).Value), cur, vbTextCompare) = 0 Then at = i: Exit For
    Next i
    If at = 0 Then at = 1 Else at = at + delta
    Dim okMonth As Boolean
    If at >= 1 And at <= n Then okMonth = (Txt(ls.Cells(at, 5).Value) <> "")      ' (VBA tests both sides of an Or: never read row 0)
    If Not okMonth Then
        MsgBox "That is the " & IIf(delta < 0, "first", "last") & " month the roster published.", vbInformation, "Calendar"
        Exit Sub
    End If
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    UnprotectWs ws
    ws.Range(CAL_MONTH).NumberFormat = "@"
    ws.Range(CAL_MONTH).Value = Txt(ls.Cells(at, 5).Value)
    ProtectWs ws
    Application.EnableEvents = evWas
    MemberShowCalendar
End Sub

Public Sub MemberCalPrev()
    StepMonth -1
End Sub

Public Sub MemberCalNext()
    StepMonth 1
End Sub

Public Sub MemberCalPrint()
    On Error Resume Next
    WsOf(SH_CAL).PrintOut Copies:=1
    If Err.Number <> 0 Then MsgBox "The calendar couldn't be printed: " & Err.Description, vbExclamation, "Calendar"
End Sub


' ====================================================================
'  POSTINGS
' ====================================================================
Private Function CloseTime() As String        ' published by the master; a default only until the first publish
    CloseTime = Txt(Meta("Posting close time"))
    If CloseTime = "" Then CloseTime = "8:00 AM"
End Function

' The moment a posting stops taking sign-ups: its closing date at the published time
Private Function ClosesAt(ByVal closesOn As Date) As Date
    ClosesAt = Int(closesOn)
    If IsDate(CloseTime()) Then ClosesAt = ClosesAt + TimeValue(CloseTime())
End Function

Private Function SignupsOn() As Boolean
    SignupsOn = (LCase$(Txt(Meta("Sign-ups"))) = "on")
End Function

' The person's place on a posting's published list: their rank (1 = first), 0 = not
' on the list. signedTxt = what the master published for them ("" / "Yes" / "8 h").
Private Function MyPlace(ByRef l As Variant, ByVal postId As String, ByVal nm As String, ByRef signedTxt As String) As Long
    signedTxt = ""
    If Not IsArray(l) Then Exit Function
    Dim j As Long
    For j = 2 To UBound(l, 1)
        If StrComp(Txt(l(j, 1)), postId, vbTextCompare) = 0 And StrComp(Txt(l(j, 3)), nm, vbTextCompare) = 0 Then
            If UBound(l, 2) >= 5 Then signedTxt = Txt(l(j, 5))
            MyPlace = CLng(Num(l(j, 2)))
            Exit Function
        End If
    Next j
End Function

' The newest sign-up file this person has in the inbox for a posting (their own
' request, read back so the sheet can say "sent"). False = none / can't read.
Private Function InboxMine(ByVal postId As String, ByVal nm As String, ByRef action As String, _
                           ByRef hrs As Double, ByRef stamp As Date) As Boolean
    action = "": hrs = 0: stamp = 0
    Dim folder As String: folder = Txt(Meta("Signup folder"))
    If folder = "" Or Not SignupsOn() Then Exit Function
    Dim files() As String, n As Long, f As String, i As Long, ff As Integer, ln As String, p As Long
    Dim emp As String, act As String, h As Double, st As Date, s As String
    On Error GoTo Quit
    f = Dir$(folder & postId & "__*.txt")
    Do While f <> ""
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    For i = 1 To n
        emp = "": act = "": h = 0: st = 0
        ff = FreeFile
        Open folder & files(i) For Input Access Read Shared As #ff
        Do While Not EOF(ff)
            Line Input #ff, ln
            p = InStr(ln, "=")
            If p > 1 Then
                s = Trim$(Mid$(ln, p + 1))
                Select Case LCase$(Trim$(Left$(ln, p - 1)))
                    Case "employee": emp = s
                    Case "action": act = LCase$(s)
                    Case "hours": If IsNumeric(s) Then h = CDbl(s)
                    Case "stamp"
                        If s Like "####-##-## ##:##:##*" Then
                            st = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2))) + _
                                 TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
                        End If
                End Select
            End If
        Loop
        Close #ff
        If StrComp(emp, nm, vbTextCompare) = 0 And st >= stamp And (act = "signup" Or act = "withdraw") Then
            action = act: hrs = h: stamp = st
            InboxMine = True
        End If
    Next i
    Exit Function
Quit:
    On Error Resume Next
    Close #ff
End Function

' What the sheet says about this person's sign-up, and whether they count as signed
' up right now. A file sent after the last publish wins over the published status.
Private Function MyStatus(ByVal postId As String, ByVal nm As String, ByVal signedTxt As String, ByRef isSigned As Boolean) As String
    Dim act As String, h As Double, st As Date
    isSigned = (signedTxt <> "")
    If InboxMine(postId, nm, act, h, st) Then
        If st > Dt(Meta("Published at")) Then
            isSigned = (act = "signup")
            If act = "signup" Then
                MyStatus = "Sent " & Format$(st, "mmm d h:nn AM/PM") & IIf(h > 0, " (" & NumTxt(h) & " h)", "")
            Else
                MyStatus = "Withdrawn " & Format$(st, "mmm d h:nn AM/PM")
            End If
            Exit Function
        End If
    End If
    If signedTxt <> "" Then MyStatus = "Signed up" & IIf(StrComp(signedTxt, "Yes", vbTextCompare) = 0, "", " (" & signedTxt & ")")
End Function

' M2: mine first (the postings the picked person is qualified for: on the list,
' or holding a code in the unit = out of department), then ONE table of every
' open posting, then the full list of the posting picked in the box at the top.
Public Sub MemberShowPostings()
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
    Dim ws As Worksheet: Set ws = WsOf(SH_POST)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    On Error GoTo Finish
    ClearBelow ws, 4
    Dim i As Long, j As Long, k As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "mb_su" Then ws.Shapes(i).Delete
    Next i
    Dim p As Variant: p = Data("dPost")
    Dim l As Variant: l = Data("dPostList")
    Dim rv As Variant: rv = Data("dRoster")
    Dim who As String: who = WhoAmI()
    Dim r As Long: r = 5
    Dim nP As Long
    If IsArray(p) Then nP = UBound(p, 1) - 1
    If nP < 1 Then
        ws.Cells(r, 1).Value = "No open postings."
        GoTo Finish
    End If
    ' soonest shift first
    Dim ord() As Long, t As Long
    ReDim ord(1 To nP)
    For i = 1 To nP
        ord(i) = i + 1
    Next i
    For i = 2 To nP
        t = ord(i): j = i - 1
        Do While j >= 1
            If Dt(p(ord(j), 3)) <= Dt(p(t, 3)) Then Exit Do
            ord(j + 1) = ord(j): j = j - 1
        Loop
        ord(j + 1) = t
    Next i

    ' ---- 1. the postings this person can sign up for ----
    Dim meRow As Long, place As Long, signedTxt As String, statusTxt As String, isSigned As Boolean
    Dim nMine As Long, codeCol As Long, placeTxt As String, isClosed As Boolean, idTxt As String
    If who <> "" Then
        meRow = RosterRowOf(rv, who)
        SectionBar ws, r, "POSTINGS " & UCase$(who) & " CAN SIGN UP FOR"
        r = r + 1
        Header ws, r, Array("Posting", "Unit", "Shift date", "Shift", "Hours", "My place", "Closes", "Sign-up", "", "")
        r = r + 1
        For i = 1 To nP
            k = ord(i)
            idTxt = Txt(p(k, 1))
            place = MyPlace(l, idTxt, who, signedTxt)
            placeTxt = ""
            If place > 0 Then
                placeTxt = "#" & place & " on the list"
            ElseIf meRow > 0 Then
                codeCol = UnitCodeCol(Txt(p(k, 2)))
                If codeCol > 0 Then
                    If codeCol <= UBound(rv, 2) Then
                        If Txt(rv(meRow, codeCol)) <> "" Then placeTxt = "out of department"
                    End If
                End If
            End If
            If placeTxt <> "" Then
                nMine = nMine + 1
                statusTxt = MyStatus(idTxt, who, signedTxt, isSigned)
                isClosed = (Now > ClosesAt(Dt(p(k, 8))))
                ws.Cells(r, 1).Value = idTxt
                ws.Cells(r, 2).Value = Txt(p(k, 2))
                ws.Cells(r, 3).Value = Dt(p(k, 3)): ws.Cells(r, 3).NumberFormat = "ddd mmm d"
                ws.Cells(r, 4).Value = Txt(p(k, 4))
                ws.Cells(r, 5).Value = Num(p(k, 5))
                ws.Cells(r, 6).Value = placeTxt
                ws.Cells(r, 7).NumberFormat = "@"              ' "Mon Jan 12 8:00 AM" would become a date-time
                ws.Cells(r, 7).Value = Format$(Dt(p(k, 8)), "ddd mmm d") & " " & CloseTime() & IIf(isClosed, " (closed)", "")
                ws.Cells(r, 8).Value = statusTxt
                ws.Rows(r).RowHeight = 19
                If SignupsOn() And Not isClosed Then
                    If isSigned Then
                        Btn ws, "mb_su" & nMine, "Withdraw", "'MemberWithdraw """ & idTxt & """'", _
                            ws.Cells(r, 9).Left + 2, ws.Cells(r, 9).Top + 2, 72, 15, LIGHT, INK
                    Else
                        Btn ws, "mb_su" & nMine, "Sign me up", "'MemberSignUp """ & idTxt & """'", _
                            ws.Cells(r, 9).Left + 2, ws.Cells(r, 9).Top + 2, 72, 15, GREEN, vbWhite
                    End If
                End If
                r = r + 1
            End If
        Next i
        If nMine = 0 Then
            ws.Cells(r, 1).Value = "None right now: no open posting lists you, and none is in a unit you hold a code in."
            ws.Cells(r, 1).Font.Italic = True
            r = r + 1
        End If
        ws.Cells(r, 1).Value = IIf(SignupsOn(), "Sign me up sends your request to the supervisors' roster. The supervisor still confirms it; the paper sheet counts too. ", _
                                   "Sign up on the posted paper sheet. ") & _
                               "'Out of department' counts only if nobody on the list takes the shift."
        ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 100, 100)
        r = r + 2
    Else
        ws.Cells(r, 1).Value = "Pick your name on Home (Find me) to see the postings you can sign up for."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 2
    End If

    ' ---- 2. every open posting, one row each ----
    SectionBar ws, r, "ALL OPEN POSTINGS"
    r = r + 1
    Dim hdrRow As Long: hdrRow = r
    Header ws, r, Array("Posting", "Unit", "Shift date", "Shift", "Hours", "Positions", "Reason", "Closes", "Listed", "Signed up")
    r = r + 1
    For i = 1 To nP
        k = ord(i)
        ws.Cells(r, 1).Value = Txt(p(k, 1))
        ws.Cells(r, 2).Value = Txt(p(k, 2))
        ws.Cells(r, 3).Value = Dt(p(k, 3)): ws.Cells(r, 3).NumberFormat = "ddd mmm d"
        ws.Cells(r, 4).Value = Txt(p(k, 4))
        ws.Cells(r, 5).Value = Num(p(k, 5))
        ws.Cells(r, 6).Value = Num(p(k, 6))
        ws.Cells(r, 7).Value = Txt(p(k, 7))
        ws.Cells(r, 8).Value = Dt(p(k, 8)): ws.Cells(r, 8).NumberFormat = "ddd mmm d"
        If UBound(p, 2) >= 10 Then
            ws.Cells(r, 9).Value = Num(p(k, 9))
            ws.Cells(r, 10).Value = Num(p(k, 10))
        End If
        r = r + 1
    Next i
    ws.Range(ws.Cells(hdrRow + 1, 4), ws.Cells(r - 1, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(hdrRow + 1, 9), ws.Cells(r - 1, 10)).HorizontalAlignment = xlCenter
    With ws.Range(ws.Cells(hdrRow, 1), ws.Cells(r - 1, 10))
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(228, 232, 238)
        .AutoFilter
    End With
    r = r + 1

    ' ---- 3. the full list of the posting picked at the top ----
    Dim pick As String: pick = Txt(ws.Range(POST_PICK).Value)
    If pick <> "" Then
        SectionBar ws, r, "LIST FOR " & UCase$(pick) & "  (call order, hours as of the posting date)"
        r = r + 1
        Header ws, r, Array("#", "Name", "Hours posted", "Signed up")
        r = r + 1
        k = 0
        If IsArray(l) Then
            For j = 2 To UBound(l, 1)
                If StrComp(Txt(l(j, 1)), pick, vbTextCompare) = 0 Then
                    If Num(l(j, 2)) > 0 Then ws.Cells(r, 1).Value = Num(l(j, 2)) Else ws.Cells(r, 1).Value = "out of dept"
                    ws.Cells(r, 2).Value = Txt(l(j, 3))
                    If Txt(l(j, 4)) <> "" Then ws.Cells(r, 3).Value = Num(l(j, 4))
                    If UBound(l, 2) >= 5 Then ws.Cells(r, 4).Value = Txt(l(j, 5))
                    ws.Cells(r, 1).HorizontalAlignment = xlCenter
                    ws.Cells(r, 3).HorizontalAlignment = xlCenter
                    r = r + 1
                    k = k + 1
                End If
            Next j
        End If
        If k = 0 Then ws.Cells(r, 2).Value = "(no list published for " & pick & ")": r = r + 1
    Else
        ws.Cells(r, 1).Value = "Pick a posting in the box at the top to see everyone listed on it."
        ws.Cells(r, 1).Font.Size = 8.5: ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 100, 100)
    End If
Finish:
    On Error Resume Next
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub

Private Sub SectionBar(ByVal ws As Worksheet, ByVal r As Long, ByVal cap As String)
    With ws.Cells(r, 1)
        .Value = cap
        .Font.Bold = True: .Font.Size = 10: .Font.Color = NAVY
    End With
    ws.Rows(r).RowHeight = 20
End Sub


' ====================================================================
'  SIGN ME UP / WITHDRAW  (M2)
'  Each click writes ONE new text file into the inbox folder the master
'  published. Nothing is edited or deleted, and this workbook is never
'  written to. The file is a request: the supervisors' roster reads it
'  when the posting is awarded, and the supervisor confirms.
' ====================================================================
Public Sub MemberSignUp(ByVal postId As String)
    SendSignup postId, "signup"
End Sub

Public Sub MemberWithdraw(ByVal postId As String)
    SendSignup postId, "withdraw"
End Sub

' letters, digits, dot, dash and underscore only (the login goes in a file name)
Private Function SafeName(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9._-]" Then SafeName = SafeName & ch
    Next i
    If SafeName = "" Then SafeName = "user"
End Function

Private Sub SendSignup(ByVal postId As String, ByVal action As String)
    Const TTL As String = "Overtime posting"
    Dim who As String: who = WhoAmI()
    If who = "" Then MsgBox "Pick your name in Find me on the Home sheet first.", vbInformation, TTL: Exit Sub
    If Not SignupsOn() Then MsgBox "Sign-up from this file isn't switched on. Sign up on the posted paper sheet.", vbInformation, TTL: Exit Sub
    Dim folder As String: folder = Txt(Meta("Signup folder"))
    Dim ok As Boolean
    On Error Resume Next
    If folder <> "" Then ok = (Len(Dir$(folder, vbDirectory)) > 0)
    On Error GoTo 0
    If Not ok Then
        MsgBox "The sign-up folder can't be reached from this PC, so nothing was sent." & vbNewLine & vbNewLine & _
               "Sign up on the posted paper sheet, and tell your supervisor.", vbExclamation, TTL
        Exit Sub
    End If
    ' the posting, as published
    Dim p As Variant: p = Data("dPost")
    Dim i As Long, k As Long
    If IsArray(p) Then
        For i = 2 To UBound(p, 1)
            If StrComp(Txt(p(i, 1)), postId, vbTextCompare) = 0 Then k = i: Exit For
        Next i
    End If
    If k = 0 Then MsgBox postId & " is no longer open in this file. See your supervisor.", vbExclamation, TTL: Exit Sub
    If Now > ClosesAt(Dt(p(k, 8))) Then
        MsgBox postId & " closed " & Format$(Dt(p(k, 8)), "ddd mmm d") & " at " & CloseTime() & ". See your supervisor.", vbExclamation, TTL
        Exit Sub
    End If
    Dim what As String
    what = postId & "   " & Txt(p(k, 2)) & "   " & Format$(Dt(p(k, 3)), "dddd mmm d, yyyy") & "   " & Txt(p(k, 4)) & "   " & NumTxt(Num(p(k, 5))) & " h"

    Dim hrs As Double, a As Variant
    If action = "signup" Then
        a = Application.InputBox(what & vbNewLine & vbNewLine & "Signing up: " & who & vbNewLine & vbNewLine & _
                                 "For the WHOLE shift, leave this blank and click OK." & vbNewLine & _
                                 "For PART of it, enter the hours you can work.", TTL, Type:=2)
        If VarType(a) = vbBoolean Then Exit Sub                    ' Cancel
        If Trim$(CStr(a)) <> "" Then
            If Not IsNumeric(Trim$(CStr(a))) Then MsgBox "Enter a number of hours, or leave it blank for the whole shift.", vbExclamation, TTL: Exit Sub
            hrs = CDbl(Trim$(CStr(a)))
            If hrs <= 0 Then MsgBox "The hours must be more than 0.", vbExclamation, TTL: Exit Sub
            If hrs >= Num(p(k, 5)) Then hrs = 0                    ' the whole shift
        End If
        If MsgBox("Sign up " & who & " for" & vbNewLine & what & vbNewLine & _
                  IIf(hrs > 0, "for " & NumTxt(hrs) & " of the " & NumTxt(Num(p(k, 5))) & " hours?", "for the whole shift?") & vbNewLine & vbNewLine & _
                  "Your request goes to the supervisors' roster with this PC's login (" & Environ$("USERNAME") & ").", _
                  vbOKCancel + vbQuestion, TTL) <> vbOK Then Exit Sub
    Else
        If MsgBox("Withdraw " & who & "'s sign-up for" & vbNewLine & what & "?", vbOKCancel + vbQuestion, TTL) <> vbOK Then Exit Sub
    End If

    ' one new file; never an existing one
    Dim login As String: login = SafeName(Environ$("USERNAME"))
    Dim stamp As Date: stamp = Now
    Dim path As String, n As Long, ff As Integer, em As String
    path = folder & UCase$(postId) & "__" & login & "__" & Format$(stamp, "yyyymmdd-hhnnss") & ".txt"
    Do While Len(Dir$(path)) > 0 And n < 20
        n = n + 1
        path = folder & UCase$(postId) & "__" & login & "__" & Format$(stamp, "yyyymmdd-hhnnss") & "-" & n & ".txt"
    Loop
    On Error GoTo Failed
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "posting=" & UCase$(postId)
    Print #ff, "employee=" & who
    Print #ff, "hours=" & IIf(hrs > 0, Replace(CStr(hrs), ",", "."), "")
    Print #ff, "action=" & action
    Print #ff, "login=" & Environ$("USERNAME")
    Print #ff, "machine=" & Environ$("COMPUTERNAME")
    Print #ff, "stamp=" & Format$(stamp, "yyyy-mm-dd hh:nn:ss")
    Close #ff
    On Error GoTo 0
    MemberShowPostings
    DrawHomeCard
    MsgBox IIf(action = "signup", "Sign-up sent.", "Withdrawal sent.") & vbNewLine & vbNewLine & _
           who & vbNewLine & what & IIf(action = "signup" And hrs > 0, vbNewLine & NumTxt(hrs) & " hours", "") & vbNewLine & _
           Format$(stamp, "ddd mmm d, h:nn AM/PM") & vbNewLine & vbNewLine & _
           "The supervisor confirms sign-ups when the posting is awarded.", vbInformation, TTL
    Exit Sub
Failed:
    em = Err.Description
    On Error Resume Next
    Close #ff
    MsgBox "The sign-up could NOT be sent (" & em & ")." & vbNewLine & vbNewLine & _
           "Sign up on the posted paper sheet, and tell your supervisor.", vbExclamation, TTL
End Sub


' ====================================================================
'  WHERE I STAND  (M2: the card on Home for the picked name)
'  Everything here is a published value looked up for one person.
' ====================================================================
Private Sub CardLine(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal txtVal As String, _
                     Optional ByVal boldVal As Boolean = False)
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 9: ws.Cells(r, 2).Font.Color = RGB(100, 100, 100)
    With ws.Cells(r, 3)
        .NumberFormat = "@"
        .Value = txtVal
        .Font.Bold = boldVal
    End With
    r = r + 1
End Sub

Private Sub DrawHomeCard()
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    Dim who As String: who = WhoAmI()
    UnprotectWs ws
    ClearBelow ws, CARD_ROW
    Dim r As Long: r = CARD_ROW
    If who = "" Or Not HasData() Then GoTo Finish
    Dim rv As Variant: rv = Data("dRoster")
    Dim meRow As Long: meRow = RosterRowOf(rv, who)
    If meRow = 0 Then
        ws.Cells(r, 2).Value = who & " isn't in the published roster. Ask your supervisor."
        GoTo Finish
    End If
    With ws.Cells(r, 2)
        .Value = "WHERE I STAND  -  " & who
        .Font.Bold = True: .Font.Size = 13: .Font.Color = NAVY
    End With
    r = r + 1
    ws.Cells(r, 2).Value = "As published " & Format$(Dt(Meta("Published at")), "ddd mmm d, h:mm AM/PM") & _
                           ", Kronos hours through " & Format$(Dt(Meta("Kronos data through")), "mmm d") & "."
    ws.Cells(r, 2).Font.Size = 8.5: ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(100, 100, 100)
    r = r + 2

    CardLine ws, r, "Crew / home unit", Txt(rv(meRow, RC_CREW)) & "   |   " & IIf(Txt(rv(meRow, RC_HOME)) = "", "(no home unit on file)", Txt(rv(meRow, RC_HOME)))
    CardLine ws, r, "Total hours", CStr(HoursText(rv(meRow, RC_TOTAL), rv(meRow, RC_PROV))) & _
             IIf(Num(rv(meRow, RC_PROV)) > 0, "   (* includes " & NumTxt(Num(rv(meRow, RC_PROV))) & " h accepted, not yet in Kronos)", ""), True
    CardLine ws, r, "Worked", NumTxt(Num(rv(meRow, RC_WORKED)))
    CardLine ws, r, "Refused", NumTxt(Num(rv(meRow, RC_REFUSED))) & IIf(RefusedKinds(rv, meRow) <> "", "   =   " & RefusedKinds(rv, meRow), "")

    ' call position in every unit they hold a code in
    Dim units As Variant: units = UnitList()
    Dim i As Long, c As Long, posTxt As String
    For i = LBound(units) To UBound(units)
        c = RC_FIXED + 2 * (i - LBound(units)) + 1
        If c + 1 <= UBound(rv, 2) Then
            If Txt(rv(meRow, c)) <> "" Then
                posTxt = posTxt & IIf(posTxt = "", "", "     ") & Txt(units(i)) & " #" & Txt(rv(meRow, c + 1)) & " (" & Txt(rv(meRow, c)) & ")"
            End If
        End If
    Next i
    CardLine ws, r, "Call position", IIf(posTxt = "", "(no unit code on file)", posTxt)
    r = r + 1

    ' sick time: counts only (decided Oct 12, 2026) - never a date of a sick day
    Dim av As Variant: av = Data("dAttend")
    Dim j As Long, found As Boolean
    If IsArray(av) Then
        For j = 2 To UBound(av, 1)
            If StrComp(Txt(av(j, 1)), who, vbTextCompare) = 0 And UBound(av, 2) >= 6 Then
                CardLine ws, r, "Sick time", "sick days used " & NumTxt(Num(av(j, 2))) & "   |   paid sick days left " & NumTxt(Num(av(j, 3))) & _
                         "   |   occurrences used " & NumTxt(Num(av(j, 4))) & "   |   personal days used " & NumTxt(Num(av(j, 5)))
                CardLine ws, r, "", Txt(av(j, 6)), True
                found = True
                Exit For
            End If
        Next j
    End If
    If Not found Then CardLine ws, r, "Sick time", "not published - ask your supervisor"
    ws.Cells(r, 3).Value = "Counts since Jan 1, or since your counters restarted. Ask your supervisor for the days behind them."
    ws.Cells(r, 3).Font.Size = 8: ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(100, 100, 100)
    r = r + 2

    ' open postings they can sign up for (the buttons are on the Postings sheet)
    Dim p As Variant: p = Data("dPost")
    Dim l As Variant: l = Data("dPostList")
    Dim place As Long, signedTxt As String, isSigned As Boolean, statusTxt As String, codeCol As Long, placeTxt As String, nMine As Long
    ws.Cells(r, 2).Value = "Open postings I can sign up for"
    ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Color = NAVY
    r = r + 1
    If IsArray(p) Then
        For i = 2 To UBound(p, 1)
            place = MyPlace(l, Txt(p(i, 1)), who, signedTxt)
            placeTxt = ""
            If place > 0 Then
                placeTxt = "#" & place & " on the list"
            Else
                codeCol = UnitCodeCol(Txt(p(i, 2)))
                If codeCol > 0 And codeCol <= UBound(rv, 2) Then
                    If Txt(rv(meRow, codeCol)) <> "" Then placeTxt = "out of department"
                End If
            End If
            If placeTxt <> "" Then
                nMine = nMine + 1
                statusTxt = MyStatus(Txt(p(i, 1)), who, signedTxt, isSigned)
                ws.Cells(r, 2).Value = Txt(p(i, 1)) & "  " & Txt(p(i, 2))
                ws.Cells(r, 3).NumberFormat = "@"
                ws.Cells(r, 3).Value = Format$(Dt(p(i, 3)), "ddd mmm d") & "  " & Txt(p(i, 4)) & "  " & NumTxt(Num(p(i, 5))) & " h   |   " & placeTxt & _
                                       "   |   closes " & Format$(Dt(p(i, 8)), "ddd mmm d") & IIf(statusTxt <> "", "   |   " & statusTxt, "")
                r = r + 1
            End If
        Next i
    End If
    If nMine = 0 Then
        ws.Cells(r, 3).Value = "none right now"
        r = r + 1
    Else
        ws.Cells(r, 3).Value = IIf(SignupsOn(), "Open Postings has the Sign me up buttons.", "Sign up on the posted paper sheet.")
        ws.Cells(r, 3).Font.Size = 8: ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(100, 100, 100)
        r = r + 1
    End If
    r = r + 1

    ' the next days where they are among the first three free in their home unit
    Dim nv As Variant: nv = Data("dNext")
    Dim homeU As String: homeU = Txt(rv(meRow, RC_HOME))
    Dim nextTxt As String, nNext As Long
    If IsArray(nv) And homeU <> "" Then
        For i = 2 To UBound(nv, 1)
            If Num(nv(i, 4)) <= 3 Then
                If StrComp(Txt(nv(i, 5)), who, vbTextCompare) = 0 And StrComp(Txt(nv(i, 1)), homeU, vbTextCompare) = 0 Then
                    nNext = nNext + 1
                    If nNext <= 8 Then nextTxt = nextTxt & IIf(nextTxt = "", "", "     ") & Format$(Dt(nv(i, 2)), "ddd mmm d") & " " & Txt(nv(i, 3)) & " #" & NumTxt(Num(nv(i, 4)))
                End If
            End If
        Next i
    End If
    ws.Cells(r, 2).Value = "Near the top of the call list"
    ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Color = NAVY
    r = r + 1
    ws.Cells(r, 3).Value = IIf(nextTxt = "", "Not in the first three for " & IIf(homeU = "", "any unit", homeU) & " in the published days.", _
                               homeU & ":  " & nextTxt & IIf(nNext > 8, "  ... and " & (nNext - 8) & " more", ""))
    r = r + 1
    ws.Cells(r, 3).Value = "Who is free changes with time off and switches: this is as of the last publish."
    ws.Cells(r, 3).Font.Size = 8: ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(100, 100, 100)
Finish:
    ProtectWs ws
    MarkClean
End Sub


' ====================================================================
'  MY HISTORY
' ====================================================================
Public Sub MemberShowHistory()
    DrawHistory Txt(WsOf(SH_HIST).Range(HIST_NAME).Value)
End Sub

Private Sub DrawHistory(ByVal nm As String)
    Dim ws As Worksheet: Set ws = WsOf(SH_HIST)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    On Error GoTo Fail
    ClearBelow ws, 4
    If nm = "" Then
        ws.Cells(5, 1).Value = "Pick a name above."
        ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    End If
    Dim v As Variant: v = Data("dLog")
    Dim hasNote As Boolean
    If IsArray(v) Then hasNote = (UBound(v, 2) >= 11)      ' Note is published only when the master says so
    Dim nc As Long: nc = IIf(hasNote, 9, 8)
    If hasNote Then
        Header ws, 5, Array("Shift date", "Ref", "Type", "Unit / shift", "Hours", "Outcome", "Charged", "Pay period", "Note")
        ws.Columns("I").ColumnWidth = 36
    Else
        Header ws, 5, Array("Shift date", "Ref", "Type", "Unit / shift", "Hours", "Outcome", "Charged", "Pay period")
    End If
    Dim perDays As Long: perDays = CLng(Num(Meta("Pay period days")))
    If perDays <= 0 Then perDays = 14
    Dim r As Long: r = 6
    Dim i As Long, idx() As Long, n As Long
    If IsArray(v) Then
        ReDim idx(1 To UBound(v, 1))
        For i = 2 To UBound(v, 1)
            If StrComp(Txt(v(i, 7)), nm, vbTextCompare) = 0 Then n = n + 1: idx(n) = i
        Next i
    End If
    ws.Range("A3").Value = nm & ": " & n & " entries this year, newest first."
    If n = 0 Then
        ws.Cells(6, 1).Value = "Nothing logged for " & nm & " this year."
        ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    End If
    ' newest first (by shift date)
    Dim j As Long, t As Long
    For i = 2 To n
        t = idx(i): j = i - 1
        Do While j >= 1
            If Dt(v(idx(j), 1)) >= Dt(v(t, 1)) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    Dim out() As Variant: ReDim out(1 To n, 1 To nc)
    Dim totals As Object: Set totals = CreateObject("Scripting.Dictionary")
    Dim yearTot As Double, ps As String, kCall As Double, kPost As Double, kOther As Double
    For i = 1 To n
        out(i, 1) = Dt(v(idx(i), 1))
        out(i, 2) = Txt(v(idx(i), 2))
        out(i, 3) = Txt(v(idx(i), 3))
        out(i, 4) = Txt(v(idx(i), 4)) & " / " & Txt(v(idx(i), 5))
        out(i, 5) = Num(v(idx(i), 6))
        out(i, 6) = Txt(v(idx(i), 8))
        out(i, 7) = Num(v(idx(i), 9))
        out(i, 8) = Format$(Dt(v(idx(i), 10)), "mmm d")
        If hasNote Then out(i, 9) = Txt(v(idx(i), 11))
        ps = Format$(Dt(v(idx(i), 10)), "yyyy-mm-dd")
        totals(ps) = Num(totals(ps)) + Num(v(idx(i), 9))
        yearTot = yearTot + Num(v(idx(i), 9))
        Select Case LCase$(Txt(v(idx(i), 3)))              ' M2: Refused by kind
            Case "callout": kCall = kCall + Num(v(idx(i), 9))
            Case "posting": kPost = kPost + Num(v(idx(i), 9))
            Case Else: kOther = kOther + Num(v(idx(i), 9))
        End Select
    Next i
    With ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, nc))
        .Value = out
        .RowHeight = 17
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(228, 232, 238)
    End With
    ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).NumberFormat = "ddd mmm d"
    ws.Range(ws.Cells(6, 5), ws.Cells(5 + n, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 7), ws.Cells(5 + n, 7)).HorizontalAlignment = xlCenter
    r = 5 + n + 2
    Header ws, r, Array("Pay period", "Charged hours", "", "", "", "", "", "")
    r = r + 1
    Dim keys As Variant: keys = totals.Keys
    Dim ks() As String, nk As Long
    nk = totals.Count
    If nk > 0 Then
        ReDim ks(1 To nk)
        For i = 0 To nk - 1
            ks(i + 1) = CStr(keys(i))
        Next i
        SortStrings ks, nk
        For i = nk To 1 Step -1
            ws.Cells(r, 1).Value = Format$(CDate(ks(i)), "mmm d") & " - " & Format$(CDate(ks(i)) + perDays - 1, "mmm d")
            ws.Cells(r, 2).Value = Num(totals(ks(i)))
            ws.Cells(r, 2).HorizontalAlignment = xlCenter
            r = r + 1
        Next i
    End If
    ws.Cells(r, 1).Value = "Year to date"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = yearTot
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).HorizontalAlignment = xlCenter
    r = r + 1
    ws.Cells(r, 1).Value = "  call-out refusals": ws.Cells(r, 2).Value = kCall: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    r = r + 1
    ws.Cells(r, 1).Value = "  posting refusals": ws.Cells(r, 2).Value = kPost: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    If Abs(kOther) > 0.0001 Then
        r = r + 1
        ws.Cells(r, 1).Value = "  other (adjustments)": ws.Cells(r, 2).Value = kOther: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    End If
Fail:                                            ' reached normally, or from an error anywhere above
    On Error Resume Next
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  NAVIGATION + EVENTS (from the members' ThisWorkbook)
' ====================================================================
Public Sub MemberGoRoster(): WsOf(SH_ROSTER).Activate: End Sub
Public Sub MemberGoNext(): WsOf(SH_NEXT).Activate: End Sub
Public Sub MemberGoCalendar(): WsOf(SH_CAL).Activate: End Sub
Public Sub MemberGoPostings(): WsOf(SH_POST).Activate: End Sub
Public Sub MemberGoHistory(): WsOf(SH_HIST).Activate: End Sub

Public Sub MemberOnOpen()
    On Error Resume Next
    MemberRefreshAll
    WsOf(SH_HOME).Activate
    ActiveWindow.DisplayGridlines = False
    MarkClean
End Sub

' a picker cell changed -> redraw that view
Public Sub MemberSheetChanged(ByVal Sh As Object, ByVal Target As Range)
    If mDrawing Then Exit Sub
    On Error Resume Next
    Select Case Sh.Name
        Case SH_ROSTER
            If Not Application.Intersect(Target, Sh.Range(ROSTER_CREW)) Is Nothing Then MemberShowAll
            If Not Application.Intersect(Target, Sh.Range(ROSTER_FIND)) Is Nothing Then ApplyMe Txt(Sh.Range(ROSTER_FIND).Value), SH_ROSTER
        Case SH_NEXT
            If Not Application.Intersect(Target, Sh.Range(NEXT_UNIT & "," & NEXT_DATE & "," & NEXT_SHIFT)) Is Nothing Then MemberShowNext
        Case SH_CAL
            If Not Application.Intersect(Target, Sh.Range(CAL_MONTH & "," & CAL_UNIT)) Is Nothing Then MemberShowCalendar
        Case SH_HIST
            If Not Application.Intersect(Target, Sh.Range(HIST_NAME)) Is Nothing Then MemberShowHistory
        Case SH_POST
            If Not Application.Intersect(Target, Sh.Range(POST_PICK)) Is Nothing Then MemberShowPostings
        Case SH_HOME                                   ' M2: Find me drives every sheet and the card below it
            If Not Application.Intersect(Target, Sh.Range(HOME_FIND)) Is Nothing Then ApplyMe Txt(Sh.Range(HOME_FIND).Value), SH_HOME
            MarkClean
    End Select
End Sub

Public Sub MemberSheetActivated(ByVal Sh As Object)
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    MarkClean
End Sub
