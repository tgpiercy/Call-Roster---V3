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

Public Const MEM_VERSION As String = "Members M6.1 (Oct 28)"     ' v29.0.3 (M6.1, the alpha only): Act as for the admins' logins only (the template is rebuilt by the panel); v29 (M6): the status word "Booking voided", My history's PDF to the user's Downloads; v28 (M5, item 1): members' requests - ask for time off or a switch, the partner's Agree, My requests (the template is rebuilt); v27 (M4): me first, the tiles, since you last looked, the header band, How it works
Private Const PW As String = "members"            ' same as MEMBER_PASSWORD in the master
' v29.0.2 (the alpha only; removed in the release after): Act as - MemberLogin() stands in for the PC's login in every file
' name and login= line while a chair is stored in this file (Alt+F8 > MemberActAs with MEM_CHAIR_PASSWORD; MemberActAsClear);
' ignored after MEM_ALPHA_END_ISO; a freshly published file holds no chair
Private Const MEM_CHAIR_PASSWORD As String = "alpha2026"   ' the same word as CHAIR_PASSWORD in the master
Private Const MEM_ALPHA_END_ISO As String = "2026-11-20"
Private Const ACTAS_NAME As String = "ActAsChair"
' v29.0.3 (review 1.5): the Windows logins that may set or use a chair in this file - the roster's admins. Any other login
' is itself whatever name the file holds (a member could add the name in Name Manager, or read the password in the code)
Private Const MEM_ACTAS_USERS As String = "piercyg,lanthierr"

Private Const SH_HOME As String = "Home"
Private Const SH_ROSTER As String = "Roster"
Private Const SH_NEXT As String = "Who's next"
Private Const SH_CAL As String = "Calendar"
Private Const SH_POST As String = "Postings"
Private Const SH_HIST As String = "My history"
Private Const SH_LISTS As String = "dLists"       ' picker lists, rebuilt by MemberRefreshAll
Private Const SH_HELP As String = "How it works"  ' M4: the rules in plain words, inside the file
Private Const SH_REQ As String = "Requests"       ' M5 (v28 item 1): ask for time off or a switch, the partner's Agree, My requests
Private Const APP_KEY As String = "CallOutRoster"  ' M4: Excel's per-user setting store (this PC user), never the shared file
Private Const APP_SEC As String = "Members"
Private Const LAYOUT_CELL As String = "Z1"        ' M4: where BuildMemberTemplate writes the version (Home; was B19)

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
Private Const HOME_FIND As String = "C6"          ' M4: me first - the name sits at the top
Private Const POST_PICK As String = "C2"           ' Postings: "show the list for"
Private Const REQ_KIND As String = "C4"            ' M5: the request's typed cells (their labels sit one column left)
Private Const REQ_FROM As String = "E4"
Private Const REQ_TO As String = "G4"
Private Const REQ_SHIFT As String = "C5"
Private Const REQ_PARTNER As String = "E5"
Private Const REQ_NOTE As String = "C6"
Private Const REQ_FIRST_ROW As Long = 9            ' M5: the lists start here (rows 4-7 are the ask; row 3 the subtitle)
Private Const CARD_ROW As Long = 21                ' Home: the rest of the card starts here, under the tiles (rows 7-12) and the buttons (14-20)
Private Const CAL_COLW As Double = 17              ' calendar: width of each half of a day
Private Const CAL_MAX_LINES As Long = 6            ' calendar: lines shown per day half; the rest become "... and N more"
Private Const US As String = "~|~"                 ' separators for the calendar's per-day line lists
Private Const RS As String = "~#~"

Private mDrawing As Boolean
Private mViewErr As String                         ' M2.1: a view that failed to draw says so here (its own handler ends quietly)
Private mLoadNote As String                        ' M2.1: the note MemberRefreshAll left on Home ("" = none)
Private mSeen As Object                            ' M4: "since you last looked", worked out once a session per name (name -> line)


' M5 (v28 item 1): the Requests page's constants (declared here, above the first procedure)
Private Const RQ_TTL As String = "Members' requests"
' dRequests, as the master publishes it (the contract test ties these to modRequests.PublishRequestsTable)
Private Const DQ_EMP As Long = 1
Private Const DQ_KIND As Long = 2
Private Const DQ_FROM As Long = 3
Private Const DQ_TO As Long = 4
Private Const DQ_SHIFT As Long = 5
Private Const DQ_PARTNER As Long = 6
Private Const DQ_STATUS As Long = 7
Private Const DQ_ANSWER As Long = 8
Private Const DQ_DECIDED As Long = 9
' a request file read back (ReadRequestFolder rows)
Private Const FQ_ID As Long = 1
Private Const FQ_ACTION As Long = 2
Private Const FQ_KIND As Long = 3
Private Const FQ_EMP As Long = 4
Private Const FQ_FROM As Long = 5
Private Const FQ_TO As Long = 6
Private Const FQ_SHIFT As Long = 7
Private Const FQ_PARTNER As Long = 8
Private Const FQ_NOTE As Long = 9
Private Const FQ_LOGIN As Long = 10
Private Const FQ_STAMP As Long = 11
Private Const FQ_N As Long = 11


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

' The version whose BuildMemberTemplate laid the sheets out (Home Z1 since M4 - B19 before: only the build writes it)
Public Function MemberLayoutVersion() As String
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    If Not ws Is Nothing Then MemberLayoutVersion = Trim$(Txt(ws.Range(LAYOUT_CELL).Value))
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
    Set ws = WsOf(SH_REQ)                            ' M5: the lists under the ask
    If Not ws Is Nothing Then
        UnprotectWs ws
        ClearBelow ws, REQ_FIRST_ROW - 1
        ws.Cells(REQ_FIRST_ROW, 1).Value = "Not available - see the note on the Home sheet."
        ProtectWs ws
    End If
    Set ws = WsOf(SH_HOME)
    If Not ws Is Nothing Then
        UnprotectWs ws
        ClearTiles ws                                ' M4: the tiles too
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
    ws.Rows(1).RowHeight = 40                        ' M4 (review M2): the header band's height, before any row-2 button is placed
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
    For Each nm In Array("dMeta", "dRoster", "dNext", "dCal", "dPost", "dPostList", "dLog", "dAttend", "dWorked", "dNotFree", "dRequests", SH_LISTS)   ' M3.1: + dWorked; M4: + dNotFree; M5: + dRequests
        Set ws = WsOf(CStr(nm), True)
        UnprotectWs ws
        ws.Visible = xlSheetVeryHidden
        ProtectWs ws
    Next nm
    ' view sheets, in order
    For Each nm In ViewSheets()
        Set ws = WsOf(CStr(nm), True)
        ws.Visible = xlSheetVisible
    Next nm
    Dim i As Long, order As Variant
    order = ViewSheets()
    For i = 0 To UBound(order)
        ThisWorkbook.Worksheets(CStr(order(i))).Move Before:=ThisWorkbook.Worksheets(i + 1)
    Next i
    ' the blank workbook's own sheet(s) - by index, backwards, so none is skipped
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set sh = ThisWorkbook.Worksheets(i)
        If sh.Visible = xlSheetVisible Then
            If InStr(1, "|" & Join(ViewSheets(), "|") & "|", "|" & sh.Name & "|") = 0 Then sh.Delete
        End If
    Next i
    Application.DisplayAlerts = True

    LayoutHome
    LayoutRoster
    LayoutNext
    LayoutCalendar
    LayoutPostings
    LayoutHistory
    LayoutRequests                                   ' M5
    LayoutHelp                                       ' M4
    For Each nm In ViewSheets()                      ' M4: one header band on every sheet but Home, the version in every footer
        Set ws = WsOf(CStr(nm))
        UnprotectWs ws
        If CStr(nm) <> SH_HOME Then HeaderBand ws
        On Error Resume Next
        ws.PageSetup.RightFooter = "&7" & MEM_VERSION
        On Error GoTo 0
        ProtectWs ws
    Next nm
    MemberRefreshAll
    For Each nm In ViewSheets()      ' M3: the template ships at 100 % (Larger type is per session)
        ThisWorkbook.Worksheets(CStr(nm)).Activate
        ActiveWindow.Zoom = 100
    Next nm
    ThisWorkbook.Worksheets(SH_HOME).Activate
    Application.ScreenUpdating = True
    MsgBox "Template built (" & MEM_VERSION & "). Save it as 'Call Out Roster - Members TEMPLATE.xlsm' next to the master. " & _
           "The master's Member View button fills it and publishes the members' copy.", vbInformation, "Members' template"
End Sub

' M4 (item 34): me first. Rows 2-4 the title and the published stamp, 5 the warning line, 6 Find me (remembered on
' this PC), 7 the name, 8 "since you last looked", 9-12 the four tiles, 14-20 the six buttons, 21 on the rest of the
' card. Row heights are set BEFORE the buttons are placed: shapes don't move with cells.
Private Sub LayoutHome()
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    UnprotectWs ws
    Dim sk As Long                                   ' (review H2) a rebuild in place: every old button goes (M3.1's mb_zoom sat over a tile)
    For sk = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(sk).Name, 3) = "mb_" Then ws.Shapes(sk).Delete
    Next sk
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:I").ColumnWidth = 18
    ws.Rows(1).RowHeight = 6
    ws.Rows(2).RowHeight = 36
    With ws.Range("B2")
        .Value = "OVERTIME CALL OUT ROSTER"
        .Font.Size = 20: .Font.Bold = True: .Font.Color = NAVY
    End With
    ws.Range("B3").Font.Size = 9: ws.Range("B3").Font.Color = RGB(100, 100, 100)
    ws.Range("B4").Font.Size = 8: ws.Range("B4").Font.Color = RGB(130, 135, 145)
    ws.Range("B5:I5").Merge                          ' M2.1: the warning line (HomeNote)
    ws.Range("B5").WrapText = True
    ws.Range("B5").VerticalAlignment = xlTop
    ws.Rows(5).RowHeight = 26
    ws.Rows(6).RowHeight = 22
    ws.Rows(7).RowHeight = 22
    ws.Rows(8).RowHeight = 30
    ws.Rows(9).RowHeight = 15
    ws.Rows(10).RowHeight = 26
    ws.Rows(11).RowHeight = 34
    ws.Rows(12).RowHeight = 20
    ws.Rows(13).RowHeight = 10
    Dim r As Long
    For r = 14 To 20
        ws.Rows(r).RowHeight = 15
    Next r
    ' the band on Home: Larger type and How it works beside the title
    Btn ws, "mb_band_zoom", "Larger type", "MemberZoomToggle", ws.Range("G2").Left, ws.Range("B2").Top + 8, 90, 20, LIGHT, INK
    Btn ws, "mb_band_help", "How it works", "MemberGoHelp", ws.Range("G2").Left + 96, ws.Range("B2").Top + 8, 100, 20, LIGHT, INK
    ' Find me, remembered on this PC (Excel's own per-user setting store), Not me
    Picker ws, HOME_FIND, "", "Find me:"
    ws.Range(HOME_FIND).ColumnWidth = 26
    ' 26 points clear of the box: Excel draws the drop-down arrow just outside its right edge
    Btn ws, "mb_notme", "Not me", "MemberNotMe", ws.Range("D6").Left + 26, ws.Range("D6").Top + 1, 80, 20, vbWhite, RGB(100, 100, 100)
    Btn ws, "mb_find", "Show me on the roster", "MemberFindFromHome", ws.Range("D6").Left + 112, ws.Range("D6").Top + 1, 150, 20, LIGHT, INK
    With ws.Range("F6")
        .Value = "Remembered on this PC, never in the shared file."
        .Font.Size = 8: .Font.Italic = True: .Font.Color = RGB(130, 135, 145)
    End With
    ' the seven buttons below the tiles (M5: four and three; MY REQUESTS opens the Requests page)
    Dim top As Double: top = ws.Range("B14").Top
    Dim lft As Double: lft = ws.Range("B14").Left
    Btn ws, "mb_roster", "ROSTER", "MemberGoRoster", lft, top, 115, 44, NAVY, vbWhite
    Btn ws, "mb_next", "WHO'S NEXT", "MemberGoNext", lft + 125, top, 115, 44, GREEN, vbWhite
    Btn ws, "mb_cal", "CALENDAR", "MemberGoCalendar", lft + 250, top, 115, 44, RGB(31, 56, 100), vbWhite
    Btn ws, "mb_post", "OPEN POSTINGS", "MemberGoPostings", lft + 375, top, 115, 44, RGB(93, 64, 55), vbWhite
    Btn ws, "mb_hist", "MY HISTORY", "MemberGoHistory", lft, top + 54, 115, 44, RGB(106, 27, 154), vbWhite
    Btn ws, "mb_req", "MY REQUESTS", "MemberGoRequests", lft + 125, top + 54, 115, 44, RGB(0, 110, 90), vbWhite
    Btn ws, "mb_helpbig", "HOW IT WORKS", "MemberGoHelp", lft + 250, top + 54, 115, 44, LIGHT, INK
    With ws.Range(LAYOUT_CELL)                       ' the layout version: only the build writes it (MemberLayoutVersion)
        .Value = MEM_VERSION
        .Font.Size = 6: .Font.Color = vbWhite
    End With
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
    ws.Columns("I").ColumnWidth = 11
    ws.Columns("J").ColumnWidth = 62                     ' M3: the Why line
    Picker ws, HIST_NAME, "", "Name:"
    ' M3: the same rows as the master's member-safe history report, printable from here too
    Btn ws, "mb_hist_pdf", "Print / PDF", "MemberHistoryPrint", ws.Range("E2").Left, ws.Range("E2").Top - 1, 90, 20, NAVY, vbWhite
    ProtectWs ws
End Sub

' M6 (v29): the signed-in user's Downloads folder - Windows' own answer first (a relocated Downloads folder included), then
' the profile's Downloads, then the user's TEMP (never the members' folder: read-only to members once the files split)
Private Function DownloadsFolder() As String
    On Error Resume Next                            ' never raises
    Dim sh As Object, ns As Object, p As String
    Set sh = CreateObject("Shell.Application")
    If Not sh Is Nothing Then
        Set ns = sh.Namespace("shell:Downloads")
        If Not ns Is Nothing Then p = ns.Self.Path
    End If
    If p = "" Or Dir$(p, vbDirectory) = "" Then p = Environ$("USERPROFILE") & Application.PathSeparator & "Downloads"
    If Dir$(p, vbDirectory) = "" Then p = Environ$("TEMP")
    DownloadsFolder = p
End Function

' M3: My history as a PDF ("History <name> <date>.pdf"), then "print as well?". M6 (v29; Graeme, Oct 27): the PDF goes to the
' signed-in user's Downloads folder (DownloadsFolder), never beside the members' file.
Public Sub MemberHistoryPrint()
    Dim ws As Worksheet: Set ws = WsOf(SH_HIST)
    Dim nm As String: nm = Txt(ws.Range(HIST_NAME).Value)
    If nm = "" Then MsgBox "Pick a name above first.", vbInformation, "My history": Exit Sub
    Dim pdfPath As String, em As String, folder As String
    folder = DownloadsFolder()
    pdfPath = folder & Application.PathSeparator & "History " & SafeName(Replace(nm, " ", "_")) & " " & Format$(Date, "yyyy-mm-dd") & ".pdf"
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Err.Number <> 0 Then em = Err.Description
    Err.Clear
    Dim doPrint As Boolean
    If em <> "" Then
        doPrint = (MsgBox("The PDF couldn't be saved in " & folder & " (" & em & "). Print it instead?", vbYesNo + vbExclamation, "My history") = vbYes)
    Else
        doPrint = (MsgBox("PDF saved in your " & IIf(InStr(1, folder, "Downloads", vbTextCompare) > 0, "Downloads folder", "folder " & folder) & ":" & vbNewLine & pdfPath & vbNewLine & vbNewLine & _
                          "Print My history as well?", vbYesNo + vbInformation, "My history") = vbYes)
    End If
    If doPrint Then
        Err.Clear
        ws.PrintOut Copies:=1
        If Err.Number <> 0 Then MsgBox "It couldn't be printed: " & Err.Description, vbExclamation, "My history"
    End If
    On Error GoTo 0
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
    mViewErr = ""
    On Error GoTo Done
    stepTxt = "the layout check"
    note = LayoutProblem()                         ' sheets an older version laid out: write nothing into their cells
    If note = "" Then
        ' the file always starts with nobody picked, even if it was saved with a name showing
        stepTxt = "the pickers"
        SetCell WsOf(SH_HOME), HOME_FIND, ""
        SetCell WsOf(SH_ROSTER), ROSTER_FIND, ""
        SetCell WsOf(SH_HIST), HIST_NAME, ""
        SetCell WsOf(SH_POST), POST_PICK, ""
        stepTxt = "the lists"
        BuildLists
    End If
    stepTxt = "Home"
    DrawHome
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
        stepTxt = "the Requests"
        MemberShowRequests                           ' M5
        stepTxt = "How it works"
        DrawHelp                                     ' M4
    End If
Done:
    en = Err.Number: em = Err.Description          ' reached normally (0) or from an error anywhere above
    On Error Resume Next
    If en <> 0 Then
        note = "This file didn't finish loading (" & stepTxt & ": " & em & "). Tell your supervisor."
    ElseIf mViewErr <> "" Then
        note = "This file didn't finish loading (" & mViewErr & "). Tell your supervisor."
    End If
    mLoadNote = note
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
    SetList WsOf(SH_REQ), REQ_KIND, "Time off,Switch"              ' M5: the request's pickers
    SetList WsOf(SH_REQ), REQ_SHIFT, "Days,Nights"
    SetList WsOf(SH_REQ), REQ_PARTNER, ListRef("A", nNames)
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
        ws.Range("B3").Value = StampText() & "   |   Kronos through " & Format$(Dt(Meta("Kronos data through")), "mmm d") & _
                               "   |   pay period " & Format$(Dt(Meta("Pay period start")), "mmm d") & " - " & _
                               Format$(Dt(Meta("Pay period end")), "mmm d")
        ws.Range("B3").Font.Color = IIf(StampOld(), RGB(176, 0, 32), RGB(100, 100, 100))
        ws.Range("B3").Font.Bold = StampOld()
        ws.Range("B4").Value = "Hours and call order come from the supervisors' roster (" & Txt(Meta("Master version")) & ") and are published here as values. " & _
                               "Leave shows only as 'Leave'. Sick time is counts only. Charged: " & Txt(Meta("Charge rule")) & "."
    Else
        ws.Range("B3").Value = "Not published yet - the supervisors' roster fills this file with its Member View button."
        ws.Range("B4").Value = ""
    End If
    ProtectWs ws
    StampBands                                       ' M4: the same stamp in every sheet's header band
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
    If nm = "" Then MsgBox "Pick your name in the Find me box first.", vbInformation, "Find me": Exit Sub
    ApplyMe nm, SH_ROSTER
End Sub

' Home: the "Show me on the roster" button
Public Sub MemberFindFromHome()
    Dim nm As String: nm = WhoAmI()
    If nm = "" Then MsgBox "Pick your name first.", vbInformation, "Find me": Exit Sub
    ApplyMe nm, SH_ROSTER
End Sub

' Home: back to the neutral views, and this PC forgets the name (M4)
Public Sub MemberNotMe()
    RememberMe ""
    ApplyMe "", SH_HOME
End Sub

' M4 (item 34): the name this PC user picked last time, from Excel's per-user setting store (never the shared file)
Private Function RememberedMe() As String
    On Error Resume Next
    RememberedMe = GetSetting(APP_KEY, APP_SEC, "Me", "")
End Function

Private Sub RememberMe(ByVal nm As String)
    On Error Resume Next
    If nm = "" Then DeleteSetting APP_KEY, APP_SEC, "Me" Else SaveSetting APP_KEY, APP_SEC, "Me", nm
End Sub

' The name picked in Find me on Home ("" = nobody). The published file starts with no name; M4: the open puts
' back the name this PC user picked last time (RememberedMe).
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
    mViewErr = ""
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
    MemberShowRequests                               ' M5 (review): the Requests page follows the picked name
    DrawHomeCard
    If nm <> "" And stayOn = SH_HOME Then RememberMe nm     ' M4: the next open starts on this name (review: Home's Find me only, not a colleague looked up on the Roster)
Done:
    en = Err.Number: em = Err.Description          ' reached normally (0) or from an error anywhere above
    On Error Resume Next
    If en = 0 And mViewErr <> "" Then en = 1: em = mViewErr
    ' a note left by the load stays: it is still true
    HomeNote IIf(en <> 0, "Couldn't show everything for " & nm & " (" & em & "). Tell your supervisor.", mLoadNote)
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
    On Error GoTo Fail
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
        heads = Array("Name", "Crew", "Home unit", "Total", "Worked", "Charged (refusals)", "Charged by kind")     ' M4: plain words
    Else
        heads = Array("#", "Name", "Crew", "Code", "Total", "Worked", "Charged (refusals)", "Charged by kind")
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
    Exit Sub
Fail:                                            ' M2.1: an error here used to leave Excel's events switched off
    mViewErr = "the Roster: " & Err.Description
    On Error Resume Next
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
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
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
    Dim i As Long, n As Long, out() As Variant, anyEst As Boolean
    If IsArray(v) And unit <> "" And dTxt <> "" And sh <> "" Then
        ReDim out(1 To UBound(v, 1), 1 To 5)
        For i = 2 To UBound(v, 1)
            If StrComp(Txt(v(i, 1)), unit, vbTextCompare) = 0 And StrComp(Txt(v(i, 3)), sh, vbTextCompare) = 0 And _
               (Format$(Dt(v(i, 2)), "ddd mmm d") = dTxt Or (dPick > 0 And Int(Dt(v(i, 2))) = dPick)) Then
                n = n + 1
                out(n, 1) = Num(v(i, 4)): out(n, 2) = Txt(v(i, 5)): out(n, 3) = "": out(n, 4) = Txt(v(i, 8))
                out(n, 5) = HoursText(v(i, 6), v(i, 7))
                If UBound(v, 2) >= 9 Then                     ' M3 (M8): a planned mover's estimated hours, with ~ as the master shows them
                    If StrComp(Txt(v(i, 9)), "Yes", vbTextCompare) = 0 Then out(n, 5) = "~" & NumTxt(Num(v(i, 6))): anyEst = True
                End If
            End If
        Next i
    End If
    ws.Range("A3").Value = IIf(n = 0, "Nobody free", n & " free") & " for " & unit & ", " & dTxt & ", " & sh & _
                           ".  Order: grade code, then total hours, then plant seniority.  (as of " & _
                           Format$(Dt(Meta("Published at")), "mmm d h:mm AM/PM") & ")" & _
                           IIf(anyEst, "   ~ = estimated hours after a planned unit move.", "")
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
        Dim meNm As String: meNm = WhoAmI()          ' M4 (item 35): highlight me
        If meNm <> "" Then
            For i = 1 To n
                If StrComp(CStr(out(i, 2)), meNm, vbTextCompare) = 0 Then
                    ws.Cells(5 + i, 2).Value = meNm & "  (you)"
                    With ws.Range(ws.Cells(5 + i, 1), ws.Cells(5 + i, 5))
                        .Interior.Color = RGB(255, 242, 204): .Font.Bold = True
                    End With
                End If
            Next i
        End If
    Else
        ws.Cells(6, 2).Value = "No one is free for that shift (or the date is outside the published " & Txt(Meta("Who's next days")) & " days)."
        ws.Cells(6, 2).Font.Italic = True
    End If
    ' M4 (item 37): the people with a code in the unit who are not listed for that shift, and the reason word
    Dim fv As Variant: fv = Data("dNotFree")
    Dim rr As Long: rr = 7 + IIf(n > 0, n, 1)
    Dim nNot As Long, meNot As String: meNot = WhoAmI()
    If IsArray(fv) And unit <> "" And dTxt <> "" And sh <> "" Then
        For i = 2 To UBound(fv, 1)
            If StrComp(Txt(fv(i, 1)), unit, vbTextCompare) = 0 And StrComp(Txt(fv(i, 3)), sh, vbTextCompare) = 0 And _
               (Format$(Dt(fv(i, 2)), "ddd mmm d") = dTxt Or (dPick > 0 And Int(Dt(fv(i, 2))) = dPick)) Then
                If nNot = 0 Then
                    ws.Cells(rr, 2).Value = "Not listed for this shift"
                    ws.Cells(rr, 2).Font.Bold = True: ws.Cells(rr, 2).Font.Color = NAVY
                    rr = rr + 1
                End If
                nNot = nNot + 1
                ws.Cells(rr, 2).Value = Txt(fv(i, 4)) & IIf(StrComp(Txt(fv(i, 4)), meNot, vbTextCompare) = 0, "  (you)", "")
                ws.Cells(rr, 4).Value = Txt(fv(i, 5))
                ws.Range(ws.Cells(rr, 2), ws.Cells(rr, 4)).Font.Color = RGB(110, 110, 110)
                If StrComp(Txt(fv(i, 4)), meNot, vbTextCompare) = 0 Then ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 5)).Interior.Color = RGB(255, 242, 204)
                rr = rr + 1
            End If
        Next i
    End If
    Err.Clear                                    ' (a normal arrival carries no error)
Fail:                                            ' reached normally, or from an error anywhere above
    If Err.Number <> 0 Then mViewErr = "Who's next: " & Err.Description
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

' The month shown, as its first day (0 = can't be read). M3 (M13): the label is matched against the
' published dates' own "mmm yyyy" labels - never CDate on a plant PC whose locale may read it differently.
Private Function CalMonthStart(ByVal ws As Worksheet) As Date
    Dim mTxt As String: mTxt = Txt(ws.Range(CAL_MONTH).Value)
    Dim m0 As Date
    If VarType(ws.Range(CAL_MONTH).Value) = vbDate Then
        m0 = ws.Range(CAL_MONTH).Value
    ElseIf mTxt <> "" Then
        Dim v As Variant: v = Data("dCal")
        Dim i As Long, d As Date
        If IsArray(v) Then
            For i = 2 To UBound(v, 1)
                d = Dt(v(i, 1))
                If d <> 0 Then
                    If StrComp(Format$(d, "mmm yyyy"), mTxt, vbTextCompare) = 0 Then m0 = d: Exit For
                End If
            Next i
        End If
        If m0 = 0 Then m0 = MonthFromLabel(mTxt)             ' a month with no published line at all
    End If
    If m0 <> 0 Then CalMonthStart = DateSerial(Year(m0), Month(m0), 1)
End Function

' "Oct 2026" -> Oct 1, 2026 by the English month names the picker itself was filled with (0 = not one)
Private Function MonthFromLabel(ByVal mTxt As String) As Date
    Dim parts As Variant: parts = Split(Trim$(mTxt), " ")
    If UBound(parts) <> 1 Then Exit Function
    Dim m As Long, k As Long
    For k = 1 To 12
        If StrComp(Left$(MonthName(k, True), 3), Left$(CStr(parts(0)), 3), vbTextCompare) = 0 Then m = k: Exit For
    Next k
    If m = 0 Or Not IsNumeric(parts(1)) Then Exit Function
    MonthFromLabel = DateSerial(CLng(parts(1)), m, 1)
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

' Writes one block of lines (a day's days or nights) into the pair of columns c, c + 1.
' maxRows (M2.1): the block's height; when there are more lines than that, the last row says how many more.
Private Sub PaintLines(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal lines As String, _
                       Optional ByVal maxRows As Long = 0)
    If lines = "" Then Exit Sub
    Dim arr As Variant: arr = Split(lines, RS)
    Dim i As Long, p As Variant, col As Long
    Dim nAll As Long: nAll = UBound(arr) + 1
    Dim nShow As Long: nShow = nAll
    If maxRows > 0 And nAll > maxRows Then nShow = maxRows - 1
    For i = 0 To nShow - 1
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
    If nShow < nAll Then
        With ws.Cells(r + nShow, c)
            .Value = "... " & (nAll - nShow) & " more"
            .Font.Italic = True
            .Font.Color = RGB(100, 100, 100)
        End With
    End If
End Sub

' M2: the master's month grid. Week rows, Sunday to Saturday; each day is two
' columns (who's off | covered by) with the day lines above the night lines.
Public Sub MemberShowCalendar()
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
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
                If mxD > CAL_MAX_LINES Then mxD = CAL_MAX_LINES       ' M2.1: a crowded day doesn't stretch the whole week
                If mxN > CAL_MAX_LINES Then mxN = CAL_MAX_LINES
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
                PaintLines ws, r, c, dTxt(k), mxD
                PaintLines ws, r + mxD, c, nTxt(k), mxN
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
    On Error GoTo Finish
    Err.Clear
Finish:
    If Err.Number <> 0 Then mViewErr = "the Calendar: " & Err.Description
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
' The person's rank on a posting's published list; 0 = not on it; -1 (M3) = on it but EXEMPT (moved unit
' or left: the master never charges them on it and they can't sign up for it here)
Private Function MyPlace(ByRef l As Variant, ByVal postId As String, ByVal nm As String, ByRef signedTxt As String) As Long
    signedTxt = ""
    If Not IsArray(l) Then Exit Function
    Dim j As Long
    For j = 2 To UBound(l, 1)
        If StrComp(Txt(l(j, 1)), postId, vbTextCompare) = 0 And StrComp(Txt(l(j, 3)), nm, vbTextCompare) = 0 Then
            If UBound(l, 2) >= 5 Then signedTxt = Txt(l(j, 5))
            If IsExemptRow(l, j) Then MyPlace = -1: Exit Function
            MyPlace = CLng(Num(l(j, 2)))
            Exit Function
        End If
    Next j
End Function

' M3: the Exempt column of a dPostList row ("Yes" = moved unit / left; no reason is published)
Private Function IsExemptRow(ByRef l As Variant, ByVal j As Long) As Boolean
    If UBound(l, 2) >= 6 Then IsExemptRow = (StrComp(Txt(l(j, 6)), "Yes", vbTextCompare) = 0)
End Function

' The newest sign-up file this person has in the inbox for a posting (their own
' request, read back so the sheet can say "sent"). False = none / can't read.
' M3.1: the inbox folder as published (".\Signups" = beside this file), resolved the one way for the sign-up
' and for reading it back (the v26.5 reviews: InboxMine used the raw ".\" against Excel's current folder)
Private Function SignupFolderPath() As String
    Dim folder As String: folder = Txt(Meta("Signup folder"))
    If folder = "." Then
        folder = ThisWorkbook.Path
    ElseIf Left$(folder, 2) = ".\" Then
        folder = ThisWorkbook.Path & Mid$(folder, 2)
    End If
    If folder <> "" Then
        If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    End If
    SignupFolderPath = folder
End Function

Private Function InboxMine(ByVal postId As String, ByVal nm As String, ByRef action As String, _
                           ByRef hrs As Double, ByRef stamp As Date) As Boolean
    action = "": hrs = 0: stamp = 0
    Dim folder As String: folder = SignupFolderPath()
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

' dPost rows in shift-date order: ord(1..nP) = row numbers of p, soonest first
Private Sub SoonestFirst(ByRef p As Variant, ByVal nP As Long, ByRef ord() As Long)
    Dim i As Long, j As Long, t As Long
    If nP < 1 Then Exit Sub
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
End Sub

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
        Err.Clear
        GoTo Finish
    End If
    ' soonest shift first
    Dim ord() As Long
    SoonestFirst p, nP, ord

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
            ElseIf place < 0 Then
                placeTxt = "not on this list now (moved unit)"     ' M3: an Exempt row - no rank, no Sign me up
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
                If SignupsOn() And Not isClosed And place >= 0 Then     ' M3: no button on an Exempt row
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
                    If IsExemptRow(l, j) Then                      ' M3: moved unit / left - on the paper list, no rank here
                        ws.Cells(r, 1).Value = "-"
                        ws.Cells(r, 2).Value = Txt(l(j, 3)) & "  (not on this list now)"
                        ws.Cells(r, 2).Font.Color = RGB(130, 130, 130)
                    Else
                        If Num(l(j, 2)) > 0 Then ws.Cells(r, 1).Value = Num(l(j, 2)) Else ws.Cells(r, 1).Value = "out of dept"
                        ws.Cells(r, 2).Value = Txt(l(j, 3))
                    End If
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
    Err.Clear                                    ' (a normal arrival carries no error)
Finish:
    If Err.Number <> 0 Then mViewErr = "the Postings: " & Err.Description
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

' M2.1: the last question before a request is sent. Yes = send it for who. No = it is
' for someone else: the name is asked for and must be on the published roster.
' askTxt carries {who} where the name goes. False = nothing is sent.
Private Function ConfirmFor(ByRef who As String, ByVal askTxt As String, ByVal ttl As String) As Boolean
    Dim ans As Long, a As Variant, rv As Variant, r As Long
    ans = MsgBox(Replace(askTxt, "{who}", who) & vbNewLine & vbNewLine & _
                 "Yes = send it for " & who & "." & vbNewLine & "No = it is for someone else.", vbYesNoCancel + vbQuestion, ttl)
    If ans = vbCancel Then Exit Function
    If ans = vbNo Then
        a = Application.InputBox("Whose is it? Type the name exactly as it is on the Roster sheet.", ttl, Type:=2)
        If VarType(a) = vbBoolean Then Exit Function
        rv = Data("dRoster")
        r = RosterRowOf(rv, Trim$(CStr(a)))
        If r = 0 Then
            MsgBox "'" & Trim$(CStr(a)) & "' isn't on the published roster. Nothing was sent.", vbExclamation, ttl
            Exit Function
        End If
        who = Txt(rv(r, RC_NAME))
        If MsgBox(Replace(askTxt, "{who}", who), vbOKCancel + vbQuestion, ttl) <> vbOK Then Exit Function
    End If
    ConfirmFor = True
End Function

Private Sub SendSignup(ByVal postId As String, ByVal action As String)
    Const TTL As String = "Overtime posting"
    Dim who As String: who = WhoAmI()
    If who = "" Then MsgBox "Pick your name in Find me on the Home sheet first.", vbInformation, TTL: Exit Sub
    If RosterRowOf(Data("dRoster"), who) = 0 Then       ' M3 (M11): a typed "me" the roster doesn't know is never sent
        MsgBox "'" & who & "' isn't on the published roster, so nothing was sent. Pick your name from the Find me list on the Home sheet.", vbExclamation, TTL
        Exit Sub
    End If
    If Not SignupsOn() Then MsgBox "Sign-up from this file isn't switched on. Sign up on the posted paper sheet.", vbInformation, TTL: Exit Sub
    ' M3 (M10): the inbox folder is published as the setting reads (".\Signups" = beside this file), resolved here
    Dim folder As String: folder = SignupFolderPath()
    If folder <> "" Then
        If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    End If
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
        a = Application.InputBox(what & vbNewLine & vbNewLine & "Signing up: " & who & "  (the next box lets you change the name)" & vbNewLine & vbNewLine & _
                                 "For the WHOLE shift, leave this blank and click OK." & vbNewLine & _
                                 "For PART of it, enter the hours you can work.", TTL, Type:=2)
        If VarType(a) = vbBoolean Then Exit Sub                    ' Cancel
        If Trim$(CStr(a)) <> "" Then
            If Not IsNumeric(Trim$(CStr(a))) Then MsgBox "Enter a number of hours, or leave it blank for the whole shift.", vbExclamation, TTL: Exit Sub
            hrs = CDbl(Trim$(CStr(a)))
            If hrs <= 0 Then MsgBox "The hours must be more than 0.", vbExclamation, TTL: Exit Sub
            If hrs >= Num(p(k, 5)) Then hrs = 0                    ' the whole shift
        End If
        If Not ConfirmFor(who, "Sign up {who} for" & vbNewLine & what & vbNewLine & _
                          IIf(hrs > 0, "for " & NumTxt(hrs) & " of the " & NumTxt(Num(p(k, 5))) & " hours?", "for the whole shift?") & vbNewLine & vbNewLine & _
                          "The request goes to the supervisors' roster with this PC's login (" & MemberLogin() & ").", TTL) Then Exit Sub
    Else
        If Not ConfirmFor(who, "Withdraw {who}'s sign-up for" & vbNewLine & what & "?", TTL) Then Exit Sub
    End If

    ' one new file; never an existing one
    Dim login As String: login = SafeName(MemberLogin())
    Dim pc As String: pc = SafeName(Environ$("COMPUTERNAME"))      ' M3 (M12): two PCs on one login never share a name
    Dim stamp As Date: stamp = Now
    Dim path As String, n As Long, ff As Integer, em As String
    path = folder & UCase$(postId) & "__" & login & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & ".txt"
    Do While Len(Dir$(path)) > 0 And n < 20
        n = n + 1
        path = folder & UCase$(postId) & "__" & login & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & "-" & n & ".txt"
    Loop
    On Error GoTo Failed
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "posting=" & UCase$(postId)
    Print #ff, "employee=" & who
    Print #ff, "hours=" & IIf(hrs > 0, Replace(CStr(hrs), ",", "."), "")
    Print #ff, "action=" & action
    Print #ff, "login=" & MemberLogin()
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

' M4 (items 34-36): the name, "since you last looked", four tiles (Total hours, My call position, Am I next?, Sick
' time), then - under the buttons - the open postings I can sign up for and the days near the top of the list.
Private Sub DrawHomeCard()
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    Dim who As String: who = WhoAmI()
    UnprotectWs ws
    Dim meRow As Long                                ' (review) known at Finish whichever way it is reached
    ClearTiles ws
    ClearBelow ws, CARD_ROW
    Dim r As Long: r = CARD_ROW
    If who = "" Or Not HasData() Then GoTo Finish
    Dim rv As Variant: rv = Data("dRoster")
    meRow = RosterRowOf(rv, who)
    If meRow = 0 Then
        ws.Range("B7").Value = who & " isn't in the published roster. Ask your supervisor."
        GoTo Finish
    End If
    Dim homeU As String: homeU = Txt(rv(meRow, RC_HOME))

    ' ---- the open postings I can sign up for (listed, or out of department), soonest first ----
    Dim p As Variant: p = Data("dPost")
    Dim l As Variant: l = Data("dPostList")
    Dim place As Long, signedTxt As String, isSigned As Boolean, statusTxt As String, codeCol As Long, placeTxt As String, nMine As Long
    Dim ordC() As Long, nPc As Long, q As Long, i As Long
    Dim myIds As String, nextPost As String, nextPostTxt As String
    Dim mineRow() As Long, mineTxt() As String, mineStat() As String
    If IsArray(p) Then nPc = UBound(p, 1) - 1
    If nPc > 0 Then
        ReDim mineRow(1 To nPc): ReDim mineTxt(1 To nPc): ReDim mineStat(1 To nPc)
        SoonestFirst p, nPc, ordC
        For q = 1 To nPc
            i = ordC(q)
            place = MyPlace(l, Txt(p(i, 1)), who, signedTxt)
            placeTxt = ""
            If place > 0 Then
                placeTxt = Ordinal(place) & " on the list"
            ElseIf place = 0 Then
                codeCol = UnitCodeCol(Txt(p(i, 2)))
                If codeCol > 0 And codeCol <= UBound(rv, 2) Then
                    If Txt(rv(meRow, codeCol)) <> "" Then placeTxt = "out of department"
                End If
            End If                                       ' place < 0: Exempt - not one they can sign up for
            If placeTxt <> "" Then
                nMine = nMine + 1
                statusTxt = MyStatus(Txt(p(i, 1)), who, signedTxt, isSigned)
                mineRow(nMine) = i: mineTxt(nMine) = placeTxt: mineStat(nMine) = statusTxt
                myIds = myIds & "," & Txt(p(i, 1))
                If nextPost = "" And Not isSigned And SignupsOn() And Now <= ClosesAt(Dt(p(i, 8))) Then
                    nextPost = Txt(p(i, 1))
                    nextPostTxt = "open posting " & nextPost & ": " & Format$(Dt(p(i, 3)), "ddd mmm d") & " " & Txt(p(i, 4)) & ", " & Txt(p(i, 2))
                End If
            End If
        Next q
    End If

    ' ---- the name, the crew and home unit, since you last looked ----
    With ws.Range("B7")
        .Value = who & "     Crew " & Txt(rv(meRow, RC_CREW)) & "  -  " & IIf(homeU = "", "no home unit on file", homeU)
        .Font.Bold = True: .Font.Size = 13: .Font.Color = NAVY
    End With

    ' ---- the call positions: the home unit first ----
    Dim units As Variant: units = UnitList()
    Dim c As Long, k As Long, nIn As Long, posMain As String, posMainDetail As String, posOther As String, mainPos As Long
    For k = LBound(units) To UBound(units)
        c = RC_FIXED + 2 * (k - LBound(units)) + 1
        If c + 1 <= UBound(rv, 2) Then
            If Txt(rv(meRow, c)) <> "" Then
                nIn = 0
                For i = 2 To UBound(rv, 1)
                    If Txt(rv(i, c)) <> "" Then nIn = nIn + 1
                Next i
                If posMain = "" And (StrComp(Txt(units(k)), homeU, vbTextCompare) = 0 Or homeU = "") Then
                    mainPos = CLng(Num(rv(meRow, c + 1)))
                    posMain = Ordinal(mainPos) & " of " & nIn
                    posMainDetail = "in " & Txt(units(k)) & " (" & Txt(rv(meRow, c)) & ")"
                Else
                    posOther = posOther & IIf(posOther = "", "", "  -  ") & Ordinal(CLng(Num(rv(meRow, c + 1)))) & " of " & nIn & " in " & Txt(units(k)) & " (" & Txt(rv(meRow, c)) & ")"
                End If
            End If
        End If
    Next k
    If posMain = "" And posOther <> "" Then posMain = Split(posOther, " in ")(0): posMainDetail = posOther: posOther = ""

    With ws.Range("B8:I8")
        .Merge
        .Value = SinceText(who, myIds, homeU, mainPos)
        .WrapText = True: .VerticalAlignment = xlTop
        .Font.Size = 9: .Font.Color = RGB(60, 60, 60)
    End With

    ' ---- the four tiles ----
    Dim totTxt As String, detail As String
    totTxt = CStr(HoursText(rv(meRow, RC_TOTAL), rv(meRow, RC_PROV)))
    detail = NumTxt(Num(rv(meRow, RC_WORKED))) & " worked + " & NumTxt(Num(rv(meRow, RC_REFUSED))) & " charged (refusals)" & _
             IIf(Num(rv(meRow, RC_PROV)) > 0, " + " & NumTxt(Num(rv(meRow, RC_PROV))) & " accepted, not yet in Kronos", "") & _
             IIf(RefusedKinds(rv, meRow) <> "", "   Charged: " & RefusedKinds(rv, meRow), "")
    Tile ws, 2, "TOTAL HOURS", totTxt, detail
    Tile ws, 4, "MY CALL POSITION", IIf(posMain = "", "-", posMain), IIf(posMainDetail = "", "no unit code on file", posMainDetail) & IIf(posOther <> "", "   " & posOther, "")
    Dim nxPlace As Long, nxTotal As Long, nxD As Date, nxSh As String, nxU As String, nxEst As Boolean
    If NextParts(who, nxPlace, nxTotal, nxD, nxSh, nxU, nxEst) Then
        Tile ws, 6, "AM I NEXT?", Ordinal(nxPlace), Format$(nxD, "ddd mmm d") & " " & nxSh & ", " & nxU & " - " & Ordinal(nxPlace) & " of " & nxTotal & _
             IIf(nxEst, " (~ estimate after a planned move)", "") & IIf(nextPostTxt <> "", "   " & nextPostTxt, "")
    Else
        Tile ws, 6, "AM I NEXT?", "Not listed", "Not listed in the next " & Txt(Meta("Who's next days")) & " days (not free, or no unit code)." & _
             IIf(nextPostTxt <> "", "   " & nextPostTxt, "")
    End If
    Dim av As Variant: av = Data("dAttend")
    Dim j As Long, found As Boolean
    If IsArray(av) Then
        For j = 2 To UBound(av, 1)
            If StrComp(Txt(av(j, 1)), who, vbTextCompare) = 0 And UBound(av, 2) >= 6 Then
                Tile ws, 8, "SICK TIME", NumTxt(Num(av(j, 3))) & " left", NumTxt(Num(av(j, 2))) & " used, " & NumTxt(Num(av(j, 4))) & " occurrence(s), " & _
                     NumTxt(Num(av(j, 5))) & " personal day(s) - counts only.  " & Txt(av(j, 6))
                found = True
                Exit For
            End If
        Next j
    End If
    If Not found Then Tile ws, 8, "SICK TIME", "-", "not published - ask your supervisor"
    ' the actions under the tiles: the total opens My history; the next posting's Sign me up
    Btn ws, "mb_hometotal", "My history  >", "MemberGoHistory", ws.Range("B12").Left + 4, ws.Range("B12").Top + 2, 110, 16, vbWhite, NAVY
    If nextPost <> "" Then
        Btn ws, "mb_homesign", "Sign me up: " & nextPost, "'MemberSignUp """ & nextPost & """'", ws.Range("F12").Left + 4, ws.Range("F12").Top + 2, 160, 16, GREEN, vbWhite
    Else
        On Error Resume Next
        ws.Shapes("mb_homesign").Delete
        On Error GoTo 0
    End If

    ' ---- under the buttons: the open postings I can sign up for ----
    ws.Cells(r, 2).Value = "Open postings I can sign up for"
    ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Color = NAVY
    r = r + 1
    For q = 1 To nMine
        i = mineRow(q)
        ws.Cells(r, 2).Value = Txt(p(i, 1)) & "  " & Txt(p(i, 2))
        ws.Cells(r, 3).NumberFormat = "@"
        ws.Cells(r, 3).Value = Format$(Dt(p(i, 3)), "ddd mmm d") & "  " & Txt(p(i, 4)) & "  " & NumTxt(Num(p(i, 5))) & " h   |   " & mineTxt(q) & _
                               "   |   closes " & Format$(Dt(p(i, 8)), "ddd mmm d") & IIf(mineStat(q) <> "", "   |   " & mineStat(q), "")
        r = r + 1
    Next q
    If nMine = 0 Then
        ws.Cells(r, 3).Value = "none right now"
        r = r + 1
    Else
        ws.Cells(r, 3).Value = IIf(SignupsOn(), "Open Postings has a Sign me up button for each.", "Sign up on the posted paper sheet.")
        ws.Cells(r, 3).Font.Size = 8: ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(100, 100, 100)
        r = r + 1
    End If
    r = r + 1

    ' the next days where they are among the first three free in their home unit
    Dim nv As Variant: nv = Data("dNext")
    Dim nextTxt As String, nNext As Long
    If IsArray(nv) And homeU <> "" Then
        For i = 2 To UBound(nv, 1)
            If Num(nv(i, 4)) <= 3 Then
                If StrComp(Txt(nv(i, 5)), who, vbTextCompare) = 0 And StrComp(Txt(nv(i, 1)), homeU, vbTextCompare) = 0 Then
                    nNext = nNext + 1
                    If nNext <= 8 Then nextTxt = nextTxt & IIf(nextTxt = "", "", "     ") & Format$(Dt(nv(i, 2)), "ddd mmm d") & " " & Txt(nv(i, 3)) & " " & Ordinal(CLng(Num(nv(i, 4))))
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
    r = r + 1
    ' M4 (item 37): the days I am NOT listed in my home unit, and why (reason words only)
    Dim fv As Variant: fv = Data("dNotFree")
    Dim notTxt As String, nNot As Long
    If IsArray(fv) And homeU <> "" Then
        For i = 2 To UBound(fv, 1)
            If StrComp(Txt(fv(i, 4)), who, vbTextCompare) = 0 And StrComp(Txt(fv(i, 1)), homeU, vbTextCompare) = 0 Then
                If Txt(fv(i, 5)) <> "scheduled" Then        ' your own shifts are no news
                    nNot = nNot + 1
                    If nNot <= 6 Then notTxt = notTxt & IIf(notTxt = "", "", "     ") & Format$(Dt(fv(i, 2)), "ddd mmm d") & " " & Txt(fv(i, 3)) & ": " & Txt(fv(i, 5))
                End If
            End If
        Next i
    End If
    If notTxt <> "" Then
        ws.Cells(r, 2).Value = "Not listed (besides your own shifts)"
        ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Color = NAVY
        r = r + 1
        ws.Cells(r, 3).Value = homeU & ":  " & notTxt & IIf(nNot > 6, "  ... and " & (nNot - 6) & " more", "")
        r = r + 1
    End If
    ws.Cells(r, 3).Value = "Who is free changes with time off and switches: this is as of the last publish. Sick time counts since Jan 1, or since your counters restarted."
    ws.Cells(r, 3).Font.Size = 8: ws.Cells(r, 3).Font.Italic = True: ws.Cells(r, 3).Font.Color = RGB(100, 100, 100)
Finish:
    If who = "" Or Not HasData() Or meRow = 0 Then  ' (review) a name not on the roster leaves no buttons behind
        On Error Resume Next
        ws.Shapes("mb_homesign").Delete
        ws.Shapes("mb_hometotal").Delete
        On Error GoTo 0
    End If
    ProtectWs ws
    MarkClean
End Sub

' M4: the tiles' rows (7-12) emptied before a redraw
Private Sub ClearTiles(ByVal ws As Worksheet)
    With ws.Range("B7:I12")
        .UnMerge
        .ClearContents
        .ClearFormats
        .Font.Name = "Segoe UI": .Font.Size = 10: .Font.Color = INK
    End With
End Sub

' M4: one tile - two columns from col, the label (row 9), the figure (row 10), the line under it (row 11)
Private Sub Tile(ByVal ws As Worksheet, ByVal col As Long, ByVal label As String, ByVal figure As String, ByVal detail As String)
    With ws.Range(ws.Cells(9, col), ws.Cells(9, col + 1))
        .Merge
        .Value = label
        .Font.Size = 8: .Font.Bold = True: .Font.Color = RGB(100, 100, 100)
    End With
    With ws.Range(ws.Cells(10, col), ws.Cells(10, col + 1))
        .Merge
        .NumberFormat = "@"
        .Value = figure
        .Font.Size = 16: .Font.Bold = True: .Font.Color = NAVY
    End With
    With ws.Range(ws.Cells(11, col), ws.Cells(11, col + 1))
        .Merge
        .NumberFormat = "@"
        .Value = detail
        .WrapText = True: .VerticalAlignment = xlTop
        .Font.Size = 8: .Font.Color = RGB(60, 60, 60)
    End With
    With ws.Range(ws.Cells(9, col), ws.Cells(12, col + 1))
        .Interior.Color = RGB(240, 244, 250)
        .IndentLevel = 1
        .BorderAround xlContinuous, xlThin, , RGB(200, 210, 225)
    End With
End Sub

' 1st, 2nd, 3rd, 4th ... 11th, 12th, 13th, 21st
Private Function Ordinal(ByVal n As Long) As String
    Dim s As String
    Select Case n Mod 100
        Case 11, 12, 13: s = "th"
        Case Else
            Select Case n Mod 10
                Case 1: s = "st"
                Case 2: s = "nd"
                Case 3: s = "rd"
                Case Else: s = "th"
            End Select
    End Select
    Ordinal = n & s
End Function

' M4 (item 36): what changed since this PC user last looked at this name - new open postings in my units, postings
' that closed (awarded, no takers or cancelled), my place in my home unit. Read once a session per name, then the
' new state is kept (Excel's per-user setting store, never the shared file).
Private Function SinceText(ByVal who As String, ByVal myIds As String, ByVal homeU As String, ByVal mainPos As Long) As String
    If mSeen Is Nothing Then Set mSeen = CreateObject("Scripting.Dictionary"): mSeen.CompareMode = 1
    If mSeen.Exists(who) Then SinceText = mSeen(who): Exit Function      ' (review) once a session PER name
    On Error Resume Next
    Dim key As String: key = "Seen " & Left$(SafeName(Replace(who, " ", "_")), 60)
    Dim stored As String: stored = GetSetting(APP_KEY, APP_SEC, key, "")
    Dim s As String
    If stored = "" Then
        s = "First look at " & who & " on this PC: from the next open, what changed shows here."
    Else
        Dim parts As Variant: parts = Split(stored, "|")      ' at | ids | unit | pos
        Dim atTxt As String, oldIds As String, oldUnit As String, oldPos As Long
        If UBound(parts) >= 3 Then atTxt = parts(0): oldIds = parts(1): oldUnit = parts(2): oldPos = CLng(Val(parts(3)))
        Dim k As Variant, nNew As Long, nGone As Long
        For Each k In Split(myIds, ",")
            If CStr(k) <> "" And InStr(oldIds & ",", "," & CStr(k) & ",") = 0 Then nNew = nNew + 1
        Next k
        For Each k In Split(oldIds, ",")
            If CStr(k) <> "" And InStr(myIds & ",", "," & CStr(k) & ",") = 0 Then nGone = nGone + 1
        Next k
        Dim bits As String
        If nNew > 0 Then bits = bits & "  -  " & nNew & " new posting" & IIf(nNew = 1, "", "s") & " you can sign up for"
        If nGone > 0 Then bits = bits & "  -  " & nGone & " posting" & IIf(nGone = 1, "", "s") & " closed (awarded, no takers or cancelled)"
        If oldPos > 0 And mainPos > 0 And oldPos <> mainPos And StrComp(oldUnit, homeU, vbTextCompare) = 0 Then
            bits = bits & "  -  your " & homeU & " place moved from " & Ordinal(oldPos) & " to " & Ordinal(mainPos)
        End If
        s = "Since you last looked (" & atTxt & "):" & IIf(bits = "", "  nothing new for you.", bits & ".")
    End If
    SaveSetting APP_KEY, APP_SEC, key, Format$(Now, "ddd mmm d, h:mm AM/PM") & "|" & myIds & "|" & homeU & "|" & mainPos
    mSeen(who) = s
    SinceText = s
End Function

' M4: the earliest published shift (dNext) this person is listed for - place, how many listed, the date, shift and
' unit, and whether their hours there are an estimate. False = not listed in the published days.
Private Function NextParts(ByVal who As String, ByRef place As Long, ByRef total As Long, ByRef d As Date, _
                           ByRef sh As String, ByRef unit As String, ByRef est As Boolean) As Boolean
    Dim nv As Variant: nv = Data("dNext")
    Dim i As Long, best As Long, dd As Date
    place = 0: total = 0: d = 0: sh = "": unit = "": est = False
    If Not IsArray(nv) Then Exit Function
    For i = 2 To UBound(nv, 1)
        If StrComp(Txt(nv(i, 5)), who, vbTextCompare) = 0 Then
            dd = Dt(nv(i, 2))
            If best = 0 Or dd < d Or (dd = d And Txt(nv(i, 3)) = "Days" And sh = "Nights") Then
                best = i: d = dd: unit = Txt(nv(i, 1)): sh = Txt(nv(i, 3))
            End If
        End If
    Next i
    If best = 0 Then Exit Function
    For i = 2 To UBound(nv, 1)
        If StrComp(Txt(nv(i, 1)), unit, vbTextCompare) = 0 And Int(Dt(nv(i, 2))) = Int(d) And _
           StrComp(Txt(nv(i, 3)), sh, vbTextCompare) = 0 Then total = total + 1
    Next i
    place = CLng(Num(nv(best, 4)))
    If UBound(nv, 2) >= 9 Then est = (StrComp(Txt(nv(best, 9)), "Yes", vbTextCompare) = 0)
    NextParts = True
End Function


' M3 (C2) "Am I next?": the earliest published shift (dNext, the master's Who's next window) this person is
' listed for, with their place: "Tue Oct 21 Days, Phosphine - you are #2 of 9" (~ = an estimate after a
' planned unit move). No rule: the master worked the order out.
Private Function NextForMe(ByVal who As String) As String
    Dim place As Long, total As Long, d As Date, sh As String, unit As String, est As Boolean
    If Not NextParts(who, place, total, d, sh, unit, est) Then
        NextForMe = "Not listed in the next " & Txt(Meta("Who's next days")) & " days (not free, or no unit code)."
        Exit Function
    End If
    NextForMe = Format$(d, "ddd mmm d") & " " & sh & ", " & unit & "  -  you are " & Ordinal(place) & " of " & total & _
                IIf(est, "  (~ estimated hours after a planned unit move)", "")
End Function

' M3 (C2) "Larger type": every view sheet at 130 % or back to 100 %, for this session (the file never saves)
Public Sub MemberZoomToggle()
    On Error Resume Next
    Dim was As Worksheet: Set was = ActiveSheet
    Application.ScreenUpdating = False
    Dim ws As Worksheet: Set ws = WsOf(SH_HOME)
    ws.Activate
    Dim big As Boolean: big = (ActiveWindow.Zoom < 130)     ' the state is the window's zoom itself (review: no module flag to go stale)
    Dim nm As Variant
    For Each nm In ViewSheets()
        WsOf(CStr(nm)).Activate
        ActiveWindow.Zoom = IIf(big, 130, 100)
    Next nm
    If Not was Is Nothing Then was.Activate
    For Each nm In ViewSheets()                      ' M4: every header band's button says what a click does next
        Set ws = WsOf(CStr(nm))
        UnprotectWs ws
        ws.Shapes("mb_band_zoom").TextFrame2.TextRange.Text = IIf(big, "Normal type", "Larger type")
        ProtectWs ws
    Next nm
    Application.ScreenUpdating = True
    MarkClean
End Sub


' ====================================================================
'  MY HISTORY
' ====================================================================
Public Sub MemberShowHistory()
    DrawHistory Txt(WsOf(SH_HIST).Range(HIST_NAME).Value)
End Sub

Private Sub DrawHistory(ByVal nm As String)
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
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
    ' M3: dLog carries Kind, Call Order and Why (the master's charge sentence for a charged row); no rule is
    ' worked out here. The same columns as the master's member-safe history report, row for row.
    Dim nc As Long: nc = 10
    Header ws, 5, Array("Shift date", "Ref", "Type", "Unit / shift", "Hours", "Call order", "Outcome", "Charged", "Pay period", "Why")
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
    Dim totalsP As Object: Set totalsP = CreateObject("Scripting.Dictionary")     ' M3.1: posting refusals per period
    Dim totalsC As Object: Set totalsC = CreateObject("Scripting.Dictionary")     ' M3.1: call-out refusals per period
    Dim yearTot As Double, ps As String, kCall As Double, kPost As Double, kOther As Double
    Dim hasM3 As Boolean: hasM3 = (UBound(v, 2) >= 13)
    Dim kind As String
    For i = 1 To n
        out(i, 1) = Dt(v(idx(i), 1))
        out(i, 2) = Txt(v(idx(i), 2))
        out(i, 3) = Txt(v(idx(i), 3))
        out(i, 4) = Txt(v(idx(i), 4)) & " / " & Txt(v(idx(i), 5))
        out(i, 5) = Num(v(idx(i), 6))
        If hasM3 Then out(i, 6) = Txt(v(idx(i), 12))
        out(i, 7) = Txt(v(idx(i), 8))
        out(i, 8) = Num(v(idx(i), 9))
        out(i, 9) = Format$(Dt(v(idx(i), 10)), "mmm d")
        If hasM3 Then out(i, 10) = Txt(v(idx(i), 13))
        ps = Format$(Dt(v(idx(i), 10)), "yyyy-mm-dd")
        totals(ps) = Num(totals(ps)) + Num(v(idx(i), 9))
        yearTot = yearTot + Num(v(idx(i), 9))
        If hasM3 Then kind = LCase$(Txt(v(idx(i), 11))) Else kind = ""    ' M3: the master's Kind column (the bucket rule lives there); not IIf (it reads column 11 either way)
        Select Case kind
            Case "callout": kCall = kCall + Num(v(idx(i), 9)): totalsC(ps) = Num(totalsC(ps)) + Num(v(idx(i), 9))
            Case "posting": kPost = kPost + Num(v(idx(i), 9)): totalsP(ps) = Num(totalsP(ps)) + Num(v(idx(i), 9))
            Case Else: kOther = kOther + Num(v(idx(i), 9))
        End Select
    Next i
    With ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, nc))
        .Value = out
        .RowHeight = 17
        .VerticalAlignment = xlTop
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(228, 232, 238)
    End With
    ws.Range(ws.Cells(6, 1), ws.Cells(5 + n, 1)).NumberFormat = "ddd mmm d"
    ws.Range(ws.Cells(6, 5), ws.Cells(5 + n, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 8), ws.Cells(5 + n, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 10), ws.Cells(5 + n, 10)).WrapText = True
    ws.Range(ws.Rows(6), ws.Rows(5 + n)).AutoFit
    r = 5 + n + 2
    ' M3.1 (v26.5, Graeme Oct 19): per pay period the total charged, the posting refusals, the call-out refusals and the
    ' premium hours worked (dWorked, from the Kronos export; a dash until an export has been imported)
    Dim wkd As Variant: wkd = Data("dWorked")
    Dim worked As Object: Set worked = CreateObject("Scripting.Dictionary")
    Dim hasWorked As Boolean
    If IsArray(wkd) Then
        For i = 2 To UBound(wkd, 1)
            If StrComp(Txt(wkd(i, 1)), nm, vbTextCompare) = 0 Then
                ps = Format$(Dt(wkd(i, 2)), "yyyy-mm-dd")
                worked(ps) = Num(worked(ps)) + Num(wkd(i, 3)): hasWorked = True
                If Not totals.Exists(ps) Then totals(ps) = 0        ' a period with hours worked but nothing charged still gets a row
            End If
        Next i
    End If
    Header ws, r, Array("Pay period", "Charged hours", "Posting refusals", "Call-out refusals", "Worked (Kronos)", "", "", "", "", "")
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
            ws.Cells(r, 3).Value = Num(totalsP(ks(i)))
            ws.Cells(r, 4).Value = Num(totalsC(ks(i)))
            If hasWorked Then ws.Cells(r, 5).Value = Num(worked(ks(i))) Else ws.Cells(r, 5).Value = "-"
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
            r = r + 1
        Next i
    End If
    Dim wkTot As Double, wk As Variant
    For Each wk In worked.Keys
        wkTot = wkTot + Num(worked(wk))
    Next wk
    ws.Cells(r, 1).Value = "Year to date (every period logged)"     ' M3.1: a later period's charge is here before the card's Refused counts it
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = yearTot
    ws.Cells(r, 3).Value = kPost
    ws.Cells(r, 4).Value = kCall
    If hasWorked Then ws.Cells(r, 5).Value = wkTot Else ws.Cells(r, 5).Value = "-"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).Font.Bold = True
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    r = r + 1
    If Not hasWorked Then
        ws.Cells(r, 1).Value = "Worked (Kronos) shows once the supervisors import a Kronos export (premium hours per pay period)."
        ws.Cells(r, 1).Font.Size = 8: ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 100, 100)
        r = r + 1
    End If
    ws.Cells(r, 1).Value = "  call-out refusals": ws.Cells(r, 2).Value = kCall: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    r = r + 1
    ws.Cells(r, 1).Value = "  posting refusals": ws.Cells(r, 2).Value = kPost: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    If Abs(kOther) > 0.0001 Then
        r = r + 1
        ws.Cells(r, 1).Value = "  other (adjustments)": ws.Cells(r, 2).Value = kOther: ws.Cells(r, 2).HorizontalAlignment = xlCenter
    End If
    Err.Clear                                    ' (a normal arrival carries no error)
Fail:                                            ' reached normally, or from an error anywhere above
    If Err.Number <> 0 Then mViewErr = "My history: " & Err.Description
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
Public Sub MemberGoRequests(): MemberShowRequests: WsOf(SH_REQ).Activate: End Sub     ' M5 (redrawn when opened: the folder may have changed)

Public Sub MemberOnOpen()
    On Error Resume Next
    Application.EnableEvents = True                ' (only matters when this is run by hand: with events off, Excel never calls it on open)
    MemberRefreshAll
    ' M4 (item 34): the name this PC user picked last time, if it is still on the published roster
    Dim me_ As String: me_ = RememberedMe()
    If me_ <> "" And LayoutProblem() = "" And DataProblem() = "" Then
        If RosterRowOf(Data("dRoster"), me_) > 0 Then ApplyMe me_, SH_HOME
    End If
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

' ====================================================================
'  M4 (item 34): ONE HEADER BAND ON EVERY SHEET, THE STAMP, HOW IT WORKS
' ====================================================================
' The view sheets in tab order (Home first, How it works last)
Private Function ViewSheets() As Variant
    ViewSheets = Array(SH_HOME, SH_ROSTER, SH_NEXT, SH_CAL, SH_POST, SH_HIST, SH_REQ, SH_HELP)     ' M5: + Requests
End Function

' < Home, Larger type, How it works and the published stamp, on row 1 to the right of the title (row 1 is made taller)
Private Sub HeaderBand(ByVal ws As Worksheet)
    ws.Rows(1).RowHeight = 40
    Dim x As Double: x = ws.Range("A1").Left + 230
    Dim y As Double: y = ws.Range("A1").Top + 19
    Btn ws, "mb_band_home", "<  Home", "MemberGoHome", x, y, 60, 18, NAVY, vbWhite
    Btn ws, "mb_band_zoom", "Larger type", "MemberZoomToggle", x + 66, y, 80, 18, LIGHT, INK
    If ws.Name <> SH_HELP Then Btn ws, "mb_band_help", "How it works", "MemberGoHelp", x + 152, y, 90, 18, LIGHT, INK
    Dim s As Shape
    On Error Resume Next
    ws.Shapes("mb_stamp").Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, ws.Range("A1").Top + 2, 330, 14)
    s.Name = "mb_stamp"
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .WordWrap = msoFalse
        .MarginLeft = 0: .MarginTop = 0: .MarginBottom = 0
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = 8
    End With
End Sub

' "Published Mon Oct 19, 7:12 AM", or "Published Oct 12 - 7 days old: ask a supervisor to update" past the age the
' master published (Members stale days)
Private Function StampText() As String
    If Not HasData() Then StampText = "Not published yet": Exit Function
    Dim pub As Date: pub = Dt(Meta("Published at"))
    If StampOld() Then
        StampText = "Published " & Format$(pub, "ddd mmm d") & " - " & (Int(Date) - Int(pub)) & " days old: ask a supervisor to update"
    Else
        StampText = "Published " & Format$(pub, "ddd mmm d, h:mm AM/PM")
    End If
End Function

Private Function StampOld() As Boolean
    If Not HasData() Then Exit Function
    Dim lim As Long: lim = CLng(Num(Meta("Members stale days")))
    If lim <= 0 Then lim = 3
    StampOld = (Int(Date) - Int(Dt(Meta("Published at"))) >= lim)
End Function

' the stamp written into every band (after a publish and on open)
Private Sub StampBands()
    Dim nm As Variant, ws As Worksheet
    On Error Resume Next
    For Each nm In ViewSheets()
        Set ws = WsOf(CStr(nm))
        If Not ws Is Nothing Then
            If CStr(nm) <> SH_HOME Then
                UnprotectWs ws
                With ws.Shapes("mb_stamp").TextFrame2.TextRange
                    .Text = StampText()
                    .Font.Fill.ForeColor.RGB = IIf(StampOld(), RGB(176, 0, 32), RGB(100, 100, 100))
                    .Font.Bold = IIf(StampOld(), msoTrue, msoFalse)
                End With
                ProtectWs ws
            End If
        End If
    Next nm
End Sub

Public Sub MemberGoHome(): WsOf(SH_HOME).Activate: End Sub
Public Sub MemberGoHelp(): WsOf(SH_HELP).Activate: End Sub

Private Sub LayoutHelp()
    Dim ws As Worksheet: Set ws = WsOf(SH_HELP, True)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "HOW IT WORKS", "The rules the supervisors' roster uses, in plain words. Nothing here is worked out in this file."
    ws.Columns("A").ColumnWidth = 110
    ProtectWs ws
End Sub

' M4 (item 35): the rules once, in the file (the legend comes from the master's own text)
Private Sub DrawHelp()
    Dim ws As Worksheet: Set ws = WsOf(SH_HELP)
    If ws Is Nothing Then Exit Sub
    UnprotectWs ws
    ClearBelow ws, 4
    Dim r As Long: r = 5
    HelpPara ws, r, "Who is called first", "Everyone with a code in the unit who is free for the shift is listed. The order: the grade in the code (D1 before D2), then the lowest Total hours, then the earliest plant seniority, then the name. Total hours = the hours worked (from Kronos) + the hours charged for refusals + overtime accepted but not in Kronos yet (marked *)."
    HelpPara ws, r, "Why was I charged?", "A refusal or a message left on a call-out is charged the hours offered x 1.75, for grade 2 and below. On a posting, everyone ahead of the last person awarded who did not sign up is charged the posting's hours x 1.75. My history shows a Why line for every charge."
    HelpPara ws, r, "Not charged", "Accepting, being unavailable for a reason, a plant upset (exception), a cancelled posting, a part acceptance that closed the shift for those below it, and an Exempt row (moved unit, left, or restricted before the shift) are never charged."
    HelpPara ws, r, "~ and *", "~ = the hours someone will have after a planned unit move (used for their place only). * = the total includes overtime accepted but not yet in Kronos."
    HelpPara ws, r, "Sign-ups", "Sign me up drops a request for the supervisors' roster; the supervisor confirms it when the posting is awarded. The posted paper sheet still counts. A sign-up for part of the shift is awarded only when nobody takes the whole shift."
    HelpPara ws, r, "Who is free", "Not free: scheduled on that shift, coming off nights that morning or starting days the next morning (the rest rule), booked off, on a shift switch, inside the 24 h after a sick shift, or when the 16 h rule leaves nothing (no more than 16 h of work in a row, and 11 h rest before the next work)."
    HelpPara ws, r, "Asking for time off or a switch", "The Requests page sends a request to the supervisors' roster: time off between two dates, or a shift switch on one date with the partner who would work it. A switch waits until that partner clicks Agree on their own PC. An admin then books it through the roster's own rules or declines it with a reason; the answer shows under My requests after the next update. A request books nothing by itself; your note is read by the admin (the Requests folder beside this file is shared, like Signups)."
    HelpPara ws, r, "The calendar", Txt(Meta("Legend calendar"))
    HelpPara ws, r, "The log and the lists", Txt(Meta("Legend log"))
    ws.Cells(r + 1, 1).Value = IIf(HasData(), "As published " & Format$(Dt(Meta("Published at")), "ddd mmm d, h:mm AM/PM") & ".  " & Txt(Meta("Trust strip")), "Not published yet.")     ' (review) the template has no date
    ws.Cells(r + 1, 1).Font.Size = 8: ws.Cells(r + 1, 1).Font.Italic = True: ws.Cells(r + 1, 1).Font.Color = RGB(100, 100, 100)
    ws.Cells(r + 1, 1).WrapText = True
    ProtectWs ws
End Sub

Private Sub HelpPara(ByVal ws As Worksheet, ByRef r As Long, ByVal head As String, ByVal body As String)
    If body = "" Then Exit Sub
    ws.Cells(r, 1).Value = head
    ws.Cells(r, 1).Font.Bold = True: ws.Cells(r, 1).Font.Size = 11: ws.Cells(r, 1).Font.Color = NAVY
    ws.Cells(r + 1, 1).Value = body
    ws.Cells(r + 1, 1).WrapText = True
    ws.Rows(r + 1).AutoFit
    r = r + 3
End Sub


' ====================================================================
'  M5 (v28 item 1): REQUESTS - ask for time off or a shift switch; the partner's Agree / No; My requests; Withdraw.
'  One small text file per click into the Requests folder beside this file (as the sign-ups): the master's
'  modRequests reads them. This file reads the same folder so the named partner sees "asked of me" at once and a
'  sent request shows as waiting before the roster has published its answer. No rule lives here: the admin books
'  through the roster's own pop-ups.
' ====================================================================
' (the M5 constants - RQ_TTL, DQ_*, FQ_* - sit in the declarations section at the top of the module: a module-level
'  Const after the first procedure does not compile - review v28)

Private Function RequestsOn() As Boolean
    RequestsOn = (LCase$(Txt(Meta("Requests"))) = "on")
End Function

' the folder as published (".\Requests" = beside this file), resolved here
Private Function RequestFolderPath() As String
    Dim folder As String: folder = Txt(Meta("Request folder"))
    If folder = "." Then
        folder = ThisWorkbook.Path
    ElseIf Left$(folder, 2) = ".\" Then
        folder = ThisWorkbook.Path & Mid$(folder, 2)
    End If
    If folder <> "" Then
        If Right$(folder, 1) <> Application.PathSeparator Then folder = folder & Application.PathSeparator
    End If
    RequestFolderPath = folder
End Function

Private Sub LayoutRequests()
    Dim ws As Worksheet: Set ws = WsOf(SH_REQ, True)
    UnprotectWs ws
    Dim sk As Long
    For sk = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(sk).Name, 3) = "mb_" Then ws.Shapes(sk).Delete
    Next sk
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "REQUESTS", "Ask for time off, or a shift switch with a partner. An admin books it through the roster's rules or declines it; the answer shows below after the next update."
    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 9
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 22
    ws.Columns("F").ColumnWidth = 9
    ws.Columns("G").ColumnWidth = 22
    ws.Columns("H").ColumnWidth = 26
    ws.Columns("I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 10
    ws.Columns("K").ColumnWidth = 10
    ws.Rows(4).RowHeight = 20: ws.Rows(5).RowHeight = 20: ws.Rows(6).RowHeight = 20
    Picker ws, REQ_KIND, "", "Kind:"
    Picker ws, REQ_FROM, "", "From:"
    Picker ws, REQ_TO, "", "To:"
    Picker ws, REQ_SHIFT, "", "Shift:"
    Picker ws, REQ_PARTNER, "", "Partner:"
    Picker ws, REQ_NOTE, "", "Note:"
    ws.Range("C6:G6").Merge
    ws.Range("C6").MergeArea.Locked = False          ' (review) Picker unlocked C6 alone; the merged note cell as a whole
    With ws.Range("A7")
        .Value = "Dates like Nov 3, Nov 3 2026 or 2026-11-03; To blank = the one day. A switch is ONE shift: pick Days or Nights and the partner who would work it - they click Agree on their PC before the roster sees it. No leave words in the note."
        .Font.Size = 8: .Font.Italic = True: .Font.Color = RGB(100, 100, 100)
    End With
    ws.Range("A7:K7").Merge: ws.Range("A7").WrapText = True: ws.Rows(7).RowHeight = 26
    Btn ws, "mb_rq_send", "SEND REQUEST", "MemberSendRequest", ws.Range("I4").Left, ws.Range("I4").Top, 120, 44, GREEN, vbWhite
    ProtectWs ws
End Sub

' the folder's REQ__*.txt files as rows (FQ_*), or Empty (no folder, nothing in it, or it can't be reached)
Private Function ReadRequestFolder() As Variant
    Dim folder As String: folder = RequestFolderPath()
    If folder = "" Then Exit Function
    Dim ok As Boolean
    On Error Resume Next
    ok = (Len(Dir$(folder, vbDirectory)) > 0)
    On Error GoTo 0
    If Not ok Then Exit Function
    Dim files() As String, n As Long, f As String
    On Error Resume Next
    f = Dir$(folder & "REQ__*.txt")
    Do While f <> "" And Err.Number = 0
        n = n + 1
        ReDim Preserve files(1 To n)
        files(n) = f
        f = Dir$()
    Loop
    On Error GoTo 0
    If n = 0 Then Exit Function
    Dim rows() As Variant: ReDim rows(1 To n, 1 To FQ_N)
    Dim i As Long, k As Long, fx() As String
    For i = 1 To n
        If ReadRequestFile(folder & files(i), fx) Then
            If fx(FQ_ID) <> "" And fx(FQ_ACTION) <> "" Then
                k = k + 1
                Dim c As Long
                For c = 1 To FQ_N
                    rows(k, c) = fx(c)
                Next c
            End If
        End If
    Next i
    If k = 0 Then Exit Function
    Dim out() As Variant: ReDim out(1 To k, 1 To FQ_N)
    Dim j As Long
    For i = 1 To k
        For j = 1 To FQ_N
            out(i, j) = rows(i, j)
        Next j
    Next i
    ReadRequestFolder = out
End Function

' one file's key=value lines (False = it could not be opened: still being written, antivirus)
Private Function ReadRequestFile(ByVal fullPath As String, ByRef fx() As String) As Boolean
    ReDim fx(1 To FQ_N)
    Dim ff As Integer, ln As String, p As Long, valTxt As String, opened As Boolean, nLines As Long
    On Error GoTo Quit
    ff = FreeFile
    Open fullPath For Input Access Read Shared As #ff
    opened = True
    Do While Not EOF(ff) And nLines < 60
        Line Input #ff, ln
        nLines = nLines + 1
        p = InStr(ln, "=")
        If p > 1 Then
            valTxt = Trim$(Mid$(ln, p + 1))
            Select Case LCase$(Trim$(Left$(ln, p - 1)))
                Case "request": fx(FQ_ID) = valTxt
                Case "action": fx(FQ_ACTION) = LCase$(valTxt)
                Case "kind": fx(FQ_KIND) = valTxt
                Case "employee": fx(FQ_EMP) = valTxt
                Case "from": fx(FQ_FROM) = valTxt
                Case "to": fx(FQ_TO) = valTxt
                Case "shift": fx(FQ_SHIFT) = valTxt
                Case "partner": fx(FQ_PARTNER) = valTxt
                Case "note": fx(FQ_NOTE) = valTxt
                Case "login": fx(FQ_LOGIN) = valTxt
                Case "stamp": fx(FQ_STAMP) = valTxt
            End Select
        End If
    Loop
    Close #ff
    ReadRequestFile = True
    Exit Function
Quit:
    On Error Resume Next
    If opened Then Close #ff
End Function

' the folder's answer files for one request: agree by another login ("" = none), refused, withdrawn
Private Sub FolderAnswers(ByRef fv As Variant, ByVal rid As String, ByVal requesterLogin As String, _
                          ByRef agreeLogin As String, ByRef refused As Boolean, ByRef withdrawn As Boolean)
    agreeLogin = "": refused = False: withdrawn = False
    If Not IsArray(fv) Then Exit Sub
    Dim i As Long
    For i = 1 To UBound(fv, 1)
        If StrComp(Txt(fv(i, FQ_ID)), rid, vbTextCompare) = 0 Then
            Select Case Txt(fv(i, FQ_ACTION))
                Case "agree"
                    If StrComp(Txt(fv(i, FQ_LOGIN)), requesterLogin, vbTextCompare) <> 0 Then agreeLogin = Txt(fv(i, FQ_LOGIN))
                Case "refuse"                               ' the master's rule (re-review): a refuse from another login, a withdraw from the one that asked
                    If StrComp(Txt(fv(i, FQ_LOGIN)), requesterLogin, vbTextCompare) <> 0 Then refused = True
                Case "withdraw"
                    If StrComp(Txt(fv(i, FQ_LOGIN)), requesterLogin, vbTextCompare) = 0 Then withdrawn = True
            End Select
        End If
    Next i
End Sub

' a published dRequests row for this request (by name, kind, dates, shift and partner), 0 = none yet
Private Function PublishedRow(ByRef dq As Variant, ByVal who As String, ByVal kind As String, ByVal fromIso As String, _
                              ByVal toIso As String, ByVal shiftLbl As String, ByVal partner As String) As Long
    If Not IsArray(dq) Then Exit Function
    Dim i As Long
    For i = UBound(dq, 1) To 2 Step -1               ' the newest row of the same ask wins
        If StrComp(Txt(dq(i, DQ_EMP)), who, vbTextCompare) = 0 And StrComp(Txt(dq(i, DQ_KIND)), kind, vbTextCompare) = 0 Then
            If Format$(Dt(dq(i, DQ_FROM)), "yyyy-mm-dd") = fromIso And Format$(Dt(dq(i, DQ_TO)), "yyyy-mm-dd") = toIso Then
                If StrComp(Txt(dq(i, DQ_SHIFT)), shiftLbl, vbTextCompare) = 0 And StrComp(Txt(dq(i, DQ_PARTNER)), partner, vbTextCompare) = 0 Then
                    PublishedRow = i: Exit Function
                End If
            End If
        End If
    Next i
End Function

Public Sub MemberShowRequests()
    If LayoutProblem() <> "" Or DataProblem() <> "" Then Exit Sub      ' the note on Home says why
    Dim ws As Worksheet: Set ws = WsOf(SH_REQ)
    If ws Is Nothing Then Exit Sub
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    On Error GoTo Finish
    ClearBelow ws, REQ_FIRST_ROW - 1
    Dim i As Long, k As Long, nBtn As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "mb_rqb" Then ws.Shapes(i).Delete
    Next i
    Dim who As String: who = WhoAmI()
    Dim r As Long: r = REQ_FIRST_ROW
    If Not RequestsOn() Then
        ws.Cells(r, 1).Value = "Requests from this file aren't switched on. Ask your supervisor the usual way."
        ws.Cells(r, 1).Font.Italic = True
        Err.Clear
        GoTo Finish
    End If
    If who = "" Then
        ws.Cells(r, 1).Value = "Pick your name on Home (Find me) to ask, or to see your requests."
        ws.Cells(r, 1).Font.Italic = True
        Err.Clear
        GoTo Finish
    End If
    Dim fv As Variant: fv = ReadRequestFolder()
    Dim dq As Variant: dq = Data("dRequests")
    Dim myLogin As String: myLogin = MemberLogin()
    Dim agreeLogin As String, refused As Boolean, withdrawn As Boolean, st As String, pr As Long

    ' ---- 1. asked of me: switch requests naming me as the partner, still waiting for my answer ----
    SectionBar ws, r, "ASKED OF " & UCase$(who) & "  -  switches where you would work the shift (Agree, or No)"
    r = r + 1
    Header ws, r, Array("Asked by", "Shift date", "Shift", "", "", "", "", "", "", "", "")
    r = r + 1
    Dim nAsked As Long, dF As Date
    If IsArray(fv) Then
        For i = 1 To UBound(fv, 1)
            If Txt(fv(i, FQ_ACTION)) = "request" And StrComp(Txt(fv(i, FQ_KIND)), "Switch", vbTextCompare) = 0 Then
                If StrComp(Txt(fv(i, FQ_PARTNER)), who, vbTextCompare) = 0 And StrComp(Txt(fv(i, FQ_EMP)), who, vbTextCompare) <> 0 Then
                    FolderAnswers fv, Txt(fv(i, FQ_ID)), Txt(fv(i, FQ_LOGIN)), agreeLogin, refused, withdrawn
                    If agreeLogin = "" And Not refused And Not withdrawn And IsoDay(Txt(fv(i, FQ_FROM)), dF) Then
                        If dF >= Date Then
                            nAsked = nAsked + 1
                            ws.Cells(r, 1).Value = Txt(fv(i, FQ_EMP))
                            ws.Cells(r, 2).Value = dF: ws.Cells(r, 2).NumberFormat = "ddd mmm d"
                            ws.Cells(r, 3).Value = Txt(fv(i, FQ_SHIFT))
                            ws.Rows(r).RowHeight = 19
                            nBtn = nBtn + 1
                            Btn ws, "mb_rqb" & nBtn, "Agree", "'MemberAgreeRequest """ & Txt(fv(i, FQ_ID)) & """'", _
                                ws.Cells(r, 9).Left + 2, ws.Cells(r, 9).Top + 2, 60, 15, GREEN, vbWhite
                            nBtn = nBtn + 1
                            Btn ws, "mb_rqb" & nBtn, "No", "'MemberRefuseRequest """ & Txt(fv(i, FQ_ID)) & """'", _
                                ws.Cells(r, 10).Left + 2, ws.Cells(r, 10).Top + 2, 50, 15, LIGHT, INK
                            r = r + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
    If nAsked = 0 Then
        ws.Cells(r, 1).Value = "Nothing waiting for your answer."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If
    r = r + 1

    ' ---- 2. my requests: the folder's (sent, maybe not yet in the roster) and the published answers ----
    SectionBar ws, r, "MY REQUESTS"
    r = r + 1
    Header ws, r, Array("Kind", "From", "To", "Shift", "Partner", "Status", "Answer", "Decided", "", "", "")
    r = r + 1
    Dim nMine As Long, shown As String, kind As String, fromIso As String, toIso As String, dTo As Date, keyTxt As String
    If IsArray(fv) Then
        For i = 1 To UBound(fv, 1)
            If Txt(fv(i, FQ_ACTION)) = "request" And StrComp(Txt(fv(i, FQ_EMP)), who, vbTextCompare) = 0 Then
                kind = Txt(fv(i, FQ_KIND)): fromIso = Txt(fv(i, FQ_FROM)): toIso = Txt(fv(i, FQ_TO))
                FolderAnswers fv, Txt(fv(i, FQ_ID)), Txt(fv(i, FQ_LOGIN)), agreeLogin, refused, withdrawn
                pr = PublishedRow(dq, who, kind, fromIso, toIso, Txt(fv(i, FQ_SHIFT)), Txt(fv(i, FQ_PARTNER)))
                If pr > 0 Then
                    st = Txt(dq(pr, DQ_STATUS))
                    shown = shown & "|" & pr & "|"
                ElseIf withdrawn Then
                    st = "Withdrawn"
                ElseIf refused Then
                    st = "Declined by partner"
                ElseIf StrComp(kind, "Switch", vbTextCompare) = 0 And agreeLogin = "" Then
                    st = "Waiting for partner"
                Else
                    st = "Sent - waiting for the roster"
                End If
                If st = "Waiting" And withdrawn Then st = "Withdrawn"           ' withdrawn after the roster read it: the roster catches up at its next open
                nMine = nMine + 1
                ws.Cells(r, 1).Value = kind
                If IsoDay(fromIso, dF) Then ws.Cells(r, 2).Value = dF: ws.Cells(r, 2).NumberFormat = "ddd mmm d"
                If IsoDay(toIso, dTo) Then ws.Cells(r, 3).Value = dTo: ws.Cells(r, 3).NumberFormat = "ddd mmm d"
                ws.Cells(r, 4).Value = Txt(fv(i, FQ_SHIFT))
                ws.Cells(r, 5).Value = Txt(fv(i, FQ_PARTNER))
                ws.Cells(r, 6).Value = st
                If pr > 0 Then
                    ws.Cells(r, 7).Value = Txt(dq(pr, DQ_ANSWER))
                    If Dt(dq(pr, DQ_DECIDED)) > 0 Then ws.Cells(r, 8).Value = Dt(dq(pr, DQ_DECIDED)): ws.Cells(r, 8).NumberFormat = "ddd mmm d"
                End If
                ws.Rows(r).RowHeight = 19
                If st = "Waiting" Or st = "Waiting for partner" Or st = "Sent - waiting for the roster" Then
                    nBtn = nBtn + 1
                    Btn ws, "mb_rqb" & nBtn, "Withdraw", "'MemberWithdrawRequest """ & Txt(fv(i, FQ_ID)) & """'", _
                        ws.Cells(r, 9).Left + 2, ws.Cells(r, 9).Top + 2, 72, 15, LIGHT, INK
                End If
                r = r + 1
            End If
        Next i
    End If
    If IsArray(dq) Then                               ' the published rows no folder file matched (their files are archived)
        For i = 2 To UBound(dq, 1)
            If StrComp(Txt(dq(i, DQ_EMP)), who, vbTextCompare) = 0 And InStr(shown, "|" & i & "|") = 0 Then
                nMine = nMine + 1
                ws.Cells(r, 1).Value = Txt(dq(i, DQ_KIND))
                If Dt(dq(i, DQ_FROM)) > 0 Then ws.Cells(r, 2).Value = Dt(dq(i, DQ_FROM)): ws.Cells(r, 2).NumberFormat = "ddd mmm d"
                If Dt(dq(i, DQ_TO)) > 0 Then ws.Cells(r, 3).Value = Dt(dq(i, DQ_TO)): ws.Cells(r, 3).NumberFormat = "ddd mmm d"
                ws.Cells(r, 4).Value = Txt(dq(i, DQ_SHIFT))
                ws.Cells(r, 5).Value = Txt(dq(i, DQ_PARTNER))
                ws.Cells(r, 6).Value = Txt(dq(i, DQ_STATUS))
                ws.Cells(r, 7).Value = Txt(dq(i, DQ_ANSWER))
                If Dt(dq(i, DQ_DECIDED)) > 0 Then ws.Cells(r, 8).Value = Dt(dq(i, DQ_DECIDED)): ws.Cells(r, 8).NumberFormat = "ddd mmm d"
                ws.Rows(r).RowHeight = 19
                r = r + 1
            End If
        Next i
    End If
    If nMine = 0 Then
        ws.Cells(r, 1).Value = "No requests yet. Fill in the cells above and click SEND REQUEST."
        ws.Cells(r, 1).Font.Italic = True
        r = r + 1
    End If
    ws.Cells(r + 1, 1).Value = "A request books nothing by itself: an admin books it through the roster's own rules, or declines it with a reason. " & _
                               "Statuses: Sent / Waiting for partner / Waiting (for an admin) / Booked / Declined / Declined by partner / Withdrawn / Expired / Booking voided (the admin took the booking back - ask them). Your note is read by the admin (the shared Requests folder, like Signups)."
    ws.Cells(r + 1, 1).Font.Size = 8: ws.Cells(r + 1, 1).Font.Italic = True: ws.Cells(r + 1, 1).Font.Color = RGB(100, 100, 100)
Finish:
    If Err.Number <> 0 Then mViewErr = "the Requests: " & Err.Description
    On Error Resume Next
    ProtectWs ws
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub

' "yyyy-mm-dd" -> a date (digits in every place, no roll-over); False = not one
Private Function IsoDay(ByVal s As String, ByRef d As Date) As Boolean
    s = Trim$(s): d = 0
    If Len(s) <> 10 Then Exit Function
    If Not (s Like "####-##-##") Then Exit Function
    Dim m As Long, dd As Long: m = CLng(Mid$(s, 6, 2)): dd = CLng(Right$(s, 2))
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    d = DateSerial(CLng(Left$(s, 4)), m, dd)
    IsoDay = (Day(d) = dd)
End Function

' A typed day: "2026-11-03", "11/3" or "11/3/2026" (month first), "Nov 3", "Nov 3 2026", "Nov 3, 2026". A date with no
' year lands in the year that puts it inside today - 90 .. today + 275 (the roster's own typed-date rule). No CDate.
Private Function ReadTypedDay(ByVal s As String, ByRef d As Date) As Boolean
    s = Trim$(Replace(s, ",", " ")): d = 0
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If s = "" Then Exit Function
    If IsoDay(s, d) Then ReadTypedDay = True: Exit Function
    Dim m As Long, dd As Long, y As Long, parts As Variant, mons As String, p As Long
    mons = "janfebmaraprmayjunjulaugsepoctnovdec"
    If InStr(s, "/") > 0 Then
        parts = Split(s, "/")
        If UBound(parts) < 1 Or UBound(parts) > 2 Then Exit Function
        If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then Exit Function
        m = CLng(Val(parts(0))): dd = CLng(Val(parts(1)))
        If UBound(parts) = 2 Then
            If Not IsNumeric(parts(2)) Then Exit Function
            y = CLng(Val(parts(2))): If y < 100 Then y = y + 2000
        End If
    Else
        parts = Split(s, " ")
        If UBound(parts) < 1 Or UBound(parts) > 2 Then Exit Function
        p = MonthPos(CStr(parts(0)))
        If p = 0 Then
            p = MonthPos(CStr(parts(1)))                           ' "3 Nov"
            If p = 0 Or Not IsNumeric(parts(0)) Then Exit Function
            dd = CLng(Val(parts(0)))
        Else
            If Not IsNumeric(parts(1)) Then Exit Function
            dd = CLng(Val(parts(1)))
        End If
        m = (p - 1) \ 3 + 1
        If UBound(parts) = 2 Then
            If Not IsNumeric(parts(2)) Then Exit Function
            y = CLng(Val(parts(2))): If y < 100 Then y = y + 2000
        End If
    End If
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    If y = 0 Then
        y = Year(Date)
        d = DateSerial(y, m, dd)
        If d < Date - 90 Then d = DateSerial(y + 1, m, dd)
        If d > Date + 275 Then d = DateSerial(y - 1, m, dd)
    Else
        d = DateSerial(y, m, dd)
    End If
    ReadTypedDay = (Day(d) = dd)
End Function

' a month name's place in "janfeb..." (1, 4, 7 ...; 0 = not a month) - on a 3-letter boundary, so "Rap" is not March
Private Function MonthPos(ByVal word As String) As Long
    If Len(word) < 3 Then Exit Function
    Dim p As Long: p = InStr("janfebmaraprmayjunjulaugsepoctnovdec", LCase$(Left$(word, 3)))
    If p > 0 Then
        If (p - 1) Mod 3 = 0 Then MonthPos = p
    End If
End Function

' the roster's leave words, as the master screens them (LeaveWordsIn): a note holding one is never sent
Private Function LeaveWordIn(ByVal s As String) As String
    Dim up As String: up = " " & UCase$(s) & " "
    Dim ch As Long, i As Long
    For i = 1 To Len(up)
        ch = Asc(Mid$(up, i, 1))
        If Not ((ch >= 65 And ch <= 90) Or (ch >= 48 And ch <= 57)) Then Mid$(up, i, 1) = " "
    Next i
    Dim w As Variant
    For Each w In Array("STD", "LOA", "WSIB", "LTD", "OLV", "PAR", "PARENTAL", "DISABILITY", "SICK", "PERSONAL")
        If InStr(up, " " & CStr(w) & " ") > 0 Then LeaveWordIn = CStr(w): Exit Function
    Next w
    If InStr(up, " LEAVE OF ABSENCE ") > 0 Then LeaveWordIn = "LEAVE OF ABSENCE"
End Function

' the folder, reachable from this PC ("" = not): the one message when it is not
Private Function ReachableRequestFolder() As String
    Dim folder As String: folder = RequestFolderPath()
    Dim ok As Boolean
    On Error Resume Next
    If folder <> "" Then ok = (Len(Dir$(folder, vbDirectory)) > 0)
    On Error GoTo 0
    If ok Then ReachableRequestFolder = folder
End Function

Public Sub MemberSendRequest()
    If LayoutProblem() <> "" Or DataProblem() <> "" Then MsgBox "This file did not load fully (see the note on Home), so nothing was sent. Ask your supervisor the usual way.", vbExclamation, RQ_TTL: Exit Sub
    Dim who As String: who = WhoAmI()
    If who = "" Then MsgBox "Pick your name in Find me on the Home sheet first.", vbInformation, RQ_TTL: Exit Sub
    Dim rv As Variant: rv = Data("dRoster")
    If RosterRowOf(rv, who) = 0 Then
        MsgBox "'" & who & "' isn't on the published roster, so nothing was sent. Pick your name from the Find me list on the Home sheet.", vbExclamation, RQ_TTL
        Exit Sub
    End If
    If Not RequestsOn() Then MsgBox "Requests from this file aren't switched on. Ask your supervisor the usual way.", vbInformation, RQ_TTL: Exit Sub
    Dim ws As Worksheet: Set ws = WsOf(SH_REQ)
    Dim kind As String: kind = Trim$(Txt(ws.Range(REQ_KIND).Value))
    If kind <> "Time off" And kind <> "Switch" Then MsgBox "Pick the Kind: Time off or Switch.", vbExclamation, RQ_TTL: Exit Sub
    Dim dF As Date, dTo As Date
    If Not ReadTypedDay(Txt(ws.Range(REQ_FROM).Value), dF) Then MsgBox "Type the From date (Nov 3, Nov 3 2026 or 2026-11-03).", vbExclamation, RQ_TTL: Exit Sub
    If Trim$(Txt(ws.Range(REQ_TO).Value)) = "" Then
        dTo = dF
    ElseIf Not ReadTypedDay(Txt(ws.Range(REQ_TO).Value), dTo) Then
        MsgBox "Type the To date, or leave it blank for the one day.", vbExclamation, RQ_TTL: Exit Sub
    End If
    If dTo < dF Then MsgBox "The To date is before the From date.", vbExclamation, RQ_TTL: Exit Sub
    If dF < Date Then MsgBox "The From date is in the past.", vbExclamation, RQ_TTL: Exit Sub
    Dim shiftLbl As String, partner As String
    If kind = "Switch" Then
        shiftLbl = Trim$(Txt(ws.Range(REQ_SHIFT).Value))
        partner = Trim$(Txt(ws.Range(REQ_PARTNER).Value))
        If shiftLbl <> "Days" And shiftLbl <> "Nights" Then MsgBox "A switch is one shift: pick Days or Nights.", vbExclamation, RQ_TTL: Exit Sub
        If dTo <> dF Then MsgBox "A switch is one shift: leave the To date blank.", vbExclamation, RQ_TTL: Exit Sub
        If partner = "" Then MsgBox "Pick the partner who would work the shift.", vbExclamation, RQ_TTL: Exit Sub
        If RosterRowOf(rv, partner) = 0 Then MsgBox "'" & partner & "' isn't on the published roster.", vbExclamation, RQ_TTL: Exit Sub
        If StrComp(partner, who, vbTextCompare) = 0 Then MsgBox "The partner is you.", vbExclamation, RQ_TTL: Exit Sub
    End If
    Dim note As String: note = Trim$(Txt(ws.Range(REQ_NOTE).Value))
    Dim lw As String: lw = LeaveWordIn(note)
    If lw <> "" Then MsgBox "The note holds '" & lw & "'. The roster never shows leave words to anyone - word it another way (or leave the note out).", vbExclamation, RQ_TTL: Exit Sub
    Dim folder As String: folder = ReachableRequestFolder()
    If folder = "" Then
        MsgBox "The Requests folder can't be reached from this PC, so nothing was sent. Ask your supervisor the usual way.", vbExclamation, RQ_TTL
        Exit Sub
    End If
    Dim what As String
    If kind = "Switch" Then
        what = "Shift switch " & Format$(dF, "dddd mmm d, yyyy") & " " & shiftLbl & "  -  " & partner & " would work your shift" & vbNewLine & _
               "(" & partner & " must click Agree on their PC before the roster sees it)"
    Else
        what = "Time off " & Format$(dF, "dddd mmm d, yyyy") & IIf(dTo > dF, " to " & Format$(dTo, "dddd mmm d, yyyy"), "")
    End If
    If MsgBox("Send this request for " & who & "?" & vbNewLine & vbNewLine & what & IIf(note <> "", vbNewLine & "Note: " & note, "") & vbNewLine & vbNewLine & _
              "It goes to the supervisors' roster with this PC's login (" & MemberLogin() & "). An admin books or declines it.", _
              vbOKCancel + vbQuestion, RQ_TTL) <> vbOK Then Exit Sub
    Dim stamp As Date: stamp = Now
    Dim login As String: login = SafeName(MemberLogin())
    Dim rid As String: rid = login & "-" & Format$(stamp, "yyyymmddhhnnss")
    If WriteRequestFile(folder, rid, "request", who, kind, dF, dTo, shiftLbl, partner, note, stamp) Then
        UnprotectWs ws
        ws.Range(REQ_KIND).Value = "": ws.Range(REQ_FROM).Value = "": ws.Range(REQ_TO).Value = ""
        ws.Range(REQ_SHIFT).Value = "": ws.Range(REQ_PARTNER).Value = "": ws.Range(REQ_NOTE).Value = ""
        ProtectWs ws
        MemberShowRequests
        MsgBox "Request sent." & vbNewLine & vbNewLine & who & vbNewLine & what & vbNewLine & Format$(stamp, "ddd mmm d, h:nn AM/PM") & vbNewLine & vbNewLine & _
               IIf(kind = "Switch", "It reaches the roster once " & partner & " has clicked Agree. ", "") & _
               "The answer shows under My requests after the roster's next update.", vbInformation, RQ_TTL
    End If
End Sub

' the partner's Agree / No, the requester's Withdraw (buttons on the Requests page)
Public Sub MemberAgreeRequest(ByVal rid As String)
    AnswerRequest rid, "agree"
End Sub

Public Sub MemberRefuseRequest(ByVal rid As String)
    AnswerRequest rid, "refuse"
End Sub

Public Sub MemberWithdrawRequest(ByVal rid As String)
    AnswerRequest rid, "withdraw"
End Sub

Private Sub AnswerRequest(ByVal rid As String, ByVal action As String)
    Dim who As String: who = WhoAmI()
    If who = "" Then MsgBox "Pick your name in Find me on the Home sheet first.", vbInformation, RQ_TTL: Exit Sub
    Dim fv As Variant: fv = ReadRequestFolder()
    Dim i As Long, k As Long
    If IsArray(fv) Then
        For i = 1 To UBound(fv, 1)
            If StrComp(Txt(fv(i, FQ_ID)), rid, vbTextCompare) = 0 And Txt(fv(i, FQ_ACTION)) = "request" Then k = i: Exit For
        Next i
    End If
    If k = 0 Then MsgBox "That request is no longer in the folder (the roster has taken it, or it was withdrawn). Nothing was sent.", vbInformation, RQ_TTL: Exit Sub
    Dim emp As String: emp = Txt(fv(k, FQ_EMP))
    Dim dF As Date: IsoDay Txt(fv(k, FQ_FROM)), dF
    Dim what As String: what = Txt(fv(k, FQ_KIND)) & " " & Format$(dF, "dddd mmm d, yyyy") & " " & Txt(fv(k, FQ_SHIFT))
    Dim askTxt As String
    Select Case action
        Case "agree", "refuse"
            If StrComp(Txt(fv(k, FQ_PARTNER)), who, vbTextCompare) <> 0 Then MsgBox "This switch was asked of " & Txt(fv(k, FQ_PARTNER)) & ", not of " & who & ".", vbExclamation, RQ_TTL: Exit Sub
            If StrComp(Txt(fv(k, FQ_LOGIN)), MemberLogin(), vbTextCompare) = 0 Then
                MsgBox "The same PC login sent this request, so its Agree would not count. " & Txt(fv(k, FQ_PARTNER)) & " agrees from their own PC.", vbExclamation, RQ_TTL
                Exit Sub
            End If
            askTxt = IIf(action = "agree", "Agree to work " & emp & "'s shift?" & vbNewLine & what & vbNewLine & vbNewLine & "The admin still checks the rules (the 16 h rule, bookings) and saves the switch.", _
                      "Say no to " & emp & "'s switch?" & vbNewLine & what)
        Case "withdraw"
            If StrComp(emp, who, vbTextCompare) <> 0 Then MsgBox "Only " & emp & " can withdraw this request.", vbExclamation, RQ_TTL: Exit Sub
            If StrComp(Txt(fv(k, FQ_LOGIN)), MemberLogin(), vbTextCompare) <> 0 Then
                MsgBox "This request was sent from the PC login " & Txt(fv(k, FQ_LOGIN)) & "; withdraw it from that login (the roster counts no other).", vbExclamation, RQ_TTL
                Exit Sub
            End If
            askTxt = "Withdraw your request?" & vbNewLine & what & vbNewLine & vbNewLine & "If the roster has already booked it, tell your supervisor."
    End Select
    If MsgBox(askTxt & vbNewLine & vbNewLine & "Sent with this PC's login (" & MemberLogin() & ").", vbOKCancel + vbQuestion, RQ_TTL) <> vbOK Then Exit Sub
    Dim folder As String: folder = ReachableRequestFolder()
    If folder = "" Then MsgBox "The Requests folder can't be reached from this PC, so nothing was sent.", vbExclamation, RQ_TTL: Exit Sub
    Dim dTo As Date: IsoDay Txt(fv(k, FQ_TO)), dTo
    If WriteRequestFile(folder, rid, action, who, Txt(fv(k, FQ_KIND)), dF, dTo, Txt(fv(k, FQ_SHIFT)), Txt(fv(k, FQ_PARTNER)), "", Now) Then
        MemberShowRequests
        MsgBox Choose(InStr("agree  refuse withdraw", action) \ 7 + 1, "Agreed - the request now goes to the admin.", "Sent: you said no. The requester sees it under My requests.", "Withdrawn.") & vbNewLine & vbNewLine & what, vbInformation, RQ_TTL
    End If
End Sub

' one new file; never an existing one. The master's modRequests reads these lines (the contract test ties the keys).
Private Function WriteRequestFile(ByVal folder As String, ByVal rid As String, ByVal action As String, ByVal who As String, _
                                  ByVal kind As String, ByVal dF As Date, ByVal dTo As Date, ByVal shiftLbl As String, _
                                  ByVal partner As String, ByVal note As String, ByVal stamp As Date) As Boolean
    Dim login As String: login = SafeName(MemberLogin())
    Dim pc As String: pc = SafeName(Environ$("COMPUTERNAME"))
    Dim path As String, n As Long, ff As Integer, em As String
    path = folder & "REQ__" & rid & "__" & action & "__" & login & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & ".txt"
    Do While Len(Dir$(path)) > 0 And n < 20
        n = n + 1
        path = folder & "REQ__" & rid & "__" & action & "__" & login & "__" & pc & "__" & Format$(stamp, "yyyymmdd-hhnnss") & "-" & n & ".txt"
    Loop
    On Error GoTo Failed
    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "request=" & rid
    Print #ff, "action=" & action
    Print #ff, "kind=" & kind
    Print #ff, "employee=" & who
    Print #ff, "from=" & Format$(dF, "yyyy-mm-dd")
    Print #ff, "to=" & Format$(IIf(dTo > 0, dTo, dF), "yyyy-mm-dd")
    Print #ff, "shift=" & shiftLbl
    Print #ff, "partner=" & partner
    Print #ff, "note=" & Replace(Replace(note, vbCr, " "), vbLf, " ")
    Print #ff, "login=" & MemberLogin()
    Print #ff, "machine=" & Environ$("COMPUTERNAME")
    Print #ff, "stamp=" & Format$(stamp, "yyyy-mm-dd hh:nn:ss")
    Close #ff
    On Error GoTo 0
    WriteRequestFile = True
    Exit Function
Failed:
    em = Err.Description
    On Error Resume Next
    Close #ff
    MsgBox "The request could NOT be sent (" & em & ")." & vbNewLine & vbNewLine & "Ask your supervisor the usual way.", vbExclamation, RQ_TTL
End Function

' ====================================================================
'  ACT AS (v29.0.2, the alpha only - removed in the release after the alpha)
' ====================================================================
' The login this file sends as: the stored chair while the alpha runs, else the PC's login. Never raises.
' v29.0.3 (review 1.5): the chair counts only for a login on MEM_ACTAS_USERS (the admins running the test)
Public Function MemberLogin() As String
    On Error Resume Next
    MemberLogin = Environ$("USERNAME")
    If Not MemberMayActAs() Then Exit Function
    Dim nmObj As Name
    Set nmObj = ThisWorkbook.Names(ACTAS_NAME)
    If nmObj Is Nothing Then Exit Function
    If Int(Date) > Int(DateSerial(CLng(Left$(MEM_ALPHA_END_ISO, 4)), CLng(Mid$(MEM_ALPHA_END_ISO, 6, 2)), CLng(Mid$(MEM_ALPHA_END_ISO, 9, 2)))) Then Exit Function
    Dim s As String: s = Trim$(Replace(Replace(nmObj.RefersTo, "=", ""), """", ""))
    If s <> "" Then MemberLogin = s
End Function

' Alt+F8: the password, then any login name (letters, digits, . _ -), e.g. tester2 - so one tester can send a switch request as
' one member and Agree as another (the Agree must come from a login other than the request's)
' v29.0.3 (review 1.5): True for a Windows login on MEM_ACTAS_USERS (lower case, spaces dropped)
Private Function MemberMayActAs() As Boolean
    Dim r As String: r = LCase$(Replace(Environ$("USERNAME"), " ", ""))
    If r = "" Then Exit Function
    MemberMayActAs = (InStr(1, "," & LCase$(Replace(MEM_ACTAS_USERS, " ", "")) & ",", "," & r & ",", vbBinaryCompare) > 0)
End Function

Public Sub MemberActAs()
    On Error GoTo Failed
    If Not MemberMayActAs() Then MsgBox "Act as is for the roster's admins' own logins only (yours is " & Environ$("USERNAME") & ").", vbExclamation, "Act as (alpha test)": Exit Sub
    Dim typed As String: typed = InputBox("Act as another member login for the alpha test. Password:", "Act as (alpha test)")
    If typed = "" Then Exit Sub
    If StrComp(typed, MEM_CHAIR_PASSWORD, vbBinaryCompare) <> 0 Then MsgBox "Wrong password. Nothing changed.", vbExclamation, "Act as (alpha test)": Exit Sub
    Dim chair As String: chair = Trim$(InputBox("Login to send as (letters, digits, . _ - only; e.g. tester2; blank = cancel):", "Act as (alpha test)"))
    If chair = "" Then Exit Sub
    If StrComp(chair, SafeName(chair), vbBinaryCompare) <> 0 Then MsgBox "Letters, digits, . _ - only.", vbExclamation, "Act as (alpha test)": Exit Sub
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = ThisWorkbook.Names(ACTAS_NAME)
    On Error GoTo Failed
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=ACTAS_NAME, RefersTo:="=""" & chair & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & chair & """"
        nmObj.Visible = False
    End If
    MsgBox "This file now sends sign-ups and requests as " & chair & " (the PC's login is " & Environ$("USERNAME") & "). " & _
           "MemberActAsClear puts it back; a freshly published members' file starts with no chair.", vbInformation, "Act as (alpha test)"
    Exit Sub
Failed:
    MsgBox "Act as stopped: " & Err.Description, vbExclamation, "Act as (alpha test)"
End Sub

Public Sub MemberActAsClear()
    On Error Resume Next
    ThisWorkbook.Names(ACTAS_NAME).Delete
    On Error GoTo 0
    MsgBox "This file sends as the PC's login again (" & Environ$("USERNAME") & ").", vbInformation, "Act as (alpha test)"
End Sub
