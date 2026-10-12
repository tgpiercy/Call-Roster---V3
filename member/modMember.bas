Attribute VB_Name = "modMember"
Option Explicit

' ====================================================================
'  CALL OUT ROSTER - MEMBERS' WORKBOOK (M1)
'  This module lives in the members' file, NOT in the master. It holds
'  no rules: everything it shows was worked out by the master and
'  pushed here as values (the very hidden sheets dMeta, dRoster, dNext,
'  dCal, dPost, dPostList, dLog). The code only sorts, filters and
'  looks those values up.
'
'  BuildMemberTemplate (Alt+F8, once, in a blank .xlsm with this module
'  and the ThisWorkbook code): lays out every sheet and button.
'  MemberRefreshAll: redraws Home and the Roster from the data (the
'  master runs it after each publish; it also runs on open).
'  Everything else is a button or a picker cell.
' ====================================================================

Public Const MEM_VERSION As String = "Members M1 (Oct 12)"
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
Private Const RC_FIXED As Long = 7

' picker cells (unlocked; a change redraws the view)
Private Const ROSTER_FIND As String = "C2"
Private Const ROSTER_CREW As String = "F2"
Private Const NEXT_UNIT As String = "B2"
Private Const NEXT_DATE As String = "D2"
Private Const NEXT_SHIFT As String = "F2"
Private Const CAL_MONTH As String = "B2"
Private Const CAL_UNIT As String = "D2"
Private Const HIST_NAME As String = "B2"

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
    For Each nm In Array("dMeta", "dRoster", "dNext", "dCal", "dPost", "dPostList", "dLog", SH_LISTS)
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
    Picker ws, "C14", "", "Find me:"
    ws.Range("C14").ColumnWidth = 26
    Btn ws, "mb_find", "Show me on the roster", "MemberFindFromHome", ws.Range("D14").Left + 4, ws.Range("D14").Top - 2, 150, 20, LIGHT, INK
    With ws.Range("B17")
        .Value = "Hours and call order are worked out by the supervisors' roster and published here as values. " & _
                 "Sign-up for postings is still on the paper sheet. Leave shows only as 'Leave'."
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
    Picker ws, ROSTER_FIND, "", "Find me:"
    Picker ws, ROSTER_CREW, "", "Crew:"
    Dim x As Double: x = ws.Range("H1").Left
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
    Title ws, "COVERAGE CALENDAR", "Pick a month and a unit. Each day shows who is off (and who covers) on days and nights. " & _
                                   "OPEN = still needs cover. Leave shows only as 'Leave'."
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 52
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 52
    Picker ws, CAL_MONTH, "", "Month:"
    Picker ws, CAL_UNIT, "", "Unit:"
    ProtectWs ws
End Sub

Private Sub LayoutPostings()
    Dim ws As Worksheet: Set ws = WsOf(SH_POST)
    UnprotectWs ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = INK
    Title ws, "OPEN POSTINGS", "Postings still open for sign-up, with the people listed on each (in call order). Sign up on the paper sheet."
    ws.Columns("A").ColumnWidth = 10
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 34
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
    On Error GoTo Done
    BuildLists
    DrawHome
    MemberShowAll
    DrawNextDefault
    DrawCalendarDefault
    MemberShowPostings
    DrawHistory ""
Done:
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

' picker lists on dLists: A names, B units (+ All units), C shifts, D next dates, E months
Private Sub BuildLists()
    Dim ws As Worksheet: Set ws = WsOf(SH_LISTS, True)
    UnprotectWs ws
    ws.Cells.ClearContents
    ws.Columns("A:E").NumberFormat = "@"           ' lists are text: "Oct 2026" must not become a date
    Dim v As Variant, i As Long, nNames As Long, nUnits As Long, nMonths As Long
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
    SetList WsOf(SH_CAL), CAL_UNIT, ListRef("B", IIf(nUnits > 0, nUnits + 1, 0))
    SetList WsOf(SH_HIST), HIST_NAME, ListRef("A", nNames)
    SetList WsOf(SH_HOME), "C14", ListRef("A", nNames)
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
    ws.Range(addr).Locked = False
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

Public Sub MemberFindMe()
    Dim ws As Worksheet: Set ws = WsOf(SH_ROSTER)
    Dim nm As String: nm = Txt(ws.Range(ROSTER_FIND).Value)
    If nm = "" Then MsgBox "Pick your name in the Find me box first.", vbInformation: Exit Sub
    DrawRoster "", "A - Z", nm
End Sub

Public Sub MemberFindFromHome()
    Dim nm As String: nm = Txt(WsOf(SH_HOME).Range("C14").Value)
    If nm = "" Then MsgBox "Pick your name first.", vbInformation: Exit Sub
    UnprotectWs WsOf(SH_ROSTER)
    WsOf(SH_ROSTER).Range(ROSTER_FIND).Value = nm
    ProtectWs WsOf(SH_ROSTER)
    DrawRoster "", "A - Z", nm
    WsOf(SH_ROSTER).Activate
End Sub

Private Sub DrawRoster(ByVal unit As String, ByVal viewTxt As String, Optional ByVal highlight As String = "")
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
    ws.Rows(4).RowHeight = 30
    x = ws.Range("A4").Left: y = ws.Range("A4").Top + 4
    If UBound(units) >= LBound(units) Then
        For i = 0 To nU - 1
            Btn ws, "mb_u" & i, Txt(units(i)), "'MemberSortUnit """ & Txt(units(i)) & """'", x + i * 92, y, 88, 22, _
                IIf(Txt(units(i)) = unit, NAVY, LIGHT), IIf(Txt(units(i)) = unit, vbWhite, INK)
        Next i
    End If

    Dim heads As Variant
    If unit = "" Then
        heads = Array("Name", "Crew", "Home unit", "Total", "Worked", "Refused")
    Else
        heads = Array("#", "Name", "Crew", "Code", "Total", "Worked", "Refused")
    End If
    Header ws, 5, heads
    ws.Range("A3").Value = "View: " & viewTxt & IIf(crew <> "", "  |  crew " & crew, "") & _
                           ".  * = total includes accepted overtime not yet in Kronos."

    If Not IsArray(v) Then ProtectWs ws: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
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
                ElseIf uCode > 0 And Txt(v(i, uCode)) <> "" Then
                    n = n + 1: idx(n) = i: keys(n) = Num(v(i, uPos))
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
        ProtectWs ws: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    End If
    ReDim out(1 To n, 1 To nc)
    Dim r As Long, hiRow As Long
    For r = 1 To n
        i = idx(r)
        If unit = "" Then
            out(r, 1) = Txt(v(i, RC_NAME)): out(r, 2) = Txt(v(i, RC_CREW)): out(r, 3) = Txt(v(i, RC_HOME))
            out(r, 4) = HoursText(v(i, RC_TOTAL), v(i, RC_PROV)): out(r, 5) = Num(v(i, RC_WORKED)): out(r, 6) = Num(v(i, RC_REFUSED))
        Else
            out(r, 1) = keys(r): out(r, 2) = Txt(v(i, RC_NAME)): out(r, 3) = Txt(v(i, RC_CREW)): out(r, 4) = Txt(v(i, uCode))
            out(r, 5) = HoursText(v(i, RC_TOTAL), v(i, RC_PROV)): out(r, 6) = Num(v(i, RC_WORKED)): out(r, 7) = Num(v(i, RC_REFUSED))
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
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  CALENDAR
' ====================================================================
Private Sub DrawCalendarDefault()
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    Dim u As Variant: u = UnitList()
    UnprotectWs ws
    ws.Range(CAL_MONTH).NumberFormat = "@"
    ws.Range(CAL_MONTH).Value = Format$(Dt(Meta("Published at")), "mmm yyyy")
    If Txt(ws.Range(CAL_UNIT).Value) = "" And UBound(u) >= LBound(u) Then ws.Range(CAL_UNIT).Value = Txt(u(LBound(u)))
    ProtectWs ws
    MemberShowCalendar
End Sub

' the master's TypeColor values, copied (no rule here, just the same colours)
Private Function LineColour(ByVal code As String) As Long
    Select Case UCase$(code)
        Case "VAC": LineColour = RGB(21, 101, 192)
        Case "LV", "LOA", "STD", "PAR": LineColour = RGB(106, 27, 154)
        Case "TRN": LineColour = RGB(0, 121, 107)
        Case "PHA": LineColour = RGB(46, 125, 50)
        Case "SUP": LineColour = RGB(93, 64, 55)
        Case "SS": LineColour = RGB(194, 24, 91)
        Case "OPEN": LineColour = RGB(190, 30, 30)
        Case Else: LineColour = RGB(97, 97, 97)
    End Select
End Function

' colours each line of a multi-line cell by its type code ("VAC|LV|OPEN" ...)
Private Sub ColourLines(ByVal c As Range, ByVal cellTxt As String, ByVal types As String)
    If cellTxt = "" Then Exit Sub
    Dim lines As Variant: lines = Split(cellTxt, Chr$(10))
    Dim codes As Variant: codes = Split(types, "|")
    Dim i As Long, p As Long: p = 1
    For i = LBound(lines) To UBound(lines)
        If i <= UBound(codes) Then
            With c.Characters(p, Len(lines(i))).Font
                .Color = LineColour(CStr(codes(i)))
                .Bold = (UCase$(CStr(codes(i))) = "OPEN")
            End With
        End If
        p = p + Len(lines(i)) + 1
    Next i
End Sub

Public Sub MemberShowCalendar()
    Dim ws As Worksheet: Set ws = WsOf(SH_CAL)
    Dim mTxt As String: mTxt = Txt(ws.Range(CAL_MONTH).Value)
    Dim unit As String: unit = Txt(ws.Range(CAL_UNIT).Value)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    ClearBelow ws, 4
    Header ws, 5, Array("Date", "Day", "DAYS  (who's off  ->  covered by)", "", "NIGHTS  (who's off  ->  covered by)")
    ws.Range("A3").Value = "Month: " & mTxt & "   |   " & unit & "   |   OPEN = still needs cover.  Leave shows as 'Leave'."
    Dim v As Variant: v = Data("dCal")
    If Not IsArray(v) Or mTxt = "" Then ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    ' the month: "Oct 2026" as text, or a real date if Excel converted the picker
    Dim m0 As Date
    If VarType(ws.Range(CAL_MONTH).Value) = vbDate Then
        m0 = DateSerial(Year(ws.Range(CAL_MONTH).Value), Month(ws.Range(CAL_MONTH).Value), 1)
    Else
        On Error Resume Next
        m0 = CDate("1 " & mTxt)
        If m0 = 0 Then m0 = CDate(mTxt)
        On Error GoTo 0
        If m0 <> 0 Then m0 = DateSerial(Year(m0), Month(m0), 1)
    End If
    If m0 = 0 Then ProtectWs ws: MarkClean: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
    Dim m1 As Date: m1 = DateSerial(Year(m0), Month(m0) + 1, 1) - 1
    Dim allUnits As Boolean: allUnits = (StrComp(unit, "All units", vbTextCompare) = 0)
    Dim d As Date, r As Long, i As Long, dayTxt As String, nightTxt As String, hol As String
    Dim dayTypes As String, nightTypes As String, line As String, pre As String, code As String
    r = 6
    For d = m0 To m1
        dayTxt = "": nightTxt = "": hol = "": dayTypes = "": nightTypes = ""
        For i = 2 To UBound(v, 1)
            If Int(Dt(v(i, 1))) = Int(d) Then
                If Txt(v(i, 3)) = "*" Then
                    hol = Txt(v(i, 4))
                ElseIf allUnits Or StrComp(Txt(v(i, 3)), unit, vbTextCompare) = 0 Then
                    pre = IIf(allUnits, Txt(v(i, 3)) & ": ", "")
                    line = pre & Txt(v(i, 4)) & IIf(Txt(v(i, 5)) <> "", "  ->  " & Txt(v(i, 5)), "  ->  OPEN")
                    code = IIf(CBool(v(i, 7)), "OPEN", Txt(v(i, 6)))
                    If Txt(v(i, 2)) = "Nights" Then
                        nightTxt = nightTxt & IIf(nightTxt = "", "", Chr$(10)) & line
                        nightTypes = nightTypes & IIf(nightTypes = "", "", "|") & code
                    Else
                        dayTxt = dayTxt & IIf(dayTxt = "", "", Chr$(10)) & line
                        dayTypes = dayTypes & IIf(dayTypes = "", "", "|") & code
                    End If
                End If
            End If
        Next i
        ws.Cells(r, 1).Value = d
        ws.Cells(r, 1).NumberFormat = "mmm d"
        ws.Cells(r, 2).Value = Format$(d, "dddd") & IIf(hol <> "", Chr$(10) & hol, "")
        ws.Cells(r, 3).Value = dayTxt
        ws.Cells(r, 5).Value = nightTxt
        With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
            .WrapText = True
            .VerticalAlignment = xlTop
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Color = RGB(200, 206, 215)
        End With
        ws.Cells(r, 3).Interior.Color = RGB(255, 249, 196)        ' days: yellow
        ws.Cells(r, 5).Interior.Color = RGB(227, 236, 250)        ' nights: blue
        ColourLines ws.Cells(r, 3), dayTxt, dayTypes
        ColourLines ws.Cells(r, 5), nightTxt, nightTypes
        If Weekday(d, vbMonday) >= 6 Then ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(240, 242, 245)
        If hol <> "" Then ws.Cells(r, 2).Font.Color = RGB(176, 0, 32)
        ws.Rows(r).AutoFit
        If ws.Rows(r).RowHeight < 18 Then ws.Rows(r).RowHeight = 18
        r = r + 1
    Next d
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
End Sub


' ====================================================================
'  POSTINGS
' ====================================================================
Private Function CloseTime() As String        ' published by the master; a default only until the first publish
    CloseTime = Txt(Meta("Posting close time"))
    If CloseTime = "" Then CloseTime = "8:00 AM"
End Function

Public Sub MemberShowPostings()
    Dim ws As Worksheet: Set ws = WsOf(SH_POST)
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    UnprotectWs ws
    ClearBelow ws, 4
    Dim p As Variant: p = Data("dPost")
    Dim l As Variant: l = Data("dPostList")
    Dim r As Long: r = 5
    Dim i As Long, j As Long, k As Long
    If Not IsArray(p) Then
        ws.Cells(r, 1).Value = "No open postings."
    ElseIf UBound(p, 1) < 2 Then
        ws.Cells(r, 1).Value = "No open postings."
    Else
        For i = 2 To UBound(p, 1)
            ws.Rows(r).RowHeight = 22
            With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
                .Merge
                .Value = Txt(p(i, 1)) & "   " & Txt(p(i, 2)) & "   " & Format$(Dt(p(i, 3)), "dddd mmm d, yyyy") & "   " & Txt(p(i, 4)) & _
                         "   " & Num(p(i, 5)) & " h   " & Num(p(i, 6)) & " position(s)   " & Txt(p(i, 7)) & _
                         "   |   closes " & Format$(Dt(p(i, 8)), "ddd mmm d") & " " & CloseTime()
                .Interior.Color = NAVY
                .Font.Color = vbWhite: .Font.Bold = True
                .VerticalAlignment = xlCenter
            End With
            r = r + 1
            Header ws, r, Array("#", "Name", "Hours", "", "")
            r = r + 1
            k = 0
            If IsArray(l) Then
                For j = 2 To UBound(l, 1)
                    If StrComp(Txt(l(j, 1)), Txt(p(i, 1)), vbTextCompare) = 0 Then
                        ws.Cells(r, 1).Value = Num(l(j, 2))
                        ws.Cells(r, 2).Value = Txt(l(j, 3))
                        ws.Cells(r, 3).Value = Num(l(j, 4))
                        ws.Cells(r, 1).HorizontalAlignment = xlCenter
                        ws.Cells(r, 3).HorizontalAlignment = xlCenter
                        r = r + 1
                        k = k + 1
                    End If
                Next j
            End If
            If k = 0 Then ws.Cells(r, 2).Value = "(list not published)": r = r + 1
            ws.Cells(r, 2).Value = "Listed in call order with hours as of the posting date. Out-of-department sign-ups go on the blank lines of the paper sheet."
            ws.Cells(r, 2).Font.Size = 8: ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(100, 100, 100)
            r = r + 2
        Next i
    End If
    ProtectWs ws
    MarkClean
    Application.ScreenUpdating = True
    Application.EnableEvents = evWas
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
    ClearBelow ws, 4
    If nm = "" Then
        ws.Cells(5, 1).Value = "Pick a name above."
        ProtectWs ws: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
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
        ProtectWs ws: Application.ScreenUpdating = True: Application.EnableEvents = evWas: Exit Sub
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
    Dim yearTot As Double, ps As String
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
            If Not Application.Intersect(Target, Sh.Range(ROSTER_FIND)) Is Nothing Then MemberFindMe
        Case SH_NEXT
            If Not Application.Intersect(Target, Sh.Range(NEXT_UNIT & "," & NEXT_DATE & "," & NEXT_SHIFT)) Is Nothing Then MemberShowNext
        Case SH_CAL
            If Not Application.Intersect(Target, Sh.Range(CAL_MONTH & "," & CAL_UNIT)) Is Nothing Then MemberShowCalendar
        Case SH_HIST
            If Not Application.Intersect(Target, Sh.Range(HIST_NAME)) Is Nothing Then MemberShowHistory
    End Select
End Sub

Public Sub MemberSheetActivated(ByVal Sh As Object)
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
End Sub
