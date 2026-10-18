Attribute VB_Name = "Call_Out_Roster"
Option Explicit

' ====================================================================
'  PRINTED CALL OUT ROSTER (Blank Call Out Roster sheet) - v21
'  The Call Out Roster button (Call_Out) redraws the sheet from scratch:
'  one section per unit, everyone with a code in that unit, in call
'  order (the same LoadEligible the call-out form uses), with
'     #  |  Hours  |  Name  |  Crew  |  Phone
'  then offers to print it. Nothing on the OT Roster is sorted,
'  filtered, hidden or copied - the values are read and written.
'  Hours marked * include accepted OT not yet in Kronos (provisional).
'  The phone column is hidden for non-admins by ApplyView
'  (CALLOUT_PHONE_COL in modOvertimeEntry).
' ====================================================================

Public Const CO_VERSION As String = "Call out roster v25.3 (Oct 17)"

Private Const NAVY_C As Long = 6567966        ' RGB(30, 56, 100)
Private Const HDR_ROWS As Long = 4            ' title block repeated on every printed page

' Button: draw, show, offer to print
Sub Call_Out()
    If Not RequireSupervisor() Then Exit Sub     ' v25.3: a supervisor's job (Graeme, Oct 17)
    If Not SheetExists(CALLOUT_SHEET) Then
        MsgBox "Sheet '" & CALLOUT_SHEET & "' not found. Add a blank sheet with that name and try again.", vbCritical
        Exit Sub
    End If
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(CALLOUT_SHEET)
    Dim nPeople As Long, nUnits As Long
    Application.ScreenUpdating = False
    On Error GoTo Failed
    RenderCalloutRoster ws, nPeople, nUnits
    On Error GoTo 0
    SetSheetVisible ws, xlSheetVisible
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ws.Range("A1").Select
    On Error GoTo 0
    Application.ScreenUpdating = True

    Select Case MsgBox("Call out roster drawn: " & nPeople & " people in " & nUnits & " units, in call order." & vbNewLine & vbNewLine & _
                       "Print it now?" & vbNewLine & "(No = leave it on the tab to look at or print later.)", _
                       vbYesNo + vbQuestion, "Call Out Roster")
        Case vbYes
            Dim em As String
            On Error Resume Next
            ws.PrintOut Copies:=1
            If Err.Number <> 0 Then em = Err.Description
            On Error GoTo 0
            If em <> "" Then
                MsgBox "Couldn't print: " & em & vbNewLine & "The roster is on the " & CALLOUT_SHEET & " tab - use File > Print.", vbExclamation
            End If
    End Select
    Exit Sub

Failed:
    Dim msg As String: msg = Err.Description
    LogError "Call_Out", Err.Number, msg      ' v24.3.1
    On Error Resume Next
    ProtectSheet ws
    Application.ScreenUpdating = True
    MsgBox "The call out roster couldn't be drawn: " & msg & vbNewLine & "Send me a screenshot of this message.", vbCritical
End Sub


' --------------------------------------------------------------------
'  DRAW (sheet is wiped and redrawn every time)
' --------------------------------------------------------------------
Private Sub RenderCalloutRoster(ByVal ws As Worksheet, ByRef nPeople As Long, ByRef nUnits As Long)
    Dim lightFill As Long: lightFill = RGB(242, 245, 249)
    Dim greyTxt As Long: greyTxt = RGB(90, 90, 90)
    Dim lineCol As Long: lineCol = RGB(200, 206, 215)
    Dim i As Long, r As Long

    Application.Calculate
    UnlockSheet ws
    Application.DisplayAlerts = False

    ' ---- wipe: old sections, buttons, merges, formats ----
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Cells.UnMerge
    ws.ResetAllPageBreaks
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    ws.Cells.VerticalAlignment = xlCenter
    ws.Rows.RowHeight = 15
    ws.Columns.ColumnWidth = 8.43
    ws.Columns(1).ColumnWidth = 4.5       ' #
    ws.Columns(2).ColumnWidth = 8         ' hours
    ws.Columns(3).ColumnWidth = 30        ' name
    ws.Columns(4).ColumnWidth = 7         ' crew
    ws.Columns(CALLOUT_PHONE_COL).ColumnWidth = 15   ' phone (hidden for non-admins)
    ws.Columns(CALLOUT_PHONE_COL).Hidden = Not gAdminUnlocked

    ' ---- title block (rows 1-4, repeated on every page) ----
    Dim ps As Date: ps = PeriodStartFor(Date)
    ws.Rows(1).RowHeight = 30
    With ws.Range("A1:E1")
        .Merge
        .Value = "OVERTIME CALL OUT ROSTER"
        .Font.Size = 18: .Font.Bold = True: .Font.Color = NAVY_C
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(2).RowHeight = 16
    With ws.Range("A2:E2")
        .Merge
        .Value = "Pay period " & Format$(ps, "mmm d") & " - " & Format$(ps + 13, "mmm d, yyyy") & "     |     " & _
                 KronosStatusText() & "     |     printed " & Format$(Now, "ddd mmm d, yyyy h:mm AM/PM")
        .Font.Size = 9: .Font.Color = greyTxt
        .HorizontalAlignment = xlLeft
    End With
    ws.Rows(3).RowHeight = 14
    With ws.Range("A3:E3")
        .Merge
        .Value = "Listed in call order: grade code, then total hours, then plant seniority.  " & _
                 "* = total includes accepted overtime not yet in Kronos (data through " & _
                 Format$(EnsureDataThrough(), "mmm d") & ").  >> = above grade " & CHARGE_MAX_GRADE & " (called, not charged)." & _
                 IIf(HasPlannedMoves(), "  ~ = hours after a planned unit move's top-up.", "")
        .Font.Size = 8: .Font.Italic = True: .Font.Color = greyTxt
        .HorizontalAlignment = xlLeft
    End With
    With ws.Range("A4:E4")
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThick
        .Borders(xlEdgeBottom).Color = NAVY_C
    End With
    ws.Rows(4).RowHeight = 4

    ' ---- one section per unit ----
    r = HDR_ROWS + 2
    Dim depts As Variant: depts = DeptList()
    Dim recs() As EmpRec, n As Long, d As Long
    For d = 0 To UBound(depts)
        n = LoadEligible(CStr(depts(d)), "ALL", recs, 999, True, Date)     ' v24: today's crew
        nUnits = nUnits + 1
        nPeople = nPeople + n
        r = WriteSection(ws, r, DeptRosterTitle(CStr(depts(d))), recs, n, lightFill, greyTxt, lineCol)
    Next d

    ' ---- page ----
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(r, 5)).Address
        .PrintTitleRows = "$1:$" & HDR_ROWS
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintGridlines = False
        .LeftHeader = "": .CenterHeader = "": .RightHeader = ""
        .LeftFooter = "": .RightFooter = ""
        .CenterFooter = "Page &P of &N"
    End With
    Application.PrintCommunication = True
    ws.DisplayPageBreaks = False
    On Error GoTo 0

    Application.DisplayAlerts = True
    ProtectSheet ws
End Sub

' One unit: title bar, column headings, one row per person. Returns the next free row.
Private Function WriteSection(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String, _
                              recs() As EmpRec, ByVal n As Long, ByVal lightFill As Long, _
                              ByVal greyTxt As Long, ByVal lineCol As Long) As Long
    Dim i As Long
    ' a unit never starts in the last few lines of a page: let Excel break before it
    ws.Rows(r).RowHeight = 21
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Merge
        .Value = UCase$(title) & "   -   " & n & IIf(n = 1, " person", " people")
        .Interior.Color = NAVY_C
        .Font.Color = vbWhite: .Font.Bold = True: .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    r = r + 1
    ws.Rows(r).RowHeight = 16
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("#", "Hours", "Name", "Crew", "Phone")
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
        .Interior.Color = lightFill
        .Font.Bold = True: .Font.Size = 9: .Font.Color = greyTxt
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = lineCol
    End With
    ws.Cells(r, 3).HorizontalAlignment = xlLeft
    r = r + 1

    If n = 0 Then
        ws.Rows(r).RowHeight = 18
        With ws.Cells(r, 3)
            .Value = "nobody has a code in this unit"
            .Font.Italic = True: .Font.Color = greyTxt
        End With
        WriteSection = r + 2
        Exit Function
    End If

    Dim v() As Variant: ReDim v(1 To n, 1 To 5)
    For i = 1 To n
        v(i, 1) = i
        If recs(i).ProvHrs > 0 Or recs(i).EstHrs Then          ' v24.3: ~ = hours after a planned unit move's top-up
            v(i, 2) = Format$(recs(i).TotalHrs, "#,##0") & IIf(recs(i).ProvHrs > 0, "*", "") & IIf(recs(i).EstHrs, "~", "")
        Else
            v(i, 2) = recs(i).TotalHrs
        End If
        v(i, 3) = IIf(recs(i).Grade > CHARGE_MAX_GRADE, ">> ", "") & recs(i).EmpName & "  [" & recs(i).DeptCode & "]"
        v(i, 4) = recs(i).EmpShift
        v(i, 5) = IIf(recs(i).Phone = "", "-", recs(i).Phone)
    Next i
    Dim first As Long: first = r
    Dim last As Long: last = r + n - 1
    With ws.Range(ws.Cells(first, 1), ws.Cells(last, 5))
        .NumberFormat = "@"
        .Value = v
        .RowHeight = 18
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
        .Borders(xlInsideHorizontal).Color = lineCol
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = lineCol
    End With
    ws.Range(ws.Cells(first, 2), ws.Cells(last, 2)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(first, 2), ws.Cells(last, 2)).Value = ws.Range(ws.Cells(first, 2), ws.Cells(last, 2)).Value
    ws.Range(ws.Cells(first, 1), ws.Cells(last, 1)).Font.Color = greyTxt
    ws.Range(ws.Cells(first, 1), ws.Cells(last, 2)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(first, 4), ws.Cells(last, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(first, 3), ws.Cells(last, 3)).Font.Size = 11
    For i = 1 To n
        If recs(i).Grade > CHARGE_MAX_GRADE Then
            With ws.Range(ws.Cells(first + i - 1, 1), ws.Cells(first + i - 1, 5)).Font
                .Color = RGB(130, 135, 145): .Italic = True
            End With
        End If
    Next i
    WriteSection = last + 2
End Function
