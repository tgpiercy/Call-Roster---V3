Attribute VB_Name = "modPostingSheet"
Option Explicit

' ====================================================================
'  THE OVERTIME POSTING SHEET  (output only - redrawn every time)
' ====================================================================
Private Function PostingSheet() As Worksheet
    If Not SheetExists(POSTING_SHEET) Then
        Dim ws As Worksheet
        UnlockStructure                                 ' v25: structure protection
        Set ws = ThisWorkbook.Worksheets.Add(After:=RosterSheet())
        ws.Name = POSTING_SHEET
        RelockStructure
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

    UnlockSheet ws
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
            If i > 1 Then                            ' v27 (item 27): "=" tied with the row above on grade and the frozen hours
                If OrderTie(lst(i - 1).Grade, lst(i - 1).TotalHrs, lst(i).Grade, lst(i).TotalHrs) Then v(i, 1) = CStr(lst(i).Rank) & "="
            End If
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
        ' v27 (item 5, 5.3): the old sheet's layout, with where it came from
        ' (review M1) a footer holds 255 characters in all: kept short
        If Not blank Then .CenterFooter = "&8From " & p.ID & " (list frozen at posting) - printed " & Format$(Now, "mmm d h:mm AM/PM") & _
                                          vbLf & "&7Order: grade, lowest Total Hours at posting, earliest seniority. = tied on grade and hours."     ' v27 (item 27)
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


