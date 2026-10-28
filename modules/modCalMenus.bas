Attribute VB_Name = "modCalMenus"
Option Explicit
' v27.1 (the refactor, Oct 2026): the double-click / right-click menus on the Calendar and Needs Attention, from modCalendar; moved unchanged (tests/movecheck.py proves it)
Public Const CALMENU_VERSION As String = "Calendar menus v29.0.1 (Oct 28)"

Private mMenuKey As String          ' line (or "P:P-0001") the double-click menu is for
Private Const MENU_NAME As String = "CalLineMenu"
' ====================================================================
'  MENUS  (double-click / right-click on the Calendar or Needs Attention)
'  Keys:  A:AB-0001|serial|Days   a time-off line
'         C:CV-0001|serial|Days   a need / coverage line
'         E:NEW|serial|Days       an empty spot under a date
'         D:DAY|serial|           a date header
'         P:P-0001                a posting
'         S:SS-0001               a shift switch owed a payback
'         B:AB-0001               a time-off booking
' ====================================================================
Public Sub CalDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    On Error Resume Next
    If Target.Worksheet.Name <> CAL_SHEET Then Exit Sub
    Cancel = True                                  ' never "this cell is protected"
    Dim keyTxt As String: keyTxt = SafeStr(MapSheet().Range(Target.Cells(1, 1).Address).Value)
    On Error GoTo 0
    If keyTxt <> "" Then ShowLineMenu keyTxt
End Sub

Public Sub CalRightClick(ByVal Target As Range, ByRef Cancel As Boolean)
    On Error Resume Next
    If Target.Worksheet.Name <> CAL_SHEET Then Exit Sub
    Dim keyTxt As String: keyTxt = SafeStr(MapSheet().Range(Target.Cells(1, 1).Address).Value)
    On Error GoTo 0
    If keyTxt = "" Then Exit Sub                  ' ordinary right-click elsewhere
    Cancel = True
    ShowLineMenu keyTxt
End Sub

Private Function NAKeyAt(ByVal Target As Range) As String
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet: Set ws = Target.Worksheet
    Dim r As Long: r = Target.Cells(1, 1).Row
    If r = 4 Or r = 5 Then
        ' summary tiles: find the tile's first column
        Dim c As Long
        For c = Target.Cells(1, 1).Column To 2 Step -1
            If SafeStr(ws.Cells(r, NA_KEY + c).Value) <> "" Then NAKeyAt = SafeStr(ws.Cells(r, NA_KEY + c).Value): Exit Function
        Next c
    End If
    NAKeyAt = SafeStr(ws.Cells(r, NA_KEY).Value)
End Function

Public Sub NADoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Cancel = True
    Dim keyTxt As String: keyTxt = NAKeyAt(Target)
    If keyTxt = "" Or Left$(keyTxt, 2) = "Q:" Then Exit Sub      ' (review L17) Q: lines are for reading
    If keyTxt = "Z:more" Then NAMoreToggle: Exit Sub             ' v29 (item 5): a lite login's "More (system)" heading folds / unfolds
    If Left$(keyTxt, 2) = "J:" Then
        On Error Resume Next                            ' never raises
        Application.Goto Target.Worksheet.Cells(CLng(Mid$(keyTxt, 3)), 1), True
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "H:" Then              ' holiday to confirm: open its row on the Holidays sheet
        GoToHoliday DateSerial(CLng(Mid$(keyTxt, 3, 4)), CLng(Mid$(keyTxt, 8, 2)), CLng(Mid$(keyTxt, 11, 2)))
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "U:" Then              ' v21: Home Unit blank - jump to that cell on the roster
        GoToHomeUnit Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "M:" Then              ' M1: publish the members' file
        If MsgBox("Publish the members' file now?", vbYesNo + vbQuestion, "Member View") = vbYes Then PublishMemberView
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "X:" Then              ' v24.2: errors - open the Error Log
        ShowErrorLog
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "F:" Then              ' v27 (item 9): the receipt answers - open the Feedback sheet
        ShowFeedback
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "R:" Then              ' v28 (item 1): a member's request - book it through the prefilled pop-up
        RequestAct Mid$(keyTxt, 3), False
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "V:" Then              ' v24.3: a planned unit move - cancel it, or apply it when due
        PlannedMoveMenu Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "I:" Then              ' v22: imported line - open the convert pop-up
        ShowConvertForm
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "T:" Then              ' v22.1: attendance - open the sheet
        ShowAttendance
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "N:" Then              ' v26: below minimum staffing - the calendar on that date and group
        CalShowDate NADateInKey(keyTxt), NAUnitInKey(keyTxt)
        Exit Sub
    End If
    ShowLineMenu keyTxt
End Sub

' v26: the date serial in a "N:<unit>|<serial>|<shift>" key (0 when it can't be read)
Private Function NADateInKey(ByVal keyTxt As String) As Date
    Dim parts As Variant: parts = Split(keyTxt, "|")
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(1)) Then NADateInKey = CDate(CDbl(parts(1)))
    End If
End Function

' v26: the unit in a "N:<unit>|<serial>|<shift>" key ("" when none)
Private Function NAUnitInKey(ByVal keyTxt As String) As String
    Dim parts As Variant: parts = Split(keyTxt, "|")
    If Len(CStr(parts(0))) > 2 Then NAUnitInKey = Mid$(CStr(parts(0)), 3)
End Function

' v26: the Calendar opened on the month that holds d (today's month when d is 0) and, when a unit is
' given, on a group that holds the unit; drawn once (the draw stamp Z8 is cleared when the month or the
' group moved, so CalStale is True - review: a month change alone never made it stale)
Public Sub CalShowDate(ByVal d As Date, Optional ByVal unit As String = "")
    On Error GoTo Oops
    If d = 0 Then d = Date
    CalendarSheet
    Dim m0 As Date: m0 = DateSerial(Year(d), Month(d), 1)
    Dim moved As Boolean: moved = (CalMonth() <> m0)
    MapWrite "Z1", m0
    If unit <> "" Then
        If Not UnitInGroup(unit, CalGroup()) Then MapWrite "Z2", GroupOfUnit(unit): moved = True
    End If
    If moved Then MapWrite "Z8", "moved"              ' never equal to the Z7 counter (re-review: "" equals a never-marked map)
    ShowCalendar
    Exit Sub
Oops:
    ReportError "CalShowDate"
End Sub

Public Sub NARightClick(ByVal Target As Range, ByRef Cancel As Boolean)
    Dim keyTxt As String: keyTxt = NAKeyAt(Target)
    If keyTxt = "" Or Left$(keyTxt, 2) = "J:" Or Left$(keyTxt, 2) = "Q:" Then Exit Sub
    Cancel = True
    If Left$(keyTxt, 2) = "H:" Then
        GoToHoliday DateSerial(CLng(Mid$(keyTxt, 3, 4)), CLng(Mid$(keyTxt, 8, 2)), CLng(Mid$(keyTxt, 11, 2)))
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "U:" Then
        GoToHomeUnit Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "M:" Then
        If MsgBox("Publish the members' file now?", vbYesNo + vbQuestion, "Member View") = vbYes Then PublishMemberView
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "X:" Then
        ShowErrorLog
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "F:" Then
        ShowFeedback
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "R:" Then              ' v28 (item 1): a member's request - decline it with a reason
        RequestAct Mid$(keyTxt, 3), True
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "V:" Then
        PlannedMoveMenu Mid$(keyTxt, 3)
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "I:" Then
        ShowConvertForm
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "T:" Then
        ShowAttendance
        Exit Sub
    End If
    If Left$(keyTxt, 2) = "N:" Then              ' v26
        CalShowDate NADateInKey(keyTxt), NAUnitInKey(keyTxt)
        Exit Sub
    End If
    ShowLineMenu keyTxt
End Sub


Private Sub ShowLineMenu(ByVal keyTxt As String)
    mMenuKey = keyTxt
    LoadCache
    LoadPostCache
    Dim kind As String: kind = Left$(keyTxt, 2)

    Dim title As String, lt As String, rt As String, postId As String, d As Date, sh As String
    Dim caps() As String, macs() As String, grpBreak() As Boolean, n As Long
    ReDim caps(1 To 16): ReDim macs(1 To 16): ReDim grpBreak(1 To 16)
    d = KeyDate(keyTxt)
    sh = KeyShift(keyTxt)

    ' v27 (item 15, B2): a supervisor on an OPEN line (nobody on it yet) gets the jobs that are theirs since v25.5 -
    ' fill it now (a call-out), post it, award the posting on it; anything else, and everyone else, sees the details
    If Not gAdminUnlocked Then
        If IsSupervisor() And (kind = "A:" Or kind = "C:") Then
            If SlotTexts(keyTxt, lt, rt) Then
                If rt = "" Or UCase$(rt) Like "POSTED*" Or UCase$(rt) Like "NO TAKERS*" Then
                    title = Format$(d, "ddd mmm d") & " " & sh & "  -  " & Clip(lt, 34)
                    postId = PostIdOfLine(keyTxt)
                    AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
                    ' (review) a POSTED line's posting is still open: the award fills it, not a call-out beside it
                    If Not (UCase$(rt) Like "POSTED*") Then AddChoice caps, macs, grpBreak, n, "Fill this shift now (call-out)...", "CalMenuCallout", True
                    If Not (UCase$(rt) Like "POSTED*") Then AddChoice caps, macs, grpBreak, n, "Post this line...", "CalMenuPost", False
                    If postId <> "" Then
                        If PostRow(postId) > 0 Then
                            If SafeStr(mPostV(PostRow(postId), PS_STATUS)) = "Open" Then AddChoice caps, macs, grpBreak, n, "Award posting " & postId & "...", "CalMenuAward", (UCase$(rt) Like "POSTED*")   ' (re-review B8) the first job on a POSTED line
                        End If
                    End If
                    GoTo ShowIt
                End If
            End If
        End If
        If kind = "E:" Or kind = "D:" Then CalMenuDayDetails Else CalMenuDetails
        Exit Sub
    End If

    Select Case kind
        Case "A:", "C:"
            If Not SlotTexts(keyTxt, lt, rt) Then
                MsgBox "That line has changed - redrawing.", vbInformation, "Calendar"
                RefreshCalendar
                Exit Sub
            End If
            title = Format$(d, "ddd mmm d") & " " & sh & "  -  " & Clip(lt, 34)
            postId = PostIdOfLine(keyTxt)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If rt = "" Or UCase$(rt) Like "NO TAKERS*" Then AddChoice caps, macs, grpBreak, n, "Fill this shift now (call-out)...", "CalMenuCallout", True   ' v27 (item 15): the call-out from an open or NO TAKERS line, pre-filled (a covered or POSTED line: no)
            If Not (UCase$(rt) Like "POSTED*") Then AddChoice caps, macs, grpBreak, n, "Post this line...", "CalMenuPost", False
            AddChoice caps, macs, grpBreak, n, "Set coverage...", "CalMenuSetCoverage", (UCase$(rt) Like "POSTED*")
            If kind = "C:" Then AddChoice caps, macs, grpBreak, n, "Change the text of this line...", "CalMenuEditText", False
            If postId <> "" Then
                If PostRow(postId) > 0 Then
                    If SafeStr(mPostV(PostRow(postId), PS_STATUS)) = "Open" Then
                        AddChoice caps, macs, grpBreak, n, "Award posting " & postId & "...", "CalMenuAward", True
                        AddChoice caps, macs, grpBreak, n, "Change the Reason of posting " & postId & "...", "CalMenuEditReason", False
                    End If
                End If
                AddChoice caps, macs, grpBreak, n, "View posting " & postId, "CalMenuViewPosting", False
            End If
            If kind = "A:" Then
                AddChoice caps, macs, grpBreak, n, "Change end date of this time off...", "CalMenuEndDate", True
                AddChoice caps, macs, grpBreak, n, "Void this time off...", "CalMenuVoid", False
            End If
            AddChoice caps, macs, grpBreak, n, "Post support on this shift...", "CalMenuPostSupport", True
            AddChoice caps, macs, grpBreak, n, "Add another need on this shift...", "CalMenuAddNeed", False
            If ActiveSheet.Name = NA_SHEET Then AddChoice caps, macs, grpBreak, n, "Show on the calendar", "CalMenuGoto", True

        Case "E:"
            title = Format$(d, "ddd mmm d") & " " & sh & "  (" & IIf(sh = "Days", DaysCrew(d), NightsCrew(d)) & " Shift)"
            AddChoice caps, macs, grpBreak, n, "Post unit operator support...", "CalMenuPostSupport", False
            AddChoice caps, macs, grpBreak, n, "Post PHA coverage...", "CalMenuPostPHA", False
            AddChoice caps, macs, grpBreak, n, "Add a need (no posting yet)...", "CalMenuAddNeed", True
            AddChoice caps, macs, grpBreak, n, "Book time off from this date...", "CalMenuBookFrom", False
            AddChoice caps, macs, grpBreak, n, "What's on this day", "CalMenuDayDetails", True

        Case "D:"
            title = Format$(d, "dddd mmm d") & "  (Days " & DaysCrew(d) & ", Nights " & NightsCrew(d) & ")"
            AddChoice caps, macs, grpBreak, n, "Post support - Days...", "CalMenuPostSupportD", False
            AddChoice caps, macs, grpBreak, n, "Post support - Nights...", "CalMenuPostSupportN", False
            AddChoice caps, macs, grpBreak, n, "Add a need (no posting yet)...", "CalMenuAddNeed", True
            AddChoice caps, macs, grpBreak, n, "Book time off from this date...", "CalMenuBookFrom", False
            AddChoice caps, macs, grpBreak, n, "What's on this day", "CalMenuDayDetails", True

        Case "P:"
            postId = Mid$(keyTxt, 3)
            title = "Posting " & postId
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If PostRow(postId) > 0 Then
                If SafeStr(mPostV(PostRow(postId), PS_STATUS)) = "Open" Then
                    AddChoice caps, macs, grpBreak, n, "Award posting " & postId & "...", "CalMenuAward", True
                    AddChoice caps, macs, grpBreak, n, "Change the Reason of posting " & postId & "...", "CalMenuEditReason", False
                End If
            End If
            AddChoice caps, macs, grpBreak, n, "View posting " & postId, "CalMenuViewPosting", False
            If Not CovHasSourceCached(postId) Then
                AddChoice caps, macs, grpBreak, n, "Put on the calendar...", "CalMenuLinkPosting", True
            End If

        Case "S:"
            title = "Shift switch " & Mid$(keyTxt, 3)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            AddChoice caps, macs, grpBreak, n, "Log the payback (Shift Switch)...", "CalMenuSwitch", True

        Case "B:"
            title = "Time off " & Mid$(keyTxt, 3)
            AddChoice caps, macs, grpBreak, n, "Details...", "CalMenuDetails", False
            If MenuAbsIsLeave() Then AddChoice caps, macs, grpBreak, n, "Return from leave now (top-up)...", "CalMenuLeaveReturn", True
            AddChoice caps, macs, grpBreak, n, "Change end date...", "CalMenuEndDate", Not MenuAbsIsLeave()
            AddChoice caps, macs, grpBreak, n, "Void this time off...", "CalMenuVoid", False

        Case Else
            Exit Sub
    End Select

ShowIt:
    ' pop-up menu at the mouse
    Dim cb As Object, ctl As Object, i As Long
    On Error Resume Next
    Application.CommandBars(MENU_NAME).Delete
    On Error GoTo Fallback
    Set cb = Application.CommandBars.Add(Name:=MENU_NAME, Position:=5, Temporary:=True)   ' 5 = msoBarPopup
    Set ctl = cb.Controls.Add(Type:=1, Temporary:=True)                                    ' 1 = msoControlButton
    ctl.Caption = title
    ctl.Enabled = False
    For i = 1 To n
        Set ctl = cb.Controls.Add(Type:=1, Temporary:=True)
        ctl.Caption = caps(i)
        ctl.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macs(i)
        ctl.BeginGroup = (i = 1) Or grpBreak(i)
    Next i
    cb.ShowPopup
    Exit Sub

Fallback:
    ' no pop-up menus (Mac / restricted Excel) - numbered list instead
    Dim msg As String, pick As Long
    msg = title & vbNewLine & vbNewLine
    For i = 1 To n
        msg = msg & i & ")  " & caps(i) & vbNewLine
    Next i
    pick = AskNumber(msg & vbNewLine & "Type the number:", "Calendar", 1, 1, n)
    If pick >= 1 Then Application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macs(pick)
End Sub

Private Sub AddChoice(ByRef caps() As String, ByRef macs() As String, ByRef grpBreak() As Boolean, _
                      ByRef n As Long, ByVal cap As String, ByVal mac As String, ByVal newGroup As Boolean)
    n = n + 1
    caps(n) = cap
    macs(n) = mac
    grpBreak(n) = newGroup
End Sub

Private Function MenuPostId() As String
    If Left$(mMenuKey, 2) = "P:" Then
        MenuPostId = Mid$(mMenuKey, 3)
    Else
        LoadCache
        MenuPostId = PostIdOfLine(mMenuKey)
    End If
End Function

Private Function MenuAbsId() As String
    If Left$(mMenuKey, 2) = "A:" Or Left$(mMenuKey, 2) = "B:" Then MenuAbsId = Mid$(Split(mMenuKey, "|")(0), 3)
End Function

' unit for a new posting / need started from the menu
Private Function MenuUnit() As String
    If Left$(mMenuKey, 2) = "A:" Or Left$(mMenuKey, 2) = "C:" Then
        LoadCache
        MenuUnit = SlotUnit(mMenuKey)
    Else
        MenuUnit = SingleUnitOf(CalGroup())
    End If
End Function

Public Sub CalMenuDetails()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    If mMenuKey = "" Then Exit Sub
    Dim t As String
    Select Case Left$(mMenuKey, 2)
        Case "P:": t = PostingDetails(Mid$(mMenuKey, 3))
        Case "S:": t = SwitchDetails(Mid$(mMenuKey, 3))
        Case "B:": t = BookingDetails(Mid$(mMenuKey, 3))
        Case Else: t = LineDetails(mMenuKey)
    End Select
    If t = "" Then t = "No details found - try Refresh."
    MsgBox Replace(t, vbLf, vbNewLine), vbInformation, "Details"
    Exit Sub
Oops:
    ReportError "CalMenuDetails"
End Sub

Public Sub CalMenuDayDetails()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim d As Date: d = KeyDate(mMenuKey)
    If d = 0 Then Exit Sub
    Dim grp As String: grp = CalGroup()
    Dim t As String, sh As Variant, slots() As CalSlot, n As Long, i As Long, rt As String
    t = Format$(d, "dddd mmm d") & "  -  " & grp
    For Each sh In Array("Days", "Nights")
        n = SlotsFor(d, CStr(sh), grp, slots)
        t = t & vbNewLine & vbNewLine & UCase$(sh) & "  (" & IIf(sh = "Days", DaysCrew(d), NightsCrew(d)) & " Shift)"
        If n = 0 Then t = t & vbNewLine & "   nothing booked"
        For i = 1 To n
            rt = slots(i).RightText
            If rt = "" Then rt = "OPEN"
            t = t & vbNewLine & "   " & slots(i).LeftText & "   ->   " & rt
        Next i
    Next sh
    MsgBox t, vbInformation, "Calendar"
    Exit Sub
Oops:
    ReportError "CalMenuDayDetails"
End Sub

Private Function SwitchDetails(ByVal id As String) As String
    Dim d As Date, sh As String, u As String, offN As String, worksN As String
    If Not SwitchInfo(id, d, sh, u, offN, worksN) Then SwitchDetails = "Shift switch " & id & " wasn't found.": Exit Function
    SwitchDetails = "Shift switch " & id & vbLf & _
                    worksN & " worked " & Format$(d, "ddd mmm d") & " " & sh & " (" & u & ") for " & offN & "." & vbLf & _
                    "Payback still owed: " & offN & " works a shift for " & worksN & "." & vbLf & vbLf & _
                    "Log it in Shift Switch - pick this switch in the owed list."
End Function

Private Function BookingDetails(ByVal id As String) As String
    LoadCache
    Dim r As Long: r = RowOf(mAbs, id)
    If r = 0 Then BookingDetails = "Time off " & id & " wasn't found.": Exit Function
    Dim code As String: code = TypeCode(SafeStr(mAbs(r, AB_TYPE)))
    Dim t As String
    t = CalName(SafeStr(mAbs(r, AB_EMP))) & "  -  " & TypeLong(code) & "  (" & SafeStr(mAbs(r, AB_UNIT)) & ")" & vbLf & _
        "From " & Format$(mAbs(r, AB_START), "ddd mmm d, yyyy") & " to " & _
        IIf(IsDate(mAbs(r, AB_END)), Format$(mAbs(r, AB_END), "ddd mmm d, yyyy"), "open-ended (no return date)")
    If HrsTag(mAbs(r, AB_HOURS)) <> "" Then t = t & vbLf & "Off" & HrsTag(mAbs(r, AB_HOURS)) & " per shift"
    If Not IsLeaveCode(code) And code <> "SIC" And SafeStr(mAbs(r, AB_NOTE)) <> "" Then t = t & vbLf & "Note: " & SafeStr(mAbs(r, AB_NOTE))
    t = t & vbLf & "Booked by " & StampTxt(mAbs(r, AB_USER), mAbs(r, AB_STAMP)) & "  [" & id & "]"
    BookingDetails = t
End Function

Public Sub CalMenuPost()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    CalPostKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuPost"
End Sub

' v27 (item 15, B2): "Fill this shift now" from a calendar line - the call-out pop-up pre-filled with the line's unit,
' date, shift and hours (a partial absence's hours; else the shift's usual, the DST night's 13 / 11), and the line's
' key so the acceptance covers THIS line without a question (item 10). Everyone who may log a call-out may use it.
Public Sub CalMenuCallout()
    On Error GoTo Oops
    If mMenuKey = "" Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Dim kind As String: kind = Left$(mMenuKey, 2)
    If kind <> "A:" And kind <> "C:" Then Exit Sub
    LoadCache
    Dim d As Date: d = KeyDate(mMenuKey)
    Dim sh As String: sh = KeyShift(mMenuKey)
    Dim unit As String: unit = SlotUnit(mMenuKey)
    If d = 0 Or sh = "" Then MsgBox "That line has changed - redrawing.", vbInformation, "Calendar": RefreshCalendar: Exit Sub
    If unit = "" Then unit = SingleUnitOf(CalGroup())
    Dim hrs As Double: hrs = DefaultShiftHours(d, IsNightShift(sh), USUAL_SHIFT_HOURS)
    Dim slotKey As String: slotKey = Split(mMenuKey, "|")(0)
    Dim r As Long
    If kind = "A:" Then
        r = RowOf(mAbs, Mid$(slotKey, 3))
        If r > 0 Then
            If NumOr0(mAbs(r, AB_HOURS)) > 0 And NumOr0(mAbs(r, AB_HOURS)) < hrs Then hrs = NumOr0(mAbs(r, AB_HOURS))   ' a partial-day absence: its hours
        End If
    Else
        r = RowOf(mCov, Mid$(slotKey, 3))                 ' (review L7) a need line's own hours, as CalPostKey reads them
        If r > 0 Then
            If NumOr0(mCov(r, CV_HOURS)) > 0 And NumOr0(mCov(r, CV_HOURS)) < hrs Then hrs = NumOr0(mCov(r, CV_HOURS))
        End If
    End If
    Load frmOvertimeEntry
    frmOvertimeEntry.Prefill unit, d, sh, "Calendar line", hrs, slotKey
    frmOvertimeEntry.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuCallout"
End Sub

Public Sub CalMenuSetCoverage()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    CalSetCoverageKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuSetCoverage"
End Sub

' v24.3.2: the words on a need line (the Reason a posting wrote, or a typed need)
Public Sub CalMenuEditText()
    On Error GoTo Oops
    CalEditTextKey mMenuKey
    Exit Sub
Oops:
    ReportError "CalMenuEditText"
End Sub

' v24.3.2: an open posting's Reason, from its line or its Needs Attention entry
Public Sub CalMenuEditReason()
    On Error GoTo Oops
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    CalEditReasonId id
    Exit Sub
Oops:
    ReportError "CalMenuEditReason"
End Sub

Public Sub CalMenuPostSupport()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", KeyShift(mMenuKey)
    Exit Sub
Oops:
    ReportError "CalMenuPostSupport"
End Sub

Public Sub CalMenuPostPHA()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "PHA", KeyShift(mMenuKey)
    Exit Sub
Oops:
    ReportError "CalMenuPostPHA"
End Sub

Public Sub CalMenuPostSupportD()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", "Days"
    Exit Sub
Oops:
    ReportError "CalMenuPostSupportD"
End Sub

Public Sub CalMenuPostSupportN()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    PostNewLine "Unit Operator Support", "Nights"
    Exit Sub
Oops:
    ReportError "CalMenuPostSupportN"
End Sub

' New Posting for a NEW calendar line (support / PHA with nobody off)
Private Sub PostNewLine(ByVal reason As String, ByVal shiftLbl As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    Dim d As Date: d = KeyDate(mMenuKey)
    If d = 0 Then Exit Sub
    If shiftLbl = "" Then shiftLbl = "Days"
    If Int(d) < Int(Date) Then
        If MsgBox("That date has passed. Post it anyway?", vbYesNo + vbQuestion, "Post") = vbNo Then Exit Sub
    End If
    gCalPostKey = "NEW"
    gCalPostDate = d
    gCalPostShift = shiftLbl
    frmNewPosting.Prefill MenuUnit(), d, shiftLbl, reason, FULL_SHIFT, Format$(d, "ddd mmm d") & " " & shiftLbl & ": " & reason
    frmNewPosting.Show
    gCalPostKey = ""
    RefreshCalendar
End Sub

Public Sub CalMenuAddNeed()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim sh As String: sh = KeyShift(mMenuKey)
    If sh = "" Then sh = "Both"
    OpenTimeOff True, KeyDate(mMenuKey), sh
    Exit Sub
Oops:
    ReportError "CalMenuAddNeed"
End Sub

Public Sub CalMenuBookFrom()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    OpenTimeOff False, KeyDate(mMenuKey), ""
    Exit Sub
Oops:
    ReportError "CalMenuBookFrom"
End Sub

' Book Time Off, already filled in with the date / shift / unit
Private Sub OpenTimeOff(ByVal needMode As Boolean, ByVal d As Date, ByVal shiftLbl As String)
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If AbsTable(True) Is Nothing Then Exit Sub
    If CovTable(True) Is Nothing Then Exit Sub
    Dim unit As String: unit = MenuUnit()
    Load frmTimeOff
    On Error Resume Next
    With frmTimeOff
        If needMode Then .Controls("to_optN").Value = True
        If d > 0 Then
            .Controls("to_txtStart").Text = Format$(d, "mm/dd/yyyy")
            If needMode Then .Controls("to_txtEnd").Text = Format$(d, "mm/dd/yyyy")
        End If
        If needMode And shiftLbl <> "" Then .Controls("to_cboShift").Value = shiftLbl
        If unit <> "" Then .Controls("to_cboUnit").Value = unit
    End With
    On Error GoTo 0
    frmTimeOff.Show
    RefreshCalendar
End Sub

Public Sub CalMenuViewPosting()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    Dim rep As String
    If Not PostingOutput(id, True, False, False, rep) Then MsgBox "Couldn't show posting " & id & ". " & rep, vbExclamation, "Call Out Roster"
    Exit Sub
Oops:
    ReportError "CalMenuViewPosting"
End Sub

' Award pop-up with this posting already picked - just click LOAD
Public Sub CalMenuAward()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireSupervisor() Then Exit Sub            ' v27 (item 15): awarding to the list is a supervisor's job (v25.5; the pop-up greys the admin buttons); was RequireAdmin
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    Dim ids As Variant, ch As Variant, i As Long, idx As Long
    idx = -1
    ch = OpenPostingChoices(ids)
    If IsArray(ids) Then
        For i = LBound(ids) To UBound(ids)
            If StrComp(CStr(ids(i)), id, vbTextCompare) = 0 Then idx = i - LBound(ids)
        Next i
    End If
    Load frmPostingEntry
    If idx >= 0 Then
        On Error Resume Next
        frmPostingEntry.Controls("cmbPosting").ListIndex = idx
        On Error GoTo 0
    End If
    frmPostingEntry.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuAward"
End Sub

' A posting that never made it onto the calendar: pick its line now
Public Sub CalMenuLinkPosting()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuPostId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim p As PostRec, lst() As PListRec
    LoadPosting id, p, lst
    If p.ID = "" Then MsgBox "Posting " & id & " wasn't found.", vbExclamation, "Call Out Roster": Exit Sub
    LoadPostCache
    Dim byTxt As String, r As Long, aw As Variant, i As Long
    Select Case p.Status
        Case "Open": byTxt = "POSTED " & p.ID
        Case "Cancelled": byTxt = "NCN"
        Case "No takers": byTxt = "No takers"
        Case Else
            r = PostRow(p.ID)
            If r > 0 Then
                aw = Split(Replace(SafeStr(mPostV(r, PS_AWARDTO)), " (out of dept)", ""), ";")
                For i = 0 To UBound(aw)
                    If Trim$(aw(i)) <> "" Then byTxt = byTxt & IIf(byTxt = "", "", " / ") & CalName(Trim$(aw(i)))
                Next i
            End If
            If byTxt = "" Then byTxt = "POSTED " & p.ID
    End Select
    Dim t As String
    t = CalendarPlace(p.ShiftDate, p.ShiftLbl, p.Unit, "Posting " & p.ID & " - which calendar line does it cover?", _
                      byTxt, p.Hours, p.ID, ReasonType(p.Reason), p.Reason, IsSupportReason(p.Reason))
    If t <> "" Then
        SaveIfWritable
        MsgBox p.ID & " is on the calendar: " & t, vbInformation, "Calendar"
    End If
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuLinkPosting"
End Sub

Public Sub CalMenuEndDate()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If SetAbsenceEnd(id) Then RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuEndDate"
End Sub

' v22: is the booking under the menu a leave (so "Return from leave" applies)?
Private Function MenuAbsIsLeave() As Boolean
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Function
    Dim r As Long: r = RowOf(mAbs, id)
    If r > 0 Then
        MenuAbsIsLeave = IsLeaveCode(SafeStr(mAbs(r, AB_TYPE)))
        If MenuAbsIsLeave And UBound(mAbs, 2) >= AB_GROUP Then
            If LeaveToppedUp(SafeStr(mAbs(r, AB_GROUP))) Then MenuAbsIsLeave = False     ' already done
        End If
    End If
End Function

Public Sub CalMenuLeaveReturn()
    On Error GoTo Oops                         ' v24.3.1: the second door into the Leave pop-up (ShowLeaveForm has one)
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not AbsenceColumnsCurrent() Then
        MsgBox "The Absences table needs its v22 upgrade first: an admin runs Alt+F8 > UpgradeAbsenceColumns once.", vbExclamation, "Time off"
        Exit Sub
    End If
    Load frmLeave
    frmLeave.StartReturn id
    frmLeave.Show
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuLeaveReturn"
End Sub

Public Sub CalMenuVoid()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim id As String: id = MenuAbsId()
    If id = "" Then Exit Sub
    If VoidAbsence(id) Then RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuVoid"
End Sub

Public Sub CalMenuSwitch()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    ShowShiftSwitchForm
    RefreshCalendar
    Exit Sub
Oops:
    ReportError "CalMenuSwitch"
End Sub

Public Sub CalMenuBook()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    ShowTimeOffForm
    Exit Sub
Oops:
    ReportError "CalMenuBook"
End Sub

' from Needs Attention: open the calendar on that month with the line selected
Public Sub CalMenuGoto()
    On Error GoTo Oops                         ' v25 (I5): every menu macro runs top-level from the pop-up menu
    Dim keyTxt As String: keyTxt = mMenuKey
    Dim d As Date: d = KeyDate(keyTxt)
    If d = 0 Then Exit Sub
    CalendarSheet
    Dim moved As Boolean: moved = (CalMonth() <> DateSerial(Year(d), Month(d), 1))   ' v26 review: a month change alone never made the calendar stale
    MapWrite "Z1", DateSerial(Year(d), Month(d), 1)
    LoadCache
    Dim u As String: u = SlotUnit(keyTxt)
    If u <> "" Then
        If Not UnitInGroup(u, CalGroup()) Then MapWrite "Z2", GroupOfUnit(u): moved = True
    End If
    If moved Then MapWrite "Z8", "moved"
    ShowCalendar
    On Error Resume Next                            ' never raises
    Dim hit As Range
    Set hit = MapSheet().Range("A1:P1000").Find(What:=keyTxt, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not hit Is Nothing Then
        Application.Goto CalendarSheet().Range(hit.Address), False
    End If
    Exit Sub
Oops:
    ReportError "CalMenuGoto"
End Sub

