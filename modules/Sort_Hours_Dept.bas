Attribute VB_Name = "Sort_Hours_Dept"
Option Explicit

' ====================================================================
'  ROSTER SORT BUTTONS
'  One routine does every department. The eight subs below only exist
'  because the existing buttons call them by name.
'  Order: grade code > Total Hours (G) > Plant Seniority (J)
' ====================================================================

Sub Derivatives(): Call SortForDept("Derivatives"): End Sub
Sub Phosphine(): Call SortForDept("Phosphine"): End Sub
Sub Cypure(): Call SortForDept("Cypure"): End Sub
Sub Mat_Handler(): Call SortForDept("Material Handler"): End Sub
Sub QC_Lab(): Call SortForDept("QC Derivatives"): End Sub
Sub QC_Gas(): Call SortForDept("QC Gases"): End Sub
Sub Utilities(): Call SortForDept("Utilities"): End Sub
Sub SBP(): Call SortForDept("SBP"): End Sub


Public Sub SortForDept(ByVal dept As String)
    On Error GoTo Oops                         ' v24.3.1: the eight sort buttons are everyone's; a sort that raised left the roster unprotected
    Dim ws As Worksheet
    Dim dc As Long: dc = DeptCol(dept)
    If dc = 0 Then MsgBox "Department not mapped: " & dept, vbExclamation: Exit Sub

    Set ws = RosterSheet()
    Application.ScreenUpdating = False
    UnlockSheet ws
    EnsureRosterFilter ws                      ' keep the filter; just show everyone for the sort
    If ws.FilterMode Then ws.ShowAllData
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow < R_FIRST Then ProtectSheet ws: Application.ScreenUpdating = True: Exit Sub

    SortRoster ws, lastRow, dc
    ws.AutoFilter.Range.AutoFilter Field:=RosterFilterField(ws, dc), Criteria1:="<>"
    SetRosterStatus dept

    ProtectSheet ws
    ws.Activate
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    Exit Sub
Oops:
    ReportError "SortForDept"                  ' first after the label (any On Error clears Err); it re-protects the roster and restores the screen itself
End Sub

' Show All: everyone, A - Z
Sub ShowAllRows()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = RosterSheet()
    Application.ScreenUpdating = False
    UnlockSheet ws
    EnsureRosterFilter ws
    If ws.FilterMode Then ws.ShowAllData
    Dim lastRow As Long: lastRow = RosterLastRow()
    If lastRow >= R_FIRST Then SortRoster ws, lastRow, 0
    SetRosterStatus "All (A-Z)"
    ProtectSheet ws
    ws.Activate
    ActiveWindow.ScrollRow = 1
    Application.ScreenUpdating = True
    Exit Sub
Oops:
    ReportError "ShowAllRows"
End Sub

' kept so any old button pointing at it still works
Sub Alphabetical()
    ShowAllRows
End Sub

' dc = 0 -> by name; otherwise grade code > total hours > seniority
Private Sub SortRoster(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal dc As Long)
    With ws.Sort
        .SortFields.Clear
        If dc = 0 Then
            .SortFields.Add Key:=ws.Range(ws.Cells(R_FIRST, C_NAME), ws.Cells(lastRow, C_NAME)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Else
            .SortFields.Add Key:=ws.Range(ws.Cells(R_FIRST, dc), ws.Cells(lastRow, dc)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=ws.Range(ws.Cells(R_FIRST, C_TOTAL), ws.Cells(lastRow, C_TOTAL)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=ws.Range(ws.Cells(R_FIRST, C_SENIORITY), ws.Cells(lastRow, C_SENIORITY)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End If
        .SetRange ws.Range(ws.Cells(R_HDR, 1), ws.Cells(lastRow, C_LAST))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub
