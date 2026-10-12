Attribute VB_Name = "modCleanup"
Option Explicit

' ====================================================================
'  ONE-TIME CLEANUP: removes sheets that are no longer used.
'    - DATA validation  (dropdown lists now live in the code)
'    - Refusal Report   (retired - the roster / Member View replace it)
'  Any dropdowns or named ranges that pointed at them are removed first
'  so nothing is left showing #REF!.
' ====================================================================
Public Sub RemoveOldSheets()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub

    Dim targets As Variant: targets = Array("DATA validation")      ' v24.1: "Refusal Report" is the new report's sheet now
    Dim present As String, nm As Variant
    For Each nm In targets
        If SheetExists(CStr(nm)) Then present = present & "  - " & nm & vbNewLine
    Next nm
    If present = "" Then MsgBox "Nothing to remove - those sheets are already gone.", vbInformation: Exit Sub

    If MsgBox("Delete these sheets?" & vbNewLine & vbNewLine & present & vbNewLine & _
              "Dropdowns / names that point at them are removed too. This can't be undone " & _
              "(save a backup copy first).", vbOKCancel + vbExclamation, "Remove old sheets") <> vbOK Then Exit Sub

    Application.ScreenUpdating = False
    Dim nVal As Long, nNames As Long, nSheets As Long

    ' ---- dropdowns on other sheets that use those sheets' lists ----
    Dim ws As Worksheet, vr As Range, c As Range, f As String
    For Each ws In ThisWorkbook.Worksheets
        If Not IsTarget(ws.Name, targets) Then
            Set vr = Nothing
            On Error Resume Next
            Set vr = ws.Cells.SpecialCells(xlCellTypeAllValidation)
            On Error GoTo 0
            If Not vr Is Nothing Then
                ws.Unprotect Password:=SHEET_PASSWORD
                For Each c In vr.Cells
                    f = ""
                    On Error Resume Next
                    f = c.Validation.Formula1
                    On Error GoTo 0
                    If PointsAtTarget(f, targets) Then c.Validation.Delete: nVal = nVal + 1
                Next c
                ProtectSheet ws
            End If
        End If
    Next ws

    ' ---- named ranges that point at them ----
    Dim i As Long, rt As String
    For i = ThisWorkbook.Names.Count To 1 Step -1
        rt = ""
        On Error Resume Next
        rt = ThisWorkbook.Names(i).RefersTo
        On Error GoTo 0
        If PointsAtTarget(rt, targets) Then
            On Error Resume Next
            ThisWorkbook.Names(i).Delete
            If Err.Number = 0 Then nNames = nNames + 1
            On Error GoTo 0
        End If
    Next i

    ' ---- the sheets ----
    Dim keep As Worksheet: Set keep = RosterSheet()
    keep.Activate
    Application.DisplayAlerts = False
    For Each nm In targets
        If SheetExists(CStr(nm)) Then
            With ThisWorkbook.Worksheets(CStr(nm))
                .Unprotect Password:=SHEET_PASSWORD
                .Visible = xlSheetVisible
                .Delete
            End With
            nSheets = nSheets + 1
        End If
    Next nm
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    SaveIfWritable
    MsgBox "Removed " & nSheets & " sheet(s)" & _
           IIf(nVal > 0, ", " & nVal & " old dropdown cell(s)", "") & _
           IIf(nNames > 0, ", " & nNames & " named range(s)", "") & "." & vbNewLine & vbNewLine & _
           "Run ApplyRosterLayout to refresh the buttons.", vbInformation
End Sub

Private Function IsTarget(ByVal sheetName As String, ByVal targets As Variant) As Boolean
    Dim t As Variant
    For Each t In targets
        If StrComp(sheetName, CStr(t), vbTextCompare) = 0 Then IsTarget = True: Exit Function
    Next t
End Function

Private Function PointsAtTarget(ByVal f As String, ByVal targets As Variant) As Boolean
    Dim t As Variant
    If f = "" Then Exit Function
    For Each t In targets
        If InStr(1, f, CStr(t), vbTextCompare) > 0 Then PointsAtTarget = True: Exit Function
    Next t
End Function
