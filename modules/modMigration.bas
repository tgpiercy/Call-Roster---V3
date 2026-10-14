Attribute VB_Name = "modMigration"
Option Explicit

' ====================================================================
'  ONE-TIME MIGRATION - run once after importing the new code.
'  Safe to delete this module afterwards.
'
'  1. Saves a backup copy next to the workbook
'  2. Creates the "OT Log" sheet / tblOTLog table
'  3. Moves every typed pay-period value on the roster into the log as
'     "Opening Balance" (cell notes become the reason), replaces the
'     cells with SUMIFS formulas, and verifies all totals are unchanged
'  4. Tidies up: fixes column K, clears junk in A:B, removes the unused
'     sheets and broken names, fixes typos, updates the posting dropdown
'  5. Adds Add Adjustment / Void / Publish / Refresh buttons
' ====================================================================

Public Sub RunMigration()
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If SheetExists(LOG_SHEET) Then
        MsgBox "This workbook already has an '" & LOG_SHEET & "' sheet - migration has already been run.", vbInformation
        Exit Sub
    End If

    If MsgBox("This converts the roster so the OT Log becomes the source of truth." & vbNewLine & vbNewLine & _
              "- A backup copy is saved first" & vbNewLine & _
              "- Every typed pay-period value moves into the log as an Opening Balance" & vbNewLine & _
              "- Pay-period cells become formulas over the log" & vbNewLine & _
              "- Every employee's total is checked before anything is saved" & vbNewLine & _
              "- Unused sheets (Posted Roster, empty Notes/Callout Log/Posted OT Log) are removed" & vbNewLine & vbNewLine & _
              "Continue?", vbOKCancel + vbQuestion, "Migrate to OT Log") <> vbOK Then Exit Sub

    ' ---- 1. backup ----
    Dim bk As String
    bk = EnsureSlash(ThisWorkbook.Path) & "Backup before migration " & _
         Format$(Now, "yyyy-mm-dd hhmm") & " - " & ThisWorkbook.Name
    On Error Resume Next
    ThisWorkbook.SaveCopyAs bk
    If Err.Number <> 0 Then
        Dim e As String: e = Err.Description
        On Error GoTo 0
        MsgBox "Couldn't save the backup copy, so nothing was changed:" & vbNewLine & e, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Dim calcMode As XlCalculation: calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo Fail

    UnprotectAll
    ClearRosterFilter RosterSheet()

    ' ---- 2. log ----
    CreateLogSheet

    ' ---- 3. move values, write formulas, verify ----
    Dim summary As String
    If Not ConvertRosterToFormulas("Opening Balance", "OB", summary) Then GoTo Stopped

    ' ---- 4. tidy ----
    Dim notes As String
    notes = CleanupRoster()
    notes = notes & CleanupSheetsAndNames()
    FixTyposAndLists

    ' ---- 5. buttons ----
    AddNewButtons

    Application.Calculation = calcMode
    Application.Calculate
    ProtectAll
    ApplyView IsAdmin()
    RosterSheet().Activate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ThisWorkbook.Save

    MsgBox "Migration complete." & vbNewLine & vbNewLine & summary & vbNewLine & vbNewLine & _
           notes & "Backup saved as:" & vbNewLine & bk, vbInformation, "Migration complete"
    Exit Sub

Stopped:
    MarkNoSave
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Migration stopped before saving." & vbNewLine & _
           "Close this workbook WITHOUT saving (saving is blocked). Your backup is at:" & vbNewLine & bk, vbCritical
    Exit Sub

Fail:
    Dim msg As String: msg = Err.Description
    MarkNoSave
    On Error Resume Next
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Migration error: " & msg & vbNewLine & vbNewLine & _
           "Close this workbook WITHOUT saving. Your backup is at:" & vbNewLine & bk, vbCritical
End Sub


Private Sub UnprotectAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        UnlockSheet ws
        On Error GoTo 0
    Next ws
End Sub


Private Sub CreateLogSheet()
    Dim ws As Worksheet
    UnlockStructure                                 ' v25: structure protection
    Set ws = ThisWorkbook.Worksheets.Add(After:=RosterSheet())
    ws.Name = LOG_SHEET
    RelockStructure
    ws.Range("A1").Resize(1, L_NCOLS).Value = LogHeaders()

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, L_NCOLS), , xlYes)
    lo.Name = LOG_TABLE
    lo.TableStyle = "TableStyleMedium2"

    Dim widths As Variant
    widths = Array(10, 10, 11, 14, 8, 7, 18, 7, 7, 24, 9, 10, 14, 36, 10, 11, 8, 11, 18, 8, 30)
    Dim i As Long
    For i = 1 To L_NCOLS
        ws.Columns(i).ColumnWidth = widths(i - 1)
    Next i
    ws.Columns(L_DATE).NumberFormat = "mm/dd/yyyy"
    ws.Columns(L_PERIOD).NumberFormat = "mm/dd/yyyy"
    ws.Columns(L_STAMP).NumberFormat = "mm/dd/yyyy hh:mm AM/PM"
    ws.Rows(1).WrapText = True

    ' Conditional formatting. CF formulas are relative to the ACTIVE cell,
    ' so select row 2 first.
    ws.Activate
    ws.Range("A2").Select
    Dim oc As String: oc = ColLetter(L_OUTCOME)
    Dim outRng As Range: Set outRng = ws.Range(oc & "2:" & oc & "50000")
    AddValueCF outRng, "Accepted", RGB(198, 239, 206), RGB(0, 97, 0)
    AddValueCF outRng, "Awarded", RGB(198, 239, 206), RGB(0, 97, 0)
    AddValueCF outRng, "Refused", RGB(255, 199, 206), RGB(156, 0, 6)
    AddValueCF outRng, "Left Message", RGB(255, 235, 156), RGB(156, 101, 0)
    AddValueCF outRng, "Signed Up", RGB(255, 235, 156), RGB(156, 101, 0)
    AddValueCF outRng, "Unavailable", RGB(217, 225, 242), RGB(50, 65, 90)

    Dim fc As FormatCondition
    Set fc = ws.Range("A2:" & ColLetter(L_NCOLS) & "50000").FormatConditions.Add( _
                Type:=xlExpression, Formula1:="=$" & ColLetter(L_STATUS) & "2=""VOID""")
    fc.Font.Strikethrough = True
    fc.Font.Color = RGB(150, 150, 150)

    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Tab.Color = RGB(31, 73, 125)
End Sub

Private Sub AddValueCF(ByVal rng As Range, ByVal txt As String, ByVal fill As Long, ByVal fontCol As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & txt & """")
    fc.Interior.Color = fill
    fc.Font.Color = fontCol
End Sub


' Returns notes for the summary message
Private Function CleanupRoster() As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim lastRow As Long: lastRow = RosterLastRow()
    Dim r As Long, c As Long, msg As String, stray As String

    ' A:B only held stray copies (=D, =C) on a few rows
    ws.Range(ws.Cells(R_FIRST, 1), ws.Cells(lastRow, 2)).ClearContents
    If Trim$(SafeStr(ws.Range("C1").Value)) = "`" Then ws.Range("C1").ClearContents

    ' rows below the last named employee that still hold numbers
    For r = lastRow + 1 To lastRow + 50
        For c = C_WORKED To C_PERIODN
            If c = C_WORKED Or c >= C_PERIOD1 Then
                If Not IsEmpty(ws.Cells(r, c).Value) Then
                    stray = stray & r & " "
                    Exit For
                End If
            End If
        Next c
    Next r
    If stray <> "" Then
        msg = "CHECK: roster row(s) " & Trim$(stray) & " have hours but no employee name. " & _
              "They were left untouched and are NOT in the log." & vbNewLine & vbNewLine
    End If
    CleanupRoster = msg
End Function


Private Function CleanupSheetsAndNames() As String
    Dim msg As String
    Application.DisplayAlerts = False

    If SheetExists("Posted Roster") Then
        SetSheetVisible ThisWorkbook.Worksheets("Posted Roster"), xlSheetVisible
        UnlockStructure
        ThisWorkbook.Worksheets("Posted Roster").Delete
        RelockStructure
        msg = msg & "Removed: Posted Roster" & vbNewLine
    End If

    If SheetExists("Notes") Then
        If Application.WorksheetFunction.CountA(ThisWorkbook.Worksheets("Notes").UsedRange) = 0 Then
            SetSheetVisible ThisWorkbook.Worksheets("Notes"), xlSheetVisible
            UnlockStructure
            ThisWorkbook.Worksheets("Notes").Delete
            RelockStructure
            msg = msg & "Removed: Notes (empty)" & vbNewLine
        End If
    End If

    Dim nm As Variant, ws As Worksheet
    For Each nm In Array("Callout Log", "Posted OT Log")
        If SheetExists(CStr(nm)) Then
            Set ws = ThisWorkbook.Worksheets(CStr(nm))
            If Application.WorksheetFunction.CountA(ws.Range("A2:Z100000")) = 0 Then
                SetSheetVisible ws, xlSheetVisible
                UnlockStructure
                ws.Delete
                RelockStructure
                msg = msg & "Removed: " & nm & " (empty, replaced by OT Log)" & vbNewLine
            Else
                UnlockStructure
                ws.Name = "Legacy " & CStr(nm)
                RelockStructure
                msg = msg & "Kept as 'Legacy " & nm & "' (it had rows)" & vbNewLine
            End If
        End If
    Next nm

    If SheetExists("Entry Log") Then
        UnlockStructure
        ThisWorkbook.Worksheets("Entry Log").Name = "Legacy Entry Log"
        RelockStructure
        SetSheetVisible ThisWorkbook.Worksheets("Legacy Entry Log"), xlSheetVeryHidden
        msg = msg & "Renamed: Entry Log -> Legacy Entry Log (still hidden, history kept)" & vbNewLine
    End If

    Application.DisplayAlerts = True

    ' broken defined names (#REF!)
    Dim i As Long, sRef As String, nmText As String
    On Error Resume Next
    For i = ThisWorkbook.Names.Count To 1 Step -1
        sRef = "": nmText = ""
        sRef = ThisWorkbook.Names(i).RefersTo
        nmText = ThisWorkbook.Names(i).Name
        If InStr(1, sRef, "#REF!", vbTextCompare) > 0 Then
            ThisWorkbook.Names(i).Delete
            msg = msg & "Removed broken name: " & nmText & vbNewLine
        End If
    Next i
    On Error GoTo 0

    If msg <> "" Then msg = msg & vbNewLine
    CleanupSheetsAndNames = msg
End Function


Private Sub FixTyposAndLists()
    Dim nm As Variant
    For Each nm In Array(POSTING_SHEET)
        If SheetExists(CStr(nm)) Then
            ThisWorkbook.Worksheets(CStr(nm)).Cells.Replace What:="Classificaiton", _
                Replacement:="Classification", LookAt:=xlPart, MatchCase:=False
        End If
    Next nm
    If SheetExists(CALLOUT_SHEET) Then
        ThisWorkbook.Worksheets(CALLOUT_SHEET).Cells.Replace What:="Deivatives", _
            Replacement:="Derivatives", LookAt:=xlPart, MatchCase:=False
    End If
End Sub


Private Sub AddNewButtons()
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim x As Double: x = ws.Range(ColLetter(C_PERIOD1) & "1").Left + 6
    Dim y As Double: y = ws.Range("A1").Top + 10
    AddBtn ws, "Add Adjustment", "AddAdjustment", x, y
    AddBtn ws, "Void Entry", "VoidEntry", x, y + 34
    AddBtn ws, "Publish Member View", "PublishMemberView", x + 122, y
    AddBtn ws, "Refresh Roster Formulas", "RefreshRosterFormulas", x + 122, y + 34
End Sub

Private Sub AddBtn(ByVal ws As Worksheet, ByVal cap As String, ByVal mac As String, _
                   ByVal l As Double, ByVal t As Double)
    Dim b As Object
    Set b = ws.Buttons.Add(l, t, 116, 28)
    b.Caption = cap
    b.OnAction = mac
End Sub
