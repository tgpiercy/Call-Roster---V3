Attribute VB_Name = "modReceipt"
Option Explicit

' ====================================================================
'  RECEIPTS (v27, item 7 - briefs/v27.md; review 5.1)
'
'  One card after every save, one layout on every form (Graeme, Oct 20):
'    WHAT WAS RECORDED   one line per row written (Ref ID, type, person, hours)
'    WHO IT CHANGED      "Joe Smith: 120 -> 141 h" from the roster's Total Hours
'                        before and after (RosterTotalsMap), or "nobody"
'    WHERE TO SEE IT     the calendar line, the OT Log rows, the sheet, the
'                        members' file after the next update
'    IF THIS IS WRONG    Undo last / Void Entry for an admin; "ask <admin>" for
'                        a supervisor
'  and the trust strip at the foot. Print receipt saves a PDF beside the
'  roster (the Receipt output sheet, redrawn each time); Email me this opens
'  a mail draft with the text. The card is frmReceipt (code-built); it reads
'  gReceipt. Under SmokeTest the text goes to gReceiptLog instead of the
'  card, so the steps can assert its lines. Holds no rule: every figure
'  comes from the save that just happened.
' ====================================================================

Public Const RCPT_VERSION As String = "Receipt v26.5.5 (Oct 21)"     ' stamped with the release at release prep (v27)
Public Const RECEIPT_SHEET As String = "Receipt"                     ' output only: drawn for a print, hidden again

Public Type Receipt
    Title As String
    Recorded As String
    Changed As String
    WhereTxt As String
    WrongTxt As String
    Trust As String
End Type

Public gReceipt As Receipt
Public gReceiptLog As String                ' SmokeTest: every receipt's text, appended (the card is not shown)

' The roster's Total Hours before a save, for ReceiptChanges (name -> "Total|Refused|periods")
Public Function ReceiptBefore() As Object
    On Error Resume Next                          ' (review M4) a receipt must never stop a save: an empty map when the roster cannot be read
    Set ReceiptBefore = CreateObject("Scripting.Dictionary")
    Set ReceiptBefore = RosterTotalsMap()
End Function

' Who the save changed: every roster name whose Total Hours moved, "name: before -> after h (+d)"; new or gone names
' said so; "Nobody's Total Hours changed." when nothing moved. Never raises (a receipt must not stop a save).
Public Function ReceiptChanges(ByVal before As Object) As String
    On Error GoTo Quiet
    Application.Calculate
    Dim after As Object: Set after = RosterTotalsMap()
    Dim k As Variant, s As String, tb As Double, ta As Double, n As Long
    For Each k In after.Keys
        ta = TotalOf(CStr(after(k)))
        If before.Exists(k) Then
            tb = TotalOf(CStr(before(k)))
            If Abs(ta - tb) > 0.001 Then
                n = n + 1
                If n <= 12 Then s = s & "  " & k & ": " & HoursText(tb) & " -> " & HoursText(ta) & " h (" & IIf(ta > tb, "+", "") & HoursText(ta - tb) & ")" & vbNewLine
            End If
        Else
            n = n + 1
            If n <= 12 Then s = s & "  " & k & ": new on the roster, " & HoursText(ta) & " h" & vbNewLine
        End If
    Next k
    For Each k In before.Keys
        If Not after.Exists(k) Then
            n = n + 1
            If n <= 12 Then s = s & "  " & k & ": no longer on the roster" & vbNewLine
        End If
    Next k
    If n > 12 Then s = s & "  ... and " & (n - 12) & " more" & vbNewLine
    If s = "" Then s = "  Nobody's Total Hours changed." & vbNewLine
    ReceiptChanges = s
    Exit Function
Quiet:
    ReceiptChanges = "  (the before / after figures could not be read: " & Err.Description & ")" & vbNewLine
End Function

Private Function TotalOf(ByVal packed As String) As Double
    Dim p As Long: p = InStr(packed, "|")
    If p > 0 Then packed = Left$(packed, p - 1)
    TotalOf = Val(Replace(packed, ",", "."))
End Function

' The "if this is wrong" line for whoever is logged in
Public Function WrongText() As String
    If IsAdmin() Then
        WrongText = "  Undo my last log entry (the Desk) voids your newest OT Log entry of today; Void Entry on the bar voids any Ref ID." & vbNewLine
    Else
        WrongText = "  Ask " & PRIMARY_ADMIN_NAME & " to void it (Void Entry); the rows stay in the log, marked VOID." & vbNewLine
    End If
End Function

' Fills the receipt (every part a block of lines, "" allowed) and shows it
Public Sub ShowReceiptNow(ByVal title As String, ByVal recorded As String, ByVal changed As String, ByVal whereTxt As String, ByVal wrongTxt As String)
    gReceipt.Title = title
    gReceipt.Recorded = recorded
    gReceipt.Changed = changed
    gReceipt.WhereTxt = whereTxt
    gReceipt.WrongTxt = wrongTxt
    ShowReceipt
End Sub

' Shows gReceipt: the card, or (SmokeTest) the text into gReceiptLog; a card that cannot open falls back to a box
Public Sub ShowReceipt()
    gReceipt.Trust = TrustStrip()
    ' (review M7) a save that failed, or saving blocked: the card says so first - it must never read as a saved entry
    If Not gSmokeRun And (gSaveFailedAt > 0 Or NoSave()) Then
        If Left$(gReceipt.Title, 9) <> "NOT SAVED" Then gReceipt.Title = "NOT SAVED - " & gReceipt.Title
        gReceipt.WhereTxt = "  " & SaveBlockText() & vbNewLine & gReceipt.WhereTxt
    End If
    If gSmokeRun Then
        gReceiptLog = gReceiptLog & vbNewLine & "---- " & ReceiptText(gReceipt)
        Exit Sub
    End If
    On Error GoTo Fallback
    frmReceipt.Show
    Exit Sub
Fallback:
    LogError "ShowReceipt", Err.Number, Err.Description
    MsgBox ReceiptText(gReceipt), vbInformation, IIf(gReceipt.Title = "", "Saved", gReceipt.Title)
End Sub

' The receipt as plain text (the print, the email, the SmokeTest log)
Public Function ReceiptText(ByRef r As Receipt) As String
    ReceiptText = r.Title & vbNewLine & _
                  Format$(Now, "ddd mmm d, yyyy h:mm AM/PM") & "  -  " & CurrentUser() & vbNewLine & vbNewLine & _
                  "WHAT WAS RECORDED" & vbNewLine & IIf(r.Recorded = "", "  (nothing written to the OT Log)" & vbNewLine, r.Recorded) & vbNewLine & _
                  "WHO IT CHANGED" & vbNewLine & IIf(r.Changed = "", "  Nobody's Total Hours changed." & vbNewLine, r.Changed) & vbNewLine & _
                  "WHERE TO SEE IT" & vbNewLine & r.WhereTxt & vbNewLine & _
                  "IF THIS IS WRONG" & vbNewLine & r.WrongTxt & vbNewLine & _
                  r.Trust
End Function

' One line per OT Log row written: "  S-0012  Refused  Joe Smith  12 h  charged 21" (data as AppendLogRows takes it)
Public Function RowsText(ByRef data As Variant, ByVal nRows As Long) As String
    Dim i As Long, s As String
    For i = 1 To nRows
        s = s & "  " & SafeStr(data(i, L_REF)) & "  " & SafeStr(data(i, L_OUTCOME)) & IIf(SafeStr(data(i, L_OUTCOME)) = "", SafeStr(data(i, L_TASK)), "") & _
            "  " & SafeStr(data(i, L_EMP)) & "  " & HoursText(NumOr0(data(i, L_BASEHRS))) & " h" & _
            IIf(NumOr0(data(i, L_CHARGED)) <> 0, "  charged " & HoursText(NumOr0(data(i, L_CHARGED))), "") & vbNewLine
    Next i
    RowsText = s
End Function

' "Print receipt": the text onto the Receipt sheet, a PDF beside the roster, the sheet hidden again
Public Sub ReceiptPdf()
    On Error GoTo Oops
    Dim ws As Worksheet
    If Not SheetExists(RECEIPT_SHEET) Then
        UnlockStructure
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RECEIPT_SHEET
        RelockStructure
        SetTabColor ws, RGB(30, 56, 100)
    End If
    Set ws = ThisWorkbook.Worksheets(RECEIPT_SHEET)
    SetSheetVisible ws, xlSheetVisible
    UnlockSheet ws
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    Dim lines As Variant: lines = Split(ReceiptText(gReceipt), vbNewLine)
    Dim i As Long
    For i = 0 To UBound(lines)
        ws.Cells(i + 1, 1).NumberFormat = "@"
        ws.Cells(i + 1, 1).Value = CStr(lines(i))
        If i = 0 Then ws.Cells(1, 1).Font.Bold = True: ws.Cells(1, 1).Font.Size = 13
        If CStr(lines(i)) = "WHAT WAS RECORDED" Or CStr(lines(i)) = "WHO IT CHANGED" Or CStr(lines(i)) = "WHERE TO SEE IT" Or CStr(lines(i)) = "IF THIS IS WRONG" Then ws.Cells(i + 1, 1).Font.Bold = True
    Next i
    ws.Columns(1).ColumnWidth = 110
    ws.Columns(1).WrapText = True
    On Error Resume Next
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(UBound(lines) + 1, 1)).Address
        .Orientation = xlPortrait: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .CenterFooter = "Call Out Roster receipt"
    End With
    On Error GoTo Oops
    ProtectSheet ws
    Dim nm As String: nm = gReceipt.Title
    Dim bad As Variant
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        nm = Replace(nm, CStr(bad), " ")
    Next bad
    ReportPdf ws, "Receipt " & Left$(Trim$(nm), 60) & " " & Format$(Now, "hhmm"), "Receipt", True
    SetSheetVisible ws, xlSheetVeryHidden
    Exit Sub
Oops:
    On Error Resume Next
    SetSheetVisible ThisWorkbook.Worksheets(RECEIPT_SHEET), xlSheetVeryHidden   ' (review M6) never left as a tab
    ReportError "ReceiptPdf"
End Sub

' "Email me this": a mail draft with the receipt's text (the address is typed in the mail app)
Public Sub ReceiptMail()
    On Error GoTo Oops
    If Not OpenMailTo("", "Call Out Roster receipt: " & gReceipt.Title, ReceiptText(gReceipt)) Then
        MsgBox "No mail app could be opened. Print receipt saves the same text as a PDF beside the roster.", vbExclamation, "Receipt"
    End If
    Exit Sub
Oops:
    ReportError "ReceiptMail"
End Sub
