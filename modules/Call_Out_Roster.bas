Attribute VB_Name = "Call_Out_Roster"
Option Explicit

' ====================================================================
'  PRINTED CALL OUT ROSTER (Blank Call Out Roster sheet)
'  Each department section is found by its title, cleared, and refilled
'  in call order (same order the callout form uses). Nothing on the
'  main roster is re-sorted or filtered.
' ====================================================================

Sub Call_Out()
    If Not RequireAdmin() Then Exit Sub
    If Not SheetExists(CALLOUT_SHEET) Then MsgBox "Sheet '" & CALLOUT_SHEET & "' not found.", vbCritical: Exit Sub

    Dim wsDest As Worksheet: Set wsDest = ThisWorkbook.Worksheets(CALLOUT_SHEET)
    Dim depts As Variant: depts = DeptList()
    Dim i As Long, n As Long, msg As String
    Dim recs() As EmpRec

    Application.Calculate
    Application.ScreenUpdating = False
    wsDest.Unprotect Password:=SHEET_PASSWORD

    For i = 0 To UBound(depts)
        n = LoadEligible(CStr(depts(i)), "ALL", recs)
        msg = msg & WriteRosterSection(wsDest, CStr(depts(i)), recs, n)
    Next i

    ' current pay period on the header
    Dim ps As Date: ps = PeriodStartFor(Date)
    wsDest.Range("J6").Value = ps
    wsDest.Range("K6").Value = ps + 13

    ProtectSheet wsDest
    Application.ScreenUpdating = True
    wsDest.Activate
    wsDest.Range("A4").Select

    If msg = "" Then
        MsgBox "Call out roster updated successfully!", vbInformation
    Else
        MsgBox "Call out roster updated, but:" & vbNewLine & vbNewLine & msg, vbExclamation
    End If
End Sub


' Returns "" on success, or a line describing a problem
Private Function WriteRosterSection(ByVal ws As Worksheet, ByVal dept As String, _
                                    recs() As EmpRec, ByVal n As Long) As String
    Dim title As String: title = DeptRosterTitle(dept)
    Dim f As Range
    Set f = ws.Columns(2).Find(What:=title, LookIn:=xlFormulas, LookAt:=xlPart, MatchCase:=False)
    If f Is Nothing Then
        WriteRosterSection = "- " & dept & ": section title '" & title & "' not found" & vbNewLine
        Exit Function
    End If

    Dim firstRow As Long: firstRow = f.Row + 3          ' title, blank, header, then data
    Dim lastRow As Long: lastRow = SectionEndRow(ws, f.Row)
    Dim slots As Long: slots = lastRow - firstRow + 1
    If slots < 1 Then
        WriteRosterSection = "- " & dept & ": no rows available under its title" & vbNewLine
        Exit Function
    End If

    ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, 4)).ClearContents

    Dim i As Long, m As Long
    If n < slots Then m = n Else m = slots
    For i = 1 To m
        ws.Cells(firstRow + i - 1, 1).Value = recs(i).TotalHrs
        ws.Cells(firstRow + i - 1, 2).Value = recs(i).EmpName
        ws.Cells(firstRow + i - 1, 3).Value = recs(i).EmpShift
        ws.Cells(firstRow + i - 1, 4).Value = recs(i).Phone
    Next i

    If n > slots Then
        WriteRosterSection = "- " & dept & ": " & n & " employees but only " & slots & _
                             " rows - the last " & (n - slots) & " were left off. Insert rows in that section." & vbNewLine
    End If
End Function

' Last data row of a section = row before the next section title,
' or the bottom of the print area for the last section
Private Function SectionEndRow(ByVal ws As Worksheet, ByVal titleRow As Long) As Long
    Dim r As Long
    For r = titleRow + 1 To titleRow + 300
        If InStr(1, SafeStr(ws.Cells(r, 2).Value), "Call Out Roster", vbTextCompare) > 0 Then
            SectionEndRow = r - 1
            Exit Function
        End If
    Next r

    Dim pa As String: pa = ws.PageSetup.PrintArea
    If pa <> "" Then
        On Error Resume Next
        Dim rng As Range: Set rng = ws.Range(pa)
        On Error GoTo 0
        If Not rng Is Nothing Then
            SectionEndRow = rng.Row + rng.Rows.Count - 1
            Exit Function
        End If
    End If
    SectionEndRow = titleRow + 3 + 10
End Function
