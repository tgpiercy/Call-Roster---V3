Attribute VB_Name = "modSystem"
Option Explicit

' ====================================================================
'  SYSTEM CHECK + OPENING CHECKS (v20)
'
'  Alt+F8 > SystemCheck  - run after every update. Shows each module's
'  version, flags leftovers ("modCalendar1", modFixShift ...) and checks
'  the sheets, the OT Log layout, holidays and the Kronos import.
'  Listing modules and checking the pasted form code needs Excel's
'  "Trust access to the VBA project object model" (File > Options >
'  Trust Center > Trust Center Settings > Macro Settings). Without it,
'  the module versions are still shown.
'
'  OpenChecks (from Workbook_Open):
'   - read-only copy: says who has the file open for editing
'   - file opened from SharePoint / OneDrive: AutoSave off (each submit
'     already saves), so two people's edits can't be merged silently
' ====================================================================

Public Const SYS_VERSION As String = "System v20 (Oct 9)"
Public Const EXPECTED_RELEASE As String = "v20"


' ---------------- SystemCheck ----------------
' Problems come first; the message stays short enough for one MsgBox.
Public Sub SystemCheck()
    Dim warn As String, info As String, nOk As Long

    ' ---- modules: each one's version constant (a missing module won't compile) ----
    VerCheck "modOvertimeEntry", CORE_VERSION, warn, nOk
    VerCheck "modCalendar", CAL_VERSION, warn, nOk
    VerCheck "modHolidays", HOL_VERSION, warn, nOk
    VerCheck "modShiftCycle", ROT_VERSION, warn, nOk
    VerCheck "modShiftSwitch", SS_VERSION, warn, nOk
    VerCheck "modPostings", POST_VERSION, warn, nOk
    VerCheck "ImportKronosData", KR_VERSION, warn, nOk
    VerCheck "modUI", UI_VERSION, warn, nOk
    VerCheck "modSystem", SYS_VERSION, warn, nOk
    info = "Modules: " & nOk & " of 9 on " & EXPECTED_RELEASE & vbNewLine
    If InStr(ThisWorkbookVersion(), EXPECTED_RELEASE) = 0 Then
        warn = warn & "- ThisWorkbook code is older than " & EXPECTED_RELEASE & ": re-paste ThisWorkbook.txt" & vbNewLine
    End If

    ' ---- VBA project (needs trusted access) ----
    Dim proj As Object, comp As Object, nm As String, trusted As Boolean, cnt As Long
    On Error Resume Next
    Set proj = ThisWorkbook.VBProject
    trusted = Not proj Is Nothing
    If trusted Then
        cnt = proj.VBComponents.Count
        If Err.Number <> 0 Then trusted = False
    End If
    On Error GoTo 0
    If Not trusted Then
        info = info & "Forms / leftovers: not checked (VBA project access is off)" & vbNewLine
    Else
        Dim compNames As String, f As Variant
        For Each comp In proj.VBComponents
            nm = comp.Name
            compNames = compNames & "|" & LCase$(nm) & "|"
            If nm Like "*[0-9]" And Not nm Like "Sheet*" And Not nm Like "Module*" And Not nm Like "Class*" And Not nm Like "Chart*" Then
                warn = warn & "- '" & nm & "' looks like a duplicate from an import: remove it" & vbNewLine
            End If
            Select Case LCase$(nm)
                Case "modfixshift", "refusal_report", "module1", "class1"
                    warn = warn & "- '" & nm & "' is left over from an earlier version: remove it" & vbNewLine
            End Select
        Next comp
        For Each f In Array("frmShiftSwitch", "frmOvertimeEntry", "frmNewPosting", "frmPostingEntry", "frmTimeOff")
            If InStr(compNames, "|" & LCase$(CStr(f)) & "|") = 0 Then
                warn = warn & "- form " & f & " is missing" & vbNewLine
            End If
        Next f
        For Each f In Array("frmShiftSwitch", "frmOvertimeEntry", "frmNewPosting")      ' forms changed in v20
            If InStr(compNames, "|" & LCase$(CStr(f)) & "|") > 0 Then
                If InStr(FormMarker(proj, CStr(f)), EXPECTED_RELEASE) = 0 Then
                    warn = warn & "- " & f & " code is older than " & EXPECTED_RELEASE & ": re-paste " & f & ".txt" & vbNewLine
                End If
            End If
        Next f
        info = info & "Forms and leftovers: checked" & vbNewLine
    End If

    ' ---- data ----
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(LOG_SHEET).ListObjects(LOG_TABLE)
    On Error GoTo 0
    If lo Is Nothing Then
        warn = warn & "- OT Log table not found" & vbNewLine
    ElseIf Not LogLayoutCurrent(lo) Then
        warn = warn & "- OT Log is in the old layout: run UpgradeOTLogLayout" & vbNewLine
    Else
        info = info & "OT Log: " & Format$(lo.ListRows.Count, "#,##0") & " rows" & vbNewLine
    End If
    info = info & KronosStatusText() & vbNewLine
    If KronosIsStale() Then warn = warn & "- Kronos data is old: run Import Kronos" & vbNewLine
    Dim hv As Variant, k As Long
    hv = HolidaysToConfirm(Date, Date + HOL_CONFIRM_DAYS - 1)
    If IsArray(hv) Then
        For k = LBound(hv) To UBound(hv)
            warn = warn & "- " & Mid$(hv(k), 12) & " " & Left$(hv(k), 10) & _
                   " is on a weekend: enter the weekday off on the Holidays sheet" & vbNewLine
        Next k
    End If
    info = info & "Holidays sheet: " & IIf(SheetExists(HOL_SHEET), "present", "created on next open (admin)") & vbNewLine
    If ThisWorkbook.ReadOnly Then warn = warn & "- This copy is read-only" & vbNewLine

    If Len(warn) > 700 Then warn = Left$(warn, InStrRev(Left$(warn, 700), vbNewLine)) & "- ...and more" & vbNewLine
    Dim msg As String
    If warn = "" Then
        msg = "Everything looks right." & vbNewLine & vbNewLine & info
    Else
        msg = "CHECK:" & vbNewLine & warn & vbNewLine & info
    End If
    MsgBox msg, IIf(warn = "", vbInformation, vbExclamation), "System check  (" & EXPECTED_RELEASE & ")"
End Sub

Private Sub VerCheck(ByVal modName As String, ByVal ver As String, ByRef warn As String, ByRef nOk As Long)
    If InStr(ver, EXPECTED_RELEASE) > 0 Then
        nOk = nOk + 1
    Else
        warn = warn & "- " & modName & " is " & ver & ": remove it and import the " & EXPECTED_RELEASE & " file" & vbNewLine
    End If
End Sub

' ThisWorkbook exposes CodeVersion from v20 on; older code doesn't have it
Private Function ThisWorkbookVersion() As String
    Dim o As Object: Set o = ThisWorkbook
    On Error Resume Next
    ThisWorkbookVersion = o.CodeVersion
    On Error GoTo 0
    If ThisWorkbookVersion = "" Then ThisWorkbookVersion = "before v20"
End Function

' the "' @version ..." line in a form's code, or "no version line"
Private Function FormMarker(ByVal proj As Object, ByVal formName As String) As String
    On Error Resume Next
    Dim cm As Object: Set cm = proj.VBComponents(formName).CodeModule
    Dim n As Long: n = cm.CountOfLines
    If n > 60 Then n = 60
    Dim i As Long, ln As String
    For i = 1 To n
        ln = Trim$(cm.Lines(i, 1))
        If Left$(ln, 10) = "' @version" Then FormMarker = Trim$(Mid$(ln, 11)): Exit Function
    Next i
    FormMarker = "no version line (before v20)"
End Function


' ---------------- opening checks ----------------
Public Sub OpenChecks()
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then
        Dim who As String: who = LockOwner()
        MsgBox IIf(who <> "", who & " has", "Someone else has") & " the roster open for editing, so this copy is read-only." & _
               vbNewLine & vbNewLine & "You can look at everything, but nothing can be logged or saved." & vbNewLine & _
               "Close it and open it again once they're out.", vbInformation, "Read-only"
        Exit Sub
    End If
    ' SharePoint / OneDrive: AutoSave off - every submit saves anyway, and
    ' two people's copies must never be merged in the background.
    If LCase$(Left$(ThisWorkbook.FullName, 4)) = "http" Then
        Dim wbo As Object: Set wbo = ThisWorkbook      ' late-bound: older Excel has no AutoSaveOn
        If wbo.AutoSaveOn Then wbo.AutoSaveOn = False
    End If
End Sub

' Name of whoever holds the edit lock on a network share (from Excel's
' "~$" owner file next to the workbook), or "" if it can't be read.
Public Function LockOwner() As String
    On Error GoTo Done
    If LCase$(Left$(ThisWorkbook.FullName, 4)) = "http" Then Exit Function
    Dim p As String: p = ThisWorkbook.Path & Application.PathSeparator & "~$" & ThisWorkbook.Name
    If Dir$(p, vbHidden) = "" Then Exit Function
    Dim f As Integer: f = FreeFile
    Dim b() As Byte, n As Long
    Open p For Binary Access Read Shared As #f
    n = LOF(f)
    If n < 2 Then Close #f: Exit Function
    If n > 200 Then n = 200
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    Dim L As Long: L = b(0)
    If L < 1 Or L > n - 1 Then Exit Function
    Dim i As Long, s As String
    For i = 1 To L
        s = s & Chr$(b(i))
    Next i
    LockOwner = Trim$(s)
    Exit Function
Done:
    On Error Resume Next
    Close #f
End Function
