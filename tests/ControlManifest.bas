Attribute VB_Name = "ControlManifest"
Option Explicit

' ====================================================================
'  ControlManifest (tooling, Oct 2026): lists every control of the two
'  DESIGNER forms (frmOvertimeEntry, frmPostingEntry) so the repo's
'  checkers can verify handler and control names (deep review T2).
'
'  Once, in the roster (needs "Trust access to the VBA project object
'  model", which is on): Alt+F11 > File > Import File > this file,
'  Alt+F8 > WriteControlManifest, then remove this module again
'  (right-click > Remove > No). It writes "control-manifest.txt" beside
'  the roster: one line per control, "form | name | type | caption".
'  Send that file back. Nothing in the roster changes.
' ====================================================================

Public Sub WriteControlManifest()
    Dim forms As Variant: forms = Array("frmOvertimeEntry", "frmPostingEntry")
    Dim out As String, f As Variant, comp As Object, ctl As Object, cap As String, n As Long
    On Error GoTo Oops
    For Each f In forms
        Set comp = Nothing
        On Error Resume Next
        Set comp = ThisWorkbook.VBProject.VBComponents(CStr(f))
        On Error GoTo Oops
        If comp Is Nothing Then
            out = out & CStr(f) & " | (not in this workbook)" & vbCrLf
        Else
            For Each ctl In comp.Designer.Controls
                cap = ""
                On Error Resume Next
                cap = CStr(ctl.Caption)
                On Error GoTo Oops
                out = out & CStr(f) & " | " & ctl.Name & " | " & TypeName(ctl) & " | " & Replace(cap, vbCrLf, " ") & vbCrLf
                n = n + 1
            Next ctl
        End If
    Next f
    Dim p As String: p = ThisWorkbook.Path & "\control-manifest.txt"
    Dim ff As Integer: ff = FreeFile
    Open p For Output As #ff
    Print #ff, "control-manifest " & Format$(Now, "yyyy-mm-dd hh:nn") & " from " & ThisWorkbook.Name
    Print #ff, out
    Close #ff
    MsgBox n & " control(s) listed in:" & vbCrLf & p & vbCrLf & vbCrLf & "Send that file. Then remove this module (Alt+F11, right-click ControlManifest > Remove > No).", vbInformation, "Control manifest"
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description & vbCrLf & "Is 'Trust access to the VBA project object model' on?", vbExclamation, "Control manifest"
End Sub
