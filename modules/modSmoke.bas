Attribute VB_Name = "modSmoke"
Option Explicit

' ====================================================================
'  modSmoke (v24.4): pop-up openers for the Install panel's test rows
'
'  The Roster Updater's Install panel runs these through Application.Run
'  with the arguments a test row names ("SmokePosting Phosphine +21 Days":
'  the panel turns +21 into today plus 21 days and passes a Date). Each
'  one opens a pop-up pre-filled and returns when it is closed; the admin
'  then marks the row Pass or Fail. Nothing here writes to any table: a
'  pop-up that is cancelled leaves nothing, and one that is submitted goes
'  through the pop-up's own code, as a click would.
' ====================================================================

Public Const SMO_VERSION As String = "Smoke openers v29 (Oct 28)"

' Log Call-Out, pre-filled (unit, date, Days / Nights)
Public Sub SmokeCallout(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String)
    On Error GoTo Oops
    If Not RequireWritable() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    frmOvertimeEntry.Prefill unit, Int(d), shiftLbl, "From the Install panel"
    frmOvertimeEntry.Show
    Exit Sub
Oops:
    ReportError "SmokeCallout"
End Sub

' New Posting, pre-filled (unit, date, Days / Nights; reason "Unit Operator Support", the usual hours)
Public Sub SmokePosting(ByVal unit As String, ByVal d As Date, ByVal shiftLbl As String)
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not PostingTablesReady() Then Exit Sub
    If LogTable() Is Nothing Then Exit Sub
    frmNewPosting.Prefill unit, Int(d), shiftLbl, "Unit Operator Support", USUAL_SHIFT_HOURS, _
                          Format$(Int(d), "ddd mmm d") & " " & shiftLbl & ": from the Install panel"
    frmNewPosting.Show
    Exit Sub
Oops:
    ReportError "SmokePosting"
End Sub

' v25.5: Sick Call, as the Desk opens it (the pop-up picks its own person and date; d goes on the status line)
Public Sub SmokeSick(ByVal d As Date)
    On Error GoTo Oops
    Application.StatusBar = "Install panel: Sick Call for " & Format$(Int(d), "ddd mmm d") & " - look at the pop-up at scale, then Close"
    ShowSickCallForm
    Application.StatusBar = False
    Exit Sub
Oops:
    Application.StatusBar = False
    ReportError "SmokeSick"
End Sub

' v25.5: Award Posting, as the Desk opens it (pick any open posting, LOAD, look, then close without submitting)
Public Sub SmokeAward(ByVal d As Date)
    On Error GoTo Oops
    Application.StatusBar = "Install panel: Award Posting (" & Format$(Int(d), "ddd mmm d") & ") - LOAD any posting, look at the pop-up at scale, then close it"
    ShowPostingForm
    Application.StatusBar = False
    Exit Sub
Oops:
    Application.StatusBar = False
    ReportError "SmokeAward"
End Sub

' Shift Switch (the pop-up takes its dates by hand; d is shown in the title bar's status line only)
Public Sub SmokeSwitch(ByVal d As Date)
    On Error GoTo Oops
    Application.StatusBar = "Install panel: Shift Switch for " & Format$(Int(d), "ddd mmm d") & " - type the dates, or Cancel"
    ShowShiftSwitchForm
    Application.StatusBar = False
    Exit Sub
Oops:
    Application.StatusBar = False
    ReportError "SmokeSwitch"
End Sub
