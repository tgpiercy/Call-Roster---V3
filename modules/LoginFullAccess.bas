Attribute VB_Name = "LoginFullAccess"
Option Explicit

' Button: "Data Entry / Full Login"
' Admin access now comes from your Windows login (ADMIN_USERS in
' modOvertimeEntry) instead of a typed password.
Sub LoginFullAccess()
    On Error GoTo Oops
    If IsAdmin() Then
        ApplyView True
        RosterSheet().Activate
        MsgBox "Full access granted (" & CurrentUser() & ").", vbInformation, "Call Out Roster"
    Else
        ApplyView False
        MsgBox "Your Windows login (" & CurrentUser() & ") is not on the admin list.", vbExclamation, "Call Out Roster"
    End If
    Exit Sub
Oops:
    ReportError "LoginFullAccess"
End Sub
