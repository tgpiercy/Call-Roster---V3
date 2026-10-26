Attribute VB_Name = "modPostingMail"
Option Explicit
' v27.1 (the refactor, Oct 2026): the award email (Outlook draft, mailto fallback), from modPostings; moved unchanged (tests/movecheck.py proves it)
Public Const PMAIL_VERSION As String = "Posting mail v27.5 (Oct 25)"

' ====================================================================
'  EMAIL (Outlook draft - you review and click Send)
' ====================================================================
' CC list for an award email, by unit. missingGroups returns the names of
' any group whose address isn't set yet (shown as a warning).
Private Function AwardCC(ByVal unit As String, ByRef missingGroups As String) As String
    Dim cc As String
    missingGroups = ""
    Select Case unit
        Case "Cypure", "Material Handler"              ' Material Handler = Fumigation (gases)
            AddCC cc, missingGroups, CC_GASES_SUPS, "Gases supervisors"
        Case "QC Derivatives", "QC Gases"              ' Quality Control Lab
            AddCC cc, missingGroups, CC_QC_SUP, "Quality Control supervisor"
        Case "Utilities"
            AddCC cc, missingGroups, CC_UTIL_SUP, "Utilities supervisor"
        Case Else                                      ' Phosphine, Derivatives, SBP
            AddCC cc, missingGroups, CC_SHIFT_SUPS, "Shift supervisors"
    End Select
    AddCC cc, missingGroups, CC_OPS_SUPT, "Operations superintendent"
    If Trim$(AWARD_CC) <> "" Then AddCC cc, missingGroups, AWARD_CC, ""
    AwardCC = cc
End Function

Private Sub AddCC(ByRef cc As String, ByRef missingGroups As String, ByVal addr As String, ByVal groupName As String)
    addr = Trim$(addr)
    If addr = "" Then
        If groupName <> "" Then missingGroups = missingGroups & "  " & groupName & vbNewLine
    ElseIf InStr(1, "; " & cc & ";", "; " & addr & ";", vbTextCompare) = 0 Then
        cc = cc & IIf(cc = "", "", "; ") & addr
    End If
End Sub

Public Function EmailFor(ByVal empName As String) As String
    Dim ws As Worksheet: Set ws = RosterSheet()
    Dim r As Long
    For r = R_FIRST To RosterLastRow()
        If StrComp(Trim$(SafeStr(ws.Cells(r, C_NAME).Value)), empName, vbTextCompare) = 0 Then
            EmailFor = Trim$(SafeStr(ws.Cells(r, C_EMAIL).Value))
            Exit Function
        End If
    Next r
End Function

Public Sub EmailAward(ByRef p As PostRec, ByRef lst() As PListRec, ByVal n As Long, ByVal awardNames As String, _
                       Optional ByVal outName As String = "")
    Dim toList As String, missing As String, em As String, i As Long
    For i = 1 To n
        If lst(i).Awarded Then
            em = EmailFor(lst(i).EmpName)
            If em = "" Or InStr(em, "@") = 0 Then
                missing = missing & "  " & lst(i).EmpName & vbNewLine
            Else
                toList = toList & IIf(toList = "", "", "; ") & em
            End If
        End If
    Next i
    If outName <> "" Then
        em = EmailFor(outName)
        If em = "" Or InStr(em, "@") = 0 Then
            missing = missing & "  " & outName & vbNewLine
        Else
            toList = toList & IIf(toList = "", "", "; ") & em
        End If
    End If

    Dim subj As String
    subj = "Overtime awarded - " & p.ID & " - " & p.Unit & " - " & _
           Format$(p.ShiftDate, "ddd mmm d, yyyy") & " (" & p.Crew & ")"
    ' plain-text version for the fallback (v20)
    Dim plain As String
    plain = "The overtime posting below has been awarded." & vbCrLf & vbCrLf & _
            "Posting: " & p.ID & vbCrLf & "Unit: " & p.Unit & vbCrLf & _
            "Shift date: " & Format$(p.ShiftDate, "dddd, mmmm d, yyyy") & vbCrLf & _
            "Shift / crew: " & p.ShiftLbl & " / " & p.Crew & vbCrLf & _
            "Hours: " & CStr(p.Hours) & vbCrLf & "Reason: " & p.Reason & vbCrLf & _
            "Awarded to: " & awardNames & vbCrLf & vbCrLf & _
            "Awarded by " & CurrentUser() & " on " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & "."

    Dim ccMissing As String, ccList As String
    ccList = AwardCC(p.Unit, ccMissing)
    If missing <> "" Or ccMissing <> "" Then
        MsgBox IIf(missing <> "", "No email address on the roster (column K) for:" & vbNewLine & missing & vbNewLine, "") & _
               IIf(ccMissing <> "", "No CC address is set for:" & vbNewLine & ccMissing & "(" & PRIMARY_ADMIN_NAME & " sets it in the settings)" & vbNewLine, "") & _
               "Add them to the email before sending.", vbExclamation, "Call Out Roster"
    End If

    Dim ol As Object, mi As Object, shown As Boolean
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        ' new Outlook / no Outlook: open a plain email in the default mail app instead
        MailtoFallback toList, ccList, subj, plain, "Outlook couldn't be opened"
        Exit Sub
    End If

    Dim body As String
    body = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:11pt'>" & _
           "<p>The overtime posting below has been awarded.</p>" & _
           "<table cellpadding='4' style='border-collapse:collapse'>" & _
           HtmlRow("Posting", p.ID) & _
           HtmlRow("Unit", p.Unit) & _
           HtmlRow("Shift date", Format$(p.ShiftDate, "dddd, mmmm d, yyyy")) & _
           HtmlRow("Shift / crew", p.ShiftLbl & " / " & p.Crew) & _
           HtmlRow("Hours", CStr(p.Hours)) & _
           HtmlRow("Reason", p.Reason) & _
           HtmlRow("Awarded to", awardNames) & _
           "</table><p>Awarded by " & CurrentUser() & " on " & Format$(Now, "mmm d, yyyy h:mm AM/PM") & ".</p></div>"

    On Error GoTo MailFail
    Set mi = ol.CreateItem(0)
    mi.To = toList
    mi.CC = ccList
    mi.Subject = subj
    mi.Display                                   ' opens with your signature
    shown = True
    Dim h As String, b1 As Long, b2 As Long
    h = mi.HTMLBody
    b1 = InStr(1, h, "<body", vbTextCompare)
    If b1 > 0 Then b2 = InStr(b1, h, ">")
    If b2 > 0 Then
        mi.HTMLBody = Left$(h, b2) & body & Mid$(h, b2 + 1)
    Else
        mi.HTMLBody = body & h
    End If
    On Error GoTo 0
    Exit Sub
MailFail:
    Dim mailErr As String: mailErr = Err.Description
    On Error GoTo 0
    If shown Then
        MsgBox "Award saved. The email opened, but the award details couldn't be added to it (" & mailErr & ")." & _
               vbNewLine & "Type them in before sending.", vbExclamation, "Call Out Roster"
    Else
        MailtoFallback toList, ccList, subj, plain, "Outlook couldn't create the email (" & mailErr & ")"
    End If
End Sub

' v20: when Outlook can't be driven (new Outlook, no Outlook), open a plain
' email in the default mail app via mailto:, or show the text to copy.
Private Sub MailtoFallback(ByVal toList As String, ByVal ccList As String, ByVal subj As String, _
                           ByVal plain As String, ByVal why As String)
    Dim url As String
    url = "mailto:" & UrlEnc(Replace(Replace(toList, " ", ""), ";", ",")) & _
          "?cc=" & UrlEnc(Replace(Replace(ccList, " ", ""), ";", ",")) & _
          "&subject=" & UrlEnc(subj) & "&body=" & UrlEnc(plain)
    Dim ok As Boolean
    If Len(url) < 2000 Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink Address:=url
        ok = (Err.Number = 0)
        On Error GoTo 0
    End If
    If ok Then
        MsgBox "Award saved. " & why & ", so the email opened in your default mail app instead." & vbNewLine & _
               "Check the To and CC lines, then send.", vbInformation, "Award email"
    Else
        MsgBox "Award saved. " & why & " and no mail app could be opened." & vbNewLine & vbNewLine & _
               "Send this yourself:" & vbNewLine & vbNewLine & _
               "To: " & toList & vbNewLine & "CC: " & ccList & vbNewLine & "Subject: " & subj & vbNewLine & vbNewLine & _
               plain, vbExclamation, "Award email"
    End If
End Sub

' v25.3 (the Desk's "Something went wrong"): a plain email in the default mail app through mailto:; False
' when no mail app opened (the caller shows the text instead). The body is cut to what a mailto: link carries.
Public Function OpenMailTo(ByVal toList As String, ByVal subj As String, ByVal plain As String) As Boolean
    Dim url As String
    url = "mailto:" & UrlEnc(Replace(Replace(toList, " ", ""), ";", ",")) & "?subject=" & UrlEnc(subj) & "&body=" & UrlEnc(Left$(plain, 700))
    If Len(url) >= 2000 Then Exit Function
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=url
    OpenMailTo = (Err.Number = 0)
    On Error GoTo 0
End Function

' percent-encodes text for a mailto: link (UTF-8)
Private Function UrlEnc(ByVal t As String) As String
    Dim i As Long, ch As String, c As Long, out As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        c = AscW(ch) And &HFFFF&
        If ch Like "[A-Za-z0-9]" Or InStr("-_.~@,", ch) > 0 Then
            out = out & ch
        ElseIf c < &H80& Then
            out = out & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < &H800& Then
            out = out & "%" & Hex$(&HC0& Or (c \ &H40&)) & "%" & Hex$(&H80& Or (c And &H3F&))
        Else
            out = out & "%" & Hex$(&HE0& Or (c \ &H1000&)) & "%" & Hex$(&H80& Or ((c \ &H40&) And &H3F&)) & _
                  "%" & Hex$(&H80& Or (c And &H3F&))
        End If
    Next i
    UrlEnc = out
End Function

Private Function HtmlRow(ByVal k As String, ByVal v As String) As String
    HtmlRow = "<tr><td style='color:#555;padding-right:14px'><b>" & k & "</b></td><td>" & _
              Replace(Replace(v, "&", "&amp;"), "<", "&lt;") & "</td></tr>"
End Function


