Attribute VB_Name = "modDates"
Option Explicit
' v27.1 (the refactor, Oct 2026): the one typed-date parser (ParseTypedDate / TypedDateCore / DateText), from modPeople; moved unchanged (tests/movecheck.py proves it)
Public Const DATES_VERSION As String = "Typed dates v29.1 (Oct 28)"

' ====================================================================
'  TYPED DATES
' ====================================================================
' v25 (D-C7): THE one typed-date rule, for every date box and date prompt. "" = d holds the
' date; otherwise the word for the hint: "blank", "not a date", "check the year". No year typed
' ("Jan 12", "1/12") = this year, or NEXT year when that is more than 90 days back (typing
' January in December). More than ~13 months from today either way is refused. A time typed
' after the date is ignored ("Oct 12 7pm", "10/12 7:00"). Pure: today is passed in (SelfTest
' suite typeddate; the year of a yearless date comes from today, never from the PC clock).
' v25 review: read by hand, never through CDate - CDate took "Feb 29" as Feb 2029 and "Sep 31" as
' Sep 1931, and a slash date in whatever order the PC's locale says. A yearless date lands in the
' year that puts it inside today - 90 .. today + 275 (one 365-day window: "Dec 20" typed on Jan 3
' is LAST December, "Jan 5" typed in December is next January).
Public Function TypedDateCore(ByVal t As String, ByVal today As Date, ByRef d As Date) As String
    t = Trim$(t)
    If t = "" Then TypedDateCore = "blank": Exit Function
    Dim y As Long, m As Long, dy As Long, hasYear As Boolean, dd As Date
    If Not TypedDateParts(t, y, m, dy, hasYear) Then TypedDateCore = "not a date": Exit Function
    If Not hasYear Then
        y = Year(today)
        If m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then TypedDateCore = "not a date": Exit Function
        dd = DateSerial(y, m, IIf(dy > 28, 28, dy))    ' the year is chosen first ("Feb 29" typed before a leap February)
        If dd < today - 90 Then
            y = y + 1
        ElseIf dd > today + 275 Then
            y = y - 1
        End If
    End If
    If Not ValidYmd(y, m, dy) Then TypedDateCore = "not a date": Exit Function
    dd = DateSerial(y, m, dy)
    If Abs(dd - today) > 400 Then TypedDateCore = "check the year": Exit Function
    d = dd
End Function

' The pieces of a typed date: "Oct 12", "October 12", "12 Oct", "Oct 12, 2026", "Oct 12 26",
' "2026-10-12", "10/12", "10/12/2026" (month / day, as the boxes' hints say). A time after it
' ("7:00", "7pm", "7 pm") is dropped. False = none of these.
Private Function TypedDateParts(ByVal t As String, ByRef y As Long, ByRef m As Long, ByRef dy As Long, ByRef hasYear As Boolean) As Boolean
    Dim raw As Variant, toks() As String, n As Long, w As String, i As Long, clockSeen As Boolean
    raw = Split(Trim$(Replace(Replace(t, ",", " "), ".", " ")), " ")
    ReDim toks(0 To UBound(raw) + 1)
    For i = 0 To UBound(raw)
        w = Trim$(CStr(raw(i)))
        If w = "" Then
        ElseIf InStr(w, ":") > 0 Then                               ' 7:00 - the time, dropped
            clockSeen = True
        ElseIf LCase$(w) Like "#*[ap]m" Then                        ' 7pm
        ElseIf LCase$(w) = "am" Or LCase$(w) = "pm" Then            ' "7 pm": the 7 goes too ("7:00 pm": nothing more to drop)
            If n > 0 And Not clockSeen Then
                If toks(n - 1) Like "#*" And Not (toks(n - 1) Like "*[-/]*") Then n = n - 1
            End If
        Else
            toks(n) = w: n = n + 1
        End If
    Next i
    hasYear = False
    If n = 1 Then
        If toks(0) Like "####-##-##" Or toks(0) Like "####-#-#" Or toks(0) Like "####-##-#" Or toks(0) Like "####-#-##" Then
            Dim p As Variant: p = Split(toks(0), "-")
            y = CLng(p(0)): m = CLng(p(1)): dy = CLng(p(2)): hasYear = True
            TypedDateParts = True
        ElseIf InStr(toks(0), "/") > 0 Then
            p = Split(toks(0), "/")
            If UBound(p) < 1 Or UBound(p) > 2 Then Exit Function
            For i = 0 To UBound(p)
                If Not (CStr(p(i)) Like "#" Or CStr(p(i)) Like "##" Or (i = 2 And CStr(p(i)) Like "####")) Then Exit Function
            Next i
            m = CLng(p(0)): dy = CLng(p(1))
            If UBound(p) = 2 Then
                y = CLng(p(2)): hasYear = True
                If y < 100 Then y = y + 2000
            End If
            TypedDateParts = True
        End If
        Exit Function
    End If
    If n < 2 Or n > 3 Then Exit Function
    Dim mo As Long
    mo = MonthIndex(toks(0))
    If mo > 0 Then
        If Not IsDigits(toks(1)) Then Exit Function
        m = mo: dy = CLng(toks(1))
    Else
        mo = MonthIndex(toks(1))
        If mo = 0 Or Not IsDigits(toks(0)) Then Exit Function
        m = mo: dy = CLng(toks(0))
    End If
    If n = 3 Then
        If Not IsDigits(toks(2)) Then Exit Function
        y = CLng(toks(2)): hasYear = True
        If y < 100 Then y = y + 2000
    End If
    TypedDateParts = True
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    IsDigits = (s <> "" And Len(s) <= 4 And s Like String$(Len(s), "#"))
End Function

' "Oct" / "October" / "oct" -> 10; 0 when not a month name
Private Function MonthIndex(ByVal s As String) As Long
    Dim k As Long, names As Variant
    names = Array("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")
    s = LCase$(s)
    If Len(s) < 3 Then Exit Function
    For k = 0 To 11
        If Left$(s, 3) = names(k) Then
            If s Like "[a-z]*" And Not (s Like "*[!a-z]*") Then MonthIndex = k + 1
            Exit Function
        End If
    Next k
End Function

Private Function ValidYmd(ByVal y As Long, ByVal m As Long, ByVal dy As Long) As Boolean
    If y < 100 Or y > 9999 Or m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then Exit Function
    ValidYmd = (Day(DateSerial(y, m, dy)) = dy)
End Function

' Reads a typed date (no time part) as of today: True and d, or False (the box shows why).
Public Function ParseTypedDate(ByVal t As String, ByRef d As Date) As Boolean
    ParseTypedDate = (TypedDateCore(t, Date, d) = "")
End Function

' v25: the hint under a date box: "= Mon Oct 12, 2026" in green, or the word in red ("" when blank)
Public Sub TypedDateHint(ByVal t As String, ByVal lbl As Object, Optional ByVal suffix As String = "")
    Dim d As Date, w As String
    w = TypedDateCore(t, Date, d)
    If w = "blank" Then
        lbl.Caption = "": lbl.ForeColor = RGB(120, 120, 120)
    ElseIf w = "" Then
        lbl.Caption = "= " & Format$(d, "ddd mmm d, yyyy") & suffix: lbl.ForeColor = RGB(0, 110, 60)
    Else
        lbl.Caption = w: lbl.ForeColor = RGB(190, 30, 30)
    End If
End Sub

' v25: how every date box is pre-filled (no ambiguity about month / day order)
Public Function DateText(ByVal d As Date) As String
    DateText = Format$(d, "mmm d, yyyy")
End Function

' A typed date that must carry its year (a seniority date can be decades
' back, so nothing is guessed): "Mar 3, 2011", "3/3/2011". The pop-up shows
' the date back in words before anything is saved.
Public Function ParseFullDate(ByVal t As String, ByRef d As Date) As Boolean
    t = Trim$(t)
    If t = "" Then Exit Function
    If Not IsDate(t) Then Exit Function
    If Not TypedHasYear(t) Then Exit Function
    d = Int(CDate(t))
    ParseFullDate = (Year(d) >= 1950 And d <= Date + 400)
End Function

' True when the text includes a year: a 3+ digit number, three numbers
' (1/12/27), or a month name plus two numbers (Jan 12 27). v25: a time typed
' after the date ("10/12 7:00", "Oct 12 7pm") is left out first.
Private Function TypedHasYear(ByVal t As String) As Boolean
    Dim i As Long, ch As String, tok As String
    Dim nNum As Long, hasAlpha As Boolean, bigNum As Boolean
    Dim p As Long: p = InStr(t, ":")
    If p > 0 Then t = Left$(t, InStrRev(Left$(t, p), " "))
    Dim w As Variant, kept As String
    For Each w In Split(t, " ")
        If Not (LCase$(CStr(w)) Like "#*[ap]m") Then kept = kept & CStr(w) & " "
    Next w
    t = kept & " "
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch Like "[0-9]" Then
            tok = tok & ch
        Else
            If tok <> "" Then
                nNum = nNum + 1
                If Len(tok) >= 3 Then bigNum = True
                tok = ""
            End If
            If ch Like "[A-Za-z]" Then hasAlpha = True
        End If
    Next i
    TypedHasYear = bigNum Or nNum >= 3 Or (hasAlpha And nNum >= 2)
End Function


