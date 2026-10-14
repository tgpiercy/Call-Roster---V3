Attribute VB_Name = "modShiftCycle"
Option Explicit

' ====================================================================
'  12-HOUR SHIFT ROTATION
'  Taken from the plant's 12 Hour Shift Schedule (SHIFT2026). It is a
'  28-day cycle, so it works for any date in any year:
'    day 1 of the cycle = Sunday Dec 28, 2025 (B days / A nights)
'  Days = 7 AM - 7 PM.  Nights = 7 PM - 7 AM (dated by the evening it starts).
'
'  Available crew for overtime = the two crews OFF that day, minus:
'    - DAY OT:   the crew that worked NIGHTS the night before
'                (they finish at 7 AM that morning - not eligible)
'    - NIGHT OT: the crew that starts DAYS the next morning at 7 AM
'                (they'd be working straight through)
' ====================================================================

Public Const ROT_VERSION As String = "Shift cycle v24.4 (Oct 14)"

Private Const ROT_ANCHOR As Date = #12/28/2025#
Private Const ROT_DAYS As String = "BDDAACCCBBDDAAACCBBDDDAACCBB"     ' crew on days, cycle day 1..28
Private Const ROT_NIGHTS As String = "ACCBBDDDAACCBBBDDAACCCBBDDAA"   ' crew on nights, cycle day 1..28
Private Const ROT_LEN As Long = 28

Private Function RotIdx(ByVal d As Date) As Long
    Dim k As Long
    k = DateDiff("d", ROT_ANCHOR, Int(d)) Mod ROT_LEN
    If k < 0 Then k = k + ROT_LEN
    RotIdx = k + 1
End Function

Public Function DaysCrew(ByVal d As Date) As String
    DaysCrew = Mid$(ROT_DAYS, RotIdx(d), 1)
End Function

Public Function NightsCrew(ByVal d As Date) As String
    NightsCrew = Mid$(ROT_NIGHTS, RotIdx(d), 1)
End Function

' The two crews with the day off, e.g. "A/B"
Public Function CrewsOff(ByVal d As Date) As String
    Dim c As Variant, s As String
    For Each c In Array("A", "B", "C", "D")
        If c <> DaysCrew(d) And c <> NightsCrew(d) Then s = s & IIf(s = "", "", "/") & c
    Next c
    CrewsOff = s
End Function

' Which shift a roster crew code works on date d: "Days", "Nights", "" (off) or "?" (unknown code)
'   A, B, C, D   the 4-crew rotation - days and nights
'   A/B, C/D     that pair's DAY shifts only (on whenever A or B is on days)
'   DAYS         Monday - Friday days, except plant holidays (v20, see modHolidays)
Public Function CrewShiftOn(ByVal crewCode As String, ByVal d As Date) As String
    Dim s As String: s = UCase$(Replace(Trim$(crewCode), " ", ""))
    If Len(s) = 1 And InStr("ABCD", s) > 0 Then
        If s = DaysCrew(d) Then CrewShiftOn = "Days"
        If s = NightsCrew(d) Then CrewShiftOn = "Nights"
    ElseIf s Like "[ABCD]/[ABCD]" Then
        If InStr(s, DaysCrew(d)) > 0 Then CrewShiftOn = "Days"
    ElseIf s = "DAYS" Then
        If Weekday(d, vbMonday) <= 5 Then
            If Not DaysHolidayOff(d) Then CrewShiftOn = "Days"
        End If
    Else
        CrewShiftOn = "?"
    End If
End Function

Public Function IsNightShift(ByVal shiftLbl As String) As Boolean
    IsNightShift = (UCase$(Left$(Trim$(shiftLbl), 1)) = "N")
End Function

' Crew(s) eligible for OT on date d, Days or Nights. note = why a crew
' that's off isn't eligible (blank if both off crews are eligible).
Public Function AvailableCrew(ByVal d As Date, ByVal isNight As Boolean, Optional ByRef note As String) As String
    Dim c As Variant, s As String, blocked As String
    note = ""
    If isNight Then
        blocked = DaysCrew(d + 1)
    Else
        blocked = NightsCrew(d - 1)
    End If
    For Each c In Array("A", "B", "C", "D")
        If c <> DaysCrew(d) And c <> NightsCrew(d) Then
            If c = blocked Then
                If isNight Then
                    note = c & " starts days next morning"
                Else
                    note = c & " coming off nights"
                End If
            Else
                s = s & IIf(s = "", "", "/") & c
            End If
        End If
    Next c
    AvailableCrew = s
End Function

' One-line summary for the pop-ups, e.g. "On: D days, C nights | Off: A/B"
Public Function RotationSummary(ByVal d As Date) As String
    RotationSummary = "On: " & DaysCrew(d) & " days, " & NightsCrew(d) & " nights  |  Off: " & CrewsOff(d)
End Function


' ====================================================================
'  DAYLIGHT SAVING NIGHTS (v21)
'  Clocks go back on the first Sunday in November at 2 AM: the night
'  shift that STARTS the Saturday before works 13 hours. Clocks go
'  forward on the second Sunday in March: that Saturday's night shift
'  works 11 hours. Worked out by rule for any year (like modHolidays).
'  Day shifts are never affected. DST_USE_ACTUAL_HOURS (modOvertimeEntry)
'  decides whether the forms default to the actual hours or just warn.
' ====================================================================
Public Function DstFallBack(ByVal y As Long) As Date          ' the Saturday night of 13 h
    DstFallBack = NthSunday(y, 11, 1) - 1
End Function

Public Function DstSpringForward(ByVal y As Long) As Date     ' the Saturday night of 11 h
    DstSpringForward = NthSunday(y, 3, 2) - 1
End Function

Private Function NthSunday(ByVal y As Long, ByVal m As Long, ByVal n As Long) As Date
    Dim d As Date: d = DateSerial(y, m, 1)
    d = d + ((8 - Weekday(d, vbSunday)) Mod 7)        ' first Sunday
    NthSunday = d + 7 * (n - 1)
End Function

' Actual length of the night shift starting on date d: 13, 11 or 12
Public Function DstNightHours(ByVal d As Date) As Double
    Dim k As Date: k = Int(d)
    If k = DstFallBack(Year(k)) Then
        DstNightHours = 13
    ElseIf k = DstSpringForward(Year(k)) Then
        DstNightHours = 11
    Else
        DstNightHours = 12
    End If
End Function

' True when the night starting on d is a clock-change night
Public Function IsDstNight(ByVal d As Date) As Boolean
    IsDstNight = (DstNightHours(d) <> 12)
End Function

' One-line warning for the forms ("" when d is not a change night)
Public Function DstNote(ByVal d As Date) As String
    Dim h As Double: h = DstNightHours(d)
    If h = 12 Then Exit Function
    Dim way As String: way = IIf(h > 12, "clocks go back", "clocks go forward")
    If DST_USE_ACTUAL_HOURS Then
        DstNote = "DST night: " & h & " h (" & way & ")"
    Else
        DstNote = "DST night: " & way & " - actual shift is " & h & " h, 12 h kept"
    End If
End Function

' Hours the forms should default to for a shift on d: the actual DST
' hours for a change night (when the setting is on), else the usual
Public Function DefaultShiftHours(ByVal d As Date, ByVal isNight As Boolean, ByVal usual As Double) As Double
    DefaultShiftHours = usual
    If isNight And DST_USE_ACTUAL_HOURS Then
        If IsDstNight(d) Then DefaultShiftHours = DstNightHours(d)
    End If
End Function
