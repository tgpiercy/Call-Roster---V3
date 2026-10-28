Attribute VB_Name = "modLite"
Option Explicit

' ====================================================================
'  v29 (briefs/v29.md items 3a / 3b; briefs/admin-lite.md): THE ADMIN LITE DESK'S DATA - the sources behind Ray's cards.
'  Graeme, Oct 27, 2026: Ray (an admin on LITE_DESK_USERS) opens on a simpler Desk drawn on the FULL admin view, with
'  one action per card. This module holds what the cards show; the drawing itself (DrawLiteDesk) comes in item 3b and
'  reads only these functions. Nothing here writes to the roster or the OT Log: the three daily-check MARKS are hidden
'  names (DailyKronos / DailyNeeds / DailyPublish = "yyyy-mm-dd hh:nn|login", set by the import, Needs Attention and
'  the publish when they finish), and the counts are read from the tables the roster already keeps. Pure cores
'  (suites postexc, litemark, litedaily, litebar) are mirrored in tests/rules.py.
' ====================================================================
Public Const LITE_VERSION As String = "Lite v28 (Oct 27)"      ' new in the v29 build; stamped on the current release until the v29 release prep (test_drift)

Public Const LITE_CLOSING_HOURS As Double = 24      ' a posting with nobody signed up that closes within this many hours is an exception
Public Const LITE_STAFF_DAYS As Long = 14           ' Band 4: below minimum in the next N days
Private Const DAILY_KRONOS As String = "DailyKronos"
Private Const DAILY_NEEDS As String = "DailyNeeds"
Private Const DAILY_PUBLISH As String = "DailyPublish"

' ====================================================================
'  PURE
' ====================================================================
' A posting's exception for the Lite Desk's "Posting exceptions" card: "no takers" = Open, past its closing moment,
' nobody signed up (the admin decides: No takers, Cancel or Award Out of Dept); "closing soon" = Open, not closed yet,
' closing within LITE_CLOSING_HOURS, nobody signed up; "" otherwise (any other status, a sign-up, or time to spare).
' Suite postexc.
Public Function PostingExceptionCore(ByVal status As String, ByVal closed As Boolean, ByVal hoursToClose As Double, _
                                     ByVal nSignups As Long) As String
    If Trim$(status) <> "Open" Then Exit Function
    If nSignups > 0 Then Exit Function
    If closed Then
        PostingExceptionCore = "no takers"
    ElseIf hoursToClose <= LITE_CLOSING_HOURS Then
        PostingExceptionCore = "closing soon"
    End If
End Function

' A daily-check mark ("yyyy-mm-dd hh:nn|login") read against today (ISO): "hh:nn (login)" when it was made today,
' "" when not today, blank or malformed. Suite litemark.
Public Function DailyMarkText(ByVal mark As String, ByVal todayIso As String) As String
    Dim p As Long: p = InStr(mark, "|")
    If p = 0 Then Exit Function
    Dim stamp As String: stamp = Left$(mark, p - 1)
    If Len(stamp) < 16 Then Exit Function
    If Left$(stamp, 10) <> todayIso Then Exit Function
    DailyMarkText = Mid$(stamp, 12, 5) & " (" & Mid$(mark, p + 1) & ")"
End Function

' The daily check's collapsed line: all three marks made today -> "Done today: Kronos hh:nn (who); Needs Attention
' hh:nn (who); members' file hh:nn (who)"; "" while any step is not done today (the band shows its three steps).
' Suite litedaily.
Public Function DailyCheckLineCore(ByVal todayIso As String, ByVal kronosMark As String, ByVal needsMark As String, _
                                   ByVal publishMark As String) As String
    Dim k As String, n As String, p As String
    k = DailyMarkText(kronosMark, todayIso)
    n = DailyMarkText(needsMark, todayIso)
    p = DailyMarkText(publishMark, todayIso)
    If k = "" Or n = "" Or p = "" Then Exit Function
    DailyCheckLineCore = "Done today: Kronos " & k & "; Needs Attention " & n & "; members' file " & p
End Function

' How many cells of a bar nCells wide a figure fills against the largest figure: 0 for nothing, at least 1 for
' anything, rounded half up, never past nCells. Suite litebar.
Public Function BarCellsCore(ByVal hours As Double, ByVal maxHours As Double, ByVal nCells As Long) As Long
    If hours <= 0 Or maxHours <= 0 Or nCells <= 0 Then Exit Function
    Dim x As Double: x = hours / maxHours * nCells
    BarCellsCore = Int(x + 0.5)
    If BarCellsCore < 1 Then BarCellsCore = 1
    If BarCellsCore > nCells Then BarCellsCore = nCells
End Function

' ====================================================================
'  THE DAILY CHECK MARKS (hidden names; "anyone has done it today")
' ====================================================================
Private Function LiteNameGet(ByVal key As String) As String
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then Exit Function
    LiteNameGet = Replace(Replace(Mid$(nmObj.RefersTo, 2), """", ""), "'", "")
End Function

Private Sub LiteNameSet(ByVal key As String, ByVal v As String)
    On Error Resume Next                            ' never raises
    Dim nmObj As Name: Set nmObj = ThisWorkbook.Names(key)
    If nmObj Is Nothing Then
        ThisWorkbook.Names.Add Name:=key, RefersTo:="=""" & v & """", Visible:=False
    Else
        nmObj.RefersTo = "=""" & v & """"
    End If
End Sub

' Called when a daily-check step finishes: "kronos" (SyncKronosData, H written), "needs" (ShowNeedsAttention),
' "publish" (PublishMembers, the file saved). Writes "yyyy-mm-dd hh:nn|login"; a practice copy marks too (its Desk
' shows them; nothing else reads them). Never raises; a saved workbook stays marked saved (the mark rides on the next save).
Public Sub LiteMark(ByVal step As String)
    On Error Resume Next                            ' never raises
    Dim key As String
    Select Case LCase$(step)
        Case "kronos": key = DAILY_KRONOS
        Case "needs": key = DAILY_NEEDS
        Case "publish": key = DAILY_PUBLISH
        Case Else: Exit Sub
    End Select
    Dim wasSaved As Boolean: wasSaved = ThisWorkbook.Saved
    LiteNameSet key, Format$(Now, "yyyy-mm-dd hh:nn") & "|" & LCase$(CurrentUser())
    If wasSaved Then ThisWorkbook.Saved = True
End Sub

' One step's text for the band: "hh:nn (login)" when done today, else "". step as LiteMark's.
Public Function DailyStepText(ByVal step As String) As String
    Dim key As String
    Select Case LCase$(step)
        Case "kronos": key = DAILY_KRONOS
        Case "needs": key = DAILY_NEEDS
        Case "publish": key = DAILY_PUBLISH
        Case Else: Exit Function
    End Select
    DailyStepText = DailyMarkText(LiteNameGet(key), Format$(Date, "yyyy-mm-dd"))
End Function

' The collapsed line when all three are done today, else ""
Public Function DailyCheckLine() As String
    DailyCheckLine = DailyCheckLineCore(Format$(Date, "yyyy-mm-dd"), LiteNameGet(DAILY_KRONOS), LiteNameGet(DAILY_NEEDS), LiteNameGet(DAILY_PUBLISH))
End Function

' ====================================================================
'  POSTING EXCEPTIONS
' ====================================================================
' The open postings nobody has signed up for that are past closing or close within LITE_CLOSING_HOURS, oldest closing
' first; oldestId / oldestWhat name the first (the card's button opens the Award pop-up on it). Reads tblPostings and
' tblPostingList once. Never raises (0 on any trouble).
Public Function PostingExceptionCount(Optional ByRef oldestId As String, Optional ByRef oldestWhat As String) As Long
    On Error GoTo Quit
    oldestId = "": oldestWhat = ""
    Dim pv As Variant: pv = LiteTable(PT_TABLE)
    If Not IsArray(pv) Then Exit Function
    ' sign-ups per posting
    Dim signed As Object: Set signed = CreateObject("Scripting.Dictionary")
    signed.CompareMode = vbTextCompare
    Dim lv As Variant: lv = LiteTable(PL_TABLE)
    Dim i As Long, id As String
    If IsArray(lv) Then
        For i = 1 To UBound(lv, 1)
            If CBool(SafeBool(lv(i, PL_SIGNED))) Then
                id = Trim$(SafeStr(lv(i, PL_ID)))
                If id <> "" Then signed(id) = signed(id) + 1
            End If
        Next i
    End If
    Dim cl As Date, what As String, oldest As Date, n As Long, nSigned As Long
    For i = 1 To UBound(pv, 1)
        id = Trim$(SafeStr(pv(i, PS_ID)))
        cl = AsDate(pv(i, PS_CLOSES))
        If id <> "" And cl > 0 Then
            nSigned = 0                               ' (not an IIf: reading a missing Dictionary key would add it)
            If signed.Exists(id) Then nSigned = CLng(signed(id))
            what = PostingExceptionCore(SafeStr(pv(i, PS_STATUS)), PostingClosed(cl), (PostingCloseAt(cl) - Now) * 24, nSigned)
            If what <> "" Then
                n = n + 1
                If oldest = 0 Or cl < oldest Then oldest = cl: oldestId = id: oldestWhat = what
            End If
        End If
    Next i
    PostingExceptionCount = n
Quit:
End Function

' ====================================================================
'  OT HOURS THIS PAY PERIOD BY UNIT (Band 4's bars)
' ====================================================================
' The non-void Hours Charged in the OT Log whose Period Start is the current pay period, summed per unit in DeptList's
' order (+ "(other)" for a unit name the roster does not know, only when it has hours). A 2-D array (1 To n, 1 To 2):
' unit, hours; Empty when the log is empty. Display only; never a rule. Never raises.
Public Function UnitChargedThisPeriod() As Variant
    On Error GoTo Quit
    Dim v As Variant: v = LogRowsNow()
    If Not IsArray(v) Then Exit Function
    Dim depts As Variant: depts = DeptList()
    Dim nD As Long: nD = UBound(depts) - LBound(depts) + 1
    Dim out() As Variant: ReDim out(1 To nD + 1, 1 To 2)
    Dim u As Long
    For u = 1 To nD
        out(u, 1) = CStr(depts(LBound(depts) + u - 1)): out(u, 2) = 0#
    Next u
    out(nD + 1, 1) = "(other)": out(nD + 1, 2) = 0#
    Dim p0 As Date: p0 = PeriodStartFor(Date)
    Dim i As Long, k As Long, hit As Boolean, unit As String
    For i = 1 To UBound(v, 1)
        If UCase$(SafeStr(v(i, L_STATUS))) <> "VOID" And Val(SafeStr(v(i, L_CHARGED))) > 0 Then
            If Int(AsDate(v(i, L_PERIOD))) = Int(p0) Then
                unit = Trim$(SafeStr(v(i, L_DEPT))): hit = False
                For k = 1 To nD
                    If StrComp(out(k, 1), unit, vbTextCompare) = 0 Then out(k, 2) = out(k, 2) + Val(SafeStr(v(i, L_CHARGED))): hit = True: Exit For
                Next k
                If Not hit Then out(nD + 1, 2) = out(nD + 1, 2) + Val(SafeStr(v(i, L_CHARGED)))
            End If
        End If
    Next i
    If out(nD + 1, 2) = 0 Then                       ' drop "(other)" when nothing landed there
        Dim trimmed() As Variant: ReDim trimmed(1 To nD, 1 To 2)
        For u = 1 To nD
            trimmed(u, 1) = out(u, 1)
            trimmed(u, 2) = out(u, 2)
        Next u
        UnitChargedThisPeriod = trimmed
    Else
        UnitChargedThisPeriod = out
    End If
Quit:
End Function

' ====================================================================
'  HELPERS
' ====================================================================
' a table anywhere in the workbook as a 2-D array; Empty when missing or empty. Never raises.
Private Function LiteTable(ByVal tableName As String) As Variant
    On Error Resume Next                            ' never raises
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects(tableName)
        If Not lo Is Nothing Then LiteTable = TableData(lo): Exit Function
    Next ws
End Function

Private Function SafeBool(ByVal v As Variant) As Boolean
    On Error Resume Next                            ' never raises
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then SafeBool = v: Exit Function
    Dim s As String: s = UCase$(Trim$(CStr(v)))
    SafeBool = (s = "TRUE" Or s = "YES" Or s = "Y" Or s = "1" Or s = "X")
End Function
