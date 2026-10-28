Attribute VB_Name = "modRequests"
Option Explicit

' ====================================================================
'  v28 (item 1): MEMBERS' REQUESTS - time off and shift switches asked for from the members' file.
'  The members' file writes one small text file per click into REQUEST_FOLDER (beside the members' file, as the
'  sign-ups do): the request itself, the named partner's agree / refuse (a switch), the requester's withdraw.
'  ReadRequests folds them into tblRequests on the admin-only Requests sheet; Needs Attention lists the ones
'  waiting for an admin; a double-click opens the time-off / switch pop-up PREFILLED (the request id rides in
'  gRequestId; the pop-up's save calls RequestBooked with the AB- / SS- id it became); a right-click declines with
'  a reason. A request writes NOTHING on the roster by itself: the admin's save through the existing pop-ups does,
'  and their rules apply. The publish carries dRequests (no note, no login). Graeme, Oct 25, 2026: admins only;
'  a switch request reaches the admin only once a DIFFERENT login than the requester's has agreed in the file.
' ====================================================================
Public Const RQ_VERSION As String = "Requests v29.0.1 (Oct 28)"

' tblRequests columns
Public Const RQ_ID As Long = 1          ' Request ID  RQ-0001
Public Const RQ_KIND As Long = 2        ' Time off / Switch
Public Const RQ_EMP As Long = 3         ' Employee (the roster's spelling, as the members' file published it)
Public Const RQ_LOGIN As Long = 4       ' the requester's Windows login (never published)
Public Const RQ_FROM As Long = 5
Public Const RQ_TO As Long = 6
Public Const RQ_SHIFT As Long = 7       ' a switch: Days / Nights
Public Const RQ_PARTNER As Long = 8     ' a switch: who would work it
Public Const RQ_PLOGIN As Long = 9      ' the partner's login that agreed (never published)
Public Const RQ_AGREED As Long = 10     ' when the partner agreed
Public Const RQ_NOTE As Long = 11       ' the member's note (never published)
Public Const RQ_STATUS As Long = 12
Public Const RQ_RECEIVED As Long = 13   ' the request file's stamp
Public Const RQ_BY As Long = 14         ' Decided By
Public Const RQ_AT As Long = 15         ' Decided At
Public Const RQ_ANSWER As Long = 16     ' the admin's line to the member (published, screened)
Public Const RQ_REF As Long = 17        ' the AB- / SS- id the request became
Public Const RQ_FILE As Long = 18       ' the files' own id (login-stamp), the key the agree / withdraw files carry
Public Const RQ_N As Long = 18

' the statuses (the members' file shows these words)
Public Const RQ_ST_PARTNER As String = "Waiting for partner"
Public Const RQ_ST_WAITING As String = "Waiting"
Public Const RQ_ST_BOOKED As String = "Booked"
Public Const RQ_ST_DECLINED As String = "Declined"
Public Const RQ_ST_REFUSED As String = "Declined by partner"
Public Const RQ_ST_WITHDRAWN As String = "Withdrawn"
Public Const RQ_ST_EXPIRED As String = "Expired"
Public Const RQ_ST_VOIDED As String = "Booking voided"   ' v29 (item 6; Graeme Oct 27): the absence / switch a Booked request became was voided

' a request file's lines, by position (RequestFileRead)
Public Const RF_REQUEST As Long = 0     ' request=  the files' own id
Public Const RF_KIND As Long = 1
Public Const RF_EMP As Long = 2
Public Const RF_FROM As Long = 3        ' yyyy-mm-dd
Public Const RF_TO As Long = 4
Public Const RF_SHIFT As Long = 5
Public Const RF_PARTNER As Long = 6
Public Const RF_NOTE As Long = 7
Public Const RF_LOGIN As Long = 8
Public Const RF_STAMP As Long = 9       ' yyyy-mm-dd hh:nn:ss
Public Const RF_ACTION As Long = 10     ' request / agree / refuse / withdraw
Public Const RF_N As Long = 10

Private Const RQ_PUBLISH_DAYS As Long = 60    ' decided requests stay in dRequests this many days
Private Const RQ_PREFIX As String = "RQ"

Public gRequestId As String             ' the request a prefilled pop-up is booking ("" = none); RequestBooked reads it

' ---- the table (admin-only Requests sheet, made on first use) ----
Public Function RequestsTable(Optional ByVal create As Boolean = False) As ListObject
    Dim ws As Worksheet, lo As ListObject
    If SheetExists(RQ_SHEET) Then
        Set ws = ThisWorkbook.Worksheets(RQ_SHEET)
        On Error Resume Next
        Set lo = ws.ListObjects(RQ_TABLE)
        On Error GoTo 0
    End If
    If lo Is Nothing And create Then
        If ws Is Nothing Then
            Dim was As Object: Set was = ThisWorkbook.ActiveSheet
            UnlockStructure
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = RQ_SHEET
            RelockStructure
            SetTabColor ws, RGB(0, 128, 70)
            On Error Resume Next
            If Not was Is Nothing Then was.Activate
            On Error GoTo 0
        End If
        UnlockSheet ws
        ws.Range("A1").Resize(1, RQ_N).Value = Array("Request ID", "Kind", "Employee", "Login", "From", "To", "Shift", "Partner", _
                                                      "Partner Login", "Agreed At", "Note", "Status", "Received", "Decided By", _
                                                      "Decided At", "Answer", "Ref", "File")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(2, RQ_N), , xlYes)
        lo.Name = RQ_TABLE
        lo.TableStyle = "TableStyleLight9"
        ws.Columns(RQ_FROM).NumberFormat = "ddd mmm d, yyyy": ws.Columns(RQ_TO).NumberFormat = "ddd mmm d, yyyy"
        ws.Columns(RQ_AGREED).NumberFormat = "mmm d, yyyy h:mm AM/PM": ws.Columns(RQ_RECEIVED).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ws.Columns(RQ_AT).NumberFormat = "mmm d, yyyy h:mm AM/PM"
        ws.Columns(RQ_EMP).ColumnWidth = 24: ws.Columns(RQ_PARTNER).ColumnWidth = 24
        ws.Columns(RQ_NOTE).ColumnWidth = 36: ws.Columns(RQ_ANSWER).ColumnWidth = 36: ws.Columns(RQ_FILE).ColumnWidth = 30
        ProtectSheet ws
        SetSheetVisible ws, IIf(gAdminUnlocked, xlSheetVisible, xlSheetVeryHidden)
    End If
    Set RequestsTable = lo
End Function

' the table as values (Empty when no table or no rows)
Public Function RequestRows() As Variant
    Dim lo As ListObject: Set lo = RequestsTable(False)
    If lo Is Nothing Then Exit Function
    RequestRows = TableData(lo)
End Function

' ---- the folder (REQUEST_FOLDER resolved: ".\Requests" = beside the members' file, as SIGNUP_FOLDER) ----
Public Function RequestFolderPath() As String
    Dim p As String: p = Trim$(REQUEST_FOLDER)
    If p = "" Then Exit Function
    If p = "." Then
        p = PublishFolderPath()
    ElseIf Left$(p, 2) = ".\" Then
        p = PublishFolderPath() & Mid$(p, 3)
    End If
    RequestFolderPath = EnsureSlash(p)
End Function

' Makes the folder (and its done\) if it isn't there. True when it exists afterwards.
Public Function EnsureRequestFolder() As Boolean
    Dim p As String: p = RequestFolderPath()
    If p = "" Then Exit Function
    On Error Resume Next                            ' never raises
    If Not FolderExists(p) Then MkDir Left$(p, Len(p) - 1)
    If FolderExists(p) And Not FolderExists(p & "done\") Then MkDir p & "done"
    On Error GoTo 0
    EnsureRequestFolder = FolderExists(p)
End Function

' ---- the open step (Workbook_Open 10, before the Desk lands): an admin's open of the live roster folds the new files in;
'      nothing is archived (an open never saves); never raises ----
Public Sub RequestsOnOpen()
    On Error Resume Next                            ' never raises
    If Not REQUESTS_ON Then Exit Sub
    If Not IsAdmin() Then Exit Sub
    If Not IsLiveRoster() Then Exit Sub
    ReadRequests False
End Sub

' Alt+F8 / the Desk tile's verb: read the folder now and say what waits
Public Sub ImportRequests()
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    If Not REQUESTS_ON Then MsgBox "Members' requests are switched off (REQUESTS_ON).", vbInformation, "Members' requests": Exit Sub
    Dim n As Long: n = ReadRequests(True)
    Tell IIf(n = 0, "No member's request is waiting for an admin.", n & " request(s) waiting for an admin - Needs Attention lists them (double-click one to book it, right-click to decline)."), _
         vbInformation, "Members' requests"
    Exit Sub
Oops:
    ReportError "ImportRequests"
End Sub

' ---- READ THE FOLDER: every REQ__*.txt folded into the table; returns how many rows are Waiting (for an admin) ----
' A request file makes a row (Waiting, or Waiting for partner for a switch); an agree file from a DIFFERENT login
' moves a switch to Waiting; refuse / withdraw / a To date in the past end it; a Booked / Declined / Booking voided row never changes here.
' The files of an ended request move to done\. Writes the table under UnlockSheet / ProtectSheet; saves when byHand.
Public Function ReadRequests(ByVal byHand As Boolean) As Long
    Dim folder As String: folder = RequestFolderPath()
    If folder = "" Then Exit Function
    If Not EnsureRequestFolder() Then Exit Function
    Dim lo As ListObject: Set lo = RequestsTable(True)
    If lo Is Nothing Then Exit Function
    On Error GoTo Oops                              ' (review) the sheet is re-protected and the error logged, then raised to the caller
    ' 1. the files, grouped by their own id
    Dim files() As String, nF As Long, f As String
    On Error Resume Next
    f = Dir$(folder & "REQ__*.txt")
    Do While f <> "" And Err.Number = 0
        nF = nF + 1
        ReDim Preserve files(1 To nF)
        files(nF) = f
        f = Dir$()
    Loop
    On Error GoTo Oops                              ' (re-review) the handler is back on for the whole fold
    Dim byId As Object: Set byId = CreateObject("Scripting.Dictionary"): byId.CompareMode = 1    ' id -> the request facts (a String())
    Dim agreeOf As Object: Set agreeOf = CreateObject("Scripting.Dictionary"): agreeOf.CompareMode = 1   ' id -> every agree as "login|stamp;login|stamp" (the requester's own is set aside below)
    Dim refusedOf As Object: Set refusedOf = CreateObject("Scripting.Dictionary"): refusedOf.CompareMode = 1
    Dim withdrawnOf As Object: Set withdrawnOf = CreateObject("Scripting.Dictionary"): withdrawnOf.CompareMode = 1
    Dim i As Long, fx() As String, rid As String, why As String
    For i = 1 To nF
        If RequestFileRead(folder & files(i), fx) Then
            rid = fx(RF_REQUEST)
            If rid <> "" Then
                why = RequestCheckCore(fx(RF_KIND), fx(RF_EMP), fx(RF_FROM), fx(RF_TO), fx(RF_SHIFT), fx(RF_PARTNER), fx(RF_ACTION))
                If why = "" Then
                    Select Case fx(RF_ACTION)
                        Case "request"
                            If Not byId.Exists(rid) Then byId.Add rid, fx
                        Case "agree"
                            agreeOf(rid) = SafeStr(agreeOf(rid)) & IIf(SafeStr(agreeOf(rid)) = "", "", ";") & fx(RF_LOGIN) & "|" & fx(RF_STAMP)
                        Case "refuse"
                            refusedOf(rid) = SafeStr(refusedOf(rid)) & ";" & fx(RF_LOGIN) & ";"       ' every login that refused (re-review)
                        Case "withdraw"
                            withdrawnOf(rid) = SafeStr(withdrawnOf(rid)) & ";" & fx(RF_LOGIN) & ";"   ' every login that withdrew
                    End Select
                End If
            End If
        End If
    Next i
    ' 2. the table: the rows the files already made, by their id
    Dim v As Variant: v = TableData(lo)
    Dim rowIdx As Object: Set rowIdx = CreateObject("Scripting.Dictionary"): rowIdx.CompareMode = 1
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If SafeStr(v(i, RQ_FILE)) <> "" Then rowIdx(SafeStr(v(i, RQ_FILE))) = i
        Next i
    End If
    UnlockSheet lo.Parent
    Dim kk As Variant, r As Long, st As String, agreeLogin As String, agreeAt As Date, reqLogin As String, kind As String
    Dim toD As Date, fromD As Date, newRows As Long
    ' 2a. new requests -> new rows
    For Each kk In byId.Keys
        If Not rowIdx.Exists(CStr(kk)) Then
            fx = byId(kk)
            IsoDate fx(RF_FROM), fromD: IsoDate fx(RF_TO), toD
            Dim newRow As ListRow
            If lo.DataBodyRange Is Nothing Then
                Set newRow = lo.ListRows.Add
            ElseIf lo.ListRows.Count = 1 And Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
                Set newRow = lo.ListRows(1)                      ' the blank row a new table is made with (as OpensOnOpen)
            Else
                Set newRow = lo.ListRows.Add
            End If
            r = newRow.Index
            With lo.DataBodyRange
                .Cells(r, RQ_ID).Value = NextRequestId()
                .Cells(r, RQ_KIND).Value = fx(RF_KIND)
                .Cells(r, RQ_EMP).Value = fx(RF_EMP)
                .Cells(r, RQ_LOGIN).Value = fx(RF_LOGIN)
                .Cells(r, RQ_FROM).Value = fromD
                .Cells(r, RQ_TO).Value = toD
                .Cells(r, RQ_SHIFT).Value = fx(RF_SHIFT)
                .Cells(r, RQ_PARTNER).Value = fx(RF_PARTNER)
                .Cells(r, RQ_NOTE).NumberFormat = "@"
                .Cells(r, RQ_NOTE).Value = Left$(fx(RF_NOTE), 200)
                .Cells(r, RQ_STATUS).Value = RQ_ST_WAITING             ' set below
                .Cells(r, RQ_RECEIVED).Value = StampValue(fx(RF_STAMP))
                .Cells(r, RQ_FILE).NumberFormat = "@"
                .Cells(r, RQ_FILE).Value = CStr(kk)
            End With
            newRows = newRows + 1
        End If
    Next kk
    If newRows > 0 Then
        v = TableData(lo)
        rowIdx.RemoveAll
        For i = 1 To UBound(v, 1)
            If SafeStr(v(i, RQ_FILE)) <> "" Then rowIdx(SafeStr(v(i, RQ_FILE))) = i
        Next i
    End If
    ' 2b. every row still OPEN (Waiting / Waiting for partner / blank): the status from the files (agree / refuse /
    '     withdraw) and the calendar (expired). Booked, Declined, Declined by partner, Withdrawn, Expired and (v29) Booking voided are final
    '     (review: a final row was recomputed once its files were archived and came back as Waiting)
    Dim nWaiting As Long, agreeAll As Variant, aj As Long, aParts As Variant, wd As Boolean, rf As Boolean
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            rid = SafeStr(v(i, RQ_FILE))
            st = SafeStr(v(i, RQ_STATUS))
            If (st = RQ_ST_WAITING Or st = RQ_ST_PARTNER Or st = "") And rid <> "" Then
                kind = SafeStr(v(i, RQ_KIND)): reqLogin = SafeStr(v(i, RQ_LOGIN))
                agreeLogin = "": agreeAt = 0
                If SafeStr(v(i, RQ_PLOGIN)) <> "" Then agreeLogin = SafeStr(v(i, RQ_PLOGIN))        ' agreed on an earlier read
                If agreeLogin = "" And agreeOf.Exists(rid) Then
                    agreeAll = Split(CStr(agreeOf(rid)), ";")                                        ' the first agree from a login OTHER than the requester's counts
                    For aj = 0 To UBound(agreeAll)
                        aParts = Split(CStr(agreeAll(aj)) & "|", "|")
                        If StrComp(Trim$(CStr(aParts(0))), Trim$(reqLogin), vbTextCompare) <> 0 And Trim$(CStr(aParts(0))) <> "" Then
                            agreeLogin = CStr(aParts(0)): agreeAt = StampValue(CStr(aParts(1)))
                            Exit For
                        End If
                    Next aj
                End If
                wd = False: rf = False
                If withdrawnOf.Exists(rid) Then wd = (InStr(1, CStr(withdrawnOf(rid)), ";" & reqLogin & ";", vbTextCompare) > 0)   ' some withdraw from the login that asked
                If refusedOf.Exists(rid) Then rf = LoginOtherThan(CStr(refusedOf(rid)), reqLogin)                                 ' some refuse from another login (the partner)
                toD = 0
                If IsDate(v(i, RQ_TO)) Then toD = CDate(v(i, RQ_TO))
                st = RequestStatusCore(kind, agreeLogin, reqLogin, rf, wd, "", Format$(toD, "yyyy-mm-dd"), Format$(Date, "yyyy-mm-dd"))
                If st <> SafeStr(v(i, RQ_STATUS)) Then lo.DataBodyRange.Cells(i, RQ_STATUS).Value = st
                If agreeLogin <> "" And SafeStr(v(i, RQ_PLOGIN)) = "" Then
                    lo.DataBodyRange.Cells(i, RQ_PLOGIN).Value = agreeLogin
                    If agreeAt > 0 Then lo.DataBodyRange.Cells(i, RQ_AGREED).Value = agreeAt
                End If
            End If
            If st = RQ_ST_WAITING Then nWaiting = nWaiting + 1
        Next i
    End If
    ProtectSheet lo.Parent
    ' (review) the files of an ended request leave the folder only once the table that says so is SAVED: an open
    ' never saves, so it archives nothing (the next read redoes the same work); the by-hand read saves, then archives
    If byHand Then
        If SaveIfWritable() Then ArchiveEndedRequests
    End If
    ReadRequests = nWaiting
    Exit Function
Oops:
    Dim en As Long, em As String: en = Err.Number: em = Err.Description
    On Error Resume Next
    ProtectSheet lo.Parent
    LogError "ReadRequests", en, em
    On Error GoTo 0                                 ' (re-review) or the raise below would be swallowed by Resume Next
    Err.Raise en, , em
End Function

' every ended request's files -> done\, called only after a SAVE went through (never raises)
Public Sub ArchiveEndedRequests()
    On Error Resume Next                            ' never raises
    Dim folder As String: folder = RequestFolderPath()
    If folder = "" Then Exit Sub
    If Not FolderExists(folder & "done\") Then Exit Sub
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Sub
    Dim i As Long, st As String
    For i = 1 To UBound(v, 1)
        st = SafeStr(v(i, RQ_STATUS))
        If st = RQ_ST_BOOKED Or st = RQ_ST_DECLINED Or st = RQ_ST_REFUSED Or st = RQ_ST_WITHDRAWN Or st = RQ_ST_EXPIRED Or st = RQ_ST_VOIDED Then
            If SafeStr(v(i, RQ_FILE)) <> "" Then ArchiveRequestFiles folder, RequestFilesNamed(folder, SafeStr(v(i, RQ_FILE)))
        End If
    Next i
End Sub

' True when one of the ";"-wrapped logins differs from the requester's (a refuse from the partner counts; the requester's own does not)
Private Function LoginOtherThan(ByVal logins As String, ByVal requesterLogin As String) As Boolean
    Dim parts As Variant, i As Long
    parts = Split(logins, ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            If StrComp(Trim$(CStr(parts(i))), Trim$(requesterLogin), vbTextCompare) <> 0 Then LoginOtherThan = True: Exit Function
        End If
    Next i
End Function

' the files of an ended request -> done\ (never raises; a file still open elsewhere stays and is skipped next time)
Private Sub ArchiveRequestFiles(ByVal folder As String, ByVal fileList As String)
    On Error Resume Next                            ' never raises
    Dim parts As Variant, i As Long
    parts = Split(fileList, ";")
    For i = 0 To UBound(parts)
        If Trim$(CStr(parts(i))) <> "" Then
            If Len(Dir$(folder & CStr(parts(i)))) > 0 Then Name folder & CStr(parts(i)) As folder & "done\" & CStr(parts(i))
        End If
    Next i
End Sub

' RQ-0001, RQ-0002 ... : the table's highest + 1, never below the IdMark
Private Function NextRequestId() As String
    Dim v As Variant: v = RequestRows()
    Dim mx As Long, i As Long, s As String, k As Long
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            s = UCase$(SafeStr(v(i, RQ_ID)))
            If Left$(s, 3) = RQ_PREFIX & "-" Then
                k = CLng(Val(Mid$(s, 4)))
                If k > mx Then mx = k
            End If
        Next i
    End If
    If IdMark(RQ_PREFIX) > mx Then mx = IdMark(RQ_PREFIX)
    SetIdMark RQ_PREFIX, mx + 1
    NextRequestId = RQ_PREFIX & "-" & Format$(mx + 1, "0000")
End Function

' ---- one file: key=value lines into fx(RF_*) (False = could not be opened: the PC still writing it, antivirus) ----
Public Function RequestFileRead(ByVal fullPath As String, ByRef fx() As String) As Boolean
    ReDim fx(0 To RF_N)
    Dim ff As Integer, ln As String, p As Long, valTxt As String, opened As Boolean, nLines As Long
    On Error GoTo Quit
    ff = FreeFile
    Open fullPath For Input Access Read Shared As #ff
    opened = True
    Do While Not EOF(ff) And nLines < 60
        Line Input #ff, ln
        nLines = nLines + 1
        p = InStr(ln, "=")
        If p > 1 Then
            valTxt = Trim$(Mid$(ln, p + 1))
            Select Case LCase$(Trim$(Left$(ln, p - 1)))
                Case "request": fx(RF_REQUEST) = Left$(valTxt, 80)
                Case "kind": fx(RF_KIND) = valTxt
                Case "employee": fx(RF_EMP) = Left$(valTxt, 80)
                Case "from": fx(RF_FROM) = valTxt
                Case "to": fx(RF_TO) = valTxt
                Case "shift": fx(RF_SHIFT) = valTxt
                Case "partner": fx(RF_PARTNER) = Left$(valTxt, 80)
                Case "note": fx(RF_NOTE) = Left$(valTxt, 400)
                Case "login": fx(RF_LOGIN) = Left$(valTxt, 40)
                Case "stamp": fx(RF_STAMP) = valTxt
                Case "action": fx(RF_ACTION) = LCase$(valTxt)
            End Select
        End If
    Loop
    Close #ff
    RequestFileRead = True
    Exit Function
Quit:
    On Error Resume Next
    If opened Then Close #ff
End Function

' PURE (suite reqcheck): a file's facts -> "" (taken) or why it is not. A request needs a kind, a name, dates (ISO,
' To >= From); a switch is ONE shift (To = From), Days / Nights, with a partner who is not the requester. An agree /
' refuse / withdraw needs only the action (its id ties it to the request). An unknown action is refused.
Public Function RequestCheckCore(ByVal kind As String, ByVal emp As String, ByVal fromTxt As String, ByVal toTxt As String, _
                                 ByVal shiftLbl As String, ByVal partner As String, ByVal action As String) As String
    Dim dF As Date, dT As Date
    Select Case action
        Case "agree", "refuse", "withdraw"
            Exit Function
        Case "request"
        Case Else
            RequestCheckCore = "unknown action '" & action & "'": Exit Function
    End Select
    If kind <> "Time off" And kind <> "Switch" Then RequestCheckCore = "unknown kind '" & kind & "'": Exit Function
    If Trim$(emp) = "" Then RequestCheckCore = "no name": Exit Function
    If Not IsoDate(fromTxt, dF) Then RequestCheckCore = "the From date can't be read": Exit Function
    If Not IsoDate(toTxt, dT) Then RequestCheckCore = "the To date can't be read": Exit Function
    If dT < dF Then RequestCheckCore = "To is before From": Exit Function
    If kind = "Switch" Then
        If dT <> dF Then RequestCheckCore = "a switch is one shift": Exit Function
        If shiftLbl <> "Days" And shiftLbl <> "Nights" Then RequestCheckCore = "a switch names Days or Nights": Exit Function
        If Trim$(partner) = "" Then RequestCheckCore = "a switch names a partner": Exit Function
        If StrComp(Trim$(partner), Trim$(emp), vbTextCompare) = 0 Then RequestCheckCore = "the partner is the requester": Exit Function
    End If
End Function

' PURE (suite reqstatus): the status from what the files say. decided = "Booked" / "Declined" / "Booking voided" (an
' admin's word wins; v29: a void of the booking too); withdrawn / refused end it; a To date before today = Expired; a
' switch with no agree from a login OTHER than the requester's = Waiting for partner; else Waiting. Dates as yyyy-mm-dd
' ("" = none).
Public Function RequestStatusCore(ByVal kind As String, ByVal agreeLogin As String, ByVal requesterLogin As String, _
                                  ByVal refused As Boolean, ByVal withdrawn As Boolean, ByVal decided As String, _
                                  ByVal toTxt As String, ByVal todayTxt As String) As String
    Dim dT As Date, dN As Date
    If decided = RQ_ST_BOOKED Or decided = RQ_ST_DECLINED Or decided = RQ_ST_VOIDED Then RequestStatusCore = decided: Exit Function
    If withdrawn Then RequestStatusCore = RQ_ST_WITHDRAWN: Exit Function
    If refused Then RequestStatusCore = RQ_ST_REFUSED: Exit Function
    If IsoDate(toTxt, dT) And IsoDate(todayTxt, dN) Then
        If dT < dN Then RequestStatusCore = RQ_ST_EXPIRED: Exit Function
    End If
    If kind = "Switch" Then
        If Trim$(agreeLogin) = "" Then RequestStatusCore = RQ_ST_PARTNER: Exit Function
        If StrComp(Trim$(agreeLogin), Trim$(requesterLogin), vbTextCompare) = 0 Then RequestStatusCore = RQ_ST_PARTNER: Exit Function
    End If
    RequestStatusCore = RQ_ST_WAITING
End Function

' "yyyy-mm-dd" -> a date, digits in every place (no CDate: the files come from other PCs). False = not a date.
Public Function IsoDate(ByVal s As String, ByRef d As Date) As Boolean
    s = Trim$(s): d = 0
    If Len(s) <> 10 Then Exit Function
    If Not (s Like "####-##-##") Then Exit Function
    Dim m As Long, dd As Long: m = CLng(Mid$(s, 6, 2)): dd = CLng(Right$(s, 2))
    If m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    d = DateSerial(CLng(Left$(s, 4)), m, dd)
    IsoDate = (Day(d) = dd)                         ' Feb 31 rolls over: not a date
End Function

' "yyyy-mm-dd hh:nn:ss" -> date + time (0 when it can't be read)
Private Function StampValue(ByVal s As String) As Date
    Dim d As Date
    s = Trim$(s)
    If Len(s) < 19 Then Exit Function
    If Not IsoDate(Left$(s, 10), d) Then Exit Function
    If Not (Mid$(s, 11) Like " ##:##:##*") Then Exit Function
    On Error Resume Next                            ' never raises
    StampValue = d + TimeSerial(CLng(Mid$(s, 12, 2)), CLng(Mid$(s, 15, 2)), CLng(Mid$(s, 18, 2)))
End Function

' ---- what waits (the Desk tile, Needs Attention) ----
Public Function RequestsWaitingCount() As Long
    If Not REQUESTS_ON Then Exit Function
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, RQ_STATUS)) = RQ_ST_WAITING Then RequestsWaitingCount = RequestsWaitingCount + 1
    Next i
End Function

' one row's words for Needs Attention: "Time off Nov 3 - Nov 5" / "Switch Nov 3 Days with Sam K (agreed)" + the note
Public Function RequestLineText(ByRef v As Variant, ByVal r As Long) As String
    Dim s As String, dF As Date, dT As Date
    If IsDate(v(r, RQ_FROM)) Then dF = CDate(v(r, RQ_FROM))
    If IsDate(v(r, RQ_TO)) Then dT = CDate(v(r, RQ_TO))
    If SafeStr(v(r, RQ_KIND)) = "Switch" Then
        s = "Switch " & Format$(dF, "ddd mmm d") & " " & SafeStr(v(r, RQ_SHIFT)) & " with " & CalName(SafeStr(v(r, RQ_PARTNER))) & _
            IIf(SafeStr(v(r, RQ_PLOGIN)) <> "", " (agreed)", "")
    Else
        s = "Time off " & Format$(dF, "ddd mmm d") & IIf(dT > dF, " - " & Format$(dT, "ddd mmm d"), "")
    End If
    If SafeStr(v(r, RQ_NOTE)) <> "" Then s = s & "   " & Chr$(34) & Left$(SafeStr(v(r, RQ_NOTE)), 60) & Chr$(34)
    RequestLineText = s & "   " & SafeStr(v(r, RQ_ID))
End Function

' ---- the admin acts: a double-click (book) or a right-click (decline) on a Needs Attention row ----
Public Sub RequestAct(ByVal id As String, ByVal decline As Boolean)
    On Error GoTo Oops
    If Not RequireAdmin() Then Exit Sub
    If Not RequireWritable() Then Exit Sub
    Dim lo As ListObject: Set lo = RequestsTable(False)
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then MsgBox id & " is not a request.", vbExclamation, "Members' requests": Exit Sub
    If SafeStr(v(r, RQ_STATUS)) <> RQ_ST_WAITING Then
        MsgBox id & " is " & SafeStr(v(r, RQ_STATUS)) & " - nothing to do.", vbInformation, "Members' requests": Exit Sub
    End If
    Dim emp As String: emp = SafeStr(v(r, RQ_EMP))
    Dim dF As Date, dT As Date
    If IsDate(v(r, RQ_FROM)) Then dF = CDate(v(r, RQ_FROM))
    If IsDate(v(r, RQ_TO)) Then dT = CDate(v(r, RQ_TO))
    If decline Then
        Dim why As String
        why = AskRequired("Decline " & id & " (" & CalName(emp) & ": " & RequestLineText(v, r) & ")?" & vbNewLine & vbNewLine & _
                          "The reason below is published in the members' file, where every member can read it: no leave words, no phone numbers or emails.", "Decline a request")
        If why = "" Then Exit Sub
        If LeaveWordsIn(why) <> "" Then
            MsgBox "The reason holds a leave word (" & LeaveWordsIn(why) & ") and would reach every member. Word it another way.", vbExclamation, "Decline a request"
            Exit Sub
        End If
        If PrivateValueKind(why) <> "" Then          ' (review MED-4) a phone-like number or an address would refuse every publish for 60 days
            MsgBox "The reason looks like it holds " & PrivateValueKind(why) & ", which the members' file never carries. Word it another way.", vbExclamation, "Decline a request"
            Exit Sub
        End If
        MarkDecided id, RQ_ST_DECLINED, why, ""
        If SaveIfWritable() Then ArchiveEndedRequests
        ShowReceiptNow id & " declined", _
                       "  " & id & "  " & SafeStr(v(r, RQ_KIND)) & "  " & emp & "  " & RequestLineText(v, r) & vbNewLine, _
                       "  Nothing on the roster changed." & vbNewLine, _
                       "  Requests sheet (admin): Declined, with your reason. The member sees it after the next members' file update." & vbNewLine, _
                       "  Nothing to undo: the member may ask again." & vbNewLine
        Exit Sub
    End If
    ' book it: the pop-up, prefilled; its save calls RequestBooked with the AB- / SS- id
    gRequestId = id
    If SafeStr(v(r, RQ_KIND)) = "Switch" Then
        If SwitchTable(True) Is Nothing Then gRequestId = "": Exit Sub
        frmShiftSwitch.Prefill emp, SafeStr(v(r, RQ_PARTNER)), dF, SafeStr(v(r, RQ_SHIFT)), HomeUnitCached(emp), id
        frmShiftSwitch.Show
    Else
        If AbsTable(True) Is Nothing Then gRequestId = "": Exit Sub
        If CovTable(True) Is Nothing Then gRequestId = "": Exit Sub
        frmTimeOff.Prefill emp, dF, IIf(dT > 0, dT, dF), id        ' v29: dTo = 0 would mean open-ended (Fix an entry); a request has two dates
        frmTimeOff.Show
        RefreshCalendar
    End If
    gRequestId = ""
    Exit Sub
Oops:
    gRequestId = ""
    ReportError "RequestAct"
End Sub

' the pop-ups call this after a save: the request a prefilled pop-up was booking becomes Booked with that ref - ONLY
' when what was saved is that request (the same person and date, and for a switch the same partner working it; review
' MED-1: an admin may save something else from the prefilled pop-up); then the file is saved and the files archived
' v29 (item 6, Graeme Oct 27 - the two v28 gaps): shiftLbl = the shift saved (Days / Nights), compared for a switch too
' (v28 matched the date and the partner only); the past-sick path of Book Time Off now calls this with the SC- id
Public Sub RequestBooked(ByVal ref As String, ByVal emp As String, ByVal d As Date, ByVal worksName As String, _
                         Optional ByVal shiftLbl As String = "")
    On Error Resume Next                            ' never raises (the booking itself is saved; this is the bookkeeping)
    If gRequestId = "" Then Exit Sub
    Dim id As String: id = gRequestId
    gRequestId = ""
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Sub
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Sub
    Dim same As Boolean
    same = (StrComp(Trim$(SafeStr(v(r, RQ_EMP))), Trim$(emp), vbTextCompare) = 0)
    If same And SafeStr(v(r, RQ_KIND)) = "Switch" Then
        If IsDate(v(r, RQ_FROM)) Then same = (Int(CDate(v(r, RQ_FROM))) = Int(d))
        If same Then same = (StrComp(Trim$(SafeStr(v(r, RQ_PARTNER))), Trim$(worksName), vbTextCompare) = 0)
        If same And Trim$(shiftLbl) <> "" And Trim$(SafeStr(v(r, RQ_SHIFT))) <> "" Then
            same = (StrComp(Trim$(SafeStr(v(r, RQ_SHIFT))), Trim$(shiftLbl), vbTextCompare) = 0)
        End If
    ElseIf same Then
        ' time off: the booking's start may sit inside the asked range (re-review MED-2: the pop-up offers to trim the
        ' days off at the ends - Nov 3-9 asked, booked from Nov 4 - and that IS the request)
        If IsDate(v(r, RQ_FROM)) And IsDate(v(r, RQ_TO)) Then same = (Int(d) >= Int(CDate(v(r, RQ_FROM))) And Int(d) <= Int(CDate(v(r, RQ_TO))))
    End If
    If Not same Then
        Dim stNow As String: stNow = SafeStr(v(r, RQ_STATUS))      ' (re-review LOW-6) a re-entry's request is "Booking voided", not waiting
        Tell id & " stays " & LCase$(stNow) & ": what you saved (" & emp & ", " & Format$(d, "ddd mmm d") & IIf(worksName <> "", ", " & worksName & " works", "") & _
             IIf(shiftLbl <> "", ", " & shiftLbl, "") & ") is not that request." & _
             IIf(stNow = RQ_ST_WAITING Or stNow = RQ_ST_PARTNER, " Book it from Needs Attention again, or decline it.", " Tell the member how it stands."), vbExclamation, "Members' requests"
        Exit Sub
    End If
    MarkDecided id, RQ_ST_BOOKED, "", ref
    If SaveIfWritable() Then ArchiveEndedRequests
End Sub

' v29 (item 6, Graeme Oct 27): VoidAbsence / VoidSwitch call this with the id they voided - every Booked request whose
' Ref is that id becomes "Booking voided" (final; the members' file shows the word at the next publish; archived with the
' others). The Answer is the plain words "booking voided": the VOID REASON IS NEVER PUBLISHED (CLAUDE.md Privacy - void
' notes stay in the master; the review found "in hospital, see HR" would have reached every member). A past-sick booking
' from a request is several SC- rows: the request is marked only when NO live SC- row for that person is left inside
' its From..To (one voided morning out of five leaves the request Booked). Returns the request id(s) marked, ";"-joined
' ("" = none) - the status bar names them for the admin. The void saves. Never raises.
Public Function RequestBookingVoided(ByVal ref As String, ByVal why As String) As String
    On Error Resume Next                            ' never raises
    ref = UCase$(Trim$(ref))
    If ref = "" Then Exit Function
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, ids As String, nextSc As String
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, RQ_STATUS)) = RQ_ST_BOOKED And UCase$(Trim$(SafeStr(v(i, RQ_REF)))) = ref Then
            If Left$(ref, 3) = "SC-" Then
                nextSc = OtherSickDayLive(SafeStr(v(i, RQ_EMP)), v(i, RQ_FROM), v(i, RQ_TO), ref)
                If nextSc <> "" Then                  ' a day of several voided: still booked - the Ref moves to the next live day (re-review MED-1)
                    MarkDecided SafeStr(v(i, RQ_ID)), RQ_ST_BOOKED, "", nextSc
                    GoTo NextRow
                End If
            End If
            MarkDecided SafeStr(v(i, RQ_ID)), RQ_ST_VOIDED, "booking voided", ""
            ids = ids & IIf(ids = "", "", ";") & SafeStr(v(i, RQ_ID))
            MarkCalDirty                            ' the Requests sheet changed (the Desk's figures redraw on activate)
        End If
NextRow:
    Next i
    If ids <> "" Then ShowBarText "Members' request " & Replace(ids, ";", ", ") & " now reads ""Booking voided"" for the member (the next publish shows it)."
    RequestBookingVoided = ids
End Function

' the earliest live SC- row for emp, other than ref, starting inside from..to (a request booked as several past sick days):
' its id, or "" when none is left
Private Function OtherSickDayLive(ByVal emp As String, ByVal fromV As Variant, ByVal toV As Variant, ByVal ref As String) As String
    On Error Resume Next                            ' never raises
    If Not IsDate(fromV) Then Exit Function
    Dim d0 As Date, d1 As Date: d0 = Int(CDate(fromV))
    If IsDate(toV) Then d1 = Int(CDate(toV)) Else d1 = d0
    Dim v As Variant: v = AbsTableData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, id As String, sd As Date, best As Date
    For i = 1 To UBound(v, 1)
        id = UCase$(Trim$(SafeStr(v(i, 1))))
        If Left$(id, 3) = "SC-" And id <> ref And UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" Then
            If StrComp(Trim$(SafeStr(v(i, AB_EMP))), Trim$(emp), vbTextCompare) = 0 And IsDate(v(i, AB_START)) Then
                sd = Int(CDate(v(i, AB_START)))
                If sd >= d0 And sd <= d1 Then
                    If best = 0 Or sd < best Then best = sd: OtherSickDayLive = id
                End If
            End If
        End If
    Next i
End Function

' v29 (item 6, review MED-4): the Booked request whose Ref is this id ("" = none) - Fix an entry reads it BEFORE the void,
' so the re-entry re-books the same request with the new Ref (gRequestId -> RequestBooked). Never raises.
Public Function RequestIdForRef(ByVal ref As String) As String
    On Error Resume Next                            ' never raises
    ref = UCase$(Trim$(ref))
    If ref = "" Then Exit Function
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Function
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, RQ_STATUS)) = RQ_ST_BOOKED And UCase$(Trim$(SafeStr(v(i, RQ_REF)))) = ref Then RequestIdForRef = SafeStr(v(i, RQ_ID)): Exit Function
    Next i
End Function

' v29 (item 6): the SC- id of the past sick day booked for emp starting on d (the newest when several) - Book Time Off's
' sick path marks its request with it. "" when none. Never raises.
Public Function SickRefFor(ByVal emp As String, ByVal d As Date, Optional ByVal dTo As Date = 0, Optional ByRef startAt As Date) As String
    ' review MED-5: the EARLIEST live SC- start inside d..dTo (the first asked day may have had no scheduled shift);
    ' the last row among equals (the newest id); startAt = that day (RequestBooked's date test)
    On Error Resume Next                            ' never raises
    If dTo = 0 Then dTo = d
    startAt = 0
    Dim v As Variant: v = AbsTableData()
    If Not IsArray(v) Then Exit Function
    Dim i As Long, id As String, sd As Date, best As Date
    For i = 1 To UBound(v, 1)
        id = UCase$(Trim$(SafeStr(v(i, 1))))
        If Left$(id, 3) = "SC-" And UCase$(SafeStr(v(i, AB_STATUS))) <> "VOID" Then
            If StrComp(Trim$(SafeStr(v(i, AB_EMP))), Trim$(emp), vbTextCompare) = 0 And IsDate(v(i, AB_START)) Then
                sd = Int(CDate(v(i, AB_START)))
                If sd >= Int(d) And sd <= Int(dTo) Then
                    If best = 0 Or sd <= best Then best = sd: SickRefFor = id: startAt = sd
                End If
            End If
        End If
    Next i
End Function

' ---- IntegrityCheck 17 "Requests" (v29 item 6; Graeme Oct 27: Check while the request's To date is today or later, Note once
'      past, never PROBLEM): every Booked request's Ref must name a live absence (AB- / SC-) or switch (SS-) for the same
'      people. Rows of (severity 1 = Check / 0 = Note, text, fix); Empty when every one is fine. Reads only. ----
Public Function RequestRefProblemList() As Variant
    Dim v As Variant: v = RequestRows()
    If Not IsArray(v) Then Exit Function
    Dim absV As Variant: absV = AbsTableData()
    Dim swV As Variant: swV = TableData(SwitchTable(False))
    Dim out() As Variant: ReDim out(1 To 40, 1 To 3)
    Dim n As Long, i As Long, j As Long, ref As String, kind As String, found As Boolean, tStatus As String, namesOk As Boolean, who As String
    Dim prob As String, toTxt As String
    For i = 1 To UBound(v, 1)
        If SafeStr(v(i, RQ_STATUS)) = RQ_ST_BOOKED Then
            ref = UCase$(Trim$(SafeStr(v(i, RQ_REF)))): kind = SafeStr(v(i, RQ_KIND))
            found = False: tStatus = "": namesOk = True: who = ""
            If ref <> "" Then
                If kind = "Switch" Then
                    If IsArray(swV) Then
                        For j = 1 To UBound(swV, 1)
                            If UCase$(Trim$(SafeStr(swV(j, SW_ID)))) = ref Then
                                found = True: tStatus = SafeStr(swV(j, SW_STATUS))
                                namesOk = (StrComp(Trim$(SafeStr(swV(j, SW_OFF))), Trim$(SafeStr(v(i, RQ_EMP))), vbTextCompare) = 0 And _
                                           StrComp(Trim$(SafeStr(swV(j, SW_WORKS))), Trim$(SafeStr(v(i, RQ_PARTNER))), vbTextCompare) = 0)
                                who = SafeStr(swV(j, SW_OFF)) & " / " & SafeStr(swV(j, SW_WORKS))
                                Exit For
                            End If
                        Next j
                    End If
                Else
                    If IsArray(absV) Then
                        For j = 1 To UBound(absV, 1)
                            If UCase$(Trim$(SafeStr(absV(j, 1)))) = ref Then
                                found = True: tStatus = SafeStr(absV(j, AB_STATUS))
                                namesOk = (StrComp(Trim$(SafeStr(absV(j, AB_EMP))), Trim$(SafeStr(v(i, RQ_EMP))), vbTextCompare) = 0)
                                who = SafeStr(absV(j, AB_EMP))
                                Exit For
                            End If
                        Next j
                    End If
                End If
            End If
            toTxt = "": If IsDate(v(i, RQ_TO)) Then toTxt = Format$(CDate(v(i, RQ_TO)), "yyyy-mm-dd")
            ' (review LOW-9) a request whose dates ended more than RQ_PUBLISH_DAYS ago is history: the members' file no
            ' longer shows it and the table never loses a row, so it is not a Note for ever
            If toTxt <> "" Then
                If CDate(v(i, RQ_TO)) < Date - RQ_PUBLISH_DAYS Then GoTo NextReq
            End If
            prob = RequestRefProblem(kind, RQ_ST_BOOKED, ref, found, tStatus, namesOk, toTxt, Format$(Date, "yyyy-mm-dd"))
            If prob <> "" Then
                n = n + 1
                If n <= 40 Then
                    out(n, 1) = IIf(Left$(prob, 5) = "Check", 1, 0)
                    out(n, 2) = SafeStr(v(i, RQ_ID)) & " (" & SafeStr(v(i, RQ_EMP)) & ", " & kind & IIf(toTxt <> "", " to " & toTxt, "") & ") is Booked, but " & _
                                RefProblemWords(Mid$(prob, InStr(prob, "|") + 1), ref, kind, who)
                    out(n, 3) = "Tell " & SafeStr(v(i, RQ_EMP)) & " how the time off / switch stands; book it again if it should be live (the request itself cannot be re-opened)"
                End If
            End If
        End If
NextReq:
    Next i
    If n = 0 Then Exit Function
    If n > 40 Then
        out(40, 2) = "... and " & (n - 39) & " more Booked requests with a Ref to look at (the first 39 are listed)": out(40, 3) = "Work through these first; the rest show once they are gone"
        n = 40
    End If
    Dim res() As Variant: ReDim res(1 To n, 1 To 3)
    For i = 1 To n
        For j = 1 To 3
            res(i, j) = out(i, j)
        Next j
    Next i
    RequestRefProblemList = res
End Function

' the check's plain words for each case (review LOW-10)
Private Function RefProblemWords(ByVal what As String, ByVal ref As String, ByVal kind As String, ByVal who As String) As String
    Select Case what
        Case "no ref": RefProblemWords = "has no Ref (nothing is recorded for it)"
        Case "wrong kind": RefProblemWords = "its Ref " & ref & " is the wrong kind (" & IIf(kind = "Switch", "a switch needs SS-", "time off needs AB- or SC-") & ")"
        Case "missing": RefProblemWords = "its Ref " & ref & " is not in the " & IIf(kind = "Switch", "Shift Switches", "Absences") & " table"
        Case "void": RefProblemWords = "its Ref " & ref & " is VOID (the void did not reach this request: a hand edit?)"
        Case "wrong person": RefProblemWords = "its Ref " & ref & " names someone else" & IIf(who <> "", " (" & who & ")", "") & " - a reused id"
        Case Else: RefProblemWords = "its Ref " & ref & " " & what
    End Select
End Function

' PURE (suite reqref): one Booked request's Ref against its target. "" when fine; else "<Check|Note>|<what>" with what =
' "no ref" (blank), "wrong kind" (a Time off request not AB- / SC-, a Switch not SS-), "missing" (no such row), "void" (the
' row is VOID), "wrong person" (the row names someone else: a reused id). Check while the request's To date (yyyy-mm-dd)
' is today or later, Note once it is past ("" dates = Check). A status other than Booked is never flagged.
Public Function RequestRefProblem(ByVal kind As String, ByVal status As String, ByVal ref As String, ByVal targetFound As Boolean, _
                                  ByVal targetStatus As String, ByVal namesMatch As Boolean, ByVal toTxt As String, ByVal todayTxt As String) As String
    If status <> RQ_ST_BOOKED Then Exit Function
    Dim what As String
    ref = UCase$(Trim$(ref))
    If ref = "" Then
        what = "no ref"
    ElseIf kind = "Switch" And Left$(ref, 3) <> "SS-" Then
        what = "wrong kind"
    ElseIf kind <> "Switch" And Left$(ref, 3) <> "AB-" And Left$(ref, 3) <> "SC-" Then
        what = "wrong kind"
    ElseIf Not targetFound Then
        what = "missing"
    ElseIf UCase$(Trim$(targetStatus)) = "VOID" Then
        what = "void"
    ElseIf Not namesMatch Then
        what = "wrong person"
    Else
        Exit Function
    End If
    Dim dT As Date, dN As Date, past As Boolean
    If IsoDate(toTxt, dT) And IsoDate(todayTxt, dN) Then past = (dT < dN)
    RequestRefProblem = IIf(past, "Note", "Check") & "|" & what
End Function

Private Sub MarkDecided(ByVal id As String, ByVal status As String, ByVal answer As String, ByVal ref As String)
    Dim lo As ListObject: Set lo = RequestsTable(False)
    If lo Is Nothing Then Exit Sub
    Dim v As Variant: v = TableData(lo)
    Dim r As Long: r = RowOf(v, id)
    If r = 0 Then Exit Sub
    UnlockSheet lo.Parent
    With lo.DataBodyRange
        .Cells(r, RQ_STATUS).Value = status
        .Cells(r, RQ_BY).Value = CurrentUser()
        .Cells(r, RQ_AT).Value = Now
        If answer <> "" Then
            .Cells(r, RQ_ANSWER).NumberFormat = "@": .Cells(r, RQ_ANSWER).Value = answer
        ElseIf status = RQ_ST_BOOKED Then
            .Cells(r, RQ_ANSWER).ClearContents       ' v29 (re-review MED-2): a re-booking after a void does not keep "booking voided"
        End If
        If ref <> "" Then .Cells(r, RQ_REF).Value = ref
    End With
    ProtectSheet lo.Parent
End Sub

' the folder's files carrying this request's id in their name (";"-joined)
Private Function RequestFilesNamed(ByVal folder As String, ByVal rid As String) As String
    On Error Resume Next                            ' never raises
    Dim f As String
    f = Dir$(folder & "REQ__" & rid & "__*.txt")
    Do While f <> "" And Err.Number = 0
        RequestFilesNamed = RequestFilesNamed & IIf(RequestFilesNamed = "", "", ";") & f
        f = Dir$()
    Loop
End Function

' ---- the publish: dRequests = Employee, Kind, From, To, Shift, Partner, Status, Answer, Decided (no note, no login);
'      every open request, and the decided ones for RQ_PUBLISH_DAYS ----
Public Function PublishRequestsTable() As Variant
    Dim v As Variant: v = RequestRows()
    Dim cap As Long: cap = 1
    If IsArray(v) Then cap = UBound(v, 1) + 1
    Dim out() As Variant: ReDim out(1 To cap, 1 To 9)
    out(1, 1) = "Employee": out(1, 2) = "Kind": out(1, 3) = "From": out(1, 4) = "To": out(1, 5) = "Shift"
    out(1, 6) = "Partner": out(1, 7) = "Status": out(1, 8) = "Answer": out(1, 9) = "Decided"
    Dim k As Long: k = 1
    Dim i As Long, st As String, keep As Boolean
    If IsArray(v) And REQUESTS_ON Then
        For i = 1 To UBound(v, 1)
            st = SafeStr(v(i, RQ_STATUS))
            keep = (st = RQ_ST_WAITING Or st = RQ_ST_PARTNER)
            If Not keep And IsDate(v(i, RQ_AT)) Then keep = (CDate(v(i, RQ_AT)) >= Date - RQ_PUBLISH_DAYS)
            If Not keep And IsDate(v(i, RQ_RECEIVED)) Then keep = (CDate(v(i, RQ_RECEIVED)) >= Date - RQ_PUBLISH_DAYS)
            If keep And SafeStr(v(i, RQ_EMP)) <> "" Then
                k = k + 1
                out(k, 1) = SafeStr(v(i, RQ_EMP)): out(k, 2) = SafeStr(v(i, RQ_KIND))
                out(k, 3) = v(i, RQ_FROM): out(k, 4) = v(i, RQ_TO): out(k, 5) = SafeStr(v(i, RQ_SHIFT))
                out(k, 6) = SafeStr(v(i, RQ_PARTNER)): out(k, 7) = st: out(k, 8) = SafeStr(v(i, RQ_ANSWER))
                If IsDate(v(i, RQ_AT)) Then out(k, 9) = v(i, RQ_AT) Else out(k, 9) = ""
            End If
        Next i
    End If
    Dim res() As Variant: ReDim res(1 To k, 1 To 9)
    Dim j As Long
    For i = 1 To k
        For j = 1 To 9
            res(i, j) = out(i, j)
        Next j
    Next i
    PublishRequestsTable = res
End Function

' the publish's screen: an Answer with a leave word, or a Kind that is not one of the two public words
Public Function RequestWordsIn(ByRef a As Variant) As String
    If Not IsArray(a) Then Exit Function
    Dim i As Long, w As String
    For i = 2 To UBound(a, 1)
        w = LeaveWordsIn(SafeStr(a(i, 8)))
        If w <> "" Then RequestWordsIn = "dRequests row " & i & " (" & SafeStr(a(i, 1)) & "): the answer holds '" & w & "'": Exit Function
        If SafeStr(a(i, 2)) <> "Time off" And SafeStr(a(i, 2)) <> "Switch" Then
            RequestWordsIn = "dRequests row " & i & ": kind '" & SafeStr(a(i, 2)) & "'": Exit Function
        End If
    Next i
End Function
