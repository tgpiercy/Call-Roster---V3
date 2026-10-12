Attribute VB_Name = "modUpdater"
Option Explicit

' ====================================================================
'  ROSTER UPDATER  (lives in its OWN workbook, never in the roster)
'
'  Make the updater workbook once: a new blank workbook, Alt+F11,
'  File > Import File > modUpdater.bas, save it as "Roster Updater.xlsm"
'  (macro-enabled). To use it, open it AND the roster, then Alt+F8:
'
'  UpdateFromFolder   - point it at an unzipped release folder (the one
'      holding modules\, paste-in\, member\ and INSTALL.md). It saves a
'      backup copy of the roster, then replaces the code of every
'      module, of ThisWorkbook and of every form with the files in that
'      folder. New modules are imported. Form designs are never
'      touched. The new code is NOT saved yet: saving stays blocked
'      until Debug > Compile is clean and you run UpdateFinished.
'  UpdateFinished     - after a clean Compile: lifts the block and saves.
'  UpdateDrop         - Compile showed an error: closes the roster
'      WITHOUT saving. The file on disk is untouched.
'  CheckUpdateFolder  - the same list, nothing changed.
'  UpdateMemberTemplate - opens the members' template next to the
'      roster, replaces its code from member\, runs its
'      BuildMemberTemplate, saves and closes it.
'
'  Needs Excel's "Trust access to the VBA project object model"
'  (File > Options > Trust Center > Trust Center Settings > Macro
'  Settings).
'
'  WHY ITS OWN WORKBOOK: v1 ran inside the roster and crashed Excel
'  (Oct 12, 2026) while rewriting the project it was running in. From
'  here none of the roster's code is running while its modules change.
'  The macros refuse to run if this module is found inside a roster.
'
'  RULES FOR THIS MODULE
'   - It calls the roster's code only by name and only when none of it
'     is being changed: the admin test before anything is touched, and
'     ApplyView in UpdateFinished after the admin says Compile was clean.
'   - It keeps no module-level variables.
'   - A file whose module name is modUpdater is skipped.
'   - The roster's no-save flag (the hidden name its own
'     Workbook_BeforeSave / BeforeClose read) is set BEFORE the first
'     change and stays set until UpdateFinished. A failure puts the old
'     code back, so that flag still has working code behind it.
' ====================================================================

Public Const UPD_VERSION As String = "Updater v2 (Jan 5)"

Private Const UPD_TITLE As String = "Update the roster code"
Private Const UPD_TEMPLATE As String = "Call Out Roster - Members TEMPLATE.xlsm"
Private Const UPD_NOSAVE As String = "NoSaveFlag"                 ' same name modOvertimeEntry.MarkNoSave uses
Private Const UPD_PENDING As String = "UpdatePending"             ' hidden name: code swapped, waiting for Compile + UpdateFinished

' VBIDE component types (late bound, no reference)
Private Const UPD_CT_MODULE As Long = 1
Private Const UPD_CT_FORM As Long = 3
Private Const UPD_CT_DOCUMENT As Long = 100

Private Type UpdItem
    Kind As String          ' module / form / workbook
    FileName As String      ' as in the folder
    FullPath As String
    CompName As String      ' the component it goes into
    CodeText As String      ' the file's text, Attribute lines taken off
    Action As String        ' replace / add
    OldText As String       ' what the component held before (put back if anything fails)
    Swapped As Boolean
    Result As String        ' filled by the run
End Type


' ====================================================================
'  THE MACROS
' ====================================================================
Public Sub UpdateFromFolder()
    UpdRun False
End Sub

Public Sub CheckUpdateFolder()
    UpdRun True
End Sub

Public Sub UpdateMemberTemplate()
    Dim wb As Workbook: Set wb = UpdTarget()
    If wb Is Nothing Then Exit Sub
    If Not UpdIsAdmin(wb) Then Exit Sub
    Dim folder As String: folder = UpdPickFolder("member")
    If folder = "" Then Exit Sub
    UpdTemplate wb, folder
End Sub

' After Debug > Compile came back clean: lifts the save block the update set,
' puts the admin view back through the NEW code and saves.
Public Sub UpdateFinished()
    Dim wb As Workbook: Set wb = UpdTarget()
    If wb Is Nothing Then Exit Sub
    If Not UpdNameSet(wb, UPD_PENDING) Then
        MsgBox "There is no update waiting to be finished in " & wb.Name & ".", vbInformation, UPD_TITLE
        Exit Sub
    End If
    If wb.ReadOnly Then
        MsgBox "This copy is read-only, so it can't be saved. Close it; nothing is kept.", vbExclamation, UPD_TITLE
        Exit Sub
    End If
    If Not Application.EnableEvents Then
        MsgBox "Excel's events are switched off in this session (an earlier macro was stopped part-way)." & vbNewLine & _
               "Run Alt+F8 > UpdateDrop, close Excel completely, open the roster and run the update again.", vbExclamation, UPD_TITLE
        Exit Sub
    End If
    If Not UpdIsAdmin(wb) Then Exit Sub
    If MsgBox("Roster: " & wb.Name & vbNewLine & vbNewLine & _
              "Did Debug > Compile VBAProject (with the ROSTER's project selected) finish with NO message?" & vbNewLine & vbNewLine & _
              "Yes = keep the new code and save." & vbNewLine & _
              "No = stop here. Close the workbook; nothing is saved and your file stays as it was.", _
              vbYesNo + vbQuestion + vbDefaultButton2, UPD_TITLE) <> vbYes Then Exit Sub

    ' the new code must run at all before anything is saved (this also puts the admin view's state back)
    Dim en As Long, em As String
    On Error Resume Next
    Application.Run "'" & Replace(wb.Name, "'", "''") & "'!ApplyView", True
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    If en <> 0 Then
        MsgBox "The new code didn't run (" & em & "), so nothing was saved and saving is still blocked." & vbNewLine & vbNewLine & _
               "Close the workbook (Don't Save if Excel asks) and send me that message.", vbCritical, UPD_TITLE
        Exit Sub
    End If

    UpdDropName wb, UPD_NOSAVE
    UpdDropName wb, UPD_PENDING
    On Error Resume Next
    wb.Save
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    If en <> 0 Or Not wb.Saved Then
        MsgBox "The save didn't go through" & IIf(em <> "", " (" & em & ")", "") & ". Press Ctrl+S and check for a message.", vbExclamation, UPD_TITLE
        Exit Sub
    End If
    MsgBox "Saved with the new code." & vbNewLine & vbNewLine & _
           "NEXT:" & vbNewLine & _
           "1. Close the workbook and open it again." & vbNewLine & _
           "2. Alt+F8 > SystemCheck." & vbNewLine & _
           "3. Run the macros the install note lists under 'run once'." & vbNewLine & _
           "4. Alt+F8 > UpdateMemberTemplate (if the release has a members' part), then Member View." & vbNewLine & vbNewLine & _
           "Keep the Roster Updater workbook for the next release; it needs no saving.", vbInformation, UPD_TITLE
End Sub

' Compile showed an error, or you changed your mind: closes this workbook
' WITHOUT saving, whatever state its code is in. The file on disk is untouched.
Public Sub UpdateDrop()
    Dim wb As Workbook: Set wb = UpdTarget()
    If wb Is Nothing Then Exit Sub
    If MsgBox("Close " & wb.Name & " WITHOUT saving?" & vbNewLine & vbNewLine & _
              "The file on disk stays exactly as it was before the update." & vbNewLine & _
              "(If a 'Compile error' box appears while it closes, click OK: it still closes without saving.)", _
              vbOKCancel + vbExclamation + vbDefaultButton2, UPD_TITLE) <> vbOK Then Exit Sub
    wb.Close SaveChanges:=False
End Sub


' ====================================================================
'  THE ROSTER
' ====================================================================
Private Sub UpdRun(ByVal checkOnly As Boolean)
    Dim wb As Workbook: Set wb = UpdTarget()
    If wb Is Nothing Then Exit Sub
    Dim why As String

    If wb.Path = "" Or LCase$(Left$(wb.Path, 4)) = "http" Then
        MsgBox "Open the roster from an ordinary folder (a drive letter or a network share) before updating it." & vbNewLine & _
               "It can't be updated while it is open from SharePoint / OneDrive or was never saved.", vbExclamation, UPD_TITLE
        Exit Sub
    End If
    If Not checkOnly Then
        If wb.ReadOnly Then
            MsgBox "This copy is open read-only, so it can't be updated." & vbNewLine & _
                   "Close it and open it with edit access.", vbExclamation, UPD_TITLE
            Exit Sub
        End If
        If Not Application.EnableEvents Then
            MsgBox "Excel's events are switched off in this session (an earlier macro was stopped part-way)." & vbNewLine & _
                   "Close Excel completely, open the roster again and run the update.", vbExclamation, UPD_TITLE
            Exit Sub
        End If
        If UpdNameSet(wb, UPD_PENDING) Then
            MsgBox "An update is already waiting in this workbook." & vbNewLine & vbNewLine & _
                   "Debug > Compile, then Alt+F8 > UpdateFinished. Or close the workbook to drop it (nothing is saved).", vbExclamation, UPD_TITLE
            Exit Sub
        End If
        If UpdNameSet(wb, UPD_NOSAVE) Then
            MsgBox "Saving is already blocked in this workbook (an earlier check failed)." & vbNewLine & _
                   "Close it WITHOUT saving, reopen it, and run the update again.", vbExclamation, UPD_TITLE
            Exit Sub
        End If
        If Not wb.Saved Then
            MsgBox wb.Name & " has changes that aren't saved yet." & vbNewLine & _
                   "Save it (Ctrl+S), then run the update again: if the update has to be dropped, nothing since the last save is kept.", vbExclamation, UPD_TITLE
            Exit Sub
        End If
    End If
    If Not UpdIsAdmin(wb) Then Exit Sub

    Dim proj As Object: Set proj = UpdProject(wb, why)
    If proj Is Nothing Then
        MsgBox why, vbExclamation, UPD_TITLE
        Exit Sub
    End If

    Dim folder As String: folder = UpdPickFolder("modules")
    If folder = "" Then Exit Sub

    ' ---- read every file first: a bad folder stops here, before anything is changed ----
    Dim items() As UpdItem, n As Long
    If Not UpdGatherRoster(folder, proj, wb, items, n, why) Then
        MsgBox "Nothing was changed." & vbNewLine & vbNewLine & why, vbExclamation, UPD_TITLE
        Exit Sub
    End If

    Dim relNew As String, relNow As String, i As Long
    For i = 1 To n
        If StrComp(items(i).CompName, "modSystem", vbTextCompare) = 0 Then
            relNew = UpdBetween(items(i).CodeText, "EXPECTED_RELEASE As String = """, """")
        End If
    Next i
    relNow = UpdCurrentRelease(proj)

    Dim nRep As Long, nAdd As Long, addTxt As String
    For i = 1 To n
        If items(i).Action = "add" Then
            nAdd = nAdd + 1
            addTxt = addTxt & "  + " & items(i).CompName & vbNewLine
        Else
            nRep = nRep + 1
        End If
    Next i
    Dim planTxt As String
    planTxt = "Roster: " & wb.Name & vbNewLine & "Folder: " & folder & vbNewLine & _
              "The roster is on " & IIf(relNow = "", "(unknown)", relNow) & "; the folder holds " & _
              IIf(relNew = "", "(unknown)", relNew) & "." & vbNewLine & vbNewLine & _
              nRep & " to replace (modules, forms, ThisWorkbook), " & nAdd & " new." & vbNewLine & addTxt

    Dim rep As String
    rep = UPD_VERSION & "   " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbNewLine & wb.FullName & vbNewLine & planTxt & vbNewLine

    If checkOnly Then
        For i = 1 To n
            rep = rep & items(i).Action & "  " & items(i).CompName & "  <-  " & items(i).FileName & vbNewLine
        Next i
        rep = rep & UpdLeftAlone(proj, items, n)
        MsgBox "CHECK ONLY - nothing was changed." & vbNewLine & vbNewLine & planTxt & vbNewLine & _
               "The full list is in:" & vbNewLine & UpdWriteReport(wb, rep, "Update check"), vbInformation, UPD_TITLE
        Exit Sub
    End If

    If MsgBox(planTxt & vbNewLine & _
              "A backup copy is saved first. Form designs are not touched." & vbNewLine & _
              "Nothing is saved until you have run Debug > Compile and then UpdateFinished." & vbNewLine & _
              "Don't touch Excel until the next message." & vbNewLine & vbNewLine & _
              "Update now?", vbOKCancel + vbQuestion, UPD_TITLE) <> vbOK Then Exit Sub

    ' ---- backup ----
    Dim bk As String
    bk = UpdSlash(wb.Path) & "Backup before update " & Format$(Now, "yyyy-mm-dd hhmmss") & " - " & wb.Name
    On Error Resume Next
    wb.SaveCopyAs bk
    why = ""
    If Err.Number <> 0 Then why = Err.Description
    On Error GoTo 0
    If why <> "" Then
        MsgBox "The backup copy couldn't be saved, so nothing was changed:" & vbNewLine & why, vbCritical, UPD_TITLE
        Exit Sub
    End If
    rep = rep & "Backup: " & bk & vbNewLine & vbNewLine

    ' ---- block saving BEFORE the first change; it stays blocked until UpdateFinished ----
    UpdSetName wb, UPD_NOSAVE
    UpdSetName wb, UPD_PENDING
    If Not UpdNameSet(wb, UPD_NOSAVE) Or Not UpdNameSet(wb, UPD_PENDING) Then
        UpdDropName wb, UPD_NOSAVE
        UpdDropName wb, UPD_PENDING
        MsgBox "The save block couldn't be set, so nothing was changed.", vbCritical, UPD_TITLE
        Exit Sub
    End If

    ' ---- swap the code (VBA editor calls raise no workbook events, so events stay as they are) ----
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim nFail As Long, backTxt As String
    Application.ScreenUpdating = False
    nFail = UpdApply(proj, items, n)
    If nFail > 0 Then backTxt = UpdRollBack(proj, items, n)
    Application.ScreenUpdating = scrWas

    For i = 1 To n
        rep = rep & IIf(items(i).Result = "", "not reached", items(i).Result) & "  " & items(i).CompName & "  <-  " & items(i).FileName & vbNewLine
    Next i
    rep = rep & UpdLeftAlone(proj, items, n)

    Dim repPath As String
    If nFail > 0 Then
        UpdDropName wb, UPD_PENDING                 ' nothing to finish; the save block stays
        rep = rep & vbNewLine & "FAILED. " & IIf(backTxt = "", "The old code was put back.", "Putting the old code back ALSO failed: " & backTxt) & vbNewLine
        repPath = UpdWriteReport(wb, rep, "Update report")
        Dim failTxt As String
        For i = 1 To n
            If Left$(items(i).Result, 6) = "FAILED" Then failTxt = failTxt & "  " & items(i).CompName & ": " & Mid$(items(i).Result, 9) & vbNewLine
        Next i
        MsgBox "The update did NOT go in:" & vbNewLine & failTxt & vbNewLine & _
               IIf(backTxt = "", "The old code was put back.", "Putting the old code back also failed (" & backTxt & ").") & vbNewLine & _
               "Saving is blocked. Close this workbook now; if Excel asks, click DON'T SAVE." & vbNewLine & _
               "(If you try to save, the roster says 'conversion failed verification': that is this block.)" & vbNewLine & vbNewLine & _
               "The file on disk is as it was. Backup copy: " & bk & vbNewLine & vbNewLine & _
               "Then install by hand from INSTALL.md, and send me the report:" & vbNewLine & repPath, vbCritical, UPD_TITLE
        Exit Sub
    End If

    rep = rep & vbNewLine & "Code swapped. Waiting for Compile and UpdateFinished." & vbNewLine
    repPath = UpdWriteReport(wb, rep, "Update report")
    MsgBox "The code is now " & IIf(relNew = "", "the folder's release", relNew) & ": " & nRep & " replaced, " & nAdd & " added." & vbNewLine & _
           "It is NOT saved yet. Do nothing else in the workbook until you finish:" & vbNewLine & vbNewLine & _
           "1. Alt+F11, click the ROSTER's project in the list on the left (" & wb.Name & ")," & vbNewLine & _
           "    then Debug > Compile VBAProject." & vbNewLine & _
           "2. NO message: Alt+F8 > UpdateFinished. It saves and tells you what is next." & vbNewLine & _
           "    A message: note the line it highlights, then Alt+F8 > UpdateDrop." & vbNewLine & _
           "    It closes without saving; your file stays exactly as it was." & vbNewLine & _
           "    (Closing by hand instead? Click DON'T SAVE.)" & vbNewLine & vbNewLine & _
           "Backup: " & bk & vbNewLine & "Report: " & repPath, vbInformation, UPD_TITLE
End Sub

' Reads the roster files of a release folder into items(1..n). False (why) = not usable.
Private Function UpdGatherRoster(ByVal folder As String, ByVal proj As Object, ByVal wb As Workbook, _
                                 ByRef items() As UpdItem, ByRef n As Long, ByRef why As String) As Boolean
    Dim names() As String, k As Long, i As Long, j As Long, base As String
    n = 0
    ReDim items(1 To 200)

    k = UpdListFiles(folder & "modules\", "*.bas", names)
    If k = 0 Then why = "No .bas files in " & folder & "modules\": Exit Function
    For i = 1 To k
        If Not UpdAddItem(items, n, "module", folder & "modules\", names(i), "", proj, why) Then Exit Function
    Next i

    If Not UpdFileExists(folder & "paste-in\ThisWorkbook.txt") Then
        why = "paste-in\ThisWorkbook.txt isn't in " & folder
        Exit Function
    End If
    If Not UpdAddItem(items, n, "workbook", folder & "paste-in\", "ThisWorkbook.txt", wb.CodeName, proj, why) Then Exit Function

    k = UpdListFiles(folder & "paste-in\", "frm*.txt", names)
    For i = 1 To k
        base = Left$(names(i), Len(names(i)) - 4)
        If Not UpdAddItem(items, n, "form", folder & "paste-in\", names(i), base, proj, why) Then Exit Function
    Next i

    ' two files for one component (a second download, "modPeople (1).bas") would overwrite each other
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(items(i).CompName, items(j).CompName, vbTextCompare) = 0 Then
                why = items(i).FileName & " and " & items(j).FileName & " are both '" & items(i).CompName & "'. Remove the extra file from the folder."
                Exit Function
            End If
        Next j
    Next i

    ' it must look like a roster release
    Dim hasCore As Boolean, hasSys As Boolean
    For i = 1 To n
        If StrComp(items(i).CompName, "modOvertimeEntry", vbTextCompare) = 0 Then hasCore = True
        If StrComp(items(i).CompName, "modSystem", vbTextCompare) = 0 Then hasSys = True
    Next i
    If Not (hasCore And hasSys) Then
        why = "That folder doesn't hold a roster release (modOvertimeEntry.bas and modSystem.bas aren't in its modules folder)."
        Exit Function
    End If
    UpdGatherRoster = True
End Function

' One file -> one item, checked. compName "" = take it from the file's VB_Name line.
' A file that is this module is skipped (True, nothing added).
Private Function UpdAddItem(ByRef items() As UpdItem, ByRef n As Long, ByVal kind As String, ByVal dirPath As String, _
                            ByVal fileName As String, ByVal compName As String, ByVal proj As Object, _
                            ByRef why As String) As Boolean
    Dim txt As String, comp As Object, wantType As Long
    If n >= UBound(items) Then why = "Too many files in the folder.": Exit Function
    If Not UpdReadFile(dirPath & fileName, txt) Then why = "Couldn't read " & dirPath & fileName: Exit Function
    If Asc(Left$(txt, 1)) > 126 Or InStr(txt, vbCrLf) = 0 Then
        why = fileName & " isn't a plain text file with Windows line ends. Unzip the release again."
        Exit Function
    End If
    If compName = "" Then compName = UpdBetween(txt, "Attribute VB_Name = """, """")
    If compName = "" Then compName = Left$(fileName, InStrRev(fileName, ".") - 1)
    If StrComp(compName, "modUpdater", vbTextCompare) = 0 Then       ' never itself, whatever the file is called
        UpdAddItem = True
        Exit Function
    End If
    txt = UpdStripAttributes(txt)
    If InStr(1, txt, "Option Explicit", vbTextCompare) = 0 Or Len(txt) < 20 Then
        why = fileName & " doesn't look like roster code (no Option Explicit)."
        Exit Function
    End If

    Select Case kind
        Case "module": wantType = UPD_CT_MODULE
        Case "form": wantType = UPD_CT_FORM
        Case Else: wantType = UPD_CT_DOCUMENT
    End Select
    Set comp = UpdFindComp(proj, compName)
    If comp Is Nothing Then
        If kind = "workbook" Then why = "The workbook's own code module (" & compName & ") wasn't found.": Exit Function
        If kind = "form" Then
            why = "The form " & compName & " isn't in this workbook yet. Add it first (Alt+F11, Insert > UserForm, " & _
                  "set its Name to " & compName & "), then run the update again. See the install note."
            Exit Function
        End If
    ElseIf comp.Type <> wantType Then
        why = "'" & compName & "' exists in this workbook but isn't a " & kind & ". Rename or remove it first."
        Exit Function
    End If
    n = n + 1
    items(n).Kind = kind
    items(n).FileName = fileName
    items(n).FullPath = dirPath & fileName
    items(n).CompName = compName
    items(n).CodeText = txt
    items(n).Action = IIf(comp Is Nothing, "add", "replace")
    items(n).OldText = ""
    items(n).Swapped = False
    items(n).Result = ""
    UpdAddItem = True
End Function

' Does the work. Stops at the first failure. Returns the number of failures (0 or 1).
Private Function UpdApply(ByVal proj As Object, ByRef items() As UpdItem, ByVal n As Long) As Long
    Dim i As Long, comp As Object, cm As Object, why As String, ok As Boolean
    On Error GoTo Failed
    For i = 1 To n
        why = "": ok = False
        If items(i).Action = "replace" Then
            Set comp = UpdFindComp(proj, items(i).CompName)
            If comp Is Nothing Then
                why = "the component disappeared"
            Else
                Set cm = comp.CodeModule
                If cm.CountOfLines > 0 Then items(i).OldText = cm.Lines(1, cm.CountOfLines)
                items(i).Swapped = True
                ok = UpdSwapCode(comp, items(i).CodeText, why)
            End If
            If ok Then items(i).Result = "replaced"
        Else
            items(i).Swapped = True
            ok = UpdImport(proj, items(i).FullPath, items(i).CompName, why)
            If ok Then items(i).Result = "added   "
        End If
        If Not ok Then
            items(i).Result = "FAILED: " & why
            UpdApply = 1
            Exit Function
        End If
    Next i
    Exit Function
Failed:
    If i >= 1 And i <= n Then items(i).Result = "FAILED: " & Err.Description
    UpdApply = 1
End Function

' After a failure: every component touched so far gets its old code back, and
' modules that were imported are emptied and removed. "" = all back as it was.
Private Function UpdRollBack(ByVal proj As Object, ByRef items() As UpdItem, ByVal n As Long) As String
    Dim i As Long, why As String, bad As String, newRes As String
    For i = n To 1 Step -1
        If items(i).Swapped Then
            why = UpdUndoOne(proj, items(i).CompName, items(i).Action, items(i).OldText, newRes)
            If why <> "" Then
                bad = bad & items(i).CompName & " (" & why & ") "
            ElseIf Left$(items(i).Result, 6) <> "FAILED" Then
                items(i).Result = newRes
            End If
        End If
    Next i
    UpdRollBack = Trim$(bad)
End Function

' One component back to what it was. "" = done; otherwise why not. An error here
' never stops the others from being put back.
Private Function UpdUndoOne(ByVal proj As Object, ByVal compName As String, ByVal action As String, _
                            ByVal oldText As String, ByRef newRes As String) As String
    Dim comp As Object, why As String
    On Error GoTo Failed
    Set comp = UpdFindComp(proj, compName)
    If action = "replace" Then
        If comp Is Nothing Then UpdUndoOne = "gone": Exit Function
        If Not UpdSwapCode(comp, oldText, why) Then UpdUndoOne = why: Exit Function
        newRes = "put back"
    Else
        If Not comp Is Nothing Then
            If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
            proj.VBComponents.Remove comp
        End If
        newRes = "taken out"
    End If
    Exit Function
Failed:
    UpdUndoOne = Err.Description
End Function

' Modules and forms in the workbook that the folder has no file for (information only)
Private Function UpdLeftAlone(ByVal proj As Object, ByRef items() As UpdItem, ByVal n As Long) As String
    Dim comp As Object, i As Long, hit As Boolean, s As String
    On Error Resume Next
    For Each comp In proj.VBComponents
        If comp.Type = UPD_CT_MODULE Or comp.Type = UPD_CT_FORM Then
            hit = False
            For i = 1 To n
                If StrComp(items(i).CompName, comp.Name, vbTextCompare) = 0 Then hit = True: Exit For
            Next i
            If Not hit Then s = s & "left alone  " & comp.Name & vbNewLine
        End If
    Next comp
    UpdLeftAlone = s
End Function

' EXPECTED_RELEASE as the workbook's modSystem has it now ("" = can't tell)
Private Function UpdCurrentRelease(ByVal proj As Object) As String
    Dim comp As Object, cm As Object, i As Long, nLines As Long, ln As String
    On Error Resume Next
    Set comp = UpdFindComp(proj, "modSystem")
    If comp Is Nothing Then Exit Function
    Set cm = comp.CodeModule
    nLines = cm.CountOfLines
    If nLines > 80 Then nLines = 80
    For i = 1 To nLines
        ln = cm.Lines(i, 1)
        If InStr(1, ln, "EXPECTED_RELEASE As String", vbTextCompare) > 0 Then
            UpdCurrentRelease = UpdBetween(ln, "= """, """")
            Exit Function
        End If
    Next i
End Function


' ====================================================================
'  THE MEMBERS' TEMPLATE
' ====================================================================
Private Sub UpdTemplate(ByVal roster As Workbook, ByVal folder As String)
    Const TTL As String = "Update the members' template"
    Dim why As String, modTxt As String, wbTxt As String, wantVer As String
    Dim tPath As String: tPath = UpdSlash(roster.Path) & UPD_TEMPLATE

    If Not UpdFileExists(tPath) Then
        MsgBox "The members' template wasn't found next to the roster:" & vbNewLine & tPath, vbExclamation, TTL
        Exit Sub
    End If
    If Not UpdReadFile(folder & "member\modMember.bas", modTxt) Or Not UpdReadFile(folder & "member\ThisWorkbook.txt", wbTxt) Then
        MsgBox "member\modMember.bas and member\ThisWorkbook.txt must both be in:" & vbNewLine & folder, vbExclamation, TTL
        Exit Sub
    End If
    If Asc(Left$(modTxt, 1)) > 126 Or InStr(modTxt, vbCrLf) = 0 Or Asc(Left$(wbTxt, 1)) > 126 Or InStr(wbTxt, vbCrLf) = 0 Then
        MsgBox "The files in " & folder & "member\ aren't plain text with Windows line ends. Unzip the release again. Nothing was changed.", vbExclamation, TTL
        Exit Sub
    End If
    wantVer = UpdBetween(modTxt, "MEM_VERSION As String = """, """")
    modTxt = UpdStripAttributes(modTxt)
    wbTxt = UpdStripAttributes(wbTxt)
    If InStr(1, modTxt, "Option Explicit", vbTextCompare) = 0 Or InStr(1, wbTxt, "Option Explicit", vbTextCompare) = 0 Or wantVer = "" Then
        MsgBox "The files in " & folder & "member\ don't look like the members' code. Nothing was changed.", vbExclamation, TTL
        Exit Sub
    End If

    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.Name, UPD_TEMPLATE, vbTextCompare) = 0 Then
            MsgBox "The template is open. Close it (without saving), then run UpdateMemberTemplate again.", vbExclamation, TTL
            Exit Sub
        End If
    Next w

    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim secWas As Long: secWas = Application.AutomationSecurity
    Dim daWas As Boolean: daWas = Application.DisplayAlerts
    Dim tw As Workbook, tproj As Object, comp As Object, ok As Boolean, built As String
    Dim savedIt As Boolean, opened As Boolean
    On Error GoTo Failed
    Application.EnableEvents = False               ' the template's own open code must not run yet
    Application.AutomationSecurity = 1             ' msoAutomationSecurityLow: its macros load, so its rebuild can run
    Set tw = Workbooks.Open(Filename:=tPath, UpdateLinks:=0)
    opened = True
    Application.AutomationSecurity = secWas
    If tw.ReadOnly Then Err.Raise vbObjectError + 901, , "the template opened read-only (is it open on another PC?)"
    Set tproj = UpdProject(tw, why)
    If tproj Is Nothing Then Err.Raise vbObjectError + 902, , why

    ' a leftover copy from an earlier import (modMember1 ...) would make every name ambiguous: empty it and remove it
    Dim extras As String, nm As Variant
    For Each comp In tproj.VBComponents
        If comp.Type = UPD_CT_MODULE Then
            If LCase$(comp.Name) Like "modmember?*" Then extras = extras & "|" & comp.Name
        End If
    Next comp
    For Each nm In Split(Mid$(extras, 2), "|")
        If CStr(nm) <> "" Then
            Set comp = UpdFindComp(tproj, CStr(nm))
            If Not comp Is Nothing Then
                If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
                tproj.VBComponents.Remove comp
            End If
        End If
    Next nm

    Set comp = UpdFindComp(tproj, "modMember")
    If comp Is Nothing Then
        ok = UpdImport(tproj, folder & "member\modMember.bas", "modMember", why)
    ElseIf comp.Type <> UPD_CT_MODULE Then
        ok = False: why = "'modMember' in the template isn't a module"
    Else
        ok = UpdSwapCode(comp, modTxt, why)
    End If
    If Not ok Then Err.Raise vbObjectError + 903, , "modMember: " & why
    Set comp = UpdFindComp(tproj, tw.CodeName)
    If comp Is Nothing Then Err.Raise vbObjectError + 904, , "the template's own code module wasn't found"
    If Not UpdSwapCode(comp, wbTxt, why) Then Err.Raise vbObjectError + 905, , "ThisWorkbook: " & why

    ' rebuild its sheets with ITS code. It shows its own confirm and its own "Template built" message;
    ' that message says to save it - it is saved here.
    Application.EnableEvents = True
    tw.Activate
    Application.Run "'" & Replace(tw.Name, "'", "''") & "'!BuildMemberTemplate"

    built = ""
    On Error Resume Next
    built = CStr(tw.Worksheets("Home").Range("B19").Value)
    On Error GoTo Failed
    If StrComp(Trim$(built), wantVer, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 906, , "the rebuild didn't finish (the template's Home sheet says '" & built & "', expected '" & wantVer & "')"
    End If
    Application.DisplayAlerts = False
    tw.Save
    savedIt = True
    tw.Close SaveChanges:=False
    Set tw = Nothing
    Application.DisplayAlerts = daWas
    Application.EnableEvents = evWas
    roster.Activate
    MsgBox "The members' template is on " & wantVer & " and saved." & vbNewLine & vbNewLine & _
           "Now click Member View in the roster to publish the members' file.", vbInformation, TTL
    Exit Sub
Failed:
    why = Err.Description
    On Error Resume Next
    Application.AutomationSecurity = secWas
    Application.DisplayAlerts = False
    If Not tw Is Nothing Then tw.Close SaveChanges:=False
    Application.DisplayAlerts = daWas
    Application.EnableEvents = evWas
    roster.Activate
    If savedIt Then
        MsgBox "The members' template was updated and saved, but closing it gave: " & why & vbNewLine & _
               "Close it by hand if it is still open, then click Member View.", vbExclamation, TTL
    Else
        MsgBox "The members' template was NOT updated: " & why & vbNewLine & vbNewLine & _
               IIf(opened, "It was closed without saving, so", "It couldn't be opened, so") & " the file on disk is as it was." & vbNewLine & _
               "Do it by hand from the install note (the members' part), or send me this message.", vbCritical, TTL
    End If
End Sub


' ====================================================================
'  PIECES (no calls into any other module)
' ====================================================================
' The admin test, by name, BEFORE anything is changed. If it can't be run
' (the project doesn't compile), the admin is asked instead.
Private Function UpdIsAdmin(ByVal wb As Workbook) As Boolean
    Dim v As Variant, en As Long
    On Error Resume Next
    v = Application.Run("'" & Replace(wb.Name, "'", "''") & "'!IsAdmin")
    en = Err.Number
    On Error GoTo 0
    If en <> 0 Then
        UpdIsAdmin = (MsgBox("The admin list couldn't be checked (the roster's code isn't running)." & vbNewLine & vbNewLine & _
                             "Only the roster's administrator should update its code. Continue?", _
                             vbYesNo + vbExclamation + vbDefaultButton2, UPD_TITLE) = vbYes)
        Exit Function
    End If
    If VarType(v) = vbBoolean Then UpdIsAdmin = CBool(v)
    If Not UpdIsAdmin Then
        MsgBox "Your Windows login (" & Environ$("USERNAME") & ") is not on the admin list of " & wb.Name & ".", vbCritical, UPD_TITLE
    End If
End Function

' The roster to work on: the one open workbook (other than this one) whose
' project holds modOvertimeEntry. Nothing (after a message) = can't tell.
Private Function UpdTarget() As Workbook
    Dim why As String, w As Workbook, hit As Workbook, n As Long, activeIsOne As Boolean
    Dim own As Object: Set own = UpdProject(ThisWorkbook, why)
    If own Is Nothing Then
        MsgBox why, vbExclamation, UPD_TITLE
        Exit Function
    End If
    If Not UpdFindComp(own, "modOvertimeEntry") Is Nothing Then
        MsgBox "This updater is sitting INSIDE a roster (" & ThisWorkbook.Name & "), and it must never run from there." & vbNewLine & vbNewLine & _
               "Remove modUpdater from the roster (Alt+F11, right-click it, Remove, No)." & vbNewLine & _
               "Then make the updater workbook: a new blank workbook, File > Import File > modUpdater.bas, " & _
               "save it as 'Roster Updater.xlsm', and run this from there with the roster open.", vbCritical, UPD_TITLE
        Exit Function
    End If
    For Each w In Application.Workbooks
        If Not w Is ThisWorkbook Then
            If UpdIsRoster(w) Then
                n = n + 1
                Set hit = w
                If w Is ActiveWorkbook Then activeIsOne = True
            End If
        End If
    Next w
    If n = 0 Then
        MsgBox "No roster is open. Open the roster workbook, then run this again.", vbExclamation, UPD_TITLE
        Exit Function
    End If
    If n > 1 Then
        If Not activeIsOne Then
            MsgBox n & " rosters are open. Click the one to update so it is in front (or close the others), then run this again.", vbExclamation, UPD_TITLE
            Exit Function
        End If
        Set hit = ActiveWorkbook
    End If
    Set UpdTarget = hit
End Function

Private Function UpdIsRoster(ByVal w As Workbook) As Boolean
    Dim why As String, proj As Object
    On Error Resume Next
    Set proj = UpdProject(w, why)
    If Not proj Is Nothing Then UpdIsRoster = Not UpdFindComp(proj, "modOvertimeEntry") Is Nothing
End Function

' The workbook's VBA project, or Nothing with the reason in why
Private Function UpdProject(ByVal wb As Workbook, ByRef why As String) As Object
    Dim proj As Object, cnt As Long, en As Long, prot As Long
    On Error Resume Next
    Set proj = wb.VBProject
    If Not proj Is Nothing Then prot = proj.Protection
    Err.Clear
    If Not proj Is Nothing And prot <> 1 Then cnt = proj.VBComponents.Count
    en = Err.Number
    On Error GoTo 0
    If Not proj Is Nothing And prot = 1 Then
        why = "The VBA project of " & wb.Name & " is locked with a password." & vbNewLine & _
              "Unlock it first (Alt+F11, double-click the project, enter the password), then run this again."
        Exit Function
    End If
    If proj Is Nothing Or en <> 0 Or cnt = 0 Then
        why = "Excel isn't allowing access to the VBA project of " & wb.Name & "." & vbNewLine & vbNewLine & _
              "Tick: File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
              """Trust access to the VBA project object model"". Then run this again."
        Exit Function
    End If
    Set UpdProject = proj
End Function

Private Function UpdFindComp(ByVal proj As Object, ByVal compName As String) As Object
    Dim comp As Object
    For Each comp In proj.VBComponents
        If StrComp(comp.Name, compName, vbTextCompare) = 0 Then
            Set UpdFindComp = comp
            Exit Function
        End If
    Next comp
End Function

' Replaces ALL the code of one component with txt. The component stays, so
' Excel can never bring it back under a second name. If the new text doesn't
' go in, the old text is put back before returning False.
Private Function UpdSwapCode(ByVal comp As Object, ByVal txt As String, ByRef why As String) As Boolean
    Dim cm As Object, nOld As Long, want As Long, oldTxt As String, emptied As Boolean
    On Error GoTo Failed
    Set cm = comp.CodeModule
    nOld = cm.CountOfLines
    If nOld > 0 Then oldTxt = cm.Lines(1, nOld)
    If nOld > 0 Then cm.DeleteLines 1, nOld
    emptied = True
    If txt <> "" Then cm.AddFromString txt
    want = UpdLineCount(txt)
    If cm.CountOfLines < want - 3 Then Err.Raise vbObjectError + 910, , "only " & cm.CountOfLines & " of " & want & " lines went in"
    UpdSwapCode = True
    Exit Function
Failed:
    why = Err.Description
    If emptied Then
        On Error Resume Next
        If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
        If oldTxt <> "" Then cm.AddFromString oldTxt
    End If
End Function

' A module the workbook doesn't have yet: a normal import
Private Function UpdImport(ByVal proj As Object, ByVal fullPath As String, ByVal compName As String, ByRef why As String) As Boolean
    Dim comp As Object
    On Error GoTo Failed
    Set comp = proj.VBComponents.Import(fullPath)
    If StrComp(comp.Name, compName, vbTextCompare) <> 0 Then
        why = "it came in as '" & comp.Name & "' instead of '" & compName & "'"
        On Error Resume Next                        ' the stray copy is emptied and removed; the reason above is kept
        If comp.CodeModule.CountOfLines > 0 Then comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
        proj.VBComponents.Remove comp
        Exit Function
    End If
    UpdImport = True
    Exit Function
Failed:
    why = Err.Description
End Function

Private Function UpdNameSet(ByVal wb As Workbook, ByVal nameTxt As String) As Boolean
    Dim nmObj As Name
    On Error Resume Next
    Set nmObj = wb.Names(nameTxt)
    On Error GoTo 0
    UpdNameSet = Not nmObj Is Nothing
End Function

' UPD_NOSAVE is the hidden name the roster's own Workbook_BeforeSave and Workbook_BeforeClose read
Private Sub UpdSetName(ByVal wb As Workbook, ByVal nameTxt As String)
    On Error Resume Next
    If Not UpdNameSet(wb, nameTxt) Then wb.Names.Add Name:=nameTxt, RefersTo:="=TRUE", Visible:=False
End Sub

Private Sub UpdDropName(ByVal wb As Workbook, ByVal nameTxt As String)
    On Error Resume Next
    wb.Names(nameTxt).Delete
End Sub

' Asks for the release folder. Accepts the folder itself or the folder above
' it (a zip unpacks into one folder of its own). mustHave = the subfolder
' that proves it ("modules" or "member"). "" = cancelled or not a release.
Private Function UpdPickFolder(ByVal mustHave As String) As String
    Dim picked As String
    With Application.FileDialog(4)                  ' msoFileDialogFolderPicker
        .Title = "Pick the unzipped release folder (the one holding modules, paste-in and member)"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        picked = UpdSlash(CStr(.SelectedItems(1)))
    End With
    If UpdDirExists(picked & mustHave) Then UpdPickFolder = picked: Exit Function
    If StrComp(Right$(picked, Len(mustHave) + 2), "\" & mustHave & "\", vbTextCompare) = 0 Then      ' they picked modules\ itself
        UpdPickFolder = Left$(picked, Len(picked) - Len(mustHave) - 1)
        Exit Function
    End If
    Dim subs() As String, k As Long, i As Long, hit As String, nHit As Long
    k = UpdListDirs(picked, subs)
    For i = 1 To k
        If UpdDirExists(picked & subs(i) & "\" & mustHave) Then nHit = nHit + 1: hit = picked & subs(i) & "\"
    Next i
    If nHit = 1 Then UpdPickFolder = hit: Exit Function
    MsgBox "That folder doesn't hold a release: there is no '" & mustHave & "' folder in it." & vbNewLine & picked & vbNewLine & vbNewLine & _
           "Unzip the release first, then pick the folder that holds modules, paste-in and member.", vbExclamation, UPD_TITLE
End Function

Private Function UpdListFiles(ByVal dirPath As String, ByVal pattern As String, ByRef names() As String) As Long
    Dim f As String, n As Long, ext As String
    ext = Mid$(pattern, InStrRev(pattern, "."))
    On Error GoTo Done
    f = Dir$(dirPath & pattern)
    Do While f <> ""
        If StrComp(Right$(f, Len(ext)), ext, vbTextCompare) = 0 Then      ' Dir also matches longer extensions
            n = n + 1
            ReDim Preserve names(1 To n)
            names(n) = f
        End If
        f = Dir$()
    Loop
Done:
    UpdListFiles = n
End Function

Private Function UpdListDirs(ByVal dirPath As String, ByRef names() As String) As Long
    Dim f As String, n As Long, all() As String, k As Long, i As Long
    On Error GoTo Done
    f = Dir$(dirPath & "*", vbDirectory)
    Do While f <> ""
        If f <> "." And f <> ".." Then
            k = k + 1
            ReDim Preserve all(1 To k)
            all(k) = f
        End If
        f = Dir$()
    Loop
    For i = 1 To k
        If UpdDirExists(dirPath & all(i)) Then
            n = n + 1
            ReDim Preserve names(1 To n)
            names(n) = all(i)
        End If
    Next i
Done:
    UpdListDirs = n
End Function

Private Function UpdDirExists(ByVal p As String) As Boolean
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    On Error Resume Next
    UpdDirExists = ((GetAttr(p) And vbDirectory) = vbDirectory)
End Function

Private Function UpdFileExists(ByVal p As String) As Boolean
    On Error Resume Next
    UpdFileExists = ((GetAttr(p) And vbDirectory) = 0)
End Function

Private Function UpdSlash(ByVal p As String) As String
    If p <> "" And Right$(p, 1) <> "\" Then p = p & "\"
    UpdSlash = p
End Function

' The whole file as text (the release files are plain ASCII)
Private Function UpdReadFile(ByVal fullPath As String, ByRef txt As String) As Boolean
    Dim ff As Integer, nBytes As Long
    txt = ""
    On Error GoTo Failed
    ff = FreeFile
    Open fullPath For Binary Access Read Shared As #ff
    nBytes = LOF(ff)
    If nBytes > 0 Then
        txt = Space$(nBytes)
        Get #ff, 1, txt
    End If
    Close #ff
    UpdReadFile = (nBytes > 0)
    Exit Function
Failed:
    On Error Resume Next
    Close #ff
End Function

' A .bas file starts with "Attribute VB_Name = ..." lines. They belong to the
' file format, not to the code window: pasted in, they are a syntax error.
Private Function UpdStripAttributes(ByVal txt As String) As String
    Dim p As Long
    Do While Left$(txt, 10) = "Attribute "
        p = InStr(txt, vbLf)
        If p = 0 Then txt = "": Exit Do
        txt = Mid$(txt, p + 1)
    Loop
    UpdStripAttributes = txt
End Function

' The text between two markers ("" = not found)
Private Function UpdBetween(ByVal txt As String, ByVal before As String, ByVal after As String) As String
    Dim p As Long, q As Long
    p = InStr(1, txt, before, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(before)
    q = InStr(p, txt, after, vbTextCompare)
    If q = 0 Then Exit Function
    UpdBetween = Mid$(txt, p, q - p)
End Function

Private Function UpdLineCount(ByVal txt As String) As Long
    Dim p As Long, n As Long
    If txt = "" Then Exit Function
    p = InStr(txt, vbLf)
    Do While p > 0
        n = n + 1
        p = InStr(p + 1, txt, vbLf)
    Loop
    If Right$(txt, 1) <> vbLf Then n = n + 1
    UpdLineCount = n
End Function

' Writes the report next to the workbook; returns its path
Private Function UpdWriteReport(ByVal wb As Workbook, ByVal rep As String, ByVal what As String) As String
    Dim ff As Integer, fullPath As String
    fullPath = UpdSlash(wb.Path) & what & " " & Format$(Now, "yyyy-mm-dd hhmmss") & ".txt"
    On Error GoTo Failed
    ff = FreeFile
    Open fullPath For Output As #ff
    Print #ff, rep
    Close #ff
    UpdWriteReport = fullPath
    Exit Function
Failed:
    On Error Resume Next
    Close #ff
    UpdWriteReport = "(the report couldn't be written)"
End Function
