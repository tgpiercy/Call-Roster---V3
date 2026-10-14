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
'  v3 (Oct 13, 2026): v2 said "No roster is open" with the roster open
'  (UpdateMemberTemplate, first real run). The roster is now known by its
'  code OR by its OT Log table, and when nothing is recognised the box
'  lists what Excel can see and offers the workbook in front.
'
'  v4 (Oct 14, 2026, briefs/v24.4.md): THE INSTALL PANEL. Alt+F8 >
'  InstallPanel draws an Install sheet in THIS workbook from the release's
'  install-steps.txt (zip root, from v24.4): groups of steps, one button
'  enabled at a time, each click runs one step (backup, check, swap,
'  compile, finish + reopen, a roster macro, the members' template, a
'  manual row, a Pass / Fail test with a pre-filled pop-up) and writes a
'  row to the roster's hidden "Install Log" sheet (tblInstallLog). The log
'  is the state: closing Excel part-way and opening again resumes at the
'  next step. The engine below is unchanged; the panel calls its pieces.
'  With trust access off, swap and compile show the by-hand list instead.
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

Public Const UPD_VERSION As String = "Updater v4 (Oct 14)"

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

' v4: the Install panel (see THE INSTALL PANEL at the end of the module)
Private Const INST_SHEET As String = "Install"
Private Const INST_LOG_SHEET As String = "Install Log"        ' in the ROSTER (modSystem.IL_SHEET)
Private Const INST_LOG_TABLE As String = "tblInstallLog"
Private Const INST_TITLE As String = "Install panel"
Private Const INST_ROW0 As Long = 10                          ' first step row on the sheet
Private Const INST_COL_KIND As Long = 11                      ' K: kind (hidden), L caption, M argument, N flag, O step id, P group
Private Const INST_MAX As Long = 80

Private Type InstStep
    Kind As String
    Caption As String
    Arg As String
    Flag As String
    Id As String            ' "S01" ... in list order; the log key
    Group As String         ' "" or the group title the step opens
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

    Dim why As String
    If Not UpdFinish(wb, why) Then
        MsgBox why, vbCritical, UPD_TITLE
        Exit Sub
    End If
    MsgBox "Saved with the new code." & vbNewLine & vbNewLine & _
           "NEXT:" & vbNewLine & _
           "1. Close the workbook and open it again." & vbNewLine & _
           "2. Alt+F8 > SystemCheck." & vbNewLine & _
           "3. Run the macros the install note lists under 'run once'." & vbNewLine & _
           "4. Alt+F8 > UpdateMemberTemplate (if the release has a members' part), then Member View." & vbNewLine & vbNewLine & _
           "Keep the Roster Updater workbook for the next release (v4: its Install sheet is saved with it).", vbInformation, UPD_TITLE
End Sub

' v4: the finish itself (UpdateFinished and the panel's finish step): the new code must run
' at all before anything is saved (ApplyView also puts the admin view's state back), then the
' save block and the pending mark come off and the workbook is saved. False = why, and the
' block stays on.
Private Function UpdFinish(ByVal wb As Workbook, ByRef why As String) As Boolean
    Dim en As Long, em As String
    On Error Resume Next
    Application.Run "'" & Replace(wb.Name, "'", "''") & "'!ApplyView", True
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    If en <> 0 Then
        why = "The new code didn't run (" & em & "), so nothing was saved and saving is still blocked." & vbNewLine & vbNewLine & _
              "Close the workbook (Don't Save if Excel asks) and send me that message."
        Exit Function
    End If
    UpdDropName wb, UPD_NOSAVE
    UpdDropName wb, UPD_PENDING
    On Error Resume Next
    wb.Save
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    If en <> 0 Or Not wb.Saved Then
        why = "The save didn't go through" & IIf(em <> "", " (" & em & ")", "") & ". Press Ctrl+S and check for a message."
        Exit Function
    End If
    UpdFinish = True
End Function

' Compile showed an error, or you changed your mind: closes this workbook
' WITHOUT saving, whatever state its code is in. The file on disk is untouched.
Public Sub UpdateDrop()
    Dim wb As Workbook: Set wb = UpdTarget()
    If wb Is Nothing Then Exit Sub
    If Not UpdNameSet(wb, UPD_PENDING) And Not UpdNameSet(wb, UPD_NOSAVE) Then
        If MsgBox("No update is waiting in " & wb.Name & "." & vbNewLine & vbNewLine & _
                  "Close it WITHOUT saving anyway? Anything not saved in it is lost.", _
                  vbYesNo + vbExclamation + vbDefaultButton2, UPD_TITLE) <> vbYes Then Exit Sub
        wb.Close SaveChanges:=False
        Exit Sub
    End If
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
    On Error GoTo Oops                              ' v25 (C4): a catch-all; the swap itself has its own handling
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

    ' ---- backup, save block, swap (v4: the pieces the Install panel also uses) ----
    Dim bk As String, repPath As String
    If Not UpdBackup(wb, "update", bk, why) Then
        MsgBox "The backup copy couldn't be saved, so nothing was changed:" & vbNewLine & why, vbCritical, UPD_TITLE
        Exit Sub
    End If
    rep = rep & "Backup: " & bk & vbNewLine & vbNewLine
    If Not UpdSwapAll(wb, items, n, rep, repPath, why) Then
        If repPath = "" Then MsgBox why, vbCritical, UPD_TITLE: Exit Sub          ' the save block couldn't be set: nothing was touched
        MsgBox "The update did NOT go in:" & vbNewLine & why & vbNewLine & _
               "Saving is blocked. Close this workbook now; if Excel asks, click DON'T SAVE." & vbNewLine & _
               "(If you try to save, the roster says 'conversion failed verification': that is this block.)" & vbNewLine & vbNewLine & _
               "The file on disk is as it was. Backup copy: " & bk & vbNewLine & vbNewLine & _
               "Then install by hand from INSTALL.md, and send me the report:" & vbNewLine & repPath, vbCritical, UPD_TITLE
        Exit Sub
    End If
    MsgBox "The code is now " & IIf(relNew = "", "the folder's release", relNew) & ": " & nRep & " replaced, " & nAdd & " added." & vbNewLine & _
           "It is NOT saved yet. Do nothing else in the workbook until you finish:" & vbNewLine & vbNewLine & _
           "1. Alt+F11, click the ROSTER's project in the list on the left (" & wb.Name & ")," & vbNewLine & _
           "    then Debug > Compile VBAProject." & vbNewLine & _
           "2. NO message: Alt+F8 > UpdateFinished. It saves and tells you what is next." & vbNewLine & _
           "    A message: note the line it highlights, then Alt+F8 > UpdateDrop." & vbNewLine & _
           "    It closes without saving; your file stays exactly as it was." & vbNewLine & _
           "    Do NOT press Ctrl+S (with a compile error the block may not hold). Closing by hand: DON'T SAVE." & vbNewLine & vbNewLine & _
           "Backup: " & bk & vbNewLine & "Report: " & repPath, vbInformation, UPD_TITLE
    Exit Sub
Oops:
    MsgBox "The update stopped part-way: " & Err.Description & vbNewLine & vbNewLine & _
           "If the code was already swapped, saving is blocked: close the roster WITHOUT saving (Don't Save) and run it again.", vbCritical, UPD_TITLE
End Sub

' v4: a copy of the roster beside it, "Backup before <label> <stamp> - <name>". False = why.
Private Function UpdBackup(ByVal wb As Workbook, ByVal label As String, ByRef bk As String, ByRef why As String) As Boolean
    bk = UpdSlash(wb.Path) & "Backup before " & label & " " & Format$(Now, "yyyy-mm-dd hhmmss") & " - " & wb.Name
    why = ""
    On Error Resume Next
    ' v25 (I2): the roster saves its own copy in the restricted view when it can (its code runs
    ' before a swap); an older roster, or one whose code can't run, is copied as it is
    Application.Run "'" & Replace(wb.Name, "'", "''") & "'!SaveRestrictedCopy", bk
    If Err.Number <> 0 Then
        Err.Clear
        wb.SaveCopyAs bk
        If Err.Number <> 0 Then why = Err.Description
    End If
    On Error GoTo 0
    UpdBackup = (why = "")
End Function

' v4: the save block goes on, the code is swapped (rolled back on a failure), the report is
' written beside the roster. items come from UpdGatherRoster. True = swapped and waiting for
' Compile + finish; False = nothing of the release is in (why says what failed; the block stays).
Private Function UpdSwapAll(ByVal wb As Workbook, ByRef items() As UpdItem, ByVal n As Long, _
                            ByVal rep As String, ByRef repPath As String, ByRef why As String) As Boolean
    Dim proj As Object: Set proj = UpdProject(wb, why)
    If proj Is Nothing Then Exit Function
    ' ---- block saving BEFORE the first change; it stays blocked until the finish ----
    UpdSetName wb, UPD_NOSAVE
    UpdSetName wb, UPD_PENDING
    If Not UpdNameSet(wb, UPD_NOSAVE) Or Not UpdNameSet(wb, UPD_PENDING) Then
        UpdDropName wb, UPD_NOSAVE
        UpdDropName wb, UPD_PENDING
        why = "The save block couldn't be set, so nothing was changed."
        Exit Function
    End If

    ' ---- swap the code (VBA editor calls raise no workbook events, so events stay as they are) ----
    Dim scrWas As Boolean: scrWas = Application.ScreenUpdating
    Dim nFail As Long, backTxt As String, i As Long
    Application.ScreenUpdating = False
    nFail = UpdApply(proj, items, n)
    If nFail > 0 Then backTxt = UpdRollBack(proj, items, n)
    Application.ScreenUpdating = scrWas

    For i = 1 To n
        rep = rep & IIf(items(i).Result = "", "not reached", items(i).Result) & "  " & items(i).CompName & "  <-  " & items(i).FileName & vbNewLine
    Next i
    rep = rep & UpdLeftAlone(proj, items, n)

    If nFail > 0 Then
        UpdDropName wb, UPD_PENDING                 ' nothing to finish; the save block stays
        rep = rep & vbNewLine & "FAILED. " & IIf(backTxt = "", "The old code was put back.", "Putting the old code back ALSO failed: " & backTxt) & vbNewLine
        repPath = UpdWriteReport(wb, rep, "Update report")
        Dim failTxt As String
        For i = 1 To n
            If Left$(items(i).Result, 6) = "FAILED" Then failTxt = failTxt & "  " & items(i).CompName & ": " & Mid$(items(i).Result, 9) & vbNewLine
        Next i
        why = failTxt & IIf(backTxt = "", "The old code was put back.", "Putting the old code back also failed (" & backTxt & ").")
        Exit Function
    End If
    rep = rep & vbNewLine & "Code swapped. Waiting for Compile and the finish." & vbNewLine
    repPath = UpdWriteReport(wb, rep, "Update report")
    UpdSwapAll = True
End Function

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
' v4: a Function - True when the template was rebuilt and saved (the Install panel's template step)
Private Function UpdTemplate(ByVal roster As Workbook, ByVal folder As String) As Boolean
    Const TTL As String = "Update the members' template"
    Dim why As String, modTxt As String, wbTxt As String, wantVer As String
    Dim tPath As String: tPath = UpdSlash(roster.Path) & UPD_TEMPLATE

    If Not UpdFileExists(tPath) Then
        MsgBox "The members' template wasn't found next to the roster:" & vbNewLine & tPath, vbExclamation, TTL
        Exit Function
    End If
    If Not UpdReadFile(folder & "member\modMember.bas", modTxt) Or Not UpdReadFile(folder & "member\ThisWorkbook.txt", wbTxt) Then
        MsgBox "member\modMember.bas and member\ThisWorkbook.txt must both be in:" & vbNewLine & folder, vbExclamation, TTL
        Exit Function
    End If
    If Asc(Left$(modTxt, 1)) > 126 Or InStr(modTxt, vbCrLf) = 0 Or Asc(Left$(wbTxt, 1)) > 126 Or InStr(wbTxt, vbCrLf) = 0 Then
        MsgBox "The files in " & folder & "member\ aren't plain text with Windows line ends. Unzip the release again. Nothing was changed.", vbExclamation, TTL
        Exit Function
    End If
    wantVer = UpdBetween(modTxt, "MEM_VERSION As String = """, """")
    modTxt = UpdStripAttributes(modTxt)
    wbTxt = UpdStripAttributes(wbTxt)
    If InStr(1, modTxt, "Option Explicit", vbTextCompare) = 0 Or InStr(1, wbTxt, "Option Explicit", vbTextCompare) = 0 Or wantVer = "" Then
        MsgBox "The files in " & folder & "member\ don't look like the members' code. Nothing was changed.", vbExclamation, TTL
        Exit Function
    End If

    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.Name, UPD_TEMPLATE, vbTextCompare) = 0 Then
            MsgBox "The template is open. Close it (without saving), then run UpdateMemberTemplate again.", vbExclamation, TTL
            Exit Function
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
    built = CStr(Application.Run("'" & Replace(tw.Name, "'", "''") & "'!MemberLayoutVersion"))     ' v25 (I7): the template says which layout it built, as modPublish asks it
    If built = "" Then built = CStr(tw.Worksheets("Home").Range("B19").Value)                       ' (older templates: the cell the function reads)
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
    UpdTemplate = True
    Exit Function
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
        UpdTemplate = True
        MsgBox "The members' template was updated and saved, but closing it gave: " & why & vbNewLine & _
               "Close it by hand if it is still open, then click Member View.", vbExclamation, TTL
    Else
        MsgBox "The members' template was NOT updated: " & why & vbNewLine & vbNewLine & _
               IIf(opened, "It was closed without saving, so", "It couldn't be opened, so") & " the file on disk is as it was." & vbNewLine & _
               "Do it by hand from the install note (the members' part), or send me this message.", vbCritical, TTL
    End If
End Function


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

' The roster to work on: the one open workbook (other than this one) that
' is a roster. Nothing (after a message) = can't tell.
' v3: a roster is known by its code (modOvertimeEntry) OR by its data (the
' OT Log table), so a project that can't be read no longer hides it. When
' nothing is recognised the message lists what Excel can see and why each
' was passed over, and the admin can pick the workbook in front.
Private Function UpdTarget() As Workbook
    Dim why As String, w As Workbook, hit As Workbook, n As Long, activeIsOne As Boolean
    Dim seen As String, nOther As Long
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
        If Not UpdSame(w, ThisWorkbook) Then
            nOther = nOther + 1
            why = ""
            If UpdIsRoster(w, why) Then
                n = n + 1
                Set hit = w
                If UpdSame(w, ActiveWorkbook) Then activeIsOne = True
            ElseIf nOther <= 8 Then
                seen = seen & "  " & w.Name & ": " & why & vbNewLine
            End If
        End If
    Next w
    If n = 0 Then
        Set hit = UpdAskTarget(nOther, seen)
        If hit Is Nothing Then Exit Function
    ElseIf n > 1 Then
        If Not activeIsOne Then
            MsgBox n & " rosters are open. Click the one to update so it is in front (or close the others), then run this again.", vbExclamation, UPD_TITLE
            Exit Function
        End If
        Set hit = ActiveWorkbook
    End If
    ' an updater left INSIDE that roster (the Oct 12 attempt) puts a second UpdateFromFolder in the Alt+F8 list
    Dim tproj As Object: Set tproj = UpdProject(hit, why)
    If Not tproj Is Nothing Then
        If Not UpdFindComp(tproj, "modUpdater") Is Nothing Then
            MsgBox hit.Name & " still has a modUpdater module inside it from an earlier attempt." & vbNewLine & vbNewLine & _
                   "Remove it first: Alt+F11, find modUpdater under the ROSTER's project, right-click > Remove > No. " & _
                   "Save the roster, then run this again from Roster Updater.xlsm.", vbExclamation, UPD_TITLE
            Exit Function
        End If
    End If
    Set UpdTarget = hit
End Function

' Two workbook variables for the same open file (compared by path, which
' does not depend on how Excel handed the object over)
Private Function UpdSame(ByVal a As Workbook, ByVal b As Workbook) As Boolean
    If a Is Nothing Or b Is Nothing Then Exit Function
    Dim fa As String, fb As String
    On Error Resume Next
    fa = a.FullName
    fb = b.FullName
    On Error GoTo 0
    If fa = "" Or fb = "" Then Exit Function
    UpdSame = (StrComp(fa, fb, vbTextCompare) = 0)
End Function

' True when w is a roster. why = the reason it was passed over.
Private Function UpdIsRoster(ByVal w As Workbook, ByRef why As String) As Boolean
    Dim proj As Object, comp As Object, pwhy As String, ws As Worksheet, lo As Object
    why = ""
    On Error Resume Next
    Set proj = UpdProject(w, pwhy)
    If proj Is Nothing Then
        why = "its VBA project can't be read"
    Else
        Set comp = Nothing
        Set comp = UpdFindComp(proj, "modOvertimeEntry")
        If Not comp Is Nothing Then
            UpdIsRoster = True
            Exit Function
        End If
        why = "its code has no modOvertimeEntry module"
    End If
    ' the data says so even when the code can't: the OT Log table
    For Each ws In w.Worksheets
        Set lo = Nothing
        Set lo = ws.ListObjects("tblOTLog")
        If Not lo Is Nothing Then
            why = ""
            UpdIsRoster = True
            Exit Function
        End If
    Next ws
    why = why & ", and it holds no OT Log table"
End Function

' No open workbook passed the roster test: say what Excel can see, and let
' the admin use the workbook in front. Nothing = stop.
Private Function UpdAskTarget(ByVal nOther As Long, ByVal seen As String) As Workbook
    Dim aw As Workbook
    On Error Resume Next
    Set aw = ActiveWorkbook
    On Error GoTo 0
    If nOther = 0 Then
        MsgBox "This Excel has no other workbook open, so the roster can't be found." & vbNewLine & vbNewLine & _
               "If the roster IS open on your screen, it is in a second Excel that this one can't see. " & _
               "Close Roster Updater.xlsm, click the roster, open Roster Updater.xlsm with File > Open from there, " & _
               "and run this again.", vbExclamation, UPD_TITLE
        Exit Function
    End If
    If aw Is Nothing Then
        MsgBox "No open workbook was recognised as a roster:" & vbNewLine & seen & vbNewLine & _
               "Click the roster so it is in front, then run this again. Send me this box.", vbExclamation, UPD_TITLE
        Exit Function
    End If
    If UpdSame(aw, ThisWorkbook) Then
        MsgBox "No open workbook was recognised as a roster:" & vbNewLine & seen & vbNewLine & _
               "Click the roster so it is in front, then press Alt+F8 and run this again. Send me this box.", vbExclamation, UPD_TITLE
        Exit Function
    End If
    If MsgBox("No open workbook was recognised as a roster:" & vbNewLine & seen & vbNewLine & _
              "The workbook in front is " & aw.Name & "." & vbNewLine & _
              "Is that the roster?   Yes = use it.   No = stop." & vbNewLine & vbNewLine & _
              "(Send me a photo of this box either way: it says why the roster wasn't recognised.)", _
              vbYesNo + vbQuestion + vbDefaultButton2, UPD_TITLE) = vbYes Then Set UpdAskTarget = aw
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

' (v3: one component whose name can't be read no longer stops the search)
Private Function UpdFindComp(ByVal proj As Object, ByVal compName As String) As Object
    Dim comp As Object, nm As String
    On Error Resume Next
    For Each comp In proj.VBComponents
        nm = ""
        nm = comp.Name
        If nm <> "" Then
            If StrComp(nm, compName, vbTextCompare) = 0 Then
                Set UpdFindComp = comp
                Exit Function
            End If
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


' ====================================================================
'  THE INSTALL PANEL (v4)
'
'  State lives in two places only: the Install sheet's cells in THIS
'  workbook (the folder picked, the parsed step list) and the roster's
'  Install Log table (what has been done). Every click re-reads both, so
'  no module-level variable is needed and a reopened roster is found again.
'
'  install-steps.txt (zip root, generated by tests/build_zip.py):
'    # comment            VERSION vNN   EXPECT vNN   MEMBER yes|no   FORMS ...
'    [n. Group title]
'    kind | caption | argument | flag or expected result
'  kinds: backup, check, swap, compile, finish (flag reopen), macro (argument
'  = macro name + arguments; flag confirm), template, manual, test (argument
'  = opener macro + arguments, flag = the expected result). A date argument
'  +N means today plus N days; yyyy-mm-dd is taken as written.
' ====================================================================
' Alt+F8 > InstallPanel: draws the panel (and picks the folder when none is set)
Public Sub InstallPanel()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = InstSheet(True)
    ws.Activate
    If Trim$(CStr(ws.Range("C3").Value)) = "" Then
        InstallPickFolder
    Else
        InstRefresh
    End If
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

' The "Pick the update folder" button: reads the release's install-steps.txt
Public Sub InstallPickFolder()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = InstSheet(True)
    Dim folder As String: folder = UpdPickFolder("modules")
    If folder = "" Then Exit Sub
    Dim why As String
    If Not UpdFileExists(folder & "install-steps.txt") Then
        MsgBox "That release has no install-steps.txt (releases before v24.4 don't)." & vbNewLine & folder & vbNewLine & vbNewLine & _
               "Install it with Alt+F8 > UpdateFromFolder and its INSTALL.md instead.", vbExclamation, INST_TITLE
        Exit Sub
    End If
    If Not InstParse(folder & "install-steps.txt", ws, why) Then
        MsgBox "The step list couldn't be read: " & why, vbExclamation, INST_TITLE
        Exit Sub
    End If
    ws.Range("C3").Value = folder
    ws.Range("C6").Value = ""
    InstRefresh
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

Public Sub InstallRefresh()
    On Error GoTo Oops
    InstRefresh
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

' A step button: runs that step if it is the next one
Public Sub InstallStepClick()
    On Error GoTo Oops
    Dim nm As String: nm = CStr(Application.Caller)
    If Left$(nm, 10) <> "inst_step_" Then Exit Sub
    InstRunIndex CLng(Mid$(nm, 11))
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

' The "Skip this step" button: the next step is marked SKIPPED with a note (required)
Public Sub InstallSkip()
    On Error GoTo Oops
    Dim ws As Worksheet: Set ws = InstSheet(False)
    If ws Is Nothing Then Exit Sub
    Dim steps() As InstStep, n As Long
    n = InstSteps(ws, steps)
    If n = 0 Then MsgBox "Pick the update folder first.", vbInformation, INST_TITLE: Exit Sub
    Dim wb As Workbook: Set wb = InstTarget()
    If wb Is Nothing Then Exit Sub
    Dim st As Object: Set st = InstReadLog(wb, InstVersion(ws))
    Dim i As Long: i = InstNext(steps, n, st)
    If i = 0 Then MsgBox "Every step is done.", vbInformation, INST_TITLE: Exit Sub
    Dim note As String
    note = Trim$(VBA.InputBox("Skip step " & steps(i).Id & " (" & steps(i).Caption & ")?" & vbNewLine & vbNewLine & _
                              "Say why (it goes on the record; a step done by hand counts as skipped):", INST_TITLE))
    If note = "" Then Exit Sub
    InstLog wb, InstVersion(ws), steps(i), "SKIPPED", note
    InstRefresh
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

' The "Show the log" button: the roster's Install Log sheet, in front
Public Sub InstallShowLog()
    On Error GoTo Oops
    Dim wb As Workbook: Set wb = InstTarget()
    If wb Is Nothing Then Exit Sub
    Dim why As String, en As Long
    ' v25: the roster's structure is protected, so the roster shows the sheet itself (ShowInstallLog)
    On Error Resume Next
    Application.Run "'" & Replace(wb.Name, "'", "''") & "'!ShowInstallLog"
    en = Err.Number
    On Error GoTo Oops
    If en = 0 Then Exit Sub
    Dim lo As Object: Set lo = InstEnsureLog(wb, why)
    If lo Is Nothing Then MsgBox why, vbExclamation, INST_TITLE: Exit Sub
    wb.Activate
    lo.Parent.Visible = -1                          ' xlSheetVisible (the roster's ApplyView hides it again for non-admins)
    lo.Parent.Activate
    Exit Sub
Oops:
    MsgBox "This didn't finish: " & Err.Description, vbExclamation, INST_TITLE
End Sub

' ---------------- the sheet ----------------
Private Function InstSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(INST_SHEET)
    On Error GoTo 0
    If ws Is Nothing And create Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = INST_SHEET
    End If
    If Not ws Is Nothing And create Then InstDrawFrame ws      ' every time: the buttons carry this workbook's current name
    Set InstSheet = ws
End Function

Private Sub InstDrawFrame(ByVal ws As Worksheet)
    With ws
        .Range("B1:B6").ClearContents                   ' the labels only: C3:C6 and B7 (state) and K:P (the list) stay
        .Range("B8:G8").ClearContents
        .Range("B1").Value = "CALL OUT ROSTER - INSTALL PANEL"
        .Range("B1").Font.Bold = True: .Range("B1").Font.Size = 16
        .Range("B2").Value = UPD_VERSION & ". One button at a time, in order. Each click runs one step and writes it to the roster's Install Log."
        .Range("B3").Value = "Update folder:": .Range("B4").Value = "Roster:": .Range("B5").Value = "Release:": .Range("B6").Value = "Status:"
        .Range("B3:B6").Font.Bold = True
        .Range("B8").Value = "#": .Range("C8").Value = "Step": .Range("D8").Value = "Result": .Range("E8").Value = "When": .Range("F8").Value = "Note"
        .Range("B8:G8").Font.Bold = True
        .Columns("A").ColumnWidth = 2: .Columns("B").ColumnWidth = 6: .Columns("C").ColumnWidth = 70
        .Columns("D").ColumnWidth = 10: .Columns("E").ColumnWidth = 16: .Columns("F").ColumnWidth = 60: .Columns("G").ColumnWidth = 14
        .Columns("K:P").Hidden = True
    End With
    InstButton ws, "inst_pick", "Pick the update folder...", "InstallPickFolder", ws.Range("C7").Left, ws.Range("C7").Top, 150, True
    InstButton ws, "inst_refresh", "Refresh", "InstallRefresh", ws.Range("C7").Left + 160, ws.Range("C7").Top, 70, True
    InstButton ws, "inst_skip", "Skip this step...", "InstallSkip", ws.Range("C7").Left + 240, ws.Range("C7").Top, 110, True
    InstButton ws, "inst_log", "Show the log", "InstallShowLog", ws.Range("C7").Left + 360, ws.Range("C7").Top, 90, True
End Sub

Private Sub InstButton(ByVal ws As Worksheet, ByVal nm As String, ByVal cap As String, ByVal macro As String, _
                       ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal enabled As Boolean)
    Dim shp As Object
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(5, x, y, w, 18)       ' 5 = msoShapeRoundedRectangle
    shp.Name = nm
    shp.TextFrame.Characters.Text = cap
    shp.TextFrame.HorizontalAlignment = -4108          ' xlCenter
    shp.TextFrame.Characters.Font.Size = 9
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.Characters.Font.Color = IIf(enabled, RGB(255, 255, 255), RGB(90, 90, 90))
    shp.Fill.ForeColor.RGB = IIf(enabled, RGB(31, 78, 121), RGB(217, 217, 217))
    shp.Line.Visible = False
    shp.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macro
End Sub

' ---------------- the step list ----------------
' install-steps.txt -> the hidden columns K:P (one row per step from INST_ROW0) and C5 (VERSION), B7 (header line)
Private Function InstParse(ByVal path As String, ByVal ws As Worksheet, ByRef why As String) As Boolean
    Dim txt As String
    If Not UpdReadFile(path, txt) Then why = "it couldn't be read": Exit Function
    Dim lines() As String: lines = Split(Replace(txt, vbCr, ""), vbLf)
    Dim i As Long, j As Long, ln As String, f() As String, n As Long, grp As String, ver As String, expect As String, hdr As String
    Dim rows() As String: ReDim rows(1 To INST_MAX, 1 To 6)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If ln = "" Or Left$(ln, 1) = "#" Then
            ' skip
        ElseIf Left$(ln, 1) = "[" Then
            If Len(ln) > 2 Then grp = Mid$(ln, 2, Len(ln) - 2)
        ElseIf InStr(ln, "|") = 0 Then
            If UCase$(Left$(ln, 8)) = "VERSION " Then ver = Trim$(Mid$(ln, 9))
            If UCase$(Left$(ln, 7)) = "EXPECT " Then expect = Trim$(Mid$(ln, 8))
            If UCase$(Left$(ln, 7)) = "MEMBER " Or UCase$(Left$(ln, 6)) = "FORMS " Then hdr = hdr & IIf(hdr = "", "", "   ") & ln
        Else
            f = Split(ln, "|")
            If Trim$(f(0)) = "" Then why = "line " & (i + 1) & " has no step kind": Exit Function
            If UBound(f) < 1 Then why = "line " & (i + 1) & " has no caption": Exit Function
            If Trim$(f(1)) = "" Then why = "line " & (i + 1) & " has no caption": Exit Function
            Select Case LCase$(Trim$(f(0)))
                Case "backup", "check", "swap", "compile", "finish", "macro", "template", "manual", "test"
                Case Else: why = "line " & (i + 1) & ": '" & Trim$(f(0)) & "' isn't a step kind": Exit Function
            End Select
            n = n + 1
            If n > INST_MAX Then why = "more than " & INST_MAX & " steps": Exit Function
            rows(n, 1) = LCase$(Trim$(f(0)))
            rows(n, 2) = Trim$(f(1))
            If UBound(f) >= 2 Then rows(n, 3) = Trim$(f(2))
            If UBound(f) >= 3 Then rows(n, 4) = Trim$(f(3))
            rows(n, 5) = "S" & Format$(n, "00")
            rows(n, 6) = grp
            grp = ""
        End If
    Next i
    If n = 0 Then why = "no steps in it": Exit Function
    If ver = "" Then why = "no VERSION line": Exit Function
    ' only now the sheet changes (a bad file leaves the old list alone)
    Dim rng As Range: Set rng = ws.Range(ws.Cells(INST_ROW0, INST_COL_KIND), ws.Cells(INST_ROW0 + INST_MAX, INST_COL_KIND + 5))
    rng.ClearContents
    rng.NumberFormat = "@"                              ' "Oct 12", "1,234", "TRUE" stay the words they are
    For i = 1 To n
        For j = 1 To 6
            ws.Cells(INST_ROW0 + i - 1, INST_COL_KIND + j - 1).Value = rows(i, j)
        Next j
    Next i
    ws.Range("C5").NumberFormat = "@"
    ws.Range("C5").Value = ver & IIf(expect <> "" And expect <> ver, " (SystemCheck expects " & expect & ")", "")
    ws.Range("B7").Value = hdr
    InstParse = True
End Function

Private Function InstSteps(ByVal ws As Worksheet, ByRef steps() As InstStep) As Long
    Dim n As Long, r As Long
    ReDim steps(1 To INST_MAX)
    For r = INST_ROW0 To INST_ROW0 + INST_MAX - 1
        If Trim$(CStr(ws.Cells(r, INST_COL_KIND).Value)) = "" Then Exit For
        n = n + 1
        steps(n).Kind = CStr(ws.Cells(r, INST_COL_KIND).Value)
        steps(n).Caption = CStr(ws.Cells(r, INST_COL_KIND + 1).Value)
        steps(n).Arg = CStr(ws.Cells(r, INST_COL_KIND + 2).Value)
        steps(n).Flag = CStr(ws.Cells(r, INST_COL_KIND + 3).Value)
        steps(n).Id = CStr(ws.Cells(r, INST_COL_KIND + 4).Value)
        steps(n).Group = CStr(ws.Cells(r, INST_COL_KIND + 5).Value)
    Next r
    InstSteps = n
End Function

Private Function InstVersion(ByVal ws As Worksheet) As String
    Dim s As String: s = Trim$(CStr(ws.Range("C5").Value))
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    InstVersion = s
End Function

' ---------------- the log (in the roster) ----------------
' The Install Log sheet and its table, made on first use. Nothing = why (the sheet is protected:
' open the roster in admin view; or the workbook can't take a sheet).
Private Function InstEnsureLog(ByVal wb As Workbook, ByRef why As String) As Object
    Dim ws As Worksheet, lo As Object
    On Error Resume Next
    Set ws = wb.Worksheets(INST_LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        If wb.ProtectStructure Then                   ' v25: only the roster's own code can add a sheet now
            why = "The roster has no " & INST_LOG_SHEET & " sheet yet and its structure is protected (v25): open it in ADMIN view once " & _
                  "(it makes the sheet at open), then click the step again."
            Exit Function
        End If
        Dim evWas As Boolean: evWas = Application.EnableEvents
        Dim wasSaved As Boolean: wasSaved = wb.Saved
        Dim actWas As Object: Set actWas = ActiveSheet
        On Error GoTo Failed
        Application.EnableEvents = False                ' (review M1) no roster sheet events while the sheet is made
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = INST_LOG_SHEET
        ws.Range("A1:G1").Value = Array("Version", "Step", "Kind", "Status", "When", "User", "Note")
        ws.Range("A1:G1").Font.Bold = True
        Set lo = ws.ListObjects.Add(1, ws.Range("A1:G1"), , 1)    ' xlSrcRange, xlYes
        lo.Name = INST_LOG_TABLE
        ws.Columns("A:G").ColumnWidth = 18: ws.Columns("B").ColumnWidth = 60: ws.Columns("G").ColumnWidth = 70
        ws.Visible = 2                                 ' xlSheetVeryHidden: the roster's ApplyView shows it to admins
        On Error Resume Next
        actWas.Parent.Activate                         ' (Worksheets.Add made the roster the active workbook)
        actWas.Activate
        On Error GoTo 0
        Application.EnableEvents = evWas
        wb.Saved = wasSaved
    Else
        On Error Resume Next
        Set lo = ws.ListObjects(INST_LOG_TABLE)
        On Error GoTo 0
        If lo Is Nothing Then why = "The roster's " & INST_LOG_SHEET & " sheet has no " & INST_LOG_TABLE & " table.": Exit Function
    End If
    If ws.ProtectContents Then                      ' (no Unprotect call: without the password Excel would prompt for one)
        why = "The roster's " & INST_LOG_SHEET & " sheet is protected: open the roster in ADMIN view (its ProtectSheet leaves this sheet open), then click the step again."
        Exit Function
    End If
    Set InstEnsureLog = lo
    Exit Function
Failed:
    why = "The " & INST_LOG_SHEET & " sheet couldn't be made in " & wb.Name & ": " & Err.Description
    On Error Resume Next                                ' (re-review) a half-made sheet goes; the roster is as it was
    Application.DisplayAlerts = False
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    Application.EnableEvents = evWas
    wb.Saved = wasSaved
    actWas.Parent.Activate
    actWas.Activate
End Function

Private Sub InstLog(ByVal wb As Workbook, ByVal ver As String, ByRef st As InstStep, ByVal status As String, ByVal note As String)
    Dim why As String
    Dim lo As Object: Set lo = InstEnsureLog(wb, why)
    If lo Is Nothing Then Err.Raise vbObjectError + 950, , why
    ' (review H2 / M1) a log row must not make the roster "unsaved" (the check step wants it saved) and must not
    ' run the roster's sheet events (right after a swap that code isn't compiled yet)
    Dim wasSaved As Boolean: wasSaved = wb.Saved
    Dim evWas As Boolean: evWas = Application.EnableEvents
    Dim en As Long, em As String
    Application.EnableEvents = False
    On Error GoTo Failed
    Dim lr As Object: Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 1).Value = ver
    lr.Range.Cells(1, 2).Value = st.Id & "  " & st.Caption
    lr.Range.Cells(1, 3).Value = st.Kind
    lr.Range.Cells(1, 4).Value = status
    lr.Range.Cells(1, 5).Value = Now
    lr.Range.Cells(1, 6).Value = Environ$("USERNAME")
    lr.Range.Cells(1, 7).Value = "'" & Left$(note, 250)
    Application.EnableEvents = evWas
    wb.Saved = wasSaved
    Exit Sub
Failed:
    en = Err.Number: em = Err.Description
    Application.EnableEvents = evWas
    On Error Resume Next
    wb.Saved = wasSaved
    On Error GoTo 0
    Err.Raise en, , em
End Sub

' The latest row per step id for this version: key = step id, item = status & "|" & when & "|" & note
Private Function InstReadLog(ByVal wb As Workbook, ByVal ver As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set InstReadLog = d
    Dim ws As Worksheet, lo As Object, v As Variant, i As Long, id As String
    On Error Resume Next
    Set ws = wb.Worksheets(INST_LOG_SHEET)
    If ws Is Nothing Then Exit Function
    Set lo = ws.ListObjects(INST_LOG_TABLE)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = lo.DataBodyRange.Value
    On Error GoTo 0
    If Not IsArray(v) Then Exit Function
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(CStr(v(i, 1))), ver, vbTextCompare) = 0 Then
            id = Trim$(CStr(v(i, 2)))
            If InStr(id, " ") > 0 Then id = Left$(id, InStr(id, " ") - 1)
            d(id) = CStr(v(i, 4)) & "|" & Format$(v(i, 5), "mmm d hh:nn") & "|" & CStr(v(i, 7))
        End If
    Next i
End Function

Private Function InstStatusOf(ByVal st As Object, ByVal id As String) As String
    If st.Exists(id) Then InstStatusOf = Split(st(id), "|")(0)
End Function

Private Function InstIsDone(ByVal status As String) As Boolean
    Select Case UCase$(status)
        Case "OK", "PASS", "DONE", "SKIPPED": InstIsDone = True
    End Select
End Function

' The next step to run (0 = all done)
Private Function InstNext(ByRef steps() As InstStep, ByVal n As Long, ByVal st As Object) As Long
    Dim i As Long
    For i = 1 To n
        If Not InstIsDone(InstStatusOf(st, steps(i).Id)) Then InstNext = i: Exit Function
    Next i
End Function

' ---------------- drawing ----------------
Private Sub InstRefresh()
    Dim ws As Worksheet: Set ws = InstSheet(True)
    Dim steps() As InstStep, n As Long, i As Long, r As Long
    n = InstSteps(ws, steps)
    ' old rows and buttons off
    ws.Range(ws.Cells(INST_ROW0, 2), ws.Cells(INST_ROW0 + 2 * INST_MAX, 7)).Clear
    Dim k As Long
    For k = ws.Shapes.Count To 1 Step -1              ' backwards: deleting inside For Each skips shapes
        If Left$(ws.Shapes(k).Name, 10) = "inst_step_" Then ws.Shapes(k).Delete
    Next k
    If n = 0 Then
        ws.Range("C6").Value = "Pick the update folder (the unzipped release that holds install-steps.txt)."
        Exit Sub
    End If
    Dim wb As Workbook, st As Object, nxt As Long, status As String, parts() As String
    On Error Resume Next
    Set wb = InstRosterQuiet()
    On Error GoTo 0
    If wb Is Nothing Then
        ws.Range("C4").Value = "(no roster open in this Excel - open it, then Refresh)"
        Set st = CreateObject("Scripting.Dictionary")
    Else
        ws.Range("C4").Value = wb.FullName
        Set st = InstReadLog(wb, InstVersion(ws))
    End If
    nxt = InstNext(steps, n, st)
    r = INST_ROW0
    For i = 1 To n
        If steps(i).Group <> "" Then
            ws.Cells(r, 3).Value = steps(i).Group
            ws.Cells(r, 3).Font.Bold = True
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Interior.Color = RGB(221, 235, 247)
            r = r + 1
        End If
        ws.Cells(r, 2).Value = steps(i).Id
        ws.Cells(r, 3).Value = steps(i).Caption & IIf(steps(i).Kind = "test" And steps(i).Flag <> "", vbLf & "   expect: " & steps(i).Flag, "")
        ws.Cells(r, 3).WrapText = True
        status = InstStatusOf(st, steps(i).Id)
        If status <> "" Then
            parts = Split(st(steps(i).Id), "|")
            ws.Cells(r, 4).Value = parts(0)
            ws.Cells(r, 5).Value = parts(1)
            ws.Cells(r, 6).Value = parts(2)
            ws.Cells(r, 6).WrapText = True
            ws.Cells(r, 4).Font.Color = IIf(InstIsDone(status), RGB(0, 110, 0), RGB(180, 0, 0))
            ws.Cells(r, 4).Font.Bold = True
        End If
        If Not InstIsDone(status) Then
            InstButton ws, "inst_step_" & i, InstButtonText(steps(i), status), "InstallStepClick", _
                       ws.Cells(r, 7).Left + 2, ws.Cells(r, 7).Top + 2, ws.Columns(7).Width - 4, (i = nxt)
        End If
        r = r + 1
    Next i
    ws.Rows(INST_ROW0 & ":" & r).AutoFit
    If nxt = 0 Then
        ws.Range("C6").Value = "Every step is done for " & InstVersion(ws) & ". The record is on the roster's Install Log sheet."
    ElseIf wb Is Nothing Then
        ws.Range("C6").Value = "Open the roster (and nothing else that looks like one), then Refresh."
    Else
        ws.Range("C6").Value = "Next: " & steps(nxt).Id & "  " & steps(nxt).Caption
    End If
End Sub

Private Function InstButtonText(ByRef st As InstStep, ByVal status As String) As String
    Select Case st.Kind
        Case "manual": InstButtonText = "Done"
        Case "test": InstButtonText = "Pass / Fail"
        Case Else: InstButtonText = IIf(UCase$(status) = "FAILED" Or UCase$(status) = "FAIL", "Retry", "Run")
    End Select
End Function

' The roster for the panel: the one open roster, or, with a practice COPY (or a backup / FINAL opened to
' look at) beside it, the one whose name carries none of those words; failing that the one in front.
Private Function InstRosterQuiet() As Workbook
    Dim w As Workbook, why As String, n As Long, hit As Workbook, nLive As Long, live As Workbook, u As String
    For Each w In Application.Workbooks
        If Not UpdSame(w, ThisWorkbook) Then
            If UpdIsRoster(w, why) Then
                n = n + 1: Set hit = w
                u = UCase$(w.Name)
                If InStr(u, "COPY") = 0 And InStr(u, "DRY RUN") = 0 And InStr(u, "BACKUP") = 0 And InStr(u, "FINAL") = 0 Then
                    nLive = nLive + 1: Set live = w
                End If
            End If
        End If
    Next w
    If n = 1 Then Set InstRosterQuiet = hit: Exit Function
    If nLive = 1 Then Set InstRosterQuiet = live: Exit Function
    If n > 1 Then
        On Error Resume Next
        If UpdIsRoster(ActiveWorkbook, why) Then Set InstRosterQuiet = ActiveWorkbook
    End If
End Function

' InstRosterQuiet with the message when there is none
Private Function InstTarget() As Workbook
    Set InstTarget = InstRosterQuiet()
    If InstTarget Is Nothing Then
        MsgBox "Open the roster in this Excel (one roster; a practice COPY beside it is passed over), then click again.", vbExclamation, INST_TITLE
    End If
End Function

' The open workbook at that path, or Nothing
Private Function InstByPath(ByVal fullPath As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, fullPath, vbTextCompare) = 0 Then Set InstByPath = w: Exit Function
    Next w
End Function

' ---------------- running a step ----------------
Private Sub InstRunIndex(ByVal i As Long)
    Dim ws As Worksheet: Set ws = InstSheet(False)
    If ws Is Nothing Then Exit Sub
    Dim steps() As InstStep, n As Long
    n = InstSteps(ws, steps)
    If i < 1 Or i > n Then Exit Sub
    Dim folder As String: folder = UpdSlash(Trim$(CStr(ws.Range("C3").Value)))
    If folder = "" Or Not UpdDirExists(folder) Then MsgBox "Pick the update folder first.", vbInformation, INST_TITLE: Exit Sub
    Dim wb As Workbook: Set wb = InstTarget()
    If wb Is Nothing Then Exit Sub
    Dim why As String
    If InstEnsureLog(wb, why) Is Nothing Then MsgBox why, vbExclamation, INST_TITLE: Exit Sub     ' (review M2) before the step, not after
    Dim ver As String: ver = InstVersion(ws)
    Dim st As Object: Set st = InstReadLog(wb, ver)
    Dim nxt As Long: nxt = InstNext(steps, n, st)
    If nxt = 0 Then MsgBox "Every step is done.", vbInformation, INST_TITLE: InstRefresh: Exit Sub
    If i <> nxt Then
        MsgBox "Step " & steps(nxt).Id & " comes first:" & vbNewLine & steps(nxt).Caption & vbNewLine & vbNewLine & _
               "(Skip this step... marks it skipped with a note if it was done by hand.)", vbInformation, INST_TITLE
        Exit Sub
    End If
    Dim status As String, note As String, fp As String
    fp = wb.FullName                                    ' (review H3) the roster is found again by its PATH: a practice
    ws.Range("C6").Value = "Running " & steps(i).Id & " " & steps(i).Caption & " ..."     '  COPY a step opened never takes the log
    InstRunStep wb, folder, ver, steps(i), status, note
    If status <> "" Then
        Set wb = InstByPath(fp)                         ' (a reopen keeps the path)
        If wb Is Nothing Then
            MsgBox "Step " & steps(i).Id & " ended " & status & " but the roster (" & fp & ") can't be found to record it: " & note & vbNewLine & vbNewLine & _
                   "Open it again, then Skip this step... with that result as the note.", vbExclamation, INST_TITLE
        Else
            On Error Resume Next
            InstLog wb, ver, steps(i), status, note
            why = Err.Description
            On Error GoTo 0
            If why <> "" Then
                MsgBox "Step " & steps(i).Id & " ended " & status & " (" & note & ") but the row couldn't be written to the roster's Install Log: " & why & _
                       vbNewLine & vbNewLine & "Fix that (ADMIN view), then Skip this step... with that result as the note.", vbExclamation, INST_TITLE
            End If
        End If
    End If
    InstRefresh
    If status = "FAILED" Or status = "FAIL" Then
        MsgBox steps(i).Id & " " & steps(i).Caption & ": " & status & vbNewLine & vbNewLine & note & vbNewLine & vbNewLine & _
               "Retry it, or Skip this step... with a note.", vbExclamation, INST_TITLE
    End If
End Sub

' One step. status "" = nothing to record (cancelled); otherwise OK / FAILED / DONE / PASS / FAIL and a note.
Private Sub InstRunStep(ByRef wb As Workbook, ByVal folder As String, ByVal ver As String, ByRef st As InstStep, _
                        ByRef status As String, ByRef note As String)
    Dim why As String, ok As Boolean, bk As String, repPath As String, rep As String
    Dim items() As UpdItem, n As Long, proj As Object, trusted As Boolean
    status = "": note = ""
    Set proj = UpdProject(wb, why)
    trusted = Not proj Is Nothing
    Select Case st.Kind
        Case "backup"
            If UpdBackup(wb, ver, bk, why) Then status = "OK": note = bk Else status = "FAILED": note = why

        Case "check"
            If wb.Path = "" Or LCase$(Left$(wb.Path, 4)) = "http" Then
                status = "FAILED": note = "Open the roster from an ordinary folder (not SharePoint / OneDrive)."
            ElseIf wb.ReadOnly Then
                status = "FAILED": note = "The roster is open read-only."
            ElseIf UpdNameSet(wb, UPD_PENDING) Or UpdNameSet(wb, UPD_NOSAVE) Then
                status = "FAILED": note = "Saving is already blocked in the roster (an earlier update or check): close it WITHOUT saving and open it again."
            ElseIf Not wb.Saved Then
                status = "FAILED": note = "The roster has unsaved changes: press Ctrl+S in it, then Retry."
            ElseIf Not UpdIsAdmin(wb) Then
                status = "FAILED": note = "Not an admin login."
            ElseIf Not trusted Then
                status = "OK": note = "Trust access to the VBA project is OFF: the swap and compile steps will show the by-hand list. " & why
            ElseIf Not UpdGatherRoster(folder, proj, wb, items, n, why) Then
                status = "FAILED": note = why
            Else
                status = "OK": note = n & " code files in the folder; the roster is on " & IIf(UpdCurrentRelease(proj) = "", "(unknown)", UpdCurrentRelease(proj)) & _
                                      "; the folder holds " & ver & "."
            End If

        Case "swap"
            If Not trusted Then
                InstManualSwap folder, status, note
            ElseIf UpdNameSet(wb, UPD_PENDING) Then
                status = "FAILED": note = "A swap is already waiting for its finish in the roster (or an earlier one was dropped): close it WITHOUT saving, open it again and Retry."
            ElseIf Not UpdGatherRoster(folder, proj, wb, items, n, why) Then
                status = "FAILED": note = why
            Else
                rep = UPD_VERSION & "   " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbNewLine & wb.FullName & vbNewLine & "Install panel, " & ver & vbNewLine & vbNewLine
                If UpdSwapAll(wb, items, n, rep, repPath, why) Then
                    status = "OK": note = n & " components swapped; not saved yet. Report: " & repPath
                Else
                    status = "FAILED"
                    If repPath = "" Then            ' the save block couldn't be set: nothing was touched
                        note = why
                    Else
                        note = why & " Report: " & repPath & ". Close the roster WITHOUT saving (Don't Save), open it again, then Retry or install by hand."
                    End If
                End If
            End If

        Case "compile"
            If Not trusted Then
                If MsgBox("Alt+F11, click the ROSTER's project, Debug > Compile VBAProject." & vbNewLine & vbNewLine & _
                          "Did it finish with NO message?", vbYesNo + vbQuestion, INST_TITLE) = vbYes Then
                    status = "DONE": note = "compiled by hand"
                Else
                    status = "FAILED": note = "a compile message: note the line it highlights"
                End If
            Else
                Select Case UpdCompile(wb, why)
                    Case 1: status = "OK": note = "Compile VBAProject: clean"
                    Case 0: status = "FAILED": note = why
                    Case Else
                        If MsgBox("The compile couldn't be run from here (" & why & ")." & vbNewLine & vbNewLine & _
                                  "Alt+F11, click the ROSTER's project, Debug > Compile VBAProject. Did it finish with NO message?", _
                                  vbYesNo + vbQuestion, INST_TITLE) = vbYes Then
                            status = "DONE": note = "compiled by hand (" & why & ")"
                        Else
                            status = "FAILED": note = "a compile message: note the line it highlights"
                        End If
                End Select
            End If

        Case "finish"
            If wb.ReadOnly Then
                status = "FAILED": note = "The roster is open read-only, so it can't be saved. Close it; nothing is kept."
                Exit Sub
            End If
            If Not Application.EnableEvents Then
                status = "FAILED": note = "Excel's events are switched off in this session (an earlier macro was stopped part-way): " & _
                                           "Alt+F8 > UpdateDrop, close Excel completely, open the roster again and resume."
                Exit Sub
            End If
            If Not UpdNameSet(wb, UPD_PENDING) Then
                If Not trusted Then
                    If MsgBox("Save the roster (Ctrl+S) now. Saved?", vbYesNo + vbQuestion, INST_TITLE) = vbYes Then status = "DONE": note = "saved by hand" Else Exit Sub
                Else
                    status = "OK": note = "nothing was waiting to be finished (already saved)"
                End If
            ElseIf UpdFinish(wb, why) Then
                status = "OK": note = "saved with the new code"
            Else
                status = "FAILED": note = why
                Exit Sub
            End If
            If LCase$(st.Arg) = "reopen" Or LCase$(st.Flag) = "reopen" Then
                If UpdReopen(wb, why) Then
                    note = note & "; closed and reopened"
                Else
                    note = note & "; it could NOT be reopened by the panel (" & why & "): open it yourself, then Refresh"
                End If
            End If

        Case "template"
            If UpdTemplate(wb, folder) Then status = "OK": note = "members' template rebuilt and saved" Else status = "FAILED": note = "see the box"

        Case "macro"
            InstRunMacro wb, st, status, note

        Case "manual"
            If MsgBox(st.Caption & IIf(st.Arg <> "", vbNewLine & vbNewLine & st.Arg, "") & vbNewLine & vbNewLine & _
                      "OK = done.   Cancel = not yet.", vbOKCancel + vbInformation, INST_TITLE) = vbOK Then status = "DONE": note = "done by hand"

        Case "test"
            InstRunTest wb, st, status, note

        Case Else
            status = "FAILED": note = "unknown step kind '" & st.Kind & "' in install-steps.txt"
    End Select
End Sub

' Trust access off: the exact by-hand list for the swap, and a Done question
Private Sub InstManualSwap(ByVal folder As String, ByRef status As String, ByRef note As String)
    Dim names() As String, k As Long, i As Long, s As String
    k = UpdListFiles(folder & "modules\", "*.bas", names)
    s = "BY HAND (trust access to the VBA project is off):" & vbNewLine & vbNewLine & "Remove and import, from " & folder & "modules\:" & vbNewLine
    For i = 1 To k
        If StrComp(names(i), "modUpdater.bas", vbTextCompare) <> 0 Then s = s & "  " & names(i) & vbNewLine
    Next i
    k = UpdListFiles(folder & "paste-in\", "*.txt", names)
    s = s & vbNewLine & "Re-paste (View Code, Ctrl+A, paste), from " & folder & "paste-in\:" & vbNewLine
    For i = 1 To k
        s = s & "  " & names(i) & vbNewLine
    Next i
    Dim p As String: p = UpdWriteReport(ThisWorkbook, s, "By-hand install list")
    If MsgBox(Left$(s, 900) & IIf(Len(s) > 900, "..." & vbNewLine, "") & vbNewLine & "The full list is in " & p & vbNewLine & vbNewLine & _
              "OK = done by hand.   Cancel = not yet.", vbOKCancel + vbInformation, INST_TITLE) = vbOK Then
        status = "DONE": note = "swapped by hand; list: " & p
    End If
End Sub

' Compile VBAProject on the roster's project through the VBA editor's own command.
' 1 = clean, 0 = an error (the command stays enabled), -1 = couldn't be run (why).
Private Function UpdCompile(ByVal wb As Workbook, ByRef why As String) As Long
    Dim proj As Object: Set proj = UpdProject(wb, why)
    If proj Is Nothing Then UpdCompile = -1: Exit Function
    Dim vbe As Object, ctl As Object
    On Error Resume Next
    Set vbe = Application.VBE
    If vbe Is Nothing Then why = "the VBA editor object isn't reachable": UpdCompile = -1: Exit Function
    Err.Clear
    Set vbe.ActiveVBProject = proj
    If Err.Number <> 0 Then why = "the roster's project couldn't be made the active one: " & Err.Description: UpdCompile = -1: Exit Function
    Err.Clear
    If StrComp(vbe.ActiveVBProject.fileName, wb.FullName, vbTextCompare) <> 0 Then
        why = "the VBA editor's active project isn't the roster's": UpdCompile = -1: Exit Function
    End If
    Set ctl = vbe.CommandBars.FindControl(, 578)        ' 578 = Debug > Compile VBAProject
    If ctl Is Nothing Then why = "the Compile command wasn't found in the VBA editor": UpdCompile = -1: Exit Function
    If Not ctl.Enabled Then why = "the Compile command is already disabled (compiled before, or not this project)": UpdCompile = -1: Exit Function
    Err.Clear
    ctl.Execute
    If Err.Number <> 0 Then why = "the Compile command didn't run: " & Err.Description: UpdCompile = -1: Exit Function
    DoEvents
    Set ctl = vbe.CommandBars.FindControl(, 578)
    If ctl Is Nothing Then why = "the Compile command disappeared": UpdCompile = -1: Exit Function
    If ctl.Enabled Then
        why = "Compile VBAProject found an error (the command is still enabled). Alt+F11 shows the line it highlighted; " & _
              "note it, then Alt+F8 > UpdateDrop closes the roster without saving (the file on disk is as it was)."
        UpdCompile = 0
    Else
        UpdCompile = 1
    End If
End Function

' Close and open the roster again (after a save): the open code runs, the admin view is fresh
Private Function UpdReopen(ByRef wb As Workbook, ByRef why As String) As Boolean
    Dim fullPath As String: fullPath = wb.FullName
    On Error GoTo Failed
    If Not wb.Saved Then wb.Save
    wb.Close SaveChanges:=False
    Set wb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0)
    ThisWorkbook.Activate
    UpdReopen = True
    Exit Function
Failed:
    why = Err.Description
End Function

' A roster macro: "Name arg arg ..." (a Function returning Boolean marks itself; a Sub is OK when it
' raised nothing; "| confirm" asks after it)
Private Sub InstRunMacro(ByVal wb As Workbook, ByRef st As InstStep, ByRef status As String, ByRef note As String)
    Dim nm As String, args() As Variant, k As Long, v As Variant, en As Long, em As String
    If Not InstParseCall(st.Arg, nm, args, k, em) Then status = "FAILED": note = em: Exit Sub
    If nm = "" Then status = "FAILED": note = "no macro named on the step": Exit Sub
    On Error Resume Next
    v = InstRun(wb, nm, args, k)
    en = Err.Number: em = Err.Description
    On Error GoTo 0
    If en <> 0 Then status = "FAILED": note = nm & ": " & em: Exit Sub
    If VarType(v) = vbBoolean Then
        If CBool(v) Then status = "OK": note = nm & " said yes" Else status = "FAILED": note = nm & " said no (see its box)"
        If LCase$(st.Flag) <> "confirm" Then Exit Sub
        If status = "FAILED" Then Exit Sub
    End If
    If LCase$(st.Flag) = "confirm" Or VarType(v) <> vbBoolean Then
        Select Case MsgBox(st.Caption & vbNewLine & vbNewLine & "Did " & nm & " finish without a problem?" & vbNewLine & vbNewLine & _
                           "Yes = OK.   No = failed (say why next).   Cancel = record nothing.", vbYesNoCancel + vbQuestion, INST_TITLE)
            Case vbYes: status = "OK": note = nm & " confirmed"
            Case vbNo: status = "FAILED": note = Trim$(VBA.InputBox("What went wrong with " & nm & "?", INST_TITLE)): If note = "" Then note = "failed (no note)"
            Case Else: status = "": note = ""
        End Select
    End If
End Sub

' A test row: the opener (if any) runs, then Pass / Fail against the expected result
Private Sub InstRunTest(ByVal wb As Workbook, ByRef st As InstStep, ByRef status As String, ByRef note As String)
    Dim nm As String, args() As Variant, k As Long, en As Long, em As String, opened As String
    If Trim$(st.Arg) <> "" Then
        If Not InstParseCall(st.Arg, nm, args, k, em) Then status = "FAILED": note = em: Exit Sub
        On Error Resume Next
        InstRun wb, nm, args, k
        en = Err.Number: em = Err.Description
        On Error GoTo 0
        If en <> 0 Then status = "FAILED": note = nm & " couldn't be run: " & em: Exit Sub
        opened = nm & InstArgsText(args, k)
    End If
    Select Case MsgBox(st.Caption & vbNewLine & vbNewLine & "Expected: " & st.Flag & vbNewLine & vbNewLine & _
                       "Yes = PASS.   No = FAIL (say what you saw next).   Cancel = record nothing.", vbYesNoCancel + vbQuestion, INST_TITLE)
        Case vbYes: status = "PASS": note = IIf(opened = "", "checked", opened)
        Case vbNo
            status = "FAIL"
            note = Trim$(VBA.InputBox("What did you see?", INST_TITLE))
            If note = "" Then note = "failed (no note)"
            If opened <> "" Then note = opened & ": " & note
        Case Else: status = "": note = ""
    End Select
End Sub

' "SmokePosting Phosphine +21 Days" -> name + typed arguments (+N / -N = today plus N days as a Date;
' yyyy-mm-dd = that date; a whole number = a number; anything else = text)
Private Function InstParseCall(ByVal txt As String, ByRef nm As String, ByRef args() As Variant, ByRef k As Long, ByRef why As String) As Boolean
    Dim toks() As String, i As Long, s As String, inQ As Boolean, c As String, p As Long, nTok As Long
    ReDim args(1 To 4)
    ReDim toks(1 To 16)
    k = 0: nm = ""
    txt = Trim$(txt)
    If txt = "" Then InstParseCall = True: Exit Function
    ' split on spaces; a "quoted part" is one token (a unit like "Material Handler")
    For p = 1 To Len(txt)
        c = Mid$(txt, p, 1)
        If c = Chr$(34) Then
            inQ = Not inQ
        ElseIf c = " " And Not inQ Then
            If s <> "" Then
                nTok = nTok + 1
                If nTok > 16 Then why = "too many words on the step": Exit Function
                toks(nTok) = s
                s = ""
            End If
        Else
            s = s & c
        End If
    Next p
    If inQ Then why = "a quote isn't closed on the step": Exit Function
    If s <> "" Then
        nTok = nTok + 1
        If nTok > 16 Then why = "too many words on the step": Exit Function
        toks(nTok) = s
    End If
    If nTok = 0 Then InstParseCall = True: Exit Function
    nm = toks(1)
    For i = 2 To nTok
        s = toks(i)
        k = k + 1
        If k > 4 Then why = "more than 4 arguments on " & nm: Exit Function
        If (Left$(s, 1) = "+" Or Left$(s, 1) = "-") And IsNumeric(Mid$(s, 2)) And InStr(s, ".") = 0 Then
            args(k) = CDate(Date + CLng(s))
        ElseIf Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" And IsNumeric(Left$(s, 4)) Then
            args(k) = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 6, 2)), CLng(Mid$(s, 9, 2)))
        ElseIf IsNumeric(s) Then
            args(k) = CDbl(s)
        Else
            args(k) = s
        End If
    Next i
    InstParseCall = True
End Function

Private Function InstArgsText(ByRef args() As Variant, ByVal k As Long) As String
    Dim i As Long, s As String
    For i = 1 To k
        s = s & " " & IIf(VarType(args(i)) = vbDate, Format$(args(i), "ddd mmm d"), CStr(args(i)))
    Next i
    InstArgsText = s
End Function

' Application.Run with 0 to 4 arguments (an error raises to the caller)
Private Function InstRun(ByVal wb As Workbook, ByVal nm As String, ByRef args() As Variant, ByVal k As Long) As Variant
    Dim full As String: full = "'" & Replace(wb.Name, "'", "''") & "'!" & nm
    Select Case k
        Case 0: InstRun = Application.Run(full)
        Case 1: InstRun = Application.Run(full, args(1))
        Case 2: InstRun = Application.Run(full, args(1), args(2))
        Case 3: InstRun = Application.Run(full, args(1), args(2), args(3))
        Case Else: InstRun = Application.Run(full, args(1), args(2), args(3), args(4))
    End Select
End Function
