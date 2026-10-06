' =====================================================================
'  EXCEL AUDIT  v3.1  (VBA)
'
'  What changed against v2
'    - RESULT INSIDE AuditTool: three sheets are written at the end
'        Audit_Summary   totals + one row per sheet
'        Audit_Findings  error cells and suspicious cells
'        Audit_Changes   formulas that changed between the old and the new file
'    - SUMMARY, NOT EVERY CELL: one row per DIFFERENT formula of a sheet
'      (with the number of cells that use it), not one row per cell
'    - FAST WRITING: output is collected in memory and written once per sheet.
'      No file stays open during the run.
'
'  INSTALL: open AuditTool.xlsm > Alt+F11 > Modulo1 > Ctrl+A > paste this > Ctrl+S
'  RUN    : Alt+F8 > RunAudit > Esegui
'  Output folder: C:\Users\<you>\AuditOut (local disk), wherever AuditTool itself is saved.
'  Do not open the CSV files while it runs. Progress: audit_log.txt in that folder (Notepad).
' =====================================================================

Option Explicit

Private Const MAX_CELLS As Long = 200000     ' cells read per block
Private Const MAX_LIST As Long = 20000       ' max rows listed on a report sheet

Private mStage As String
Private mLog As String
Private mBuf() As String
Private mBufN As Long

Private mSheets As Collection        ' one item per sheet (both files)
Private mFinds As Collection         ' errors and suspicious cells of the NEW file
Private mChk As Object               ' NEW file: sheet + check -> number of cells
Private mCnt(0 To 1) As Object       ' 0 = New, 1 = Old: sheet + formula -> number of cells
Private mEx(0 To 1) As Object        ' sheet + formula -> example cell
Private mShNames(0 To 1) As Object   ' sheet names

Public Sub RunAudit()
    Dim newPath As String, oldPath As String, outDir As String, runName As String
    Dim calc As XlCalculation, sec As Long, t0 As Double, i As Long
    Dim totF As Long, totE As Long, totFl As Long

    If ThisWorkbook.path = "" Then
        MsgBox "Save AuditTool.xlsm first.", vbExclamation
        Exit Sub
    End If

    newPath = PickFile("Select the NEW (current) workbook")
    If newPath = "" Then Exit Sub
    oldPath = PickFile("Select the OLD (previous) workbook - Cancel to skip the comparison")

    calc = Application.Calculation
    sec = Application.AutomationSecurity
    t0 = Timer
    On Error GoTo Fail

    mStage = "preparing"
    Set mSheets = New Collection
    Set mFinds = New Collection
    Set mChk = CreateObject("Scripting.Dictionary")
    For i = 0 To 1
        Set mCnt(i) = CreateObject("Scripting.Dictionary")
        Set mEx(i) = CreateObject("Scripting.Dictionary")
        Set mShNames(i) = CreateObject("Scripting.Dictionary")
    Next i

    ' output goes to a LOCAL folder (C:\Users\<you>\AuditOut), never to a network drive
    outDir = Environ$("USERPROFILE") & "\AuditOut"
    mStage = "creating the output folder " & outDir
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    runName = Left$(CleanName(BaseName(newPath)), 40)
    mLog = outDir & "\audit_log.txt"
    LogLine "=== v3.1 run '" & runName & "' started " & Format$(Now, "dd/mm/yyyy hh:nn")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3          ' macros of the opened files stay disabled

    DumpWorkbook newPath, outDir, runName, "New", 0
    If oldPath <> "" Then DumpWorkbook oldPath, outDir, runName, "Old", 1

    WriteReport newPath, oldPath, Timer - t0, totF, totE, totFl

    mStage = "finished"
    LogLine "=== Finished in " & Format$((Timer - t0) / 60, "0.0") & " min"
    Application.ScreenUpdating = True
    MsgBox "Audit finished in " & Format$((Timer - t0) / 60, "0.0") & " min." & vbCrLf & vbCrLf & _
           "Formulas: " & totF & vbCrLf & _
           "Error cells: " & totE & vbCrLf & _
           "Suspicious cells: " & totFl & vbCrLf & vbCrLf & _
           "See the sheets Audit_Summary, Audit_Findings, Audit_Changes." & vbCrLf & _
           "CSV files and log: " & outDir, vbInformation
Done:
    Application.StatusBar = False
    Application.AutomationSecurity = sec
    Application.Calculation = calc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    LogLine "STOPPED at: " & mStage & " | " & Err.Description
    MsgBox "Audit stopped." & vbCrLf & vbCrLf & "Step: " & mStage & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation
    Resume Done
End Sub

Private Sub DumpWorkbook(ByVal filePath As String, ByVal outDir As String, ByVal runName As String, _
                         ByVal role As String, ByVal idx As Long)
    Dim wb As Workbook, ws As Worksheet
    Dim fFind As String, fPat As String, fStr As String
    Dim nF As Long, nE As Long, nFlag As Long, nPat As Long
    Dim k As Long, nSheets As Long, tS As Double, tRead As Double
    Dim res As String, info As String, msg As String, pre As String, vis As String
    Dim dCnt As Object, dEx As Object, key As Variant

    On Error GoTo CleanFail
    mStage = role & " file: opening " & filePath
    If IsOpen(filePath) Then Err.Raise vbObjectError + 1, , "Close this workbook first: " & filePath
    Set wb = Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    fFind = outDir & "\findings__" & runName & "__" & role & ".csv"
    fPat = outDir & "\patterns__" & runName & "__" & role & ".csv"
    fStr = outDir & "\structure__" & runName & "__" & role & ".csv"
    mStage = role & " file: creating the CSV files in " & outDir
    WriteText fFind, "Run,Role,File,Sheet,Address,Row,Col,Check,Formula", False
    WriteText fPat, "Run,Role,File,Sheet,Formula,Cells,ExampleCell", False
    WriteText fStr, "Run,Role,File,Kind,Name,Info,Formulas,Patterns,Errors,Flags", False

    nSheets = wb.Worksheets.Count
    For Each ws In wb.Worksheets
        k = k + 1
        nF = 0: nE = 0: nFlag = 0: nPat = 0: info = "": tRead = 0
        tS = Timer
        mStage = role & " file: sheet " & k & " of " & nSheets & " '" & ws.Name & "'"
        Application.StatusBar = "Audit - " & mStage
        LogLine "-> " & mStage

        Set dCnt = CreateObject("Scripting.Dictionary")
        Set dEx = CreateObject("Scripting.Dictionary")
        pre = Q(runName) & "," & Q(role) & "," & Q(wb.Name) & "," & Q(ws.Name) & ","
        vis = VisText(ws.Visible)

        ' --- read the sheet (findings go to the memory buffer)
        BufReset
        res = DumpSheet(ws, pre, role, dCnt, dEx, nF, nE, nFlag, info, tRead)
        If res <> "" Then info = "ERROR: " & res

        mStage = role & " file: saving the findings of sheet '" & ws.Name & "'"
        If mBufN > 0 Then WriteText fFind, BufText(), True

        ' --- one row per different formula
        BufReset
        For Each key In dCnt.Keys
            BufAdd pre & Q(CStr(key)) & "," & dCnt.Item(key) & "," & Q(CStr(dEx.Item(key)))
            mCnt(idx).Item(ws.Name & vbTab & CStr(key)) = dCnt.Item(key)
            mEx(idx).Item(ws.Name & vbTab & CStr(key)) = dEx.Item(key)
        Next key
        nPat = dCnt.Count
        mShNames(idx).Item(ws.Name) = 1
        mStage = role & " file: saving the formulas of sheet '" & ws.Name & "'"
        If mBufN > 0 Then WriteText fPat, BufText(), True

        mStage = role & " file: saving the summary of sheet '" & ws.Name & "'"
        WriteText fStr, Csv(runName, role, wb.Name, "Sheet", ws.Name, vis & " | " & info, _
                            nF, nPat, nE, nFlag), True
        mSheets.Add Array(role, ws.Name, vis, info, nF, nPat, nE, nFlag, Timer - tS)
        LogLine "   done in " & Format$(Timer - tS, "0.0") & " s (reading cells " & _
                Format$(tRead, "0.0") & " s) | " & info & " | formulas " & nF & _
                " | different " & nPat & " | errors " & nE & " | flags " & nFlag
    Next ws

    mStage = role & " file: named ranges and external links"
    BufReset
    CollectNames wb, runName, role
    CollectLinks wb, runName, role
    If mBufN > 0 Then WriteText fStr, BufText(), True

    mStage = role & " file: closing"
    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:
    msg = "Error " & Err.Number & ": " & Err.Description
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Err.Raise vbObjectError + 2, , msg
End Sub

' Returns "" when the sheet was read, or the error text when it could not be read.
Private Function DumpSheet(ws As Worksheet, ByVal pre As String, ByVal role As String, _
                           dCnt As Object, dEx As Object, _
                           nF As Long, nE As Long, nFlag As Long, _
                           info As String, tRead As Double) As String
    Dim ur As Range, rng As Range, blk As Range
    Dim urR As Long, urC As Long, nR As Long, nC As Long
    Dim chunk As Long, startR As Long, rowsNow As Long
    Dim aF As Variant, aV As Variant, hf As Variant
    Dim i As Long, j As Long, sheetRow As Long, t1 As Double
    Dim fx As String, er As String, flag As String, chk As String
    Dim hasFx As Boolean, doCell As Boolean

    On Error GoTo SheetFail

    Set ur = ws.UsedRange
    urR = ur.Row + ur.Rows.Count - 1
    urC = ur.Column + ur.Columns.Count - 1
    info = "empty"
    If Application.WorksheetFunction.CountA(ur) = 0 Then Exit Function

    ' real last row and column (the used range is often much larger than the data)
    nR = LastFilled(ws, urR, urC, True)
    nC = LastFilled(ws, nR, urC, False)
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(nR, nC))
    info = rng.Address(False, False)
    If urR > nR + 100 Or urC > nC + 20 Then
        info = info & " (used range extends to " & Addr(urR, urC) & ")"
    End If

    chunk = MAX_CELLS \ nC
    If chunk < 1 Then chunk = 1
    startR = 1

    Do While startR <= nR
        rowsNow = chunk
        If startR + rowsNow - 1 > nR Then rowsNow = nR - startR + 1
        Set blk = rng.Rows(startR).Resize(rowsNow)

        If Application.WorksheetFunction.CountA(blk) > 0 Then
            hf = blk.HasFormula                 ' True = all, False = none, Null = some
            If IsNull(hf) Then hasFx = True Else hasFx = CBool(hf)
            t1 = Timer
            aV = To2D(blk.Value2)
            If hasFx Then aF = To2D(blk.FormulaR1C1)   ' R1C1 = same text for a copied formula
            tRead = tRead + (Timer - t1)

            If Not hasFx Then
                ' no formulas in this block: only look for typed error values
                For i = 1 To rowsNow
                    For j = 1 To nC
                        If IsError(aV(i, j)) Then
                            nE = nE + 1
                            AddFinding pre, role, ws.Name, startR + i - 1, j, _
                                       "Error value " & ErrText(aV(i, j)), ""
                        End If
                    Next j
                Next i
            Else
                For i = 1 To rowsNow
                    sheetRow = startR + i - 1
                    For j = 1 To nC
                        doCell = Not IsEmpty(aV(i, j))
                        If Not doCell Then
                            If VarType(aF(i, j)) = vbString Then doCell = (Len(aF(i, j)) > 0)
                        End If

                        If doCell Then
                            fx = "": er = "": flag = ""
                            If IsFx(aF(i, j)) Then
                                fx = aF(i, j)
                                nF = nF + 1
                                If dCnt.Exists(fx) Then
                                    dCnt.Item(fx) = dCnt.Item(fx) + 1
                                Else
                                    dCnt.Item(fx) = 1
                                    dEx.Item(fx) = Addr(sheetRow, j)
                                End If
                                If InStr(fx, "#REF!") > 0 Then flag = "Broken reference in formula"
                                If j > 1 And j < nC Then
                                    If IsFx(aF(i, j - 1)) And IsFx(aF(i, j + 1)) Then
                                        If aF(i, j - 1) = aF(i, j + 1) And aF(i, j - 1) <> fx Then
                                            flag = AddFlag(flag, "Formula differs from both neighbours")
                                        End If
                                    End If
                                End If
                            ElseIf j > 1 And j < nC Then
                                If IsNum(aV(i, j)) Then
                                    If IsFx(aF(i, j - 1)) And IsFx(aF(i, j + 1)) Then
                                        If aF(i, j - 1) = aF(i, j + 1) Then
                                            flag = "Typed number between identical formulas"
                                        End If
                                    End If
                                End If
                            End If

                            If IsError(aV(i, j)) Then
                                er = ErrText(aV(i, j))
                                nE = nE + 1
                            End If

                            If er <> "" Or flag <> "" Then
                                If flag <> "" Then nFlag = nFlag + 1
                                If er <> "" And flag <> "" Then
                                    chk = "Error value " & er & "; " & flag
                                ElseIf er <> "" Then
                                    chk = "Error value " & er
                                Else
                                    chk = flag
                                End If
                                AddFinding pre, role, ws.Name, sheetRow, j, chk, fx
                            End If
                        End If
                    Next j
                Next i
            End If
        End If
        startR = startR + rowsNow
    Loop
    Exit Function

SheetFail:
    DumpSheet = "Error " & Err.Number & ": " & Err.Description
End Function

Private Sub AddFinding(ByVal pre As String, ByVal role As String, ByVal sheetName As String, _
                       ByVal r As Long, ByVal c As Long, ByVal chk As String, ByVal fx As String)
    Dim a As String, k As String
    a = Addr(r, c)
    BufAdd pre & Q(a) & "," & r & "," & c & "," & Q(chk) & "," & Q(fx)
    If role = "New" Then
        k = sheetName & vbTab & chk
        If mChk.Exists(k) Then mChk.Item(k) = mChk.Item(k) + 1 Else mChk.Item(k) = 1
        If mFinds.Count < MAX_LIST Then mFinds.Add Array(sheetName, a, chk, fx)
    End If
End Sub

' Last row (byRow = True) or last column that contains anything, inside rows 1..nR, columns 1..nC.
Private Function LastFilled(ws As Worksheet, ByVal nR As Long, ByVal nC As Long, _
                            ByVal byRow As Boolean) As Long
    Dim lo As Long, hi As Long, m As Long, r As Range
    lo = 1
    If byRow Then hi = nR Else hi = nC
    Do While lo < hi
        m = (lo + hi) \ 2
        If byRow Then
            Set r = ws.Range(ws.Cells(m + 1, 1), ws.Cells(hi, nC))
        Else
            Set r = ws.Range(ws.Cells(1, m + 1), ws.Cells(nR, hi))
        End If
        If Application.WorksheetFunction.CountA(r) > 0 Then lo = m + 1 Else hi = m
    Loop
    LastFilled = lo
End Function

' ---------- report sheets inside AuditTool ----------

Private Sub WriteReport(ByVal newPath As String, ByVal oldPath As String, ByVal secs As Double, _
                        totF As Long, totE As Long, totFl As Long)
    Dim ws As Worksheet, arr() As Variant, it As Variant, parts() As String, key As Variant
    Dim i As Long, j As Long, n As Long, r As Long, totP As Long, nNew As Long
    Dim chg As Collection, sh As String, cOld As Long, cNew As Long, note As String

    ' ----- Audit_Summary
    mStage = "report: sheet Audit_Summary"
    Set ws = ReportSheet("Audit_Summary")
    n = mSheets.Count
    ReDim arr(1 To n + 1, 1 To 9)
    arr(1, 1) = "File": arr(1, 2) = "Sheet": arr(1, 3) = "Visibility": arr(1, 4) = "Range read"
    arr(1, 5) = "Formulas": arr(1, 6) = "Different formulas": arr(1, 7) = "Error cells"
    arr(1, 8) = "Suspicious cells": arr(1, 9) = "Seconds"
    i = 1
    For Each it In mSheets
        i = i + 1
        arr(i, 1) = it(0): arr(i, 2) = T(CStr(it(1))): arr(i, 3) = it(2): arr(i, 4) = T(CStr(it(3)))
        arr(i, 5) = it(4): arr(i, 6) = it(5): arr(i, 7) = it(6): arr(i, 8) = it(7)
        arr(i, 9) = Round(it(8), 1)
        If it(0) = "New" Then
            nNew = nNew + 1
            totF = totF + it(4): totP = totP + it(5): totE = totE + it(6): totFl = totFl + it(7)
        End If
    Next it
    ws.Range("A1").Value = "FORMULA AUDIT"
    ws.Range("A2").Value = "New file":            ws.Range("B2").Value = T(newPath)
    ws.Range("A3").Value = "Old file":            ws.Range("B3").Value = T(IIf(oldPath = "", "(not compared)", oldPath))
    ws.Range("A4").Value = "Run on":              ws.Range("B4").Value = "'" & Format$(Now, "dd/mm/yyyy hh:nn")
    ws.Range("A5").Value = "Duration (min)":      ws.Range("B5").Value = Round(secs / 60, 1)
    ws.Range("A6").Value = "Sheets (new file)":   ws.Range("B6").Value = nNew
    ws.Range("A7").Value = "Formulas":            ws.Range("B7").Value = totF
    ws.Range("A8").Value = "Different formulas":  ws.Range("B8").Value = totP
    ws.Range("A9").Value = "Error cells":         ws.Range("B9").Value = totE
    ws.Range("A10").Value = "Suspicious cells":   ws.Range("B10").Value = totFl
    ws.Range("A12").Resize(n + 1, 9).Value = arr
    ws.Range("A1").Font.Bold = True
    ws.Range("A12:I12").Font.Bold = True
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 30
    ws.Columns("E:I").ColumnWidth = 16
    ws.Range("B2:B10").HorizontalAlignment = xlLeft

    ' ----- Audit_Findings
    mStage = "report: sheet Audit_Findings"
    Set ws = ReportSheet("Audit_Findings")
    ws.Range("A1").Value = "ERROR CELLS AND SUSPICIOUS CELLS IN THE NEW FILE"
    ws.Range("A1").Font.Bold = True
    n = mChk.Count
    If n = 0 Then
        ws.Range("A3").Value = "Nothing found."
    Else
        ReDim arr(1 To n + 1, 1 To 3)
        arr(1, 1) = "Sheet": arr(1, 2) = "Check": arr(1, 3) = "Cells"
        i = 1
        For Each key In mChk.Keys
            i = i + 1
            parts = Split(CStr(key), vbTab, 2)
            arr(i, 1) = T(parts(0)): arr(i, 2) = parts(1): arr(i, 3) = mChk.Item(key)
        Next key
        ws.Range("A3").Resize(n + 1, 3).Value = arr
        ws.Range("A3:C3").Font.Bold = True

        r = n + 6
        note = "CELL LIST"
        If mFinds.Count >= MAX_LIST Then
            note = note & " (first " & MAX_LIST & " cells only; the full list is in the findings CSV)"
        End If
        ws.Cells(r - 1, 1).Value = note
        ws.Cells(r - 1, 1).Font.Bold = True
        n = mFinds.Count
        ReDim arr(1 To n + 1, 1 To 4)
        arr(1, 1) = "Sheet": arr(1, 2) = "Cell": arr(1, 3) = "Check": arr(1, 4) = "Formula"
        i = 1
        For Each it In mFinds
            i = i + 1
            arr(i, 1) = T(CStr(it(0))): arr(i, 2) = "'" & it(1)
            arr(i, 3) = it(2): arr(i, 4) = T(CStr(it(3)))
        Next it
        ws.Cells(r, 1).Resize(n + 1, 4).Value = arr
        ws.Cells(r, 1).Resize(1, 4).Font.Bold = True
    End If
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 80

    ' ----- Audit_Changes
    mStage = "report: sheet Audit_Changes"
    Set ws = ReportSheet("Audit_Changes")
    ws.Range("A1").Value = "FORMULA CHANGES: OLD FILE vs NEW FILE"
    ws.Range("A1").Font.Bold = True
    If oldPath = "" Then
        ws.Range("A3").Value = "No old file was selected, so nothing was compared."
    Else
        Set chg = New Collection
        For Each key In mShNames(0).Keys
            If Not mShNames(1).Exists(key) Then chg.Add Array(CStr(key), "Sheet added", "", 0, 0, "")
        Next key
        For Each key In mShNames(1).Keys
            If Not mShNames(0).Exists(key) Then chg.Add Array(CStr(key), "Sheet removed", "", 0, 0, "")
        Next key
        For Each key In mCnt(0).Keys
            parts = Split(CStr(key), vbTab, 2)
            sh = parts(0)
            If mShNames(1).Exists(sh) And chg.Count < MAX_LIST Then
                cNew = mCnt(0).Item(key)
                If mCnt(1).Exists(key) Then cOld = mCnt(1).Item(key) Else cOld = 0
                If cOld = 0 Then
                    chg.Add Array(sh, "New formula (not in the old file)", parts(1), cOld, cNew, mEx(0).Item(key))
                ElseIf cNew < cOld Then
                    chg.Add Array(sh, "Used in fewer cells", parts(1), cOld, cNew, mEx(0).Item(key))
                ElseIf cNew > cOld Then
                    chg.Add Array(sh, "Used in more cells", parts(1), cOld, cNew, mEx(0).Item(key))
                End If
            End If
        Next key
        For Each key In mCnt(1).Keys
            parts = Split(CStr(key), vbTab, 2)
            sh = parts(0)
            If mShNames(0).Exists(sh) And chg.Count < MAX_LIST Then
                If Not mCnt(0).Exists(key) Then
                    chg.Add Array(sh, "Formula no longer used", parts(1), mCnt(1).Item(key), 0, mEx(1).Item(key))
                End If
            End If
        Next key

        n = chg.Count
        If n = 0 Then
            ws.Range("A3").Value = "No differences: same sheets and same formulas in both files."
        Else
            If n >= MAX_LIST Then ws.Range("A2").Value = "Only the first " & MAX_LIST & " changes are listed."
            ReDim arr(1 To n + 1, 1 To 6)
            arr(1, 1) = "Sheet": arr(1, 2) = "Change": arr(1, 3) = "Formula"
            arr(1, 4) = "Cells in old file": arr(1, 5) = "Cells in new file": arr(1, 6) = "Example cell"
            i = 1
            For Each it In chg
                i = i + 1
                arr(i, 1) = T(CStr(it(0))): arr(i, 2) = it(1): arr(i, 3) = T(CStr(it(2)))
                arr(i, 4) = it(3): arr(i, 5) = it(4): arr(i, 6) = "'" & it(5)
            Next it
            ws.Range("A3").Resize(n + 1, 6).Value = arr
            ws.Range("A3:F3").Font.Bold = True
        End If
    End If
    ws.Columns("A").ColumnWidth = 34
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 80
    ws.Columns("D:F").ColumnWidth = 18

    ThisWorkbook.Worksheets("Audit_Summary").Activate
End Sub

Private Function ReportSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    Else
        ws.Cells.Clear
    End If
    Set ReportSheet = ws
End Function

' Text that Excel must not read as a formula, number or date.
Private Function T(ByVal s As String) As String
    If Len(s) > 1000 Then s = Left$(s, 1000) & " ..."
    Select Case Left$(s, 1)
        Case "=", "+", "-", "@", "'": T = "'" & s
        Case Else: T = s
    End Select
End Function

' ---------- output: memory buffer + one short write ----------

Private Sub BufReset()
    ReDim mBuf(0 To 255)
    mBufN = 0
End Sub

Private Sub BufAdd(ByVal s As String)
    If mBufN > UBound(mBuf) Then ReDim Preserve mBuf(0 To mBufN * 2 + 1)
    mBuf(mBufN) = s
    mBufN = mBufN + 1
End Sub

Private Function BufText() As String
    ReDim Preserve mBuf(0 To mBufN - 1)
    BufText = Join(mBuf, vbCrLf)
End Function

Private Sub WriteText(ByVal filePath As String, ByVal content As String, ByVal addToEnd As Boolean)
    Dim attempt As Long, e As String
    For attempt = 1 To 3
        e = TryWrite(filePath, content, addToEnd)
        If e = "" Then Exit Sub
        LogLine "   write failed (attempt " & attempt & "): " & e & " | " & filePath
        Application.Wait Now + TimeSerial(0, 0, 2)
    Next attempt
    Err.Raise vbObjectError + 3, , e & " | file: " & filePath
End Sub

Private Function TryWrite(ByVal filePath As String, ByVal content As String, _
                          ByVal addToEnd As Boolean) As String
    Dim f As Integer
    On Error GoTo Bad
    f = FreeFile
    If addToEnd Then
        Open filePath For Append As #f
    Else
        Open filePath For Output As #f
    End If
    Print #f, content
    Close #f
    Exit Function
Bad:
    TryWrite = "Error " & Err.Number & ": " & Err.Description
    Resume CloseIt
CloseIt:
    On Error Resume Next
    Close #f
End Function

Private Sub CollectNames(wb As Workbook, ByVal runName As String, ByVal role As String)
    Dim nm As Name
    On Error Resume Next
    For Each nm In wb.Names
        BufAdd Csv(runName, role, wb.Name, "Name", nm.Name, SafeRefersTo(nm), 0, 0, 0, 0)
    Next nm
End Sub

Private Sub CollectLinks(wb As Workbook, ByVal runName As String, ByVal role As String)
    Dim lnk As Variant, i As Long
    On Error Resume Next
    lnk = wb.LinkSources(xlExcelLinks)
    If IsArray(lnk) Then
        For i = LBound(lnk) To UBound(lnk)
            BufAdd Csv(runName, role, wb.Name, "Link", CStr(lnk(i)), "", 0, 0, 0, 0)
        Next i
    End If
End Sub

Private Sub LogLine(ByVal s As String)
    Dim f As Integer
    On Error Resume Next
    If mLog = "" Then Exit Sub
    f = FreeFile
    Open mLog For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & "  " & s
    Close #f
End Sub

' ---------- helpers ----------

Private Function PickFile(ByVal title As String) As String
    Dim v As Variant
    v = Application.GetOpenFilename("Excel files (*.xls*), *.xls*", , title)
    If VarType(v) = vbBoolean Then PickFile = "" Else PickFile = CStr(v)
End Function

Private Function IsOpen(ByVal p As String) As Boolean
    Dim wb As Workbook, n As String
    n = Mid$(p, InStrRev(p, "\") + 1)
    For Each wb In Workbooks
        If StrComp(wb.Name, n, vbTextCompare) = 0 Then IsOpen = True
    Next wb
End Function

Private Function To2D(v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If IsArray(v) Then
        To2D = v
    Else
        a(1, 1) = v
        To2D = a
    End If
End Function

Private Function IsFx(v As Variant) As Boolean
    If VarType(v) = vbString Then IsFx = (Left$(v, 1) = "=")
End Function

Private Function IsNum(v As Variant) As Boolean
    IsNum = (VarType(v) = vbDouble)
End Function

Private Function AddFlag(ByVal existing As String, ByVal extra As String) As String
    If existing = "" Then AddFlag = extra Else AddFlag = existing & "; " & extra
End Function

Private Function ErrText(v As Variant) As String
    Dim s As String, n As Long
    s = CStr(v)
    n = CLng(Val(Mid$(s, InStrRev(s, " ") + 1)))
    Select Case n
        Case 2000: ErrText = "#NULL!"
        Case 2007: ErrText = "#DIV/0!"
        Case 2015: ErrText = "#VALUE!"
        Case 2023: ErrText = "#REF!"
        Case 2029: ErrText = "#NAME?"
        Case 2036: ErrText = "#NUM!"
        Case 2042: ErrText = "#N/A"
        Case Else: ErrText = "#ERROR " & n
    End Select
End Function

Private Function SafeRefersTo(nm As Name) As String
    On Error Resume Next
    SafeRefersTo = nm.RefersTo
End Function

Private Function VisText(ByVal v As Long) As String
    Select Case v
        Case -1: VisText = "Visible"
        Case 0: VisText = "Hidden"
        Case Else: VisText = "Very hidden"
    End Select
End Function

Private Function Addr(ByVal r As Long, ByVal c As Long) As String
    Dim s As String
    Do While c > 0
        s = Chr$(65 + (c - 1) Mod 26) & s
        c = (c - 1) \ 26
    Loop
    Addr = s & r
End Function

Private Function BaseName(ByVal p As String) As String
    Dim s As String
    s = Mid$(p, InStrRev(p, "\") + 1)
    If InStrRev(s, ".") > 1 Then s = Left$(s, InStrRev(s, ".") - 1)
    BaseName = s
End Function

Private Function CleanName(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9 ._-]" Then r = r & ch Else r = r & "_"
    Next i
    CleanName = r
End Function

Private Function Q(ByVal s As String) As String
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    Q = """" & Replace(s, """", """""") & """"
End Function

Private Function Csv(ParamArray parts() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & ","
        s = s & Q(CStr(parts(i)))
    Next i
    Csv = s
End Function
