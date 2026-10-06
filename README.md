
' =====================================================================
'  EXCEL AUDIT DUMP  v2  (VBA)
'  Same output as v1 (the Power BI queries do not change), but:
'    - skips the empty area of each sheet (finds the real last row/column)
'    - skips blocks that contain no formulas
'    - one problem sheet no longer stops the whole audit
'    - if it stops, the message names the exact step
'    - writes a log: AuditOut\audit_log.txt (one line per sheet, with seconds)
'
'  INSTALL: open AuditTool.xlsm > Alt+F11 > Modulo1 > Ctrl+A > paste this > Ctrl+S
'  RUN    : Alt+F8 > RunAudit > Esegui
' =====================================================================

Option Explicit

Private Const MAX_CELLS As Long = 200000     ' cells read per block
Private mStage As String
Private mLog As String

Public Sub RunAudit()
    Dim newPath As String, oldPath As String, outDir As String, runName As String
    Dim calc As XlCalculation, sec As Long, t0 As Double

    If ThisWorkbook.Path = "" Then
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

    outDir = ThisWorkbook.Path & "\AuditOut"
    mStage = "creating the output folder " & outDir
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    runName = CleanName(BaseName(newPath))
    mLog = outDir & "\audit_log.txt"
    LogLine "=== Run '" & runName & "' started " & Format$(Now, "dd/mm/yyyy hh:nn")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3          ' macros of the opened files stay disabled

    DumpWorkbook newPath, outDir, runName, "New"
    If oldPath <> "" Then DumpWorkbook oldPath, outDir, runName, "Old"

    mStage = "finished"
    LogLine "=== Finished in " & Format$((Timer - t0) / 60, "0.0") & " min"
    MsgBox "Audit finished in " & Format$((Timer - t0) / 60, "0.0") & " min." & vbCrLf & _
           "Files written to:" & vbCrLf & outDir, vbInformation
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

Private Sub DumpWorkbook(ByVal path As String, ByVal outDir As String, _
                         ByVal runName As String, ByVal role As String)
    Dim wb As Workbook, ws As Worksheet
    Dim fC As Integer, fS As Integer
    Dim nF As Long, nE As Long, nFlag As Long
    Dim k As Long, nSheets As Long, tS As Double
    Dim res As String, info As String, msg As String

    On Error GoTo CleanFail
    mStage = role & " file: opening " & path
    If IsOpen(path) Then Err.Raise vbObjectError + 1, , "Close this workbook first: " & path
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    mStage = role & " file: creating the CSV files in " & outDir
    fC = FreeFile
    Open outDir & "\cells__" & runName & "__" & role & ".csv" For Output As #fC
    Print #fC, "Run,Role,File,Sheet,Address,Row,Col,Formula,Error,Flag"
    fS = FreeFile
    Open outDir & "\structure__" & runName & "__" & role & ".csv" For Output As #fS
    Print #fS, "Run,Role,File,Kind,Name,Info,Formulas,Errors,Flags"

    nSheets = wb.Worksheets.Count
    For Each ws In wb.Worksheets
        k = k + 1
        nF = 0: nE = 0: nFlag = 0: info = ""
        tS = Timer
        mStage = role & " file: sheet " & k & " of " & nSheets & " '" & ws.Name & "'"
        Application.StatusBar = "Audit - " & mStage
        LogLine "-> " & mStage

        res = DumpSheet(ws, fC, runName, role, wb.Name, nF, nE, nFlag, info)
        If res <> "" Then info = "ERROR: " & res

        mStage = role & " file: writing the summary of sheet '" & ws.Name & "'"
        Print #fS, Csv(runName, role, wb.Name, "Sheet", ws.Name, _
                       VisText(ws.Visible) & " | " & info, nF, nE, nFlag)
        LogLine "   done in " & Format$(Timer - tS, "0.0") & " s | " & info & _
                " | formulas " & nF & " | errors " & nE & " | flags " & nFlag
    Next ws

    mStage = role & " file: named ranges"
    DumpNames wb, fS, runName, role
    mStage = role & " file: external links"
    DumpLinks wb, fS, runName, role

    mStage = role & " file: closing"
    Close #fC
    Close #fS
    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:
    msg = "Error " & Err.Number & ": " & Err.Description
    Close
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Err.Raise vbObjectError + 2, , msg
End Sub

' Returns "" when the sheet was read, or the error text when it could not be read.
Private Function DumpSheet(ws As Worksheet, ByVal f As Integer, ByVal runName As String, _
                           ByVal role As String, ByVal fileName As String, _
                           nF As Long, nE As Long, nFlag As Long, info As String) As String
    Dim ur As Range, rng As Range, blk As Range
    Dim urR As Long, urC As Long, nR As Long, nC As Long
    Dim chunk As Long, startR As Long, rowsNow As Long
    Dim aF As Variant, aV As Variant, hf As Variant
    Dim i As Long, j As Long, sheetRow As Long
    Dim fx As String, er As String, flag As String, pre As String
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

    pre = Q(runName) & "," & Q(role) & "," & Q(fileName) & "," & Q(ws.Name) & ","
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
            aV = To2D(blk.Value2)

            If Not hasFx Then
                ' no formulas in this block: only look for typed error values
                For i = 1 To rowsNow
                    For j = 1 To nC
                        If IsError(aV(i, j)) Then
                            nE = nE + 1
                            WriteCell f, pre, startR + i - 1, j, "", ErrText(aV(i, j)), ""
                        End If
                    Next j
                Next i
            Else
                aF = To2D(blk.FormulaR1C1)      ' R1C1 = same text for a copied formula
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

                            If fx <> "" Or er <> "" Or flag <> "" Then
                                If flag <> "" Then nFlag = nFlag + 1
                                WriteCell f, pre, sheetRow, j, fx, er, flag
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

' Last row (byRow = True) or last column that contains anything, inside rows 1..nR, columns 1..nC.
' Binary search with COUNTA: it also sees hidden and filtered cells.
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

Private Sub WriteCell(ByVal f As Integer, ByVal pre As String, ByVal r As Long, ByVal c As Long, _
                      ByVal fx As String, ByVal er As String, ByVal flag As String)
    Print #f, pre & Q(Addr(r, c)) & "," & r & "," & c & "," & Q(fx) & "," & Q(er) & "," & Q(flag)
End Sub

Private Sub DumpNames(wb As Workbook, ByVal f As Integer, ByVal runName As String, ByVal role As String)
    Dim nm As Name
    On Error Resume Next
    For Each nm In wb.Names
        Print #f, Csv(runName, role, wb.Name, "Name", nm.Name, SafeRefersTo(nm), 0, 0, 0)
    Next nm
End Sub

Private Sub DumpLinks(wb As Workbook, ByVal f As Integer, ByVal runName As String, ByVal role As String)
    Dim links As Variant, i As Long
    On Error Resume Next
    links = wb.LinkSources(xlExcelLinks)
    If IsArray(links) Then
        For i = LBound(links) To UBound(links)
            Print #f, Csv(runName, role, wb.Name, "Link", CStr(links(i)), "", 0, 0, 0)
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
