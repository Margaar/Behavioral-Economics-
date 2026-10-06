' =====================================================================
'  EXCEL AUDIT DUMP  (VBA)
'  Lists every formula, every error value and every suspicious cell of a
'  workbook, plus its sheets, named ranges and external links.
'
'  SETUP (once)
'    1. Create a NEW empty workbook and save it as AuditTool.xlsm
'       (do NOT put this code in the Capital Deep Dive template).
'    2. Alt+F11 > Inserisci > Modulo > paste everything below.
'    3. Save.
'
'  USE
'    Alt+F8 > RunAudit > Esegui
'      - first dialog : pick the NEW (current) workbook
'      - second dialog: pick the OLD (previous) workbook, or Annulla to skip
'    The workbooks are opened read-only, with their own macros disabled,
'    and are never changed. Output goes to a folder "AuditOut" next to
'    AuditTool.xlsm:
'      cells__<name>__New.csv / __Old.csv        formulas, errors, flags
'      structure__<name>__New.csv / __Old.csv    sheets, names, links
'    Power BI reads that folder (see PowerQuery_Audit.txt).
' =====================================================================

Option Explicit

Private Const MAX_CELLS As Long = 500000     ' cells read per block

Public Sub RunAudit()
    Dim newPath As String, oldPath As String, outDir As String, runName As String
    Dim calc As XlCalculation, sec As Long

    If ThisWorkbook.Path = "" Then
        MsgBox "Save AuditTool.xlsm first.", vbExclamation
        Exit Sub
    End If

    newPath = PickFile("Select the NEW (current) workbook")
    If newPath = "" Then Exit Sub
    oldPath = PickFile("Select the OLD (previous) workbook - Cancel to skip the comparison")

    outDir = ThisWorkbook.Path & "\AuditOut"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    runName = CleanName(BaseName(newPath))

    calc = Application.Calculation
    sec = Application.AutomationSecurity
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3          ' macros of the opened files stay disabled

    DumpWorkbook newPath, outDir, runName, "New"
    If oldPath <> "" Then DumpWorkbook oldPath, outDir, runName, "Old"

    MsgBox "Audit files written to:" & vbCrLf & outDir, vbInformation
Done:
    Application.AutomationSecurity = sec
    Application.Calculation = calc
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Audit stopped: " & Err.Description, vbExclamation
    Resume Done
End Sub

Private Sub DumpWorkbook(ByVal path As String, ByVal outDir As String, _
                         ByVal runName As String, ByVal role As String)
    Dim wb As Workbook, ws As Worksheet, nm As Name
    Dim fC As Integer, fS As Integer
    Dim nF As Long, nE As Long, nFlag As Long
    Dim links As Variant, i As Long, msg As String

    On Error GoTo CleanFail
    If IsOpen(path) Then Err.Raise vbObjectError + 1, , "Close this workbook first: " & path

    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)

    fC = FreeFile
    Open outDir & "\cells__" & runName & "__" & role & ".csv" For Output As #fC
    Print #fC, "Run,Role,File,Sheet,Address,Row,Col,Formula,Error,Flag"
    fS = FreeFile
    Open outDir & "\structure__" & runName & "__" & role & ".csv" For Output As #fS
    Print #fS, "Run,Role,File,Kind,Name,Info,Formulas,Errors,Flags"

    For Each ws In wb.Worksheets
        nF = 0: nE = 0: nFlag = 0
        DumpSheet ws, fC, runName, role, wb.Name, nF, nE, nFlag
        Print #fS, Csv(runName, role, wb.Name, "Sheet", ws.Name, _
                       VisText(ws.Visible) & " | " & ws.UsedRange.Address(False, False), nF, nE, nFlag)
    Next ws

    For Each nm In wb.Names
        Print #fS, Csv(runName, role, wb.Name, "Name", nm.Name, SafeRefersTo(nm), 0, 0, 0)
    Next nm

    links = wb.LinkSources(xlExcelLinks)
    If Not IsEmpty(links) Then
        For i = LBound(links) To UBound(links)
            Print #fS, Csv(runName, role, wb.Name, "Link", CStr(links(i)), "", 0, 0, 0)
        Next i
    End If

    Close #fC
    Close #fS
    wb.Close SaveChanges:=False
    Exit Sub

CleanFail:
    msg = Err.Description
    Close
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Err.Raise vbObjectError + 2, , msg
End Sub

Private Sub DumpSheet(ws As Worksheet, ByVal f As Integer, ByVal runName As String, _
                      ByVal role As String, ByVal fileName As String, _
                      nF As Long, nE As Long, nFlag As Long)
    Dim ur As Range, nR As Long, nC As Long, r0 As Long, c0 As Long
    Dim chunk As Long, startR As Long, rowsNow As Long
    Dim aF As Variant, aV As Variant, i As Long, j As Long
    Dim fx As String, er As String, flag As String, pre As String
    Dim sheetRow As Long

    Set ur = ws.UsedRange
    If Application.WorksheetFunction.CountA(ur) = 0 Then Exit Sub
    nR = ur.Rows.Count: nC = ur.Columns.Count
    r0 = ur.Row: c0 = ur.Column
    pre = Q(runName) & "," & Q(role) & "," & Q(fileName) & "," & Q(ws.Name) & ","

    chunk = MAX_CELLS \ nC
    If chunk < 1 Then chunk = 1
    startR = 1

    Do While startR <= nR
        rowsNow = chunk
        If startR + rowsNow - 1 > nR Then rowsNow = nR - startR + 1
        With ur.Rows(startR).Resize(rowsNow)
            aF = To2D(.FormulaR1C1)      ' R1C1 = same text for a copied formula, wherever it sits
            aV = To2D(.Value2)
        End With

        For i = 1 To rowsNow
            sheetRow = r0 + startR + i - 2
            For j = 1 To nC
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
                            If aF(i, j - 1) = aF(i, j + 1) Then flag = "Typed number between identical formulas"
                        End If
                    End If
                End If

                If IsError(aV(i, j)) Then
                    er = ErrText(aV(i, j))
                    nE = nE + 1
                End If

                If fx <> "" Or er <> "" Or flag <> "" Then
                    If flag <> "" Then nFlag = nFlag + 1
                    Print #f, pre & Q(Addr(sheetRow, c0 + j - 1)) & "," & sheetRow & "," & _
                              (c0 + j - 1) & "," & Q(fx) & "," & Q(er) & "," & Q(flag)
                End If
            Next j
        Next i
        startR = startR + rowsNow
    Loop
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
