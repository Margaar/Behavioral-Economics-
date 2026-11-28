=====================================================================
 C)
=====================================================================

ИДЕЯ
- Сейчас: CDD.xlsm  -> fnReadSheet + BlockConfig -> FactData -> страницы.
- Будет:  CP-Tool (по файлу на компанию) -> fnCPRead -> CP_Raw (длинная таблица)
          -> CPMap (словарь: строка CP-Tool -> Topic/Block/Item CDD) -> FactData_CP
          -> FactData = CDD / CP-Tool / оба (переключатель DataSource).
- fnCPRead работает как fnReadSheet: ищет таблицы по заголовкам периодов
  (Actual 2025, Forecast 2026, Plan 2027..., Forecast 2026 SII Review),
  а не по адресам ячеек. Читает и скрытые листы.
- Power Query видит ТОЛЬКО значения, формулы он не читает. Если нужно
  "по формулам" — см. блок 10 (VBA-выгрузка формул в таблицу-карту).

ПЕРЕД НАЧАЛОМ
- Сделай копию .pbix.
- Код в Power BI я не запускал: идём по шагам, после каждого смотри результат.


#####################################################################
 1. ПАРАМЕТРЫ (Home > Gestisci parametri > Nuovo)
#####################################################################
  SiteUrl     Testo   https://generali.sharepoint.com/sites/GCSC
  FileFilter  Testo   CP-Tool 27-29
  DataSource  Testo   CP-Tool        (варианты: CP-Tool / CDD / Both)


#####################################################################
 2. CPSheets  (Carica: OFF) — какие листы CP-Tool читать
#####################################################################
>> Включай листы по мере маппинга. Чем меньше листов, тем быстрее обновление.

let
    Src = #table(type table [CPSheet = text, Include = logical], {
        {"Free Tangible Capital analysis",  true},
        {"LGAAP TCAP analysis",             true},
        {"SII - Solvency analysis",         true},
        {"SII - BS analysis",               true},
        {"Remittance & CapNeeds analysis",  true},
        {"Capital Movement analysis",       true},
        {"Sensitivities analysis",          true},
        {"ORMT analysis",                   true},
        {"Local Solvency analysis",         false}
    })
in
    Src


#####################################################################
 3. fnCPRead  (функция) — универсальный парсер листа CP-Tool
#####################################################################
>> Query vuota > Editor avanzato > вставить > имя fnCPRead

(wb as binary, fileName as text, optional sheetFilter as nullable list) as table =>
let
    Src     = Excel.Workbook(wb, null, true),
    Sheets0 = Table.SelectRows(Src, each [Kind] = "Sheet"),
    Sheets  = if sheetFilter = null then Sheets0
              else Table.SelectRows(Sheets0, each List.Contains(sheetFilter, [Name])),

    // --- код компании XX999 ---
    IsCode  = (t as text) as logical =>
                Text.Length(t) = 5
                and List.AllTrue(List.Transform(Text.ToList(Text.Start(t, 2)), each _ >= "A" and _ <= "Z"))
                and List.AllTrue(List.Transform(Text.ToList(Text.End(t, 3)),   each _ >= "0" and _ <= "9")),

    // --- распознать заголовок периода ---
    ParsePeriod = (v as any) as nullable record =>
        let
            s   = if v is text then Text.Trim(v)
                  else if v is number and v >= 2000 and v <= 2100 then Text.From(v) else null,
            w   = if s = null then {} else Text.Split(s, " "),
            yr  = List.First(List.Select(List.Transform(w, each try Number.From(_) otherwise null),
                                         each _ <> null and _ >= 2000 and _ <= 2100), null),
            typ = if s = null or yr = null then null
                  else if Text.StartsWith(s, "Actual")   then "Actual"
                  else if Text.StartsWith(s, "Forecast") then "Forecast"
                  else if Text.StartsWith(s, "Plan")     then "Plan"
                  else if s = Text.From(yr)              then "Year"
                  else null
        in
            if typ = null then null
            else [Label = s, Type = typ, Year = yr,
                  Scenario = if Text.Contains(s, "Review") then "SII Review" else "Base"],

    // --- код и имя компании: первые 6 строк любого листа ---
    FirstTxt = List.Combine(List.Transform(Sheets[Data], (t) =>
                 List.Combine(List.Transform(List.FirstN(Table.ToRows(t), 6),
                     (r) => List.Select(r, each _ is text))))),
    Hit      = List.First(List.Select(FirstTxt, each IsCode(Text.Trim(Text.BeforeDelimiter(_, " - ")))), null),
    Entity   = if Hit = null then fileName else Text.Trim(Text.BeforeDelimiter(Hit, " - ")),
    EntName  = if Hit = null then null else Text.Trim(Text.AfterDelimiter(Hit, " - ")),

    // --- чтение одного листа ---
    ReadOne = (tbl as table, sh as text) as list =>
        let
            Raw   = Table.ToRows(tbl),
            N     = List.Count(Raw),
            Hdrs  = List.Select(List.Positions(Raw), (i) =>
                        List.Count(List.Select(Raw{i}, each ParsePeriod(_) <> null)) >= 2),
            Blocks = List.Transform(List.Positions(Hdrs), (k) =>
                let
                    h     = Hdrs{k},
                    hEnd  = if k < List.Count(Hdrs) - 1 then Hdrs{k + 1} else N,
                    hRow  = Raw{h},
                    Cols0 = List.Select(List.Positions(hRow), (c) => ParsePeriod(hRow{c}) <> null),
                    // если на строке второй блок с теми же периодами (как у FR158 справа) — отрезаем
                    Cols  = List.Accumulate(Cols0, [seen = {}, out = {}, stop = false], (st, c) =>
                                let l = ParsePeriod(hRow{c})[Label] in
                                if st[stop] or List.Contains(st[seen], l) then st & [stop = true]
                                else [seen = st[seen] & {l}, out = st[out] & {c}, stop = false])[out],
                    LeftTxt = List.Select(List.FirstN(List.Positions(hRow), Cols{0}), each hRow{_} is text),
                    LC    = if List.IsEmpty(LeftTxt) then 1 else List.Last(LeftTxt),
                    Sec   = if hRow{LC} is text then Text.Trim(hRow{LC}) else null,
                    RowsIx = List.Numbers(h + 1, hEnd - h - 1),
                    Recs  = List.Combine(List.Transform(RowsIx, (r) =>
                                let row = Raw{r}, it = row{LC} in
                                if not (it is text) or Text.Trim(it) = "" or Text.StartsWith(Text.Trim(it), "*") then {}
                                else List.Transform(List.Select(Cols, each row{_} is number), (c) =>
                                    let p = ParsePeriod(hRow{c}) in
                                    [Sheet = sh, SectionNo = k + 1, Section = Sec, RowNo = r + 1, ColNo = c + 1,
                                     Item = Text.Trim(it), PeriodLabel = p[Label], PeriodType = p[Type],
                                     Year = p[Year], Scenario = p[Scenario], Value = row{c}])))
                in
                    Recs)
        in
            List.Combine(Blocks),

    All  = List.Combine(List.Transform(Table.ToRecords(Sheets), each ReadOne([Data], [Name]))),
    Tbl  = Table.FromRecords(All, type table [
              Sheet = text, SectionNo = Int64.Type, Section = text, RowNo = Int64.Type, ColNo = Int64.Type,
              Item = text, PeriodLabel = text, PeriodType = text, Year = Int64.Type,
              Scenario = text, Value = number]),
    Out  = Table.AddColumn(Table.AddColumn(Table.AddColumn(Tbl,
              "Entity", each Entity, type text),
              "EntityName", each EntName, type text),
              "File", each fileName, type text)
in
    Out


#####################################################################
 4. TEST на одном файле (Carica: OFF), потом удалить
#####################################################################
let
    F = Table.SelectRows(SharePoint.Files(SiteUrl, [ApiVersion = 15]),
            each Text.Contains([Name], "GR002")){0},
    R = fnCPRead(F[Content], F[Name], {"Free Tangible Capital analysis"})
in
    R

>> Проверка (GR002): Item = "Free Tangible Capital",
>>   PeriodLabel "Forecast 2026" -> 0 ; "Forecast 2026 SII Review" -> 9370899,2
>>   Section = "Free Tangible Capital Reconciliation vs LGAAP SHs' Equity" (или похожее).


#####################################################################
 5. CP_Files (OFF) + CP_Raw (OFF) + CP_Errors (ON, контроль)
#####################################################################

--- CP_Files ---
let
    Src  = SharePoint.Files(SiteUrl, [ApiVersion = 15]),
    Xl   = Table.SelectRows(Src, each
             Text.Contains([Name], FileFilter)
             and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
             and not Text.StartsWith([Name], "~$")),
    Keep = Table.SelectColumns(Xl, {"Name", "Folder Path", "Date modified", "Content"})
in
    Keep

--- CP_Raw --- (по каждой компании — только самый свежий файл)
let
    Sheets  = List.Buffer(Table.SelectRows(CPSheets, each [Include] = true)[CPSheet]),
    Res     = Table.AddColumn(CP_Files, "R", each try fnCPRead([Content], [Name], Sheets)),
    OK      = Table.SelectRows(Res, each not [R][HasError]),
    WithMod = Table.AddColumn(OK, "T", (row) =>
                Table.AddColumn(row[R][Value], "Modified", each row[Date modified], type datetime)),
    Comb    = Table.Combine(WithMod[T]),
    MaxMod  = Table.Group(Comb, {"Entity"}, {{"MaxMod", each List.Max([Modified]), type datetime}}),
    J       = Table.ExpandTableColumn(
                Table.NestedJoin(Comb, {"Entity"}, MaxMod, {"Entity"}, "M", JoinKind.Inner), "M", {"MaxMod"}),
    Latest  = Table.RemoveColumns(Table.SelectRows(J, each [Modified] = [MaxMod]), {"MaxMod"})
in
    Latest

--- CP_Errors ---
let
    Sheets = List.Buffer(Table.SelectRows(CPSheets, each [Include] = true)[CPSheet]),
    Res    = Table.AddColumn(CP_Files, "R", each try fnCPRead([Content], [Name], Sheets)),
    Bad    = Table.SelectRows(Res, each [R][HasError]),
    Msg    = Table.AddColumn(Bad, "Error", each [R][Error][Message], type text),
    Keep   = Table.SelectColumns(Msg, {"Name", "Folder Path", "Error"})
in
    Keep


#####################################################################
 6. CP_Catalog (ON) — список всех строк CP-Tool, основа для маппинга
#####################################################################
let
    G = Table.Group(CP_Raw, {"Sheet", "SectionNo", "Section", "Item"}, {
          {"Entities", each List.Count(List.Distinct([Entity])), Int64.Type},
          {"MinRow",   each List.Min([RowNo]), Int64.Type},
          {"Periods",  each Text.Combine(List.Sort(List.Distinct([PeriodLabel])), " | "), type text}}),
    S = Table.Sort(G, {{"Sheet", Order.Ascending}, {"SectionNo", Order.Ascending}, {"MinRow", Order.Ascending}})
in
    S

>> Entities = во скольких компаниях строка есть. Если сильно меньше остальных —
>> у части компаний другой шаблон (как FR158).


#####################################################################
 7. CPMap — СЛОВАРЬ CP-Tool -> CDD (главная ручная работа, делается 1 раз)
#####################################################################
>> Где хранить: Excel "CPMap.xlsx" в той же библиотеке SharePoint,
>> внутри — Таблица Excel (Ctrl+T) с именем CPMap.
>> Как заполнить: в Power BI открой CP_Catalog > выдели всё > Copia tabella >
>> вставь в CPMap.xlsx и добавь колонки справа.

 Колонки CPMap:
   CPSheet | CPSection | CPItem      <- из CP_Catalog (Sheet, Section, Item)
   Include      TRUE/FALSE           <- брать строку или нет
   Topic                              <- как в BlockConfig CDD (Solvency, Remittance, FTC...)
   CDDSheet                           <- имя листа CDD (как колонка Sheet в FactData)
   BlockNo | Block | Section | Group  <- как в FactData CDD для этой страницы
   Item                               <- ТОЧНОЕ имя строки, как в DimItem
   ItemOrder                          <- порядок строки (идёт в RowNo)
   SignHint | IsTotal | Measure       <- как в FactData CDD
   Scale                              <- 0,000001 для сумм (-> m), 1 для % и коэффициентов

--- CPMap (OFF) ---
let
    F   = Table.SelectRows(SharePoint.Files(SiteUrl, [ApiVersion = 15]), each [Name] = "CPMap.xlsx"){0}[Content],
    Wb  = Excel.Workbook(F, null, true),
    T   = Wb{[Item = "CPMap", Kind = "Table"]}[Data],
    Ty  = Table.TransformColumnTypes(T, {
            {"Include", type logical}, {"BlockNo", Int64.Type}, {"ItemOrder", Int64.Type},
            {"IsTotal", type logical}, {"Scale", type number}}),
    Inc = Table.SelectRows(Ty, each [Include] = true)
in
    Inc

--- CPMap_Suggest (ON, помощник) --- подсказывает Item CDD, если подписи совпадают
let
    Cdd  = Table.Distinct(Table.SelectColumns(FactData_CDD, {"Topic", "Sheet", "Block", "Item"})),
    CddK = Table.AddColumn(Cdd, "K", each Text.Lower(Text.Trim([Item])), type text),
    Cat  = Table.AddColumn(CP_Catalog, "K", each Text.Lower(Text.Trim([Item])), type text),
    J    = Table.NestedJoin(Cat, {"K"}, CddK, {"K"}, "C", JoinKind.LeftOuter),
    E    = Table.ExpandTableColumn(J, "C", {"Topic", "Sheet", "Block", "Item"},
                                         {"Sugg Topic", "Sugg CDDSheet", "Sugg Block", "Sugg Item"}),
    R    = Table.RemoveColumns(E, {"K"})
in
    R
>> Строки с заполненным "Sugg Item" — готовые кандидаты. Пустые — маппить руками.


#####################################################################
 8. ПЕРИОДЫ И ВАЛЮТА
#####################################################################
--- PeriodFix (OFF) --- переименование периодов CP-Tool в периоды DimPeriod CDD.
>> ЗАПОЛНИ по DimPeriod CDD. Что не указано — остаётся как в CP-Tool.
let
    R = [
        #"Actual 2025"              = "Act 2025",
        #"Forecast 2026"            = "2026",
        #"Plan 2027"                = "Plan 2027",
        #"Plan 2028"                = "Plan 2028",
        #"Plan 2029"                = "Plan 2029",
        #"Forecast 2026 SII Review" = "2026 SII Review"
    ]
in
    R

--- FXRates (OFF) --- плановые курсы YE26 (LC за 1 EUR). Пусто = курс неизвестен.
let
    T = #table(type table [Currency = text, LCperEUR = number], {
        {"EUR", 1}, {"CHF", null}, {"CZK", null}, {"HUF", null}, {"PLN", null},
        {"RON", null}, {"RSD", null}, {"GBP", null}})
in
    T

--- CcyByEntity (OFF) --- валюта по стране (как в EntityMap FTC)
let
    Ccy = [CH="CHF", LI="CHF", CS="RSD", RS="RSD", CZ="CZK", HU="HUF", PL="PLN", RO="RON",
           GB="GBP", AR="ARS", BR="BRL", CL="CLP", EC="USD", MX="MXN", CO="COP", MY="MYR",
           PH="PHP", IN="INR", CN="CNY", HK="HKD", TH="THB", ID="IDR", US="USD", TR="TRY"],
    E   = Table.Distinct(Table.SelectColumns(CP_Raw, {"Entity"})),
    C   = Table.AddColumn(E, "Currency", each Record.FieldOrDefault(Ccy, Text.Start([Entity], 2), "EUR"), type text)
in
    C


#####################################################################
 9. FactData_CP (OFF) и новая FactData (ON)
#####################################################################

--- FactData_CP --- в той же схеме, что FactData CDD
let
    J1   = Table.NestedJoin(CP_Raw, {"Sheet", "Section", "Item"}, CPMap, {"CPSheet", "CPSection", "CPItem"},
                            "M", JoinKind.Inner),
    E1   = Table.ExpandTableColumn(J1, "M",
             {"Topic", "CDDSheet", "BlockNo", "Block", "Section", "Group", "Item", "ItemOrder",
              "SignHint", "IsTotal", "Measure", "Scale"},
             {"Topic", "CDDSheet", "BlockNo", "Block", "CDDSection", "Group", "CDDItem", "ItemOrder",
              "SignHint", "IsTotal", "Measure", "Scale"}),
    J2   = Table.NestedJoin(E1, {"Entity"}, CcyByEntity, {"Entity"}, "C", JoinKind.LeftOuter),
    E2   = Table.ExpandTableColumn(J2, "C", {"Currency"}),
    J3   = Table.NestedJoin(E2, {"Currency"}, FXRates, {"Currency"}, "F", JoinKind.LeftOuter),
    E3   = Table.ExpandTableColumn(J3, "F", {"LCperEUR"}),
    Calc = Table.AddColumn(E3, "V", each
             let sc = if [Scale] = null then 1 else [Scale],
                 fx = if [Scale] = 1 or [LCperEUR] = null then 1 else [LCperEUR]   // % не конвертируем
             in  [Value] * sc / fx, type number),
    Flag = Table.AddColumn(Calc, "FXStatus", each
             if [Currency] = "EUR" or [Scale] = 1 then "OK"
             else if [LCperEUR] = null then "LOCAL CCY (no FX)" else "Converted", type text),
    Ent  = Table.AddColumn(Flag, "EntityLabel", each
             [Entity] & (if [EntityName] = null then "" else " – " & [EntityName]), type text),
    Per  = Table.AddColumn(Ent, "CDDPeriod", each Record.FieldOrDefault(PeriodFix, [PeriodLabel], [PeriodLabel]), type text),
    Out  = Table.SelectColumns(Per, {
             "Topic", "CDDSheet", "BlockNo", "Block", "EntityLabel", "Group", "Scenario", "CDDSection",
             "ItemOrder", "RowNo", "CDDItem", "SignHint", "IsTotal", "CDDPeriod", "PeriodType", "Year",
             "Measure", "V", "Currency", "FXStatus", "File"}),
    Ren  = Table.RenameColumns(Out, {
             {"CDDSheet", "Sheet"}, {"EntityLabel", "Entity"}, {"CDDSection", "Section"},
             {"CDDItem", "Item"}, {"CDDPeriod", "Period"}, {"V", "Value"}, {"RowNo", "CPRowNo"}}),
    Row  = Table.AddColumn(Ren, "RowNo", each if [ItemOrder] <> null then [ItemOrder] else [CPRowNo], Int64.Type),
    Src  = Table.AddColumn(Table.RemoveColumns(Row, {"ItemOrder"}), "Source", each "CP-Tool", type text)
in
    Src

>> PeriodType: CP-Tool даёт Actual/Forecast/Plan. Если в CDD другие значения
>> (например "Act"/"Fcst"), поправь здесь же через Table.ReplaceValue.

--- ПЕРЕКЛЮЧЕНИЕ FactData ---
 ВАЖНО: НЕ переименовывай запрос FactData — иначе все меры DAX переедут на новое имя.
 a) Создай Query vuota "FactData_CDD" (Carica: OFF) и вставь в неё ТЕКУЩИЙ код FactData
    (Editor avanzato > Ctrl+A > Ctrl+C из FactData > вставить).
 b) В FactData замени код на:

let
    Cdd0 = FactData_CDD,
    Cdd  = if Table.HasColumns(Cdd0, "Source") then Cdd0
           else Table.AddColumn(Cdd0, "Source", each "CDD", type text),
    Cp   = FactData_CP,
    Out  = if DataSource = "CP-Tool" then Cp
           else if DataSource = "CDD" then Cdd
           else Table.Combine({Cdd, Cp})
in
    Out

>> Если раньше была интеграция CMP с колонкой Source — она остаётся внутри FactData_CDD.


#####################################################################
 10. "ПО ФОРМУЛАМ": VBA-выгрузка формул (Power Query формулы не видит)
#####################################################################
>> Зачем: (1) понять, из каких ячеек CP-Tool/CP-Reporting берутся строки CDD,
>> и быстро заполнить CPMap; (2) увидеть логику расчёта в CP-Tool (например FTC).
>> Включает скрытые и very hidden листы. Положи в свой AuditTool.xlsm (модуль),
>> открой целевой файл, сделай его активным, запусти DumpFormulas.
>> Результат: новая книга Sheet | Cell | RowLabel | ColHeader | Formula | Value.
>> Ссылки вида '[...]Free Tangible Capital analysis'!D30 в колонке Formula
>> и есть готовая строка для CPMap.

Option Explicit

Sub DumpFormulas()
    Dim src As Workbook, out As Workbook, o As Worksheet, ws As Worksheet
    Dim f As Range, c As Range, r As Long, k As Long, hdr As String
    Set src = ActiveWorkbook
    Set out = Workbooks.Add
    Set o = out.Sheets(1)
    o.Range("A1:I1").Value = Array("Sheet", "Visible", "Cell", "Row", "Col", "RowLabel", "ColHeader", "Formula", "Value")
    r = 2
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For Each ws In src.Worksheets
        Set f = Nothing
        On Error Resume Next
        Set f = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0
        If Not f Is Nothing Then
            For Each c In f.Cells
                hdr = ""
                For k = c.Row - 1 To Application.Max(1, c.Row - 60) Step -1
                    If IsPeriod(CStr(ws.Cells(k, c.Column).Text)) Then
                        hdr = CStr(ws.Cells(k, c.Column).Text): Exit For
                    End If
                Next k
                o.Cells(r, 1).Value = ws.Name
                o.Cells(r, 2).Value = ws.Visible            ' -1 видим, 0 скрыт, 2 very hidden
                o.Cells(r, 3).Value = c.Address(False, False)
                o.Cells(r, 4).Value = c.Row
                o.Cells(r, 5).Value = c.Column
                o.Cells(r, 6).Value = Trim(CStr(ws.Cells(c.Row, 2).Text))
                o.Cells(r, 7).Value = hdr
                o.Cells(r, 8).Value = "'" & c.Formula
                If IsError(c.Value) Then o.Cells(r, 9).Value = "#ERR" Else o.Cells(r, 9).Value = c.Value
                r = r + 1
            Next c
        End If
    Next ws
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    o.ListObjects.Add(xlSrcRange, o.Range("A1").CurrentRegion, , xlYes).Name = "Formulas"
    o.Columns.AutoFit
    MsgBox (r - 2) & " formule esportate da " & src.Name
End Sub

Private Function IsPeriod(ByVal s As String) As Boolean
    s = Trim(s)
    IsPeriod = (s Like "Actual ####*" Or s Like "Forecast ####*" Or s Like "Plan ####*" _
             Or s Like "Act ####*" Or s Like "####" Or s Like "####[*]" Or s Like "HY##")
End Function

>> RowLabel берётся из колонки B (так в CP-Tool). В CDD подписи могут быть в
>> другой колонке — тогда смотри Formula/Cell, этого достаточно для маппинга.


#####################################################################
 11. ПРОВЕРКА
#####################################################################
 [ ] CP_Errors пустая (или понятно, почему файл не читается).
 [ ] CP_Catalog: строки FTC/LGAAP/Solvency есть, Entities ~ число компаний.
 [ ] DataSource = CP-Tool, FTC страница: GR002 Forecast 2026 = 0,
     SII Review = 9,37 (m) — как в таблице FTC pre/post.
 [ ] DataSource = Both: на одной странице сравнить CDD vs CP-Tool по Source.
 [ ] Колонка FXStatus: "LOCAL CCY (no FX)" = курсы ещё не внесены.
 [ ] Страницы без маппинга (SAA, Debt, Initiatives, HMS, Local RAF...) будут
     пустыми при DataSource = CP-Tool — для них оставить CDD (DataSource = Both).
=====================================================================
