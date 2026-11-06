=====================================================================
ПРАВКИ ПОСЛЕ CheckBlocks  (fnReadSheet целиком + BlockConfig целиком)
=====================================================================

>> ЧТО ЧИНИМ (по твоему фото CheckBlocks)
>>  - строка 18 "P&C Operating Result|Life Operating Result" — это не лист, а SectionPrefix
>>    от IFRS, съехавший в отдельную строку -> в новом BlockConfig его нет, он в строке IFRS.
>>  - Sensitivities 0 rows: в каждой таблице ОДИН период (Forecast 2026), а движок ждал >= 2.
>>    -> если на листе нет ни одного заголовка с 2+ периодами, берутся заголовки с одним.
>>  - ORMT Impacts / Solvency AoM 0 rows: в строке заголовка, видимо, есть служебные числа
>>    (скрытые колонки) -> правило "в заголовке ноль чисел" ослаблено до "периодов больше, чем чисел".
>>  - Solvency AoM: EOF и SCR имеют одинаковые периоды -> дубль-фильтр теперь учитывает группу
>>    (Eligible Own Funds / SCR), иначе колонки SCR пропали бы. Правые "Actual 2026" (ввод) убираются.
>>  - Local GAAP Details: на листе ТРИ таблицы (Total / Life / Non-Life) -> три строки,
>>    каждая читает от своего заголовка до следующего "LGAAP".
>>  - Non-Life Translation Framework 0 rows: две таблицы рядом (итог слева, по компании справа)
>>    -> две строки с разными окнами.
>>  - "YE 2025 / For 2026" (Translation Framework) переводятся в "Actual 2025 / Forecast 2026",
>>    чтобы периоды совпадали с остальными листами.
>>  - ORMT: SideGroup = true (<ORMT #1>, <ORMT #2> слева -> колонка Section).


#####################################################################
 ШАГ 1. fnReadSheet — ЗАМЕНИТЬ ЦЕЛИКОМ
#####################################################################
>> Trasforma dati > fnReadSheet > Editor avanzato > Ctrl+A > вставить > Fine.
>> (Это версия для CPR. В CDD-файле старую не трогай.)

(Sheet as table, optional FromText as nullable text, optional ToText as nullable text,
 optional SideGroup as nullable logical, optional SectionPrefix as nullable text) as table =>
let
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = List.Buffer(Table.ToRows(Clean)),
    NR    = List.Count(Rows),
    NC    = Table.ColumnCount(Sheet),
    Txt   = (v) => if v = null then null
                   else Text.Trim(Text.Replace(Text.Replace(Text.From(v), "#(cr)", ""), "#(lf)", " ")),

    // column window: FromText (rightmost match) ... first ToText to the right (or sheet end)
    FindCols = (t as text) as list => List.Combine(List.Transform(Rows, (row) =>
                   List.Select(List.Positions(row), (c) => row{c} is text and Text.StartsWith(Txt(row{c}), t)))),
    C0 = if FromText = null or FromText = "" then 0
         else let f = FindCols(FromText) in
              if List.IsEmpty(f) then error Error.Record("FromText not found", FromText) else List.Max(f),
    C1 = if ToText = null or ToText = "" then NC - 1
         else let f = List.Select(FindCols(ToText), (c) => c > C0) in
              if List.IsEmpty(f) then NC - 1 else List.Min(f) - 1,
    WinCols = List.Buffer(List.Numbers(C0, C1 - C0 + 1)),

    IsPeriod = (v) =>
        if v is number then v >= 2000 and v <= 2100 and Number.Round(v) = v
        else if v is text then
            (let s = Text.Upper(Txt(v)), d = Text.Select(s, {"0".."9"}) in
                (Text.Length(d) = 4 and Text.StartsWith(d, "20"))
                or (Text.StartsWith(s, "HY") and Text.Length(d) = 2)
                or s = "BENCHMARK")
        else false,
    IsPerLabel = (v) => v is text and List.AnyTrue(List.Transform({"ACT", "FOR", "PLAN", "YE"},
                     (p) => Text.StartsWith(Text.Upper(Txt(v)), p))),

    // header rows: 2+ periods and more periods than numbers;
    // if the window has none, single-period headers (Forecast 2026) are used
    RowInfo = (r as number) =>
        let cells = List.Transform(WinCols, (c) => Rows{r}{c}),
            pers  = List.Select(cells, IsPeriod),
            nNum  = List.Count(List.Select(cells, (v) => v is number and not IsPeriod(v)))
        in [nPer = List.Count(pers), nNum = nNum, lone = List.Count(pers) = 1 and nNum = 0 and IsPerLabel(pers{0})],
    Infos    = List.Buffer(List.Transform(List.Numbers(0, NR), RowInfo)),
    HdrMulti = List.Select(List.Numbers(0, NR), (r) => Infos{r}[nPer] >= 2 and Infos{r}[nNum] < Infos{r}[nPer]),
    HdrLone  = List.Select(List.Numbers(0, NR), (r) => Infos{r}[lone]),
    HdrRows  = List.Buffer(if List.IsEmpty(HdrMulti) then HdrLone else HdrMulti),

    Junk = (t) => t = null or t = "" or t = "0" or Text.StartsWith(t, "*")
                  or Text.Upper(t) = "TO ALLOCATE" or Text.StartsWith(Text.Upper(t), "WHERE NEEDED"),
    FillFwd = (lst as list) as list => List.Accumulate(lst, {}, (s, v) =>
                  s & {if v <> null and v <> "" then v else (if List.IsEmpty(s) then null else List.Last(s))}),
    Prefixes = if SectionPrefix = null or SectionPrefix = "" then {} else Text.Split(SectionPrefix, "|"),
    IsSec = (t) => t <> null and List.AnyTrue(List.Transform(Prefixes, (p) => Text.StartsWith(t, p))),

    // entity: "Company #n" up to 7 rows above the header, from the label area to 3 columns left of the window
    P0Of  = (h) => List.Min(List.Select(WinCols, (c) => IsPeriod(Rows{h}{c}))),
    EntOf = (h) =>
        let p0   = P0Of(h),
            cLo  = List.Max({0, C0 - 3}),
            cols = List.Numbers(p0 - 1, List.Max({0, p0 - cLo}), -1),
            cells = List.Combine(List.Transform(List.Select(List.Numbers(h, 8, -1), (r) => r >= 0),
                        (r) => List.Transform(cols, (c) => Txt(Rows{r}{c}))))
        in List.First(List.Select(cells, (t) => t <> null and Text.StartsWith(Text.Upper(t), "COMPANY #")), null),
    EntFill = List.Buffer(FillFwd(List.Transform(HdrRows, EntOf))),

    ReadBlock = (k as number) as list =>
        let
            h       = HdrRows{k},
            hEnd    = if k + 1 < List.Count(HdrRows) then HdrRows{k + 1} else NR,
            HRow    = Rows{h},
            p0      = P0Of(h),
            BodyAll = List.Numbers(h + 1, hEnd - h - 1),
            HasText = (c) => List.AnyTrue(List.Transform(BodyAll, (r) =>
                          Rows{r}{c} is text and not Junk(Txt(Rows{r}{c})))),
            LabCand = List.Select(List.Reverse(List.Select(WinCols, (c) => c < p0)), HasText),
            LabelCol = if List.IsEmpty(LabCand) then null else LabCand{0},

            ValArea = List.Select(WinCols, (c) => c >= p0),
            HFill   = FillFwd(List.Transform(ValArea, (c) => if IsPeriod(HRow{c}) then Txt(HRow{c}) else null)),

            SubR     = h + 1,
            SubOK    = SubR < hEnd,
            SubCells = if SubOK then List.Transform(ValArea, (c) => Rows{SubR}{c})
                       else List.Repeat({null}, List.Count(ValArea)),
            IsSubTxt = (v) => v is text and not IsPeriod(v)
                       and not List.Contains({"", "N.A.", "N.A", "-", "N/A"}, Text.Upper(Txt(v))),
            SubLab   = if SubOK and LabelCol <> null then Txt(Rows{SubR}{LabelCol}) else null,
            HasSub   = SubOK and List.Count(List.Select(SubCells, IsSubTxt)) >= 2
                       and (SubLab = null or SubLab = "" or Text.StartsWith(Text.Lower(SubLab), "amounts")),
            ValPos   = List.Select(List.Positions(ValArea), (i) =>
                           if HasSub then IsSubTxt(SubCells{i}) and HFill{i} <> null
                           else IsPeriod(HRow{ValArea{i}})),

            // group row above the periods (Eligible Own Funds / Solvency Capital Requirement ...)
            GCells  = if h > 0 then List.Transform(ValArea, (c) =>
                          let t = Txt(Rows{h - 1}{c}) in if t = null or t = "" or IsPeriod(t) then null else t)
                      else List.Repeat({null}, List.Count(ValArea)),
            UseGrp  = List.Count(List.RemoveNulls(GCells)) >= 2,
            GFill   = if UseGrp then FillFwd(GCells) else List.Repeat({null}, List.Count(ValArea)),

            // same group + period + sub-header twice (H "Actual 2026" and yellow J "Actual 2026") -> keep the first
            ValKey  = List.Transform(ValPos, (i) =>
                          (if GFill{i} = null then "" else GFill{i}) & "|" & HFill{i} & "|"
                          & (if HasSub then Txt(SubCells{i}) else "")),
            ValPosU = List.Select(ValPos, (i) =>
                          List.PositionOf(ValKey, ValKey{List.PositionOf(ValPos, i)}) = List.PositionOf(ValPos, i)),

            BodyStart = if HasSub then h + 2 else h + 1,
            Body    = List.Numbers(BodyStart, List.Max({0, hEnd - BodyStart})),
            UseSide = SideGroup = true and LabelCol <> null and LabelCol > 0,
            SecVals = List.Transform(Body, (r) =>
                          if UseSide then Txt(Rows{r}{LabelCol - 1})
                          else let l = Txt(Rows{r}{LabelCol}) in if IsSec(l) then l else null),
            SecFill = FillFwd(SecVals),

            HLabel  = Txt(HRow{LabelCol}),
            HLeft   = List.Select(List.Transform(List.Reverse(List.Select(WinCols, (c) => c < LabelCol)),
                          (c) => Txt(HRow{c})), (t) => t <> null and t <> "" and not IsPeriod(t)),
            FirstIt = List.First(List.Select(List.Transform(Body, (r) => Txt(Rows{r}{LabelCol})),
                          (t) => not Junk(t)), null),
            BlockNm = if HLabel <> null and HLabel <> "" and HLabel <> "0" and not IsPeriod(HLabel) then HLabel
                      else if not List.IsEmpty(HLeft) then HLeft{0}
                      else if FirstIt <> null then FirstIt
                      else "Block " & Text.From(k + 1),
            EntV    = EntFill{k},

            Recs = List.Combine(List.Transform(List.Positions(Body), (j) =>
                       let r = Body{j}, lab = Txt(Rows{r}{LabelCol}) in
                       if Junk(lab) then {}
                       else List.Transform(ValPosU, (i) => [
                           BlockNo  = k + 1,
                           Block    = BlockNm,
                           Entity   = EntV,
                           Grp      = GFill{i},
                           Section  = SecFill{j},
                           RowNo    = r + 1,
                           RawItem  = lab,
                           Period   = HFill{i},
                           Measure  = if HasSub then Txt(SubCells{i}) else null,
                           RawValue = Rows{r}{ValArea{i}}
                       ])))
        in
            if LabelCol = null then {} else Recs,

    Long  = List.Combine(List.Transform(List.Positions(HdrRows), ReadBlock)),
    Cols0 = {"BlockNo", "Block", "Entity", "Grp", "Section", "RowNo", "RawItem", "Period", "Measure", "RawValue"},
    T0 = if List.IsEmpty(Long) then #table(Cols0, {}) else Table.FromRecords(Long),

    // YE 2025 -> Actual 2025, For 2026 -> Forecast 2026, Act 2025 -> Actual 2025
    T0b = Table.TransformColumns(T0, {{"Period", (p) =>
              let s = Text.Upper(Text.Trim(p)), d = Text.Select(p, {"0".."9"}) in
              if Text.Length(d) <> 4 then p
              else if Text.StartsWith(s, "YE") or Text.StartsWith(s, "ACT ") then "Actual " & d
              else if Text.StartsWith(s, "FOR ") then "Forecast " & d
              else p, type text}}),

    T1 = Table.AddColumn(T0b, "Scenario", each
            if Text.Contains(Text.Upper((if [Grp] = null then "" else [Grp]) & " " & [Block]), "PREVIOUS")
            then "Previous" else "Current", type text),
    T2 = Table.AddColumn(T1, "Year", each
            let d = Text.Select([Period], {"0".."9"}) in
            if Text.Length(d) = 4 then Number.From(d)
            else if Text.Length(d) = 2 then 2000 + Number.From(d) else null, Int64.Type),
    T3 = Table.AddColumn(T2, "PeriodType", each
            let s = Text.Upper([Period]) in
            if Text.Contains(s, "PRO FORMA") then "Pro Forma SII"
            else if Text.StartsWith(s, "HY") then "Half-Year"
            else if Text.StartsWith(s, "ACT") then "Actual"
            else if Text.StartsWith(s, "FOR") then "Forecast"
            else if Text.StartsWith(s, "PLAN") then "Plan"
            else if s = "BENCHMARK" then "Benchmark"
            else let rest = Text.Trim(Text.Remove([Period], {"0".."9", "*"})) in
                 if rest = "" then "Year" else rest, type text),
    T4 = Table.AddColumn(T3, "SignHint", each
            let t = [RawItem] in
            if Text.Contains(t, "(+/-)") then "+/-"
            else if Text.Contains(t, "(-/+)") then "-/+"
            else if Text.Contains(t, "(-)") then "-"
            else if Text.Contains(t, "(+)") then "+" else null, type text),
    T5 = Table.AddColumn(T4, "Item", each
            Text.Trim(Text.Remove(List.Accumulate({"(+/-)", "(-/+)", "(-)", "(+)"}, [RawItem],
                (s, p) => Text.Replace(s, p, "")), {"*"})), type text),
    T6 = Table.AddColumn(T5, "IsTotal", each Text.StartsWith(Text.Upper([Item]), "TOTAL"), type logical),
    T7 = Table.AddColumn(T6, "Value", each
            if [RawValue] is number then [RawValue] else try Number.From([RawValue]) otherwise null, type number),
    T8 = Table.RenameColumns(T7, {{"Grp", "Group"}}),
    Result = Table.SelectColumns(T8, {"BlockNo", "Block", "Entity", "Group", "Scenario", "Section", "RowNo",
                 "Item", "SignHint", "IsTotal", "Period", "PeriodType", "Year", "Measure", "Value"})
in
    Result


#####################################################################
 ШАГ 2. BlockConfig — ЗАМЕНИТЬ ЦЕЛИКОМ
#####################################################################
>> BlockConfig > Editor avanzato > Ctrl+A > вставить > Fine.
>> Названия Topic оставила твои; новые: Local GAAP Total и Non-Life TF by Company.

let
    Config = #table(
        type table [Topic = text, Sheet = text, FromText = nullable text, ToText = nullable text,
                    SideGroup = logical, SectionPrefix = nullable text],
        {
            {"Total Distributions BU",           "Total Distributions BU",          null,             "Actual 2026 vs",      false, null},
            {"Solvency II",                      "Solvency II Entity",              null,             "Actual 2026 vs",      false, null},
            {"Local Solvency",                   "Local Solvency Entity",           null,             "Actual 2026 vs",      false, null},
            {"SST",                              "SST Entity",                      null,             "Actual 2026 vs",      false, null},
            {"Holding",                          "Holding Entity",                  null,             "Actual 2026 vs",      false, null},
            {"Other Sectoral",                   "Other Sect. Solvency Entity",     null,             "Actual 2026 vs",      false, null},
            {"Local GAAP Total",                 "Local GAAP Details",              "LGAAP Total",    "LGAAP",               false, null},
            {"Local GAAP Life",                  "Local GAAP Details",              "LGAAP Life",     "LGAAP",               false, null},
            {"Local GAAP Non-Life",              "Local GAAP Details",              "LGAAP Non-Life", "LGAAP",               false, null},
            {"IFRS",                             "IFRS Details",                    null,             "Year on Year",        false, "P&C Operating Result|Life Operating Result"},
            {"Local LGAAP",                      "Holding LGAAP",                   null,             null,                  false, null},
            {"Non-Life between LGAAP and IFRS",  "Non-Life Translation Framework",  null,             "For companies where", false, null},
            {"Non-Life TF by Company",           "Non-Life Translation Framework",  "Companies in MM", null,                 false, null},
            {"BU Profitability",                 "Profitability Ratios BU",         null,             null,                  false, null},
            {"Profitability Ratios",             "Profitability Ratios LEs",        null,             null,                  false, null},
            {"ORMT",                             "ORMT Impacts",                    null,             null,                  true,  null},
            {"Solvency II AoM",                  "Solvency AoM",                    null,             null,                  false, null},
            {"Own Funds",                        "OF Capital Generation Items",     null,             null,                  false, null},
            {"Sensitivities",                    "Sensitivities",                   null,             null,                  false, "Impact of"}
        })
in
    Config


#####################################################################
 ШАГ 3. ПРОВЕРКА CheckBlocks — что ожидаю
#####################################################################
>> Total Distributions BU   2 блока: Distributions and Other Capital Movements / Distribution ratios
>>                          (сейчас 1 — если останется 1, пришли фото строк 7-12 этого листа)
>> Local GAAP Total/Life/Non-Life  по 1 блоку каждый, примерно по 100-140 строк, НЕ 414
>> Non-Life ... (2 строки) по 1 блоку, ~17 строк x 3 периода
>> ORMT                     1-2 блока, Section = <ORMT #1> ...
>> Solvency II AoM          1 блок; в FactData колонка Group = Eligible Own Funds / Solvency Capital Requirement
>> Sensitivities            2-3 блока (LGAAP Net Result / IFRS Net Result / Other sens справа)
>> Остальные                как сейчас.
>>
>> Если где-то всё ещё "0 rows | 0 blocks": пришли фото этого листа с номерами строк
>> и буквами колонок (видно, где заголовок периодов).

>> ИЗВЕСТНЫЕ ОГРАНИЧЕНИЯ
>> - Sensitivities: строки <Other sens #1-3> в нижней левой таблице попадут в блок правой
>>   жёлтой таблицы (она начинается на той же высоте). Это пустые шаблонные строки — не страшно.
>> - Листы без "Company #" (BU-листы, Holding LGAAP, Local GAAP Details) идут с Entity = пусто.
>>   Когда таблицы будут повторены по компаниям (как пишут в жёлтых заметках), Entity подхватится сам.
>> - Profitability: левые таблицы (C:E) берутся, правые копии (J:L) отбрасываются как дубль.
