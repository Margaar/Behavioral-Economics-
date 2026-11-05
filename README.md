=====================================================================
 CAPITAL PERFORMANCE REVIEW (CPR) -> POWER BI   |   СКЕЛЕТ из копии CDD
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ИДЕЯ
>> Движок тот же: Workbook -> fnReadSheet -> BlockConfig -> FactData -> DimItem/DimPeriod.
>> Меняем: какой файл читаем, список листов (BlockConfig), FactData без CDD-логики,
>> порядок периодов (Forecast 2026 раньше Actual 2026) и одну правку в fnReadSheet
>> (в CPR колонка "Actual 2026" повторяется дважды: H и жёлтая J -> берём один раз).
>>
>> Что видно в файле CPR (по фото):
>>   - листы "... Entity": подпись в E, периоды F:H (Actual 2025 | Forecast 2026 | Actual 2026),
>>     жёлтая колонка J "Actual 2026" (ввод), справа таблицы вариаций "Actual 2026 vs ..." (выход).
>>   - Company #1, #2 ... над каждым блоком (как в CDD) -> Entity подхватится сам.
>>   - Local GAAP Details: ДВЕ таблицы рядом (Life слева, Non-Life/Holding справа) -> 2 строки конфига.
>>   - IFRS Details: справа служебные колонки (1/2, Life/Non-Life) -> окно режем до "Year on Year".
>>   - Файл-шаблон пустой (везде "-"), поэтому цифр не будет: строим структуру.


#####################################################################
 ШАГ 0. КОПИЯ И ПАПКА
#####################################################################
>> 1) Сохрани копию CDD .pbix как "CPR_Report.pbix" и работай только в ней.
>> 2) Создай новую папку, положи туда ТОЛЬКО файл Capital Performance Review Dashboard.
>> 3) Trasforma dati > Gestisci parametri > FolderPath = путь к новой папке.


#####################################################################
 ШАГ 1. УДАЛИТЬ ЛИШНИЕ ЗАПРОСЫ (иначе обновление упадёт: этих листов в CPR нет)
#####################################################################
>> Trasforma dati > слева правый клик > Elimina:
>>   List_Options, List_ORMT, List_Initiatives, List_Debt (и List_DebtExisting/New, если есть),
>>   List_RAF / List_LocalRAF, CompanyMap, CMP_Raw, CMP_Accounts, CMP_Fact, AccountMap,
>>   CDD_Items, IncludeCMP, DebugSAA.
>> ОСТАВИТЬ: FolderPath, Files, Workbook, fnReadSheet, fnReadList, BlockConfig, CheckBlocks,
>>   FactData, DimItem, DimPeriod, EntityList, DimEntity, Thresholds, SummaryIndex, Scale-таблицы.
>> Если Power BI скажет "используется другим запросом" — сначала сделай ШАГ 5 (новый FactData),
>> потом удаляй.


#####################################################################
 ШАГ 2. Workbook — читать файл CPR
#####################################################################
>> Workbook > Editor avanzato > Ctrl+A > вставить:

let
    F  = Table.SelectRows(Files, each Text.Contains([Name], "Capital Performance")),
    Wb = Excel.Workbook(F{0}[Content], null, true)
in
    Wb


#####################################################################
 ШАГ 3. fnReadSheet — одна правка (дубль периода)
#####################################################################
>> fnReadSheet > Editor avanzato. Найди строку, которая начинается с  ValPos   =
>> Она выглядит так (3 строки):
>>
>>            ValPos   = List.Select(List.Positions(ValArea), (i) =>
>>                           if HasSub then IsSubTxt(SubCells{i}) and HFill{i} <> null
>>                           else IsPeriod(HRow{ValArea{i}})),
>>
>> СРАЗУ ПОСЛЕ неё вставь две строки:

            ValKey   = List.Transform(ValPos, (i) => HFill{i} & "|" & (if HasSub then Txt(SubCells{i}) else "")),
            ValPosU  = List.Select(ValPos, (i) => List.PositionOf(ValKey, ValKey{List.PositionOf(ValPos, i)}) = List.PositionOf(ValPos, i)),

>> Потом ниже, в шаге Recs, замени
>>        else List.Transform(ValPos, (i) => [
>> на
>>        else List.Transform(ValPosU, (i) => [
>>
>> Что делает: если в строке заголовка один и тот же период встречается дважды
>> (H "Actual 2026" и J "Actual 2026"), берётся только первый (H). В CDD такого нет,
>> поэтому на логику CDD это не влияет.
>> Если в H пусто, а данные только в J — скажи, переключу на "последний".


#####################################################################
 ШАГ 4. BlockConfig — листы CPR
#####################################################################
>> BlockConfig > Editor avanzato > Ctrl+A > вставить.
>> ToText "Actual 2026 vs" = начало таблиц вариаций справа: окно чтения заканчивается
>> перед ними. В следующем году поменять на "Actual 2027 vs".

let
    Config = #table(
        type table [Topic = text, Sheet = text, FromText = nullable text, ToText = nullable text,
                    SideGroup = logical, SectionPrefix = nullable text],
        {
            {"Total Distributions BU", "Total Distributions BU",      null,             "Actual 2026 vs", false, null},
            {"Solvency II",            "Solvency II Entity",          null,             "Actual 2026 vs", false, null},
            {"Local Solvency",         "Local Solvency Entity",       null,             "Actual 2026 vs", false, null},
            {"SST",                    "SST Entity",                  null,             "Actual 2026 vs", false, null},
            {"Holding",                "Holding Entity",              null,             "Actual 2026 vs", false, null},
            {"Other Sectoral",         "Other Sect. Solvency Entity", null,             "Actual 2026 vs", false, null},
            {"Local GAAP Life",        "Local GAAP Details",          null,             "LGAAP Non-Life", false, null},
            {"Local GAAP Non-Life",    "Local GAAP Details",          "LGAAP Non-Life", null,             false, null},
            {"IFRS",                   "IFRS Details",                null,             "Year on Year",   false,
                                       "P&C Operating Result|Life Operating Result"}
        })
in
    Config

>> Имена листов должны совпасть ТОЧНО (пробелы, точки). Листы, которые я не видела целиком
>> ("Holding L...", вводные Capital Composition / Participations / IC Assets and Liabilities),
>> добавим после ШАГА 8 — пришли их список.


#####################################################################
 ШАГ 5. FactData — чистая версия для CPR
#####################################################################
>> FactData > Editor avanzato > Ctrl+A > вставить.
>> Убрано всё CDD-специфичное: Remittance Previous, NoEntCountry, CompanyMap, CMP.
>> Пустые строки НЕ фильтруем: шаблон пустой, а структура строк нужна для страниц.

let
    WithData = Table.AddColumn(BlockConfig, "Data", each
        try fnReadSheet(Workbook{[Item = [Sheet], Kind = "Sheet"]}[Data],
                        [FromText], [ToText], [SideGroup], [SectionPrefix])
        otherwise null),
    Found    = Table.SelectRows(WithData, each [Data] <> null),
    Expanded = Table.ExpandTableColumn(Found, "Data",
                   {"BlockNo", "Block", "Entity", "Group", "Scenario", "Section", "RowNo",
                    "Item", "SignHint", "IsTotal", "Period", "PeriodType", "Year", "Measure", "Value"}),
    Removed  = Table.RemoveColumns(Expanded, {"FromText", "ToText", "SideGroup", "SectionPrefix"}),
    Typed    = Table.TransformColumnTypes(Removed, {
                   {"BlockNo", Int64.Type}, {"RowNo", Int64.Type}, {"Year", Int64.Type},
                   {"Value", type number}, {"IsTotal", type logical},
                   {"Block", type text}, {"Entity", type text}, {"Group", type text},
                   {"Section", type text}, {"Item", type text}, {"Period", type text},
                   {"Measure", type text}}),
    WithSrc  = Table.AddColumn(Typed, "Source", each "CPR", type text)
in
    WithSrc

>> Колонку Source оставила, потому что на неё могут ссылаться старые меры/колонки.


#####################################################################
 ШАГ 6. DimPeriod — порядок как в CPR: Actual 2025 | Forecast 2026 | Actual 2026 | Plan 2027
#####################################################################
>> DimPeriod > Editor avanzato > Ctrl+A > вставить.
>> Отличие от CDD: внутри одного года Forecast идёт РАНЬШЕ Actual.

let
    Src     = Table.SelectColumns(FactData, {"Period", "Year", "PeriodType"}),
    Grouped = Table.Group(Src, {"Period"},
                  {{"Year", each List.Max([Year]), Int64.Type},
                   {"PeriodType", each List.First([PeriodType]), type text}},
                  GroupKind.Global, Comparer.OrdinalIgnoreCase),
    Ordered = Table.AddColumn(Grouped, "PeriodOrder", each
                  (if [Year] = null then 9999 else [Year]) * 10
                  + (if [PeriodType] = "Forecast" then 1
                     else if [PeriodType] = "Half-Year" then 2
                     else if [PeriodType] = "Actual" then 3
                     else if [PeriodType] = "Year" then 5
                     else if [PeriodType] = "Plan" then 6 else 7), Int64.Type)
in
    Ordered


#####################################################################
 ШАГ 7. НОВЫЙ ЗАПРОС CPR_Info (шапка с листа Start)  — load ON
#####################################################################
>> Nuova origine > Query vuota > Editor avanzato > вставить > имя CPR_Info.
>> Ищет подпись на листе Start и берёт первое непустое значение справа.
>> Пригодится для подзаголовка страниц: год, квартал, юрлицо, валюта, масштаб.

let
    Sheet = Workbook{[Item = "Start", Kind = "Sheet"]}[Data],
    Rows  = List.Buffer(Table.ToRows(Table.ReplaceErrorValues(Sheet,
                List.Transform(Table.ColumnNames(Sheet), each {_, null})))),
    Txt   = (v) => if v = null then null else Text.Trim(Text.From(v)),
    Get   = (lab as text) => List.First(List.RemoveNulls(List.Transform(Rows, (row) =>
                let flags = List.Transform(row, (v) => let t = Txt(v) in t <> null and Text.StartsWith(t, lab)),
                    p     = List.PositionOf(flags, true)
                in if p < 0 then null
                   else List.First(List.Select(List.Skip(row, p + 1), (v) => Txt(v) <> null and Txt(v) <> ""), null))),
            null),
    Result = #table(type table [Field = text, Value = text], {
        {"Reference year",    Txt(Get("Reference year"))},
        {"Quarter",           Txt(Get("Quarter of submission"))},
        {"Entity IFICC",      Txt(Get("Entity IFICC"))},
        {"Entity Legal Name", Txt(Get("Entity Legal Name"))},
        {"Currency",          Txt(Get("Currency of Reporting"))},
        {"Scale",             Txt(Get("Scale of Reporting"))}})
in
    Result

>> Если лист называется не "Start" — поменяй имя в первой строке.
>> Если какое-то поле пустое — подпись в файле начинается иначе: пришли фото листа Start крупно.


#####################################################################
 ШАГ 8. ПРОВЕРКА В POWER QUERY (до Chiudi e applica)
#####################################################################
>> a) CheckBlocks: по каждой строке "N rows | M blocks: ...".
>>    Ожидаю для листов Entity блоки: Solvency Evolution / Remittance /
>>    Distributions and Other Capital Movements / Distribution ratios ...
>>    ERROR "ToText not found" = на этом листе нет текста "Actual 2026 vs" -> поставь null.
>>    0 blocks = не нашёлся заголовок с периодами -> пришли фото.
>> b) FactData: колонка Entity = Company #1 ... ; Period = Actual 2025 / Forecast 2026 / Actual 2026
>>    (у Total Distributions ещё Plan 2027). Двойного "Actual 2026" быть не должно.
>> c) Список всех листов: Nuova origine > Query vuota > вставить > посмотреть > потом удалить:

let
    S = Table.SelectColumns(Table.SelectRows(Workbook, each [Kind] = "Sheet"), {"Name", "Hidden"})
in
    S

>>    Пришли мне фото этого списка: добавлю в BlockConfig остальные листы
>>    (Holding L..., Capital Composition, Participations, IC Assets and Liabilities и т.д.).


#####################################################################
 ШАГ 9. МОДЕЛЬ (после Chiudi e applica)
#####################################################################
>> Связи DimItem[Item] -> FactData[Item] и DimPeriod[Period] -> FactData[Period] уже есть.
>> DimPeriod[Period] > Ordina per colonna > PeriodOrder (если сбилось).
>>
>> Порядок блоков как в Excel. FactData > Nuova colonna:

Block Order =
CALCULATE ( MIN ( FactData[RowNo] ), ALLEXCEPT ( FactData, FactData[Block] ) )

>> Потом FactData[Block] > Ordina per colonna > Block Order.
>>
>> Старые CDD-меры и вычисляемые колонки (SAA ..., LK ..., Sens ..., MC ... и т.д.) не мешают
>> обновлению; удалить их можно потом. Страницы CDD удали, оставь одну как шаблон дизайна
>> (шапка + меню + блок "Key messages").


#####################################################################
 ШАГ 10. МЕРЫ CPR (Nuova misura, папка "CPR")
#####################################################################
>> Год отчёта = последний год с Actual (2026). Меры работают в матрице БЕЗ периодов в колонках.

CPR Year =
CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( DimPeriod ), FactData[PeriodType] = "Actual" )

Act CY =
VAR y = [CPR Year]
RETURN CALCULATE ( [Value Current], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Actual", FactData[Year] = y )

Act PY =
VAR y = [CPR Year]
RETURN CALCULATE ( [Value Current], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Actual", FactData[Year] = y - 1 )

Fcst CY =
VAR y = [CPR Year]
RETURN CALCULATE ( [Value Current], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Forecast", FactData[Year] = y )

Plan NY =
VAR y = [CPR Year]
RETURN CALCULATE ( [Value Current], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Plan", FactData[Year] = y + 1 )

>> Вариации как в правых таблицах Excel: для коэффициентов — в п.п., для сумм — в %.

CPR Is Ratio =
VAR it = SELECTEDVALUE ( FactData[Item] )
RETURN CONTAINSSTRING ( it, "Ratio" ) || CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( it, "p.p." )

Var vs PY =
VAR cy = [Act CY]
VAR py = [Act PY]
RETURN IF ( NOT ISBLANK ( cy ) || NOT ISBLANK ( py ),
           IF ( [CPR Is Ratio], ( cy - py ) * 100, DIVIDE ( cy - py, ABS ( py ) ) ) )

Var vs Fcst =
VAR cy = [Act CY]
VAR fc = [Fcst CY]
RETURN IF ( NOT ISBLANK ( cy ) || NOT ISBLANK ( fc ),
           IF ( [CPR Is Ratio], ( cy - fc ) * 100, DIVIDE ( cy - fc, ABS ( fc ) ) ) )

>> ФОРМАТ для Var vs PY и Var vs Fcst: выдели меру > Strumenti misura > Formato > Dinamico >
>> в строке формул (слева "Formato") вставь:
>>
>>   IF ( [CPR Is Ratio], "+0.0 ""p.p."";-0.0 ""p.p."";0.0 ""p.p.""", "+0.0%;-0.0%;0.0%" )
>>
>> Подпись колонок в визуале: переименуй поле в визуале (двойной клик) на
>> "Actual vs PY" и "Actual vs Forecast" — год в названии не нужен, он идёт из CPR Year.

CPR Subtitle =
VAR ent = SELECTEDVALUE ( FactData[Entity], "All entities" )
VAR q   = LOOKUPVALUE ( CPR_Info[Value], CPR_Info[Field], "Quarter" )
VAR nm  = LOOKUPVALUE ( CPR_Info[Value], CPR_Info[Field], "Entity Legal Name" )
RETURN "Capital Performance Review · " & [CPR Year]
       & IF ( NOT ISBLANK ( q ) && q <> "", " " & q, "" )
       & IF ( NOT ISBLANK ( nm ) && nm <> "", " · " & nm, "" )
       & " · " & ent


#####################################################################
 ШАГ 11. СТРАНИЦЫ (одна страница = один лист Excel)
#####################################################################
>> Страницы: Total Distributions BU | Solvency II | Local Solvency | SST | Holding |
>>           Other Sectoral | Local GAAP (Life + Non-Life) | IFRS
>> Шаблон страницы (сделай одну, потом Duplica pagina и поменяй фильтр Topic):
>>
>> 1) Filtri in questa pagina: FactData[Topic] = "Solvency II" (на Local GAAP: оба Local GAAP ...).
>> 2) Слайсер Entity (Elenco a discesa) в ленте фильтров. На Total Distributions BU не нужен.
>> 3) МАТРИЦА "как в Excel" (основной блок):
>>      Righe:   FactData[Block] > FactData[Item]
>>      Colonne: DimPeriod[Period]
>>      Valori:  Value Current
>>      Righe: стрелка у Item > "Mostra elementi senza dati"  (шаблон пустой — иначе строк не видно)
>>      Layout: Struttura (rientrato); Totali e Subtotali: Disattivati.
>> 4) МАТРИЦА "вариации" справа от неё (как правые таблицы Excel):
>>      Righe:   FactData[Block] > FactData[Item]
>>      Valori:  Act CY | Var vs PY | Var vs Fcst
>>      Без колонок периода. Цвет Var: Formattazione condizionale > Regole: <0 красный, >0 зелёный.
>> 5) Шапка: карточка/текст с мерой CPR Subtitle.
>> 6) Блок "Key messages" справа — пустой (комментарии сделаем после теста на данных).
>> 7) IFRS: в Righe добавь FactData[Section] перед Item (P&C / Life).
>>
>> Summary и автокомментарии — следующий шаг, когда будут реальные данные или тестовый файл.


#####################################################################
 ЧТО ПРИСЛАТЬ МНЕ
#####################################################################
>> [ ] фото CheckBlocks
>> [ ] фото списка листов (ШАГ 8c)
>> [ ] фото CPR_Info
>> [ ] если есть: заполненный (хоть тестовый) CPR — без цифр комментарии и пороги не проверить
