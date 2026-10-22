=====================================================================
 ОДНА ПАПКА: CDD-файл (локальные вводы) + CMP Reporting Data (центральные данные)
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ИДЕЯ
>> - CMP Reporting Data уже "длинная" таблица (Entity x Account x Time x Value) —
>>   парсер не нужен, только перевод в формат FactData:
>>     Time FY24 + Versions Plan  -> Period "Plan 2024"  (как в CDD, общий DimPeriod)
>>     AccountCode                -> Topic + Item (через таблицу AccountMap)
>>     Legal Entity Code AT004    -> Entity "AT004 – <имя из EntityList>" (как CompanyMap)
>>     Value Reporting Currency   -> / 1 000 000  (CDD в EUR M)
>> - Оба источника складываются в один FactData с колонкой Source = "CDD" / "CMP".
>>   Если одна и та же цифра есть в обоих (Topic+Entity+Item+Period) — берём CDD
>>   (там уже утверждённый/локальный ввод), иначе CMP.
>> - Все страницы и меры работают как раньше: CMP-строки выглядят для них как CDD-строки.
>> Сделай копию .pbix перед началом.

>> ЧТО ДАСТ CMP, А ЧТО НЕТ (честно)
>> CMP — центральные цифры: Solvency (EOF, SCR и модули), Remittance/дивиденды,
>> Net Result и т.п. Локальных листов (SAA, ORMT, Options, Debt, Initiatives, HMS,
>> Local RAF, Sensitivities) в CMP нет — они остаются из CDD (сейчас пустые/dummy).


#####################################################################
 ШАГ 0. ПАПКА
#####################################################################
>> Положи в одну папку (лучше синхронизированная OneDrive-папка на ПК):
>>   2nd Capital Deep Dive Reporting Dashboard_v3.0.xlsm
>>   7.4 CMP Reporting Data.xlsx
>> Имена могут меняться, важно: в имени CDD-файла есть "Capital Deep Dive",
>> в имени CMP-файла есть "CMP". Закрой файлы в Excel перед Aggiorna.


#####################################################################
 ШАГ 1. ПАРАМЕТР FolderPath
#####################################################################
>> Home > Gestisci parametri > Nuovo parametro:
>>   Nome: FolderPath   Tipo: Testo   Valore corrente: C:\Users\...\CDD_Test\
>> (путь к папке, с "\" в конце). Позже для SharePoint поменяем только этот шаг.


#####################################################################
 ШАГ 2. ЗАПРОС Files (Carica: OFF)
#####################################################################

let
    Src = Folder.Files(FolderPath),
    Xl  = Table.SelectRows(Src, each
              List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))
              and not Text.StartsWith([Name], "~$"))
in
    Xl


#####################################################################
 ШАГ 3. ЗАПРОС Workbook — заменить источник (имя запроса НЕ менять)
#####################################################################
>> Trasforma dati > Workbook > Editor avanzato > Ctrl+A > вставить.
>> Всё остальное (BlockConfig, FactData, List_...) продолжит работать.

let
    F  = Table.SelectRows(Files, each Text.Contains([Name], "Capital Deep Dive")),
    Wb = Excel.Workbook(F{0}[Content], null, true)
in
    Wb


#####################################################################
 ШАГ 4. ЗАПРОС CMP_Raw (Carica: OFF) — колонки по позиции, имена стандартные
#####################################################################

let
    F     = Table.SelectRows(Files, each Text.Contains([Name], "CMP")),
    Wb    = Excel.Workbook(F{0}[Content], null, true),
    Sh    = Table.SelectRows(Wb, each [Kind] = "Sheet"){0}[Data],
    Prom  = Table.PromoteHeaders(Sh, [PromoteAllScalars = true]),
    Std   = {"Plan", "Version", "Time", "LECode", "LEName", "Segment", "Process", "AccountCode",
             "Description", "Counterparty", "InjectionType", "CmpScenario", "LocalCcy",
             "ValueLocal", "ReportingCcy", "FXRate", "ValueEUR"},
    Old   = List.FirstN(Table.ColumnNames(Prom), List.Count(Std)),
    Ren   = Table.RenameColumns(Prom, List.Zip({Old, List.FirstN(Std, List.Count(Old))})),
    Typed = Table.TransformColumnTypes(Ren, {{"ValueEUR", type number}, {"ValueLocal", type number}})
in
    Typed

>> ПРОВЕРКА: колонки Plan, Version, Time, LECode ... ValueEUR; в ValueEUR числа.
>> Если в файле порядок колонок другой — пришли фото заголовков.


#####################################################################
 ШАГ 5. СПИСОК СЧЕТОВ CMP_Accounts (для заполнения маппинга, Carica: OFF)
#####################################################################

let
    G = Table.Group(CMP_Raw, {"Process", "AccountCode", "Description"},
            {{"Rows", each Table.RowCount(_), Int64.Type},
             {"Sum EUR m", each List.Sum([ValueEUR]) / 1000000, type number}})
in
    Table.Sort(G, {{"Process", Order.Ascending}, {"AccountCode", Order.Ascending}})

>> Открой его и пришли фото / скопируй (правый клик по таблице > Copia l'intera tabella):
>> по этому списку я заполню AccountMap целиком (какой счёт = какая строка CDD).


#####################################################################
 ШАГ 6. ТАБЛИЦА AccountMap (Home > Immettere i dati, имя AccountMap)
#####################################################################
>> Колонки: AccountCode | Topic | Item | Sign | Order | IsTotal
>> Item пиши ТОЧНО как строка в CDD (меры ищут по тексту).
>> Sign = 1 или -1 (если знак в CMP противоположен CDD). Order = порядок строки.
>> Стартовые строки (вставь, потом дополним по ШАГУ 5):

AccountCode	Topic	Item	Sign	Order	IsTotal
SII_OFLW_DIV	Remittance by Company	Dividends Outflow (excl. interim dividends)	1	1	FALSE
LGAP_RES_NET	Solvency	Local GAAP Net Result	1	50	FALSE

>> Типы: Sign, Order = Numero intero; IsTotal = Vero/Falso.
>> Счета, которых нет в AccountMap, НЕ теряются: попадут в Topic "CMP (unmapped)".


#####################################################################
 ШАГ 7. ЗАПРОС CMP_Fact (Carica: OFF) — CMP в формате FactData
#####################################################################

let
    Src    = CMP_Raw,
    // сценарий: оставь один (см. проверку ниже); Segment = Total
    Scen   = Table.SelectRows(Src, each [CmpScenario] <> null and Text.StartsWith(Text.From([CmpScenario]), "Aft")),
    Seg    = Table.SelectRows(Scen, each [Segment] = null or Text.Upper(Text.From([Segment])) = "TOTAL"),
    // последний цикл планирования, если в файле их несколько
    LastPl = List.Max(List.RemoveNulls(Seg[Plan])),
    Cyc    = Table.SelectRows(Seg, each [Plan] = LastPl),

    Yr     = Table.AddColumn(Cyc, "Year", each
                 let d = Text.Select(Text.From([Time]), {"0".."9"}) in
                 if Text.Length(d) = 2 then 2000 + Number.From(d)
                 else if Text.Length(d) = 4 then Number.From(d) else null, Int64.Type),
    PType  = Table.AddColumn(Yr, "PeriodType", each
                 let v = Text.Upper(Text.From([Version])) in
                 if Text.StartsWith(v, "ACT") then "Actual"
                 else if Text.StartsWith(v, "FOR") then "Forecast"
                 else if Text.StartsWith(v, "PLAN") then "Plan" else Text.From([Version]), type text),
    Per    = Table.AddColumn(PType, "Period", each
                 (if [PeriodType] = "Actual" then "Act " else if [PeriodType] = "Forecast" then "For "
                  else if [PeriodType] = "Plan" then "Plan " else [PeriodType] & " ") & Text.From([Year]), type text),

    // имя компании как в CompanyMap: "CODE – имя из EntityList"
    JE     = Table.NestedJoin(Per, {"LECode"}, EntityList, {"Code"}, "E", JoinKind.LeftOuter),
    XE     = Table.ExpandTableColumn(JE, "E", {"Entity", "Country"}, {"EntName", "Country"}),
    Ent    = Table.AddColumn(XE, "Entity", each Text.From([LECode]) & " – " &
                 (if [EntName] <> null then [EntName] else Text.From([LEName])), type text),

    // маппинг счетов
    JM     = Table.NestedJoin(Ent, {"AccountCode"}, AccountMap, {"AccountCode"}, "M", JoinKind.LeftOuter),
    XM     = Table.ExpandTableColumn(JM, "M", {"Topic", "Item", "Sign", "Order", "IsTotal"},
                 {"MTopic", "MItem", "MSign", "MOrder", "MIsTotal"}),
    Out    = Table.AddColumn(XM, "R", each [
                 Topic    = if [MTopic] <> null then [MTopic] else "CMP (unmapped)",
                 Sheet    = "CMP",
                 BlockNo  = null,
                 Block    = "CMP " & Text.From([Process]),
                 Group    = null,
                 Scenario = "Current",
                 Section  = null,
                 RowNo    = if [MOrder] <> null then [MOrder] else 9000,
                 Item     = if [MItem] <> null then [MItem] else Text.From([Description]),
                 SignHint = null,
                 IsTotal  = if [MIsTotal] = null then false else [MIsTotal],
                 Measure  = null,
                 Value    = if [ValueEUR] = null then null
                            else [ValueEUR] / 1000000 * (if [MSign] = null then 1 else [MSign])]),
    Keep   = Table.SelectColumns(Out, {"R", "Entity", "Period", "PeriodType", "Year", "Country"}),
    Exp    = Table.ExpandRecordColumn(Keep, "R",
                 {"Topic", "Sheet", "BlockNo", "Block", "Group", "Scenario", "Section", "RowNo",
                  "Item", "SignHint", "IsTotal", "Measure", "Value"}),
    // несколько строк CMP на одну ячейку (контрагенты и т.п.) -> сумма
    Agg    = Table.Group(Exp, {"Topic", "Sheet", "Block", "Entity", "Scenario", "Item", "IsTotal",
                               "Period", "PeriodType", "Year", "Country"},
                 {{"Value", each List.Sum([Value]), type number},
                  {"RowNo", each List.Min([RowNo]), Int64.Type}}),
    Result = Table.AddColumn(Agg, "Source", each "CMP", type text)
in
    Result

>> ПРОВЕРКА CMP_Fact:
>>   - Entity "AT004 – ...", Period "Act 2021", "For 2022", "Plan 2024"...
>>   - Dividends Outflow AT004 Plan 2026: CMP -3 161 479,2 -> у нас -3,16 (EUR m).
>>   - Если строк 0: посмотри в CMP_Raw колонку CmpScenario (distinct) и пришли —
>>     поменяю фильтр "Aft"; то же для Segment.


#####################################################################
 ШАГ 8. ДОБАВИТЬ CMP В КОНЕЦ FactData
#####################################################################
>> FactData > Editor avanzato. Внизу сейчас:
>>     Final2 = Table.RenameColumns(...)
>> in
>>     Final2
>> Поставь запятую после строки Final2 и замени "in / Final2" на:

    WithSrc = Table.AddColumn(Final2, "Source", each "CDD", type text),
    CddKeys = Table.Distinct(Table.SelectColumns(
                  Table.SelectRows(WithSrc, each [Value] <> null), {"Topic", "Entity", "Item", "Period"})),
    CmpOnly = Table.NestedJoin(CMP_Fact, {"Topic", "Entity", "Item", "Period"},
                  CddKeys, {"Topic", "Entity", "Item", "Period"}, "K", JoinKind.LeftAnti),
    CmpRows = Table.RemoveColumns(CmpOnly, {"K"}),
    All     = Table.Combine({WithSrc, CmpRows}),
    Final3  = Table.TransformColumnTypes(All, {{"Value", type number}, {"Year", Int64.Type},
                  {"RowNo", Int64.Type}, {"Source", type text}})
in
    Final3

>> Chiudi e applica. DimItem / DimPeriod подхватят новые Item и Period сами.


#####################################################################
 ШАГ 9. ПРОВЕРКА В ОТЧЁТЕ
#####################################################################
>> [ ] Слайсер Entity: появились компании из CMP (AT004 – ...).
>> [ ] Remittance by Company: у AT004 строки Dividends по годам (в EUR m, со знаком как в CDD).
>> [ ] Solvency: Local GAAP Net Result у AT004 = CMP / 1 000 000.
>> [ ] Новый слайсер/фильтр Source (CDD / CMP) — удобно сверять, откуда цифра.
>> [ ] Topic "CMP (unmapped)" — что ещё не замаплено (видно в Dati > FactData).
>> ВАЖНО: годы CMP (FY21–FY26, цикл 23-25) и CDD-dummy (2025–2029) почти не пересекаются,
>> поэтому в тесте часть страниц покажет CMP-годы отдельно — это нормально для теста.


#####################################################################
 ПОТОМ: SharePoint вместо локальной папки
#####################################################################
>> Меняется только запрос Files:
>>   Src = SharePoint.Files("https://generali-my.sharepoint.com/personal/<твоя_папка>", [ApiVersion = 15]),
>>   + фильтр по [Folder Path], содержащему имя папки.
>> Остальное не трогаем.
