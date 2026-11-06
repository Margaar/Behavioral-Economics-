=====================================================================
 
 Основа: engine (fnReadSheet CPR, BlockConfig 19 строк, FactData, DimPeriod, CPR_Info).
=====================================================================

>> ИДЕЯ (чем отличается от CDD)
>> В CDD на каждую страницу были свои меры. В CPR все листы устроены одинаково
>> (Actual 2025 | Forecast 2026 | Actual 2026), поэтому всё сделано ОДНИМ набором мер:
>>   - CPR Value / Act CY / Fcst CY / Act PY / Var vs Fcst / Var vs PY  -> все матрицы
>>   - CPR Comment                                                   -> комментарий ЛЮБОЙ страницы
>>   - SummaryIndex (таблица)  -> для каждой страницы: ключевая строка (KeyItem) и правило светофора
>>   - CPR Headline / CPR Status -> одна строка + точка на Summary, сами берут ключевую строку
>>   - Thresholds -> пороги отклонения Actual vs Forecast; лимиты Hard/Soft берутся ИЗ ФАЙЛА.
>> Новая страница/лист = 1 строка в BlockConfig + 1 строка в SummaryIndex. Меры не трогаешь.
>>
>> Логика CPR = Performance review: главный вопрос "Actual 2026 совпал с Forecast 2026?".
>> Светофор:
>>   коэффициенты (Solvency, SST, RoRC...): красный ниже Hard Limit или хуже Forecast на 10+ п.п.,
>>     жёлтый ниже Soft Limit или хуже Forecast на 5+ п.п.
>>   суммы (Remittance, VNP, Operating result...): отклонение от Forecast > 20% красный, > 10% жёлтый
>>     (в любую сторону: "надо объяснить")
>>   Info (ORMT, Sensitivities): серый, без оценки.
>>
>> ПОРЯДОК: ШАГ 1-2 Power Query -> 3 модель -> 4-5 таблицы -> 6-9 меры -> 10-12 страницы -> 13 проверка


#####################################################################
 ШАГ 1. FactData — добавить ключ строки (в конец запроса)
#####################################################################
>> Trasforma dati > FactData > Editor avanzato. Последний шаг у тебя WithSrc.
>> 1) В конце строки WithSrc поставь запятую.
>> 2) Замени последние две строки "in / WithSrc" на:

    WithKey  = Table.AddColumn(WithSrc, "LineKey", each
                   [Topic] & "|" & (if [Block] = null then "" else [Block]) & "|"
                   & (if [Item] = null then "" else [Item]), type text)
in
    WithKey


#####################################################################
 ШАГ 2. НОВЫЙ ЗАПРОС DimLine (строки отчёта в порядке Excel)  — load ON
#####################################################################
>> Nuova origine > Query vuota > Editor avanzato > вставить > имя DimLine.
>> Зачем: одна и та же строка ("Total Remittance - Cash view basis") стоит на разных листах
>> на разных местах. DimLine даёт каждой строке своё место для сортировки
>> (одинаковые названия различаются невидимым символом, на экране их не видно).

let
    Src  = Table.Group(FactData, {"Topic", "Block", "Item"},
               {{"LineOrder", each List.Min([RowNo]), Int64.Type}},
               GroupKind.Global, Comparer.OrdinalIgnoreCase),
    Key  = Table.AddColumn(Src, "LineKey", each [Topic] & "|" & [Block] & "|" & [Item], type text),
    ByIt = Table.Group(Key, {"Item"},
               {{"T", each Table.AddIndexColumn(Table.Sort(_, {{"LineOrder", Order.Ascending}}),
                                                "Dup", 0, 1, Int64.Type)}},
               GroupKind.Global, Comparer.OrdinalIgnoreCase),
    Exp  = Table.ExpandTableColumn(ByIt, "T", {"Topic", "Block", "LineOrder", "LineKey", "Dup"}),
    Line = Table.AddColumn(Exp, "Line", each [Item] & Text.Repeat(Character.FromNumber(8203), [Dup]), type text),
    BOrd = Table.Group(Line, {"Block"}, {{"BlockOrder", each List.Min([LineOrder]), Int64.Type}}),
    Join = Table.NestedJoin(Line, {"Block"}, BOrd, {"Block"}, "B", JoinKind.LeftOuter),
    Out  = Table.ExpandTableColumn(Join, "B", {"BlockOrder"})
in
    Out

>> Проверка: колонки Item, Topic, Block, LineOrder, LineKey, Dup, Line, BlockOrder.
>> Chiudi e applica.


#####################################################################
 ШАГ 3. МОДЕЛЬ
#####################################################################
>> Visualizzazione modello, связи (1 -> *, направление Singolo):
>>   DimLine[LineKey]   1 -> *  FactData[LineKey]
>>   DimPeriod[Period]  1 -> *  FactData[Period]      (если уже есть — ок)
>>   Если есть связь DimItem -> FactData: можно оставить, в матрицах DimItem не используем.
>>
>> Сортировка (Visualizzazione dati > выдели колонку > Ordina per colonna):
>>   DimLine[Line]     -> LineOrder
>>   DimLine[Block]    -> BlockOrder
>>   DimPeriod[Period] -> PeriodOrder
>>
>> Riepilogo "Non riepilogare": FactData[Year], FactData[RowNo], FactData[BlockNo],
>>   DimLine[LineOrder], DimLine[BlockOrder], DimLine[Dup].
>>
>> Вычисляемая колонка FactData (Strumenti tabella > Nuova colonna):
>> 1 = строка-коэффициент (%, п.п., лимиты, RoRC/RoCE). Нужна для формата и для правил.

RatioFlag =
VAR u = " " & UPPER ( FactData[Item] )
RETURN
IF (
    CONTAINSSTRING ( u, " RATIO" ) || CONTAINSSTRING ( u, "(%)" ) || CONTAINSSTRING ( u, "P.P." )
        || CONTAINSSTRING ( u, " LIMIT" ) || CONTAINSSTRING ( u, "CORRIDOR" )
        || CONTAINSSTRING ( u, "TAX RATE" ) || u = " RORC" || u = " ROCE",
    1, 0 )

>> Если какая-то строка-процент не попала (видно по формату 0 вместо %) — добавь её слово в формулу.


#####################################################################
 ШАГ 4. ТАБЛИЦА Thresholds  (Modellazione > Nuova tabella)
#####################################################################

Thresholds =
DATATABLE (
    "Key", STRING, "Area", STRING, "Rule", STRING, "Value", DOUBLE, "Unit", STRING, "Status", STRING,
    {
        { "RATIO_DEV_AMBER", "Ratios",  "Amber if Actual ratio is below Forecast by more than", 5,    "pp",    "Draft" },
        { "RATIO_DEV_RED",   "Ratios",  "Red if Actual ratio is below Forecast by more than",   10,   "pp",    "Draft" },
        { "AMT_DEV_AMBER",   "Amounts", "Amber if Actual differs from Forecast by more than",   0.10, "share", "Draft" },
        { "AMT_DEV_RED",     "Amounts", "Red if Actual differs from Forecast by more than",     0.20, "share", "Draft" }
    } )

>> Hard Limit / Soft Limit не здесь: они читаются из строк "Hard Limit" и "Soft Limit" каждого
>> Entity-листа (свои у каждой компании).


#####################################################################
 ШАГ 5. ТАБЛИЦА SummaryIndex  (Modellazione > Nuova tabella)
#####################################################################
>> Order    порядок на Summary
>> Area     группа (красный подзаголовок)
>> Section  название на Summary
>> Topic    ТОЧНО как Topic в BlockConfig
>> Page     ТОЧНОЕ имя вкладки Power BI (для кнопки)
>> KeyItem  начало текста ключевой строки (Item)
>> KeyBlock начало текста Block / Group / Section, если KeyItem встречается дважды ("" = не нужно)
>> Rule     RATIO_LIMIT (коэфф. + лимиты из файла) | RATIO | AMOUNT | INFO
>> KeyFig   TRUE = попадает в "KEY FIGURES" Executive Summary
>> ByCompany TRUE = показывать на Summary BU (на листе есть Company #)

SummaryIndex =
DATATABLE (
    "Order", INTEGER, "Area", STRING, "Section", STRING, "Topic", STRING, "Page", STRING,
    "KeyItem", STRING, "KeyBlock", STRING, "Rule", STRING, "KeyFig", BOOLEAN, "ByCompany", BOOLEAN,
    {
        {  1, "Distributions",      "Total Distributions BU", "Total Distributions BU",          "Total Distributions BU", "Total Remittance",               "",                                  "AMOUNT",      TRUE,  FALSE },
        {  2, "Solvency",           "Solvency II",            "Solvency II",                     "Solvency II",            "Solvency Ratio (%)",             "",                                  "RATIO_LIMIT", TRUE,  TRUE  },
        {  3, "Solvency",           "Local Solvency",         "Local Solvency",                  "Local Solvency",         "Solvency Ratio",                 "",                                  "RATIO_LIMIT", FALSE, TRUE  },
        {  4, "Solvency",           "SST",                    "SST",                             "SST",                    "SST Ratio",                      "",                                  "RATIO_LIMIT", FALSE, TRUE  },
        {  5, "Solvency",           "Other Sectoral",         "Other Sectoral",                  "Other Sectoral",         "Solvency Ratio alike",           "",                                  "RATIO",       FALSE, TRUE  },
        {  6, "Solvency",           "Solvency II AoM",        "Solvency II AoM",                 "Solvency II AoM",        "Total End of Period",            "Eligible Own Funds",                "AMOUNT",      FALSE, TRUE  },
        {  7, "Holding",            "Holding",                "Holding",                         "Holding",                "Total Remittance",               "",                                  "AMOUNT",      TRUE,  TRUE  },
        {  8, "Capital generation", "Own Funds Generation",   "Own Funds",                       "Own Funds",              "Life SII VNP",                   "",                                  "AMOUNT",      FALSE, TRUE  },
        {  9, "Profitability",      "BU Profitability",       "BU Profitability",                "BU Profitability",       "RoRC",                           "RoRC",                              "RATIO",       TRUE,  FALSE },
        { 10, "Profitability",      "LE Profitability",       "Profitability Ratios",            "LE Profitability",       "RoCE",                           "RoCE incl",                         "RATIO",       FALSE, TRUE  },
        { 11, "Accounting",         "Local GAAP",             "Local GAAP Total",                "Local GAAP",             "Gross Written Premiums",         "",                                  "AMOUNT",      FALSE, FALSE },
        { 12, "Accounting",         "Holding Local GAAP",     "Local LGAAP",                     "Holding LGAAP",          "Operating result",               "",                                  "AMOUNT",      FALSE, FALSE },
        { 13, "Accounting",         "IFRS",                   "IFRS",                            "IFRS",                   "P&C Operating Result",           "",                                  "AMOUNT",      FALSE, TRUE  },
        { 14, "Accounting",         "Non-Life LGAAP vs IFRS", "Non-Life between LGAAP and IFRS", "Non-Life LGAAP vs IFRS", "IFRS P&C Result gross of tax",   "",                                  "AMOUNT",      FALSE, FALSE },
        { 15, "Risk",               "ORMT",                   "ORMT",                            "ORMT",                   "Impact on Solvency Ratio (p.p.)", "",                                 "INFO",        FALSE, TRUE  },
        { 16, "Risk",               "Sensitivities",          "Sensitivities",                   "Sensitivities",          "+50 bps Interest Rates",         "Impact of Sensitivities on Local",  "INFO",        FALSE, TRUE  }
    } )

>> SummaryIndex[Section] > Ordina per colonna > Order.
>> KeyItem "Total End of Period" (AoM) и "Solvency Ratio" (Local Solvency) — по шаблону я их
>> не видела полностью: проверь в ШАГЕ 13 (страница Debug), при необходимости поменяй текст.
>> Новая страница = новая строка здесь.


#####################################################################
 ШАГ 6. БАЗОВЫЕ МЕРЫ  (Home > Nuova misura, папка "CPR")
#####################################################################
>> Если в новом файле есть старые Value / Value Current — оставь, CPR-меры от них не зависят.

CPR Value =
VAR v   = SUM ( FactData[Value] )
VAR isR = MAX ( FactData[RatioFlag] ) = 1
VAR nE  = DISTINCTCOUNT ( FactData[Entity] )
RETURN IF ( isR && nE > 1, BLANK (), v )

>> (коэффициенты нельзя складывать: если выбрано несколько компаний, строка-процент пустая)

CPR Year =
CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData ), FactData[PeriodType] = "Actual" )

Act CY =
VAR y = [CPR Year]
RETURN CALCULATE ( [CPR Value], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Actual", FactData[Year] = y )

Act PY =
VAR y = [CPR Year]
RETURN CALCULATE ( [CPR Value], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Actual", FactData[Year] = y - 1 )

Fcst CY =
VAR y = [CPR Year]
RETURN CALCULATE ( [CPR Value], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Forecast", FactData[Year] = y )

Plan NY =
VAR y = [CPR Year]
RETURN CALCULATE ( [CPR Value], REMOVEFILTERS ( DimPeriod ),
           FactData[PeriodType] = "Plan", FactData[Year] = y + 1 )

Var vs Fcst =
VAR a = [Act CY]
VAR f = [Fcst CY]
VAR isR = MAX ( FactData[RatioFlag] ) = 1
RETURN IF ( NOT ISBLANK ( a ) && NOT ISBLANK ( f ),
           IF ( isR, ( a - f ) * 100, DIVIDE ( a - f, ABS ( f ) ) ) )

Var vs PY =
VAR a = [Act CY]
VAR p = [Act PY]
VAR isR = MAX ( FactData[RatioFlag] ) = 1
RETURN IF ( NOT ISBLANK ( a ) && NOT ISBLANK ( p ),
           IF ( isR, ( a - p ) * 100, DIVIDE ( a - p, ABS ( p ) ) ) )

Delta vs Fcst = IF ( NOT ISBLANK ( [Act CY] ) && NOT ISBLANK ( [Fcst CY] ), [Act CY] - [Fcst CY] )

>> ФОРМАТ (выдели меру > Strumenti misura > Formato: Dinamico > слева в строке формул "Formato"):
>>   CPR Value, Act CY, Act PY, Fcst CY, Plan NY, Delta vs Fcst:
>>       IF ( MAX ( FactData[RatioFlag] ) = 1, "0.0%", "#,##0;(#,##0);-" )
>>   Var vs Fcst, Var vs PY:
>>       IF ( MAX ( FactData[RatioFlag] ) = 1, "+0.0 ""pp"";-0.0 ""pp"";0.0 ""pp""", "+0.0%;-0.0%;0.0%" )

CPR Subtitle =
VAR q  = LOOKUPVALUE ( CPR_Info[Value], CPR_Info[Field], "Quarter" )
VAR nm = LOOKUPVALUE ( CPR_Info[Value], CPR_Info[Field], "Entity Legal Name" )
VAR ent = IF ( ISFILTERED ( FactData[Entity] ), SELECTEDVALUE ( FactData[Entity], "Selected companies" ), "All companies" )
RETURN "Capital Performance Review · Actual " & [CPR Year]
       & IF ( q <> "", " " & q, "" ) & IF ( nm <> "", " · " & nm, "" ) & " · " & ent


#####################################################################
 ШАГ 7. КЛЮЧЕВАЯ СТРОКА СТРАНИЦЫ (по SummaryIndex)  — папка "CPR Key"
#####################################################################
>> Key Topic: на Summary = строка SummaryIndex, на странице = Topic из фильтра страницы.

Key Topic =
COALESCE ( SELECTEDVALUE ( SummaryIndex[Topic] ), SELECTEDVALUE ( DimLine[Topic] ), SELECTEDVALUE ( FactData[Topic] ) )

Key Act CY =
VAR t  = [Key Topic]
VAR k  = LOOKUPVALUE ( SummaryIndex[KeyItem], SummaryIndex[Topic], t )
VAR kb = LOOKUPVALUE ( SummaryIndex[KeyBlock], SummaryIndex[Topic], t )
RETURN IF ( NOT ISBLANK ( k ),
    CALCULATE ( [Act CY], FactData[Topic] = t, LEFT ( FactData[Item], LEN ( k ) ) = k,
        FILTER ( ALL ( FactData[Block], FactData[Group], FactData[Section] ),
            kb = "" || LEFT ( FactData[Block], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Group], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Section], LEN ( kb ) ) = kb ) ) )

Key Fcst CY =
VAR t  = [Key Topic]
VAR k  = LOOKUPVALUE ( SummaryIndex[KeyItem], SummaryIndex[Topic], t )
VAR kb = LOOKUPVALUE ( SummaryIndex[KeyBlock], SummaryIndex[Topic], t )
RETURN IF ( NOT ISBLANK ( k ),
    CALCULATE ( [Fcst CY], FactData[Topic] = t, LEFT ( FactData[Item], LEN ( k ) ) = k,
        FILTER ( ALL ( FactData[Block], FactData[Group], FactData[Section] ),
            kb = "" || LEFT ( FactData[Block], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Group], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Section], LEN ( kb ) ) = kb ) ) )

Key Act PY =
VAR t  = [Key Topic]
VAR k  = LOOKUPVALUE ( SummaryIndex[KeyItem], SummaryIndex[Topic], t )
VAR kb = LOOKUPVALUE ( SummaryIndex[KeyBlock], SummaryIndex[Topic], t )
RETURN IF ( NOT ISBLANK ( k ),
    CALCULATE ( [Act PY], FactData[Topic] = t, LEFT ( FactData[Item], LEN ( k ) ) = k,
        FILTER ( ALL ( FactData[Block], FactData[Group], FactData[Section] ),
            kb = "" || LEFT ( FactData[Block], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Group], LEN ( kb ) ) = kb
                    || LEFT ( FactData[Section], LEN ( kb ) ) = kb ) ) )

Key Hard =
VAR t = [Key Topic]
RETURN CALCULATE ( [Act CY], FactData[Topic] = t, FactData[Item] = "Hard Limit" )

Key Soft =
VAR t = [Key Topic]
RETURN CALCULATE ( [Act CY], FactData[Topic] = t, FactData[Item] = "Soft Limit" )

>> Светофор ОДНОЙ компании (или BU, если на листе нет Company #)
Key Status Ent =
VAR t    = [Key Topic]
VAR rule = LOOKUPVALUE ( SummaryIndex[Rule], SummaryIndex[Topic], t )
VAR isR  = rule IN { "RATIO", "RATIO_LIMIT" }
VAR a    = [Key Act CY]
VAR f    = [Key Fcst CY]
VAR hard = IF ( rule = "RATIO_LIMIT", [Key Hard] )
VAR soft = IF ( rule = "RATIO_LIMIT", [Key Soft] )
VAR rA   = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "RATIO_DEV_AMBER" )
VAR rR   = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "RATIO_DEV_RED" )
VAR aA   = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "AMT_DEV_AMBER" )
VAR aR   = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "AMT_DEV_RED" )
VAR d    = ( a - f ) * 100
VAR rel  = ABS ( DIVIDE ( a - f, ABS ( f ) ) )
VAR sLim = SWITCH ( TRUE (),
               NOT ISBLANK ( hard ) && hard > 0 && a < hard, 3,
               NOT ISBLANK ( soft ) && soft > 0 && a < soft, 2, 1 )
VAR sDev = SWITCH ( TRUE (),
               ISBLANK ( f ) || f = 0, 1,
               isR && d <= -rR, 3,
               isR && d <= -rA, 2,
               isR, 1,
               rel > aR, 3,
               rel > aA, 2, 1 )
RETURN SWITCH ( TRUE (),
    rule = "INFO", BLANK (),
    ISBLANK ( a ) || ( a = 0 && ( ISBLANK ( f ) || f = 0 ) ), BLANK (),
    MAX ( sLim, sDev ) )

>> Светофор раздела = худшая компания
CPR Status =
VAR t    = [Key Topic]
VAR rule = LOOKUPVALUE ( SummaryIndex[Rule], SummaryIndex[Topic], t )
VAR ents = CALCULATETABLE ( VALUES ( FactData[Entity] ), FactData[Topic] = t )
VAR hasF = NOT ISBLANK ( [Key Act CY] ) || NOT ISBLANK ( [Key Fcst CY] )
         || NOT ISEMPTY ( FILTER ( ents, NOT ISBLANK ( [Key Act CY] ) || NOT ISBLANK ( [Key Fcst CY] ) ) )
RETURN IF ( rule = "INFO", IF ( hasF, 0 ), MAXX ( ents, [Key Status Ent] ) )

>> Одна строка с цифрой. Коэффициент: худшая компания. Сумма: сумма компаний.
CPR Headline =
VAR t    = [Key Topic]
VAR k    = LOOKUPVALUE ( SummaryIndex[KeyItem], SummaryIndex[Topic], t )
VAR rule = LOOKUPVALUE ( SummaryIndex[Rule], SummaryIndex[Topic], t )
VAR isR  = rule IN { "RATIO", "RATIO_LIMIT" } || CONTAINSSTRING ( k, "p.p." )
VAR y    = [CPR Year]
VAR fm   = IF ( isR, "0.0%", "#,##0" )
VAR ents =
    FILTER (
        ADDCOLUMNS ( CALCULATETABLE ( VALUES ( FactData[Entity] ), FactData[Topic] = t ),
            "@a", [Key Act CY], "@f", [Key Fcst CY], "@p", [Key Act PY] ),
        NOT ISBLANK ( [@a] ) || NOT ISBLANK ( [@f] ) )
VAR n    = COUNTROWS ( ents )
VAR wRow = TOPN ( 1, ents, COALESCE ( [@a], [@f] ), ASC )
VAR wE   = MAXX ( wRow, FactData[Entity] )
VAR a    = IF ( isR, MAXX ( wRow, [@a] ), SUMX ( ents, [@a] ) )
VAR f    = IF ( isR, MAXX ( wRow, [@f] ), SUMX ( ents, [@f] ) )
VAR p    = IF ( isR, MAXX ( wRow, [@p] ), SUMX ( ents, [@p] ) )
VAR vsF  = IF ( NOT ISBLANK ( a ) && NOT ISBLANK ( f ),
               IF ( isR, FORMAT ( ( a - f ) * 100, "+0.0;-0.0;0.0" ) & " pp",
                         FORMAT ( DIVIDE ( a - f, ABS ( f ) ), "+0%;-0%;0%" ) ) & " vs Forecast" )
VAR vsP  = IF ( NOT ISBLANK ( a ) && NOT ISBLANK ( p ),
               IF ( isR, FORMAT ( ( a - p ) * 100, "+0.0;-0.0;0.0" ) & " pp",
                         FORMAT ( DIVIDE ( a - p, ABS ( p ) ), "+0%;-0%;0%" ) ) & " vs " & ( y - 1 ) )
VAR who  = IF ( n > 1, IF ( isR, " – lowest: " & wE, " (" & n & " companies)" ), "" )
RETURN SWITCH ( TRUE (),
    ISBLANK ( k ) || n = 0, BLANK (),
    ISBLANK ( a ), k & " Forecast " & y & ": " & FORMAT ( f, fm ) & who & "; Actual " & y & " not yet available.",
    k & " " & y & ": " & FORMAT ( a, fm ) & who
      & IF ( vsF <> "" || vsP <> "", " (" & vsF & IF ( vsF <> "" && vsP <> "", "; " ) & vsP & ")" ) & "." )


#####################################################################
 ШАГ 8. КОММЕНТАРИЙ ДЛЯ ЛЮБОЙ СТРАНИЦЫ  — папка "CPR Comment"
#####################################################################
>> Без выбора: ключевая строка + 3 самых больших отклонения сумм от Forecast
>> + 2 самых больших по коэффициентам + сколько строк ещё без Actual.
>> Клик по строке матрицы: эта строка — Actual, vs Forecast, vs прошлый год, Plan.

CPR Comment =
VAR nl  = UNICHAR ( 10 )
VAR y   = [CPR Year]
VAR nE  = DISTINCTCOUNT ( FactData[Entity] )
VAR who = IF ( nE = 1, COALESCE ( SELECTEDVALUE ( FactData[Entity] ), "Business Unit" ),
              nE & " companies (amounts summed, ratios only per company)" )
-- page mode
VAR L =
    FILTER (
        ADDCOLUMNS ( SUMMARIZE ( FactData, DimLine[Line], DimLine[Block] ),
            "@a", [Act CY], "@f", [Fcst CY], "@r", CALCULATE ( MAX ( FactData[RatioFlag] ) ) ),
        NOT ISBLANK ( [@a] ) || NOT ISBLANK ( [@f] ) )
VAR amt  = FILTER ( L, [@r] = 0 && NOT ISBLANK ( [@a] ) && NOT ISBLANK ( [@f] ) && [@a] <> [@f] )
VAR rat  = FILTER ( L, [@r] = 1 && NOT ISBLANK ( [@a] ) && NOT ISBLANK ( [@f] ) && [@a] <> [@f] )
VAR topA = TOPN ( 3, amt, ABS ( [@a] - [@f] ), DESC )
VAR topR = TOPN ( 2, rat, ABS ( [@a] - [@f] ), DESC )
VAR nNoA = COUNTROWS ( FILTER ( L, ISBLANK ( [@a] ) && NOT ISBLANK ( [@f] ) && [@f] <> 0 ) ) + 0
VAR head = [CPR Headline]
VAR txtA = CONCATENATEX ( topA,
               "- " & DimLine[Line] & ": " & FORMAT ( [@a], "#,##0" ) & " vs " & FORMAT ( [@f], "#,##0" )
               & " forecast (" & FORMAT ( [@a] - [@f], "+#,##0;-#,##0" )
               & IF ( [@f] <> 0, ", " & FORMAT ( DIVIDE ( [@a] - [@f], ABS ( [@f] ) ), "+0%;-0%" ), "" ) & ")",
               nl, ABS ( [@a] - [@f] ), DESC )
VAR txtR = CONCATENATEX ( topR,
               "- " & DimLine[Line] & ": " & FORMAT ( [@a], "0.0%" ) & " vs " & FORMAT ( [@f], "0.0%" )
               & " forecast (" & FORMAT ( ( [@a] - [@f] ) * 100, "+0.0;-0.0" ) & " pp)",
               nl, ABS ( [@a] - [@f] ), DESC )
VAR pageText =
    UPPER ( who ) & " – ACTUAL " & y & nl & nl
    & IF ( head <> "", head & nl & nl, "" )
    & IF ( COUNTROWS ( L ) = 0, "No figures reported yet.",
          IF ( NOT ISEMPTY ( topA ), "Largest deviations vs Forecast " & y & ":" & nl & txtA & nl & nl, "" )
          & IF ( NOT ISEMPTY ( topR ), "Ratios vs Forecast:" & nl & txtR & nl & nl, "" )
          & IF ( ISEMPTY ( topA ) && ISEMPTY ( topR ), "Actual " & y & " in line with Forecast on all lines." & nl, "" )
          & IF ( nNoA > 0, nNoA & " line(s) have a Forecast but no Actual " & y & " yet.", "" ) )
-- row mode
VAR ln  = SELECTEDVALUE ( DimLine[Line] )
VAR a   = [Act CY]
VAR f   = [Fcst CY]
VAR p   = [Act PY]
VAR pl  = [Plan NY]
VAR isR = MAX ( FactData[RatioFlag] ) = 1
VAR fm  = IF ( isR, "0.0%", "#,##0" )
VAR dF  = IF ( isR, FORMAT ( ( a - f ) * 100, "+0.0;-0.0;0.0" ) & " pp",
                    FORMAT ( a - f, "+#,##0;-#,##0;0" )
                    & IF ( f <> 0, " (" & FORMAT ( DIVIDE ( a - f, ABS ( f ) ), "+0%;-0%;0%" ) & ")", "" ) )
VAR dP  = IF ( isR, FORMAT ( ( a - p ) * 100, "+0.0;-0.0;0.0" ) & " pp",
                    FORMAT ( a - p, "+#,##0;-#,##0;0" )
                    & IF ( p <> 0, " (" & FORMAT ( DIVIDE ( a - p, ABS ( p ) ), "+0%;-0%;0%" ) & ")", "" ) )
VAR rowText =
    UPPER ( ln ) & " – " & who & nl & nl
    & IF ( ISBLANK ( a ), "No Actual " & y & " yet.",
          "Actual " & y & ": " & FORMAT ( a, fm )
          & IF ( NOT ISBLANK ( f ), nl & "vs Forecast " & y & ": " & dF, "" )
          & IF ( NOT ISBLANK ( p ), nl & "vs Actual " & ( y - 1 ) & ": " & dP, "" ) )
    & IF ( ISBLANK ( a ) && NOT ISBLANK ( f ), nl & "Forecast " & y & ": " & FORMAT ( f, fm ), "" )
    & IF ( NOT ISBLANK ( pl ), nl & "Plan " & ( y + 1 ) & ": " & FORMAT ( pl, fm ), "" )
RETURN IF ( HASONEVALUE ( DimLine[Line] ), rowText, pageText )


#####################################################################
 ШАГ 9. МЕРЫ SUMMARY  — папка "Summary"
#####################################################################

Summary Dot = IF ( NOT ISBLANK ( [CPR Headline] ), UNICHAR ( 11044 ) )

Summary Color =
SWITCH ( [CPR Status], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", 0, "#9E9E9E", "#D0D0D0" )

Summary Label =
VAR s = [CPR Status]
RETURN SWITCH ( TRUE (),
    s = 3, "Action needed", s = 2, "To explain", s = 1, "In line", s = 0, "Info",
    NOT ISBLANK ( [CPR Headline] ), "No Actual yet" )

Summary Nav = SELECTEDVALUE ( SummaryIndex[Page], "" )

Summary Button Text =
IF ( HASONEVALUE ( SummaryIndex[Section] ), "Open " & SELECTEDVALUE ( SummaryIndex[Section] ) & "  →",
     "Select a line to open its page" )

>> ---- EXECUTIVE SUMMARY (страница Summary; на Summary BU — для выбранной компании) ----
Executive Summary =
VAR nl  = UNICHAR ( 10 )
VAR b   = UNICHAR ( 8226 ) & " "
VAR y   = [CPR Year]
VAR who = IF ( ISFILTERED ( FactData[Entity] ), SELECTEDVALUE ( FactData[Entity], "Selected companies" ), "Business Unit" )
VAR t   = ADDCOLUMNS ( SummaryIndex, "@s", [CPR Status], "@txt", [CPR Headline] )
VAR tx  = FILTER ( t, NOT ISBLANK ( [@txt] ) )
VAR nG  = COUNTROWS ( FILTER ( tx, [@s] = 1 ) ) + 0
VAR nA  = COUNTROWS ( FILTER ( tx, [@s] = 2 ) ) + 0
VAR nR  = COUNTROWS ( FILTER ( tx, [@s] = 3 ) ) + 0
VAR nRt = nG + nA + nR
VAR verdict = SWITCH ( TRUE (),
    nRt = 0, "no Actual " & y & " figures to assess yet",
    nR = 0 && nA = 0, "Actual " & y & " in line with Forecast and limits",
    nR = 0, "broadly in line, " & nA & " area(s) to explain",
    nR & " area(s) need action, " & nA & " to explain" )
VAR kTxt = CONCATENATEX ( FILTER ( tx, SummaryIndex[KeyFig] ), b & [@txt], nl, SummaryIndex[Order], ASC )
VAR rTxt = CONCATENATEX ( FILTER ( tx, [@s] = 3 ), b & SummaryIndex[Section] & ": " & [@txt], nl, SummaryIndex[Order], ASC )
VAR aTxt = CONCATENATEX ( FILTER ( tx, [@s] = 2 ), b & SummaryIndex[Section] & ": " & [@txt], nl, SummaryIndex[Order], ASC )
VAR dev =
    FILTER (
        ADDCOLUMNS ( SUMMARIZE ( FactData, FactData[Topic], DimLine[Line] ),
            "@a", [Act CY], "@f", [Fcst CY], "@r", CALCULATE ( MAX ( FactData[RatioFlag] ) ) ),
        [@r] = 0 && NOT ISBLANK ( [@a] ) && NOT ISBLANK ( [@f] ) && [@f] <> 0 && [@a] <> [@f] )
VAR top3 = TOPN ( 3, dev, ABS ( [@a] - [@f] ), DESC )
VAR dTxt = CONCATENATEX ( top3,
               b & FactData[Topic] & " – " & DimLine[Line] & ": " & FORMAT ( [@a] - [@f], "+#,##0;-#,##0" )
               & " (" & FORMAT ( DIVIDE ( [@a] - [@f], ABS ( [@f] ) ), "+0%;-0%" ) & ") vs Forecast",
               nl, ABS ( [@a] - [@f] ), DESC )
VAR miss = CONCATENATEX ( FILTER ( t, ISBLANK ( [@s] ) && SummaryIndex[Rule] <> "INFO" ),
               SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
RETURN
who & " – Capital Performance Review " & y & ": " & verdict
& IF ( nRt > 0, " (" & nG & " of " & nRt & " areas in line)", "" ) & "."
& IF ( kTxt <> "", nl & nl & "KEY FIGURES" & nl & kTxt, "" )
& IF ( nR > 0, nl & nl & "ACTION NEEDED" & nl & rTxt, "" )
& IF ( nA > 0, nl & nl & "TO EXPLAIN" & nl & aTxt, "" )
& IF ( dTxt <> "", nl & nl & "LARGEST DEVIATIONS vs FORECAST" & nl & dTxt, "" )
& IF ( miss <> "", nl & nl & "NO ACTUAL " & y & " YET" & nl & b & miss, "" )

>> ---- ПО КОМПАНИЯМ (страница Summary BU) ----
BU Worst =
MAXX (
    FILTER ( ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] ), "@s", [CPR Status] ),
             [@s] IN { 1, 2, 3 } ),
    [@s] )

BU Line =
VAR rows = ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] ), "@s", [CPR Status] )
VAR n  = COUNTROWS ( FILTER ( rows, [@s] IN { 1, 2, 3 } ) ) + 0
VAR g  = COUNTROWS ( FILTER ( rows, [@s] = 1 ) ) + 0
VAR rl = CONCATENATEX ( FILTER ( rows, [@s] = 3 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR al = CONCATENATEX ( FILTER ( rows, [@s] = 2 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR sr = CALCULATE ( [Key Act CY], SummaryIndex[Topic] = "Solvency II" )
RETURN IF ( n > 0,
    IF ( NOT ISBLANK ( sr ), "SII ratio " & FORMAT ( sr, "0%" ) & " · ", "" )
    & g & " of " & n & " in line"
    & IF ( rl <> "", " · Action: " & rl, "" )
    & IF ( al <> "", " · Explain: " & al, "" ) )

BU Dot = IF ( NOT ISBLANK ( [BU Worst] ), UNICHAR ( 11044 ) )

BU Color = SWITCH ( [BU Worst], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", "#9E9E9E" )

BU Executive Summary =
VAR nl  = UNICHAR ( 10 )
VAR b   = UNICHAR ( 8226 ) & " "
VAR one = ISFILTERED ( FactData[Entity] ) && HASONEVALUE ( FactData[Entity] )
VAR companyText = CALCULATE ( [Executive Summary], SummaryIndex[ByCompany] = TRUE () )
VAR ents =
    FILTER (
        ADDCOLUMNS ( CALCULATETABLE ( VALUES ( FactData[Entity] ), NOT ISBLANK ( FactData[Entity] ) ),
            "@w", [BU Worst], "@line", [BU Line] ),
        NOT ISBLANK ( [@w] ) )
VAR cN = COUNTROWS ( ents ) + 0
VAR cR = COUNTROWS ( FILTER ( ents, [@w] = 3 ) ) + 0
VAR cA = COUNTROWS ( FILTER ( ents, [@w] = 2 ) ) + 0
VAR cG = COUNTROWS ( FILTER ( ents, [@w] = 1 ) ) + 0
VAR discuss = CONCATENATEX ( FILTER ( ents, [@w] >= 2 ), b & FactData[Entity] & ": " & [@line], nl,
                  [@w], DESC, FactData[Entity], ASC )
VAR okList  = CONCATENATEX ( FILTER ( ents, [@w] = 1 ), FactData[Entity], ", ", FactData[Entity], ASC )
VAR portfolio =
    IF ( cN = 0, "No company figures for Actual " & [CPR Year] & " yet.",
        cN & " companies: " & cR & " need action, " & cA & " to explain, " & cG & " in line."
        & IF ( discuss <> "", nl & nl & "COMPANIES TO DISCUSS" & nl & discuss, "" )
        & IF ( okList <> "", nl & nl & "IN LINE" & nl & b & okList, "" )
        & nl & nl & "Select a company for its summary." )
RETURN IF ( one, companyText, portfolio )


#####################################################################
 ШАГ 10. ШАБЛОН СТРАНИЦЫ (сделай одну, потом Duplica pagina)
#####################################################################
>> Сетка как в CDD (Design_Guide): шапка 0,0 1280x56 | меню слева | основной блок | комментарий справа.
>>
>> 1) Фильтр страницы: Filtri in questa pagina > DimLine[Topic] = нужный Topic (см. ШАГ 11).
>>    ВАЖНО: фильтр именно по DimLine[Topic], не по FactData — иначе в матрице появятся строки других листов.
>> 2) Подзаголовок: Scheda (nuova) с мерой CPR Subtitle.
>> 3) Слайсер FactData[Entity] (Elenco a discesa, Selezione singola Off) — только где есть Company #.
>>    Фильтр слайсера: Entity "non è vuoto".
>> 4) МАТРИЦА A "как в Excel":
>>      Righe:   DimLine[Block] > DimLine[Line]
>>      Colonne: DimPeriod[Period]
>>      Valori:  CPR Value
>>      Стрелка у Line в Righe > "Mostra elementi senza dati" (структура видна даже без цифр)
>>      Layout: Struttura; Subtotali righe/colonne: Off; "Espandi tutto" (иконка двойной стрелки вниз).
>> 5) МАТРИЦА B "Performance" справа от A (те же строки):
>>      Righe:   DimLine[Block] > DimLine[Line]
>>      Valori:  Act CY | Fcst CY | Delta vs Fcst | Var vs Fcst | Var vs PY
>>      Переименуй поля в визуале: "Actual", "Forecast", "Δ vs Fcst", "Δ vs Fcst %/pp", "Δ vs PY %/pp".
>>      Цвет: Var vs Fcst > Formattazione condizionale > Colore carattere > Regole: <0 #A50E0E, >0 #2E7D32.
>>      (Совет: в A и B одинаковая высота строк, тогда строки идут ровно рядом.)
>> 6) КОММЕНТАРИЙ "Key messages": Scheda/Tabella с мерой CPR Comment, A capo automatico On.
>>    Клик по строке в матрице A или B меняет комментарий на эту строку.
>> 7) Кнопка "← Summary": Inserisci > Pulsante > Indietro, или Navigazione pagina > Summary.


#####################################################################
 ШАГ 11. СТРАНИЦЫ — ЧТО МЕНЯТЬ НА КАЖДОЙ
#####################################################################
>>  Вкладка (имя = SummaryIndex[Page])   DimLine[Topic] в фильтре страницы            Особенности
>>  -----------------------------------  -------------------------------------------  ------------------------------------
>>  Summary                              (без фильтра)                                ШАГ 12
>>  Summary BU                           (без фильтра)                                ШАГ 12
>>  Total Distributions BU               Total Distributions BU                       без слайсера Entity; в A виден Plan 2027
>>  Solvency II                          Solvency II                                  + карточка CPR Headline над матрицей
>>  Local Solvency                       Local Solvency                               + карточка CPR Headline
>>  SST                                  SST                                          + карточка CPR Headline
>>  Other Sectoral                       Other Sectoral
>>  Solvency II AoM                      Solvency II AoM                              A: Colonne = FactData[Group] > DimPeriod[Period];
>>                                                                                    B: Colonne = FactData[Group]
>>  Holding                              Holding
>>  Own Funds                            Own Funds
>>  BU Profitability                     BU Profitability                             без слайсера Entity
>>  LE Profitability                     Profitability Ratios
>>  Local GAAP                           Local GAAP Total, Local GAAP Life,            Righe: DimLine[Topic] > DimLine[Line]
>>                                       Local GAAP Non-Life                          (Topic = таблица), без слайсера
>>  Holding LGAAP                        Local LGAAP                                  без слайсера
>>  IFRS                                 IFRS
>>  Non-Life LGAAP vs IFRS               Non-Life between LGAAP and IFRS,             Righe: DimLine[Topic] > DimLine[Line]
>>                                       Non-Life TF by Company                       слайсер Entity влияет только на правую таблицу
>>  ORMT                                 ORMT                                         Righe: FactData[Section] > DimLine[Line];
>>                                                                                    "Mostra elementi senza dati" OFF
>>  Sensitivities                        Sensitivities                                только матрица A (там один период Forecast);
>>                                                                                    матрицу B не ставить
>>
>> Имена вкладок должны ТОЧНО совпасть с SummaryIndex[Page], иначе кнопка на Summary не сработает.
>> Скрыть служебные таблицы: Visualizzazione modello > правый клик > Nascondi: CPR_Info, Thresholds,
>> FactData[LineKey], DimLine[LineKey], DimLine[Dup].


#####################################################################
 ШАГ 12. SUMMARY-СТРАНИЦЫ
#####################################################################
>> ---- Summary (первая страница) ----
>> 1) Подзаголовок CPR Subtitle.
>> 2) Executive Summary: Scheda/Tabella с мерой Executive Summary, A capo On, ~760 x 300 сверху.
>> 3) Таблица разделов под ней:
>>      Colonne: SummaryIndex[Area] | Summary Dot | SummaryIndex[Section] | CPR Headline | Summary Label
>>      Summary Dot: Formattazione condizionale > Colore carattere > Valore campo > Summary Color
>>      Summary Label: тот же цвет (Summary Color).
>>      Сортировка: по SummaryIndex[Order] (добавь Order и сузь колонку до минимума, или Ordina per colonna).
>> 4) Кнопка перехода: слайсер SummaryIndex[Section] (Elenco a discesa, Selezione singola) +
>>    Inserisci > Pulsante > Vuoto > Azione: Navigazione pagina > Destinazione: fx > Summary Nav;
>>    Testo pulsante: fx > Summary Button Text.
>>    Слайсер НЕ должен фильтровать Executive Summary: выдели слайсер > Formato > Modifica interazioni >
>>    на Executive Summary "Nessuna".
>>
>> ---- Summary BU ----
>> 1) Слайсер FactData[Entity] (Elenco, Selezione singola On), фильтр Entity "non è vuoto".
>> 2) BU Executive Summary: Scheda/Tabella справа, A capo On.
>> 3) "Companies at a glance": Tabella  FactData[Entity] | BU Dot | BU Line
>>    BU Dot: Colore carattere > Valore campo > BU Color. Фильтр визуала: BU Dot "non è vuoto".
>> 4) Тепловая карта: Matrice
>>      Righe: FactData[Entity]   Colonne: SummaryIndex[Section]   Valori: Summary Dot
>>      Colore carattere > Valore campo > Summary Color; Totali Off; Intestazioni colonna: A capo On.
>>      Фильтры визуала: SummaryIndex[ByCompany] = True; FactData[Entity] "non è vuoto".
>> 5) Под картой таблица как на Summary (Area | Dot | Section | CPR Headline | Label),
>>    фильтр визуала ByCompany = True — при выборе компании показывает её строки.
>>
>> (HTML-вид из CDD подходит тоже: в Summary HTML замени [Summary Status] на [CPR Status]
>>  и [Summary Text] на [CPR Headline].)


#####################################################################
 ШАГ 13. ПРОВЕРКА НА DUMMY DATA
#####################################################################
>> Сделай скрытую страницу "Debug" с таблицей:
>>   SummaryIndex[Section] | SummaryIndex[KeyItem] | Key Act CY | Key Fcst CY | Key Act PY |
>>   Key Hard | Key Soft | CPR Status | CPR Headline
>>
>> [ ] Key Act CY и Key Fcst CY пустые у раздела, где в Excel цифры ЕСТЬ -> неверный KeyItem/KeyBlock:
>>     найди точный текст строки в FactData[Item] и поправь SummaryIndex.
>> [ ] Солвенси: если Key Act CY = 1,85 -> формат 185% (верно). Если 185 -> коэффициенты в файле
>>     хранятся в процентах, а не долях: скажи — поменяю *100 и форматы.
>> [ ] Hard/Soft пустые на Solvency II/Local Solvency/SST -> строки называются иначе
>>     ("Hard limit"?) — поправь текст в мерах Key Hard / Key Soft.
>> [ ] Для теста светофора поставь в dummy: Actual SII ratio на 6 п.п. ниже Forecast -> жёлтый;
>>     ниже Soft Limit -> жёлтый; ниже Hard Limit -> красный; Remittance +25% к Forecast -> красный.
>> [ ] Матрица A: строки в порядке Excel, колонки Actual 2025 | Forecast 2026 | Actual 2026.
>> [ ] Матрица B: Var vs Fcst у коэффициентов в pp, у сумм в %.
>> [ ] Несколько компаний в слайсере: строки-проценты пустые (так задумано), суммы складываются.
>> [ ] Executive Summary: KEY FIGURES = Total Distributions BU, Solvency II, Holding, BU Profitability.
>> [ ] Клик по строке матрицы -> комментарий по строке.
>> [ ] Ошибка в формуле -> фото жёлтой строки под формулой.
