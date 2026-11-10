=====================================================================
 
=====================================================================


#####################################################################
 ШАГ 1. ОШИБКА BU Line ("La sintassi per 'rows' non è corretta")
#####################################################################
>> ROWS — зарезервированное слово DAX (как раньше top, ln, avg). Переименовала в tR.
>> BU Executive Summary заработает сам: его ошибка только из-за BU Line.
>> Замени формулу BU Line целиком:

BU Line =
VAR tR = ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] ), "@s", [CPR Status] )
VAR n  = COUNTROWS ( FILTER ( tR, [@s] IN { 1, 2, 3 } ) ) + 0
VAR g  = COUNTROWS ( FILTER ( tR, [@s] = 1 ) ) + 0
VAR rl = CONCATENATEX ( FILTER ( tR, [@s] = 3 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR al = CONCATENATEX ( FILTER ( tR, [@s] = 2 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR sr = CALCULATE ( [Key Act CY], SummaryIndex[Topic] = "Solvency II" )
RETURN IF ( n > 0,
    IF ( NOT ISBLANK ( sr ), "SII ratio " & FORMAT ( sr, "0%" ) & " · ", "" )
    & g & " of " & n & " in line"
    & IF ( rl <> "", " · Action: " & rl, "" )
    & IF ( al <> "", " · Explain: " & al, "" ) )


#####################################################################
 ШАГ 2. ПОРЯДОК ПЕРИОДОВ И СТРОК (сейчас Actual 2026 стоит раньше Forecast,
        а строки идут по алфавиту)
#####################################################################
>> 2a) FactData > Nuova colonna:

PeriodSort =
FactData[Year] * 10
    + SWITCH ( FactData[PeriodType], "Forecast", 1, "Half-Year", 2, "Actual", 3, "Plan", 6, 9 )

>> 2b) Visualizzazione dati > FactData, выделить колонку > Strumenti colonna > Ordina per colonna:
>>       Period -> PeriodSort
>>       Block  -> BlockOrder
>>       Line   -> LineOrder
>>     (Если Line/LineOrder/BlockOrder нет в FactData — сначала шаг из прошлого сообщения
>>      про хвост FactData в Power Query.)
>>
>> 2c) Во ВСЕХ матрицах: Righe = FactData[Block] > FactData[Line]   (не Item!)
>>     Colonne = FactData[Period].
>>     На твоём фото LE Profitability в строках стоит Item — поэтому алфавит.


#####################################################################
 ШАГ 3. ФИЛЬТРЫ — проверка
#####################################################################
>> Фильтр страницы FactData[Topic] = "Profitability Ratios" на LE-странице — правильно.
>> Замена DimLine[Topic] -> FactData[Topic] в мерах — тоже правильно.
>> Слайсер Entity: ок. На BU-страницах (Total Distributions BU, Profitability Ratios BU,
>> Local GAAP Details, Holding LGAAP) слайсер не нужен — там нет Company #.
>>
>> Вкладки у тебя названы как листы Excel, поэтому кнопка с Summary не найдёт страницы.
>> Замени SummaryIndex целиком (Page = имена ТВОИХ вкладок). Проверь последние четыре по
>> своим вкладкам: ORMT Impacts, Solvency AoM, OF Capital Generation Items, Sensitivities.

SummaryIndex =
DATATABLE (
    "Order", INTEGER, "Area", STRING, "Section", STRING, "Topic", STRING, "Page", STRING,
    "KeyItem", STRING, "KeyBlock", STRING, "Rule", STRING, "KeyFig", BOOLEAN, "ByCompany", BOOLEAN,
    {
        {  1, "Distributions",      "Total Distributions BU", "Total Distributions BU",          "Total Distributions BU",         "Total Remittance",                "",                                 "AMOUNT",      TRUE,  FALSE },
        {  2, "Solvency",           "Solvency II",            "Solvency II",                     "Solvency II Entity",             "Solvency Ratio (%)",              "",                                 "RATIO_LIMIT", TRUE,  TRUE  },
        {  3, "Solvency",           "Local Solvency",         "Local Solvency",                  "Local Solvency Entity",          "Solvency Ratio",                  "",                                 "RATIO_LIMIT", FALSE, TRUE  },
        {  4, "Solvency",           "SST",                    "SST",                             "SST Entity",                     "SST Ratio",                       "",                                 "RATIO_LIMIT", FALSE, TRUE  },
        {  5, "Solvency",           "Other Sectoral",         "Other Sectoral",                  "Other Sect. Solvency Entity",    "Solvency Ratio alike",            "",                                 "RATIO",       FALSE, TRUE  },
        {  6, "Solvency",           "Solvency II AoM",        "Solvency II AoM",                 "Solvency AoM",                   "Total End of Period",             "Eligible Own Funds",               "AMOUNT",      FALSE, TRUE  },
        {  7, "Holding",            "Holding",                "Holding",                         "Holding Entity",                 "Total Remittance",                "",                                 "AMOUNT",      TRUE,  TRUE  },
        {  8, "Capital generation", "Own Funds Generation",   "Own Funds",                       "OF Capital Generation Items",    "Life SII VNP",                    "",                                 "AMOUNT",      FALSE, TRUE  },
        {  9, "Profitability",      "BU Profitability",       "BU Profitability",                "Profitability Ratios BU",        "RoRC",                            "RoRC",                             "RATIO",       TRUE,  FALSE },
        { 10, "Profitability",      "LE Profitability",       "Profitability Ratios",            "Profitability Ratios LEs",       "RoCE",                            "RoCE incl",                        "RATIO",       FALSE, TRUE  },
        { 11, "Accounting",         "Local GAAP",             "Local GAAP Total",                "Local GAAP Details",             "Gross Written Premiums",          "",                                 "AMOUNT",      FALSE, FALSE },
        { 12, "Accounting",         "Holding Local GAAP",     "Local LGAAP",                     "Holding LGAAP",                  "Operating result",                "",                                 "AMOUNT",      FALSE, FALSE },
        { 13, "Accounting",         "IFRS",                   "IFRS",                            "IFRS Details",                   "P&C Operating Result",            "",                                 "AMOUNT",      FALSE, TRUE  },
        { 14, "Accounting",         "Non-Life LGAAP vs IFRS", "Non-Life between LGAAP and IFRS", "Non-Life Translation Framework", "IFRS P&C Result gross of tax",    "",                                 "AMOUNT",      FALSE, FALSE },
        { 15, "Risk",               "ORMT",                   "ORMT",                            "ORMT Impacts",                   "Impact on Solvency Ratio (p.p.)", "",                                 "INFO",        FALSE, TRUE  },
        { 16, "Risk",               "Sensitivities",          "Sensitivities",                   "Sensitivities",                  "+50 bps Interest Rates",          "Impact of Sensitivities on Local", "INFO",        FALSE, TRUE  }
    } )


#####################################################################
 ШАГ 4. HTML-ТАБЛИЦА "КАК В EXCEL"  (визуал HTML Content)
#####################################################################
>> Одна мера на все страницы: берёт Topic из фильтра страницы и компанию из слайсера.
>> Вид: название блока тёмно-красным, под ним красная линия, периоды справа красным,
>> строки "o/w" с отступом, Total/Ratio жирные, колонка Actual (текущий год) подсвечена,
>> пусто = "-", отрицательные в скобках.

CPR Table HTML =
VAR yA   = [CPR Year]
VAR pers = SUMMARIZE ( FactData, FactData[Period], FactData[PeriodSort], FactData[PeriodType], FactData[Year] )
VAR nP   = COUNTROWS ( pers )
VAR thS  = "padding:4px 8px;text-align:right;color:#A50E0E;font-weight:600;white-space:nowrap;border-bottom:2px solid #C21B17"
VAR hdr  = CONCATENATEX ( pers,
               "<th style='" & thS
               & IF ( FactData[PeriodType] = "Actual" && FactData[Year] = yA, ";background:#F2E3E2", "" )
               & "'>" & FactData[Period] & "</th>",
               "", FactData[PeriodSort], ASC )
VAR blocks = SUMMARIZE ( FactData, FactData[Block], FactData[BlockOrder] )
VAR body =
    CONCATENATEX ( blocks,
        VAR bl    = FactData[Block]
        VAR lines = CALCULATETABLE (
                        SUMMARIZE ( FactData, FactData[Line], FactData[LineOrder], FactData[RatioFlag] ),
                        FactData[Block] = bl )
        RETURN
        "<table style='border-collapse:collapse;width:100%;margin:0 0 18px 0'>"
        & "<tr><th style='text-align:left;padding:4px 8px;color:#A50E0E;font-weight:700;border-bottom:2px solid #C21B17'>"
        & bl & "</th>" & hdr & "</tr>"
        & CONCATENATEX ( lines,
            VAR ln   = FactData[Line]
            VAR rf   = FactData[RatioFlag]
            VAR u    = UPPER ( TRIM ( ln ) )
            VAR bold = LEFT ( u, 5 ) = "TOTAL" || ( rf = 1 && CONTAINSSTRING ( u, "RATIO" ) )
                       || u IN { "RORC", "ROCE" }
            VAR ind  = IF ( LEFT ( u, 3 ) = "O/W" || LEFT ( u, 1 ) = "<", "26px", "8px" )
            VAR cells =
                CONCATENATEX ( pers,
                    VAR p  = FactData[Period]
                    VAR hi = FactData[PeriodType] = "Actual" && FactData[Year] = yA
                    VAR v  = CALCULATE ( [CPR Value], FactData[Block] = bl, FactData[Line] = ln, FactData[Period] = p )
                    VAR tx = IF ( ISBLANK ( v ), "-",
                                 IF ( rf = 1, FORMAT ( v, "0.0%" ), FORMAT ( v, "#,##0;(#,##0);0" ) ) )
                    RETURN "<td style='padding:3px 8px;text-align:right;border-bottom:1px solid #E3E3E3;white-space:nowrap"
                           & IF ( hi, ";background:#FBF3F2", "" )
                           & IF ( NOT ISBLANK ( v ) && v < 0, ";color:#A50E0E", "" ) & "'>" & tx & "</td>",
                    "", FactData[PeriodSort], ASC )
            RETURN
            "<tr" & IF ( bold, " style='background:#F7F7F7'", "" ) & ">"
            & "<td style='padding:3px 8px 3px " & ind & ";border-bottom:1px solid #E3E3E3"
            & IF ( bold, ";font-weight:700", "" ) & "'>" & ln & "</td>" & cells & "</tr>",
            "", FactData[LineOrder], ASC )
        & "</table>",
        "", FactData[BlockOrder], ASC )
VAR who = IF ( HASONEVALUE ( FactData[Entity] ) && NOT ISBLANK ( SELECTEDVALUE ( FactData[Entity] ) ),
              "<div style='color:#A50E0E;font-weight:700;font-size:13px;margin:0 0 8px 0'>"
              & SELECTEDVALUE ( FactData[Entity] ) & "</div>", "" )
RETURN
IF ( nP = 0, BLANK (),
    "<div style='font-family:Segoe UI;font-size:11px;color:#333'>" & who & body & "</div>" )

>> Поставить на страницу:
>>   Visualizzazioni > HTML Content (из AppSource, как в CDD) > Values = CPR Table HTML.
>>   Размер: основной блок (как матрица A), Titolo off, Bordo off, фон белый.
>>   Над ним текстбокс с названием листа (например "Business Unit – Total Distributions")
>>   и Scheda с CPR Subtitle.
>>   Матрицу B (Act CY / Fcst CY / Var ...) оставь рядом — HTML-таблица не кликабельна,
>>   а по клику на строку матрицы B меняется комментарий.
>>   Если клик по матрице B фильтрует и HTML-таблицу (она станет из одной строки):
>>   выдели матрицу B > Formato > Modifica interazioni > на HTML-визуале "Nessuna".
>>
>> Комментарий тоже в HTML (пункты списком):

CPR Comment HTML =
"<div style='font-family:Segoe UI;font-size:11px;line-height:1.45;color:#333'>"
& SUBSTITUTE ( [CPR Comment], UNICHAR ( 10 ), "<br>" )
& "</div>"

>> ЕСЛИ ОШИБКА "FactData[BlockOrder] не найден" -> в FactData нет колонок из хвоста Power Query
>> (LineOrder / BlockOrder / Line). Пришли фото списка колонок FactData.
>> ЕСЛИ HTML пустой при выбранных нескольких компаниях — нормально только для процентов;
>> суммы должны быть.
