=====================================================================
 SAA — сборка страницы
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> КАК УСТРОЕН ЛИСТ SAA (по фото), для каждой компании:
>>
>> A) ALLOCATION (строки 8-22): классы активов x
>>    Act 2025 MV | Expected 2026 MV | Net Balance 2027 (Redemptions, Purchases,
>>    Sales, Net Balance) | SAA 2027 (Target MV, Target %, Lower, Upper)
>>    + Total MV, Duration Fixed Income, Duration Gap
>>    У КАЖДОЙ компании две такие таблицы (DebugSAA: блоки 1-2, 4-5, 7-8).
>>    Читаем: первая = Life, вторая = Non-Life (у Life-компании вторая пустая).
>> B) ТАБЛИЦА СПРАВА (AC-AG): Sales & Redemptions, Current Yield %,
>>    Total Return Yield %, Current Income, Total Return.
>>    Парсер читает её вместе с блоком (Period = SAA 2027). Доработка не нужна.
>> C) SCR ABSORPTION (строки 41-53): Net SCR 2026 | SCR Purchases 2027 |
>>    SCR Sales & Redemptions 2027. Парсер v4 читает (двойной заголовок).
>> D) ПАРАМЕТРЫ (55-57): BU Return on Risk Capital, PH's Participation %,
>>    Nominal Tax Rate %.
>>
>> ВАЖНО: в dummy-файле лист SAA ПУСТОЙ (жёлтые ячейки без цифр).
>> Как с HMS: сохрани копию, внеси тестовые цифры для Company #1 (обе таблицы,
>> 4-5 классов + Target %, Lower, Upper — один класс специально вне границ,
>> + пара цифр в SCR и Current Yield), направь источник на копию, Aggiorna.


>> ШАГ 0. ПРОВЕРКА ПАРСЕРА (CheckBlocks)
>> Trasforma dati > CheckBlocks > строка SAA. Ожидаю:
>>   "... rows | N blocks: Cash & Cash Equivalent"  (имя одно на все блоки — это нормально)
>>   N = на каждую компанию 1-2 allocation + 1 SCR (Company #1: 3, #2: 2 ...).
>> Если ERROR или N странное — пришли фото этой строки.


>> ШАГ 1-2. ПРОПУСТИТЬ. DebugSAA показал: у allocation-блоков 15 столбцов
>> (10 основных + 5 справа: Sales & Redemptions, Current Yield, Total Return
>> Yield, Current Income, Total Return), т.е. парсер уже читает таблицу справа.
>> fnSAAYields не нужна — она бы задвоила эти цифры.
>> Если уже добавила её в FactData — верни конец запроса к  "in <твой шаг>".


>> ШАГ 3. КОЛОНКИ в FactData (Visualizzazione tabella > FactData > Nuova colonna), по порядку

SAA Table =
IF ( FactData[Topic] = "SAA",
     SWITCH ( TRUE (),
         CONTAINSSTRING ( FactData[Item], "Return on Risk" )
             || CONTAINSSTRING ( FactData[Item], "Participation" )
             || CONTAINSSTRING ( FactData[Item], "Tax Rate" ), "Parameters",
         CONTAINSSTRING ( FactData[Period], "SCR" ), "SCR",
         "Allocation" ) )

>> Life / Non-Life: номер таблицы того же типа внутри компании
SAA Part =
VAR e = FactData[Entity]
VAR b = FactData[BlockNo]
VAR t = IF ( FactData[SAA Table] = "Parameters", "SCR", FactData[SAA Table] )
RETURN
IF ( FactData[Topic] = "SAA",
     VAR blocks =
         DISTINCT ( SELECTCOLUMNS (
             FILTER ( FactData,
                 FactData[Topic] = "SAA" && FactData[Entity] = e && FactData[SAA Table] = t ),
             "B", FactData[BlockNo] ) )
     VAR n   = COUNTROWS ( FILTER ( blocks, [B] < b ) ) + 1
     VAR tot = COUNTROWS ( blocks )
     RETURN IF ( tot <= 1, "Total", SWITCH ( n, 1, "Life", 2, "Non-Life", "Part " & n ) ) )

>> Чистое имя столбца (вместо длинного текста из Excel)
SAA Col =
IF ( FactData[Topic] = "SAA",
     VAR m = UPPER ( FactData[Measure] )
     VAR p = FactData[PeriodType]
     RETURN SWITCH ( TRUE (),
         FactData[SAA Table] = "Parameters",          "Value",
         CONTAINSSTRING ( m, "CURRENT YIELD" ),        "Current Yield %",
         CONTAINSSTRING ( m, "TOTAL RETURN YIELD" ),   "Total Return Yield %",
         CONTAINSSTRING ( m, "CURRENT INCOME" ),       "Current Income",
         CONTAINSSTRING ( m, "TOTAL RETURN OF" ),      "Total Return",
         CONTAINSSTRING ( m, "SALES AND REDEMPTIONS" ),"Sales & Redemptions",
         CONTAINSSTRING ( m, "NET SCR" ),              "Net SCR " & FactData[Year],
         CONTAINSSTRING ( m, "PURCHASES DURING" ),     "SCR Purchases " & FactData[Year],
         CONTAINSSTRING ( m, "SALES & REDEMPTION" ),   "SCR Sales & Red. " & FactData[Year],
         p IN { "Actual", "Expected" },                "MV " & FactData[Period],
         CONTAINSSTRING ( m, "REDEMPTION" ),           "Redemptions",
         CONTAINSSTRING ( m, "PURCHASES" ),            "Purchases",
         CONTAINSSTRING ( m, "SALES" ),                "Sales",
         CONTAINSSTRING ( m, "NET BALANCE" ),          "Net Balance",
         CONTAINSSTRING ( m, "TARGET MV" ) && CONTAINSSTRING ( m, "%" ), "Target %",
         CONTAINSSTRING ( m, "TARGET MV" ),            "Target MV",
         CONTAINSSTRING ( m, "LOWER" ),                "Lower Bound %",
         CONTAINSSTRING ( m, "UPPER" ),                "Upper Bound %",
         FactData[Measure] ) )

SAA Col Order =
IF ( FactData[Topic] = "SAA",
     VAR c = FactData[SAA Col]
     RETURN SWITCH ( TRUE (),
         LEFT ( c, 6 ) = "MV Act", 1,   LEFT ( c, 6 ) = "MV Exp", 2,
         c = "Redemptions", 3,          c = "Purchases", 4,
         c = "Sales", 5,                c = "Net Balance", 6,
         c = "Target MV", 7,            c = "Target %", 8,
         c = "Lower Bound %", 9,        c = "Upper Bound %", 10,
         c = "Sales & Redemptions", 11, c = "Current Yield %", 12,
         c = "Total Return Yield %", 13, c = "Current Income", 14,
         c = "Total Return", 15,
         LEFT ( c, 7 ) = "Net SCR", 21, LEFT ( c, 8 ) = "SCR Purc", 22,
         LEFT ( c, 8 ) = "SCR Sale", 23,
         c = "Value", 30, 99 ) )

SAA Line =
IF ( FactData[Topic] = "SAA", FactData[Item] )

SAA Order =
IF ( FactData[Topic] = "SAA",
     VAR it = FactData[Item]
     RETURN SWITCH ( TRUE (),
         CONTAINSSTRING ( it, "Cash" ), 1,
         CONTAINSSTRING ( it, "Government" ), 2,
         CONTAINSSTRING ( it, "Corporate" ), 3,
         CONTAINSSTRING ( it, "Private Debt" ), 4,
         CONTAINSSTRING ( it, "Listed Equity" ), 5,
         CONTAINSSTRING ( it, "Infrastructure" ), 6,
         CONTAINSSTRING ( it, "Multi Asset" ), 7,
         CONTAINSSTRING ( it, "Real Estate" ), 8,
         CONTAINSSTRING ( it, "Other Inv" ), 9,
         LEFT ( it, 5 ) = "Total", 10,
         CONTAINSSTRING ( it, "Duration Fixed" ), 11,
         CONTAINSSTRING ( it, "Duration Gap" ), 12,
         CONTAINSSTRING ( it, "Return on Risk" ), 13,
         CONTAINSSTRING ( it, "Participation" ), 14,
         CONTAINSSTRING ( it, "Tax Rate" ), 15,
         99 ) )

>> Потом сортировка:
>>   SAA Col  > Ordina per colonna > SAA Col Order
>>   SAA Line > Ordina per colonna > SAA Order


>> ШАГ 4. МЕРЫ (Home > Nuova misura), по одной.
>> Значения в листе уже в EUR M -> без Scale (слайсер Scale на этой странице не нужен).

SAA Value = SUM ( FactData[Value] )

>> ФОРМАТ для SAA Value: Strumenti misura > Formato > "Dinamico" > вставить:
VAR c = SELECTEDVALUE ( FactData[SAA Col] )
VAR l = SELECTEDVALUE ( FactData[SAA Line] )
RETURN SWITCH ( TRUE (),
    CONTAINSSTRING ( l, "Duration" ), "0.0",
    CONTAINSSTRING ( c, "%" ) || CONTAINSSTRING ( l, "%" )
        || CONTAINSSTRING ( l, "Return on Risk" ), "0.0%",
    "#,##0.0" )

>> --- по строке (класс актива), для графика, цвета и комментария ---

SAA L Exp =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", LEFT ( FactData[SAA Col], 6 ) = "MV Exp" )

SAA L Tgt =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Target MV" )

SAA L Pct =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Target %" )

SAA L Lo =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Lower Bound %" )

SAA L Hi =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Upper Bound %" )

>> Красный Target %, если вне границ (Formattazione condizionale > Colore carattere > Valore campo)
SAA Pct Color =
VAR c  = SELECTEDVALUE ( FactData[SAA Col] )
VAR pc = [SAA L Pct]
VAR lo = [SAA L Lo]
VAR hi = [SAA L Hi]
RETURN IF ( c = "Target %" && NOT ISBLANK ( pc )
            && ( ( NOT ISBLANK ( lo ) && pc < lo ) || ( NOT ISBLANK ( hi ) && pc > hi ) ),
            "#C00000" )

>> --- итоги по компании (классы 1-9, без строки Total) ---

SAA MV Act =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    LEFT ( FactData[SAA Col], 6 ) = "MV Act", FactData[SAA Order] <= 9 )

SAA MV Exp =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    LEFT ( FactData[SAA Col], 6 ) = "MV Exp", FactData[SAA Order] <= 9 )

SAA Tgt MV =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Target MV", FactData[SAA Order] <= 9 )

SAA Purch =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Purchases", FactData[SAA Order] <= 9 )

SAA Sales =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Sales", FactData[SAA Order] <= 9 )

SAA Redem =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Redemptions", FactData[SAA Order] <= 9 )

SAA NetBal =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Net Balance", FactData[SAA Order] <= 9 )

SAA Income =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
    FactData[SAA Col] = "Current Income", FactData[SAA Order] <= 9 )

SAA Dur Gap =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FactData[Topic] = "SAA", LEFT ( FactData[SAA Col], 6 ) = "MV Exp",
    CONTAINSSTRING ( FactData[Item], "Duration Gap" ) )

>> --- SCR и параметры: если у компании одна SCR-таблица (Part = Total),
>>     она подходит и к Life, и к Non-Life ---

SAA SCR Purch =
VAR p = SELECTEDVALUE ( FactData[SAA Part] )
RETURN CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FILTER ( ALL ( FactData[SAA Part] ), ISBLANK ( p ) || FactData[SAA Part] IN { p, "Total" } ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR",
    LEFT ( FactData[SAA Col], 8 ) = "SCR Purc", FactData[SAA Order] <= 9 )

SAA SCR Sales =
VAR p = SELECTEDVALUE ( FactData[SAA Part] )
RETURN CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FILTER ( ALL ( FactData[SAA Part] ), ISBLANK ( p ) || FactData[SAA Part] IN { p, "Total" } ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR",
    LEFT ( FactData[SAA Col], 8 ) = "SCR Sale", FactData[SAA Order] <= 9 )

SAA SCR Net =
VAR p = SELECTEDVALUE ( FactData[SAA Part] )
RETURN CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FILTER ( ALL ( FactData[SAA Part] ), ISBLANK ( p ) || FactData[SAA Part] IN { p, "Total" } ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR",
    LEFT ( FactData[SAA Col], 7 ) = "Net SCR", FactData[SAA Order] <= 9 )

SAA RoRC =
VAR p = SELECTEDVALUE ( FactData[SAA Part] )
RETURN CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
    FILTER ( ALL ( FactData[SAA Part] ), ISBLANK ( p ) || FactData[SAA Part] IN { p, "Total" } ),
    FactData[Topic] = "SAA", CONTAINSSTRING ( FactData[Item], "Return on Risk" ) )

>> --- АВТОКОММЕНТАРИЙ ---

SAA Comment =
VAR ent = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR one = HASONEVALUE ( FactData[SAA Part] )
VAR prt = IF ( one, SELECTEDVALUE ( FactData[SAA Part] ) )
VAR yA  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Actual" )
VAR yE  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Expected" )
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR act = [SAA MV Act]
VAR ex  = [SAA MV Exp]
VAR tg  = [SAA Tgt MV]
VAR pu  = [SAA Purch]
VAR sa  = [SAA Sales]
VAR re  = [SAA Redem]
VAR nb  = [SAA NetBal]
VAR inc = [SAA Income]
VAR dg  = [SAA Dur Gap]
VAR sP  = [SAA SCR Purch]
VAR sS  = [SAA SCR Sales]
VAR sN  = [SAA SCR Net]
VAR ror = [SAA RoRC]
VAR hasData = NOT ISBLANK ( act ) || NOT ISBLANK ( ex ) || NOT ISBLANK ( tg )
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"

VAR cls = CALCULATETABLE (
              SUMMARIZE ( FactData, FactData[SAA Part], FactData[SAA Line] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Order] <= 9 )
VAR x   = ADDCOLUMNS ( cls, "@e", [SAA L Exp], "@t", [SAA L Tgt],
                            "@p", [SAA L Pct], "@lo", [SAA L Lo], "@hi", [SAA L Hi] )
VAR mv  = FILTER ( x, NOT ISBLANK ( [@t] ) || NOT ISBLANK ( [@e] ) )
VAR top = TOPN ( 1, mv, ABS ( [@t] - [@e] ), DESC )
VAR topN = MAXX ( top, FactData[SAA Line] & IF ( NOT one, " (" & FactData[SAA Part] & ")" ) )
VAR topD = MAXX ( top, [@t] - [@e] )
VAR withPct = FILTER ( x, NOT ISBLANK ( [@p] ) )
VAR br  = FILTER ( withPct,
              ( NOT ISBLANK ( [@lo] ) && [@p] < [@lo] ) || ( NOT ISBLANK ( [@hi] ) && [@p] > [@hi] ) )
VAR nBr = COUNTROWS ( br )
VAR brTxt = CONCATENATEX ( br,
              FactData[SAA Line] & IF ( NOT one, " (" & FactData[SAA Part] & ")" ) & " "
              & FORMAT ( [@p], "0.0%" ) & " vs " & FORMAT ( [@lo], "0%" ) & "–" & FORMAT ( [@hi], "0%" ), "; " )
VAR pctSum = SUMX ( withPct, [@p] )
VAR gap = ex + nb - tg

VAR L1 = ent & IF ( one && prt <> "Total", " – " & prt ) & ": investments "
         & FORMAT ( act, fmt ) & " at Act " & yA & ", " & FORMAT ( ex, fmt ) & " expected for " & yE
         & ", SAA " & yS & " target " & FORMAT ( tg, fmt )
         & " (" & FORMAT ( tg - ex, fmtS ) & " vs " & yE & ")."
VAR L2 = IF ( NOT ISBLANK ( nb ) || NOT ISBLANK ( pu ),
         "Net balance " & yS & ": " & FORMAT ( nb, fmtS ) & " (purchases " & FORMAT ( pu, fmt )
         & ", sales " & FORMAT ( sa, fmt ) & ", redemptions/distributions " & FORMAT ( re, fmt ) & ")." )
VAR L3 = IF ( NOT ISBLANK ( topD ) && topD <> 0,
         "Largest reallocation: " & topN & " " & FORMAT ( topD, fmtS ) & " vs " & yE & "." )
VAR L4 = IF ( COUNTROWS ( withPct ) > 0,
         IF ( nBr = 0, "All target weights are within the SAA bounds.",
              nBr & IF ( nBr = 1, " asset class", " asset classes" ) & " outside the bounds: " & brTxt & "." ) )
VAR L5 = IF ( one && COUNTROWS ( withPct ) > 0 && ABS ( pctSum - 1 ) > 0.001,
         "Check: target weights add up to " & FORMAT ( pctSum, "0.0%" ) & ", not 100%." )
VAR L6 = IF ( NOT ISBLANK ( tg ) && NOT ISBLANK ( nb ) && ABS ( gap ) >= 1,
         "Expected " & yE & " + net balance differs from the target by " & FORMAT ( gap, fmtS )
         & " (market effect or input to check)." )
VAR L7 = IF ( NOT ISBLANK ( inc ),
         "Expected current income on purchases: " & FORMAT ( inc, "#,##0.0" ) & "." )
VAR L8 = IF ( one && NOT ISBLANK ( dg ), "Duration gap " & yE & ": " & FORMAT ( dg, "0.0" ) & " years." )
VAR L9 = IF ( NOT ISBLANK ( sP ) || NOT ISBLANK ( sS ),
         "SCR absorption " & yS & ": " & FORMAT ( sP, fmtS ) & " from purchases, "
         & FORMAT ( sS, fmtS ) & " from sales & redemptions (net " & FORMAT ( sP + sS, fmtS ) & ")"
         & IF ( NOT ISBLANK ( sN ), "; net SCR " & yE & ": " & FORMAT ( sN, fmt ) ) & "." )
VAR L10 = IF ( one && NOT ISBLANK ( ror ),
          "BU return on risk capital " & yS & ": " & FORMAT ( ror, "0.0%" ) & "." )
VAR lines = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ),
              ( 6, L6 ), ( 7, L7 ), ( 8, L8 ), ( 9, L9 ), ( 10, L10 ) }
RETURN
IF ( NOT hasData, ent & ": no SAA data filled in.",
     CONCATENATEX ( FILTER ( lines, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )


>> ШАГ 5. СТРАНИЦА "SAA"
>>
>> Фильтр страницы: Topic = SAA.
>>
>> 1) Слайсер Entity (selezione singola) и слайсер FactData[SAA Part]
>>    (selezione singola; у Company #1 будет Life / Non-Life, у остальных Total).
>>
>> 2) Matrix "Asset allocation"
>>    Righe: SAA Line   Colonne: SAA Col   Valori: SAA Value
>>    Filtri visual: SAA Table = Allocation; SAA Col Order <= 10
>>    Formato > Subtotali di riga / Totali di colonna: Disattivato
>>    Colore carattere (Valori) > fx > Valore campo > SAA Pct Color
>>
>> 3) Matrix "Purchases: yields and income" (таблица справа)
>>    То же, фильтр SAA Col Order tra 11 e 15.
>>
>> 4) Matrix "Expected SCR absorption"
>>    Righe: SAA Line   Colonne: SAA Col   Valori: SAA Value
>>    Filtri: SAA Table = SCR; SAA Order <= 15
>>    (в SCR-блок Company #1/#2 попадает строка "Company #2"/"#3" — этот фильтр её убирает)
>>    Если при выборе Life таблица пустая (одна SCR-таблица на компанию):
>>    выдели слайсер SAA Part > Formato > Modifica interazioni > на этой
>>    матрице нажми значок "Nessuna".
>>
>> 5) Matrix/Tabella "Parameters": Righe SAA Line, Valori SAA Value,
>>    Filtri: SAA Table = Parameters; SAA Order <= 15.
>>
>> 6) Grafico a barre raggruppate "Expected vs SAA target"
>>    Asse Y: SAA Line   Valori: SAA L Exp, SAA L Tgt
>>    Filtro visual: SAA Order <= 9
>>
>> 7) Comment box: Scheda (nuova) или Casella di testo-карточка с SAA Comment,
>>    перенос текста включён.
>>
>> ПРОВЕРКА на тестовых цифрах:
>>   - матрица = жёлтая таблица Excel 1:1 (порядок строк/столбцов);
>>   - класс, который ты поставила вне границ, красный и есть в комментарии;
>>   - "Net balance" = сумма столбца Net Balance в Excel.
