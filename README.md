=====================================================================
 4 PAGES: REMITTANCE BY COMPANY | LIFE KPIs | NON-LIFE KPIs | DEBT MANAGEMENT
 Строки с ">>" не копировать. Копировать только код.
 Все листы в EUR M -> меры без Scale (слайсер Scale на этих страницах не нужен).
=====================================================================

>> ШАГ 0. ПРОВЕРКА ПАРСЕРА (5 минут, до всего остального)
>> Trasforma dati > CheckBlocks. Строки:
>>   Remittance by Company -> ожидаю N blocks, Entity = Company #1..#4
>>   Life KPIs             -> 4 блока на компанию (NBV, NBV/PVNBP, NB Guar., IF Guar.)
>>   Non-Life KPIs         -> 2 блока на компанию (основная таблица + RoRC)
>> Потом Visualizzazione tabella > FactData, фильтр Topic = каждая из трёх:
>> пришли по одному фото колонок Block | Entity | Scenario | Section | Item | Period | Measure.
>> Меры ниже написаны под то, что видно на твоих фото Excel; если имена
>> строк во входной таблице другие — поправлю одно слово.


>> ===== РЕЗУЛЬТАТ ПРОВЕРКИ (твои фото) и 2 правки в Power Query =====
>> Remittance by Company: 4 блока ok. Во входе только Dividends / Interim (со знаком -).
>>    У Company #4 строка Previous не распознана (Group пустой): previous-годы
>>    попали в Current -> ПРАВКА 2.
>> Life KPIs: 8 блоков ok (NBV / PVNBP / NB Guar. / IF Guar. x 2 компании),
>>    но Entity ПУСТОЙ: ярлык "Company #1" стоит не слева от таблицы -> ПРАВКА 1.
>>    Вторая таблица — PVNBP в EUR, маржа считается мерой LK Margin.
>> Non-Life KPIs: 4 блока ok, Entity / Section / Gross-Net ok, SII 2025-2029
>>    = -1 / 48 / 98 / 128 / 143 как в Excel.

>> ПРАВКА 1. fnReadSheet: искать "Company #..." сначала слева (как сейчас),
>> потом в остальных колонках окна, потом во всей строке.
>> fnReadSheet > Editor avanzato > найди блок от "EntOf = (h) =>" до строки
>> "...EntNames, (n) => Text.StartsWith(Text.Upper(t), n))))), null),"
>> и замени его целиком на:

    EntOf = (h) =>
        let p0 = P0Of(h),
            Rws   = List.Select(List.Numbers(h, 7, -1), (r) => r >= 0),
            IsEnt = (t) => t <> null and
                       (Text.StartsWith(Text.Upper(t), "COMPANY #")
                        or List.AnyTrue(List.Transform(EntNames, (n) => Text.StartsWith(Text.Upper(t), n)))),
            Find  = (cols as list) => List.First(List.Select(List.Combine(List.Transform(Rws,
                        (r) => List.Transform(cols, (c) => Txt(Rows{r}{c})))), IsEnt), null),
            L = Find(List.Reverse(List.Select(WinCols, (c) => c < p0))),
            R = if L <> null then L else Find(List.Select(WinCols, (c) => c >= p0)),
            A = if R <> null then R else Find(List.Numbers(0, NC))
        in A,

>> (запятая в конце нужна: дальше идёт строка EntFill = ...)
>> Другие страницы не меняются: где компания находилась слева, она находится так же.

>> ПРАВКА 2. FactData: если у Remittance by Company нет группы Current/Previous,
>> второй одинаковый год в той же строке = Previous.
>> FactData > Editor avanzato. Внизу:   in  <ПоследнийШаг>
>> замени эти 2 строки на (вместо Typed — имя твоего последнего шага, 2 места):

    ,Idx     = Table.AddIndexColumn(Typed, "FIdx", 0, 1, Int64.Type),
    FirstI   = Table.Group(Table.SelectRows(Idx, each [Topic] = "Remittance by Company"),
                   {"BlockNo", "RowNo", "Period"}, {{"FirstIdx", each List.Min([FIdx]), Int64.Type}, {"Cnt", each Table.RowCount(_), Int64.Type}}),
    Joined   = Table.NestedJoin(Idx, {"BlockNo", "RowNo", "Period"}, FirstI, {"BlockNo", "RowNo", "Period"},
                   "F", JoinKind.LeftOuter),
    Expanded2 = Table.ExpandTableColumn(Joined, "F", {"FirstIdx", "Cnt"}),
    Scen     = Table.AddColumn(Expanded2, "Scen", each
                   if [Topic] = "Remittance by Company" and ( [Group] = null or [Group] = "" )
                      and [FirstIdx] <> null and [FIdx] > [FirstIdx]
                   then "Previous" else [Scenario], type text),
    Final    = Table.RenameColumns(Table.RemoveColumns(Scen, {"Scenario", "FIdx", "FirstIdx", "Cnt"}), {{"Scen", "Scenario"}})
in
    Final

>> ПРОВЕРКА после Chiudi e applica: FactData, Topic = Remittance by Company, Entity = Company #4,
>> Item = Dividends Outflow, Period 2026:  Current = -171.6, Previous = -388.7.
>> Если наоборот (Current -388.7) — в Excel previous-колонки стоят слева; тогда
>> в строке Scen замени  [FIdx] > [FirstIdx]  на  ( [Cnt] > 1 and [FIdx] = [FirstIdx] ).
>> Life KPIs: Entity теперь Company #1 / Company #2.


#####################################################################
 A) REMITTANCE BY COMPANY
#####################################################################

>> Колонки не нужны. Меры (Home > Nuova misura):

>> Во входной таблице есть только Dividends Outflow и Interim Dividends Outflow,
>> со знаком минус (outflow). Total = сумма двух, со знаком плюс, как в выходной таблице.
RBC Dividends =
- CALCULATE ( SUM ( FactData[Value] ),
    FactData[Topic] = "Remittance by Company", FactData[Scenario] = "Current",
    CONTAINSSTRING ( FactData[Item], "Dividends Outflow" )
        && NOT CONTAINSSTRING ( FactData[Item], "Interim" ) )

RBC Interim =
- CALCULATE ( SUM ( FactData[Value] ),
    FactData[Topic] = "Remittance by Company", FactData[Scenario] = "Current",
    CONTAINSSTRING ( FactData[Item], "Interim" ) )

RBC Total = [RBC Dividends] + [RBC Interim]

RBC Previous =
- CALCULATE ( SUM ( FactData[Value] ),
    FactData[Topic] = "Remittance by Company", FactData[Scenario] = "Previous" )

RBC Delta =
IF ( NOT ISBLANK ( [RBC Previous] ), [RBC Total] - [RBC Previous] )

>> Формат всех RBC: Numero intero, разделитель тысяч.

RBC Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR oneEnt = HASONEVALUE ( FactData[Entity] )
VAR oneYr  = HASONEVALUE ( FactData[Year] )
VAR ent = SELECTEDVALUE ( FactData[Entity] )
VAR yr  = SELECTEDVALUE ( FactData[Year] )
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), REMOVEFILTERS ( FactData[Year] ),
              FactData[Topic] = "Remittance by Company", FactData[Scenario] = "Current",
              NOT ISBLANK ( FactData[Year] ) )
VAR y0 = MINX ( yrs, FactData[Year] )
VAR yN = MAXX ( yrs, FactData[Year] )
VAR byYear = CONCATENATEX ( yrs, FactData[Year] & ": " & FORMAT ( [RBC Total], fmt ), ", ", FactData[Year], ASC )
VAR tot    = SUMX ( yrs, [RBC Total] )
VAR yrsP   = FILTER ( yrs, NOT ISBLANK ( [RBC Previous] ) )
VAR dlt    = SUMX ( yrsP, [RBC Delta] )
VAR prevT  = SUMX ( yrsP, [RBC Previous] )
VAR inter  = SUMX ( yrs, [RBC Interim] )
VAR ents = CALCULATETABLE ( VALUES ( FactData[Entity] ), REMOVEFILTERS ( FactData[Entity], DimEntity ),
               FactData[Topic] = "Remittance by Company", NOT ISBLANK ( FactData[Entity] ) )
VAR e2   = FILTER ( ADDCOLUMNS ( ents, "@t", SUMX ( yrs, [RBC Total] ),
                                       "@d", SUMX ( yrsP, [RBC Delta] ) ),
                    NOT ISBLANK ( [@t] ) && [@t] <> 0 )
VAR allT = SUMX ( e2, [@t] )
VAR nEnt = COUNTROWS ( e2 )
VAR topE = TOPN ( 1, e2, [@t], DESC )
VAR bigName = MAXX ( topE, FactData[Entity] )
VAR topV = MAXX ( topE, [@t] )
VAR worst  = TOPN ( 1, FILTER ( e2, [@d] < 0 ), [@d], ASC )
VAR worstN = MAXX ( worst, FactData[Entity] )
VAR worstV = MAXX ( worst, [@d] )
-- page mode: all companies
VAR P1 = "Total remittance " & y0 & "–" & yN & ": " & FORMAT ( tot, fmt ) & " from " & nEnt
         & IF ( nEnt = 1, " company", " companies" ) & " (" & byYear & ")."
VAR P2 = IF ( nEnt > 1, "Largest contributor: " & bigName & " with " & FORMAT ( topV, fmt )
         & " (" & FORMAT ( DIVIDE ( topV, allT ), "0%" ) & ")." )
VAR P3 = IF ( NOT ISBLANK ( prevT ),
         "Vs previous plan: " & FORMAT ( dlt, fmtS ) & " (" & FORMAT ( DIVIDE ( dlt, ABS ( prevT ) ), "+0%;-0%" ) & ")"
         & IF ( NOT ISBLANK ( worstV ), "; largest decrease " & worstN & " " & FORMAT ( worstV, fmtS ) ) & "." )
VAR P4 = IF ( inter <> 0, "Interim dividends: " & FORMAT ( inter, fmt ) & " ("
         & FORMAT ( DIVIDE ( inter, tot ), "0%" ) & " of total remittance)." )
-- company mode
VAR dByYear = CONCATENATEX ( yrsP, FactData[Year] & " " & FORMAT ( [RBC Delta], fmtS ), ", ", FactData[Year], ASC )
VAR shareC  = DIVIDE ( tot, CALCULATE ( SUMX ( yrs, [RBC Total] ), REMOVEFILTERS ( FactData[Entity], DimEntity ) ) )
VAR C1 = ent & ": remittance " & FORMAT ( tot, fmt ) & " over " & y0 & "–" & yN & " (" & byYear & "), "
         & FORMAT ( shareC, "0%" ) & " of the country total."
VAR C2 = IF ( NOT ISBLANK ( prevT ), "Vs previous plan: " & FORMAT ( dlt, fmtS ) & " (" & dByYear & ")." )
VAR C3 = IF ( inter <> 0, "Of which interim dividends: " & FORMAT ( inter, fmt ) & "." )
-- single cell: company + year
VAR cT = [RBC Total]
VAR cP = [RBC Previous]
VAR Y1 = ent & " " & yr & ": total remittance " & FORMAT ( cT, fmt )
         & " (dividends " & FORMAT ( [RBC Dividends], fmt )
         & IF ( NOT ISBLANK ( [RBC Interim] ), ", interim " & FORMAT ( [RBC Interim], fmt ) ) & ")"
         & IF ( NOT ISBLANK ( cP ), "; previous plan " & FORMAT ( cP, fmt ) & ", delta " & FORMAT ( cT - cP, fmtS ) ) & "."
VAR pg = { ( 1, P1 ), ( 2, P2 ), ( 3, P3 ), ( 4, P4 ) }
VAR co = { ( 1, C1 ), ( 2, C2 ), ( 3, C3 ) }
RETURN
SWITCH ( TRUE (),
    ISBLANK ( tot ), "No remittance data.",
    oneEnt && oneYr, Y1,
    oneEnt, CONCATENATEX ( FILTER ( co, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ),
    CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> СТРАНИЦА "Remittance by Company" (фильтр страницы Topic = Remittance by Company)
>> 1) Matrix: Righe FactData[Entity]; Colonne FactData[Year];
>>    Valori: RBC Total, RBC Dividends, RBC Interim, RBC Previous, RBC Delta
>>    Formato > Valori > Opzioni > "Cambia valori nelle righe" = Attivato
>>      (меры встают строками под каждой компанией, как в Excel)
>>    Переименуй в Valori: Total Remittance (Cash View basis) | Dividends Outflow
>>      (excl. interim) | Interim Dividends Outflow | Previous Planned Remittance |
>>      Delta Current vs Previous
>>    Totali di colonna: Disattivato. RBC Delta: Colore carattere < 0 красный (Regole).
>> 2) Istogramma in pila: Asse X FactData[Year], Asse Y RBC Total, Legenda FactData[Entity].
>> 3) Comment box: RBC Comment.
>>    Без выбора -> вся страна; клик по компании -> компания; клик по ячейке -> компания+год.
>> ПРОВЕРКА (твоё фото): ES003 2026 total 129, previous 360, delta -231/-232 (округление);
>>    ES005 2026: 68 = 10 + 59 (dividends + interim).


#####################################################################
 B) LIFE KPIs
#####################################################################

>> ШАГ B1. КОЛОНКИ в FactData (Nuova colonna), по порядку

>> Какая из 4 таблиц: по строкам, которые есть в блоке
LK Table =
IF ( FactData[Topic] = "Life KPIs",
     VAR b   = FactData[BlockNo]
     VAR blk = FILTER ( FactData, FactData[Topic] = "Life KPIs" && FactData[BlockNo] = b )
     RETURN SWITCH ( TRUE (),
         COUNTROWS ( FILTER ( blk, CONTAINSSTRING ( FactData[Item], "Mathematical Reserves" ) ) ) > 0, "IF Guarantees",
         COUNTROWS ( FILTER ( blk, CONTAINSSTRING ( FactData[Item], "NB Premiums" ) ) ) > 0, "NB Guarantees",
         COUNTROWS ( FILTER ( blk, CONTAINSSTRING ( FactData[Item], "VNP" )
                                   || CONTAINSSTRING ( FactData[Item], "NBV official" ) ) ) > 0, "NBV",
         "PVNBP" ) )

>> Номер строки внутри таблицы (0 = первая строка)
LK Order =
IF ( FactData[Topic] = "Life KPIs",
     VAR b = FactData[BlockNo]
     RETURN FactData[RowNo]
            - MINX ( FILTER ( FactData, FactData[Topic] = "Life KPIs" && FactData[BlockNo] = b ), FactData[RowNo] ) )

>> Имя строки; одинаковые названия в разных таблицах разводим невидимым символом
LK Line =
IF ( FactData[Topic] = "Life KPIs",
     FactData[Item]
     & SWITCH ( FactData[LK Table], "PVNBP", UNICHAR ( 8203 ),
                "IF Guarantees", UNICHAR ( 8203 ) & UNICHAR ( 8203 ), "" ) )

>> Сортировка: LK Line > Ordina per colonna > LK Order.
>> Если ошибка "più valori di LK Order per lo stesso LK Line" — пришли фото, значит
>> у компаний разное число строк, поправлю.

>> ШАГ B2. МЕРЫ

LK Value = SUM ( FactData[Value] )

>> Формат LK Value: Dinamico >
VAR t  = SELECTEDVALUE ( FactData[LK Table] )
VAR it = SELECTEDVALUE ( FactData[LK Line] )
RETURN IF ( CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( it, "Guarantee-linked" )
            || CONTAINSSTRING ( it, "Average Guarantee" ), "0.0%", "#,##0" )

>> Маржа NBV / PVNBP (во входе есть только PVNBP в EUR; маржу считаем, как в Excel)
LK Margin =
VAR it    = SUBSTITUTE ( SELECTEDVALUE ( FactData[LK Line] ), UNICHAR ( 8203 ), "" )
VAR isTot = CONTAINSSTRING ( it, "Total" )
VAR pv = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[LK Line], FactData[LK Table] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "PVNBP",
             FILTER ( ALL ( FactData[Item] ),
                 IF ( isTot, CONTAINSSTRING ( FactData[Item], "Total" ), FactData[Item] = it ) ) )
VAR nb = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[LK Line], FactData[LK Table] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             FILTER ( ALL ( FactData[Item] ),
                 IF ( isTot, CONTAINSSTRING ( FactData[Item], "Total NBV" ), FactData[Item] = it ) ) )
RETURN DIVIDE ( nb, pv )

>> Формат LK Margin: Percentuale, 1 decimale.

LK Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR rowSel = ISFILTERED ( FactData[LK Line] ) && HASONEVALUE ( FactData[LK Line] )
VAR yA = CALCULATE ( MIN ( FactData[Year] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs" )
VAR yZ = CALCULATE ( MAX ( FactData[Year] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs" )
VAR pA = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yA )
VAR pZ = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yZ )
-- page mode
VAR vA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yA )
VAR vZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yZ )
VAR nA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "Total NBV" ), FactData[Year] = yA )
VAR nZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
             CONTAINSSTRING ( FactData[Item], "Total NBV" ), FactData[Year] = yZ )
VAR lob  = CALCULATETABLE ( VALUES ( FactData[LK Line] ),
               REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
               FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[LK Order] <= 3 )
VAR lobX = FILTER ( ADDCOLUMNS ( lob, "@v",
               CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ),
                   FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[Year] = yZ ) ),
               NOT ISBLANK ( [@v] ) && [@v] <> 0 )
VAR lobT = TOPN ( 1, lobX, [@v], DESC )
VAR lobN = MAXX ( lobT, FactData[LK Line] )
VAR lobV = MAXX ( lobT, [@v] )
VAR rec  = CALCULATETABLE ( VALUES ( FactData[LK Line] ),
               REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
               FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[LK Order] > 3,
               NOT CONTAINSSTRING ( FactData[Item], "Total" ) )
VAR recX = FILTER ( ADDCOLUMNS ( rec, "@v",
               CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ),
                   FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV", FactData[Year] = yZ ) ),
               NOT ISBLANK ( [@v] ) && [@v] <> 0 )
VAR recT = TOPN ( 1, recX, ABS ( [@v] ), DESC )
VAR recN = MAXX ( recT, FactData[LK Line] )
VAR recV = MAXX ( recT, [@v] )
VAR pvZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[LK Table] = "PVNBP",
             CONTAINSSTRING ( FactData[Item], "Total" ), FactData[Year] = yZ )
VAR mgZ = DIVIDE ( nZ, pvZ )
VAR gNB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Guarantee-linked NB" ), FactData[Year] = yZ )
VAR aNB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Average Guarantee NB" ), FactData[Year] = yZ )
VAR gIF = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Guarantee-linked Mathematical" ), FactData[Year] = yZ )
VAR aIF = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[LK Line], DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", CONTAINSSTRING ( FactData[Item], "Average Guarantee on Mathematical" ), FactData[Year] = yZ )
VAR L1 = ent & ": Total Life SII VNP (net of tax) " & FORMAT ( vA, fmt ) & " in " & pA & ", "
         & FORMAT ( vZ, fmt ) & " in " & pZ & " (" & FORMAT ( vZ - vA, fmtS ) & ")."
VAR L2 = IF ( NOT ISBLANK ( nA ) || NOT ISBLANK ( nZ ),
         "NBV official (gross of tax): " & FORMAT ( nA, fmt ) & " in " & pA & ", " & FORMAT ( nZ, fmt ) & " in " & pZ & "." )
VAR L3 = IF ( NOT ISBLANK ( lobV ), "Largest line of business in " & pZ & ": " & lobN & " (" & FORMAT ( lobV, fmt ) & ")." )
VAR L4 = IF ( NOT ISBLANK ( recV ), "Main item from NBV to VNP in " & pZ & ": " & recN & " " & FORMAT ( recV, fmtS ) & "." )
VAR L5 = IF ( NOT ISBLANK ( mgZ ) && mgZ <> 0, "NBV / PVNBP margin " & pZ & ": " & FORMAT ( mgZ, "0.0%" ) & "." )
VAR L6 = IF ( NOT ISBLANK ( gNB ) && gNB <> 0, "Guarantee-linked NB premiums " & pZ & ": " & FORMAT ( gNB, "0%" )
         & IF ( NOT ISBLANK ( aNB ) && aNB <> 0, ", average guarantee " & FORMAT ( aNB, "0.00%" ) ) & "." )
VAR L7 = IF ( NOT ISBLANK ( gIF ) && gIF <> 0, "Guarantee-linked mathematical reserves " & pZ & ": " & FORMAT ( gIF, "0%" )
         & IF ( NOT ISBLANK ( aIF ) && aIF <> 0, ", average guarantee " & FORMAT ( aIF, "0.00%" ) ) & "." )
-- row mode: one line clicked
VAR rl    = SELECTEDVALUE ( FactData[LK Line] )
VAR rTbl  = CALCULATE ( MAX ( FactData[LK Table] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ),
                FactData[Topic] = "Life KPIs" )
VAR isPct = CONTAINSSTRING ( rl, "(%)" ) || CONTAINSSTRING ( rl, "Guarantee-linked" )
            || CONTAINSSTRING ( rl, "Average Guarantee" )
VAR rf    = IF ( isPct, "0.0%", fmt )
VAR per   = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Year], FactData[Period] ),
                REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ), FactData[Topic] = "Life KPIs" )
VAR series = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ) ), rf ),
                ", ", FactData[Year], ASC )
VAR rA = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yA )
VAR rZ = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year] ),
             FactData[Topic] = "Life KPIs", FactData[Year] = yZ )
VAR R1 = rl & " (" & ent & ", " & rTbl & "): " & series & "."
VAR R2 = IF ( NOT ISBLANK ( rA ) || NOT ISBLANK ( rZ ),
         "Change " & pA & " to " & pZ & ": "
         & IF ( isPct, FORMAT ( ( rZ - rA ) * 100, "+0.0;-0.0;0.0" ) & " pp", FORMAT ( rZ - rA, fmtS ) ) & "." )
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ), ( 6, L6 ), ( 7, L7 ) }
RETURN
IF ( rowSel, R1 & UNICHAR ( 10 ) & R2,
     CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> СТРАНИЦА "Life KPIs" (фильтр страницы Topic = Life KPIs; слайсер Entity, selezione singola)
>> 4 матрицы, у всех: Righe LK Line; Colonne DimPeriod[Period]; Valori LK Value;
>>   Totali: Disattivato. Фильтр визуала LK Table =
>>   1) NBV             (заголовок "NBV")
>>   2) PVNBP           (заголовок "NBV / PVNBP (IFRS 17)"; в Valori ВМЕСТО LK Value -> LK Margin)
>>   3) NB Guarantees   (заголовок "NB Guarantees - on NB Premiums")
>>   4) IF Guarantees   (заголовок "IF Guarantees - on Mathematical Reserves")
>> + Comment box: LK Comment (клик по строке -> ряд по годам и изменение).
>> ПРОВЕРКА (ES003): Total Life SII VNP 27 / 19 / 18 / 20 / 21; Expense assumptions 39 / 34 / 32 / 34 / 36.


#####################################################################
 C) NON-LIFE KPIs
#####################################################################

>> ШАГ C1. КОЛОНКИ в FactData

NL Group =
IF ( FactData[Topic] = "Non-Life KPIs",
     VAR b = FactData[BlockNo]
     VAR isR = COUNTROWS ( FILTER ( FactData, FactData[Topic] = "Non-Life KPIs"
                   && FactData[BlockNo] = b && FactData[PeriodType] = "Benchmark" ) ) > 0
     RETURN SWITCH ( TRUE (),
         isR, "RoRC",
         CONTAINSSTRING ( FactData[Section], "Earned" ), "Earned Premiums",
         CONTAINSSTRING ( FactData[Section], "SII" ), "SII CY Technical Result",
         CONTAINSSTRING ( FactData[Section], "CoR" ), "CoR",
         "Other" ) )

NL Order =
IF ( FactData[Topic] = "Non-Life KPIs",
     VAR b = FactData[BlockNo]
     RETURN FactData[RowNo]
            - MINX ( FILTER ( FactData, FactData[Topic] = "Non-Life KPIs" && FactData[BlockNo] = b ), FactData[RowNo] )
            + IF ( FactData[NL Group] = "RoRC", 100, 0 ) )

NL Line =
IF ( FactData[Topic] = "Non-Life KPIs",
     FactData[Item]
     & SWITCH ( FactData[NL Group], "CoR", UNICHAR ( 8203 ),
                "RoRC", UNICHAR ( 8203 ) & UNICHAR ( 8203 ), "" ) )

>> Сортировка: NL Line > Ordina per colonna > NL Order.

>> ШАГ C2. МЕРЫ

NL Value = SUM ( FactData[Value] )

>> Формат NL Value: Dinamico >
IF ( SELECTEDVALUE ( FactData[NL Group] ) IN { "CoR", "RoRC" }, "0.0%", "#,##0" )

NL Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR rowSel = ISFILTERED ( FactData[NL Line] ) && HASONEVALUE ( FactData[NL Line] )
VAR yA = CALCULATE ( MIN ( FactData[Year] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
VAR yZ = CALCULATE ( MAX ( FactData[Year] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
VAR pA = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC", FactData[Year] = yA )
VAR pZ = CALCULATE ( MIN ( FactData[Period] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC", FactData[Year] = yZ )
VAR epA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yA )
VAR epZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )
VAR epG = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Gross", FactData[Year] = yZ )
VAR moZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             FactData[Item] = "Motor", FactData[Measure] = "Net", FactData[Year] = yZ )
VAR crA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "CoR",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yA )
VAR crZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "CoR",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )
VAR siA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "SII CY Technical Result", FactData[Year] = yA )
VAR siZ = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "SII CY Technical Result", FactData[Year] = yZ )
VAR roA = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "RoRC", CONTAINSSTRING ( FactData[Item], "RoRC" ),
             FactData[Measure] = "Net", FactData[PeriodType] = "Actual" )
VAR roB = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( FactData[NL Line], DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
             FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "RoRC", CONTAINSSTRING ( FactData[Item], "RoRC" ),
             FactData[Measure] = "Net", FactData[PeriodType] = "Benchmark" )
VAR cagr = IF ( epA > 0 && epZ > 0 && yZ > yA, ( epZ / epA ) ^ ( 1 / ( yZ - yA ) ) - 1 )
VAR N1 = IF ( NOT ISBLANK ( epZ ) && epZ <> 0,
         ent & ": net earned premiums (excl. GC&C) " & FORMAT ( epA, fmt ) & " in " & pA & ", "
         & FORMAT ( epZ, fmt ) & " in " & pZ & IF ( NOT ISBLANK ( cagr ), " (CAGR " & FORMAT ( cagr, "+0.0%;-0.0%" ) & ")" ) & "." )
VAR N2 = IF ( NOT ISBLANK ( epZ ) && epZ <> 0 && NOT ISBLANK ( moZ ),
         "Motor " & FORMAT ( DIVIDE ( moZ, epZ ), "0%" ) & " of net premiums in " & pZ
         & IF ( NOT ISBLANK ( epG ) && epG <> 0, "; retention (net / gross) " & FORMAT ( DIVIDE ( epZ, epG ), "0%" ) ) & "." )
VAR N3 = IF ( NOT ISBLANK ( crZ ) && crZ <> 0,
         "Net CoR (excl. GC&C): " & FORMAT ( crA, "0.0%" ) & " in " & pA & ", " & FORMAT ( crZ, "0.0%" ) & " in " & pZ
         & IF ( crZ < 1, " – below 100% (underwriting profit).", " – above 100% (underwriting loss)." ) )
VAR N4 = IF ( NOT ISBLANK ( siZ ) || NOT ISBLANK ( siA ),
         "Non-life SII CY technical result: " & FORMAT ( siA, fmt ) & " in " & pA & ", " & FORMAT ( siZ, fmt ) & " in " & pZ
         & " (" & FORMAT ( siZ - siA, fmtS ) & ")." )
VAR N5 = IF ( NOT ISBLANK ( roA ) && roA <> 0,
         "RoRC (excl. GC&C, net) " & FORMAT ( roA, "0.0%" )
         & IF ( NOT ISBLANK ( roB ) && roB <> 0, " vs benchmark " & FORMAT ( roB, "0.0%" )
                & IF ( roA >= roB, " – above benchmark.", " – below benchmark." ), "." ) )
-- row mode
VAR rl  = SELECTEDVALUE ( FactData[NL Line] )
VAR rG  = CALCULATE ( MAX ( FactData[NL Group] ), REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ),
              FactData[Topic] = "Non-Life KPIs" )
VAR rf  = IF ( rG IN { "CoR", "RoRC" }, "0.0%", fmt )
VAR per = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[PeriodType], FactData[Year], FactData[Period] ),
              REMOVEFILTERS ( DimPeriod, FactData[Period], FactData[Year], FactData[Measure] ), FactData[Topic] = "Non-Life KPIs" )
VAR sN  = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod, FactData[Measure] ), FactData[Measure] = "Net" ), rf ), ", ", FactData[Year], ASC )
VAR sG  = CONCATENATEX ( per, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod, FactData[Measure] ), FactData[Measure] = "Gross" ), rf ), ", ", FactData[Year], ASC )
VAR R1 = rl & " (" & ent & ")" & UNICHAR ( 10 ) & "Net: " & sN & "." & UNICHAR ( 10 ) & "Gross: " & sG & "."
VAR pg = { ( 1, N1 ), ( 2, N2 ), ( 3, N3 ), ( 4, N4 ), ( 5, N5 ) }
VAR txt = CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC )
RETURN
IF ( rowSel, R1, IF ( txt = "", ent & ": no Non-Life KPI data filled in.", txt ) )

>> СТРАНИЦА "Non-Life KPIs" (фильтр страницы Topic = Non-Life KPIs; слайсер Entity, selezione singola)
>> 1) Matrix: Righe NL Line; Colonne DimPeriod[Period] и под ним FactData[Measure]
>>    (Gross / Net); Valori NL Value; фильтр NL Group <> RoRC (Filtro di base: всё кроме RoRC)
>>    Formato > Layout: "Struttura" не нужен; Totali: Disattivato.
>>    Если Gross/Net идут не в том порядке: Measure > Ordina per colonna не нужен,
>>    просто в Colonne > меню поля > Ordina crescente.
>> 2) Matrix RoRC: то же, фильтр NL Group = RoRC (колонки Act 2025 | HY26 | Benchmark).
>> 3) Comment box: NL Comment (клик по строке -> ряд Net и Gross по годам).
>> ПРОВЕРКА (ES003): Non-life SII CY technical result -1 / 48 / 98 / 128 / 143.


#####################################################################
 D) DEBT MANAGEMENT  (это список, не FactData)
#####################################################################

>> ШАГ D1. ЗАПРОС List_Debt (Trasforma dati > Nuova origine > Query vuota >
>> Editor avanzato > вставить > имя List_Debt, Carica: ON).
>> Если у тебя уже есть List_DebtExisting / List_DebtNew — их можно удалить,
>> этот запрос читает обе таблицы и даёт колонкам короткие имена.

let
    Sheet = Workbook{[Item = "Debt Management", Kind = "Sheet"]}[Data],
    Map = (n as text) as text =>
        let u = Text.Upper(n) in
        if Text.StartsWith(u, "ID") then "ID"
        else if Text.Contains(u, "BORROWER") then "Borrower"
        else if Text.Contains(u, "LENDER") then "Lender"
        else if Text.Contains(u, "MATURITY") then "Maturity"
        else if Text.Contains(u, "CALL DATE") then "CallDate"
        else if Text.Contains(u, "TIERING") then "Tiering"
        else if Text.StartsWith(u, "CURRENCY") then "Currency"
        else if Text.Contains(u, "TRANCHE NUMBER") then "Tranche"
        else if Text.Contains(u, "UPDATED") then "PrincipalAfter"
        else if Text.Contains(u, "EXPECTED MONTH") then "ExpectedDate"
        else if Text.Contains(u, "PRINCIPAL AMOUNT AT") and Text.Contains(u, "EUR") then "PrincipalEUR"
        else if Text.Contains(u, "PRINCIPAL AMOUNT AT") then "PrincipalCCY"
        else if Text.Contains(u, "TRANCHE") and Text.Contains(u, "EUR") then "TrancheEUR"
        else if Text.Contains(u, "TRANCHE") then "TrancheCCY"
        else n,
    Std = (t as table, typ as text) as table =>
        let r = Table.RenameColumns(t, List.Transform(Table.ColumnNames(t), each {_, Map(_)}))
        in Table.AddColumn(r, "Type", each typ, type text),
    Ex  = Std(fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 1, 37, 0), "Existing"),
    Nw  = Std(fnReadList(Sheet, "Proposals on Existing Debt", null, "Borrower", 1, 2, 35, 0), "New"),
    All = Table.Combine({Ex, Nw}),
    ToNum = (v) => if v is number then v else try Number.From(v) otherwise null,
    Num = Table.TransformColumns(All,
              {{"PrincipalEUR", ToNum, type number}, {"TrancheEUR", ToNum, type number},
               {"PrincipalAfter", ToNum, type number}, {"PrincipalCCY", ToNum, type number},
               {"TrancheCCY", ToNum, type number}}, null, MissingField.Ignore),
    Yr  = Table.AddColumn(Num, "Year", each
              let d = try Record.Field(_, "ExpectedDate") otherwise null in
              if d is date then Date.Year(d)
              else if d is datetime then Date.Year(DateTime.Date(d))
              else if d is number then Date.Year(Date.From(d))
              else try Date.Year(Date.FromText(Text.From(d), [Culture = "en-GB"])) otherwise null,
              Int64.Type)
in
    Yr

>> Проверка: 2 строки-примера в Excel (одна в Existing, одна в New) -> Aggiorna ->
>> List_Debt показывает их с колонками ID, Borrower, ..., TrancheEUR, Type, Year.
>> Если в dummy таблица пустая — запрос вернёт 0 строк, это нормально.

>> ШАГ D2. МЕРЫ (в таблице List_Debt)

Debt Instruments = DISTINCTCOUNT ( List_Debt[ID] )

>> Номинал на YE2025: по одному разу на ID (транши повторяют номинал)
Debt Principal YE =
SUMX ( CALCULATETABLE ( VALUES ( List_Debt[ID] ), List_Debt[Type] = "Existing" ),
       CALCULATE ( MAX ( List_Debt[PrincipalEUR] ) ) )

Debt Reimbursements = CALCULATE ( SUM ( List_Debt[TrancheEUR] ), List_Debt[TrancheEUR] < 0 )

Debt Issuances = CALCULATE ( SUM ( List_Debt[TrancheEUR] ), List_Debt[TrancheEUR] > 0 )

Debt Net = SUM ( List_Debt[TrancheEUR] )

Debt Comment =
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR n    = [Debt Instruments]
VAR yrs  = CALCULATETABLE ( VALUES ( List_Debt[Year] ), NOT ISBLANK ( List_Debt[Year] ) )
VAR byY  = CONCATENATEX ( yrs, List_Debt[Year] & ": reimbursements " & FORMAT ( [Debt Reimbursements], fmt )
              & ", issuances " & FORMAT ( [Debt Issuances], fmt ) & ", net " & FORMAT ( [Debt Net], fmtS ),
              "; ", List_Debt[Year], ASC )
VAR nNew = CALCULATE ( DISTINCTCOUNT ( List_Debt[ID] ), List_Debt[Type] = "New" )
VAR aNew = CALCULATE ( SUM ( List_Debt[TrancheEUR] ), List_Debt[Type] = "New" )
VAR tier = ADDCOLUMNS ( VALUES ( List_Debt[Tiering] ), "@p", [Debt Principal YE] )
VAR tTop = TOPN ( 1, FILTER ( tier, NOT ISBLANK ( [@p] ) ), [@p], DESC )
VAR nFx  = CALCULATE ( DISTINCTCOUNT ( List_Debt[ID] ), List_Debt[Currency] <> "EUR", NOT ISBLANK ( List_Debt[Currency] ) )
VAR D1 = n & IF ( n = 1, " instrument", " instruments" ) & " in the proposals; outstanding principal at YE2025 "
         & FORMAT ( [Debt Principal YE], fmt ) & " (EUR m)."
VAR D2 = IF ( NOT ISBLANK ( byY ), byY & "." )
VAR D3 = IF ( nNew > 0, "New issuances: " & nNew & " for " & FORMAT ( aNew, fmt ) & "." )
VAR D4 = IF ( NOT ISBLANK ( MAXX ( tTop, [@p] ) ), "Largest tier by principal: " & MAXX ( tTop, List_Debt[Tiering] )
         & " (" & FORMAT ( MAXX ( tTop, [@p] ), fmt ) & ")." )
VAR D5 = IF ( nFx > 0, nFx & " instrument(s) not in EUR." )
VAR pg = { ( 1, D1 ), ( 2, D2 ), ( 3, D3 ), ( 4, D4 ), ( 5, D5 ) }
RETURN
IF ( n = 0 || ISBLANK ( n ), "No reimbursement or issuance proposals for 2026–2027.",
     CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> Формат: Debt Principal YE, Reimbursements, Issuances, Net -> Numero, 1 decimale.

>> СТРАНИЦА "Debt Management" (без фильтра Topic — это другая таблица)
>> 1) Tabella "Proposals on Existing Debt": фильтр Type = Existing; колонки
>>    ID, Borrower, Lender, Maturity, CallDate, Tiering, Currency, PrincipalCCY,
>>    PrincipalEUR, Tranche, TrancheEUR, PrincipalAfter, ExpectedDate
>>    (в Valori у числовых колонок поставь "Non riepilogare").
>> 2) Tabella "New Issuances": фильтр Type = New; ID, Borrower, Lender, Maturity,
>>    CallDate, Tiering, Currency, Tranche, TrancheCCY, TrancheEUR, PrincipalAfter, ExpectedDate.
>> 3) 4 карточки: Debt Principal YE | Debt Reimbursements | Debt Issuances | Debt Net.
>> 4) Istogramma: Asse X List_Debt[Year], Valori Debt Reimbursements + Debt Issuances.
>> 5) Comment box: Debt Comment (клик по Borrower в таблице -> комментарий по нему).
