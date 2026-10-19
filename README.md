=====================================================================
 SENSITIVITY ANALYSES ON NET RESULT AND SOLVENCY — сборка страницы
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ШАГ 0. ПРОВЕРКА (2 минуты)
>> Dati > FactData > фильтр Topic = Sensitivities. Посмотри колонки:
>>   Section: должно быть "Impact of Sensitivities on Local GAAP Net Result"
>>            и "Impact of Sensitivities on Solvency Ratio (%)"
>>   Item:    шоки (+50 bps Interest Rates, -25% Equity...) и строки
>>            Total Distributions..., Local GAAP Net Result, Solvency Ratio (%)
>>   Entity:  Cajamar Vida S.A..., Cajamar Seguros Generales... (НЕ пусто)
>>   Value у строки "No Volatility Adjustment", ES005, Act 2025:
>>            0,251  -> ppFactor = 100 (в мерах ниже уже так)
>>            25,1   -> поменяй ppFactor на 1 в двух мерах
>> Если Entity пустой или Section другой — пришли фото, поправлю.


>> ШАГ 1. ИСПРАВИТЬ NoScale
>> Строка "Solvency Ratio (%)" сейчас пересчитывается в тысячи.
>> Dati > FactData > колонка NoScale > заменить формулу:

NoScale =
VAR it  = UPPER ( FactData[Item] )
VAR itw = " " & SUBSTITUTE ( SUBSTITUTE ( SUBSTITUTE ( it, "(", " " ), ")", " " ), ":", " " ) & " "
VAR ctx = UPPER ( FactData[Block] & " " & FactData[Section] & " " & FactData[Measure] )
RETURN
    IF (
        FactData[Topic] = "Sensitivities",
        CONTAINSSTRING ( UPPER ( FactData[Section] & "" ), "SOLVENCY" ) || CONTAINSSTRING ( it, "(%)" ),
        CONTAINSSTRING ( it, "(%)" ) || CONTAINSSTRING ( itw, " RATIO " ) || CONTAINSSTRING ( it, "P.P." )
            || CONTAINSSTRING ( it, "TAX RATE" ) || CONTAINSSTRING ( it, "(YEARS)" )
            || CONTAINSSTRING ( ctx, "%" ) || CONTAINSSTRING ( ctx, "COR (" )
            || CONTAINSSTRING ( ctx, "RORC" ) || CONTAINSSTRING ( ctx, "PVNBP" )
    )


>> ШАГ 2. ПЯТЬ КОЛОНОК в FactData (Dati > FactData > Nuova colonna), по порядку.
>> Если Sens Metric / Sens Shock уже есть из прошлой инструкции — замени их формулы.

>> Номер строки внутри блока компании (для порядка строк как в Excel)
Sens RelRow =
IF ( FactData[Topic] = "Sensitivities",
     FactData[RowNo]
     - CALCULATE ( MIN ( FactData[RowNo] ), ALLEXCEPT ( FactData, FactData[Topic], FactData[BlockNo] ) ) )

>> Название строки. Шоки встречаются дважды (Net Result и Solvency),
>> поэтому к строкам блока Solvency добавлен невидимый символ.
Sens Line =
IF ( FactData[Topic] = "Sensitivities",
     FactData[Item]
     & IF ( CONTAINSSTRING ( FactData[Section] & "", "Solvency" ) && NOT CONTAINSSTRING ( FactData[Item], "Solvency Ratio" ),
            UNICHAR ( 8203 ), "" ) )

>> Порядок строк
Sens Order =
IF ( FactData[Topic] = "Sensitivities",
     CALCULATE ( MIN ( FactData[Sens RelRow] ),
         ALLEXCEPT ( FactData, FactData[Topic], FactData[Section], FactData[Item] ) ) )

>> Только строки-шоки (без заголовков, без Distributions / Net Result / Ratio)
Sens Shock =
IF ( FactData[Topic] = "Sensitivities"
         && LEFT ( FactData[Section] & "", 9 ) = "Impact of"
         && LEFT ( FactData[Item], 9 ) <> "Impact of"
         && NOT CONTAINSSTRING ( FactData[Item], "Solvency Ratio" ),
     FactData[Item] )

>> На что влияет шок
Sens Metric =
IF ( NOT ISBLANK ( FactData[Sens Shock] ),
     IF ( CONTAINSSTRING ( FactData[Section], "Solvency" ), "Solvency", "Net Result" ) )

>> Потом: колонка Sens Line > Ordina per colonna > Sens Order.


>> ШАГ 3. МЕРЫ (Home > Nuova misura), по одной.

>> Значение для таблицы: суммы в выбранных единицах, шоки на ratio в пунктах
Sens Value =
VAR ppFactor = 100
VAR v = CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities" )
RETURN IF ( SELECTEDVALUE ( FactData[Sens Metric] ) = "Solvency", v * ppFactor, v )

>> ФОРМАТ для Sens Value: выдели меру > Strumenti misura > Formato > "Dinamico"
>> > в строке формул выбери "Formato" (стрелка слева) и вставь:
>>
>> VAR it = SELECTEDVALUE ( FactData[Item] )
>> VAR m  = SELECTEDVALUE ( FactData[Sens Metric] )
>> RETURN SWITCH ( TRUE (),
>>     CONTAINSSTRING ( it, "Solvency Ratio" ), "0%",
>>     m = "Solvency", "+0.0 pp;-0.0 pp;0.0 pp",
>>     "#,##0;-#,##0;-" )
>>
>> (без ">>" в начале строк). Если "Dinamico" нет в твоей версии — напиши мне.

Sens NR =
CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Net Result" )

Sens Ratio pp =
VAR ppFactor = 100
RETURN CALCULATE ( [Value Current], FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency" ) * ppFactor

>> Коэффициент до шоков: строка "Solvency Ratio (%)" этого же листа (в %)
Sens Base Ratio =
CALCULATE ( [Value Current], REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[Sens Line], FactData[Sens Order], FactData[Sens Shock], FactData[Sens Metric] ),
    FactData[Topic] = "Sensitivities", CONTAINSSTRING ( FactData[Item], "Solvency Ratio" ) ) * 100

Sens Ratio After =
IF ( NOT ISBLANK ( [Sens Ratio pp] ), [Sens Base Ratio] + [Sens Ratio pp] )


>> ШАГ 4. КОММЕНТАРИЙ
>> Без выбора: последний год плана. Клик по ячейке: этот год. Клик по шоку: этот шок.

Sens Comment =
VAR hardLim = 110
VAR softLim = 130
VAR fmt  = "#,##0.0"
VAR fmtP = "+0.0;-0.0;0.0"
VAR nl   = UNICHAR ( 10 )
VAR y    = MAX ( FactData[Year] )
VAR nEnt = CALCULATE ( DISTINCTCOUNT ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Sensitivities" )
VAR who  = IF ( nEnt = 1,
               CALCULATE ( SELECTEDVALUE ( FactData[Entity] ), REMOVEFILTERS ( DimItem ), FactData[Topic] = "Sensitivities" ),
               nEnt & " companies (simple sum)" )
VAR base = CALCULATE ( [Sens Base Ratio], FactData[Year] = y )
VAR shocks =
    CALCULATETABLE (
        ADDCOLUMNS ( VALUES ( FactData[Sens Shock] ), "@pp", [Sens Ratio pp], "@nr", [Sens NR] ),
        REMOVEFILTERS ( DimItem ),
        REMOVEFILTERS ( FactData[Sens Line], FactData[Sens Order], FactData[Sens Shock], FactData[Sens Metric] ),
        FactData[Topic] = "Sensitivities", FactData[Year] = y )
VAR real = FILTER ( shocks, NOT ISBLANK ( FactData[Sens Shock] )
                            && NOT ( LEFT ( FactData[Sens Shock], 1 ) IN { "[", "<" } )
                            && ( [@pp] <> 0 || [@nr] <> 0 ) )
VAR worst   = TOPN ( 1, real, [@pp], ASC )
VAR wName   = MAXX ( worst, FactData[Sens Shock] )
VAR wPP     = MAXX ( worst, [@pp] )
VAR wAfter  = base + wPP
VAR worstNR = TOPN ( 1, FILTER ( real, [@nr] < 0 ), [@nr], ASC )
VAR nBreach = COUNTROWS ( FILTER ( real, base + [@pp] < hardLim ) )
VAR allTxt  = CONCATENATEX ( FILTER ( real, [@pp] <> 0 ),
                  "- " & FactData[Sens Shock] & ": " & FORMAT ( [@pp], fmtP ) & " pp", nl, [@pp], ASC )
VAR pageText =
    "SENSITIVITIES " & y & ": " & UPPER ( who ) & nl & nl
    & IF ( COUNTROWS ( real ) = 0, "No sensitivity results reported for " & y & ".",
           "Solvency ratio before shocks: " & FORMAT ( base, "0" ) & "%." & nl
           & "Most severe shock: " & wName & " (" & FORMAT ( wPP, fmtP ) & " pp), ratio after shock "
           & FORMAT ( wAfter, "0" ) & "%"
           & IF ( wAfter < hardLim, ": BELOW the " & hardLim & "% threshold.",
                  IF ( wAfter < softLim, ": between the " & hardLim & "% and " & softLim & "% thresholds.",
                       ": still above the " & softLim & "% threshold." ) ) & nl
           & IF ( nBreach > 0, nBreach & " shock(s) would bring the ratio below " & hardLim & "%.",
                  "No shock brings the ratio below " & hardLim & "%." ) & nl
           & IF ( NOT ISEMPTY ( worstNR ),
                  "Largest hit on Local GAAP net result: " & MAXX ( worstNR, FactData[Sens Shock] )
                  & " (" & FORMAT ( MAXX ( worstNR, [@nr] ), fmt ) & ").",
                  "No material impact on Local GAAP net result." )
           & nl & nl & "Impact on the solvency ratio:" & nl & allTxt )
VAR sh  = SELECTEDVALUE ( FactData[Sens Shock] )
VAR pp  = CALCULATE ( [Sens Ratio pp], FactData[Year] = y )
VAR nr  = CALCULATE ( [Sens NR], FactData[Year] = y )
VAR aft = base + pp
VAR rowText =
    UPPER ( sh ) & " " & y & ": " & who & nl & nl
    & "Impact on Local GAAP net result: " & FORMAT ( nr + 0, fmt ) & nl
    & "Impact on solvency ratio: " & FORMAT ( pp + 0, fmtP ) & " pp, from " & FORMAT ( base, "0" )
    & "% to " & FORMAT ( aft, "0" ) & "%"
    & IF ( aft < hardLim, ": BELOW the " & hardLim & "% threshold.",
           IF ( aft < softLim, ": between the thresholds.", "." ) )
RETURN IF ( HASONEVALUE ( FactData[Sens Shock] ) && NOT ISBLANK ( sh ), rowText, pageText )


>> ШАГ 5. СТРАНИЦА "Sensitivity analyses on net result and solvency"
>>
>> Фильтр страницы: Topic = Sensitivities; PeriodType = Actual, Forecast, Plan.
>> Срезы: Entity (можно несколько), Scale = Millions. RefScenario/CompScenario не нужны.
>>
>> a) МАТРИЦА как в Excel
>>    Righe:   Entity, под ним Sens Line   (компании одна под другой, как в Excel)
>>    Colonne: DimPeriod[Period]
>>    Valori:  Sens Value
>>    Фильтр визуала: Sens Line "non è vuoto"
>>    Formato > Subtotali di riga: Disattivato; Totali di colonna: Disattivato
>>    Чтобы видеть заголовки "Impact of Sensitivities on ..." (у них нет цифр):
>>      в Righe стрелка у Sens Line > "Mostra elementi senza dati"
>>    Цвет: Formattazione condizionale > Colore carattere > Regole:
>>      Sens Value < 0 -> красный; > 0 -> зелёный (только на шоки — ок и для всех)
>>
>> b) "ТОРНАДО" по коэффициенту (последний год или выбранный)
>>    Grafico a barre raggruppate: Asse Y = Sens Shock, Asse X = Sens Ratio pp
>>    Фильтр визуала: Sens Metric = Solvency; Year = 2029 (или срез Year)
>>    Ordina > Sens Ratio pp > Crescente (худший шок сверху)
>>
>> c) КАРТОЧКИ: Sens Base Ratio (подпись "Solvency ratio before shocks, %"),
>>    и минимальный ratio после шоков — по желанию.
>>
>> d) КОММЕНТАРИЙ: таблица с Sens Comment, A capo automatico = Attivato.


>> ПРОВЕРКА по твоему фото Excel (Cajamar Vida, ES005, Scale = Millions)
>>   Total Distributions  Act 2025 = -82;  Local GAAP Net Result Act 2025 = 82
>>   Solvency Ratio (%)   Act 2025 = 208%; Plan 2029 = 216%
>>   No Volatility Adjustment  Act 2025 = +25.1 pp
>>   +50 bps Interest Rates    Act 2025 = -3.7 pp
>>   -25% Equity               Act 2025 = -1.2 pp
