=====================================================================
 SAA v2 — три таблицы как в Excel: LIFE | NON-LIFE | CAPITAL CONSUMPTION
 + новый комментарий. Строки с ">>" не копировать.
 Колонки из шага 3 (SAA Part, SAA Line, SAA Order ...) и меры шага 4
 (SAA L Exp, SAA MV Act, SAA RoRC ...) ОСТАЮТСЯ — они нужны ниже.
=====================================================================

>> ШАГ A. ОДНА ПРАВКА КОЛОНКИ SAA Table
>> (лишняя строка "Company #2" внизу SCR-блока уходит в "Other",
>>  тогда фильтры по SAA Order больше не нужны)
>> Выдели колонку SAA Table > Ctrl+A в строке формул > вставь:

SAA Table =
IF ( FactData[Topic] = "SAA",
     SWITCH ( TRUE (),
         FactData[SAA Order] = 99, "Other",
         CONTAINSSTRING ( FactData[Item], "Return on Risk" )
             || CONTAINSSTRING ( FactData[Item], "Participation" )
             || CONTAINSSTRING ( FactData[Item], "Tax Rate" ), "Parameters",
         CONTAINSSTRING ( FactData[Period], "SCR" ), "SCR",
         "Allocation" ) )


=====================================================================
>> ШАГ B. МЕРЫ ДЛЯ ТАБЛИЦ LIFE и NON-LIFE (одни и те же меры, обе матрицы)
>> Столбцы как в Excel: Act MV % | Exp MV | Exp MV % | Redemptions |
>> Purchases | Sales | Net Balance | Target MV | Target MV % | Lower | Upper
=====================================================================

SAA a Act % =
VAR lineNm  = SELECTEDVALUE ( FactData[SAA Line] )
VAR v   = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", LEFT ( FactData[SAA Col], 6 ) = "MV Act" )
VAR tot = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col], FactData[SAA Line] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", LEFT ( FactData[SAA Col], 6 ) = "MV Act",
              FactData[SAA Order] <= 9 )
RETURN SWITCH ( TRUE (),
    CONTAINSSTRING ( lineNm, "Duration" ), v,
    LEFT ( lineNm, 5 ) = "Total", tot,
    DIVIDE ( v, tot ) )

SAA a Exp MV =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( CONTAINSSTRING ( lineNm, "Duration" ), BLANK (), [SAA L Exp] )

SAA a Exp % =
VAR lineNm  = SELECTEDVALUE ( FactData[SAA Line] )
VAR v   = [SAA L Exp]
VAR tot = CALCULATE ( [SAA L Exp], REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
RETURN SWITCH ( TRUE (),
    CONTAINSSTRING ( lineNm, "Duration" ), v,
    LEFT ( lineNm, 5 ) = "Total", BLANK (),
    DIVIDE ( v, tot ) )

SAA a Redemptions =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Redemptions" )

SAA a Purchases =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" )

SAA a Sales =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Sales" )

SAA a Net Balance =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Net Balance" )

SAA a Target MV = [SAA L Tgt]

SAA a Target % =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total", BLANK (), [SAA L Pct] )

SAA a Lower =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total" || CONTAINSSTRING ( lineNm, "Duration" ), BLANK (), [SAA L Lo] )

SAA a Upper =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total" || CONTAINSSTRING ( lineNm, "Duration" ), BLANK (), [SAA L Hi] )

>> Красный Target %, если вне границ
SAA a Pct Color =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
VAR pc = [SAA L Pct]
VAR lo = [SAA L Lo]
VAR hi = [SAA L Hi]
RETURN IF ( NOT CONTAINSSTRING ( lineNm, "Duration" ) && NOT ISBLANK ( pc )
            && ( ( NOT ISBLANK ( lo ) && pc < lo ) || ( NOT ISBLANK ( hi ) && pc > hi ) ),
            "#C00000" )

>> ФОРМАТЫ (выдели меру > Strumenti misura > Formato):
>>   SAA a Act %, SAA a Exp %, SAA a Target %  ->  "Dinamico", вставить:
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN SWITCH ( TRUE (),
    CONTAINSSTRING ( lineNm, "Duration" ), "0.0",
    LEFT ( lineNm, 5 ) = "Total", "#,##0",
    "0.0%" )
>>   SAA a Lower, SAA a Upper  ->  Percentuale, 1 decimale
>>   остальные SAA a ...       ->  Numero intero с разделителем тысяч


=====================================================================
>> ШАГ C. МЕРЫ ДЛЯ CAPITAL CONSUMPTION OVERVIEW
>> Столбцы: Net SCR 2026 | Purchases | SCR | SCR Abs % | Sales & Red. | SCR |
>> SCR Abs % | Net Balance | Net SCR | Current Yield % | Total Return %
>> Purchases / Sales & Red. = Life + Non-Life (как в Excel).
>> Проверено по твоему фото: Gov 112/40 = 280%; Net SCR 112+67 = 179;
>> Net Balance 40+40 = 80, итог 105+80 = 185.
=====================================================================

SAA c Net SCR 2026 =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR", LEFT ( FactData[SAA Col], 7 ) = "Net SCR" )

SAA c Purchases =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total",
    CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col], FactData[SAA Line], FactData[SAA Part] ),
        FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases",
        FactData[SAA Order] <= 9 ),
    CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col], FactData[SAA Part] ),
        FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" ) )

SAA c Sales Red =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total",
    CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col], FactData[SAA Line], FactData[SAA Part] ),
        FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
        FactData[SAA Col] IN { "Sales", "Redemptions" }, FactData[SAA Order] <= 9 ),
    CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col], FactData[SAA Part] ),
        FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
        FactData[SAA Col] IN { "Sales", "Redemptions" } ) )

SAA c SCR Purch =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR", LEFT ( FactData[SAA Col], 8 ) = "SCR Purc" )

SAA c SCR SalesRed =
CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
    FactData[Topic] = "SAA", FactData[SAA Table] = "SCR", LEFT ( FactData[SAA Col], 8 ) = "SCR Sale" )

SAA c Abs Purch % = DIVIDE ( [SAA c SCR Purch], [SAA c Purchases] )

SAA c Abs SalesRed % = DIVIDE ( [SAA c SCR SalesRed], [SAA c Sales Red] )

SAA c Net Balance = [SAA c Purchases] + [SAA c Sales Red]

SAA c Net SCR = [SAA c SCR Purch] + [SAA c SCR SalesRed]

>> Доходность покупок: средневзвешенная по покупкам Life и Non-Life
SAA c Yield % =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
VAR t = ADDCOLUMNS (
            CALCULATETABLE ( VALUES ( FactData[SAA Part] ), REMOVEFILTERS ( FactData[SAA Part] ),
                FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" ),
            "@p", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" ),
            "@y", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Current Yield %" ) )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total", BLANK (),
            DIVIDE ( SUMX ( t, [@p] * [@y] ), SUMX ( t, [@p] ), MAXX ( t, [@y] ) ) )

SAA c Total Return % =
VAR lineNm = SELECTEDVALUE ( FactData[SAA Line] )
VAR t = ADDCOLUMNS (
            CALCULATETABLE ( VALUES ( FactData[SAA Part] ), REMOVEFILTERS ( FactData[SAA Part] ),
                FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" ),
            "@p", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Purchases" ),
            "@y", CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
                      FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Col] = "Total Return Yield %" ) )
RETURN IF ( LEFT ( lineNm, 5 ) = "Total", BLANK (),
            DIVIDE ( SUMX ( t, [@p] * [@y] ), SUMX ( t, [@p] ), MAXX ( t, [@y] ) ) )

>> ФОРМАТЫ: SAA c ...% -> Percentuale 1 decimale; остальные SAA c -> Numero intero.
>> В Excel при нулевых покупках стоит "n. a.", у нас ячейка пустая.
>>
>> Блок под таблицей: "BU Return on Risk Capital 2027" = мера SAA RoRC (уже есть).
>> "Net Current Income / Net Total Return of Purchases on SCR" — пришли формулу
>> этих двух ячеек из Excel (клик по ячейке > фото строки формул), сделаю 1:1.


=====================================================================
>> ШАГ D. КОММЕНТАРИЙ (4 меры; SAA Comment ЗАМЕНЯЕТ старую)
=====================================================================

>> 1) Текст по одной части (Life или Non-Life), вся компания
SAA Part Text =
VAR prt = SELECTEDVALUE ( FactData[SAA Part] )
VAR lbl = SWITCH ( prt, "Life", "Life & SLT", "Non-Life", "Non-Life & NSLT", prt )
VAR yA  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Actual" )
VAR yE  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Expected" )
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR act = [SAA MV Act]
VAR ex  = [SAA MV Exp]
VAR tg  = [SAA Tgt MV]
VAR pu  = [SAA Purch]
VAR sa  = [SAA Sales]
VAR re  = [SAA Redem]
VAR nb  = [SAA NetBal]
VAR dg  = [SAA Dur Gap]
VAR s   = COALESCE ( act, 0 ) + COALESCE ( ex, 0 ) + COALESCE ( tg, 0 )
VAR cls = CALCULATETABLE ( VALUES ( FactData[SAA Line] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Order] <= 9 )
VAR x   = ADDCOLUMNS ( cls, "@e", [SAA L Exp], "@t", [SAA L Tgt],
                            "@p", [SAA L Pct], "@lo", [SAA L Lo], "@hi", [SAA L Hi] )
VAR mv  = FILTER ( x, NOT ISBLANK ( [@t] ) || NOT ISBLANK ( [@e] ) )
VAR bigRow  = TOPN ( 1, mv, ABS ( [@t] - [@e] ), DESC )
VAR bigName = MAXX ( bigRow, FactData[SAA Line] )
VAR bigD    = MAXX ( bigRow, [@t] - [@e] )
VAR withPct = FILTER ( x, NOT ISBLANK ( [@p] ) && [@p] <> 0 )
VAR br  = FILTER ( withPct,
              ( NOT ISBLANK ( [@lo] ) && [@p] < [@lo] ) || ( NOT ISBLANK ( [@hi] ) && [@p] > [@hi] ) )
VAR nBr = COUNTROWS ( br )
VAR brTxt = CONCATENATEX ( br, FactData[SAA Line] & " " & FORMAT ( [@p], "0.0%" )
              & " vs " & FORMAT ( [@lo], "0%" ) & "–" & FORMAT ( [@hi], "0%" ), "; " )
VAR pctSum = SUMX ( withPct, [@p] )
VAR L1 = lbl & ": investments " & FORMAT ( act, fmt ) & " at Act " & yA & ", " & FORMAT ( ex, fmt )
         & " expected for " & yE & ", SAA " & yS & " target " & FORMAT ( tg, fmt )
         & " (" & FORMAT ( tg - ex, fmtS ) & " vs " & yE & ")."
VAR L2 = IF ( NOT ISBLANK ( nb ),
         "Net balance " & yS & ": " & FORMAT ( nb - ex, fmtS ) & " (purchases " & FORMAT ( pu, fmt )
         & ", sales " & FORMAT ( sa, fmt ) & ", redemptions/distributions " & FORMAT ( re, fmt )
         & "), MV after rebalancing " & FORMAT ( nb, fmt ) & "." )
VAR L3 = IF ( NOT ISBLANK ( bigD ) && bigD <> 0,
         "Largest reallocation: " & bigName & " " & FORMAT ( bigD, fmtS ) & " vs " & yE & "." )
VAR L4 = IF ( COUNTROWS ( withPct ) > 0,
         IF ( nBr = 0, "All target weights are within the SAA bounds.",
              nBr & IF ( nBr = 1, " asset class", " asset classes" ) & " outside the bounds: " & brTxt & "." ) )
VAR L5 = IF ( COUNTROWS ( withPct ) > 0 && ABS ( pctSum - 1 ) > 0.001,
         "Check: target weights add up to " & FORMAT ( pctSum, "0.0%" ) & ", not 100%." )
VAR L6 = IF ( NOT ISBLANK ( tg ) && NOT ISBLANK ( nb ) && ABS ( nb - tg ) >= 1,
         "MV after net balance differs from the SAA target by " & FORMAT ( nb - tg, fmtS )
         & " (market effect or input to check)." )
VAR L7 = IF ( NOT ISBLANK ( dg ) && dg <> 0, "Duration gap " & yE & ": " & FORMAT ( dg, "0.0" ) & " years." )
VAR lines = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ), ( 6, L6 ), ( 7, L7 ) }
RETURN IF ( s = 0, BLANK (),
            CONCATENATEX ( FILTER ( lines, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> 2) Текст по одной части для выбранной строки (клик по классу актива)
SAA Part Row Text =
VAR prt = SELECTEDVALUE ( FactData[SAA Part] )
VAR lbl = SWITCH ( prt, "Life", "Life & SLT", "Non-Life", "Non-Life & NSLT", prt )
VAR lineNm  = SELECTEDVALUE ( FactData[SAA Line] )
VAR yA  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Actual" )
VAR yE  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Expected" )
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR lAct = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( FactData[SAA Col] ),
               FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", LEFT ( FactData[SAA Col], 6 ) = "MV Act" )
VAR lExp = [SAA L Exp]
VAR lTg  = [SAA L Tgt]
VAR lPct = [SAA L Pct]
VAR lLo  = [SAA L Lo]
VAR lHi  = [SAA L Hi]
VAR lNb  = [SAA a Net Balance]
VAR lPu  = [SAA a Purchases]
VAR lSa  = [SAA a Sales]
VAR lRe  = [SAA a Redemptions]
VAR s    = COALESCE ( lAct, 0 ) + COALESCE ( lExp, 0 ) + COALESCE ( lTg, 0 ) + COALESCE ( lNb, 0 )
VAR lPos = IF ( ISBLANK ( lPct ), "",
           IF ( NOT ISBLANK ( lLo ) && lPct < lLo, " – BELOW the lower bound",
           IF ( NOT ISBLANK ( lHi ) && lPct > lHi, " – ABOVE the upper bound", " – within the bounds" ) ) )
VAR isDur = CONTAINSSTRING ( lineNm, "Duration" )
VAR R1 = IF ( isDur,
           FORMAT ( lAct, "0.0" ) & " years at Act " & yA & ", " & FORMAT ( lExp, "0.0" ) & " expected for " & yE
           & IF ( NOT ISBLANK ( lPct ), ", SAA " & yS & " target " & FORMAT ( lPct, "0.0" ) ) & ".",
           "MV " & FORMAT ( lAct, fmt ) & " at Act " & yA & ", " & FORMAT ( lExp, fmt ) & " expected for " & yE
           & ", " & FORMAT ( lNb, fmt ) & " after net balance " & yS & " (" & FORMAT ( lNb - lExp, fmtS ) & ")." )
VAR R2 = IF ( NOT isDur && ( NOT ISBLANK ( lPu ) || NOT ISBLANK ( lSa ) || NOT ISBLANK ( lRe ) ),
           "Purchases " & FORMAT ( lPu, fmt ) & ", sales " & FORMAT ( lSa, fmt )
           & ", redemptions " & FORMAT ( lRe, fmt ) & "." )
VAR R3 = IF ( NOT isDur && NOT ISBLANK ( lTg ) && LEFT ( lineNm, 5 ) <> "Total",
           "SAA " & yS & " target " & FORMAT ( lTg, fmt ) & " = " & FORMAT ( lPct, "0.0%" )
           & IF ( NOT ISBLANK ( lLo ) || NOT ISBLANK ( lHi ),
                  " (bounds " & FORMAT ( lLo, "0%" ) & "–" & FORMAT ( lHi, "0%" ) & ")" ) & lPos & "." )
VAR rl = { ( 1, R1 ), ( 2, R2 ), ( 3, R3 ) }
RETURN IF ( s = 0, BLANK (),
            lbl & ": " & CONCATENATEX ( FILTER ( rl, [Value2] <> "" ), [Value2], " ", [Value1], ASC ) )

>> 3) Capital consumption (вся компания или выбранная строка)
SAA CC Text =
VAR rowSel = ISFILTERED ( FactData[SAA Line] ) && HASONEVALUE ( FactData[SAA Line] )
VAR lineNm  = SELECTEDVALUE ( FactData[SAA Line] )
VAR yE  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "Expected" )
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR fmt  = "#,##0"
VAR fmtS = "+#,##0;-#,##0;0"
// page totals (asset classes 1-9)
VAR pP  = CALCULATE ( [SAA c Purchases],    REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
VAR pSR = CALCULATE ( [SAA c Sales Red],    REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
VAR pSP = CALCULATE ( [SAA c SCR Purch],    REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
VAR pSS = CALCULATE ( [SAA c SCR SalesRed], REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
VAR pB  = CALCULATE ( [SAA c Net SCR 2026], REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
VAR pInc = CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[SAA Part], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
              FactData[SAA Col] = "Current Income", FactData[SAA Order] <= 9 )
VAR ror = [SAA RoRC]
VAR cls = CALCULATETABLE ( VALUES ( FactData[SAA Line] ),
              REMOVEFILTERS ( FactData[SAA Line], FactData[SAA Col], FactData[Item] ),
              FactData[Topic] = "SAA", FactData[SAA Table] = "SCR", FactData[SAA Order] <= 9 )
VAR x   = FILTER ( ADDCOLUMNS ( cls, "@n", [SAA c Net SCR] ), NOT ISBLANK ( [@n] ) )
VAR topSCR  = TOPN ( 1, x, [@n], DESC )
VAR topName = MAXX ( topSCR, FactData[SAA Line] )
VAR topVal  = MAXX ( topSCR, [@n] )
VAR P1 = IF ( NOT ISBLANK ( pSP ) || NOT ISBLANK ( pSS ),
         "Capital consumption " & yS & ": purchases " & FORMAT ( pP, fmt ) & " with SCR " & FORMAT ( pSP, fmt )
         & IF ( pP <> 0, " (" & FORMAT ( DIVIDE ( pSP, pP ), "0%" ) & " of purchases)" )
         & "; sales & redemptions " & FORMAT ( pSR, fmt ) & " with SCR " & FORMAT ( pSS, fmt )
         & "; net SCR " & FORMAT ( pSP + pSS, fmt )
         & IF ( NOT ISBLANK ( pB ), " (expected SCR absorption " & yE & ": " & FORMAT ( pB, fmt ) & ")" ) & "." )
VAR P2 = IF ( NOT ISBLANK ( topVal ), "Largest SCR consumer: " & topName & " (" & FORMAT ( topVal, fmt ) & ")." )
VAR P3 = IF ( NOT ISBLANK ( pInc ) && pP <> 0,
         "Expected current income on purchases: " & FORMAT ( pInc, "#,##0.0" )
         & " (" & FORMAT ( DIVIDE ( pInc, pP ), "0.0%" ) & " average yield)." )
VAR P4 = IF ( NOT ISBLANK ( ror ), "BU return on risk capital " & yS & ": " & FORMAT ( ror, "0.0%" ) & "." )
// selected line
VAR lB  = [SAA c Net SCR 2026]
VAR lP  = [SAA c Purchases]
VAR lSP = [SAA c SCR Purch]
VAR lSR = [SAA c Sales Red]
VAR lSS = [SAA c SCR SalesRed]
VAR lY  = [SAA c Yield %]
VAR Q1 = IF ( NOT ISBLANK ( lSP ) || NOT ISBLANK ( lSS ) || NOT ISBLANK ( lB ),
         "Capital consumption " & yS & ": purchases " & FORMAT ( lP, fmt ) & " with SCR " & FORMAT ( lSP, fmt )
         & IF ( lP <> 0, " (" & FORMAT ( DIVIDE ( lSP, lP ), "0%" ) & ")" )
         & ", sales & redemptions " & FORMAT ( lSR, fmt ) & " with SCR " & FORMAT ( lSS, fmt )
         & ", net SCR " & FORMAT ( lSP + lSS, fmt )
         & IF ( NOT ISBLANK ( lB ), " (expected " & yE & ": " & FORMAT ( lB, fmt ) & ")" ) & "." )
VAR Q2 = IF ( NOT ISBLANK ( lY ) && lY <> 0, "Current yield of purchases " & FORMAT ( lY, "0.0%" ) & "." )
VAR pg = { ( 1, P1 ), ( 2, P2 ), ( 3, P3 ), ( 4, P4 ) }
VAR rw = { ( 1, Q1 ), ( 2, Q2 ) }
RETURN IF ( rowSel,
            CONCATENATEX ( FILTER ( rw, [Value2] <> "" ), [Value2], " ", [Value1], ASC ),
            CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> 4) Итоговый комментарий (замени старую SAA Comment целиком)
SAA Comment =
VAR ent    = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR rowSel = ISFILTERED ( FactData[SAA Line] ) && HASONEVALUE ( FactData[SAA Line] )
VAR parts  = CALCULATETABLE ( VALUES ( FactData[SAA Part] ), REMOVEFILTERS ( FactData[SAA Line] ),
                 FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" )
VAR t      = FILTER ( ADDCOLUMNS ( parts, "@t", IF ( rowSel, [SAA Part Row Text], [SAA Part Text] ) ),
                 [@t] <> "" )
VAR body   = CONCATENATEX ( t, [@t], UNICHAR ( 10 ), FactData[SAA Part], ASC )
VAR cc     = [SAA CC Text]
VAR head   = IF ( rowSel, SELECTEDVALUE ( FactData[SAA Line] ) & " – " & ent,
                  ent & " – Strategic Asset Allocation (EUR m)" )
RETURN
IF ( body = "" && cc = "", ent & ": no SAA data filled in.",
     head & IF ( body <> "", UNICHAR ( 10 ) & body ) & IF ( cc <> "", UNICHAR ( 10 ) & cc ) )


=====================================================================
>> ШАГ E. СТРАНИЦА (как в Excel, сверху вниз)
=====================================================================
>> Удали: слайсер SAA Part, старые матрицы (Asset Allocation с SAA Col,
>>        Purchases yields, Expected SCR absorption, Parameters).
>> Оставь: слайсер Entity (selezione singola), фильтр страницы Topic = SAA.
>>
>> 1) Matrix "Total Investments (Life and SLT)"
>>    Righe: SAA Line
>>    Valori (в этом порядке): SAA a Act %, SAA a Exp MV, SAA a Exp %,
>>      SAA a Redemptions, SAA a Purchases, SAA a Sales, SAA a Net Balance,
>>      SAA a Target MV, SAA a Target %, SAA a Lower, SAA a Upper
>>    Filtri visual: SAA Table = Allocation; SAA Part = Life
>>    Переименуй столбцы в visual (двойной клик по полю в Valori):
>>      Act 2025 MV (%) | Exp 2026 MV | Exp 2026 MV (%) | Redemptions |
>>      Purchases | Sales | Net Balance | Target MV | Target MV (%) |
>>      Lower Bound | Upper Bound
>>    Formato > Totali di riga/colonna: Disattivato
>>    SAA a Target % > Formattazione condizionale > Colore carattere >
>>      Valore campo > SAA a Pct Color
>>
>> 2) Matrix "Total Investments (Non-life and NSLT)"
>>    Ctrl+C / Ctrl+V первой матрицы, фильтр SAA Part = Non-Life.
>>
>> 3) Matrix "Capital Consumption Overview"
>>    Righe: SAA Line
>>    Valori: SAA c Net SCR 2026, SAA c Purchases, SAA c SCR Purch,
>>      SAA c Abs Purch %, SAA c Sales Red, SAA c SCR SalesRed,
>>      SAA c Abs SalesRed %, SAA c Net Balance, SAA c Net SCR,
>>      SAA c Yield %, SAA c Total Return %
>>    Filtro visual: SAA Table = SCR
>>    Переименуй: Net SCR 2026 | Purchases | SCR | SCR Absorption (%) |
>>      Sales & Redemptions | SCR | SCR Absorption (%) | Net Balance |
>>      Net SCR | Current Yield (%) | Total Return (%)
>>
>> 4) Карточка: SAA RoRC (BU Return on Risk Capital).
>>
>> 5) Comment box: SAA Comment.
>>    Клик по строке в любой из 3 матриц -> комментарий по этому классу
>>    (Life, Non-Life и capital consumption сразу).
>>
>> ПРОВЕРКА по фото Excel (Company #1):
>>   Life: Cash Act 5.0%, Exp 55 / 5.2%, Total 995 / 1,055 / Net Bal 1,080 / Target 1,177
>>   CC:   Gov  Net SCR 2026 111, Purchases 40, SCR 112, 280%, S&R 40, SCR 67,
>>         167.5%, Net Balance 80, Net SCR 179; итог Purchases 105, S&R 80, NB 185
