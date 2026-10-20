=====================================================================
 SUMMARY v1 — библиотека порогов + 2 summary-страницы (Total и Companies)
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ИДЕЯ
>> 1) Таблица Thresholds = библиотека порогов. Её утверждают и меняют руками,
>>    в DAX ни одного числа порога нет — меры читают их отсюда.
>> 2) По каждой странице 2 меры: <PREFIX> Status (1 зелёный / 2 жёлтый / 3 красный /
>>    0 информационно) и <PREFIX> Headline (одно предложение).
>> 3) SummaryIndex = список разделов (Section, Prefix, Page). Одна мера SWITCH собирает
>>    все заголовки и статусы -> одна таблица на summary-странице.
>> 4) Клик по строке + кнопка "Open page" -> переход на страницу. Слайсер Entity
>>    синхронизирован на всех страницах -> на странице компании переход сразу
>>    открывает страницу, отфильтрованную по этой компании.


#####################################################################
 ШАГ 1. БИБЛИОТЕКА ПОРОГОВ (Thresholds)
#####################################################################
>> Home > Immettere i dati. Назови таблицу Thresholds. Колонки:
>>   Key | Section | Rule | Value | Unit | Status
>> Скопируй блок ниже (разделители — табуляция) и вставь Ctrl+V в первую ячейку.
>> Тип колонки Value = Numero decimale. Дробь пишется С ЗАПЯТОЙ (итальянская локаль).
>> Status: Draft -> после утверждения поменяй на Approved.
>> Изменить порог потом: Trasforma dati > Thresholds > шестерёнка у шага "Origine".

Key	Section	Rule	Value	Unit	Status
SOL_RED	Solvency	Red if lowest Solvency ratio over the plan is below	1,10	ratio	Draft
SOL_AMBER	Solvency	Amber if lowest Solvency ratio over the plan is below	1,30	ratio	Draft
SENS_RED	Sensitivities	Red if ratio after worst shock is below	1,10	ratio	Draft
SENS_AMBER	Sensitivities	Amber if ratio after worst shock is below	1,30	ratio	Draft
REM_AMBER_DROP	Remittance	Amber if remittance vs previous plan falls by more than	0	share	Draft
REM_RED_DROP	Remittance	Red if remittance vs previous plan falls by more than	0,10	share	Draft
CN_AMBER_INC	Capital Needs	Amber if capital needs vs previous plan grow by more than	0	share	Draft
CN_RED_INC	Capital Needs	Red if capital needs vs previous plan grow by more than	0,10	share	Draft
RH_MIN_CASH	Remittance Holding	Red if holding cash at end of plan is below	0	EUR m	Draft
CHECK_TOL	All bridges	Red if a reconciliation check differs from zero by more than	1	EUR m	Draft
FTC_RED	Free Tangible Capital	Red if FTC in the last plan year is below	0	EUR m	Draft
SAA_AMBER_BREACH	SAA	Amber if number of asset classes outside bounds is at least	1	count	Draft
SAA_RED_BREACH	SAA	Red if number of asset classes outside bounds is at least	2	count	Draft
SAA_SUM_TOL	SAA	Amber if target weights differ from 100% by more than	0,001	share	Draft
LK_RED_DROP	Life KPIs	Red if VNP (net of tax) first to last year falls by more than	0,10	share	Draft
NL_COR_AMBER	Non-Life KPIs	Amber if net CoR (excl. GC&C) in last plan year is above	1,00	ratio	Draft
NL_COR_RED	Non-Life KPIs	Red if net CoR (excl. GC&C) in last plan year is above	1,05	ratio	Draft

>> Chiudi e applica. Таблица ни с чем не связана (отношения не нужны).


#####################################################################
 ШАГ 2. СПИСОК РАЗДЕЛОВ (SummaryIndex)
#####################################################################
>> Modellazione > Nuova tabella > вставить. Page = ТОЧНОЕ имя вкладки Power BI.
>> Совет: убери "(!)" из имён вкладок СЕЙЧАС и больше их не переименовывай —
>> иначе кнопки перестанут находить страницу.
>> ByCompany = показывать ли раздел на странице компаний (TRUE, если в данных есть Entity).

SummaryIndex =
DATATABLE (
    "Order", INTEGER, "Area", STRING, "Section", STRING, "Prefix", STRING, "Page", STRING, "ByCompany", BOOLEAN,
    {
        {  1, "Capital",    "Solvency",                "SOL",  "Solvency, Distributions & FTC", TRUE  },
        {  2, "Capital",    "Solvency AoM",            "AoM",  "Solvency AoM",                  TRUE  },
        {  3, "Capital",    "Sensitivities",           "SENS", "Sens on Net Result & Solvency", TRUE  },
        {  4, "Capital",    "Free Tangible Capital",   "FTC",  "Free Tangible Capital",         TRUE  },
        {  5, "Capital",    "Capital Needs",           "CN",   "Capital Needs",                 FALSE },
        {  6, "Remittance", "Total Remittance",        "TR",   "Total Remittance",              FALSE },
        {  7, "Remittance", "Remittance by Company",   "RBC",  "Remittance by Company",         TRUE  },
        {  8, "Remittance", "Remittance Holding",      "RH",   "Remittance Holding",            FALSE },
        {  9, "Remittance", "HMS Impacts",             "HMS",  "HMS impacts",                   FALSE },
        { 10, "Investments","Strategic Asset Allocation","SAA","SAA",                           TRUE  },
        { 11, "Business",   "Life KPIs",               "LK",   "Life KPIs",                     TRUE  },
        { 12, "Business",   "Non-Life KPIs",           "NL",   "Non-Life KPIs",                 TRUE  },
        { 13, "Funding",    "Debt Management",         "DEBT", "Debt Management",               FALSE }
    } )

>> Новый раздел потом = одна строка здесь + одна ветка в Summary Text / Summary Status.
>> SummaryIndex[Section] > Ordina per colonna > Order.


#####################################################################
 ШАГ 3. HEADLINE + STATUS ПО КАЖДОЙ СТРАНИЦЕ (Home > Nuova misura)
 Меры используют уже существующие меры страниц. Если какое-то имя меры
 подчёркнуто красным — у тебя оно называется иначе, напиши мне.
#####################################################################

----- SOLVENCY -------------------------------------------------------

SOL Min Ratio =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Solvency", NOT ISBLANK ( FactData[Year] ) )
RETURN MINX ( FILTER ( ADDCOLUMNS ( yrs, "@r", [SOL Ratio] ), NOT ISBLANK ( [@r] ) ), [@r] )

SOL Status =
VAR r   = [SOL Min Ratio]
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_RED" )
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_AMBER" )
RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), r < red, 3, r < amb, 2, 1 )

SOL Headline =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Solvency", NOT ISBLANK ( FactData[Year] ) )
VAR x   = FILTER ( ADDCOLUMNS ( yrs, "@r", [SOL Ratio] ), NOT ISBLANK ( [@r] ) )
VAR mn  = MINX ( x, [@r] )
VAR mx  = MAXX ( x, [@r] )
VAR mnY = MAXX ( TOPN ( 1, x, [@r], ASC ), FactData[Year] )
VAR y1  = MAXX ( x, FactData[Year] )
VAR r1  = MAXX ( FILTER ( x, FactData[Year] = y1 ), [@r] )
VAR p1  = CALCULATE ( [SOL Ratio Prev], FactData[Year] = y1 )
RETURN IF ( NOT ISBLANK ( mn ),
    "Solvency ratio " & FORMAT ( mn, "0%" ) & "–" & FORMAT ( mx, "0%" ) & " over the plan, lowest in " & mnY
    & IF ( NOT ISBLANK ( p1 ), "; " & y1 & ": " & FORMAT ( r1, "0%" ) & " vs " & FORMAT ( p1, "0%" ) & " in the previous plan" ) & "." )

----- SOLVENCY AoM ---------------------------------------------------

AoM Status =
VAR tol = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CHECK_TOL" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_RED" )
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SOL_AMBER" )
VAR r   = DIVIDE ( [AoM EOF EoP], [AoM SCR EoP] )
VAR ce  = ABS ( ROUND ( [AoM Check EOF], 0 ) )
VAR cs  = ABS ( ROUND ( [AoM Check SCR], 0 ) )
RETURN SWITCH ( TRUE (), ISBLANK ( r ), BLANK (), ce > tol || cs > tol, 3, r < red, 3, r < amb, 2, 1 )

AoM Headline =
VAR tol = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CHECK_TOL" )
VAR r0  = DIVIDE ( [AoM EOF BoP], [AoM SCR BoP] )
VAR r1  = DIVIDE ( [AoM EOF EoP], [AoM SCR EoP] )
VAR ok  = ABS ( ROUND ( [AoM Check EOF], 0 ) ) <= tol && ABS ( ROUND ( [AoM Check SCR], 0 ) ) <= tol
RETURN IF ( NOT ISBLANK ( r1 ),
    "Solvency ratio walk " & FORMAT ( r0, "0%" ) & " → " & FORMAT ( r1, "0%" )
    & " (" & FORMAT ( ( r1 - r0 ) * 100, "+0.0;-0.0;0.0" ) & " pp); "
    & IF ( ok, "EOF and SCR bridges reconcile.", "a bridge does NOT reconcile – check the page." ) )

----- SENSITIVITIES --------------------------------------------------
>> Sens Base Ratio и Sens Ratio pp — в процентных пунктах (182 = 182%), поэтому /100.

SENS Worst After =
VAR base = [Sens Base Ratio]
VAR sh   = CALCULATETABLE ( VALUES ( FactData[Sens Shock] ), FactData[Topic] = "Sensitivities",
               NOT ISBLANK ( FactData[Sens Shock] ) )
VAR w    = MINX ( FILTER ( ADDCOLUMNS ( sh, "@pp", [Sens Ratio pp] ), NOT ISBLANK ( [@pp] ) ), [@pp] )
RETURN IF ( NOT ISBLANK ( base ) && NOT ISBLANK ( w ), ( base + w ) / 100 )

SENS Status =
VAR a   = [SENS Worst After]
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SENS_RED" )
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SENS_AMBER" )
RETURN SWITCH ( TRUE (), ISBLANK ( a ), BLANK (), a < red, 3, a < amb, 2, 1 )

SENS Headline =
VAR base = [Sens Base Ratio]
VAR sh   = CALCULATETABLE ( VALUES ( FactData[Sens Shock] ), FactData[Topic] = "Sensitivities",
               NOT ISBLANK ( FactData[Sens Shock] ) )
VAR x    = FILTER ( ADDCOLUMNS ( sh, "@pp", [Sens Ratio pp] ), NOT ISBLANK ( [@pp] ) )
VAR wRow = TOPN ( 1, x, [@pp], ASC )
VAR wN   = MAXX ( wRow, FactData[Sens Shock] )
VAR wPP  = MAXX ( wRow, [@pp] )
RETURN IF ( NOT ISBLANK ( wPP ),
    "Worst sensitivity: " & wN & " " & FORMAT ( wPP, "+0.0;-0.0" ) & " pp, ratio after shock "
    & FORMAT ( ( base + wPP ) / 100, "0%" ) & " (base " & FORMAT ( base / 100, "0%" ) & ")." )

----- FREE TANGIBLE CAPITAL ------------------------------------------

FTC Status =
VAR tol = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CHECK_TOL" )
VAR lim = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "FTC_RED" )
VAR y1  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR f   = CALCULATE ( [FTC FTC], FactData[Year] = y1 )
VAR chk = ABS ( ROUND ( CALCULATE ( [FTC Check], FactData[Year] = y1 ), 0 ) )
RETURN SWITCH ( TRUE (), ISBLANK ( f ), BLANK (), chk > tol, 3, f < lim, 3, 1 )

FTC Headline =
VAR tol = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CHECK_TOL" )
VAR y0  = CALCULATE ( MIN ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR y1  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Free Tangible Capital" )
VAR f0  = CALCULATE ( [FTC FTC], FactData[Year] = y0 )
VAR f1  = CALCULATE ( [FTC FTC], FactData[Year] = y1 )
VAR chk = ROUND ( CALCULATE ( [FTC Check], FactData[Year] = y1 ), 0 )
RETURN IF ( NOT ISBLANK ( f1 ),
    "Free Tangible Capital " & FORMAT ( f0, "#,##0" ) & " in " & y0 & ", " & FORMAT ( f1, "#,##0" ) & " in " & y1
    & IF ( ABS ( chk ) <= tol, "; reconciliation closes.", "; reconciliation gap " & FORMAT ( chk, "+#,##0;-#,##0" ) & "." ) )

----- CAPITAL NEEDS --------------------------------------------------

CN Delta % =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Capital Needs", NOT ISBLANK ( FactData[Year] ) )
VAR t   = FILTER ( ADDCOLUMNS ( yrs, "@c", [CN Total Cur], "@p", [CN Total Prev] ), NOT ISBLANK ( [@p] ) && [@p] <> 0 )
RETURN DIVIDE ( SUMX ( t, [@c] ) - SUMX ( t, [@p] ), ABS ( SUMX ( t, [@p] ) ) )

CN Status =
VAR d   = [CN Delta %]
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CN_AMBER_INC" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "CN_RED_INC" )
VAR tot = CALCULATE ( [CN Total Cur], FactData[Topic] = "Capital Needs" )
RETURN SWITCH ( TRUE (), ISBLANK ( tot ), BLANK (), ISBLANK ( d ), 1, d > red, 3, d > amb, 2, 1 )

CN Headline =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Capital Needs", NOT ISBLANK ( FactData[Year] ) )
VAR tot = SUMX ( yrs, [CN Total Cur] )
VAR d   = [CN Delta %]
RETURN IF ( NOT ISBLANK ( tot ),
    "Capital needs " & MINX ( yrs, FactData[Year] ) & "–" & MAXX ( yrs, FactData[Year] ) & ": " & FORMAT ( tot, "#,##0" )
    & IF ( NOT ISBLANK ( d ), " (" & FORMAT ( d, "+0%;-0%;0%" ) & " vs previous plan)" ) & "." )

----- TOTAL REMITTANCE -----------------------------------------------

TR Delta % =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Remittance", NOT ISBLANK ( FactData[Year] ) )
VAR t   = FILTER ( ADDCOLUMNS ( yrs, "@c", ABS ( [TR Cur] ), "@p", ABS ( [TR Prev] ) ), [@p] <> 0 )
RETURN DIVIDE ( SUMX ( t, [@c] ) - SUMX ( t, [@p] ), SUMX ( t, [@p] ) )

TR Status =
VAR d   = [TR Delta %]
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_AMBER_DROP" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_RED_DROP" )
RETURN SWITCH ( TRUE (), ISBLANK ( d ), BLANK (), d < -red, 3, d < -amb, 2, 1 )

TR Headline =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Remittance", NOT ISBLANK ( FactData[Year] ) )
VAR x   = FILTER ( ADDCOLUMNS ( yrs, "@c", ABS ( [TR Cur] ) ), [@c] <> 0 )
VAR d   = [TR Delta %]
RETURN IF ( COUNTROWS ( x ) > 0,
    "Total remittance " & MINX ( x, FactData[Year] ) & "–" & MAXX ( x, FactData[Year] ) & ": "
    & FORMAT ( SUMX ( x, [@c] ), "#,##0" )
    & IF ( NOT ISBLANK ( d ), " (" & FORMAT ( d, "+0%;-0%;0%" ) & " vs previous plan)" ) & "." )

----- REMITTANCE BY COMPANY ------------------------------------------

RBC Delta % =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Remittance by Company", NOT ISBLANK ( FactData[Year] ) )
VAR t   = FILTER ( ADDCOLUMNS ( yrs, "@c", [RBC Total], "@p", [RBC Previous] ), NOT ISBLANK ( [@p] ) && [@p] <> 0 )
RETURN DIVIDE ( SUMX ( t, [@c] ) - SUMX ( t, [@p] ), ABS ( SUMX ( t, [@p] ) ) )

RBC Status =
VAR d   = [RBC Delta %]
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_AMBER_DROP" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "REM_RED_DROP" )
RETURN SWITCH ( TRUE (), ISBLANK ( d ), BLANK (), d < -red, 3, d < -amb, 2, 1 )

RBC Headline =
VAR yrs = CALCULATETABLE ( VALUES ( FactData[Year] ), FactData[Topic] = "Remittance by Company", NOT ISBLANK ( FactData[Year] ) )
VAR x   = FILTER ( ADDCOLUMNS ( yrs, "@c", [RBC Total] ), NOT ISBLANK ( [@c] ) && [@c] <> 0 )
VAR d   = [RBC Delta %]
RETURN IF ( COUNTROWS ( x ) > 0,
    "Remittance by company " & MINX ( x, FactData[Year] ) & "–" & MAXX ( x, FactData[Year] ) & ": "
    & FORMAT ( SUMX ( x, [@c] ), "#,##0" )
    & IF ( NOT ISBLANK ( d ), " (" & FORMAT ( d, "+0%;-0%;0%" ) & " vs previous plan)" ) & "." )

----- REMITTANCE HOLDING ---------------------------------------------

RH Cash End =
VAR y1 = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Remittance Holding" )
RETURN CALCULATE ( [Value Current], FactData[Topic] = "Remittance Holding",
           CONTAINSSTRING ( FactData[Item], "(EoP)" ), FactData[Year] = y1 )

RH Status =
VAR c   = [RH Cash End]
VAR lim = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "RH_MIN_CASH" )
RETURN SWITCH ( TRUE (), ISBLANK ( c ), BLANK (), c < lim, 3, 1 )

RH Headline =
VAR y0 = CALCULATE ( MIN ( FactData[Year] ), FactData[Topic] = "Remittance Holding" )
VAR y1 = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Remittance Holding" )
VAR b0 = CALCULATE ( [Value Current], FactData[Topic] = "Remittance Holding",
             CONTAINSSTRING ( FactData[Item], "(BoP)" ), FactData[Year] = y0 )
VAR e1 = [RH Cash End]
RETURN IF ( NOT ISBLANK ( e1 ),
    "Holding cash position " & FORMAT ( b0, "#,##0" ) & " at the start of " & y0 & ", "
    & FORMAT ( e1, "#,##0" ) & " at the end of " & y1 & "." )

----- HMS IMPACTS (информационно, без статуса) ----------------------

HMS Status = IF ( NOT ISBLANK ( CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts" ) ), 0 )

HMS Headline =
VAR v = CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts", FactData[IsTotal] = TRUE () )
VAR a = IF ( ISBLANK ( v ), CALCULATE ( [Value Current], FactData[Topic] = "HMS Impacts" ), v )
RETURN IF ( NOT ISBLANK ( a ), "HMS impacts on remittance: " & FORMAT ( a, "+#,##0;-#,##0;0" ) & " over the plan." )

----- SAA ------------------------------------------------------------

SAA Breaches =
VAR t = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Entity], FactData[SAA Part], FactData[SAA Line] ),
            FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation", FactData[SAA Order] <= 9 )
VAR x = ADDCOLUMNS ( t, "@p", [SAA L Pct], "@lo", [SAA L Lo], "@hi", [SAA L Hi] )
RETURN COUNTROWS ( FILTER ( x, NOT ISBLANK ( [@p] ) && [@p] <> 0
           && ( ( NOT ISBLANK ( [@lo] ) && [@p] < [@lo] ) || ( NOT ISBLANK ( [@hi] ) && [@p] > [@hi] ) ) ) ) + 0

>> сколько таблиц (компания x Life/Non-Life), где Target % не даёт 100%
SAA Sum Off =
VAR tol = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_SUM_TOL" )
VAR t = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Entity], FactData[SAA Part] ),
            FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation" )
VAR x = ADDCOLUMNS ( t, "@s", CALCULATE ( SUM ( FactData[Value] ),
            FactData[Topic] = "SAA", FactData[SAA Table] = "Allocation",
            FactData[SAA Col] = "Target %", FactData[SAA Order] <= 9 ) )
RETURN COUNTROWS ( FILTER ( x, NOT ISBLANK ( [@s] ) && [@s] <> 0 && ABS ( [@s] - 1 ) > tol ) ) + 0

SAA Status =
VAR tg  = [SAA Tgt MV]
VAR nB  = [SAA Breaches]
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_AMBER_BREACH" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "SAA_RED_BREACH" )
RETURN SWITCH ( TRUE (), ISBLANK ( tg ), BLANK (), nB >= red, 3, nB >= amb || [SAA Sum Off] > 0, 2, 1 )

SAA Headline =
VAR yS  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "SAA", FactData[PeriodType] = "SAA" )
VAR tg  = [SAA Tgt MV]
VAR ex  = [SAA MV Exp]
VAR nB  = [SAA Breaches]
VAR so  = [SAA Sum Off]
VAR scr = CALCULATE ( [SAA c Net SCR], REMOVEFILTERS ( FactData[SAA Line] ), FactData[SAA Order] <= 9 )
RETURN IF ( NOT ISBLANK ( tg ),
    "SAA " & yS & ": target investments " & FORMAT ( tg, "#,##0" ) & " (" & FORMAT ( tg - ex, "+#,##0;-#,##0;0" ) & " vs expected); "
    & IF ( nB = 0, "all weights within bounds", nB & " asset class(es) outside bounds" )
    & IF ( so > 0, ", weights do not add up to 100% in " & so & " table(s)" )
    & IF ( NOT ISBLANK ( scr ), "; net SCR from rebalancing " & FORMAT ( scr, "+#,##0;-#,##0;0" ) ) & "." )

----- LIFE KPIs ------------------------------------------------------

LK VNP First =
VAR yA = CALCULATE ( MIN ( FactData[Year] ), FactData[Topic] = "Life KPIs" )
RETURN CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
           CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yA )

LK VNP Last =
VAR yZ = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Life KPIs" )
RETURN CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Life KPIs", FactData[LK Table] = "NBV",
           CONTAINSSTRING ( FactData[Item], "VNP" ), FactData[Year] = yZ )

LK Status =
VAR a   = [LK VNP First]
VAR z   = [LK VNP Last]
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "LK_RED_DROP" )
VAR d   = DIVIDE ( z - a, ABS ( a ) )
RETURN SWITCH ( TRUE (), ISBLANK ( z ), BLANK (), ISBLANK ( d ), 1, d < -red, 3, d < 0, 2, 1 )

LK Headline =
VAR yA = CALCULATE ( MIN ( FactData[Year] ), FactData[Topic] = "Life KPIs" )
VAR yZ = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Life KPIs" )
VAR a  = [LK VNP First]
VAR z  = [LK VNP Last]
RETURN IF ( NOT ISBLANK ( z ),
    "Life SII VNP (net of tax) " & FORMAT ( a, "#,##0" ) & " in " & yA & ", " & FORMAT ( z, "#,##0" ) & " in " & yZ
    & IF ( a <> 0, " (" & FORMAT ( DIVIDE ( z - a, ABS ( a ) ), "+0%;-0%;0%" ) & ")" ) & "." )

----- NON-LIFE KPIs --------------------------------------------------

NL CoR Last =
VAR yZ = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
RETURN CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "CoR",
           CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )

NL Status =
VAR c   = [NL CoR Last]
VAR amb = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "NL_COR_AMBER" )
VAR red = LOOKUPVALUE ( Thresholds[Value], Thresholds[Key], "NL_COR_RED" )
RETURN SWITCH ( TRUE (), ISBLANK ( c ) || c = 0, BLANK (), c > red, 3, c > amb, 2, 1 )

NL Headline =
VAR yZ = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Non-Life KPIs", FactData[NL Group] <> "RoRC" )
VAR c  = [NL CoR Last]
VAR ep = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Non-Life KPIs", FactData[NL Group] = "Earned Premiums",
             CONTAINSSTRING ( FactData[Item], "Total excl" ), FactData[Measure] = "Net", FactData[Year] = yZ )
VAR si = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Non-Life KPIs",
             FactData[NL Group] = "SII CY Technical Result", FactData[Year] = yZ )
RETURN IF ( NOT ISBLANK ( c ) || NOT ISBLANK ( si ),
    "Non-Life " & yZ & ": "
    & IF ( NOT ISBLANK ( c ) && c <> 0, "net CoR (excl. GC&C) " & FORMAT ( c, "0.0%" ) & ", ", "" )
    & IF ( NOT ISBLANK ( ep ) && ep <> 0, "net earned premiums " & FORMAT ( ep, "#,##0" ) & ", ", "" )
    & "SII CY technical result " & FORMAT ( si, "#,##0" ) & "." )

----- DEBT MANAGEMENT (информационно) -------------------------------

DEBT Status = IF ( [Debt Instruments] > 0, 0 )

DEBT Headline =
IF ( [Debt Instruments] > 0,
     [Debt Instruments] & " instrument(s) in the proposals: reimbursements " & FORMAT ( [Debt Reimbursements], "#,##0" )
     & ", issuances " & FORMAT ( [Debt Issuances], "#,##0" ) & ", net " & FORMAT ( [Debt Net], "+#,##0;-#,##0;0" ) & ".",
     "No reimbursement or issuance proposals." )


#####################################################################
 ШАГ 4. СБОРКА SUMMARY
#####################################################################

Summary Text =
SWITCH ( SELECTEDVALUE ( SummaryIndex[Prefix] ),
    "SOL",  [SOL Headline],  "AoM",  [AoM Headline],  "SENS", [SENS Headline],
    "FTC",  [FTC Headline],  "CN",   [CN Headline],   "TR",   [TR Headline],
    "RBC",  [RBC Headline],  "RH",   [RH Headline],   "HMS",  [HMS Headline],
    "SAA",  [SAA Headline],  "LK",   [LK Headline],   "NL",   [NL Headline],
    "DEBT", [DEBT Headline] )

Summary Status =
SWITCH ( SELECTEDVALUE ( SummaryIndex[Prefix] ),
    "SOL",  [SOL Status],  "AoM",  [AoM Status],  "SENS", [SENS Status],
    "FTC",  [FTC Status],  "CN",   [CN Status],   "TR",   [TR Status],
    "RBC",  [RBC Status],  "RH",   [RH Status],   "HMS",  [HMS Status],
    "SAA",  [SAA Status],  "LK",   [LK Status],   "NL",   [NL Status],
    "DEBT", [DEBT Status] )

>> Цветная точка (в таблице/матрице)
Summary Dot = IF ( NOT ISBLANK ( [Summary Status] ), UNICHAR ( 9679 ) )

Summary Color =
SWITCH ( [Summary Status], 1, "#2E7D32", 2, "#E0A100", 3, "#C62828", 0, "#9E9E9E" )

Summary Label =
SWITCH ( [Summary Status], 1, "In line", 2, "Watch", 3, "Action needed", 0, "Info" )

>> Куда вести кнопку
Summary Nav = SELECTEDVALUE ( SummaryIndex[Page], "" )

Summary Button Text =
IF ( HASONEVALUE ( SummaryIndex[Section] ), "Open " & SELECTEDVALUE ( SummaryIndex[Section] ) & "  →",
     "Select a line to open its page" )

>> Шапка: общая картина одним абзацем
Summary Overview =
VAR ent = SELECTEDVALUE ( FactData[Entity], "Country total" )
VAR t   = ADDCOLUMNS ( SummaryIndex, "@s", [Summary Status] )
VAR nG  = COUNTROWS ( FILTER ( t, [@s] = 1 ) )
VAR nA  = COUNTROWS ( FILTER ( t, [@s] = 2 ) )
VAR nR  = COUNTROWS ( FILTER ( t, [@s] = 3 ) )
VAR nAll = COUNTROWS ( FILTER ( t, [@s] IN { 1, 2, 3 } ) )
VAR rTxt = CONCATENATEX ( FILTER ( t, [@s] = 3 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR aTxt = CONCATENATEX ( FILTER ( t, [@s] = 2 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
RETURN
ent & ": " & nG & " of " & nAll & " areas in line"
& IF ( nR > 0, UNICHAR ( 10 ) & "Action needed: " & rTxt & "." )
& IF ( nA > 0, UNICHAR ( 10 ) & "Watch: " & aTxt & "." )
& IF ( nR + nA = 0, ". No issues flagged against the approved thresholds." )

>> Проверка, что все пороги утверждены (показать мелко внизу страницы)
Summary Thresholds Note =
VAR n = COUNTROWS ( FILTER ( Thresholds, Thresholds[Status] <> "Approved" ) )
RETURN IF ( n > 0, n & " threshold(s) still in Draft – statuses are indicative.", "All thresholds approved." )


#####################################################################
 ШАГ 5. СТРАНИЦЫ
#####################################################################
>> ПОДГОТОВКА (1 раз): слайсер Entity (FactData[Entity]) на всех страницах отчёта
>> синхронизирован: Visualizza > Sincronizza filtri dei dati > выдели слайсер >
>> поставь галочки "Sincronizza" на всех страницах. Тогда переход со страницы
>> компании открывает раздел уже по этой компании.
>>
>> СТРАНИЦА "Summary – Total"
>>  1) Scheda (nuova) / Casella di testo с мерой Summary Overview — сверху.
>>  2) Tabella: SummaryIndex[Area], Summary Dot, SummaryIndex[Section], Summary Text,
>>     Summary Label. Summary Dot > Formattazione condizionale > Colore carattere >
>>     Valore campo > Summary Color. Сортировка по SummaryIndex[Order].
>>     Formato > Valori > A capo automatico: Attivato (длинные предложения).
>>  3) Pulsante (Inserisci > Pulsanti > Vuoto):
>>     Azione = Navigazione pagina; Destinazione > fx > Summary Nav.
>>     Testo pulsante > fx > Summary Button Text.
>>     Работа: клик по строке таблицы -> кнопка показывает "Open <раздел> →" -> клик.
>>  4) На этой странице слайсер Entity ОЧИЩЕН (тотал по стране; позже — по миру).
>>  5) Мелко внизу: Summary Thresholds Note.
>>
>> СТРАНИЦА "Summary – Companies"
>>  1) Слайсер Entity (синхронизированный), selezione singola.
>>  2) Те же Summary Overview + Tabella + кнопка (скопируй с Total) — показывают
>>     выбранную компанию. Фильтр визуала таблицы: SummaryIndex[ByCompany] = True.
>>  3) Тепловая карта всех компаний: Matrice
>>     Righe FactData[Entity]; Colonne SummaryIndex[Section]; Valori Summary Dot
>>     (цвет как выше); фильтр ByCompany = True.
>>     Выдели слайсер > Formato > Modifica interazioni > на матрице "Nessuna"
>>     (матрица всегда показывает все компании).
>>
>> КАК ПРОВЕРИТЬ
>>  - Summary – Total: 13 строк, у каждой точка и предложение; цифры в предложениях
>>    = цифры на страницах (Solvency min ratio, Remittance total, SII technical result...).
>>  - Поменяй в Thresholds, например, SOL_AMBER на 2,00 -> Solvency станет жёлтым. Верни.
>>  - Клик по строке "SAA" -> кнопка "Open Strategic Asset Allocation →" -> страница SAA.
