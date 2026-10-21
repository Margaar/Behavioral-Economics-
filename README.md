=====================================================================
 ПОСЛЕДНИЕ СТРАНИЦЫ: LOCAL RAF | OPTIONS MONITORING | ORMT | ORMT DETAILS | MODEL CHANGES
 + добавление их в Summary.
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ОБЩЕЕ
>> - Local RAF, Options, ORMT — это СПИСКИ (отдельные запросы List_..., не FactData).
>>   Имена таблиц в DAX пишем в кавычках: 'List_RAF' (урок с Thresholds).
>> - ORMT Details и Model Changes — из FactData (парсер их уже читает).
>> - Всё в EUR M / %, слайсер Scale на этих страницах не нужен.
>> - Перед стартом удали дубли вкладок: Local RAF x2 и Options Monitoring x2
>>   (правой кнопкой по вкладке > Elimina pagina), иначе навигация из Summary путается.
>> - В dummy пусты: Options, ORMT (список) и ORMT Details (все 0,0). Для проверки
>>   внеси в тестовую копию Excel 1-2 строки в каждый (ровно под заголовки!).


#####################################################################
 ШАГ 0. ПРОВЕРКА (5 минут)
#####################################################################
>> Trasforma dati > CheckBlocks: строки "ORMT Details" и "Model Changes" без ERROR.
>> FactData, фильтр Topic:
>>   ORMT Details  -> Section = "<ORMT #1>", "<ORMT #2>"...; Entity = Company #..
>>   Model Changes -> Block у строк SCR содержит "SCR" (у EOF — "EOF");
>>                    есть строки Item "EOF (BoP)" / "SCR (BoP)" (если есть во входе).
>> Если что-то не так — пришли фото этих колонок, поправлю одно правило.


#####################################################################
 A) LOCAL RAF
#####################################################################

>> A1. ЗАПРОС List_RAF (Nuova origine > Query vuota > Editor avanzato > вставить >
>>     имя List_RAF). Старый List_LocalRAF можно удалить.
>>     Он сам находит колонки по заголовкам (Hard Limit, Soft Limit, Lower, Upper —
>>     первая пара = Currently in place, вторая = Expected/Proposed).

let
    Sheet = Workbook{[Item = "Local RAF", Kind = "Sheet"]}[Data],
    Clean = Table.ReplaceErrorValues(Sheet, List.Transform(Table.ColumnNames(Sheet), each {_, null})),
    Rows  = List.Buffer(Table.ToRows(Clean)),
    NR    = List.Count(Rows),
    Txt   = (v) => if v = null then null
                   else Text.Trim(Text.Replace(Text.Replace(Text.From(v), "#(cr)", ""), "#(lf)", " ")),
    Hits  = (t as text) as list => List.Combine(List.Transform(List.Positions(Rows), (r) =>
                List.Transform(
                    List.Select(List.Positions(Rows{r}), (c) =>
                        let x = Txt(Rows{r}{c}) in x <> null and Text.StartsWith(Text.Upper(x), Text.Upper(t))),
                    (c) => [R = r, C = c, T = Txt(Rows{r}{c})]))),
    Long  = List.First(Hits("Long Description"), null),
    Pick  = (t as text, n as number) =>
                let h = List.Sort(List.Select(Hits(t), (x) => x[C] > Long[C] and Number.Abs(x[R] - Long[R]) <= 5),
                                  (x) => x[C])
                in if List.Count(h) > n then h{n} else null,
    H = [Code = Pick("Legal Entity", 0), Seg = Pick("Consolidation", 0), Reg = Pick("Solvency Regime", 0),
         CurHard = Pick("Hard Limit", 0),  NewHard = Pick("Hard Limit", 1),
         CurSoft = Pick("Soft Limit", 0),  NewSoft = Pick("Soft Limit", 1),
         CurLower = Pick("Lower", 0),      NewLower = Pick("Lower", 1),
         CurUpper = Pick("Upper", 0),      NewUpper = Pick("Upper", 1),
         SR1 = Pick("Solvency Ratio", 0),  SR2 = Pick("Solvency Ratio", 1)],
    HdrBot = List.Max(List.Transform(List.RemoveNulls(Record.FieldValues(H)), each [R]) & {Long[R]}),
    Body   = List.Numbers(HdrBot + 1, List.Max({0, List.Min({40, NR - HdrBot - 1})})),
    Cell   = (r as number, h) => if h = null then null else Rows{r}{h[C]},
    Num    = (v) => if v is number then v else try Number.From(v) otherwise null,
    Recs   = List.Transform(Body, (r) => [
                 Company  = Txt(Rows{r}{Long[C]}),
                 Code     = Txt(Cell(r, H[Code])),
                 Segment  = Txt(Cell(r, H[Seg])),
                 Regime   = Txt(Cell(r, H[Reg])),
                 CurHard  = Num(Cell(r, H[CurHard])),  CurSoft  = Num(Cell(r, H[CurSoft])),
                 CurLower = Num(Cell(r, H[CurLower])), CurUpper = Num(Cell(r, H[CurUpper])),
                 NewHard  = Num(Cell(r, H[NewHard])),  NewSoft  = Num(Cell(r, H[NewSoft])),
                 NewLower = Num(Cell(r, H[NewLower])), NewUpper = Num(Cell(r, H[NewUpper])),
                 SR1 = Num(Cell(r, H[SR1])), SR2 = Num(Cell(r, H[SR2])),
                 SR1Label = if H[SR1] = null then null else H[SR1][T],
                 SR2Label = if H[SR2] = null then null else H[SR2][T]]),
    Kept   = List.Select(Recs, each [Company] <> null and [Company] <> "" and [Company] <> "0"),
    Result = Table.FromRecords(Kept, type table [
                 Company = text, Code = text, Segment = text, Regime = text,
                 CurHard = number, CurSoft = number, CurLower = number, CurUpper = number,
                 NewHard = number, NewSoft = number, NewLower = number, NewUpper = number,
                 SR1 = number, SR2 = number, SR1Label = text, SR2Label = text])
in
    Result

>> Проверка (dummy): 3 строки ES003 / ES005 / ES006; CurHard 1,1; CurSoft 1,3;
>> CurLower 1,3; CurUpper 1,7 / 1,8 / 1,5; SR1 1,77 / 1,95 / 2,1; SR1Label "Solvency Ratio For 2026".
>> Chiudi e applica.

>> A2. КОЛОНКИ в List_RAF (Visualizzazione tabella > List_RAF > Nuova colonna)

RAF Min SR =
VAR a = 'List_RAF'[SR1]
VAR b = 'List_RAF'[SR2]
RETURN IF ( ISBLANK ( a ), b, IF ( ISBLANK ( b ), a, MIN ( a, b ) ) )

RAF Change =
VAR d =
    ( NOT ISBLANK ( 'List_RAF'[NewHard] )  && 'List_RAF'[NewHard]  <> 'List_RAF'[CurHard] )
 || ( NOT ISBLANK ( 'List_RAF'[NewSoft] )  && 'List_RAF'[NewSoft]  <> 'List_RAF'[CurSoft] )
 || ( NOT ISBLANK ( 'List_RAF'[NewLower] ) && 'List_RAF'[NewLower] <> 'List_RAF'[CurLower] )
 || ( NOT ISBLANK ( 'List_RAF'[NewUpper] ) && 'List_RAF'[NewUpper] <> 'List_RAF'[CurUpper] )
RETURN IF ( d, "Change proposed", "No change" )

RAF Position =
VAR m = 'List_RAF'[RAF Min SR]
RETURN SWITCH ( TRUE (),
    ISBLANK ( m ), "n.a.",
    m < 'List_RAF'[CurHard], "Below hard limit",
    m < 'List_RAF'[CurSoft], "Below soft limit",
    NOT ISBLANK ( 'List_RAF'[CurLower] ) && m < 'List_RAF'[CurLower], "Below target range",
    NOT ISBLANK ( 'List_RAF'[CurUpper] ) && m > 'List_RAF'[CurUpper], "Above target range",
    "Within target range" )

>> Формат: CurHard ... NewUpper, SR1, SR2, RAF Min SR -> Strumenti colonna > Percentuale, 0 decimali.

>> A3. МЕРЫ (Home > Nuova misura; Tabella home = _Measures, папка "Local RAF")

RAF Color =
SWITCH ( SELECTEDVALUE ( 'List_RAF'[RAF Position] ),
    "Below hard limit",   "#F4C7C3",
    "Below soft limit",   "#FCE8B2",
    "Below target range", "#FCE8B2",
    "Above target range", "#D9E8F5",
    "Within target range","#D9EAD3" )

RAF Comment =
VAR n     = COUNTROWS ( 'List_RAF' )
VAR oneC  = ISFILTERED ( 'List_RAF'[Company] ) && n = 1
VAR chg   = FILTER ( 'List_RAF', 'List_RAF'[RAF Change] = "Change proposed" )
VAR nChg  = COUNTROWS ( chg ) + 0
VAR chgTxt = CONCATENATEX ( chg,
              'List_RAF'[Code] & ": hard " & FORMAT ( 'List_RAF'[CurHard], "0%" ) & " → "
              & FORMAT ( COALESCE ( 'List_RAF'[NewHard], 'List_RAF'[CurHard] ), "0%" )
              & ", soft " & FORMAT ( 'List_RAF'[CurSoft], "0%" ) & " → "
              & FORMAT ( COALESCE ( 'List_RAF'[NewSoft], 'List_RAF'[CurSoft] ), "0%" )
              & ", target range " & FORMAT ( 'List_RAF'[CurLower], "0%" ) & "–" & FORMAT ( 'List_RAF'[CurUpper], "0%" )
              & " → " & FORMAT ( COALESCE ( 'List_RAF'[NewLower], 'List_RAF'[CurLower] ), "0%" ) & "–"
              & FORMAT ( COALESCE ( 'List_RAF'[NewUpper], 'List_RAF'[CurUpper] ), "0%" ), "; " )
VAR hd    = FILTER ( ADDCOLUMNS ( 'List_RAF', "@h", 'List_RAF'[RAF Min SR] - 'List_RAF'[CurSoft] ), NOT ISBLANK ( [@h] ) )
VAR lowR  = TOPN ( 1, hd, [@h], ASC )
VAR lowC  = MAXX ( lowR, 'List_RAF'[Code] )
VAR lowSR = MAXX ( lowR, 'List_RAF'[RAF Min SR] )
VAR lowSf = MAXX ( lowR, 'List_RAF'[CurSoft] )
VAR lowH  = MAXX ( lowR, [@h] )
VAR bHard = FILTER ( 'List_RAF', 'List_RAF'[RAF Position] = "Below hard limit" )
VAR bSoft = FILTER ( 'List_RAF', 'List_RAF'[RAF Position] = "Below soft limit" )
VAR bTgt  = FILTER ( 'List_RAF', 'List_RAF'[RAF Position] = "Below target range" )
VAR aTgt  = FILTER ( 'List_RAF', 'List_RAF'[RAF Position] = "Above target range" )
-- page mode
VAR L1 = n & IF ( n = 1, " company", " companies" ) & " in the Local RAF table; "
         & IF ( nChg = 0, "current limits and target ranges are confirmed for the plan.",
                "changes proposed for " & nChg & ": " & chgTxt & "." )
VAR L2 = IF ( NOT ISBLANK ( lowH ),
         "Lowest headroom vs soft limit: " & lowC & " " & FORMAT ( lowSR, "0%" ) & " vs " & FORMAT ( lowSf, "0%" )
         & " (" & FORMAT ( lowH * 100, "+0;-0;0" ) & " pp)." )
VAR L3 = SWITCH ( TRUE (),
         COUNTROWS ( bHard ) > 0, "Below hard limit: " & CONCATENATEX ( bHard, 'List_RAF'[Code], ", " ) & ".",
         COUNTROWS ( bSoft ) > 0, "Below soft limit: " & CONCATENATEX ( bSoft, 'List_RAF'[Code], ", " ) & ".",
         "All companies are above their soft limits over the plan." )
VAR L4 = IF ( COUNTROWS ( aTgt ) > 0, "Above the upper target threshold: "
         & CONCATENATEX ( aTgt, 'List_RAF'[Code] & " (" & FORMAT ( 'List_RAF'[RAF Min SR], "0%" ) & " vs "
                                & FORMAT ( 'List_RAF'[CurUpper], "0%" ) & ")", ", " ) & " – potential excess capital." )
VAR L5 = IF ( COUNTROWS ( bTgt ) > 0, "Below the lower target threshold: " & CONCATENATEX ( bTgt, 'List_RAF'[Code], ", " ) & "." )
-- one company clicked in the table
VAR c  = SELECTEDVALUE ( 'List_RAF'[Code] )
VAR C1 = SELECTEDVALUE ( 'List_RAF'[Company] ) & " (" & c & ", " & SELECTEDVALUE ( 'List_RAF'[Segment] ) & "): hard limit "
         & FORMAT ( MAX ( 'List_RAF'[CurHard] ), "0%" ) & ", soft limit " & FORMAT ( MAX ( 'List_RAF'[CurSoft] ), "0%" )
         & ", operating target range " & FORMAT ( MAX ( 'List_RAF'[CurLower] ), "0%" ) & "–" & FORMAT ( MAX ( 'List_RAF'[CurUpper] ), "0%" ) & "."
VAR C2 = SELECTEDVALUE ( 'List_RAF'[SR1Label] ) & ": " & FORMAT ( MAX ( 'List_RAF'[SR1] ), "0%" ) & ", "
         & SELECTEDVALUE ( 'List_RAF'[SR2Label] ) & ": " & FORMAT ( MAX ( 'List_RAF'[SR2] ), "0%" )
         & " – " & LOWER ( SELECTEDVALUE ( 'List_RAF'[RAF Position] ) ) & "."
VAR C3 = IF ( nChg = 0, "No change to limits proposed.", "Proposed: " & chgTxt & "." )
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ) }
RETURN
SWITCH ( TRUE (),
    ISBLANK ( n ) || n = 0, "No companies in the Local RAF table.",
    oneC, C1 & UNICHAR ( 10 ) & C2 & UNICHAR ( 10 ) & C3,
    CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> A4. СТРАНИЦА "Local RAF" (без фильтра Topic, без Scale)
>> 1) Tabella: Company, Code, Segment, Regime, CurHard, CurSoft, CurLower, CurUpper,
>>    NewHard, NewSoft, NewLower, NewUpper, SR1, SR2, RAF Position, RAF Change.
>>    Числовые: стрелка > Non riepilogare. Переименуй как в Excel
>>    ("Current – Hard Limit", ..., "Proposed – Upper Threshold", "Solvency Ratio For 2026"...).
>>    SR1 и SR2: Formattazione condizionale > Colore sfondo > Valore campo > RAF Color.
>>    Totali: Disattivato.
>> 2) Comment box: RAF Comment (клик по строке таблицы -> текст по этой компании).
>> ПРОВЕРКА (dummy): 3 компании, изменений нет, все "Above target range"
>>    (ES003 177% > 170%, ES005 195% > 180%, ES006 210% > 150%), lowest headroom ES003 +47 pp.


#####################################################################
 B) OPTIONS MONITORING
#####################################################################

>> B1. ЗАПРОС List_Options — замени код существующего запроса целиком
>>     (читает левую таблицу, колонкам даёт короткие имена).

let
    Sheet = Workbook{[Item = "Options Monitoring", Kind = "Sheet"]}[Data],
    Raw   = try fnReadList(Sheet, null, "ID", "Option Buyer", 3, 1, 40, 0)
            otherwise fnReadList(Sheet, null, null, "Option Buyer", 3, 1, 40, 0),
    Map = (n as text) as text =>
        let u = Text.Upper(n) in
        if u = "LABEL" or Text.Contains(u, "BRIEF") then "Description"
        else if Text.Contains(u, "OPTION TYPE") then "OptionType"
        else if Text.Contains(u, "EXPIRATION") then "Expiration"
        else if Text.Contains(u, "EXPECTED EXERCISE") then "ExpectedExercise"
        else if Text.Contains(u, "INCLUDED IN") then "InPlan"
        else if Text.Contains(u, "FINANCIAL IMPACT") then "FinImpact"
        else if Text.Contains(u, "IMPACT ON SOLVENCY") then "SolvencyImpact"
        else if Text.Contains(u, "BEGINNING") then "ExerciseFrom"
        else if Text.Contains(u, "END OF PERIOD") then "ExerciseTo"
        else if Text.Contains(u, "OPTION BUYER") and Text.Contains(u, "IFICC") then "BuyerIFICC"
        else if Text.Contains(u, "OPTION BUYER") then "Buyer"
        else if Text.Contains(u, "COUNTERPARTY") and Text.Contains(u, "IFICC") then "CounterpartyIFICC"
        else if Text.Contains(u, "COUNTERPARTY") then "Counterparty"
        else if Text.EndsWith(u, "PERIOD") then "ImpactPeriod"
        else n,
    Old = Table.ColumnNames(Raw),
    New = List.Accumulate(Old, {}, (s, n) =>
              let m = Map(n) in
              s & { if not List.Contains(s, m) then m
                    else if not List.Contains(s, n) then n
                    else n & " #" & Text.From(List.Count(s)) }),
    Ren = Table.RenameColumns(Raw, List.Zip({Old, New})),
    ToNum = (v) => if v is number then v else try Number.From(v) otherwise null,
    Result = Table.TransformColumns(Ren,
                 {{"FinImpact", ToNum, type number}, {"SolvencyImpact", ToNum, type number}}, null, MissingField.Ignore)
in
    Result

>> Проверка: колонки Description, Buyer, BuyerIFICC, Counterparty, CounterpartyIFICC,
>> OptionType, ExerciseFrom, ExerciseTo, Expiration, ExpectedExercise, InPlan,
>> ImpactPeriod, FinImpact, SolvencyImpact. Если какая-то осталась с длинным
>> именем — пришли фото, добавлю правило. В dummy строк 0 — нормально.

>> B2. МЕРЫ (папка "Options")

OPT Count = COUNTROWS ( 'List_Options' ) + 0

OPT Fin Impact = SUM ( 'List_Options'[FinImpact] )

OPT Comment =
VAR n    = [OPT Count]
VAR yes  = FILTER ( 'List_Options', UPPER ( LEFT ( 'List_Options'[ExpectedExercise], 1 ) ) = "Y" )
VAR nYes = COUNTROWS ( yes ) + 0
VAR gap  = FILTER ( yes, UPPER ( LEFT ( 'List_Options'[InPlan], 1 ) ) = "N" )
VAR nGap = COUNTROWS ( gap ) + 0
VAR types = CONCATENATEX ( VALUES ( 'List_Options'[OptionType] ),
               'List_Options'[OptionType] & " " & CALCULATE ( COUNTROWS ( 'List_Options' ) ), ", " )
VAR big  = TOPN ( 1, FILTER ( 'List_Options', NOT ISBLANK ( 'List_Options'[FinImpact] ) ), ABS ( 'List_Options'[FinImpact] ), DESC )
VAR L1 = n & IF ( n = 1, " option", " options" ) & " reported (" & types & ")."
VAR L2 = nYes & " expected to be exercised over the plan"
         & IF ( nGap > 0, "; " & nGap & " of them NOT included in the local capital management plan: "
                & CONCATENATEX ( gap, 'List_Options'[Description], "; " ), "" ) & "."
VAR L3 = IF ( NOT ISBLANK ( [OPT Fin Impact] ),
         "Total expected financial impact: " & FORMAT ( [OPT Fin Impact], "+#,##0.0;-#,##0.0;0" ) & " (EUR m)"
         & "; largest: " & MAXX ( big, 'List_Options'[Description] ) & " "
         & FORMAT ( MAXX ( big, 'List_Options'[FinImpact] ), "+#,##0.0;-#,##0.0;0" ) & "." )
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ) }
RETURN
IF ( n = 0, "No options reported.",
     CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> B3. СТРАНИЦА "Options Monitoring": Tabella со всеми колонками (Non riepilogare,
>> подписи как в Excel), карточки OPT Count и OPT Fin Impact, Comment box OPT Comment.


#####################################################################
 C) ORMT (список транзакций)
#####################################################################

>> C1. ЗАПРОС List_ORMT — замени код целиком

let
    Sheet = Workbook{[Item = "ORMT", Kind = "Sheet"]}[Data],
    Raw   = fnReadList(Sheet, null, null, "Cedant", 1, 1, 60, 0),
    Map = (n as text) as text =>
        let u = Text.Upper(n) in
        if Text.StartsWith(u, "ID") then "ID"
        else if Text.StartsWith(u, "TYPE OF") then "IGType"
        else if u = "TYPE" then "Type"
        else if Text.StartsWith(u, "CEDANT") then "Cedant"
        else if Text.StartsWith(u, "REINSURER") then "Reinsurer"
        else if Text.Contains(u, "SUB TYPE") then "SubType"
        else if Text.Contains(u, "LINE OF BUSINESS") then "LoB"
        else if Text.Contains(u, "CLASSIFICATION") then "Classification"
        else if Text.Contains(u, "EXECUTION DATE") then "ExecDate"
        else if Text.Contains(u, "TIME OF EXECUTION") then "ExecValueEUR"
        else if Text.Contains(u, "APPROVED") and Text.Contains(u, "LOCAL") then "ApprovedLC"
        else if Text.Contains(u, "APPROVED") then "Approved"
        else if Text.StartsWith(u, "LOCAL") then "Currency"
        else if Text.StartsWith(u, "DESCRIPTION") then "Description"
        else if Text.StartsWith(u, "NOTES") then "Notes"
        else n,
    Old = Table.ColumnNames(Raw),
    New = List.Accumulate(Old, {}, (s, n) =>
              let m = Map(n) in
              s & { if not List.Contains(s, m) then m
                    else if not List.Contains(s, n) then n
                    else n & " #" & Text.From(List.Count(s)) }),
    Ren = Table.RenameColumns(Raw, List.Zip({Old, New})),
    ToNum = (v) => if v is number then v else try Number.From(v) otherwise null,
    Result = Table.TransformColumns(Ren,
                 {{"Approved", ToNum, type number}, {"ApprovedLC", ToNum, type number},
                  {"ExecValueEUR", ToNum, type number}}, null, MissingField.Ignore)
in
    Result

>> C2. МЕРЫ (папка "ORMT")

ORMTL Count = COUNTROWS ( 'List_ORMT' ) + 0

ORMTL Value = SUM ( 'List_ORMT'[ExecValueEUR] )

ORMTL Comment =
VAR n    = [ORMTL Count]
VAR byT  = CONCATENATEX ( VALUES ( 'List_ORMT'[Type] ),
               'List_ORMT'[Type] & " " & CALCULATE ( COUNTROWS ( 'List_ORMT' ) ), ", " )
VAR big  = TOPN ( 1, FILTER ( 'List_ORMT', NOT ISBLANK ( 'List_ORMT'[ExecValueEUR] ) ), 'List_ORMT'[ExecValueEUR], DESC )
VAR L1 = n & IF ( n = 1, " intragroup risk mitigation transaction", " intragroup risk mitigation transactions" )
         & " executed/renewed in the year (" & byT & ")."
VAR L2 = IF ( NOT ISBLANK ( [ORMTL Value] ),
         "Total value at execution: " & FORMAT ( [ORMTL Value], "#,##0.0" ) & " (EUR m); largest: "
         & MAXX ( big, 'List_ORMT'[Cedant] ) & " → " & MAXX ( big, 'List_ORMT'[Reinsurer] )
         & " (" & MAXX ( big, 'List_ORMT'[SubType] ) & ", " & FORMAT ( MAXX ( big, 'List_ORMT'[ExecValueEUR] ), "#,##0.0" ) & ")." )
VAR pg = { ( 1, L1 ), ( 2, L2 ) }
RETURN
IF ( n = 0, "No ORMT transactions executed or renewed in the year.",
     CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )

>> C3. СТРАНИЦА "ORMT": Tabella (ID, Type, Cedant, Reinsurer, IGType, SubType, LoB,
>> Classification, Currency, Approved, ApprovedLC, ExecValueEUR, ExecDate, Description, Notes),
>> карточки ORMTL Count и ORMTL Value, Comment box ORMTL Comment.


#####################################################################
 D) ORMT DETAILS (FactData, Topic = "ORMT Details")
#####################################################################

>> D1. КОЛОНКИ в FactData

ORMT Name =
IF ( FactData[Topic] = "ORMT Details",
     SUBSTITUTE ( SUBSTITUTE ( FactData[Section], "<", "" ), ">", "" ) )

ORMT RelRow =
IF ( FactData[Topic] = "ORMT Details",
     VAR e = FactData[Entity]
     VAR s = FactData[Section]
     RETURN FactData[RowNo]
            - MINX ( FILTER ( FactData, FactData[Topic] = "ORMT Details"
                       && FactData[Entity] = e && FactData[Section] = s ), FactData[RowNo] ) )

ORMT Line = IF ( FactData[Topic] = "ORMT Details", FactData[Item] )

ORMT Order =
IF ( FactData[Topic] = "ORMT Details",
     VAR it = FactData[Item]
     RETURN MINX ( FILTER ( FactData, FactData[Topic] = "ORMT Details" && FactData[Item] = it ), FactData[ORMT RelRow] ) )

>> Сортировка: ORMT Line > Ordina per colonna > ORMT Order.

>> D2. МЕРЫ (папка "ORMT Details")

ORMT Value = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "ORMT Details" )

ORMT Base EOF =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[Section], FactData[ORMT Name], FactData[ORMT Line], FactData[Item] ),
    FactData[Topic] = "Solvency", FactData[Scenario] = "Current",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Eligible Own Funds" ) )

ORMT Base SCR =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( DimItem ),
    REMOVEFILTERS ( FactData[Section], FactData[ORMT Name], FactData[ORMT Line], FactData[Item] ),
    FactData[Topic] = "Solvency", FactData[Scenario] = "Current",
    FactData[PeriodType] IN { "Actual", "Forecast", "Plan", "Year" },
    CONTAINSSTRING ( FactData[Item], "Solvency Capital Requirement" ) )

ORMT dEOF =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[ORMT Line], FactData[Item] ),
    FactData[Topic] = "ORMT Details", CONTAINSSTRING ( FactData[Item], "Impact on Eligible Own Funds" ) )

ORMT dSCR =
CALCULATE ( SUM ( FactData[Value] ),
    REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[ORMT Line], FactData[Item] ),
    FactData[Topic] = "ORMT Details", CONTAINSSTRING ( FactData[Item], "Impact on Solvency Capital Requirement" ) )

>> Влияние на Solvency Ratio в п.п.: (EOF + dEOF) / (SCR + dSCR) - EOF / SCR
ORMT pp =
VAR e  = [ORMT Base EOF]
VAR s  = [ORMT Base SCR]
VAR de = COALESCE ( [ORMT dEOF], 0 )
VAR ds = COALESCE ( [ORMT dSCR], 0 )
RETURN IF ( NOT ISBLANK ( s ) && s <> 0, ( DIVIDE ( e + de, s + ds ) - DIVIDE ( e, s ) ) * 100 )

>> Влияние на LGAAP Net Result: строка из входа, если есть; иначе -(Ceded Result)
ORMT NR =
VAR nr = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[ORMT Line], FactData[Item] ),
             FactData[Topic] = "ORMT Details", CONTAINSSTRING ( FactData[Item], "Impact on LGAAP Net Result" ) )
VAR cr = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[ORMT Line], FactData[Item] ),
             FactData[Topic] = "ORMT Details", FactData[Item] = "Ceded Result" )
RETURN IF ( NOT ISBLANK ( nr ) && nr <> 0, nr, IF ( NOT ISBLANK ( cr ), - cr ) )

>> Формат: ORMT pp -> Dinamico не нужен, Formato personalizzato  +0.0" pp";-0.0" pp";0.0" pp"
>>         ORMT Value, ORMT NR -> #,##0.0

ORMT Comment =
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR yrs  = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Year], FactData[Period] ),
               REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
               FactData[Topic] = "ORMT Details", NOT ISBLANK ( FactData[Year] ) )
VAR yZ   = MAXX ( yrs, FactData[Year] )
VAR pZ   = MAXX ( FILTER ( yrs, FactData[Year] = yZ ), FactData[Period] )
VAR secs = CALCULATETABLE ( VALUES ( FactData[ORMT Name] ),
               REMOVEFILTERS ( FactData[ORMT Line] ),
               FactData[Topic] = "ORMT Details", NOT ISBLANK ( FactData[ORMT Name] ) )
VAR x    = ADDCOLUMNS ( secs,
               "@any", CALCULATE ( SUMX ( FactData, ABS ( FactData[Value] ) ),
                           REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[ORMT Line] ),
                           FactData[Topic] = "ORMT Details" ),
               "@pp", CALCULATE ( [ORMT pp], REMOVEFILTERS ( DimPeriod ),
                           REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Year] = yZ ),
               "@nr", CALCULATE ( [ORMT NR], REMOVEFILTERS ( DimPeriod ),
                           REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Year] = yZ ) )
VAR act  = FILTER ( x, [@any] > 0 )
VAR nA   = COUNTROWS ( act ) + 0
VAR ppAll = CALCULATE ( [ORMT pp], REMOVEFILTERS ( FactData[ORMT Name], FactData[Section], FactData[ORMT Line] ),
               REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ), FactData[Year] = yZ )
VAR L1 = ent & ": " & IF ( nA = 0, "no ORMT / reinsurance impacts reported (all values zero).",
                            nA & " ORMT(s) with impacts reported." )
VAR L2 = IF ( nA > 0, "Combined impact on the Solvency Ratio in " & pZ & ": "
              & FORMAT ( ppAll, "+0.0;-0.0;0.0" ) & " pp." )
VAR L3 = IF ( nA > 0, CONCATENATEX ( act, FactData[ORMT Name] & ": " & FORMAT ( [@pp], "+0.0;-0.0;0.0" )
              & " pp, LGAAP net result " & FORMAT ( [@nr], "+#,##0.0;-#,##0.0;0.0" ), "; " ) & " (" & pZ & ")." )
VAR L4 = IF ( nA > 0, "Positive ceded result = negative impact on the cedant LGAAP P&L." )
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ) }
RETURN CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC )

>> D3. СТРАНИЦА "ORMT Details" — БЕЗ фильтра страницы Topic (нужна и Solvency)
>> 0) Слайсер Entity (FactData[Entity], selezione singola) — скопируй с другой страницы.
>> 1) Верхняя таблица KPI (EOF, SCR, Ratio, Distributions, Net Result, Distribution Ratio):
>>    скопируй готовую матрицу со страницы Solvency (Ctrl+C / Ctrl+V).
>> 2) Matrix "ORMT and Reinsurance Impacts":
>>    Righe: ORMT Name, ORMT Line  |  Colonne: DimPeriod[Period]  |  Valori: ORMT Value
>>    Filtro visual: Topic = ORMT Details. Totali: Disattivato. Expand all (⇊).
>>    Если хочешь как Excel "Hide rows equal to zero": фильтр визуала ORMT Value  ≠ 0.
>> 3) Matrix "Impact on Solvency Ratio and Net Result":
>>    Righe: ORMT Name  |  Colonne: DimPeriod[Period]  |  Valori: ORMT pp, ORMT NR
>>    Filtro visual: Topic = ORMT Details.
>> 4) Comment box: ORMT Comment.
>> ПРОВЕРКА: в dummy всё 0 -> "no ORMT / reinsurance impacts reported".
>>   Тест: Company #1, <ORMT #1>, Act 2025: Impact on EOF = -50, Impact on SCR = -100
>>   -> ORMT pp = (1654,5-50)/(953,2-100) - 1654,5/953,2 = 188,1% - 173,6% = +14,5 pp.


#####################################################################
 E) MODEL CHANGES DETAILS (FactData, Topic = "Model Changes")
#####################################################################

>> E1. КОЛОНКИ в FactData

MC Table =
IF ( FactData[Topic] = "Model Changes",
     SWITCH ( TRUE (),
         CONTAINSSTRING ( FactData[Item], "(BoP)" ), "BoP",
         CONTAINSSTRING ( FactData[Block], "SCR" ), "SCR",
         "EOF" ) )

>> Если в ШАГЕ 0 у SCR-строк Block НЕ содержит "SCR" — пришли фото, дам вариант по позиции строк.

MC Line =
IF ( FactData[Topic] = "Model Changes",
     FactData[Item] & IF ( FactData[MC Table] = "SCR", UNICHAR ( 8203 ) ) )

MC RelRow =
IF ( FactData[Topic] = "Model Changes",
     VAR e = FactData[Entity]
     VAR t = FactData[MC Table]
     RETURN FactData[RowNo]
            - MINX ( FILTER ( FactData, FactData[Topic] = "Model Changes"
                       && FactData[Entity] = e && FactData[MC Table] = t ), FactData[RowNo] ) )

MC Order =
IF ( FactData[Topic] = "Model Changes",
     VAR l = FactData[MC Line]
     RETURN MINX ( FILTER ( FactData, FactData[Topic] = "Model Changes" && FactData[MC Line] = l ), FactData[MC RelRow] ) )

>> Сортировка: MC Line > Ordina per colonna > MC Order.

>> E2. МЕРЫ (папка "Model Changes")

MC Value = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Model Changes" )

>> BoP: из листа Model Changes (EOF (BoP)/SCR (BoP)); если там пусто — со страницы Solvency AoM
MC Base EOF =
VAR b = CALCULATE ( SUM ( FactData[Value] ),
            REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
            FactData[Topic] = "Model Changes", CONTAINSSTRING ( FactData[Item], "EOF (BoP)" ) )
RETURN IF ( NOT ISBLANK ( b ) && b <> 0, b,
            CALCULATE ( [AoM EOF BoP], REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ) ) )

MC Base SCR =
VAR b = CALCULATE ( SUM ( FactData[Value] ),
            REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
            FactData[Topic] = "Model Changes", CONTAINSSTRING ( FactData[Item], "SCR (BoP)" ) )
RETURN IF ( NOT ISBLANK ( b ) && b <> 0, b,
            CALCULATE ( [AoM SCR BoP], REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ) ) )

>> Влияние строки (категории) на Solvency Ratio, п.п.: та же категория в EOF и в SCR
MC pp =
VAR k  = SUBSTITUTE ( SELECTEDVALUE ( FactData[MC Line] ), UNICHAR ( 8203 ), "" )
VAR de = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
             FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF", FactData[Item] = k )
VAR ds = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
             FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR", FactData[Item] = k )
VAR e  = [MC Base EOF]
VAR s  = [MC Base SCR]
RETURN IF ( NOT ISBLANK ( s ) && s <> 0 && NOT ISBLANK ( k ),
            ( DIVIDE ( e + COALESCE ( de, 0 ), s + COALESCE ( ds, 0 ) ) - DIVIDE ( e, s ) ) * 100 )

MC pp Total =
VAR de = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
             FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
             CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR ds = CALCULATE ( SUM ( FactData[Value] ),
             REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
             FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR",
             CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR e  = [MC Base EOF]
VAR s  = [MC Base SCR]
RETURN IF ( NOT ISBLANK ( s ) && s <> 0,
            ( DIVIDE ( e + COALESCE ( de, 0 ), s + COALESCE ( ds, 0 ) ) - DIVIDE ( e, s ) ) * 100 )

>> Проверка: "Regulatory and model changes" = сумма трёх подстрок (по каждой компании и году)
MC Check =
VAR t = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Entity], FactData[Year] ),
            REMOVEFILTERS ( FactData[MC Line], FactData[Item] ), FactData[Topic] = "Model Changes" )
RETURN
SUMX ( t,
    VAR pe = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[Item] ),
                 FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
                 CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
    VAR ke = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[Item] ),
                 FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
                 NOT CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
    VAR ps = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[Item] ),
                 FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR",
                 CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
    VAR ks = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimItem ), REMOVEFILTERS ( FactData[MC Line], FactData[Item] ),
                 FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR",
                 NOT CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
    RETURN ABS ( COALESCE ( pe, 0 ) - COALESCE ( ke, 0 ) ) + ABS ( COALESCE ( ps, 0 ) - COALESCE ( ks, 0 ) ) )

MC Comment =
VAR ent  = SELECTEDVALUE ( FactData[Entity], "All companies" )
VAR tol  = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "CHECK_TOL" )
VAR rowSel = ISFILTERED ( FactData[MC Line] ) && HASONEVALUE ( FactData[MC Line] )
VAR yrs  = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Year], FactData[Period] ),
               REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[Period], FactData[Year] ),
               FactData[Topic] = "Model Changes", NOT ISBLANK ( FactData[Year] ) )
VAR fmtP = "+0.0;-0.0;0.0"
VAR fmtS = "+#,##0;-#,##0;0"
VAR dEt  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ),
               REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[MC Line], FactData[Item] ),
               FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
               CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR dSt  = CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ),
               REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[MC Line], FactData[Item] ),
               FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR",
               CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR ppSer = CONCATENATEX ( yrs, FactData[Period] & " " & FORMAT ( CALCULATE ( [MC pp Total], REMOVEFILTERS ( DimPeriod ) ), fmtP ),
               ", ", FactData[Year], ASC )
VAR cats = SELECTCOLUMNS ( { "EIOPA changes", "other than EIOPA", "Model enhancements" }, "Cat", [Value] )
VAR cx   = ADDCOLUMNS ( cats, "@e",
               VAR k = [Cat]
               RETURN CALCULATE ( SUM ( FactData[Value] ), REMOVEFILTERS ( DimPeriod ),
                   REMOVEFILTERS ( FactData[Period], FactData[Year], FactData[MC Line], FactData[Item] ),
                   FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
                   CONTAINSSTRING ( FactData[Item], k ) ) )
VAR bigC = TOPN ( 1, FILTER ( cx, NOT ISBLANK ( [@e] ) && [@e] <> 0 ), ABS ( [@e] ), DESC )
VAR bigN = SWITCH ( MAXX ( bigC, [Cat] ), "EIOPA changes", "EIOPA changes",
               "other than EIOPA", "regulatory changes other than EIOPA", "Model enhancements", "model enhancements & other" )
VAR chk  = [MC Check]
VAR L1 = ent & ": regulatory and model changes over the plan – EOF " & FORMAT ( dEt, fmtS )
         & ", SCR " & FORMAT ( dSt, fmtS ) & " (EUR m)."
VAR L2 = "Impact on the Solvency Ratio (pp): " & ppSer & "."
VAR L3 = IF ( NOT ISBLANK ( bigN ), "Main driver on EOF: " & bigN & " (" & FORMAT ( MAXX ( bigC, [@e] ), fmtS ) & ")." )
VAR L4 = IF ( chk > tol, "Check: sub-items do NOT add up to 'Regulatory and model changes' – review the input.",
              "Check: sub-items reconcile with the total." )
-- row mode
VAR rl  = SUBSTITUTE ( SELECTEDVALUE ( FactData[MC Line] ), UNICHAR ( 8203 ), "" )
VAR rE  = CONCATENATEX ( yrs, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
              FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF", FactData[Item] = rl ) + 0, fmtS ), ", ", FactData[Year], ASC )
VAR rS  = CONCATENATEX ( yrs, FactData[Period] & " " & FORMAT ( CALCULATE ( SUM ( FactData[Value] ),
              REMOVEFILTERS ( DimPeriod ), REMOVEFILTERS ( FactData[MC Line], FactData[MC Table], FactData[Item] ),
              FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR", FactData[Item] = rl ) + 0, fmtS ), ", ", FactData[Year], ASC )
VAR rP  = CONCATENATEX ( yrs, FactData[Period] & " " & FORMAT ( CALCULATE ( [MC pp], REMOVEFILTERS ( DimPeriod ) ), fmtP ),
              ", ", FactData[Year], ASC )
VAR R1 = rl & " (" & ent & ")" & UNICHAR ( 10 ) & "EOF: " & rE & "." & UNICHAR ( 10 ) & "SCR: " & rS & "."
         & UNICHAR ( 10 ) & "Solvency Ratio impact (pp): " & rP & "."
VAR pg = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ) }
RETURN
IF ( rowSel, R1,
     IF ( ISBLANK ( dEt ) && ISBLANK ( dSt ), ent & ": no regulatory or model changes reported.",
          CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) ) )

>> E3. СТРАНИЦА "Model Changes Details" (фильтр страницы Topic = Model Changes)
>> 0) Слайсер Entity (selezione singola).
>> 1) Matrix "EOF Impacts": Righe MC Line, Colonne DimPeriod[Period], Valori MC Value,
>>    фильтр MC Table = EOF. Totali: Disattivato.
>> 2) Matrix "SCR Impacts": копия, фильтр MC Table = SCR.
>> 3) Matrix "Solvency Ratio impact (pp)": Righe MC Line, Colonne DimPeriod[Period],
>>    Valori MC pp (формат +0.0" pp"), фильтр MC Table = EOF (строки — те же категории).
>> 4) Полный ratio walk как в Excel (BoP -> ... -> EoP) уже есть на странице Solvency AoM:
>>    при желании скопируй оттуда матрицу AoM Ratio Walk.
>> 5) Comment box: MC Comment (клик по строке -> EOF, SCR и п.п. по годам).
>> ПРОВЕРКА (ES003): EOF Regulatory and model changes 114 / 10 / 30 / 6 / 16 = EIOPA changes ->
>>   MC Check = 0. Наши п.п. считаются из EOF/SCR, в dummy Excel проценты введены
>>   вручную и с ними не связаны (например 0,0% в строке total при 9,6% у EIOPA) — это dummy.


#####################################################################
 F) ДОБАВИТЬ В SUMMARY
#####################################################################

>> F1. Headline + Status (Nuova misura, папка Summary)

RAF Headline =
VAR n     = COUNTROWS ( 'List_RAF' ) + 0
VAR nChg  = COUNTROWS ( FILTER ( 'List_RAF', 'List_RAF'[RAF Change] = "Change proposed" ) ) + 0
VAR nBad  = COUNTROWS ( FILTER ( 'List_RAF', 'List_RAF'[RAF Position] IN { "Below hard limit", "Below soft limit" } ) ) + 0
VAR mn    = MINX ( 'List_RAF', 'List_RAF'[RAF Min SR] )
RETURN IF ( n > 0,
    n & " companies, lowest planned solvency ratio " & FORMAT ( mn, "0%" ) & "; "
    & IF ( nBad = 0, "all above soft limits", nBad & " below soft/hard limit" ) & "; "
    & IF ( nChg = 0, "no changes to limits proposed.", nChg & " change(s) to limits proposed." ) )

RAF Status =
VAR n     = COUNTROWS ( 'List_RAF' ) + 0
VAR nHard = COUNTROWS ( FILTER ( 'List_RAF', 'List_RAF'[RAF Position] = "Below hard limit" ) ) + 0
VAR nWch  = COUNTROWS ( FILTER ( 'List_RAF', 'List_RAF'[RAF Position] IN { "Below soft limit", "Below target range" }
                                   || 'List_RAF'[RAF Change] = "Change proposed" ) ) + 0
RETURN IF ( n > 0, SWITCH ( TRUE (), nHard > 0, 3, nWch > 0, 2, 1 ) )

MC Headline =
VAR tol = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "CHECK_TOL" )
VAR yZ  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "Model Changes" )
VAR de  = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Model Changes", FactData[MC Table] = "EOF",
              CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR ds  = CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Model Changes", FactData[MC Table] = "SCR",
              CONTAINSSTRING ( FactData[Item], "Regulatory and model" ) )
VAR ppZ = CALCULATE ( [MC pp Total], FactData[Year] = yZ )
RETURN IF ( NOT ISBLANK ( de ) || NOT ISBLANK ( ds ),
    "Regulatory & model changes over the plan: EOF " & FORMAT ( de, "+#,##0;-#,##0;0" ) & ", SCR "
    & FORMAT ( ds, "+#,##0;-#,##0;0" ) & "; ratio impact " & yZ & ": " & FORMAT ( ppZ, "+0.0;-0.0;0.0" ) & " pp"
    & IF ( [MC Check] > tol, "; sub-items do NOT add up.", "." ) )

MC Status =
VAR tol = LOOKUPVALUE ( 'Thresholds'[Value], 'Thresholds'[Key], "CHECK_TOL" )
VAR has = NOT ISBLANK ( CALCULATE ( SUM ( FactData[Value] ), FactData[Topic] = "Model Changes" ) )
RETURN IF ( has, IF ( [MC Check] > tol, 3, 0 ) )

ORMTD Headline =
VAR t   = CALCULATETABLE ( SUMMARIZE ( FactData, FactData[Entity], FactData[ORMT Name] ),
              FactData[Topic] = "ORMT Details", FactData[Value] <> 0 )
VAR n   = COUNTROWS ( t ) + 0
VAR yZ  = CALCULATE ( MAX ( FactData[Year] ), FactData[Topic] = "ORMT Details" )
VAR ppZ = CALCULATE ( [ORMT pp], FactData[Year] = yZ )
RETURN IF ( n = 0, "No ORMT / reinsurance impacts reported.",
    n & " ORMT(s) with impacts; combined solvency ratio impact " & yZ & ": " & FORMAT ( ppZ, "+0.0;-0.0;0.0" ) & " pp." )

ORMTD Status = 0

ORMTL Headline =
IF ( [ORMTL Count] = 0, "No ORMT transactions executed or renewed in the year.",
     [ORMTL Count] & " ORMT transaction(s) executed/renewed, total value " & FORMAT ( [ORMTL Value], "#,##0" ) & " (EUR m)." )

ORMTL Status = 0

OPT Headline =
VAR n    = [OPT Count]
VAR nGap = COUNTROWS ( FILTER ( 'List_Options',
               UPPER ( LEFT ( 'List_Options'[ExpectedExercise], 1 ) ) = "Y"
               && UPPER ( LEFT ( 'List_Options'[InPlan], 1 ) ) = "N" ) ) + 0
RETURN IF ( n = 0, "No options reported.",
    n & " option(s) reported, financial impact " & FORMAT ( [OPT Fin Impact], "+#,##0;-#,##0;0" )
    & IF ( nGap > 0, "; " & nGap & " expected to be exercised but not in the plan.", "." ) )

OPT Status =
VAR nGap = COUNTROWS ( FILTER ( 'List_Options',
               UPPER ( LEFT ( 'List_Options'[ExpectedExercise], 1 ) ) = "Y"
               && UPPER ( LEFT ( 'List_Options'[InPlan], 1 ) ) = "N" ) ) + 0
RETURN IF ( nGap > 0, 2, 0 )

>> F2. SummaryIndex: открой таблицу (Visualizzazione tabella > SummaryIndex > строка формул)
>> и после последней строки  { 13, ... }  поставь запятую и добавь (Page = точное имя вкладки):

        { 14, "Capital",     "Local RAF",          "RAF",   "Local RAF",             FALSE },
        { 15, "Capital",     "Model Changes",      "MC",    "Model Changes Details", TRUE  },
        { 16, "Capital",     "ORMT Impacts",       "ORMTD", "ORMT Details",          TRUE  },
        { 17, "Capital",     "ORMT Transactions",  "ORMTL", "ORMT",                  FALSE },
        { 18, "Funding",     "Options Monitoring", "OPT",   "Options Monitoring",    FALSE }

>> (у последней строки запятой в конце НЕТ). Local RAF / ORMT / Options = FALSE, потому что
>> в списках компании названы иначе (ES003...), общий слайсер Company #.. их не фильтрует.

>> F3. Замени меры Summary Text и Summary Status целиком:

Summary Text =
SWITCH ( SELECTEDVALUE ( SummaryIndex[Prefix] ),
    "SOL",  [SOL Headline],  "AoM",   [AoM Headline],   "SENS",  [SENS Headline],
    "FTC",  [FTC Headline],  "CN",    [CN Headline],    "TR",    [TR Headline],
    "RBC",  [RBC Headline],  "RH",    [RH Headline],    "HMS",   [HMS Headline],
    "SAA",  [SAA Headline],  "LK",    [LK Headline],    "NL",    [NL Headline],
    "DEBT", [DEBT Headline], "RAF",   [RAF Headline],   "MC",    [MC Headline],
    "ORMTD",[ORMTD Headline],"ORMTL", [ORMTL Headline], "OPT",   [OPT Headline] )

Summary Status =
SWITCH ( SELECTEDVALUE ( SummaryIndex[Prefix] ),
    "SOL",  [SOL Status],  "AoM",   [AoM Status],   "SENS",  [SENS Status],
    "FTC",  [FTC Status],  "CN",    [CN Status],    "TR",    [TR Status],
    "RBC",  [RBC Status],  "RH",    [RH Status],    "HMS",   [HMS Status],
    "SAA",  [SAA Status],  "LK",    [LK Status],    "NL",    [NL Status],
    "DEBT", [DEBT Status], "RAF",   [RAF Status],   "MC",    [MC Status],
    "ORMTD",[ORMTD Status],"ORMTL", [ORMTL Status], "OPT",   [OPT Status] )

>> ПРОВЕРКА Summary Total: 18 строк. Local RAF в dummy зелёный (все выше soft, изменений нет),
>> Model Changes / ORMT / Options серые (Info).
