=====================================================================
 INITIATIVES — страница + добавление в Summary
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> КАК УСТРОЕН ЛИСТ
>> Левая таблица (вывод): Initiative | Expected Timeline: Start Date, Due Date |
>> Current Status | Expected Impact: Period, on Solvency Ratio (%), on Liquidity (EUR M),
>> on P&L (EUR M) | Other Comments.
>> Одна инициатива = несколько строк (по одной на Period); имя, даты, статус и
>> комментарий — объединённые ячейки, заполнены только в первой строке.
>> Запрос ниже читает левую таблицу, протягивает имя/даты/статус вниз внутри
>> своей инициативы и даёт колонкам короткие имена.
>> В dummy лист пустой: внеси в тестовую копию 2 инициативы по 2-3 периода
>> (ровно под заголовки), одной поставь Due Date в прошлом и статус "Ongoing".


#####################################################################
 ШАГ 1. ЗАПРОС List_Initiatives — замени код существующего запроса целиком
#####################################################################
>> Trasforma dati > List_Initiatives > Editor avanzato > Ctrl+A > вставить.

let
    Sheet = Workbook{[Item = "Initiatives", Kind = "Sheet"]}[Data],
    Raw   = try fnReadList(Sheet, null, "ID", "Expected Timeline", 3, 1, 150, 0)
            otherwise fnReadList(Sheet, null, null, "Expected Timeline", 3, 1, 150, 0),
    Map = (n as text) as text =>
        let u = Text.Upper(n) in
        if u = "LABEL" or Text.StartsWith(u, "INITIATIVE") then "Initiative"
        else if Text.Contains(u, "COMMENT") then "Comments"
        else if Text.Contains(u, "START") then "StartDate"
        else if Text.Contains(u, "DUE") then "DueDate"
        else if Text.Contains(u, "STATUS") then "Status"
        else if Text.Contains(u, "SOLVENCY") then "ImpSolvency"
        else if Text.Contains(u, "LIQUIDITY") then "ImpLiquidity"
        else if Text.Contains(u, "P&L") then "ImpPL"
        else if Text.EndsWith(u, "PERIOD") then "Period"
        else n,
    Old = Table.ColumnNames(Raw),
    New = List.Accumulate(Old, {}, (s, n) =>
              let m = Map(n) in
              s & { if not List.Contains(s, m) then m
                    else if not List.Contains(s, n) then n
                    else n & " #" & Text.From(List.Count(s)) }),
    Ren = Table.RenameColumns(Raw, List.Zip({Old, New})),

    // имя инициативы вниз, потом даты/статус/комментарий — только внутри своей инициативы
    Blank0  = Table.ReplaceValue(Ren, "", null, Replacer.ReplaceValue, Table.ColumnNames(Ren)),
    FillIni = Table.FillDown(Blank0, {"Initiative"}),
    Attr    = List.Intersect({{"StartDate", "DueDate", "Status", "Comments"}, Table.ColumnNames(FillIni)}),
    Grp     = Table.Group(FillIni, {"Initiative"}, {{"T", each Table.FillDown(_, Attr), type table}}),
    Back    = Table.Combine(Grp[T]),
    Kept    = Table.SelectRows(Back, each [Initiative] <> null),

    ToNum  = (v) => if v is number then v else try Number.From(v) otherwise null,
    ToDate = (v) => if v is date then v
                    else if v is datetime then DateTime.Date(v)
                    else if v is number then Date.From(v)
                    else try Date.FromText(Text.From(v), [Culture = "en-GB"]) otherwise null,
    Typed  = Table.TransformColumns(Kept,
                 {{"ImpSolvency", ToNum, type number}, {"ImpLiquidity", ToNum, type number},
                  {"ImpPL", ToNum, type number},
                  {"StartDate", ToDate, type date}, {"DueDate", ToDate, type date},
                  {"Period", each if _ = null then null else Text.Trim(Text.From(_)), type text},
                  {"Status", each if _ = null then null else Text.Trim(Text.From(_)), type text}},
                 null, MissingField.Ignore),
    Result = Table.AddColumn(Typed, "Year", each
                 let p = try Record.Field(_, "Period") otherwise null,
                     d = if p = null then "" else Text.Select(p, {"0".."9"})
                 in if Text.Length(d) = 4 then Number.From(d)
                    else if Text.Length(d) = 2 then 2000 + Number.From(d) else null, Int64.Type)
in
    Result

>> ПРОВЕРКА: колонки Initiative, StartDate, DueDate, Status, Period, ImpSolvency,
>> ImpLiquidity, ImpPL, Comments, Year. У каждой строки-периода заполнено имя инициативы.
>> Если какая-то колонка осталась с длинным именем или Initiative пустая — пришли фото
>> этого запроса, поправлю одно правило (у листа может быть 2, а не 3 строки заголовка).
>> Chiudi e applica.


#####################################################################
 ШАГ 2. МЕРЫ (Home > Nuova misura; Tabella home = _Measures, папка "Initiatives")
#####################################################################

INI Count = DISTINCTCOUNT ( 'List_Initiatives'[Initiative] ) + 0

INI Solvency = SUM ( 'List_Initiatives'[ImpSolvency] )

INI Liquidity = SUM ( 'List_Initiatives'[ImpLiquidity] )

INI PL = SUM ( 'List_Initiatives'[ImpPL] )

>> Просрочена: Due Date прошла, а статус не "completed / closed / done"
INI Overdue =
VAR t = SUMMARIZE ( 'List_Initiatives', 'List_Initiatives'[Initiative],
            "@due", MAX ( 'List_Initiatives'[DueDate] ),
            "@st",  MAX ( 'List_Initiatives'[Status] ) )
RETURN COUNTROWS ( FILTER ( t,
           NOT ISBLANK ( [@due] ) && [@due] < TODAY ()
           && NOT ( CONTAINSSTRING ( [@st], "complet" ) || CONTAINSSTRING ( [@st], "closed" )
                    || CONTAINSSTRING ( [@st], "done" ) ) ) ) + 0

>> Формат: INI Solvency -> Personalizzato  +0.0%;-0.0%;0.0%
>>         INI Liquidity, INI PL -> +#,##0.0;-#,##0.0;0.0

INI Comment =
VAR n     = [INI Count]
VAR oneI  = ISFILTERED ( 'List_Initiatives'[Initiative] ) && n = 1
VAR fmtP  = "+0.0%;-0.0%;0.0%"
VAR fmtE  = "+#,##0.0;-#,##0.0;0.0"
VAR ini   = SUMMARIZE ( 'List_Initiatives', 'List_Initiatives'[Initiative],
                "@sol", SUM ( 'List_Initiatives'[ImpSolvency] ),
                "@liq", SUM ( 'List_Initiatives'[ImpLiquidity] ),
                "@pl",  SUM ( 'List_Initiatives'[ImpPL] ),
                "@due", MAX ( 'List_Initiatives'[DueDate] ),
                "@st",  MAX ( 'List_Initiatives'[Status] ) )
VAR byStatus = CONCATENATEX ( VALUES ( 'List_Initiatives'[Status] ),
                COALESCE ( 'List_Initiatives'[Status], "no status" ) & " "
                & CALCULATE ( DISTINCTCOUNT ( 'List_Initiatives'[Initiative] ) ), ", " )
VAR bigS  = TOPN ( 1, FILTER ( ini, NOT ISBLANK ( [@sol] ) && [@sol] <> 0 ), ABS ( [@sol] ), DESC )
VAR bigL  = TOPN ( 1, FILTER ( ini, NOT ISBLANK ( [@liq] ) && [@liq] <> 0 ), ABS ( [@liq] ), DESC )
VAR late  = FILTER ( ini, NOT ISBLANK ( [@due] ) && [@due] < TODAY ()
                && NOT ( CONTAINSSTRING ( [@st], "complet" ) || CONTAINSSTRING ( [@st], "closed" )
                         || CONTAINSSTRING ( [@st], "done" ) ) )
VAR yrs   = CALCULATETABLE ( VALUES ( 'List_Initiatives'[Year] ), NOT ISBLANK ( 'List_Initiatives'[Year] ) )
VAR bySol = CONCATENATEX ( yrs, 'List_Initiatives'[Year] & " " & FORMAT ( [INI Solvency], fmtP ),
                ", ", 'List_Initiatives'[Year], ASC )
-- page mode
VAR L1 = n & IF ( n = 1, " initiative", " initiatives" ) & " ongoing/planned (" & byStatus & ")."
VAR L2 = "Total expected impact: solvency ratio " & FORMAT ( [INI Solvency], fmtP )
         & ", liquidity " & FORMAT ( [INI Liquidity], fmtE ) & ", P&L " & FORMAT ( [INI PL], fmtE ) & " (EUR m)."
VAR L3 = IF ( NOT ISBLANK ( bySol ) && bySol <> "", "Solvency impact by year: " & bySol & "." )
VAR L4 = IF ( NOT ISBLANK ( MAXX ( bigS, [@sol] ) ),
         "Largest solvency impact: " & MAXX ( bigS, 'List_Initiatives'[Initiative] )
         & " (" & FORMAT ( MAXX ( bigS, [@sol] ), fmtP ) & ")"
         & IF ( NOT ISBLANK ( MAXX ( bigL, [@liq] ) ),
                "; largest liquidity impact: " & MAXX ( bigL, 'List_Initiatives'[Initiative] )
                & " (" & FORMAT ( MAXX ( bigL, [@liq] ), fmtE ) & ")" ) & "." )
VAR L5 = IF ( COUNTROWS ( late ) > 0, "Past due date and not completed: "
         & CONCATENATEX ( late, 'List_Initiatives'[Initiative] & " (due " & FORMAT ( [@due], "mmm-yy" ) & ")", "; " ) & "." )
-- one initiative clicked
VAR nm  = SELECTEDVALUE ( 'List_Initiatives'[Initiative] )
VAR C1  = nm & ": " & COALESCE ( SELECTEDVALUE ( 'List_Initiatives'[Status] ), "no status" )
          & ", " & FORMAT ( MIN ( 'List_Initiatives'[StartDate] ), "mmm-yy" ) & " to "
          & FORMAT ( MAX ( 'List_Initiatives'[DueDate] ), "mmm-yy" )
          & IF ( [INI Overdue] > 0, " – past due date.", "." )
VAR C2  = "Expected impact: solvency " & FORMAT ( [INI Solvency], fmtP ) & ", liquidity "
          & FORMAT ( [INI Liquidity], fmtE ) & ", P&L " & FORMAT ( [INI PL], fmtE ) & " (EUR m)"
          & IF ( NOT ISBLANK ( bySol ) && bySol <> "", "; by year: " & bySol, "" ) & "."
VAR C3  = IF ( NOT ISBLANK ( SELECTEDVALUE ( 'List_Initiatives'[Comments] ) ),
          "Comment: " & SELECTEDVALUE ( 'List_Initiatives'[Comments] ) )
VAR pg  = { ( 1, L1 ), ( 2, L2 ), ( 3, L3 ), ( 4, L4 ), ( 5, L5 ) }
VAR one = { ( 1, C1 ), ( 2, C2 ), ( 3, C3 ) }
RETURN
SWITCH ( TRUE (),
    n = 0, "No initiatives reported.",
    oneI, CONCATENATEX ( FILTER ( one, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ),
    CONCATENATEX ( FILTER ( pg, [Value2] <> "" ), [Value2], UNICHAR ( 10 ), [Value1], ASC ) )


#####################################################################
 ШАГ 3. СТРАНИЦА "Initiatives" (без фильтра Topic, без Scale)
#####################################################################
>> 1) Tabella "Overview": Initiative, StartDate, DueDate, Status, Comments
>>    (даты: формат gg/mm/aaaa или mmm-yy). Totali: Disattivato.
>>    Status: Formattazione condizionale > Colore sfondo > Regole
>>    (contiene "Complet" -> зелёный, "Ongoing" -> жёлтый, "Planned" -> серый).
>> 2) Matrix "Expected impact": Righe Initiative | Colonne Period | Valori INI Solvency,
>>    INI Liquidity, INI PL. Formato > Valori > "Cambia valori nelle righe" = Attivato,
>>    чтобы под каждой инициативой шли 3 строки как в Excel.
>> 3) Карточки: INI Count | INI Solvency | INI Liquidity | INI Overdue.
>> 4) Comment box: INI Comment (клик по инициативе в таблице -> текст по ней).
>> ПРОВЕРКА на тестовых строках: сумма "on Solvency Ratio" по инициативе = INI Solvency;
>> инициатива с прошедшей Due Date и статусом Ongoing попадает в "Past due date".


#####################################################################
 ШАГ 4. ДОБАВИТЬ В SUMMARY
#####################################################################

INI Headline =
VAR n = [INI Count]
RETURN IF ( n = 0, "No initiatives reported.",
    n & " initiative(s); expected impact on solvency " & FORMAT ( [INI Solvency], "+0.0%;-0.0%;0.0%" )
    & ", liquidity " & FORMAT ( [INI Liquidity], "+#,##0;-#,##0;0" ) & " (EUR m)"
    & IF ( [INI Overdue] > 0, "; " & [INI Overdue] & " past due date.", "." ) )

INI Status = IF ( [INI Overdue] > 0, 2, 0 )

>> SummaryIndex: после строки { 18, ... } поставь запятую и добавь
>> (Page = точное имя вкладки):

        { 19, "Business",    "Initiatives",        "INI",   "Initiatives",           FALSE }

>> Summary Text: перед последней скобкой добавь   , "INI", [INI Headline]
>> Summary Status: перед последней скобкой добавь , "INI", [INI Status]
>> Пример конца Summary Text:
>>     "ORMTD",[ORMTD Headline],"ORMTL", [ORMTL Headline], "OPT",   [OPT Headline],
>>     "INI",  [INI Headline] )
