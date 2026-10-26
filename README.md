=====================================================================
 SUMMARY BU: EXECUTIVE SUMMARY ПО КОМПАНИЯМ + "ТОЧКИ ПО ЮНИТАМ"
 Строки с ">>" не копировать. Копировать только код.
 Нужно, чтобы уже был сделан Summary_v2_Executive (ШАГ 1–3).
=====================================================================

>> ЧТО ПОЛУЧИТСЯ
>> 1) Таблица "Companies at a glance": компания | цветная точка (худший статус) |
>>    одна строка: "Solvency min 177% · 3 of 5 in line · Action: SAA · Watch: Life KPIs".
>> 2) Executive Summary, который сам меняет режим:
>>    - компания НЕ выбрана -> обзор портфеля: сколько компаний в норме,
>>      какие обсуждать и почему;
>>    - выбрана одна компания -> короткое саммари этой компании
>>      (вывод, ключевые цифры, action / watch, проверки данных).
>> 3) Клик по строке в таблице = выбор компании -> саммари переключается на неё.
>> Учитываются только разделы с SummaryIndex[ByCompany] = TRUE.


#####################################################################
 ШАГ 1. МЕРЫ ПО КОМПАНИИ (Nuova misura, папка Summary BU), по одной
#####################################################################

BU Worst =
VAR t = FILTER (
            ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ),
                         "@s", [Summary Status] ),
            [@s] IN { 1, 2, 3 } )
RETURN MAXX ( t, [@s] )

BU Rated =
COUNTROWS ( FILTER (
    ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ), "@s", [Summary Status] ),
    [@s] IN { 1, 2, 3 } ) ) + 0

BU In Line =
COUNTROWS ( FILTER (
    ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ), "@s", [Summary Status] ),
    [@s] = 1 ) ) + 0

BU Red List =
CONCATENATEX ( FILTER (
    ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ), "@s", [Summary Status] ),
    [@s] = 3 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )

BU Amber List =
CONCATENATEX ( FILTER (
    ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ), "@s", [Summary Status] ),
    [@s] = 2 ), SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )

BU Line =
VAR n  = [BU Rated]
VAR sr = [SOL Min Ratio]
VAR rl = [BU Red List]
VAR al = [BU Amber List]
RETURN IF ( n > 0,
    IF ( NOT ISBLANK ( sr ), "Solvency min " & FORMAT ( sr, "0%" ) & " · ", "" )
    & [BU In Line] & " of " & n & " in line"
    & IF ( rl <> "", " · Action: " & rl )
    & IF ( al <> "", " · Watch: " & al ) )

BU Dot = IF ( NOT ISBLANK ( [BU Worst] ), UNICHAR ( 11044 ) )

BU Color = SWITCH ( [BU Worst], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", "#9E9E9E" )

BU Label = SWITCH ( [BU Worst], 1, "In line", 2, "Watch", 3, "Action needed" )

>> Если SOL Min Ratio подчёркнута красным — у тебя она называется иначе
>> (см. Summary_v1, раздел SOLVENCY). Можно просто удалить строку с sr.


#####################################################################
 ШАГ 2. МЕРА BU EXECUTIVE SUMMARY
#####################################################################

BU Executive Summary =
VAR nl  = UNICHAR ( 10 )
VAR b   = UNICHAR ( 8226 ) & " "
VAR one = HASONEVALUE ( FactData[Entity] )
------------------ режим "одна компания" ------------------
VAR ent = SELECTEDVALUE ( FactData[Entity] )
VAR t =
    FILTER (
        ADDCOLUMNS ( FILTER ( SummaryIndex, SummaryIndex[ByCompany] = TRUE () ),
                     "@s", [Summary Status], "@txt", [Summary Text] ),
        NOT ISBLANK ( [@txt] ) && [@txt] <> "" )
VAR nG   = COUNTROWS ( FILTER ( t, [@s] = 1 ) ) + 0
VAR nA   = COUNTROWS ( FILTER ( t, [@s] = 2 ) ) + 0
VAR nR   = COUNTROWS ( FILTER ( t, [@s] = 3 ) ) + 0
VAR nAll = nG + nA + nR
VAR verdict =
    SWITCH ( TRUE (),
        nAll = 0, "no rated areas for this company",
        nR = 0 && nA = 0, "in line with the approved thresholds",
        nR = 0, "broadly in line, " & nA & " point(s) to watch",
        nR & " area(s) need action, " & nA & " to watch" )
VAR keys = { "Solvency", "Sensitivities", "Free Tangible Capital", "Remittance by Company" }
VAR kTxt = CONCATENATEX ( FILTER ( t, SummaryIndex[Section] IN keys ),
               b & [@txt], nl, SummaryIndex[Order], ASC )
VAR rTxt = CONCATENATEX ( FILTER ( t, [@s] = 3 ),
               b & SummaryIndex[Section] & ": " & [@txt], nl, SummaryIndex[Order], ASC )
VAR aTxt = CONCATENATEX ( FILTER ( t, [@s] = 2 ),
               b & SummaryIndex[Section] & ": " & [@txt], nl, SummaryIndex[Order], ASC )
VAR dq   = CONCATENATEX (
               FILTER ( t, CONTAINSSTRING ( [@txt], "NOT reconcile" )
                        || CONTAINSSTRING ( [@txt], "NOT close" )
                        || CONTAINSSTRING ( [@txt], "do not add up" ) ),
               SummaryIndex[Section], ", ", SummaryIndex[Order], ASC )
VAR companyText =
    ent & ": " & verdict & IF ( nAll > 0, " (" & nG & " of " & nAll & " areas in line)." , "." )
    & IF ( kTxt <> "", nl & nl & "KEY FIGURES" & nl & kTxt )
    & IF ( nR > 0, nl & nl & "ACTION NEEDED" & nl & rTxt )
    & IF ( nA > 0, nl & nl & "TO WATCH" & nl & aTxt )
    & IF ( dq <> "", nl & nl & "DATA CHECKS BEFORE SIGN-OFF" & nl
            & b & "Review " & dq & ": the source file does not reconcile." )
------------------ режим "все компании" ------------------
VAR ents =
    FILTER (
        ADDCOLUMNS ( CALCULATETABLE ( VALUES ( FactData[Entity] ), NOT ISBLANK ( FactData[Entity] ) ),
                     "@w", [BU Worst], "@line", [BU Line] ),
        NOT ISBLANK ( [@w] ) )
VAR cN = COUNTROWS ( ents ) + 0
VAR cR = COUNTROWS ( FILTER ( ents, [@w] = 3 ) ) + 0
VAR cA = COUNTROWS ( FILTER ( ents, [@w] = 2 ) ) + 0
VAR cG = COUNTROWS ( FILTER ( ents, [@w] = 1 ) ) + 0
VAR discuss = CONCATENATEX ( FILTER ( ents, [@w] >= 2 ),
                  b & FactData[Entity] & ": " & [@line], nl,
                  [@w], DESC, FactData[Entity], ASC )
VAR okList  = CONCATENATEX ( FILTER ( ents, [@w] = 1 ), FactData[Entity], ", ", FactData[Entity], ASC )
VAR portfolioText =
    cN & " companies: " & cR & " need action, " & cA & " to watch, " & cG & " in line."
    & IF ( discuss <> "", nl & nl & "COMPANIES TO DISCUSS" & nl & discuss )
    & IF ( okList <> "", nl & nl & "IN LINE" & nl & b & okList )
    & nl & nl & "Select a company for its summary."
RETURN IF ( one, companyText, portfolioText )

>> keys = какие разделы идут в KEY FIGURES компании (имена как в SummaryIndex[Section]).
>> Компании без оцениваемых разделов (например, холдинг "Company #4") в обзор не попадают.


#####################################################################
 ШАГ 3. (ПО ЖЕЛАНИЮ) HTML-ВЕРСИЯ С ЦВЕТНЫМИ ЗАГОЛОВКАМИ
#####################################################################
>> Только если стоит визуал HTML Content.

BU Executive Summary HTML =
VAR s  = [BU Executive Summary]
VAR h1 = SUBSTITUTE ( s,  UNICHAR ( 10 ), "<br>" )
VAR h2 = SUBSTITUTE ( h1, "KEY FIGURES",          "<b style='color:#C21B17'>KEY FIGURES</b>" )
VAR h3 = SUBSTITUTE ( h2, "ACTION NEEDED",        "<b style='color:#A50E0E'>ACTION NEEDED</b>" )
VAR h4 = SUBSTITUTE ( h3, "TO WATCH",             "<b style='color:#B07D00'>TO WATCH</b>" )
VAR h5 = SUBSTITUTE ( h4, "COMPANIES TO DISCUSS", "<b style='color:#A50E0E'>COMPANIES TO DISCUSS</b>" )
VAR h6 = SUBSTITUTE ( h5, "IN LINE",              "<b style='color:#2E7D32'>IN LINE</b>" )
VAR h7 = SUBSTITUTE ( h6, "DATA CHECKS BEFORE SIGN-OFF",
                          "<b style='color:#6E6E6E'>DATA CHECKS BEFORE SIGN-OFF</b>" )
RETURN "<div style='font-family:Segoe UI;font-size:12px;line-height:1.5;color:#333'>" & h7 & "</div>"

>> (SUBSTITUTE различает регистр: "in line" маленькими буквами в тексте не затрагивается.)


#####################################################################
 ШАГ 4. РАСКЛАДКА СТРАНИЦЫ SUMMARY BU
#####################################################################
>>  ┌──────────────────────────────────────────────────────────────┐
>>  │ шапка + слайсер Entity (Elenco a discesa)                    │
>>  ├───────────────────────────────┬──────────────────────────────┤
>>  │ BU EXECUTIVE SUMMARY          │ COMPANIES AT A GLANCE        │
>>  │ (текст/HTML, ~460 x 300)      │ (таблица, ~560 x 300)        │
>>  ├───────────────────────────────┴──────────────────────────────┤
>>  │ ТЕПЛОВАЯ КАРТА компании x разделы (как сейчас)               │
>>  └──────────────────────────────────────────────────────────────┘
>>
>> a) СЛАЙСЕР Entity: Formato > Impostazioni filtro dei dati > Selezione singola = On.
>>    Повторный клик по выбранной компании снимает выбор -> режим "все компании".
>>
>> b) BU EXECUTIVE SUMMARY
>>    Без HTML: Tabella > Colonne = BU Executive Summary; Valori: A capo automatico On;
>>    Intestazioni colonna Off; Segoe UI 10.
>>    С HTML: HTML Content > Values = BU Executive Summary HTML; Titolo off, Bordo off.
>>    Старый визуал Summary Overview / Executive Summary на этой странице убери.
>>
>> c) COMPANIES AT A GLANCE (Tabella)
>>    Colonne: FactData[Entity] | BU Dot | BU Line
>>    (то же поле Entity, что в тепловой карте и слайсере).
>>    Цвет точки: выдели BU Dot в Colonne > Formato > Elementi visivi cella >
>>    Colore carattere > fx > Formatta in base a: Valore campo > Campo: BU Color.
>>    BU Dot: Segoe UI 14, выравнивание по центру, ширина ~30 px.
>>    BU Line: A capo automatico On. Заголовки: "Company", "", "Summary".
>>    Фильтр визуала: BU Dot "non è vuoto" (уберёт строки без оценок).
>>    Сортировка "сначала проблемные": добавь в Colonne BU Worst, клик по заголовку
>>    BU Worst -> Decrescente, затем сузь колонку до минимума (Formato >
>>    Larghezza colonna) или переименуй её в "Level".
>>
>> d) ВЗАИМОДЕЙСТВИЯ
>>    Клик по строке таблицы фильтрует страницу по компании -> саммари переключается.
>>    Проверь: выдели таблицу > Formato (лента) > Modifica interazioni >
>>    на BU Executive Summary и тепловой карте значок "Filtro" (не "Nessuna").
>>
>> e) ТЕПЛОВАЯ КАРТА остаётся как есть. Подпись над ней: "Status by area".


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Ничего не выбрано: "3 companies: X need action, Y to watch, Z in line." и список
>>     "COMPANIES TO DISCUSS" с причинами, сначала красные.
>> [ ] Выбрана ES003: первая строка "ES003 – ...: ...", KEY FIGURES (Solvency, Sensitivities,
>>     FTC, Remittance by Company), ACTION / WATCH только по этой компании.
>> [ ] Цвет точки в таблице = самый плохой цвет в строке тепловой карты этой компании.
>> [ ] "x of y in line" в таблице = число зелёных точек в тепловой карте по компании.
>> [ ] Ошибка -> фото жёлтой строки под формулой.
