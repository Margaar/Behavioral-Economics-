=====================================================================
 SUMMARY v2: ИСПРАВЛЕНИЯ + EXECUTIVE SUMMARY + HTML-ВИД
 Строки с ">>" не копировать. Копировать только код.
 Перед началом: сохрани копию .pbix.
=====================================================================

>> ПОРЯДОК
>>   ШАГ 1. Sensitivities: исправить 2905% (ошибка суммирования)
>>   ШАГ 2. Debt: пустое число в тексте
>>   ШАГ 3. Мера Executive Summary
>>   ШАГ 4. Поставить на страницу Summary Total
>>   ШАГ 5. (по желанию) HTML-вид: Executive Summary + вся таблица Summary
>>   ШАГ 6. Комментарии страниц списком (HTML)
>>   ПРОВЕРКА


#####################################################################
 ШАГ 1. SENSITIVITIES: ИСПРАВИТЬ 2905%
#####################################################################
>> ПОЧЕМУ ОШИБКА
>> Sens Base Ratio = СУММА строки "Solvency Ratio" по всем фильтрам, которые остались.
>> На странице Sensitivities это ок: там выбран год и компания.
>> На Summary нет ни года, ни компании -> складываются все годы (Act 2025 ... Plan 2029)
>> и все компании: примерно 3 компании x 5 периодов x ~200% = ~2985%.
>> То же самое с ударом шока (-79,9 pp = сумма по годам и компаниям).
>> Складывать коэффициенты нельзя. Правильно: последний год, по каждой компании
>> отдельно, и показать ХУДШУЮ компанию.
>> Сама мера Sens Base Ratio и страница Sensitivities НЕ меняются.

>> 1a) НОВАЯ мера (Home > Nuova misura, папка Summary):

SENS Year =
CALCULATE ( MAX ( FactData[Year] ), REMOVEFILTERS ( DimItem ),
    FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency" )

>> 1b) НОВАЯ мера: худший шок (в pp) для текущего контекста (одна компания, один год)

SENS Ent Worst PP =
VAR sh = CALCULATETABLE ( VALUES ( FactData[Sens Shock] ),
             FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency" )
VAR x  = FILTER ( ADDCOLUMNS ( sh, "@pp", [Sens Ratio pp] ),
             NOT ISBLANK ( [@pp] ) && NOT ( LEFT ( FactData[Sens Shock], 1 ) IN { "[", "<" } ) )
RETURN MINX ( x, [@pp] )

>> 1c) ЗАМЕНИ формулу меры SENS Worst After (клик по мере > строка формул > вставить всё):

SENS Worst After =
VAR y  = [SENS Year]
VAR ents = CALCULATETABLE ( VALUES ( FactData[Entity] ),
             FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency" )
VAR t  = FILTER (
             ADDCOLUMNS ( ents,
                 "@b", CALCULATE ( [Sens Base Ratio],   FactData[Year] = y ),
                 "@w", CALCULATE ( [SENS Ent Worst PP], FactData[Year] = y ) ),
             NOT ISBLANK ( [@b] ) && NOT ISBLANK ( [@w] ) )
RETURN IF ( NOT ISEMPTY ( t ), MINX ( t, [@b] + [@w] ) / 100 )

>> SENS Status менять НЕ нужно: он берёт SENS Worst After и пороги SENS_RED / SENS_AMBER.

>> 1d) ЗАМЕНИ формулу меры SENS Headline:

SENS Headline =
VAR y  = [SENS Year]
VAR ents = CALCULATETABLE ( VALUES ( FactData[Entity] ),
             FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency" )
VAR t  = FILTER (
             ADDCOLUMNS ( ents,
                 "@b", CALCULATE ( [Sens Base Ratio],   FactData[Year] = y ),
                 "@w", CALCULATE ( [SENS Ent Worst PP], FactData[Year] = y ) ),
             NOT ISBLANK ( [@b] ) && NOT ISBLANK ( [@w] ) )
VAR n   = COUNTROWS ( t )
VAR wr  = TOPN ( 1, t, [@b] + [@w], ASC )
VAR wE  = MAXX ( wr, FactData[Entity] )
VAR wB  = MAXX ( wr, [@b] )
VAR wPP = MAXX ( wr, [@w] )
VAR shk = CALCULATETABLE (
              FILTER ( ADDCOLUMNS ( VALUES ( FactData[Sens Shock] ), "@pp", [Sens Ratio pp] ),
                       NOT ISBLANK ( [@pp] ) && NOT ( LEFT ( FactData[Sens Shock], 1 ) IN { "[", "<" } ) ),
              FactData[Topic] = "Sensitivities", FactData[Sens Metric] = "Solvency",
              FactData[Year] = y, FactData[Entity] = wE )
VAR wN  = MAXX ( TOPN ( 1, shk, [@pp], ASC ), FactData[Sens Shock] )
RETURN IF ( n > 0,
    "Worst sensitivity " & y & ": " & wN & " " & FORMAT ( wPP, "+0.0;-0.0" ) & " pp"
    & IF ( n > 1, " at " & wE, "" )
    & ", ratio " & FORMAT ( wB / 100, "0%" ) & " → " & FORMAT ( ( wB + wPP ) / 100, "0%" )
    & IF ( n > 1, " (lowest of " & n & " companies)", "" ) & "." )

>> ОЖИДАЕМО (dummy): base ~180–220%, удар шока в единицах pp (например -3,7 pp),
>> ratio после шока ~170–215%. Если снова 1000%+ — пришли фото, значит в одном году
>> у компании несколько строк "Solvency Ratio".


#####################################################################
 ШАГ 2. DEBT: ПУСТОЕ ЧИСЛО ("reimbursements , issuances ...")
#####################################################################
>> Если погашений нет, мера пустая и FORMAT даёт пустую строку. Добавляем +0.
>> ЗАМЕНИ формулу DEBT Headline:

DEBT Headline =
IF ( [Debt Instruments] > 0,
     [Debt Instruments] & " instrument(s) in the proposals: reimbursements "
     & FORMAT ( [Debt Reimbursements] + 0, "#,##0" )
     & ", issuances " & FORMAT ( [Debt Issuances] + 0, "#,##0" )
     & ", net " & FORMAT ( [Debt Net] + 0, "+#,##0;-#,##0;0" ) & ".",
     "No reimbursement or issuance proposals." )


#####################################################################
 ШАГ 3. МЕРА EXECUTIVE SUMMARY  (Nuova misura, папка Summary)
#####################################################################
>> Собирает готовые фразы [Summary Text] и статусы [Summary Status] в 4 блока:
>>   вывод одной строкой | KEY FIGURES | ACTION NEEDED / TO WATCH (с цифрами) |
>>   DATA CHECKS (ошибки в файле, а не проблемы бизнеса).
>> Работает и для страны, и для выбранной компании (слайсер Entity).

Executive Summary =
VAR ent  = SELECTEDVALUE ( FactData[Entity], "Country total" )
VAR nl   = UNICHAR ( 10 )
VAR b    = UNICHAR ( 8226 ) & " "
VAR t =
    FILTER (
        ADDCOLUMNS ( SummaryIndex, "@s", [Summary Status], "@txt", [Summary Text] ),
        NOT ISBLANK ( [@txt] ) && [@txt] <> ""
    )
VAR nG   = COUNTROWS ( FILTER ( t, [@s] = 1 ) ) + 0
VAR nA   = COUNTROWS ( FILTER ( t, [@s] = 2 ) ) + 0
VAR nR   = COUNTROWS ( FILTER ( t, [@s] = 3 ) ) + 0
VAR nAll = nG + nA + nR
VAR verdict =
    SWITCH ( TRUE (),
        nR = 0 && nA = 0, "in line with the approved thresholds",
        nR = 0, "broadly in line, " & nA & " point(s) to watch",
        nR & " area(s) need action, " & nA & " to watch" )
VAR keys = { "Solvency", "Sensitivities", "Total Remittance", "Remittance Holding", "Free Tangible Capital" }
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
RETURN
ent & ": " & verdict & " (" & nG & " of " & nAll & " areas in line)."
& IF ( kTxt <> "", nl & nl & "KEY FIGURES" & nl & kTxt )
& IF ( nR > 0, nl & nl & "ACTION NEEDED" & nl & rTxt )
& IF ( nA > 0, nl & nl & "TO WATCH" & nl & aTxt )
& IF ( dq <> "", nl & nl & "DATA CHECKS BEFORE SIGN-OFF" & nl
        & b & "Review " & dq & ": the source file does not reconcile." )

>> keys = какие разделы попадают в KEY FIGURES. Имена должны совпадать с
>> SummaryIndex[Section]. Лишнее/отсутствующее имя ошибки не даёт — просто пропускается.
>> Старая мера Summary Overview больше не нужна (не удаляй, пока не проверишь новую).


#####################################################################
 ШАГ 4. НА СТРАНИЦУ SUMMARY TOTAL
#####################################################################
>> Вариант A (без HTML, быстрый):
>> 1) Выдели визуал "Summary Overview" > в поле Valori/Colonne замени меру
>>    Summary Overview на Executive Summary.
>> 2) Formato > Valori: A capo automatico = Attivato; Segoe UI 10, #333333.
>>    Intestazioni colonna: Disattivato (заголовок "Summary Overview" оставь красной
>>    плашкой-текстбоксом, как сейчас).
>> 3) Высота визуала ~280–320 px (сейчас текст обрежется). Таблицу разделов
>>    сдвинь ниже — она остаётся как детализация.
>> 4) Summary BU: скопируй визуал туда же (Ctrl+C / Ctrl+V). Слайсер Entity
>>    поменяет первую строку на имя компании.
>>
>> Вариант B (красивее, с цветными заголовками) — ШАГ 5.


#####################################################################
 ШАГ 5. (ПО ЖЕЛАНИЮ) HTML-ВИД
#####################################################################
>> Нужен визуал HTML Content (один раз на отчёт):
>> Visualizzazioni > "..." > Ottieni altri oggetti visivi > AppSource > HTML Content > Aggiungi.
>> Если AppSource закрыт политикой — оставайся на Варианте A.

>> 5a) Executive Summary в HTML (заголовки блоков жирные и цветные):

Executive Summary HTML =
VAR s = [Executive Summary]
VAR h1 = SUBSTITUTE ( s,  UNICHAR ( 10 ), "<br>" )
VAR h2 = SUBSTITUTE ( h1, "KEY FIGURES",   "<b style='color:#C21B17'>KEY FIGURES</b>" )
VAR h3 = SUBSTITUTE ( h2, "ACTION NEEDED", "<b style='color:#A50E0E'>ACTION NEEDED</b>" )
VAR h4 = SUBSTITUTE ( h3, "TO WATCH",      "<b style='color:#B07D00'>TO WATCH</b>" )
VAR h5 = SUBSTITUTE ( h4, "DATA CHECKS BEFORE SIGN-OFF",
                          "<b style='color:#6E6E6E'>DATA CHECKS BEFORE SIGN-OFF</b>" )
RETURN "<div style='font-family:Segoe UI;font-size:12px;line-height:1.5;color:#333'>" & h5 & "</div>"

>> Поставить: HTML Content > поле Values = Executive Summary HTML. Titolo off, Bordo off.

>> 5b) Вся таблица Summary в HTML (точка | раздел | фраза, сгруппировано по областям,
>>     сверху 3 "таблетки" In line / Watch / Action needed):

Summary HTML =
VAR t = FILTER ( ADDCOLUMNS ( SummaryIndex, "@s", [Summary Status], "@txt", [Summary Text] ),
            NOT ISBLANK ( [@txt] ) && [@txt] <> "" )
VAR nG = COUNTROWS ( FILTER ( t, [@s] = 1 ) ) + 0
VAR nA = COUNTROWS ( FILTER ( t, [@s] = 2 ) ) + 0
VAR nR = COUNTROWS ( FILTER ( t, [@s] = 3 ) ) + 0
VAR who = SELECTEDVALUE ( FactData[Entity], "Country total" )
VAR head =
    "<div style='font-size:15px;font-weight:600;color:#333;margin-bottom:6px'>" & who & "</div>"
    & "<div style='margin-bottom:10px'>"
    & "<span style='display:inline-block;padding:3px 10px;margin-right:6px;border-radius:12px;background:#2E7D32;color:#fff;font-weight:600;font-size:11px'>" & nG & " in line</span>"
    & "<span style='display:inline-block;padding:3px 10px;margin-right:6px;border-radius:12px;background:#E0A100;color:#fff;font-weight:600;font-size:11px'>" & nA & " watch</span>"
    & "<span style='display:inline-block;padding:3px 10px;border-radius:12px;background:#A50E0E;color:#fff;font-weight:600;font-size:11px'>" & nR & " action needed</span>"
    & "</div>"
VAR areas = SUMMARIZE ( t, SummaryIndex[Area], "@o", MIN ( SummaryIndex[Order] ) )
VAR body =
    CONCATENATEX ( areas,
        VAR a = SummaryIndex[Area]
        RETURN
        "<div style='color:#C21B17;font-weight:600;font-size:12px;margin:12px 0 4px;padding-bottom:3px;border-bottom:1px solid #E3E3E3'>"
        & a & "</div><table style='border-collapse:collapse;width:100%'>"
        & CONCATENATEX ( FILTER ( t, SummaryIndex[Area] = a ),
            "<tr><td style='width:16px;vertical-align:top;padding:6px 0'>"
            & "<span style='display:inline-block;width:10px;height:10px;border-radius:50%;background:"
            & SWITCH ( [@s], 1, "#2E7D32", 2, "#E0A100", 3, "#A50E0E", "#9E9E9E" ) & "'></span></td>"
            & "<td style='width:170px;vertical-align:top;padding:4px 8px;font-weight:600;color:#333'>"
            & SummaryIndex[Section] & "</td>"
            & "<td style='vertical-align:top;padding:4px 0;color:#555'>" & [@txt] & "</td></tr>",
            "", SummaryIndex[Order], ASC )
        & "</table>",
        "", [@o], ASC )
RETURN
"<div style='font-family:Segoe UI;font-size:12px;line-height:1.45'>" & head & body & "</div>"

>> Поставить: второй HTML Content под Executive Summary, Values = Summary HTML.
>> Старую таблицу разделов спрячь (Visualizza > Selezione > глаз), не удаляй сразу.
>>
>> КНОПКА "Open … →": HTML-визуал не передаёт клики. Оставь узкий слайсер
>> SummaryIndex[Section] (Elenco a discesa, Selezione singola) рядом с кнопкой.
>> Слайсер не должен фильтровать HTML-визуалы: выдели слайсер > Formato >
>> Modifica interazioni > на HTML-визуалах значок "Nessuna".
>>
>> Summary BU: скопируй оба HTML-визуала; на Summary HTML добавь фильтр визуала
>> SummaryIndex[ByCompany] = True.


#####################################################################
 ШАГ 6. КОММЕНТАРИИ НА СТРАНИЦАХ СПИСКОМ (HTML, по желанию)
#####################################################################
>> Одна мера-обёртка на страницу. Шаблон (меняешь только имя меры внутри []):

SOL Comment HTML =
"<div style='font-family:Segoe UI;font-size:12px;line-height:1.45;color:#333'><ul style='margin:0;padding-left:16px'><li style='margin-bottom:5px'>"
& SUBSTITUTE ( [SOL Comment], UNICHAR ( 10 ), "</li><li style='margin-bottom:5px'>" )
& "</li></ul></div>"

>> Копии:  TR -> [TR Comment]   CN -> [CN Comment]   AoM -> [AoM Comment]
>>         Sens -> [Sens Comment]  FTC -> [FTC Comment]  RH -> [RH Comment]
>>         HMS -> [HMS Comment]  SAA -> [SAA Comment]  RBC -> [RBC Comment]
>>         LK -> [LK Comment All]  NL -> [NL Comment]  Debt -> [Debt Comment]
>>         RAF -> [RAF Comment]  OPT -> [OPT Comment]  ORMTL -> [ORMTL Comment]
>>         ORMT -> [ORMT Comment]  MC -> [MC Comment]  INI -> [INI Comment]
>> На странице замени карточку комментария на HTML Content с этой мерой
>> (то же место и размер; заголовок "Key messages" оставь текстбоксом).
>> Без AppSource: оставь карточки как есть.


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Sensitivities в Summary: "Worst sensitivity 2029: ... at <компания>, ratio 2xx% → 2xx%".
>>     Нет четырёхзначных процентов. Точка зелёная/жёлтая/красная по порогам 110/130.
>> [ ] Страница Sensitivities не изменилась (те же цифры, что раньше).
>> [ ] Debt: "reimbursements 0, issuances 38.242, net +38.242."
>> [ ] Executive Summary: первая строка "Country total: N area(s) need action...",
>>     KEY FIGURES с Solvency / Sensitivities / Remittance / Holding / FTC,
>>     ACTION NEEDED с цифрами, DATA CHECKS называет Solvency AoM и SAA (dummy).
>> [ ] Число "x of y areas in line" = число зелёных точек в таблице.
>> [ ] Summary BU: выбор компании меняет первую строку и тексты.
>> [ ] Ошибка в формуле -> пришли фото жёлтой строки под формулой.
