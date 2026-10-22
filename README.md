=====================================================================
 SUMMARY И КОММЕНТАРИИ В ЧИТАБЕЛЬНОМ ВИДЕ (HTML)
 Строки с ">>" не копировать. Копировать только код.
=====================================================================

>> ЧТО ПОЛУЧИТСЯ
>> Summary: сверху имя (страна/компания) и 3 "таблетки" In line / Watch / Action needed,
>> ниже разделы по областям (Capital, Remittance...) с красным подзаголовком,
>> в каждой строке: цветная точка | название жирным | фраза с цифрой серым.
>> Комментарии на страницах: список пунктов вместо сплошного абзаца.
>> Одна мера Summary HTML работает на обеих Summary-страницах (подстраивается под слайсер).


#####################################################################
 ШАГ 1. ДОБАВИТЬ ВИЗУАЛ HTML CONTENT (1 раз на весь отчёт)
#####################################################################
>> Report (не Power Query) > в панели Visualizzazioni нажми "..." (Altri oggetti visivi)
>>   > "Ottieni altri oggetti visivi" / "Da AppSource" > поиск: HTML Content > Aggiungi.
>> В панели появится новая иконка "HTML Content".
>> Если AppSource закрыт политикой компании -> сразу ШАГ 6 (вариант без него).


#####################################################################
 ШАГ 2. МЕРА Summary HTML  (Home > Nuova misura; Tabella home _Measures, папка Summary)
#####################################################################

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

>> Если ошибка: пришли фото жёлтой строки под формулой.
>> Если таблица SummaryIndex у тебя без колонки Area / Order / Section — скажи, поправлю.


#####################################################################
 ШАГ 3. ПОСТАВИТЬ НА СТРАНИЦУ SUMMARY – TOTAL
#####################################################################
>> 1) Кликни по пустому месту страницы > в панели Visualizzazioni иконка HTML Content.
>> 2) В поле "Values" перетащи меру Summary HTML.
>> 3) Размер: основная область (по сетке дизайна: X 196, Y 124, W 760, H 580).
>> 4) Formato визуала: Titolo off; Sfondo белый; Bordo off.
>>    В настройках визуала (Content formatting / Formattazione contenuto):
>>    "Allow opening URLs" не нужно; "Show raw HTML" = off.
>> 5) Старые визуалы Summary Overview и таблицу разделов можно удалить
>>    (или сначала спрятать: Visualizza > Selezione > значок глаза).
>> 6) Кнопка "Open … →": HTML-визуал НЕ передаёт клики в Power BI. Чтобы кнопка
>>    работала, оставь узкий слайсер справа:
>>    Filtro dei dati > поле SummaryIndex[Section] > Stile: Elenco a discesa (или Riquadro),
>>    Selezione singola: on. Выбор раздела в слайсере -> кнопка "Open <раздел> →".
>>    Слайсер НЕ должен фильтровать HTML-визуал: выдели слайсер > Formato >
>>    Modifica interazioni > на HTML-визуале значок "Nessuna".


#####################################################################
 ШАГ 4. SUMMARY – COMPANIES
#####################################################################
>> Скопируй HTML-визуал (Ctrl+C / Ctrl+V) на Summary – Companies.
>> Слайсер Entity уже там: при выборе компании заголовок станет её именем, точки и фразы —
>> по этой компании. Фильтр ByCompany = True добавь на сам HTML-визуал
>> (Filtri in questo oggetto visivo > SummaryIndex[ByCompany] = True).
>> Тепловую карту компаний оставь как есть.


#####################################################################
 ШАГ 5. КОММЕНТАРИИ НА СТРАНИЦАХ СПИСКОМ
#####################################################################
>> Для каждой страницы — одна мера-обёртка вокруг её комментария.
>> Шаблон (меняешь только имя меры в квадратных скобках внутри):

SOL Comment HTML =
"<div style='font-family:Segoe UI;font-size:12px;line-height:1.45;color:#333'><ul style='margin:0;padding-left:16px'><li style='margin-bottom:5px'>"
& SUBSTITUTE ( [SOL Comment], UNICHAR ( 10 ), "</li><li style='margin-bottom:5px'>" )
& "</li></ul></div>"

>> Имена для копий (создай те, что нужны):
>>   TR Comment HTML    -> [TR Comment]        CN Comment HTML   -> [CN Comment]
>>   AoM Comment HTML   -> [AoM Comment]       Sens Comment HTML -> [Sens Comment]
>>   FTC Comment HTML   -> [FTC Comment]       RH Comment HTML   -> [RH Comment]
>>   HMS Comment HTML   -> [HMS Comment]       SAA Comment HTML  -> [SAA Comment]
>>   RBC Comment HTML   -> [RBC Comment]       LK Comment HTML   -> [LK Comment All]
>>   NL Comment HTML    -> [NL Comment]        Debt Comment HTML -> [Debt Comment]
>>   RAF Comment HTML   -> [RAF Comment]       OPT Comment HTML  -> [OPT Comment]
>>   ORMTL Comment HTML -> [ORMTL Comment]     ORMT Comment HTML -> [ORMT Comment]
>>   MC Comment HTML    -> [MC Comment]        INI Comment HTML  -> [INI Comment]
>> На странице: замени карточку комментария на HTML Content с этой мерой
>> (тот же размер и место, заголовок "Key messages" оставь текстбоксом над ним).
>> Клик по строке матрицы продолжает менять комментарий — HTML-визуал фильтруется как обычный.
>> Пустые пункты: если в комментарии есть пустая строка, будет пустая точка списка —
>> это видно сразу, скажи, уберу.


#####################################################################
 ШАГ 6. ЕСЛИ APPSOURCE ЗАКРЫТ (без HTML)
#####################################################################
>> Summary: вместо таблицы — Matrice:
>>   Righe: SummaryIndex[Area] > SummaryIndex[Section]  (Area как группа)
>>   Valori: Summary Dot, Summary Text, Summary Label
>>   Formato > Layout: Struttura (rientrato); Intestazioni riga: Area жирным, цвет #C21B17;
>>   Valori: A capo automatico on; Spaziatura righe 6-8 px; Totali off; Griglia verticale off.
>> Комментарии: оставь карточки, но в мерах замени разделитель строк
>>   UNICHAR ( 10 )  на  UNICHAR ( 10 ) & UNICHAR ( 8226 ) & " "
>>   и добавь в начало текста  UNICHAR ( 8226 ) & " "  — получатся пункты "•".


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Summary – Total: имя "Country total", 3 таблетки, области с красными подзаголовками.
>> [ ] Число зелёных/жёлтых/красных в таблетках = число точек каждого цвета ниже.
>> [ ] Summary – Companies: выбор компании меняет заголовок и строки.
>> [ ] Комментарий страницы: пункты списком; клик по строке матрицы -> текст по строке.
