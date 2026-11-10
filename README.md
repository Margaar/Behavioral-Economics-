=====================================================================

=====================================================================

>> ЧТО ПОЛУЧИТСЯ
>>  - заголовки периодов красные, под ними красная линия
>>  - строки Total / Ratio: светло-красный фон, тёмно-красный шрифт
>>  - строки "o/w": серые, с отступом
>>  - колонка Actual текущего года подсвечена
>>  - отрицательные: красные в скобках; пустые: "-"
>> Ограничение: жирный шрифт по строкам через меры недоступен — выделяем цветом.
>>
>> ПЕРЕД НАЧАЛОМ должны быть сделаны:
>>  - колонки FactData[Line], [LineOrder], [BlockOrder] (хвост FactData в Power Query)
>>  - колонка PeriodSort и сортировка Period -> PeriodSort, Line -> LineOrder, Block -> BlockOrder
>>    (файл CPR_Fix_Sort_HTML, ШАГ 2)


#####################################################################
 ШАГ 1. ОТСТУП ДЛЯ "o/w" (Power Query)
#####################################################################
>> Trasforma dati > FactData > Editor avanzato.
>> Найди строку, которая начинается с  LK3    =  (2 строки) и замени её на:

    LK3    = Table.AddColumn(LK2, "Line", each
                 let it  = if [Item] = null then "" else [Item],
                     pad = if Text.StartsWith(Text.Lower(it), "o/w") or Text.StartsWith(it, "<")
                           then Text.Repeat(Character.FromNumber(8194), 4) else ""
                 in pad & it & Text.Repeat(Character.FromNumber(8203), [Dup]), type text),

>> Fine > Chiudi e applica.
>> Проверка: в FactData[Line] строки "o/w ..." начинаются с небольшого отступа.


#####################################################################
 ШАГ 2. МЕРЫ (Home > Nuova misura, папка "CPR Format")
#####################################################################

>> Тип строки: 1 = Total/Ratio, 2 = o/w и <custom>, 0 = обычная, 3 = уровень блока

Line Kind =
VAR u = UPPER ( TRIM ( SUBSTITUTE ( SUBSTITUTE ( SELECTEDVALUE ( FactData[Line] ),
            UNICHAR ( 8194 ), "" ), UNICHAR ( 8203 ), "" ) ) )
RETURN SWITCH ( TRUE (),
    NOT HASONEVALUE ( FactData[Line] ), 3,
    LEFT ( u, 5 ) = "TOTAL" || CONTAINSSTRING ( u, "RATIO" ) || u IN { "RORC", "ROCE" }, 1,
    LEFT ( u, 3 ) = "O/W" || LEFT ( u, 1 ) = "<", 2,
    0 )

>> Значение ячейки: пусто -> 0, чтобы формат показал "-"

CPR Cell = COALESCE ( [CPR Value], 0 )

>> ФОРМАТ для CPR Cell: выдели меру > Strumenti misura > Formato: Dinamico >
>> слева в строке формул выбери "Formato" и вставь:
>>
>>   IF ( MAX ( FactData[RatioFlag] ) = 1, "0.0%;-0.0%;-", "#,##0;(#,##0);-" )

>> Цвет фона ячейки

Cell BG =
VAR k    = [Line Kind]
VAR isCY = SELECTEDVALUE ( FactData[PeriodType] ) = "Actual"
           && SELECTEDVALUE ( FactData[Year] ) = [CPR Year]
RETURN SWITCH ( TRUE (),
    k = 1 && isCY, "#EBD3D1",
    k = 1,         "#F2E3E2",
    isCY,          "#FBF3F2",
    "#FFFFFF" )

>> Цвет шрифта ячейки

Cell Font =
VAR v = [CPR Value]
VAR k = [Line Kind]
RETURN SWITCH ( TRUE (),
    NOT ISBLANK ( v ) && v < 0, "#A50E0E",
    k = 1, "#7A0F0C",
    k = 2, "#6E6E6E",
    "#333333" )

>> Для матрицы Performance (Act CY / Fcst CY / Var ...) — тот же фон без подсветки колонки:

Row BG = IF ( [Line Kind] = 1, "#F2E3E2", "#FFFFFF" )


#####################################################################
 ШАГ 3. МАТРИЦА "КАК В EXCEL"
#####################################################################
>> 1) ПОЛЯ
>>      Righe:   FactData[Block] > FactData[Line]
>>      Colonne: FactData[Period]
>>      Valori:  CPR Cell   (вместо CPR Value)
>>      В Righe: стрелка у Line > "Mostra elementi senza dati" = On
>>      (видны все строки шаблона, пустые = "-")
>>
>> 2) УСЛОВНОЕ ФОРМАТИРОВАНИЕ (Formato > Elementi visivi cella > Serie: CPR Cell)
>>      Colore sfondo   On > fx > Formatta in base a: Valore campo > Campo: Cell BG
>>      Colore carattere On > fx > Formatta in base a: Valore campo > Campo: Cell Font
>>
>> 3) ФОРМАТ ВИЗУАЛА
>>      Stile predefinito ........ Nessuno (иначе тема перекрасит фон)
>>      Layout ................... Struttura; Rientro struttura 15
>>      Subtotali righe/colonne .. Off
>>      Icone +/- ................ Off; затем правый клик по матрице > Espandi > Tutto
>>      Intestazioni colonna ..... фон #FFFFFF; testo #A50E0E; Segoe UI Semibold 10;
>>                                 allineamento: a destra
>>      Griglia > Bordi .......... Sezione: Intestazione colonna; Posizione: In basso;
>>                                 colore #C21B17; larghezza 2
>>      Griglia orizzontale ...... On, colore #E3E3E3, 1 px
>>      Griglia verticale ........ Off
>>      Intestazioni riga ........ testo #333333, Segoe UI 9;
>>                                 "+/-" off; Блок (верхний уровень) — Segoe UI Semibold, #A50E0E
>>                                 (если в твоей версии есть "Applica impostazioni a: Livello 1")
>>      Valori ................... Segoe UI 9; allineamento a destra
>>      Titolo ................... Off (название листа — отдельным текстбоксом сверху)
>>
>> 4) МАТРИЦА PERFORMANCE рядом (Act CY | Fcst CY | Delta vs Fcst | Var vs Fcst | Var vs PY):
>>      те же поля Righe, те же настройки формата;
>>      Colore sfondo для каждого значения: Valore campo > Row BG;
>>      Colore carattere для Var vs Fcst: Regole: < 0 #A50E0E, > 0 #2E7D32.


#####################################################################
 ШАГ 4. ПЕРЕНЕСТИ НА ВСЕ СТРАНИЦЫ
#####################################################################
>> Выдели готовую матрицу > Home > Copia formato (кисть) > клик по матрице на другой странице.
>> Условное форматирование кистью НЕ копируется в некоторых версиях: тогда проще
>> скопировать саму матрицу (Ctrl+C / Ctrl+V) на другую страницу — фильтр страницы
>> (Topic) сам поменяет содержимое.


#####################################################################
 ПРОВЕРКА
#####################################################################
>> [ ] Колонки: Actual 2025 | Forecast 2026 | Actual 2026 | (Plan 2027)
>> [ ] Строки в порядке Excel, блоки в порядке Excel
>> [ ] "Total Remittance ...", "Solvency Ratio (%)", "RoRC" — светло-красный фон
>> [ ] "o/w ..." — серые с отступом
>> [ ] Колонка Actual 2026 слегка подсвечена
>> [ ] Пустые ячейки "-", отрицательные (123) красным
>> [ ] Строка Total не подсветилась -> её слово добавь в Line Kind в ветку "= 1"
>> [ ] Ошибка -> фото жёлтой строки под формулой
